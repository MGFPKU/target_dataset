# -*- coding: utf-8 -*-
"""Push local commits to GitHub via the git database API (github.com:443 blocked).

Replicates tree/blob/commit objects exactly, so remote SHAs equal local SHAs.
Pushes the whole unpushed range (oldest first), one commit per API round.
GitHub keeps timezone offsets and message bytes as sent; the message must be
extracted verbatim from the raw commit object (git log --format=%B adds a
trailing newline, which breaks sha equality).
Usage: PYTHONIOENCODING=utf-8 python push_via_api.py
"""
import base64, datetime, json, os, re, subprocess, sys, urllib.request, urllib.error

TOKEN = subprocess.check_output(['gh', 'auth', 'token']).decode().strip()
API = 'https://api.github.com'

def api(method, path, payload=None):
    req = urllib.request.Request(API + path, method=method,
        data=json.dumps(payload).encode() if payload is not None else None,
        headers={'Authorization': f'Bearer {TOKEN}',
                 'Accept': 'application/vnd.github+json',
                 'Content-Type': 'application/json'})
    with urllib.request.urlopen(req, timeout=120) as resp:
        return json.loads(resp.read().decode())

def git(repo_dir, *args):
    return subprocess.check_output(['git', '-C', repo_dir] + list(args))

def commit_meta(repo_dir, sha):
    raw = git(repo_dir, 'cat-file', '-p', sha)
    header, _, msg = raw.partition(b'\n\n')
    lines = header.decode().split('\n')
    tree = lines[0].split()[1]
    parents = [l.split()[1] for l in lines if l.startswith('parent ')]
    def who(prefix):
        l = next(l for l in lines if l.startswith(prefix))
        m = re.match(r'(\w+) (.*) <(.*)> (\d+) ([+-]\d{4})', l)
        name, email, epoch, tz = m.group(2), m.group(3), int(m.group(4)), m.group(5)
        sign = 1 if tz[0] == '+' else -1
        off = datetime.timedelta(hours=int(tz[1:3]), minutes=int(tz[3:5])) * sign
        dt = datetime.datetime.fromtimestamp(epoch, datetime.timezone(off))
        return {'name': name, 'email': email, 'date': dt.isoformat()}
    return {
        'tree': tree, 'parents': parents, 'message': msg.decode(),
        'author': who('author '), 'committer': who('committer '),
    }

def remote_sha(repo, branch='main'):
    try:
        return api('GET', f'/repos/{repo}/git/refs/heads/{branch}')['object']['sha']
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return None
        raise

def commit_paths(repo_dir, commit_sha):
    changed, deleted = [], []
    for line in git(repo_dir, 'diff-tree', '-r', '--name-status', commit_sha).decode().splitlines():
        status, _, path = line.partition('\t')
        if status in ('A', 'M', 'T'):
            changed.append(path)
        elif status == 'D':
            deleted.append(path)
    return changed, deleted

def push_repo(repo, repo_dir, commit_sha, changed_paths, deleted_paths=()):
    print(f'=== {repo} {commit_sha[:7]} ===')
    meta = commit_meta(repo_dir, commit_sha)
    tree_sha = meta['tree']

    if changed_paths or deleted_paths:
        base_tree = git(repo_dir, 'rev-parse', f'{meta["parents"][0]}^{{tree}}').decode().strip()
        if changed_paths:
            entries = []
            for path in changed_paths:
                blob_sha = git(repo_dir, 'rev-parse', f'{commit_sha}:{path}').decode().strip()
                content = git(repo_dir, 'cat-file', 'blob', f'{commit_sha}:{path}')
                got = api('POST', f'/repos/{repo}/git/blobs',
                          {'content': base64.b64encode(content).decode(), 'encoding': 'base64'})['sha']
                assert got == blob_sha, f'blob sha mismatch {path}: {got} != {blob_sha}'
                print(f'  blob {path}: {blob_sha[:10]} ok')
                entries.append({'path': path, 'mode': '100644', 'type': 'blob', 'sha': blob_sha})
            tree_sha = api('POST', f'/repos/{repo}/git/trees',
                           {'base_tree': base_tree, 'tree': entries})['sha']
            print(f'  tree (blobs): {tree_sha[:10]} ok')
            base_tree = tree_sha
        if deleted_paths:
            # separate round: mixing add+delete entries in one create-tree
            # call makes GitRPC fail with 422 BadObjectState
            del_entries = [{'path': p, 'mode': '100644', 'type': 'blob', 'sha': None}
                           for p in deleted_paths]
            tree_sha = api('POST', f'/repos/{repo}/git/trees',
                           {'base_tree': base_tree, 'tree': del_entries})['sha']
            print(f'  tree (deletes): {tree_sha[:10]} ok')
        assert tree_sha == meta['tree'], f'tree sha mismatch: {tree_sha} != {meta["tree"]}'

    new_sha = api('POST', f'/repos/{repo}/git/commits', {
        'message': meta['message'], 'tree': tree_sha,
        'parents': meta['parents'],
        'author': meta['author'], 'committer': meta['committer'],
    })['sha']
    assert new_sha == commit_sha, f'commit sha mismatch: {new_sha} != {commit_sha}'
    print(f'  commit: {new_sha[:10]} ok')
    return new_sha


def remote_has(repo, sha):
    try:
        api('GET', f'/repos/{repo}/git/commits/{sha}')
        return True
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return False
        raise


def commits_to_push(repo, d, head):
    """Walk back from head until a commit already on the remote is found."""
    todo, sha = [], head
    while not remote_has(repo, sha):
        todo.append(sha)
        parents = commit_meta(d, sha)['parents']
        if not parents:
            return todo[::-1], None
        sha = parents[0]
    return todo[::-1], sha

if __name__ == '__main__':
    base = r'D:\MGF databases\Target tracker'
    jobs = [
        ('MGFPKU/target_dataset', base + r'\target_dataset'),
        ('MGFPKU/target_visualization', base + r'\target_visualization'),
        ('MGFPKU/target_table', base + r'\target_table'),
    ]
    for repo, d in jobs:
        try:
            if not os.path.isdir(os.path.join(d, '.git')):
                print(f'=== {repo} skipped (no local repo at {d})')
                continue
            head = git(d, 'rev-parse', 'HEAD').decode().strip()
            rsha = remote_sha(repo)
            if rsha == head:
                print(f'=== {repo} {head[:7]} already on remote, skipping')
                continue
            todo, base = commits_to_push(repo, d, head)
            fast_fwd = base == rsha
            if rsha and not fast_fwd:
                msg = api('GET', f'/repos/{repo}/git/commits/{rsha}')['message']
                print(f'{repo}: remote main {rsha[:10]} ("{msg[:60]}") is not in '
                      f'local history; pushing {len(todo)} commit(s) would overwrite it')
                if '--force' not in sys.argv:
                    print('re-run with --force to overwrite remote main')
                    sys.exit(1)
            print(f'=== {repo}: {len(todo)} commit(s) to push')
            for sha in todo:
                push_repo(repo, d, sha, *commit_paths(d, sha))
            if rsha is None:
                api('POST', f'/repos/{repo}/git/refs',
                    {'ref': 'refs/heads/main', 'sha': head})
            else:
                api('PATCH', f'/repos/{repo}/git/refs/heads/main',
                    {'sha': head, 'force': not fast_fwd})
            print(f'  ref main -> {head[:10]}'
                  f'{"" if fast_fwd else " (forced)"}')
        except urllib.error.HTTPError as e:
            print(f'FAIL {repo}: HTTP {e.code} {e.read().decode()[:500]}')
            sys.exit(1)
        except Exception as e:
            print(f'FAIL {repo}: {e}')
            sys.exit(1)
    print('ALL PUSHED')
