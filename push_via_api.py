# -*- coding: utf-8 -*-
"""Push local commits to GitHub via the git database API (github.com:443 blocked).

Replicates tree/blob/commit objects exactly, so remote SHAs equal local SHAs.
GitHub keeps timezone offsets and message bytes as sent; the message must be
extracted verbatim from the raw commit object (git log --format=%B adds a
trailing newline, which breaks sha equality).
Usage: PYTHONIOENCODING=utf-8 python push_via_api.py
"""
import base64, datetime, json, re, subprocess, sys, urllib.request, urllib.error

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

def push_repo(repo, repo_dir, changed_paths):
    commit_sha = git(repo_dir, 'rev-parse', 'HEAD').decode().strip()
    print(f'=== {repo} {commit_sha[:7]} ===')
    meta = commit_meta(repo_dir, commit_sha)
    tree_sha = meta['tree']

    if changed_paths:
        entries = []
        base_tree = git(repo_dir, 'rev-parse', f'{meta["parents"][0]}^{{tree}}').decode().strip()
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
        assert tree_sha == meta['tree'], f'tree sha mismatch: {tree_sha} != {meta["tree"]}'
        print(f'  tree: {tree_sha[:10]} ok')

    new_sha = api('POST', f'/repos/{repo}/git/commits', {
        'message': meta['message'], 'tree': tree_sha,
        'parents': meta['parents'],
        'author': meta['author'], 'committer': meta['committer'],
    })['sha']
    assert new_sha == commit_sha, f'commit sha mismatch: {new_sha} != {commit_sha}'
    print(f'  commit: {new_sha[:10]} ok')

    api('PATCH', f'/repos/{repo}/git/refs/heads/main', {'sha': new_sha, 'force': False})
    print(f'  ref main -> {new_sha[:10]}')
    return new_sha

if __name__ == '__main__':
    base = r'D:\MGF databases\Target tracker'
    jobs = [
        ('MGFPKU/target_dataset', base + r'\target_dataset',
         ['Targets_cn.xlsx', 'fix_direction2_cn.py']),
        ('MGFPKU/target_visualization', base + r'\target_visualization', []),
        ('MGFPKU/target_table', base + r'\target_table', []),
    ]
    for repo, d, paths in jobs:
        try:
            push_repo(repo, d, paths)
        except urllib.error.HTTPError as e:
            print(f'FAIL {repo}: HTTP {e.code} {e.read().decode()[:500]}')
            sys.exit(1)
        except Exception as e:
            print(f'FAIL {repo}: {e}')
            sys.exit(1)
    print('ALL PUSHED')
