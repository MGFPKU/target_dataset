# -*- coding: utf-8 -*-
"""发布同步流程（每次数据集提交并推送后必须执行，确保网页端数据同步）：

1. 上传 release 资产到 tag 6（--clobber）：China_Climate_Target_Tracker_cn/en.xlsx，
   --with-sources 时附 Data_sources.zip（Data_sources/ 目录有变更才加）
2. 刷新 release 标题（YYYYMMDD）与 notes（英文，--notes-file 指定）
3. 向 target_table / target_visualization 两个 app 仓库推空提交，
   触发 Posit Connect Cloud 的 GitHub 连接自动重部署（进程重启后重新拉取 release 资产）
4. 轮询 Connect Cloud API 验证 current_revision.commit_sha 与推送提交一致且 status=published
5. 本地端到端验证：用 target_table 的 data.fetch_raw_data() 拉取最新资产并输出行数

用法（用系统 Python 运行，rsconnect.api 在系统 Python 里）：
  python publish_and_sync.py --notes-file notes.txt [--with-sources] [--verify-rows]
"""
import argparse
import datetime
import json
import os
import re
import subprocess
import sys
import time

TAG = '6'
ASSETS = ['China_Climate_Target_Tracker_cn.xlsx', 'China_Climate_Target_Tracker_en.xlsx']
APP_REPOS = [
    # (仓库路径, Connect Cloud content GUID)
    ('../target_table', '019cb37e-3fbc-652e-3fb8-5d4c5126bba1'),
    ('../target_visualization', '019cc515-6617-e2ff-d001-fea7a50074e1'),
]
SERVERS_JSON = os.path.expanduser('~/AppData/Roaming/rsconnect-python/servers.json')


def run(cmd, cwd=None):
    print('+', ' '.join(cmd) if isinstance(cmd, list) else cmd)
    r = subprocess.run(cmd, cwd=cwd, capture_output=True, text=True)
    if r.returncode != 0:
        print(r.stdout, r.stderr)
        sys.exit(f'FAILED: {cmd}')
    return r.stdout


def upload_assets(with_sources):
    assets = list(ASSETS)
    if with_sources:
        assets.append('Data_sources.zip')
    run(['gh', 'release', 'upload', TAG] + assets + ['--clobber'])


def edit_release(title, notes_file):
    run(['gh', 'release', 'edit', TAG, '--title', title, '--notes-file', notes_file])


def trigger_app_redeploys(title):
    msg = f'Trigger redeploy to pick up updated dataset release ({title})'
    for path, _ in APP_REPOS:
        run(['git', 'commit', '--allow-empty', '-m', msg], cwd=path)
        out = run(['git', 'push'], cwd=path)
        m = re.search(r'(\w+)\.\.(\w+)\s+\S+ ->', out)
        if not m:
            sys.exit('无法从 push 输出解析新提交 sha: ' + out)
        yield path, m.group(2)


def get_cloud_client():
    sys.path.insert(0, 'D:/Python/Lib/site-packages')
    from rsconnect.api import ConnectCloudServer, ConnectCloudClient
    cfg = json.load(open(SERVERS_JSON))['posit']
    server = ConnectCloudServer(account_name='mgflab', access_token=cfg.get('access_token'),
                                refresh_token=cfg.get('refresh_token'), url=cfg['url'],
                                account_id=cfg.get('account_id'))
    return ConnectCloudClient(server)


def wait_published(expected):
    """expected: {guid: sha}，轮询至 current_revision 匹配且 published。"""
    client = get_cloud_client()
    pending = dict(expected)
    deadline = time.time() + 420
    while pending and time.time() < deadline:
        for guid, sha in list(pending.items()):
            rev = client.get_content(guid).get('current_revision') or {}
            if rev.get('commit_sha') == sha and rev.get('status') == 'published':
                print(f'  {guid} -> {sha[:7]} published')
                del pending[guid]
        if pending:
            time.sleep(10)
    if pending:
        sys.exit('验证超时，未确认发布: ' + str(pending))


def verify_rows():
    """用 target_table 自身的 data.fetch_raw_data() 拉最新 release 资产并数行数。"""
    out = run(['uv', 'run', '--with', 'openpyxl', 'python', '-X', 'utf8', '-c', (
        "import openpyxl, data; "
        "buf = data.fetch_raw_data('CN'); "
        "wb = openpyxl.load_workbook(buf, read_only=True); "
        "skip = {'说明','概览','来源'}; "
        "total = sum(ws.max_row-1 for ws in wb.worksheets if ws.title not in skip); "
        "print('release CN rows:', total)")],
        cwd='../target_table')
    print(out.strip())


def main():
    ap = argparse.ArgumentParser(description='数据集发布同步：release 资产 + 英文 notes + 两 app 重部署 + API 验证')
    ap.add_argument('--notes-file', required=True, help='release notes 文本文件（英文）')
    ap.add_argument('--title', default=datetime.date.today().strftime('%Y%m%d'),
                    help='release 标题，默认今天 YYYYMMDD')
    ap.add_argument('--with-sources', action='store_true',
                    help='Data_sources/ 有变更时加传 Data_sources.zip')
    ap.add_argument('--verify-rows', action='store_true', help='本地拉最新资产验证行数')
    args = ap.parse_args()

    print(f'== 1/4 上传 release 资产到 tag {TAG}')
    upload_assets(args.with_sources)
    print(f'== 2/4 刷新标题/notes: {args.title}')
    edit_release(args.title, args.notes_file)
    print('== 3/4 推空提交触发两 app 重部署')
    expected = {}
    for path, sha in trigger_app_redeploys(args.title):
        expected[path] = sha
        print(f'  {path}: pushed {sha[:7]}')
    guid_map = dict(APP_REPOS)
    print('== 4/4 API 验证发布状态')
    wait_published({guid_map[path]: sha for path, sha in expected.items()})
    if args.verify_rows:
        print('== 附加: 本地端到端行数验证')
        verify_rows()
    print('done')


if __name__ == '__main__':
    main()
