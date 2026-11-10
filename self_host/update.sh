#!/usr/bin/env bash
# 自托管自动更新：拉取三个仓库，任一有变更则重启两个 app 服务。
# 由 cron 每 5 分钟调用一次（见 setup_vps.sh）。
set -u
BASE=/opt/climate-tracker
CHANGED=0

for repo in target_dataset target_table target_visualization; do
  cd "$BASE/$repo" || { echo "[$(date -Is)] MISSING $repo" >> /var/log/climate-tracker-update.log; continue; }
  git fetch origin main -q
  LOCAL=$(git rev-parse HEAD)
  REMOTE=$(git rev-parse origin/main)
  if [ "$LOCAL" != "$REMOTE" ]; then
    if ! git pull --ff-only origin main -q 2>/dev/null; then
      # 上游发生过 force-push（push_via_api 可能产生非快进历史），部署机直接对齐远程
      git reset --hard "$REMOTE"
    fi
    echo "[$(date -Is)] $repo updated $LOCAL -> $REMOTE" >> /var/log/climate-tracker-update.log
    CHANGED=1
  fi
done

if [ "$CHANGED" = "1" ]; then
  systemctl restart climate-table climate-viz
  echo "[$(date -Is)] services restarted" >> /var/log/climate-tracker-update.log
fi
