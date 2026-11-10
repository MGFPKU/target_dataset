#!/usr/bin/env bash
# 中国气候目标追踪器 - VPS 一键初始化（Ubuntu 22.04/24.04, root 运行）
#
# 用法:
#   sudo bash setup_vps.sh [CLOUDFLARED_TUNNEL_TOKEN]
#
# 不带 token 时跳过 cloudflared 安装提示；token 在 Cloudflare Zero Trust
# 控制台创建 tunnel 后获得（README.md 有步骤）。
set -euo pipefail

BASE=/opt/climate-tracker
SELF="$BASE/target_dataset/self_host"
GH_MIRROR="https://ghfast.top"

if [ "$(id -u)" -ne 0 ]; then echo "请用 root 运行: sudo bash $0"; exit 1; fi
TUNNEL_TOKEN="${1:-}"

echo "== 1/7 安装基础依赖"
apt-get update -qq
apt-get install -y -qq git curl ca-certificates

echo "== 2/7 安装 uv"
if [ ! -x /usr/local/bin/uv ]; then
  if curl -fsSL --connect-timeout 15 https://astral.sh/uv/install.sh | env UV_INSTALL_DIR=/usr/local/bin sh; then
    echo "uv installed from astral.sh"
  else
    echo "astral.sh 不可达，改用清华 pip 镜像安装"
    apt-get install -y -qq python3-pip
    pip3 install -q -i https://pypi.tuna.tsinghua.edu.cn/simple uv --break-system-packages 2>/dev/null \
      || pip3 install -q -i https://pypi.tuna.tsinghua.edu.cn/simple uv
    ln -sf "$(command -v uv)" /usr/local/bin/uv
  fi
fi
uv --version

echo "== 3/7 创建 tracker 用户并克隆仓库"
id tracker >/dev/null 2>&1 || useradd -m -s /bin/bash tracker
mkdir -p "$BASE"
chown tracker:tracker "$BASE"
for repo in target_dataset target_table target_visualization; do
  if [ ! -d "$BASE/$repo/.git" ]; then
    if ! git clone -q "https://github.com/MGFPKU/$repo.git" "$BASE/$repo"; then
      echo "github 直连失败，尝试镜像 ghfast.top"
      git clone -q "$GH_MIRROR/https://github.com/MGFPKU/$repo.git" "$BASE/$repo"
    fi
  fi
  cd "$BASE/$repo"
  git remote set-url origin "https://github.com/MGFPKU/$repo.git" 2>/dev/null || true
  chown -R tracker:tracker "$BASE/$repo"
done

echo "== 4/7 预构建 uv 依赖环境（首次启动免等待）"
sudo -u tracker env UV_DEFAULT_INDEX=https://pypi.tuna.tsinghua.edu.cn/simple \
  sh -c 'cd /opt/climate-tracker/target_table && uv sync -q && cd /opt/climate-tracker/target_visualization && uv sync -q'

echo "== 5/7 安装 systemd 服务"
cp "$SELF/systemd/climate-table.service" /etc/systemd/system/
cp "$SELF/systemd/climate-viz.service" /etc/systemd/system/
systemctl daemon-reload
systemctl enable --now climate-table climate-viz

echo "== 6/7 安装自动更新（每 5 分钟 git pull + 变更重启）"
install -m 755 "$SELF/update.sh" /usr/local/bin/climate-update.sh
( crontab -l 2>/dev/null | grep -v climate-update.sh; \
  echo '*/5 * * * * /usr/local/bin/climate-update.sh >> /var/log/climate-tracker-update.log 2>&1' ) | crontab -

echo "== 7/7 安装 cloudflared 并配置隧道"
if ! command -v cloudflared >/dev/null 2>&1; then
  if curl -fsSL --connect-timeout 15 -o /tmp/cloudflared.deb \
     https://pkg.cloudflare.com/cloudflared-stable-linux-amd64.deb && dpkg -i /tmp/cloudflared.deb; then
    echo "cloudflared installed from pkg.cloudflare.com"
  else
    echo "改用 GitHub 镜像下载 cloudflared"
    curl -fsSL -o /tmp/cloudflared.tgz \
      "$GH_MIRROR/https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-linux-amd64.tgz"
    tar -xzf /tmp/cloudflared.tgz -C /usr/local/bin cloudflared
  fi
fi
cloudflared --version

if [ -n "$TUNNEL_TOKEN" ]; then
  cloudflared service install "$TUNNEL_TOKEN"
  systemctl enable --now cloudflared
  echo "隧道已用 token 安装并启动。请在 Cloudflare Zero Trust 控制台为隧道添加两个 public hostname:"
  echo "  table.<你的域名> -> http://localhost:8001"
  echo "  viz.<你的域名>   -> http://localhost:8002"
else
  echo "未提供 tunnel token。稍后执行:"
  echo "  cloudflared service install <TOKEN> && systemctl enable --now cloudflared"
  echo "并在 Zero Trust 控制台添加上述两个 public hostname。"
fi

echo
echo "完成。检查状态: systemctl status climate-table climate-viz cloudflared"
echo "本地验证: curl http://127.0.0.1:8001/?lang=cn 和 http://127.0.0.1:8002/?lang=cn"
