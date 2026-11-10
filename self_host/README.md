# 中国气候目标追踪器 — 自托管部署（替代 Posit Connect Cloud）

两个 Shiny 网页 app（target_table 数据表格、target_visualization 可视化）脱离 Posit Connect Cloud，
部署到国内轻量 VPS，通过 Cloudflare Tunnel 提供公网 HTTPS（免备案、免开放端口）。

更新链路：本地推送 GitHub → VPS 每 5 分钟 `git pull` 三个仓库 → 有变更自动重启 app。
数据集（xlsx）在 target_dataset 仓库内，app 以 LOCAL_DATA 模式直接读本地文件，不依赖 GitHub API。

## 一、VPS 选购清单

| 项 | 推荐 | 说明 |
|---|---|---|
| 厂商 | 腾讯云轻量应用服务器 或 阿里云轻量 | 二选一即可 |
| 配置 | **2 核 4G**（两个 app 在 Posit 上各配 4G 属严重超配，数据集仅 ~300KB） | 2C2G 也能跑，4G 更从容 |
| 地域 | 上海 / 广州 | 国内访问最优 |
| 镜像 | Ubuntu 22.04 或 24.04 LTS | 系统镜像直接选 |
| 磁盘 | 60GB SSD 默认即可 | 仓库总量 < 100MB |
| 带宽 | 4-5 Mbps 峰值即可 | 页面数据量很小 |
| 价格 | 新用户优惠约 ¥50-70/月（标准价 ¥100-112/月，包年更便宜） | 对比 Posit Enhanced $59/月 |
| 安全组 | **无需开放任何入站端口**（Tunnel 为出站连接） | 保持默认 |

域名（必需，用于 Tunnel 的 HTTPS 域名）：

- 已有域名 → 加入 Cloudflare（免费账号即可，NS 托管到 Cloudflare）
- 没有域名 → 腾讯云/阿里云注册 .com（~¥60-80/年，需实名认证），再托管到 Cloudflare
- Cloudflare 免费计划完全够用；**无需 ICP 备案**（访客经 Cloudflare 边缘访问）

## 二、部署步骤（约 15 分钟）

1. **Cloudflare 控制台建隧道**：
   登录 dash.cloudflare.com → Zero Trust → Networks → Tunnels → Create a tunnel
   （类型选 Cloudflared）→ 复制 **token** → 暂不配置 public hostname 也可以，稍后补。

2. **VPS 上运行初始化脚本**：
   ```bash
   git clone https://github.com/MGFPKU/target_dataset.git /tmp/td   # 若 github 直连慢：git clone https://ghfast.top/https://github.com/MGFPKU/target_dataset.git /tmp/td
   sudo bash /tmp/td/self_host/setup_vps.sh <TUNNEL_TOKEN>
   ```

3. **配置域名路由**（Zero Trust 控制台 → 你的隧道 → Public Hostname）：
   - `table.<你的域名>` → `http://localhost:8001`
   - `viz.<你的域名>` → `http://localhost:8002`
   （保存后 Cloudflare 自动建 DNS 记录，1-2 分钟生效）

4. **验证**：
   ```bash
   curl http://127.0.0.1:8001/?lang=cn   # 200 即本地服务正常
   curl http://127.0.0.1:8002/?lang=en
   ```
   然后浏览器访问 `https://table.<你的域名>/?lang=cn` 和 `https://viz.<你的域名>/`。

## 三、日常更新流程（与以前相同，但更简单）

数据集或 app 代码更新后照常推送 GitHub，**不再需要 publish_and_sync.py**：
VPS 每 5 分钟自动拉取并重启（日志 `/var/log/climate-tracker-update.log`）。
想立即生效可在 VPS 上手动执行 `sudo /usr/local/bin/climate-update.sh`。

GitHub release 资产上传（`gh release upload 6 ... --clobber`）可保留作为开放数据分发渠道，
但发布脚本中 Posit 重部署/API 验证部分（第 3-5 步）不再需要。

## 四、运维速查

```bash
systemctl status climate-table climate-viz cloudflared   # 状态
systemctl restart climate-table                          # 手动重启
journalctl -u climate-table -n 50                        # app 日志
tail -f /var/log/climate-tracker-update.log              # 更新日志
```

- GitHub 直连不稳时，clone/pull 已默认走 `https://gh-proxy.com/` 镜像（origin 即镜像地址）；镜像失效时可临时换 `https://ghfast.top/` 或直连
- PyPI 依赖经清华镜像（已在 systemd 单元配置 UV_DEFAULT_INDEX），下载慢可换 `https://mirrors.aliyun.com/pypi/simple`
- 国内访客经 Cloudflare 海外边缘（100-300ms，可用）。若以后要求国内毫秒级访问，可加备案后增加境内直连线路（见下）

## 五、已知取舍

- **国内访问速度**：Cloudflare 免费版无中国边缘节点，国内访客走海外边缘，速度一般但稳定可用。
  要最优体验需 ICP 备案 + 分线路解析（境内直连 VPS、境外走 Cloudflare），备案约 1-3 周，
  可后续升级（DNS 分线路：DNSPod 免费支持）。
- **VPS 单点**：无多副本/自动扩缩容。此项目流量小、数据小，单机足够；轻量 VPS 支持快照备份。
