# 部署指南

本文面向第一次部署 WordReplace 的用户，覆盖 Docker Compose、群晖 Container Manager、反向代理 HTTPS、权限、健康检查、备份恢复、升级回滚和资源限制。

> 返回 [README](../README.md) · 快速上手看 [GETTING_STARTED.md](../GETTING_STARTED.md) · 遇到问题看 [TROUBLESHOOTING.md](TROUBLESHOOTING.md)

---

## 1. 准备工作

需要：

- 一台安装了 Docker 和 Docker Compose v2 的机器（Linux、NAS、Windows + WSL 均可）。
- 一个可用的端口，默认使用 `12344`。
- 如果要从公网访问，需要一个域名，并建议配置 HTTPS。

检查环境：

```bash
docker --version
docker compose version
```

确认 `docker compose version` 有输出即可。本文所有命令都用 `docker compose`（中间是空格），不再使用旧的 `docker-compose`。

---

## 2. Docker Compose 部署（推荐）

### 2.1 目录准备

```bash
mkdir -p wordreplace/data
cd wordreplace
```

### 2.2 新建 `docker-compose.yml`

```yaml
services:
  wordreplace:
    image: ghcr.io/marod1m/wordreplace:3.0.0
    container_name: wordreplace
    restart: unless-stopped
    ports:
      - "12344:8000"
    volumes:
      - ./data:/app/data
    environment:
      CORS_ALLOW_ORIGINS: "http://localhost:12344"
      ALLOWED_HOSTS: "*"
      MAX_UPLOAD_SIZE_MB: "50"
      MAX_EXCEL_ROWS: "20000"
      MAX_REPLACE_RULES: "500"
      RUN_CACHE_TTL_SECONDS: "1800"
      RUN_CACHE_MAX_ENTRIES: "200"
      RUN_CACHE_MAX_BYTES: "536870912"
      MAX_CONCURRENT_RUNS: "1"
      EXPORT_TOKEN_SECRET: "请改成足够长的随机字符串"
    mem_limit: 2g
    cpus: "1.5"
```

要点：

- 左边端口 `12344` 可改，右边 `8000` 不要改。
- `./data:/app/data` 用来保存 SQLite 数据库，建议保留。
- `EXPORT_TOKEN_SECRET` 必须替换成随机字符串，不要直接使用示例值。
- 生产环境建议固定镜像标签（本示例为 `3.0.0`；GitHub 发布标签为 `v3.0.0`），不要长期使用 `latest`。

生成随机密钥的示例：

```bash
openssl rand -hex 32
```

把输出粘贴到 `EXPORT_TOKEN_SECRET`。

### 2.3 启动

```bash
docker compose up -d
```

查看状态：

```bash
docker compose ps
docker compose logs -f --tail=100
```

看到日志中出现 `WordReplace 3.0.0 started` 即启动成功。

### 2.4 验证

```bash
curl -fsS http://localhost:12344/health
curl -fsS http://localhost:12344/health/ready
curl -fsS http://localhost:12344/config
```

预期分别返回：

```json
{"status":"ok"}
{"status":"ready"}
```

`/config` 会返回版本号和各项限制。

浏览器访问 `http://localhost:12344`，页面底部应显示 `v3.0.0`。

### 2.5 使用本地源码构建

如果还没有可用镜像，或者要自己改代码后构建：

```bash
docker compose up -d --build
```

构建过程会先构建 Next.js 前端，再打包 Python 运行时。首次构建时间较长属于正常现象。

---

## 3. 群晖 Container Manager 部署

### 3.1 创建目录

在 File Station 中创建：

```text
/volume1/docker/wordreplace/data
```

或者通过 SSH：

```bash
sudo mkdir -p /volume1/docker/wordreplace/data
```

### 3.2 用“项目”方式部署

1. 打开 **Container Manager**。
2. 进入 **项目** → **新增**。
3. 项目名称填 `wordreplace`。
4. 路径选择 `/volume1/docker/wordreplace`。
5. 选择“创建 docker-compose.yml”，粘贴：

```yaml
services:
  wordreplace:
    image: ghcr.io/marod1m/wordreplace:3.0.0
    container_name: wordreplace
    restart: unless-stopped
    ports:
      - "12344:8000"
    volumes:
      - /volume1/docker/wordreplace/data:/app/data
    environment:
      CORS_ALLOW_ORIGINS: "http://你的群晖IP:12344"
      ALLOWED_HOSTS: "*"
      MAX_UPLOAD_SIZE_MB: "50"
      MAX_EXCEL_ROWS: "20000"
      MAX_REPLACE_RULES: "500"
      RUN_CACHE_TTL_SECONDS: "1800"
      RUN_CACHE_MAX_ENTRIES: "200"
      RUN_CACHE_MAX_BYTES: "536870912"
      MAX_CONCURRENT_RUNS: "1"
      EXPORT_TOKEN_SECRET: "请改成足够长的随机字符串"
```

6. 点击完成并启动。
7. 浏览器访问 `http://你的群晖IP:12344`。

### 3.3 权限说明

应用进程以 UID `10001` 运行。镜像入口脚本在容器以 root 启动时会先修正 `/app/data` 权限，再切换到 `appuser`；如果仍然报权限错误，请在群晖 SSH 中执行：

```bash
sudo chown -R 10001:10001 /volume1/docker/wordreplace/data
```

### 3.4 群晖反向代理

如果需要通过群晖的“反向代理服务器”提供 HTTPS：

1. 打开 **控制面板 → 登录门户 → 高级 → 反向代理服务器**。
2. 新增规则：
   - 来源：`HTTPS`，主机名 `word.example.com`，端口 `443`。
   - 目的地：`HTTP`，主机名 `127.0.0.1`，端口 `12344`。
3. 保存后，把 `CORS_ALLOW_ORIGINS` 改成 `https://word.example.com`，把 `ALLOWED_HOSTS` 改成 `word.example.com`，然后重启容器。

---

## 4. 反向代理与 HTTPS

### 4.1 部署前的环境变量

假设最终访问地址是 `https://word.example.com`，建议设置：

```yaml
    environment:
      CORS_ALLOW_ORIGINS: "https://word.example.com"
      ALLOWED_HOSTS: "word.example.com"
      EXPORT_TOKEN_SECRET: "你的强随机密钥"
```

如果通过反向代理，应用仍然监听容器内的 `8000`，宿主机可以只绑定本机：

```yaml
    ports:
      - "127.0.0.1:12344:8000"
```

这样外网只能经过反向代理访问，不能直接访问 `12344`。

### 4.2 Caddy 示例

Caddy 会自动申请和续期 HTTPS 证书。安装 Caddy 后编辑 `/etc/caddy/Caddyfile`：

```caddyfile
word.example.com {
    encode gzip
    reverse_proxy 127.0.0.1:12344
}
```

重载：

```bash
sudo systemctl reload caddy
```

Caddy 默认不会限制上传大小，通常无需额外配置。如果前面还有其他代理，请确认请求体上限大于 `MAX_UPLOAD_SIZE_MB`。

### 4.3 Nginx 示例（HTTP）

```nginx
server {
    listen 80;
    server_name word.example.com;

    client_max_body_size 100m;

    location / {
        proxy_pass http://127.0.0.1:12344;
        proxy_http_version 1.1;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
        proxy_read_timeout 600s;
    }
}
```

保存后检查并重载：

```bash
sudo nginx -t
sudo systemctl reload nginx
```

### 4.4 Nginx 示例（HTTPS）

先准备证书（也可以用 Certbot 自动签发），然后：

```nginx
server {
    listen 80;
    server_name word.example.com;
    return 301 https://$host$request_uri;
}

server {
    listen 443 ssl;
    server_name word.example.com;

    ssl_certificate     /etc/letsencrypt/live/word.example.com/fullchain.pem;
    ssl_certificate_key /etc/letsencrypt/live/word.example.com/privkey.pem;

    client_max_body_size 100m;

    location / {
        proxy_pass http://127.0.0.1:12344;
        proxy_http_version 1.1;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
        proxy_read_timeout 600s;
    }
}
```

使用 Certbot 申请证书的常见命令：

```bash
sudo certbot --nginx -d word.example.com
```

> 反向代理的 `client_max_body_size` 要大于或等于 `MAX_UPLOAD_SIZE_MB`，否则请求还没到应用就会被代理返回 413。

---

## 5. 非 root 运行与数据权限

镜像创建了 `appuser`，UID 为 `10001`。入口脚本逻辑：

- 如果容器以 root 启动：先对 `/app/data` 执行 `chown`，然后用 `gosu appuser` 切换到非 root 用户运行应用。
- 如果容器已经以非 root 启动：直接运行应用。

检查当前用户：

```bash
docker exec wordreplace id
```

预期类似：

```text
uid=10001(appuser) gid=10001(appuser) groups=10001(appuser)
```

如果宿主机绑定目录权限不对，执行：

```bash
sudo chown -R 10001:10001 ./data
```

SELinux 开启的系统（如部分 CentOS / Fedora）可能需要在挂载后加 `:Z`：

```yaml
    volumes:
      - ./data:/app/data:Z
```

或者临时放行：

```bash
sudo chcon -Rt container_file_t ./data
```

---

## 6. 健康检查

镜像内置健康检查：

```dockerfile
HEALTHCHECK --interval=30s --timeout=5s --start-period=20s --retries=3 \
    CMD python -c "import urllib.request; urllib.request.urlopen('http://127.0.0.1:8000/health', timeout=3)"
```

含义：

- 每 30 秒检查一次。
- 单次超时 5 秒。
- 启动后前 20 秒不计失败。
- 连续失败 3 次标记为 unhealthy。

查看：

```bash
docker inspect --format '{{.State.Health.Status}}' wordreplace
```

如果 `/health` 正常但 `/health/ready` 返回 503，通常是数据库不可用或数据目录不可写。查看日志：

```bash
docker compose logs --tail=200 wordreplace
```

需要自定义健康检查时，可以在 compose 中覆盖：

```yaml
    healthcheck:
      test: ["CMD", "python", "-c", "import urllib.request; urllib.request.urlopen('http://127.0.0.1:8000/health', timeout=3)"]
      interval: 30s
      timeout: 5s
      retries: 3
      start_period: 20s
```

---

## 7. 备份与恢复

### 7.1 备份

数据库是 SQLite，建议停止服务后复制，避免写入过程中文件不一致。

```bash
docker compose stop
cp -r ./data "./data.backup-$(date +%Y%m%d-%H%M%S)"
docker compose start
```

也可以打包：

```bash
docker compose stop
tar -czf "wordreplace-data-$(date +%Y%m%d-%H%M%S).tar.gz" ./data
docker compose start
```

也可以使用项目自带脚本（会备份到 `backups/` 目录）：

```bash
./scripts/backup.sh
```

注意：替换结果保存在内存中，不需要备份，也无法通过备份恢复。需要保留结果时，请在结果过期前下载。

### 7.2 恢复

```bash
docker compose stop
cp -r ./data.backup-你的时间戳/. ./data/
docker compose start
```

如果使用 tar 包：

```bash
docker compose stop
tar -xzf wordreplace-data-你的时间戳.tar.gz
docker compose start
```

恢复后检查：

```bash
curl -fsS http://localhost:12344/health/ready
```

---

## 8. 升级与回滚

### 8.1 升级前

1. 记录当前镜像标签。
2. 备份 `./data`。
3. 在低峰时间操作。

### 8.2 升级到指定版本

修改 compose：

```yaml
    image: ghcr.io/marod1m/wordreplace:3.0.0
```

执行：

```bash
docker compose pull
docker compose up -d
```

### 8.3 升级到最新版

```bash
docker compose pull
docker compose up -d
```

> 不建议生产环境长期跟随 `latest`，因为无法预知下一次拉取会带来什么变化。

### 8.4 回滚

1. 把镜像标签改回旧版本。
2. 执行：

```bash
docker compose pull
docker compose up -d
```

3. 如果新版本改动了数据库结构，还需要恢复升级前的 `./data` 备份。
4. 验证：

```bash
curl -fsS http://localhost:12344/health
curl -fsS http://localhost:12344/health/ready
```

更完整的升级步骤见 [UPGRADE.md](UPGRADE.md)。

---

## 9. 资源限制建议

个人使用的最小建议：

- 内存：至少 1GB，推荐 2GB。
- CPU：1 核可用，2 核更流畅。
- 磁盘：镜像本身约几百 MB，数据目录很小；主要取决于系统 Docker 占用。

Compose 示例：

```yaml
    mem_limit: 2g
    cpus: "1.5"
```

注意：

- 结果缓存默认最多占 512MB（`RUN_CACHE_MAX_BYTES`），处理大文档时还会额外占用内存。
- `mem_limit` 建议大于缓存上限并留出至少 512MB 余量。
- `MAX_CONCURRENT_RUNS` 默认是 1。调大后会同时占用更多内存和 CPU，不建议在小内存机器上调大。
- 磁盘空间主要用于镜像和 SQLite 数据库，规则模板数据通常很小。

查看资源占用：

```bash
docker stats wordreplace
```

---

## 10. 停止与卸载

停止：

```bash
docker compose down
```

停止并删除本项目构建的镜像（可选）：

```bash
docker compose down --rmi local
```

数据目录不会自动删除。确认不再需要后，手动删除：

```bash
rm -rf ./data
```

删除前请务必备份。
