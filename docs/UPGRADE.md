# 升级与回滚

本文给出从备份、升级、验证到回滚的完整命令。升级前请务必先备份 `./data`。

> 返回 [README](../README.md) · 部署细节见 [DEPLOYMENT.md](DEPLOYMENT.md) · 排障见 [TROUBLESHOOTING.md](TROUBLESHOOTING.md)

---

## 0. 升级前检查

进入 docker-compose.yml 所在目录：

```bash
cd /path/to/wordreplace
```

确认当前版本与容器状态：

```bash
cat VERSION
docker compose ps
docker inspect --format '{{.Config.Image}}' wordreplace
```

查看当前镜像标签，并记录下来，回滚时要用：

```bash
docker inspect --format '{{.Config.Image}}' wordreplace
```

确认当前服务健康：

```bash
curl -fsS http://localhost:12344/health
curl -fsS http://localhost:12344/health/ready
curl -fsS http://localhost:12344/config
```

如果当前已经不正常，请先按 [TROUBLESHOOTING.md](TROUBLESHOOTING.md) 处理，不要直接升级。

---

## 1. 备份（必须）

### 1.1 停止服务

```bash
docker compose stop
```

停止可以保证 SQLite 文件复制时没有写入。

### 1.2 复制数据目录

```bash
cp -r ./data "./data.backup-$(date +%Y%m%d-%H%M%S)"
```

或打包：

```bash
tar -czf "wordreplace-data-$(date +%Y%m%d-%H%M%S).tar.gz" ./data
```

### 1.3 记录镜像标签

编辑或查看 `docker-compose.yml`，把当前的 `image:` 行复制到文本里保存。例如：

```yaml
    image: ghcr.io/marod1m/wordreplace:v2.0.5
```

### 1.4 重新启动（如果暂时不升级）

```bash
docker compose start
```

---

## 2. 升级

### 2.1 修改镜像标签

打开 `docker-compose.yml`，把 `image:` 改成目标版本。

例如从 `v2.0.5` 升级到 `v3.0.0`：

```yaml
    image: ghcr.io/marod1m/wordreplace:v3.0.0
```

如果暂时使用最新版：

```yaml
    image: ghcr.io/marod1m/wordreplace:latest
```

> 生产环境建议固定版本号，不要长期使用 `latest`。

### 2.2 拉取并启动

```bash
docker compose pull
docker compose up -d
```

如果镜像拉取失败，检查网络和镜像仓库权限。也可以稍后重试：

```bash
docker compose pull
```

### 2.3 本地源码构建升级

如果你是自己构建镜像：

```bash
docker compose down
docker compose build --no-cache
docker compose up -d
```

---

## 3. 升级后验证

### 3.1 容器状态

```bash
docker compose ps
```

状态应为 `Up`，健康检查最终应为 `healthy`：

```bash
docker inspect --format '{{.State.Health.Status}}' wordreplace
```

### 3.2 接口检查

```bash
curl -fsS http://localhost:12344/health
curl -fsS http://localhost:12344/health/ready
curl -fsS http://localhost:12344/config
```

预期：

- `/health` 返回 `{"status":"ok"}`
- `/health/ready` 返回 `{"status":"ready"}`
- `/config` 中的 `app_version` 是目标版本

### 3.3 页面检查

1. 打开 `http://你的地址:12344`。
2. 页面底部版本号应为目标版本（例如 `v3.0.0`）。
3. 上传示例 Word 和 Excel，确认能预览。
4. 添加一条规则，执行 1-2 行，确认能下载。
5. 打开“显示高级项”，确认模板库列表仍在，旧模板可加载。

### 3.4 数据检查

确认旧规则模板没有丢失：

```bash
curl -fsS http://localhost:12344/rule-templates
```

如果返回空数组，但升级前有模板，请检查：

- 是否挂载了正确的 `./data`。
- 是否误用了新的空数据目录。
- 数据库文件是否存在：

```bash
ls -lh ./data
```

### 3.5 日志检查

```bash
docker compose logs --tail=200 wordreplace
```

重点看：

- 是否有 `EXPORT_TOKEN_SECRET is using the default value` 警告。
- 是否有数据库权限错误。
- 是否有启动失败或异常堆栈。

---

## 4. 回滚

### 4.1 回滚应用版本

1. 把 `docker-compose.yml` 中的镜像标签改回旧版本，例如：

```yaml
    image: ghcr.io/marod1m/wordreplace:v2.0.5
```

2. 拉取并启动：

```bash
docker compose pull
docker compose up -d
```

3. 验证：

```bash
curl -fsS http://localhost:12344/health
curl -fsS http://localhost:12344/health/ready
curl -fsS http://localhost:12344/config
```

### 4.2 回滚数据（需要时）

如果新版本改动了数据库结构，或者升级后数据异常，需要恢复备份：

```bash
docker compose stop
cp -r ./data.backup-你的时间戳/. ./data/
docker compose start
```

使用 tar 包时：

```bash
docker compose stop
tar -xzf wordreplace-data-你的时间戳.tar.gz
docker compose start
```

恢复后再次检查：

```bash
curl -fsS http://localhost:12344/health/ready
curl -fsS http://localhost:12344/rule-templates
```

### 4.3 完全回到升级前

按顺序执行：

```bash
docker compose down
# 改回旧镜像标签
# 恢复数据目录
docker compose up -d
```

注意：

- `docker compose down` 不会删除 `./data`。
- 不要执行 `docker compose down -v`，除非你确认要删除数据卷。
- 不要手动删除 `data.backup-*`，确认升级稳定后再清理。

---

## 5. 从 v2.0.5 升级到 v3.0.0 的特别说明

v3.0.0 的主要变化：

- 支持多工作表选择。
- 支持空单元格保留占位符。
- 执行前校验 Excel 列。
- 结果明细分页。
- 导出令牌带过期时间。
- 容器改为非 root 用户（UID `10001`）。
- 合并改用 `docxcompose`。

升级时注意：

1. 由于容器改为非 root，若 `./data` 原先由 root 创建，升级后可能报权限错误，执行：

```bash
sudo chown -R 10001:10001 ./data
```

2. 旧版全局规则会在启动时自动迁移为“历史规则导入”模板（如果存在旧规则）。
3. 旧版生成的下载链接在重启后全部失效，需要重新执行。
4. 建议在升级后用 [ACCEPTANCE_CHECKLIST.md](../ACCEPTANCE_CHECKLIST.md) 逐项点一遍。

---

## 6. 常见升级问题

| 现象 | 处理 |
|---|---|
| 拉取镜像失败 | 检查网络、镜像标签是否存在、是否需要登录 GHCR。 |
| 启动后 unhealthy | 检查 `./data` 权限与日志，见 [TROUBLESHOOTING.md](TROUBLESHOOTING.md)。 |
| 页面仍是旧版本 | 强制刷新浏览器；确认容器已用新镜像重建。 |
| 模板丢失 | 确认挂载路径正确，必要时从 `data.backup-*` 恢复。 |
| 下载旧链接失效 | 正常现象，重新执行生成新链接。 |
| 权限错误 | `sudo chown -R 10001:10001 ./data`。 |

---

## 7. 升级检查清单

- [ ] 已记录当前镜像标签。
- [ ] 已停止服务并备份 `./data`。
- [ ] 已确认备份目录或 tar 包存在且大小正常。
- [ ] 已修改为目标版本标签。
- [ ] 已执行 `docker compose pull` 和 `docker compose up -d`。
- [ ] `/health` 与 `/health/ready` 正常。
- [ ] `/config` 与页面底部版本一致。
- [ ] 旧规则模板仍可读取。
- [ ] 已用示例文件完成一次预览、执行、下载。
- [ ] 已检查日志没有明显错误。
- [ ] 确认稳定后再清理旧备份。
