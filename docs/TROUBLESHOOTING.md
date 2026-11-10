# 故障排查

本文按“现象 → 原因 → 处理办法”整理常见问题。先看一遍下面的速查表，再按编号查看详细说明。

> 返回 [README](../README.md) · 部署问题看 [DEPLOYMENT.md](DEPLOYMENT.md) · 接口问题看 [API.md](API.md)

---

## 速查表

| 编号 | 现象 | 最常见原因 | 快速处理 |
|---|---|---|---|
| 1 | 上传返回 413 / 提示超过大小限制 | 文件超过 `MAX_UPLOAD_SIZE_MB`，或反向代理限制更小 | 调大应用限制和代理 `client_max_body_size` |
| 2 | 提示“文件内容与扩展名不匹配” | 文件只是改了后缀，真实内容不是 docx/xlsx | 用 Office / WPS 重新另存为正确格式 |
| 3 | 提示“以下 Excel 列不存在” | 列名不一致、选错工作表、表头有空格 | 使用列选择器，确认工作表与表头 |
| 4 | 执行后结果 0 条 | 行范围超出数据、空表、选错工作表 | 检查起止行和工作表 |
| 5 | 下载提示“导出令牌无效 / 结果已过期” | 结果超过 TTL 被清理，或 token 被改动 | 重新执行；必要时调大 TTL |
| 6 | 合并文档失败或格式异常 | 文档过大、结构特殊、依赖缺失 | 先用 ZIP；确认 `docxcompose` 已安装 |
| 7 | `./data` 权限错误 Permission denied | 宿主机目录不属于 UID 10001 | `sudo chown -R 10001:10001 ./data` |
| 8 | 启动失败，端口被占用 | `12344` 已被其他程序使用 | 改端口或停止占用程序 |
| 9 | 容器状态 unhealthy | 数据库不可写、启动未完成、健康检查失败 | 看日志与 `/health/ready` |
| 10 | 数据库 locked | 多进程同时写、备份时仍在写入 | 保持单 worker，备份前先停止容器 |
| 11 | Docker 构建失败 | 网络、依赖、磁盘或缓存问题 | 看构建日志，清理缓存后重试 |
| 12 | 版本号不一致 | `VERSION`、镜像标签、页面版本不同步 | 运行版本同步脚本并拉取正确镜像 |

---

## 1. 上传返回 413 或“超过大小限制”

### 现象

- 浏览器提示上传失败。
- 接口返回 `413`。
- 中文提示类似“Word 超过 50MB 限制”。

### 原因

应用有两层限制：

1. 应用自身的 `MAX_UPLOAD_SIZE_MB`（默认 50）。
2. 反向代理的请求体上限（Nginx 的 `client_max_body_size`、群晖反向代理、Cloudflare 等）。

任意一层更小，都会导致 413。

### 处理

1. 调大应用限制，例如改成 100MB：

```yaml
    environment:
      MAX_UPLOAD_SIZE_MB: "100"
```

2. 重启：

```bash
docker compose up -d
```

3. Nginx 同步调大：

```nginx
client_max_body_size 100m;
```

4. 群晖反向代理、Cloudflare、Traefik 等也要检查对应限制。

5. 确认浏览器到容器的链路都放行后，再试一次。

> 不要只改代理不改应用，也不要只改应用不改代理。

---

## 2. 提示“文件内容与扩展名不匹配”

### 现象

上传时提示：

```text
Word 文件内容与扩展名不匹配
Excel 文件内容与扩展名不匹配
```

### 原因

系统会检查文件头（magic number），而不只看后缀：

- `.docx`、`.xlsx` 应该是 ZIP 格式（文件头 `PK`）。
- `.xls` 应该是 OLE 格式。
- 把 `.txt`、`.doc`、`.csv` 直接改名为 `.docx` / `.xlsx` 会触发此提示。

### 处理

- 用 Microsoft Office 或 WPS 打开原文件，选择“另存为”，保存为真正的 `.docx` / `.xlsx`。
- Excel 不要直接改后缀，请另存为“Excel 工作簿（.xlsx）”。
- 旧版 `.doc` 需要先另存为 `.docx`。
- 加密文档、损坏文档也可能解析失败，请先解除密码或修复文件。

---

## 3. 提示“以下 Excel 列不存在”

### 现象

- 执行按钮变灰。
- 提示类似：`以下 Excel 列不存在：姓名、文件`。
- 后端返回 400。

### 原因

规则里的“Excel 列名”必须与当前工作表表头**完全一致**。常见问题：

- 表头前后有空格。
- 中英文括号、全角半角不一致。
- 选错了工作表。
- Excel 第一行不是表头。
- 列名被合并单元格破坏。

### 处理

1. 在 Excel 预览里确认表头文字。
2. 规则中使用下拉选择器选择列，不要手动输入近似名称。
3. 多工作表时，先在预览区切换到正确的工作表。
4. 检查“文件名来源列”是否也存在。
5. 重新上传修改后的 Excel。

> 系统会自动去除表头首尾空格，但不会自动统一全角/半角或近义词。

---

## 4. 执行后结果 0 条

### 现象

- 执行成功，但“总数”“成功”“失败”都是 0。
- 下载按钮不可用。

### 原因

- 起始行大于结束行。
- 起始行超出数据行数。
- Excel 只有表头，没有数据。
- 选错工作表（例如数据在 `Sheet2`，当前选的是 `Sheet1`）。

### 处理

1. 确认起始行和结束行。页面行号从 Excel 数据第一行开始计数，包含表头偏移的说明见明细里的“Excel 行号”。
2. 切换到有数据的工作表。
3. 确认 Excel 至少有一行数据。
4. 如果确实没有数据，0 条是正常结果，不是报错。

---

## 5. 下载提示“导出令牌无效”或“结果已过期”

### 现象

- 下载 ZIP / 合并文档 / 单文件时失败。
- 提示“结果已过期”“任务不存在或已过期”“导出令牌无效”。

### 原因

- 结果保存在**内存**中，默认 30 分钟后清理（`RUN_CACHE_TTL_SECONDS`）。
- 导出令牌带有效期，过期后无法下载。
- 容器重启后内存结果全部丢失。
- 手动修改了下载链接里的 `token`。

### 处理

1. 重新执行一次替换，使用新生成的下载链接。
2. 需要更长时间保留时，调大 TTL：

```yaml
    environment:
      RUN_CACHE_TTL_SECONDS: "3600"
```

3. 修改后重启容器。
4. 重要结果请在过期前下载保存。

> 不要通过修改 token 或复制旧链接来恢复，重新执行最可靠。

---

## 6. 合并文档失败或格式异常

### 现象

- 下载合并文档失败。
- 合并后图片丢失、页眉异常、编号错乱。

### 原因

- 合并依赖 `docxcompose`。依赖缺失时会降级为基础合并，效果较差。
- 个别复杂模板（大量文本框、嵌入对象、特殊域）合并兼容性有限。
- 文档本身损坏或包含不受支持的 OLE 对象。

### 处理

1. 先使用“下载 ZIP”，逐个检查生成文件。ZIP 不涉及文档合并，最稳定。
2. 确认镜像内已安装 `docxcompose`（官方镜像已包含）。
3. 减少单次合并数量，分批导出。
4. 简化模板中的复杂对象后再试。
5. 如果只是需要最终打印，优先使用 ZIP + 打印合并，而不是依赖 Word 合并。

---

## 7. `./data` 权限错误（Permission denied）

### 现象

- 日志出现 `Permission denied`、`unable to open database file`。
- `/health/ready` 返回 503。
- 容器反复重启。

### 原因

容器以 UID `10001` 运行，而宿主机上的 `./data` 可能属于 root 或其他用户。

### 处理

Linux：

```bash
sudo chown -R 10001:10001 ./data
docker compose restart
```

群晖：

```bash
sudo chown -R 10001:10001 /volume1/docker/wordreplace/data
```

SELinux 系统：

```yaml
    volumes:
      - ./data:/app/data:Z
```

或：

```bash
sudo chcon -Rt container_file_t ./data
```

检查容器内用户：

```bash
docker exec wordreplace id
```

预期包含 `uid=10001(appuser)`。

---

## 8. 端口冲突

### 现象

- `docker compose up -d` 失败。
- 提示 `Bind for 0.0.0.0:12344 failed: port is already allocated`。

### 处理

查看占用：

```bash
sudo ss -lptn 'sport = :12344'
```

改端口，例如改成 `12345`：

```yaml
    ports:
      - "12345:8000"
```

然后：

```bash
docker compose up -d
```

访问 `http://localhost:12345`。

> 只改左边宿主机端口，右边容器端口 `8000` 保持不变。

---

## 9. 容器状态 unhealthy

### 现象

`docker compose ps` 显示 `unhealthy`，或 `docker inspect` 的健康状态不是 healthy。

### 排查

1. 看健康检查结果：

```bash
docker inspect --format '{{json .State.Health}}' wordreplace
```

2. 看日志：

```bash
docker compose logs --tail=200 wordreplace
```

3. 手动检查接口：

```bash
docker exec wordreplace python -c "import urllib.request; print(urllib.request.urlopen('http://127.0.0.1:8000/health', timeout=3).read())"
```

4. 检查就绪接口：

```bash
curl -i http://localhost:12344/health/ready
```

### 常见原因

- 数据目录不可写（见第 7 条）。
- 启动时间超过 `start-period`，但应用仍在初始化。
- 容器内 Python 进程异常退出。
- 端口配置错误导致请求打不到应用。

如果 `/health` 正常、`/health/ready` 返回 503，优先检查数据库路径与权限。

---

## 10. 数据库 locked（SQLite 被锁）

### 现象

- 日志出现 `database is locked`。
- 保存规则模板失败。
- 请求偶发 500。

### 原因

SQLite 同时只允许一个写入者。常见触发：

- 运行了多个 worker 或多个容器副本。
- 备份时直接复制正在写入的数据库。
- 有外部工具同时打开数据库。

### 处理

1. 保持单 worker、单副本运行（本项目的缓存和 SQLite 都按单实例设计）。
2. 不要用 `--workers 2` 启动 Uvicorn。
3. 备份前先停止容器：

```bash
docker compose stop
cp -r ./data ./data.backup-$(date +%Y%m%d-%H%M%S)
docker compose start
```

4. 确认没有其他程序打开 `wordreplace.db`。

---

## 11. Docker 构建失败

### 现象

`docker compose up -d --build` 在拉取依赖或构建前端时失败。

### 常见原因与处理

1. **网络问题**：拉取 npm / pip 依赖失败。

```bash
docker compose build --no-cache
```

2. **磁盘空间不足**：

```bash
df -h
docker system df
```

清理无用镜像和缓存：

```bash
docker system prune
```

3. **前端构建失败**：先本地验证。

```bash
cd frontend
pnpm install --frozen-lockfile --ignore-scripts
pnpm build
```

4. **后端依赖失败**：检查 Python 版本与网络。

```bash
cd backend
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

5. **架构不匹配**：确认镜像支持你的 CPU 架构（官方构建支持 `linux/amd64` 与 `linux/arm64`）。

查看完整日志：

```bash
docker compose build --progress=plain
```

---

## 12. 版本号不一致

### 现象

- 页面底部版本、镜像标签、`VERSION` 文件三者不一致。
- 升级后仍显示旧版本。
- CI 的版本校验失败。

### 处理

1. 查看根目录版本：

```bash
cat VERSION
```

2. 运行同步脚本：

```bash
python3 scripts/sync_version.py
python3 scripts/sync_version.py --check
```

3. 确认 compose 中的镜像标签与版本一致，例如 `3.0.0`（GitHub 发布标签为 `v3.0.0`）。
4. 重新拉取并重建：

```bash
docker compose pull
docker compose up -d
```

5. 强制重建本地镜像（源码部署时）：

```bash
docker compose up -d --build --force-recreate
```

6. 如果浏览器仍显示旧页面，强制刷新（`Ctrl + F5`）或清理浏览器缓存。

---

## 仍然无法解决？

请记录以下信息后再反馈：

- 操作步骤（越具体越好）。
- 页面提示文字或截图。
- 相关日志：

```bash
docker compose logs --tail=300 wordreplace
```

- 当前版本：

```bash
curl -fsS http://localhost:12344/config
```

- 容器状态：

```bash
docker compose ps
docker inspect --format '{{.State.Health.Status}}' wordreplace
```

请勿在公开渠道粘贴真实姓名、身份证号、合同内容等敏感数据。
