# 快速开始

本文带你在几分钟内启动并完成第一次批量替换。当前版本为 **v3.0.0**，以根目录 [`VERSION`](VERSION) 为准。

> 返回 [README](README.md) · 部署细节见 [docs/DEPLOYMENT.md](docs/DEPLOYMENT.md) · 遇到问题见 [docs/TROUBLESHOOTING.md](docs/TROUBLESHOOTING.md)

---

## 1. 最快启动

### 方式一：源码构建

在项目根目录执行：

```bash
docker compose up -d --build
```

打开：

```text
http://localhost:12344
```

### 方式二：使用已发布镜像

```bash
docker compose -f docker-compose.example.yml pull
docker compose -f docker-compose.example.yml up -d
```

访问 `http://你的服务器IP:12344`。

### 确认启动成功

```bash
docker compose ps
curl -fsS http://localhost:12344/health
curl -fsS http://localhost:12344/health/ready
```

预期返回：

```json
{"status":"ok"}
{"status":"ready"}
```

群晖、反向代理与 HTTPS 请参考 [docs/DEPLOYMENT.md](docs/DEPLOYMENT.md)。

---

## 2. 第一次使用（5 步）

### 第 1 步：上传文件

- Word 模板：`.docx`
- Excel 数据：`.xlsx` 或 `.xls`
- 默认单文件上限 50MB（可通过 `MAX_UPLOAD_SIZE_MB` 调整）。

上传后会自动预览，不需要点击“预览”按钮。

> 页面上有“示例与快速填充”，可以直接下载两组示例文件先体验。

### 第 2 步：确认预览与工作表

- 左侧显示 Word 正文、表格、页眉页脚文字。
- 右侧显示 Excel 表头和前 25 行，以及总行数。
- 如果 Excel 有多个工作表，会出现“工作表”下拉框，请切换到你要用的那一张。

切换工作表后，预览和后续执行都会使用当前选择的工作表。

### 第 3 步：添加替换规则

一条规则由两部分组成：

- **模板关键字**：Word 里要替换的文字，例如 `【姓名】`。
- **Excel 列名**：从下拉框选择，例如 `姓名`。

点击“添加规则”。可以继续添加、编辑或删除规则。

注意：

- 列名必须与 Excel 表头完全一致。
- 如果规则引用了当前工作表不存在的列，执行按钮会变灰，并提示缺失的列名。
- 打开“显示高级项”后，可以把当前规则保存到模板库，或从模板库加载。

### 第 4 步：设置执行参数

- **起始行 / 结束行**：要处理的数据范围。页面上第一条数据对应起始行 1。
- **文件名来源列**：用哪一列生成文件名。必须选择存在的列。
- **序号来源**：
  - “不添加序号”
  - “使用替换序号”（自动 1、2、3……）
  - “使用 Excel 列”（用某一列的值作为序号）
- **序号格式**：`1`、`01`、`0001`、`1.`、`一`。
- **文件名前缀 / 后缀**：可选。
- **空值保留占位符**：勾选后，Excel 单元格为空时不会把模板占位符替换成空白。

### 第 5 步：执行并下载

点击“开始替换”。完成后会看到：

- 总数
- 成功
- 失败
- 替换次数
- 明细列表（每页 50 条，可上一页 / 下一页）

可以下载：

- **ZIP**：全部成功文件打包下载。
- **合并文档**：合并为一个 Word，图片、页眉、编号尽量保留。
- **单个文件**：在明细里点击“下载”。

也可以在明细里删除某条结果，统计数字会同步更新。

---

## 3. 结果有效期

- 结果保存在内存中，默认保留 **30 分钟**。
- 导出链接带签名和有效期。
- 过期或重启容器后，结果会消失，需要重新执行。
- 需要更长时间时，调整 `RUN_CACHE_TTL_SECONDS` 并重启。

> 重要结果请在过期前下载保存。

---

## 4. 常见问题

| 现象 | 处理 |
|---|---|
| 执行按钮不可点 | 确认已上传两个文件、至少一条规则、起止行有效、列名都存在。 |
| 提示“文件内容与扩展名不匹配” | 文件可能只是改了后缀，请用 Office / WPS 另存为正确格式。 |
| 提示 Excel 列不存在 | 检查表头文字、空格和当前工作表。 |
| 执行后 0 条 | 检查起止行、工作表和 Excel 是否有数据。 |
| 下载提示结果已过期 | 重新执行一次，或调大 `RUN_CACHE_TTL_SECONDS`。 |
| `./data` 权限错误 | 执行 `sudo chown -R 10001:10001 ./data`。 |
| 上传返回 413 | 调大 `MAX_UPLOAD_SIZE_MB`，并同步调大反向代理请求体上限。 |
| 容器 unhealthy | 查看 `docker compose logs`，确认数据库目录可写。 |

完整排查表见 [docs/TROUBLESHOOTING.md](docs/TROUBLESHOOTING.md)。

---

## 5. 升级与备份

升级前先备份 `./data`：

```bash
docker compose stop
cp -r ./data "./data.backup-$(date +%Y%m%d-%H%M%S)"
docker compose start
```

升级：

```bash
docker compose pull
docker compose up -d
```

完整步骤见 [docs/UPGRADE.md](docs/UPGRADE.md)。安全建议见 [SECURITY.md](SECURITY.md)。

---

## 6. 源码开发模式（可选）

### 后端

```bash
cd backend
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
uvicorn app.main:app --reload --port 8000
```

### 前端

```bash
cd frontend
cp .env.local.example .env.local
pnpm install --ignore-scripts
pnpm dev
```

开发时前端默认在 `http://localhost:3000`，后端在 `http://localhost:8000`。默认 CORS 已包含 `http://localhost:3000`。

更多约定见 [CONTRIBUTING.md](CONTRIBUTING.md)。

---

## 7. 下一步

- [部署指南](docs/DEPLOYMENT.md)
- [故障排查](docs/TROUBLESHOOTING.md)
- [架构说明](docs/ARCHITECTURE.md)
- [API 参考](docs/API.md)
- [升级与回滚](docs/UPGRADE.md)
- [验收清单](ACCEPTANCE_CHECKLIST.md)
- [更新日志](CHANGELOG.md)
