# API 参考

本文列出 WordReplace 提供的 HTTP 接口、表单字段、返回结构与 `curl` 示例。

> 返回 [README](../README.md) · 架构见 [ARCHITECTURE.md](ARCHITECTURE.md) · 排障见 [TROUBLESHOOTING.md](TROUBLESHOOTING.md)

---

## 基本信息

- 默认端口：容器内 `8000`，Docker 映射到宿主机 `12344`。
- 开发模式后端：`http://localhost:8000`
- Docker 部署：`http://localhost:12344`
- 接口没有统一前缀，直接挂在根路径下。
- 请求与响应默认使用 JSON；上传接口使用 `multipart/form-data`。
- 错误响应格式：

```json
{"detail": "错误说明"}
```

常见状态码：

| 状态码 | 含义 |
|---|---|
| 400 | 参数错误、文件不合法、列不存在、规则为空 |
| 403 | 导出令牌无效 |
| 404 | 任务、文件或模板不存在 |
| 413 | 上传文件超过大小限制 |
| 503 | 数据库不可用（`/health/ready`） |

本文示例先设置基础地址：

```bash
BASE=http://localhost:12344
```

如果使用开发模式，改成：

```bash
BASE=http://localhost:8000
```

---

## 1. 健康检查

### GET /health

存活检查，不访问数据库。

```bash
curl -fsS "$BASE/health"
```

响应：

```json
{"status": "ok"}
```

### GET /health/ready

就绪检查，会执行一次 `SELECT 1`。

```bash
curl -fsS "$BASE/health/ready"
```

正常响应：

```json
{"status": "ready"}
```

数据库不可用时返回 `503`：

```json
{"detail": "数据库不可用: ..."}
```

---

## 2. 应用配置

### GET /config

返回当前版本与主要限制，前端用它显示上传大小、提示信息等。

```bash
curl -fsS "$BASE/config"
```

响应示例：

```json
{
  "app_version": "3.0.0",
  "max_upload_size_mb": 50,
  "max_excel_rows": 20000,
  "max_replace_rules": 500,
  "max_preview_rows": 25,
  "run_cache_ttl_seconds": 1800
}
```

---

## 3. 文件预览

### POST /preview

上传 Word 与 Excel，返回模板文本、Excel 表头、前 25 行和工作表信息。

表单字段：

| 字段 | 类型 | 必填 | 说明 |
|---|---|---|---|
| `word_file` | 文件 | 是 | Word 模板，`.docx`。 |
| `excel_file` | 文件 | 是 | Excel 数据，`.xlsx` 或 `.xls`。 |
| `sheet_name` | 字符串 | 否 | 要预览的工作表名称；为空时使用第一个工作表。 |

```bash
curl -fsS -X POST "$BASE/preview" \
  -F "word_file=@./template.docx" \
  -F "excel_file=@./data.xlsx" \
  -F "sheet_name=Sheet1"
```

响应示例：

```json
{
  "word_text": "尊敬的【姓名】：\n……",
  "excel_columns": ["姓名", "部门", "文件名"],
  "excel_rows": [["张三", "销售部", "张三.docx"]],
  "excel_total_rows": 15,
  "sheet_names": ["Sheet1", "Sheet2"],
  "selected_sheet": "Sheet1"
}
```

说明：

- `word_text` 包含正文、表格、页眉页脚中可提取的文字。
- `excel_rows` 最多返回 `MAX_PREVIEW_ROWS`（默认 25）行。
- 工作表名称不存在时返回 400。
- 文件扩展名、文件头或大小不合法时返回 400 / 413。

---

## 4. 执行替换

### POST /replace/execute

按 Excel 行批量替换 Word 模板，生成结果并返回下载令牌。

表单字段：

| 字段 | 类型 | 必填 | 默认值 | 说明 |
|---|---|---|---|---|
| `word_file` | 文件 | 是 | - | Word 模板，`.docx`。 |
| `excel_file` | 文件 | 是 | - | Excel 数据，`.xlsx` / `.xls`。 |
| `start_row` | 整数 | 是 | - | 起始数据行，最小 1（对应 Excel 第一条数据）。 |
| `end_row` | 整数 | 是 | - | 结束数据行，必须大于等于 `start_row`。 |
| `file_name_column` | 字符串 | 是 | - | 文件名来源列，必须存在于当前工作表。 |
| `export_mode` | 字符串 | 否 | `zip` | 只接受 `zip` 或 `merge`；实际导出由导出接口完成。 |
| `rules_json` | 字符串 | 是 | - | JSON 数组，元素为 `{"keyword": "...", "excel_column": "..."}`。 |
| `use_seq_prefix` | 布尔 | 否 | `false` | 是否在文件名前加序号或序号列的值。 |
| `seq_format` | 字符串 | 否 | `1` | `1`、`01`、`0001`、`1.`、`一`。 |
| `seq_column` | 字符串 | 否 | 空 | 从该 Excel 列取序号文本；为空时使用自动序号。 |
| `name_prefix` | 字符串 | 否 | 空 | 文件名前缀。 |
| `name_suffix` | 字符串 | 否 | 空 | 文件名后缀。 |
| `sheet_name` | 字符串 | 否 | 空 | 使用的工作表；为空时使用第一个。 |
| `keep_placeholder_on_empty` | 布尔 | 否 | `false` | 为 `true` 时，Excel 空值保留模板占位符。 |

```bash
RULES='[{"keyword":"【姓名】","excel_column":"姓名"},{"keyword":"【部门】","excel_column":"部门"}]'

curl -fsS -X POST "$BASE/replace/execute" \
  -F "word_file=@./template.docx" \
  -F "excel_file=@./data.xlsx" \
  -F "start_row=1" \
  -F "end_row=10" \
  -F "file_name_column=文件名" \
  -F "export_mode=zip" \
  -F "rules_json=$RULES" \
  -F "use_seq_prefix=true" \
  -F "seq_format=01" \
  -F "seq_column=" \
  -F "name_prefix=" \
  -F "name_suffix=" \
  -F "sheet_name=Sheet1" \
  -F "keep_placeholder_on_empty=false"
```

响应示例：

```json
{
  "run_id": "a1b2c3d4e5f6",
  "export_token": "1762500000.9f8e7d6c5b4a39281706f5e4d3c2b1a0...",
  "total": 2,
  "success": 2,
  "failed": 0,
  "replacements": 4,
  "details": [
    {
      "item_id": "1a2b3c4d",
      "seq": 1,
      "row_number": 2,
      "file_name": "01_张三.docx",
      "status": "成功",
      "replace_count": 2,
      "message": ""
    },
    {
      "item_id": "5e6f7a8b",
      "seq": 2,
      "row_number": 3,
      "file_name": "02_李四.docx",
      "status": "成功",
      "replace_count": 2,
      "message": ""
    }
  ]
}
```

字段说明：

- `total`：本次处理的数据行数。
- `success` / `failed`：成功与失败数量。
- `replacements`：所有成功文档的替换次数总和。
- `details[].row_number`：对应 Excel 中的行号（表头算第 1 行，第一条数据是第 2 行）。
- `details` 会返回完整明细；如需分页读取，使用 `/export/details/{run_id}`。

常见错误：

- 起始行大于结束行：400。
- 规则格式错误或为空：400。
- 规则数量超过限制：400。
- 列不存在：400，`detail` 会列出缺失列名。
- 单次结果超过缓存字节上限：400，提示减少行数。

### POST /replace/preview（已废弃）

旧接口，固定返回 410：

```json
{"detail": "请改用 /replace/execute"}
```

---

## 5. 导出下载

所有导出接口都需要 `token` 查询参数。`token` 来自 `/replace/execute` 的 `export_token`，默认有效期 30 分钟。

### GET /export/modes

返回支持的导出模式。

```bash
curl -fsS "$BASE/export/modes"
```

```json
{"modes": ["zip", "merge"]}
```

### GET /export/details/{run_id}

分页读取结果明细。

查询参数：

| 参数 | 必填 | 默认值 | 说明 |
|---|---|---|---|
| `token` | 是 | - | 导出令牌。 |
| `page` | 否 | 1 | 页码，最小 1。 |
| `page_size` | 否 | 50 | 每页条数，范围 1-200。 |

```bash
curl -fsS "$BASE/export/details/$RUN_ID?token=$TOKEN&page=1&page_size=50"
```

响应示例：

```json
{
  "total": 120,
  "page": 1,
  "page_size": 50,
  "details": [
    {
      "item_id": "1a2b3c4d",
      "seq": 1,
      "row_number": 2,
      "file_name": "张三.docx",
      "status": "成功",
      "replace_count": 2,
      "message": ""
    }
  ]
}
```

错误：

- token 无效：403。
- run 不存在或已过期：404。
- 分页参数越界：400。

### GET /export/zip/{run_id}

下载包含全部成功文件的 ZIP。

```bash
curl -fL "$BASE/export/zip/$RUN_ID?token=$TOKEN" -o replace_results.zip
```

响应：

- `Content-Type: application/zip`
- 文件名：`replace_results.zip`
- ZIP 内重名文件会自动加 `_1`、`_2`。

### GET /export/merge/{run_id}

下载合并后的单个 Word 文档。使用 `docxcompose` 合并成功文件。

```bash
curl -fL "$BASE/export/merge/$RUN_ID?token=$TOKEN" -o replace_merged.docx
```

响应：

- `Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document`
- 文件名：`replace_merged.docx`

### GET /export/file/{run_id}/{item_id}

下载单个结果文件，支持中文文件名。

```bash
curl -fL "$BASE/export/file/$RUN_ID/$ITEM_ID?token=$TOKEN" -o result.docx
```

响应使用 RFC 5987 的 `filename*` 传递 UTF-8 文件名，浏览器会显示原始中文名。

### DELETE /export/result/{run_id}/{item_id}

删除一条结果（成功或失败明细都可删除），并返回更新后的统计。

```bash
curl -fsS -X DELETE "$BASE/export/result/$RUN_ID/$ITEM_ID?token=$TOKEN"
```

响应示例：

```json
{
  "deleted": true,
  "total": 1,
  "success": 1,
  "failed": 0,
  "replacements": 2,
  "details": [
    {
      "item_id": "5e6f7a8b",
      "seq": 1,
      "row_number": 3,
      "file_name": "李四.docx",
      "status": "成功",
      "replace_count": 2,
      "message": ""
    }
  ]
}
```

删除后：

- `total`、`success`、`failed`、`replacements` 会重新计算。
- 剩余明细的 `seq` 会重新连续编号。
- 删除失败行时，`deleted` 可能为 `true`，即使它没有对应文件。

---

## 6. 规则模板

规则模板保存在 SQLite 中，可用于复用一组“关键字 → Excel 列”的映射。

### GET /rule-templates

列出所有启用中的模板。

```bash
curl -fsS "$BASE/rule-templates"
```

响应示例：

```json
[
  {
    "id": 1,
    "name": "入职通知书",
    "creator": "张三",
    "description": "常用字段",
    "is_active": true,
    "created_at_bj": "2026-11-10 10:00:00",
    "updated_at_bj": "2026-11-10 10:00:00",
    "item_count": 3
  }
]
```

### GET /rule-templates/{template_id}

获取模板详情，包含规则项。

```bash
curl -fsS "$BASE/rule-templates/1"
```

### POST /rule-templates

创建模板。

请求体：

```json
{
  "name": "入职通知书",
  "creator": "张三",
  "description": "常用字段",
  "items": [
    {"keyword": "【姓名】", "excel_column": "姓名"},
    {"keyword": "【部门】", "excel_column": "部门"}
  ]
}
```

```bash
curl -fsS -X POST "$BASE/rule-templates" \
  -H "Content-Type: application/json" \
  -d '{
    "name": "入职通知书",
    "creator": "张三",
    "description": "常用字段",
    "items": [
      {"keyword": "【姓名】", "excel_column": "姓名"},
      {"keyword": "【部门】", "excel_column": "部门"}
    ]
  }'
```

### PUT /rule-templates/{template_id}

更新模板，字段均为可选。传入 `items` 时会整体替换原有规则项。

```bash
curl -fsS -X PUT "$BASE/rule-templates/1" \
  -H "Content-Type: application/json" \
  -d '{"description":"新的说明"}'
```

### DELETE /rule-templates/{template_id}

软删除模板（`is_active` 置为 false），返回 204。

```bash
curl -i -X DELETE "$BASE/rule-templates/1"
```

### POST /rule-templates/{template_id}/apply

按当前 Excel 列校验模板，返回可用与不可用的规则项。

请求体：

```json
{
  "excel_columns": ["姓名", "部门", "文件名"],
  "mode": "replace"
}
```

- `mode` 只接受 `replace` 或 `append`，用于前端决定覆盖还是追加；服务端返回校验结果。

```bash
curl -fsS -X POST "$BASE/rule-templates/1/apply" \
  -H "Content-Type: application/json" \
  -d '{"excel_columns":["姓名","部门","文件名"],"mode":"replace"}'
```

响应示例：

```json
{
  "valid_items": [
    {"keyword": "【姓名】", "excel_column": "姓名"},
    {"keyword": "【部门】", "excel_column": "部门"}
  ],
  "invalid_items": []
}
```

---

## 7. 旧版规则接口

这些接口仍然存在，但页面的执行流程使用的是“当前会话规则”（通过 `rules_json` 提交），不依赖这些全局规则。保留它们主要用于兼容和模板迁移。

| 方法 | 路径 | 说明 |
|---|---|---|
| GET | `/rules` | 列出规则 |
| POST | `/rules` | 创建规则，请求体 `{"keyword":"...","excel_column":"..."}` |
| PUT | `/rules/{id}` | 更新规则 |
| DELETE | `/rules/{id}` | 删除规则，返回 204 |

示例：

```bash
curl -fsS "$BASE/rules"

curl -fsS -X POST "$BASE/rules" \
  -H "Content-Type: application/json" \
  -d '{"keyword":"【姓名】","excel_column":"姓名"}'
```

---

## 8. 完整调用流程示例

以下示例假设已安装 `jq`，用于从响应中提取字段：

```bash
BASE=http://localhost:12344
RULES='[{"keyword":"【姓名】","excel_column":"姓名"}]'

RESP=$(curl -fsS -X POST "$BASE/replace/execute" \
  -F "word_file=@./template.docx" \
  -F "excel_file=@./data.xlsx" \
  -F "start_row=1" \
  -F "end_row=10" \
  -F "file_name_column=文件名" \
  -F "rules_json=$RULES")

RUN_ID=$(echo "$RESP" | jq -r '.run_id')
TOKEN=$(echo "$RESP" | jq -r '.export_token')

curl -fL "$BASE/export/zip/$RUN_ID?token=$TOKEN" -o replace_results.zip
curl -fL "$BASE/export/merge/$RUN_ID?token=$TOKEN" -o replace_merged.docx
curl -fsS "$BASE/export/details/$RUN_ID?token=$TOKEN&page=1&page_size=50" | jq
```

如果没有 `jq`，也可以手动从响应中复制 `run_id` 和 `export_token`。

---

## 9. 交互式文档

FastAPI 默认提供 Swagger UI 与 OpenAPI 文档：

- `http://localhost:12344/docs`
- `http://localhost:12344/openapi.json`

如果反向代理或自定义前端路由覆盖了这些路径，请直接使用本文的接口说明。

---

## 10. 注意事项

- 结果与令牌默认 **30 分钟** 后失效，过期后重新执行。
- 不要公开分享带 `token` 的下载链接。
- 公网部署前请修改 `EXPORT_TOKEN_SECRET`，并在反向代理层增加认证。
- 所有接口都没有内置登录校验，请勿在无保护的情况下暴露到公网。

更多安全建议见 [SECURITY.md](../SECURITY.md)。
