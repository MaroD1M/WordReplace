# 架构说明

本文说明 WordReplace 的组成、一次完整请求的流转、替换引擎、结果缓存、限制与安全模型，方便维护和二次开发。

> 返回 [README](../README.md) · 接口细节见 [API.md](API.md) · 部署见 [DEPLOYMENT.md](DEPLOYMENT.md)

---

## 1. 总体结构

WordReplace 采用“单镜像、单进程”设计：

```text
┌──────────────────────────────────────────────────────────┐
│ 浏览器                                                     │
│  Next.js 静态页面（由 FastAPI 提供）                        │
└──────────────────────────┬───────────────────────────────┘
                           │ HTTP / HTTPS
┌──────────────────────────▼───────────────────────────────┐
│ FastAPI 应用（uvicorn，默认 1 worker）                     │
│                                                          │
│  /config  /preview  /replace  /export                     │
│  /rules   /rule-templates  /health                        │
│                                                          │
│  ┌─────────────────────┐   ┌──────────────────────────┐  │
│  │ SQLite              │   │ 进程内存 RUN_CACHE        │  │
│  │ 规则 / 规则模板      │   │ 生成结果 + 导出令牌        │  │
│  └─────────────────────┘   └──────────────────────────┘  │
└──────────────────────────────────────────────────────────┘
```

- 前端使用 Next.js 静态导出（`output: "export"`），构建产物放在镜像的 `/app/frontend_dist`，由 FastAPI 直接托管。
- 后端负责上传校验、Excel 读取、Word 替换、结果缓存和下载。
- SQLite 只保存规则与规则模板；替换结果不落盘。

---

## 2. 组件划分

| 组件 | 位置 | 职责 |
|---|---|---|
| 应用入口 | `backend/app/main.py` | 创建 FastAPI、注册中间件与路由、挂载前端静态文件、启动时建表。 |
| 配置 | `backend/app/core/config.py` | 从环境变量读取配置，提供默认值与限制。 |
| 上传校验 | `backend/app/core/uploads.py` | 扩展名、文件头、大小校验，限制读取过程。 |
| Excel 读取 | `backend/app/core/excel.py` | 列出工作表、按字符串读取、检查行数与单元格上限。 |
| 预览接口 | `backend/app/api/preview.py` | 解析 Word 文本与 Excel 前 N 行，返回工作表信息。 |
| 执行接口 | `backend/app/api/replace.py` | 校验参数与规则，串行控制，调用替换引擎，写入缓存。 |
| 导出接口 | `backend/app/api/export.py` | 校验导出令牌，提供 ZIP、合并文档、单文件、分页明细、删除结果。 |
| 规则接口 | `backend/app/api/rules.py` | 旧版全局规则 CRUD。 |
| 模板库 | `backend/app/api/templates.py` | 规则模板 CRUD 与按 Excel 列应用。 |
| 替换引擎 | `backend/app/services/replace_service.py` | 文档替换、文件名处理、缓存、令牌、ZIP、合并。 |
| 模板服务 | `backend/app/services/template_service.py` | 模板持久化、旧规则迁移、列有效性校验。 |
| 数据库 | `backend/app/db/`、`backend/app/models/` | SQLAlchemy 引擎、会话、表模型。 |
| 前端 | `frontend/src/` | 上传、预览、规则、执行、结果、设置等交互。 |

---

## 3. 请求流转

### 3.1 上传与预览

```text
选择文件
  → 前端 250ms 防抖
  → POST /preview（word_file、excel_file、sheet_name）
  → 校验扩展名
  → 分块读取并限制大小
  → 校验文件头（docx/xlsx = ZIP，xls = OLE）
  → python-docx 解析正文、表格、页眉页脚文字
  → pandas 列出工作表，读取目标工作表
  → 返回 word_text、excel_columns、前 25 行、总行数、工作表列表
```

前端在文件变化、切换工作表时重新发起预览，并用 `AbortController` 取消旧请求。

### 3.2 执行替换

```text
POST /replace/execute
  → 校验起止行、导出模式、序号格式
  → 校验规则 JSON（数量、长度、非空）
  → 再次读取并校验两个上传文件
  → 获取 MAX_CONCURRENT_RUNS 信号量
  → 在线程池中执行 run_replace（避免阻塞事件循环）
  → 读取 Excel（dtype=str，空单元格为 ""）
  → 校验所有规则列与文件名列都存在
  → 逐行：深拷贝模板 → 替换 → 生成文件名 → 保存到内存
  → 汇总 details（成功 / 失败、替换次数、Excel 行号）
  → cache_run 写入内存缓存
  → sign_export_token 生成带有效期的下载令牌
  → 返回 run_id、export_token 和统计
```

`export_mode` 表单字段会校验为 `zip` 或 `merge`，但真正的导出由 `/export/zip` 与 `/export/merge` 分别触发。

### 3.3 结果缓存

```text
run_replace 结果
  → 计算占用字节（文件内容 + 明细条数 × 256）
  → 若单次结果超过 RUN_CACHE_MAX_BYTES：直接报错，不缓存
  → 清理过期记录
  → 写入 RUN_CACHE[run_id]
  → 再按“最大记录数”和“最大字节数”淘汰最旧记录
```

### 3.4 导出

```text
GET /export/zip|merge|file|details  或  DELETE /export/result
  → 校验 token 签名与过期时间
  → 从 RUN_CACHE 取 run（过期会被清理）
  → ZIP：逐个写入内存 ZIP，重名自动加 _1、_2
  → 合并：docxcompose.Composer 依次追加
  → 单文件：以 RFC 5987 文件名返回
  → 删除：移除文件和明细，重新计算统计与序号
```

---

## 4. 替换引擎

### 4.1 扫描范围

`iter_document_paragraphs` 会遍历：

- 文档正文中的全部段落，包括**嵌套表格**里的段落（通过遍历 `w:p` 元素实现）。
- 每一节的页眉、页脚、首页页眉页脚、奇偶页页眉页脚。
- 跳过“链接到前一节”的页眉页脚，避免重复处理。

### 4.2 段落文本的组成

`paragraph_text` 按文档顺序拼接以下位置的 run：

- 普通 run。
- 超链接内的 run（`w:hyperlink/w:r`）。
- 内容控件内的 run（`w:sdt/w:sdtContent/...`）。
- 修订插入内容（`w:ins`）。
- 智能标记（`w:smartTag`）。
- 简单域（`w:fldSimple`）。

因此，占位符写在内容控件、超链接、页眉页脚或嵌套表格中也能被替换。

### 4.3 保留格式的替换

占位符可能被 Word 拆分到多个 run 中（例如“【姓”和“名】”）。引擎的做法：

1. 把每个规则的占位符按长度从长到短排序，避免短规则先破坏长规则。
2. 在段落文本中查找所有出现位置。
3. 从后往前替换，避免位置偏移。
4. 跨 run 替换时：
   - 替换内容写入起始 run，保留起始 run 的字体格式。
   - 清空中间 run 的文字。
   - 保留中间 run 中的图片、绘图对象等嵌入内容。
   - 保留结束 run 中占位符之后的文字。

替换次数按“每个出现”计数，因此同一段落中出现两次会记为 2。

### 4.4 空值与占位符保留

Excel 以 `dtype=str`、`keep_default_na=False` 读取：

- 数字不会变成 `123.0`。
- 日期不会变成时间戳。
- 空单元格是空字符串 `""`，不是 `NaN`。

默认情况下，空单元格会把占位符替换为空。开启 `keep_placeholder_on_empty` 后，空值对应的规则会被跳过，模板中的原占位符保留。

### 4.5 文件名生成

1. 优先使用“文件名来源列”的值；为空时使用 `文件_行号.docx`。
2. 去掉重复的 `.docx` 后缀。
3. 替换 Windows 非法字符，去掉首尾空格和点。
4. 处理 Windows 保留名（`CON`、`PRN`、`AUX`、`NUL`、`COM1-9`、`LPT1-9`），必要时加 `_` 前缀。
5. 文件名主体最长 120 个字符。
6. 根据需要添加序号（`1`、`01`、`0001`、`1.`、`一`）、前缀、后缀；序号也可来自 Excel 列。
7. 同一批结果重名时自动加 `_1`、`_2`。
8. ZIP 内再次做一次重名保护。

### 4.6 合并

- 优先使用 `docxcompose` 的 `Composer`，依次追加文档，能较好保留图片、页眉、编号等。
- 如果运行环境缺少 `docxcompose`，降级为基础合并：复制正文元素，并在文档之间插入分页符。
- 合并只使用**成功生成**的文件；失败行不会进入合并。

---

## 5. 缓存模型

| 项目 | 说明 |
|---|---|
| 存储位置 | 进程内全局字典 `RUN_CACHE`，不落盘。 |
| 并发保护 | `threading.RLock`。 |
| 过期时间 | `RUN_CACHE_TTL_SECONDS`，默认 1800 秒。 |
| 最大记录数 | `RUN_CACHE_MAX_ENTRIES`，默认 200，超出淘汰最旧。 |
| 最大字节数 | `RUN_CACHE_MAX_BYTES`，默认 512MB，超出淘汰最旧。 |
| 单次结果上限 | 单次结果本身超过 `RUN_CACHE_MAX_BYTES` 时直接报错，不写入。 |
| 缓存键 | `run_id`（12 位十六进制）。 |
| 失效时机 | TTL 到期、容量淘汰、容器重启、手动删除结果。 |

导出令牌：

```text
token = expires_at + "." + HMAC-SHA256(EXPORT_TOKEN_SECRET, run_id + ":" + expires_at)
```

校验时同时检查：

- token 格式。
- 是否已过期。
- 签名是否匹配（`hmac.compare_digest`，防时序攻击）。
- 对应的 run 是否仍在缓存中。

令牌有效期默认等于 `RUN_CACHE_TTL_SECONDS`。修改 `EXPORT_TOKEN_SECRET` 后，旧令牌会立即失效；重启容器也会丢失所有结果。

### 为什么必须单 worker

缓存和令牌都依赖单个进程的内存。如果启动多个 worker 或多个副本：

- 执行请求落在 A 进程，下载请求可能落到 B 进程，导致找不到结果。
- SQLite 也会面临多写入者竞争。

因此默认 `MAX_CONCURRENT_RUNS=1`，且部署上保持单副本、单 worker。需要横向扩展时，应把结果缓存替换为 Redis 等外置存储，并评估 SQLite 的替换方案。

---

## 6. 限制一览

| 限制 | 环境变量 | 默认值 |
|---|---|---|
| 单文件上传大小 | `MAX_UPLOAD_SIZE_MB` | 50MB |
| Excel 数据行数 | `MAX_EXCEL_ROWS` | 20000 |
| Excel 单元格总数 | `MAX_EXCEL_CELLS` | 200000 |
| 单次规则数量 | `MAX_REPLACE_RULES` | 500 |
| 预览行数 | `MAX_PREVIEW_ROWS` | 25 |
| Word 预览字符数 | `MAX_PREVIEW_CHARS` | 20000 |
| 结果缓存有效期 | `RUN_CACHE_TTL_SECONDS` | 1800 秒 |
| 缓存记录数 | `RUN_CACHE_MAX_ENTRIES` | 200 |
| 缓存总字节 | `RUN_CACHE_MAX_BYTES` | 512MB |
| 同时执行任务数 | `MAX_CONCURRENT_RUNS` | 1 |

分页接口 `/export/details/{run_id}` 的 `page_size` 最大为 200。

---

## 7. 并发模型

- **上传与预览**：异步读取，按块限制大小；前端防抖并取消旧预览请求。
- **执行**：用 `asyncio.Semaphore(MAX_CONCURRENT_RUNS)` 限制同时执行数量；实际替换在线程池中运行，避免阻塞事件循环。
- **CPU 限制**：替换是 CPU 密集型工作，Python GIL 下单进程多线程不会线性提速；调大并发主要改善等待体验，不会显著提高总吞吐。
- **缓存与删除**：通过 `RLock` 保证字典和统计更新的一致性。
- **数据库**：SQLite 单写入者；模板保存等写操作应避免与外部工具同时进行。

个人使用建议保持：

```yaml
      MAX_CONCURRENT_RUNS: "1"
```

---

## 8. 安全模型

### 8.1 上传安全

- 只允许 `.docx`、`.xlsx`、`.xls` 扩展名。
- 检查文件头，拒绝改后缀的假文件。
- 读取过程中实时累计大小，超过上限立即返回 413。
- 空文件直接拒绝。
- Excel 读取后检查行数与单元格总数，避免超大表耗尽内存。

### 8.2 文件与内容

- 生成文件名会清理路径分隔符、Windows 非法字符和保留名，降低路径穿越与写入异常风险。
- 结果保存在内存缓存中；ZIP 打包使用临时文件（大包可能短暂落到系统临时目录），不会写入数据目录。
- 前端使用 React 渲染文本，默认转义；应用设置 CSP，限制脚本与连接来源。

### 8.3 下载与令牌

- 所有导出接口都要求 `token`。
- 令牌使用 HMAC-SHA256 签名，带过期时间，无法伪造或延长。
- 令牌与 `run_id` 绑定，换一个 run 无法复用。

### 8.4 网络与进程

- `ALLOWED_HOSTS` 不是 `*` 时，启用 `TrustedHostMiddleware`，拒绝不匹配 Host 的请求。
- `CORS_ALLOW_ORIGINS` 只允许配置的前端来源跨域调用 API。
- 默认返回安全响应头：
  - `X-Content-Type-Options: nosniff`
  - `X-Frame-Options: DENY`
  - `Referrer-Policy: no-referrer`
  - `Permissions-Policy`
  - 有前端资源时附加 CSP。
- 容器以 UID `10001` 的非 root 用户运行，启动时修正数据目录权限。

### 8.5 已知边界

- **没有内置账号体系**。公网部署必须在反向代理层增加认证，或限制为内网访问。
- 默认密钥 `change-this-secret` 仅供本地测试；公开前必须替换。
- 结果缓存不加密；同进程内可被有代码执行权限的人读取。请控制宿主机访问权限。
- 上传文件会在内存中解压解析，超大或恶意文档仍可能消耗较多内存；请结合资源限制使用。

---

## 9. 相关文档

- [API 参考](API.md)
- [部署指南](DEPLOYMENT.md)
- [故障排查](TROUBLESHOOTING.md)
- [升级与回滚](UPGRADE.md)
- [安全说明](../SECURITY.md)
