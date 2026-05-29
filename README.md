# 🚀 WordReplace 2.0

> Word + Excel 批量替换工具（单镜像版）

把 Word 模板（`.docx`）和 Excel 数据（`.xlsx/.xls`）批量合成目标文档，支持一键导出 ZIP 或合并文档。  
适合：通知书、合同、证书、函件、名册等标准化文档场景。

![WordReplace UI 预览](docs/images/ui-overview.png)

---

## 🧩 为什么用它

- 批量替换：模板关键字自动匹配 Excel 列
- 结果清晰：执行后直接看到总数/成功/失败/替换次数
- 导出灵活：支持 ZIP 批量下载 + 合并文档下载
- 部署简单：单镜像、单容器，一条命令即可启动

---

## 🐳 30 秒快速部署（推荐）

### 方式 1：本地源码一键启动

```bash
docker compose up -d --build
```

打开：`http://localhost:12344`

---

### 方式 2：直接拉取已发布镜像（最省事）

新建 `docker-compose.yml`，粘贴下面内容：

```yaml
version: "3.9"
services:
  wordreplace:
    image: ghcr.io/marod1m/wordreplace:latest
    container_name: wordreplace
    restart: unless-stopped
    ports:
      - "12344:8000" # 左侧端口可改
    # 可选：持久化数据库（推荐）
    volumes:
      - ./data:/app/data
    # 可选：安全参数（均提供默认值）
    environment:
      CORS_ALLOW_ORIGINS: "http://localhost:12344,http://localhost:3000"
      MAX_UPLOAD_SIZE_MB: "50"
      RUN_CACHE_TTL_SECONDS: "1800"
      RUN_CACHE_MAX_ENTRIES: "200"
      EXPORT_TOKEN_SECRET: "please-change-to-a-long-random-secret"
```

启动：

```bash
docker compose up -d
```

访问：`http://你的服务器IP:12344`

### 群晖快捷示例（可直接用）

如果你在群晖 Container Manager 部署，推荐直接用下面这份：

```yaml
version: "3.9"
services:
  wordreplace:
    image: ghcr.io/marod1m/wordreplace:latest
    container_name: wordreplace
    restart: unless-stopped
    ports:
      - "12344:8000" # 改成你想要的访问端口
    volumes:
      - /volume1/docker/wordreplace/data:/app/data
    environment:
      CORS_ALLOW_ORIGINS: "http://你的群晖IP:12344"
      MAX_UPLOAD_SIZE_MB: "50"
      RUN_CACHE_TTL_SECONDS: "1800"
      RUN_CACHE_MAX_ENTRIES: "200"
      EXPORT_TOKEN_SECRET: "please-change-to-a-long-random-secret"
```

说明：
- 你通常只需要改：端口、群晖 IP、数据目录路径、导出密钥。
- `/volume1/docker/wordreplace/data` 建议提前创建。

### 变量说明（environment）

| 变量名 | 默认值 | 是否可选 | 用途 |
|---|---|---|---|
| `CORS_ALLOW_ORIGINS` | `http://localhost:12344,http://localhost:3000` | 可选 | 配置允许跨域访问 API 的前端域名（逗号分隔） |
| `MAX_UPLOAD_SIZE_MB` | `50` | 可选 | 上传文件大小上限（MB），超出返回 413 |
| `RUN_CACHE_TTL_SECONDS` | `1800` | 可选 | 结果缓存有效期（秒） |
| `RUN_CACHE_MAX_ENTRIES` | `200` | 可选 | 结果缓存最大记录数，超出自动淘汰最旧记录 |
| `EXPORT_TOKEN_SECRET` | `change-this-secret`（代码默认） | 可选（生产强烈建议配置） | 导出签名密钥，防止未授权下载 |

> ℹ️ `CORS_ALLOW_ORIGINS` 需填写 **完整 Origin**（`协议://域名[:端口]`），多个值用英文逗号分隔；不要带路径（如 `/api`）。生产环境请改成你的真实域名。


### volumes 说明

| 挂载项 | 默认值 | 是否可选 | 用途 |
|---|---|---|---|
| `./data:/app/data` | 无（不挂载） | 可选（推荐） | 持久化 SQLite 数据库，容器重建后规则不丢失 |

> ⚠️ 生产环境建议：务必设置 `EXPORT_TOKEN_SECRET` 为高强度随机字符串，并将 `CORS_ALLOW_ORIGINS` 改为你的真实域名。

---

## 💾 可选：持久化数据库（推荐）

默认情况下，规则数据保存在容器内。若你会重建容器，建议映射数据库目录。

---

## ✅ 使用流程（5 步）

1. 上传 Word 模板和 Excel 数据
2. 添加替换规则（模板关键字 -> Excel 列名）
3. 设置起始行、结束行、文件名列
4. 点击开始替换
5. 下载 ZIP 或合并文档

---

## 🗂️ 项目结构（核心）

```text
backend/
  app/
    api/        # 路由
    services/   # 替换与导出逻辑
    models/     # SQLite 模型
    schemas/    # 请求/响应模型
frontend/
  src/app/      # 页面
  src/lib/api.ts
Dockerfile      # 单镜像构建（前后端合并）
```

---

## 🛠️ 开发模式（可选）

如果你要本地调试源码：

后端：
```bash
cd backend
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
uvicorn app.main:app --reload --port 8000
```

前端：
```bash
cd frontend
cp .env.local.example .env.local
pnpm install --ignore-scripts
pnpm dev
```

---

## ⚙️ GitHub 自动构建

仓库已配置自动构建并推送 GHCR：

- 工作流：`.github/workflows/docker-publish.yml`
- 触发：推送 `v*.*.*` 标签、手动触发（`workflow_dispatch`）
- 镜像：`ghcr.io/<你的仓库>`
- 标签：`latest` + 语义化版本

---

## ❓ 常见问题

- 执行按钮不可点：请确认已上传两个文件、至少一条规则、行号范围有效。
- 导出失败：请先执行替换并确认已生成结果。
- Excel 列名不生效：规则列名必须与 Excel 表头完全一致。

---

## 🔒 固定版本部署（推荐生产环境）

不建议生产环境长期使用 `latest`，建议固定版本标签：

```yaml
version: "3.9"
services:
  wordreplace:
    image: ghcr.io/marod1m/wordreplace:v2.0.1
    container_name: wordreplace
    restart: unless-stopped
    ports:
      - "12344:8000"
```

---

## 🔄 升级与回滚

### 升级到最新版本

```bash
docker compose pull
docker compose up -d
```

### 升级到指定版本

1. 修改 compose 中镜像标签（例如 `v2.0.1`）
2. 执行：

```bash
docker compose pull
docker compose up -d
```

### 回滚到旧版本

1. 将镜像标签改回历史版本（例如 `v2.0.0`）
2. 执行：

```bash
docker compose pull
docker compose up -d
```

---

## 🧹 停止与卸载

停止服务：

```bash
docker compose down
```

连同镜像一起清理（可选）：

```bash
docker compose down --rmi local
```
