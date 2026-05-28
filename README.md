# 🚀 WordReplace · Word + Excel 批量替换工具

> 🔖 当前版本：`v1.6.4`  
> 🧩 形态：Streamlit Web 应用（前后端一体）  
> 👥 开发者：MaroD1M · Codex（AI 协作开发）

---

## ✨ 项目简介

`WordReplace` 用于将 **Word 模板（.docx）** 与 **Excel 数据（.xlsx）** 自动批量合成目标文档，适合通知书、合同、证书、名册等场景。

### ✅ 你能得到什么

- 🔁 批量替换关键字（段落 + 表格）
- 🧠 两种替换模式（完整替换 / 仅括号内容）
- 📦 一键导出 ZIP、合并文档、统计 CSV、日志 TXT
- 🛡️ 文件名安全与缓存键清洗
- 🧪 CI 自动测试、语法检查、镜像构建
- 🐳 Docker 开箱即用部署

---

## 🧱 目录结构

```text
app/
  main.py         # Streamlit 页面与交互逻辑
  services.py     # 替换/合并/统计业务
  core_utils.py   # 文本/文件名/校验工具
tests/
  test_services.py
  test_core_utils.py
Dockerfile
requirements.txt
docker-compose.yml
docker-compose.example.yml
.github/workflows/docker-publish.yml
```

---

## ⚡ 快速开始（推荐 Docker）

### 一条命令启动

```bash
docker compose up -d
```

访问：`http://localhost:12344`

### 停止服务

```bash
docker compose down
```

---

## 🐳 群晖用户（最简 Compose 示例）

在群晖 Docker / Container Manager 中可直接使用以下配置。  
你通常只需要改两处：
- 端口：`12344:8501`（把 `12344` 改成你想暴露的端口）
- 镜像版本：`latest`（或改为固定版本，如 `v1.6.4`）

```yaml
version: "3.9"
services:
  wordreplace:
    image: ghcr.io/marod1m/wordreplace:latest
    container_name: wordreplace
    restart: unless-stopped
    ports:
      - "12344:8501"
    environment:
      STREAMLIT_SERVER_HEADLESS: "true"
      STREAMLIT_BROWSER_GATHER_USAGE_STATS: "false"
      STREAMLIT_SERVER_MAX_UPLOAD_SIZE: "50"
```

启动后访问：`http://群晖IP:12344`

### 群晖 3 步部署（新手版）

1. 打开 **Container Manager** → **项目** → **新增项目**。  
2. 选择“通过 compose 文件创建”，粘贴上面的 YAML（按需改端口和版本）。  
3. 点击部署，等待容器启动后访问 `http://群晖IP:你的端口`。

### 可选：启用缓存持久化（重建容器不丢缓存）

```yaml
version: "3.9"
services:
  wordreplace:
    image: ghcr.io/marod1m/wordreplace:v1.6.4
    container_name: wordreplace
    restart: unless-stopped
    ports:
      - "12344:8501"
    volumes:
      - /volume1/docker/wordreplace/cache:/home/app/.cache/batch_replacer
```

---

## 🧭 使用示例（真实场景）

### 场景：批量生成录用通知书

Word 模板中包含：`【姓名】`、`【部门】`、`【入职日期】`。  
Excel 列名对应：`姓名`、`部门`、`入职日期`。

### 操作步骤

1. 上传 `offer_template.docx`
2. 上传 `offer_data.xlsx`
3. 添加规则：
   - `【姓名】 -> 姓名`
   - `【部门】 -> 部门`
   - `【入职日期】 -> 入职日期`
4. 选择行范围（例如 1 到 200）
5. 点击“开始替换”并下载结果

### 输出文件

- `张三.docx`
- `李四.docx`
- `批量替换统计.csv`
- `操作日志.txt`

---

## 🧪 本地开发与测试

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python -m pip install pytest
.venv/bin/streamlit run app/main.py
```

访问：`http://localhost:8501`

质量校验：

```bash
.venv/bin/python -m py_compile app/main.py app/core_utils.py app/services.py
.venv/bin/python -m pytest -q tests
```

---

## 🔐 安全与镜像策略

- 基础镜像：`python:3.12-alpine3.22`
- 多阶段构建 + 非 root 用户运行
- 构建/运行阶段执行系统包更新
- CI 自动执行语法检查与单元测试
- GHCR 标签：`latest`、`semver`（如 `1.6.4`）、`sha`

---

## 🏷️ 版本发布

```bash
git add -A
git commit -m "release: v1.6.4"
git push origin main

git tag -a v1.6.4 -m "Release v1.6.4"
git push origin v1.6.4
```

---

## 📚 附加文档

- 上手文档：`GETTING_STARTED.md`
- CI 工作流：`.github/workflows/docker-publish.yml`
- 进阶部署示例：`docker-compose.example.yml`
