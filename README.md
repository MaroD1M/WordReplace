# Word+Excel 批量替换工具

> 一键把 Word 模板和 Excel 数据拼起来：保留格式、批量生成、可追踪可导出。

![Version](https://img.shields.io/badge/version-v1.6.0-blue) ![Python](https://img.shields.io/badge/python-3.10+-green) ![License](https://img.shields.io/badge/license-MIT-orange)

## 这是什么？

这是一个基于 Streamlit 的批量替换工具，适合做通知、证书、函件、合同等“模板 + 名单”场景：

- 上传 `.docx` 模板
- 上传 `.xlsx` 数据
- 配置关键字与列映射
- 一次性生成全部文档
- 支持 ZIP、合并文档、统计 CSV、日志 TXT 导出

## 核心能力

- **两种替换模式**：完整关键词替换、仅替换括号内内容
- **格式尽量保留**：支持段落与表格中的替换
- **规则管理完整**：新增/删除/撤销、导入导出 JSON、本地缓存
- **流程提示清晰**：5 步指引 + 执行前阻塞原因提示
- **结果可追溯**：替换统计、操作日志、历史记录
- **安全增强**：文件名清洗、缓存键过滤、规则导入大小限制、导出重名处理

---

## 快速开始

### 方式一：Docker Compose（推荐）

```bash
git clone https://github.com/你的用户名/WordReplace.git
cd WordReplace
docker-compose up -d
```

默认访问：`http://localhost:12344`

如需改端口，修改 `docker-compose.yml`：

```yaml
ports:
  - "8080:8501"
```

### 方式二：Docker 命令

```bash
docker run -d \
  --name WordReplace \
  -p 12344:8501 \
  -e STREAMLIT_SERVER_HEADLESS=true \
  -e STREAMLIT_BROWSER_GATHER_USAGE_STATS=false \
  ghcr.io/MaroD1M/WordReplace:latest
```

### 方式三：本地运行（开发）

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python -m pip install pytest
.venv/bin/streamlit run app/main.py
```

本地访问：`http://localhost:8501`

---

## 使用流程（和界面一致）

1. 上传 Word 模板（仅 `.docx`）
2. 上传 Excel 数据（仅 `.xlsx`，默认读取首个工作表）
3. 在右侧添加替换规则（关键字 → Excel 列）
4. 设置行范围、文件名前缀
5. 点击开始替换并下载结果

导出方式支持：

- 独立文件 ZIP
- 合并为单个 Word
- 统计 CSV
- 日志 TXT

---

## 项目结构（v1.6.0）

```text
app/
  main.py         # Streamlit UI 与状态编排
  core_utils.py   # 纯工具函数（文本/文件名/前置校验）
  services.py     # 业务服务层（替换、合并、统计）
tests/
  test_core_utils.py
  test_services.py
```

---

## 开发与测试

### 使用 Makefile（推荐）

```bash
make venv
make install
make run
make test
```

### 直接运行测试

```bash
.venv/bin/python -m pytest -q tests
```

当前测试覆盖：

- 文件名与缓存键安全处理
- 前置阻塞条件校验
- Excel 清洗与参数签名稳定
- 统计导出基础正确性

---

## Docker 自行构建（可选）

```bash
docker build -t word-replace:latest .
```

多平台构建：

```bash
docker buildx create --use
docker buildx build --platform linux/amd64,linux/arm64 -t word-replace:latest .
```

---

## 发布新版本

本仓库包含自动构建工作流：`.github/workflows/docker-publish.yml`

发布步骤：

```bash
git add .
git commit -m "feat: your change"
git push origin main

git tag -a v1.6.0 -m "Release v1.6.0"
git push origin v1.6.0
```

---

## 常见问题

### 支持哪些文件格式？

- Word：`.docx`
- Excel：`.xlsx`

### 为什么不支持 `.doc` / `.xls`？

当前实现基于 `python-docx` 和 `openpyxl`，为保证稳定性与一致性，仅支持现代 Office 格式。

### 数据会上传到外部服务器吗？

不会。应用在你部署的环境中运行；规则缓存和历史记录保存在本地缓存目录。

---

## 技术栈

- Streamlit
- Pandas
- python-docx
- openpyxl
- Docker
- GitHub Actions

## 版本历史

- **v1.6.0**：界面与交互优化、统一 `.xlsx` 支持、模块化（`core_utils` / `services`）、新增测试
- **v1.5.4**：缓存管理与预览性能优化
- **v1.2.4**：初始版本

## 许可证

MIT License
