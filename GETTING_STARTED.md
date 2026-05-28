# WordReplace 开箱即用指南

> 版本：v1.6.3  
> 适用对象：希望在 5 分钟内部署并开始批量替换 Word 模板的同学

## 1. 项目定位

WordReplace 是一个前后端一体化应用（Streamlit Web UI + Python 业务层），用于把 `.docx` 模板和 `.xlsx` 数据自动批量生成目标文档，支持 ZIP、合并文档与统计导出。

## 2. 一条命令启动（推荐）

```bash
docker compose -f docker-compose.example.yml up -d
```

访问 `http://localhost:8501`。

停止：

```bash
docker compose -f docker-compose.example.yml down
```

## 3. 标准部署

```bash
docker compose up -d
```

访问 `http://localhost:12344`。

## 4. 本地开发

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python -m pip install pytest
.venv/bin/streamlit run app/main.py
```

## 5. 本次安全修复

- 基础镜像升级到 `python:3.12-alpine3.22`
- 构建/运行阶段执行 `apk upgrade --no-cache`
- 容器改为非 root 用户 `app`
- 多阶段构建降低运行层暴露面
- CI 增加依赖安装、语法检查、单元测试
- Compose 文件增加健康检查

## 6. 验证

```bash
python3 -m py_compile app/main.py app/core_utils.py app/services.py
python3 -m pytest -q tests
```

## 7. 示例使用流程

1. 上传 Word 模板（`.docx`）
2. 上传 Excel 数据（`.xlsx`）
3. 配置关键词与列映射
4. 执行批量替换并下载结果
