# WordReplace

高可靠的 Word + Excel 批量替换工具（Streamlit）。

## 快速入口

- 开箱即用指南：`GETTING_STARTED.md`
- 标准部署文件：`docker-compose.yml`
- 示例部署文件：`docker-compose.example.yml`

## 一键启动

```bash
docker compose up -d
```

访问：`http://localhost:12344`

## 本地开发

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/streamlit run app/main.py
```

## 质量校验

```bash
python3 -m py_compile app/main.py app/core_utils.py app/services.py
python3 -m pytest -q tests
```
