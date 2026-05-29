# 文档批量替换助手 快速开始

> 版本：以根目录 `VERSION` 为准（由构建流程自动注入）  
> 架构：FastAPI + Next.js + SQLite  
> 推荐：单镜像部署

## 1) 最快启动（单镜像）

```bash
docker compose up -d --build
```

访问：`http://localhost:12344`

## 2) 使用已发布镜像

```bash
docker compose -f docker-compose.example.yml pull
docker compose -f docker-compose.example.yml up -d
```

## 3) 页面操作顺序

1. 上传 Word 模板与 Excel 数据
2. 新增替换规则
3. 设置起始行、结束行、文件名列
4. 点击开始替换
5. 下载 ZIP 或合并文档

### 示例文件说明（推荐先体验）

页面上传区可直接下载两组内置示例文件（单页模板 + 多页模板，均约 15 条数据）。  
说明：示例数据均为随机生成的模拟数据，仅用于功能演示，不包含真实个人信息。

## 4) 常见问题

- 执行按钮不可点：请确认已上传两个文件、至少一条规则、行号范围有效。
- 导出失败：请先执行替换并确认生成 `run_id`。
- Excel 列名不生效：请确保规则列名与 Excel 表头完全一致。

## 5) 源码开发模式（可选）

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
