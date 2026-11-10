# 贡献指南

感谢参与 WordReplace。本文说明本地开发、测试、构建、版本同步和提交前的检查项。

> 返回 [README](README.md) · 架构见 [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) · API 见 [docs/API.md](docs/API.md)

---

## 1. 环境要求

| 工具 | 版本 |
|---|---|
| Python | 3.12（与镜像一致） |
| Node.js | 20（与镜像一致） |
| pnpm | 10.17.1（`packageManager` 已固定） |
| Docker | 可选，用于验证镜像构建 |
| Git | 最新稳定版 |

检查：

```bash
python3 --version
node --version
pnpm --version
docker --version
```

---

## 2. 获取代码

```bash
git clone https://github.com/MaroD1M/WordReplace.git
cd WordReplace
```

如果使用自己的 fork，请把地址换成你的仓库地址。

---

## 3. 后端开发

### 3.1 创建虚拟环境

```bash
cd backend
python3 -m venv .venv
source .venv/bin/activate
```

Windows PowerShell：

```powershell
cd backend
py -3.12 -m venv .venv
.\.venv\Scripts\Activate.ps1
```

### 3.2 安装依赖

```bash
pip install --upgrade pip
pip install -r requirements.txt
```

运行测试时再安装开发依赖：

```bash
pip install -r requirements-dev.txt
```

### 3.3 启动

```bash
uvicorn app.main:app --reload --port 8000
```

- 健康检查：`http://localhost:8000/health`
- 就绪检查：`http://localhost:8000/health/ready`
- Swagger UI：`http://localhost:8000/docs`
- 默认数据库路径来自 `DATABASE_URL`（代码默认 `/app/data/wordreplace.db`，仅适合容器）。
- 本地开发建议启动前设置 `DATABASE_URL=sqlite:///./dev.db`。

示例：

```bash
DATABASE_URL=sqlite:///./dev.db uvicorn app.main:app --reload --port 8000
```

---

## 4. 前端开发

### 4.1 安装依赖

```bash
cd frontend
cp .env.local.example .env.local
pnpm install --ignore-scripts
```

> 项目在 `pnpm-workspace.yaml` 中忽略了 `sharp` 和 `unrs-resolver` 的构建脚本，因此使用 `--ignore-scripts`。

### 4.2 启动

```bash
pnpm dev
```

默认访问 `http://localhost:3000`。开发时 `.env.local` 中的 `NEXT_PUBLIC_API_BASE` 应指向后端：

```env
NEXT_PUBLIC_API_BASE=http://localhost:8000
```

### 4.3 代码位置

| 目录 | 说明 |
|---|---|
| `frontend/src/app/` | 页面与布局 |
| `frontend/src/components/` | 上传、规则、执行、结果等组件 |
| `frontend/src/hooks/` | 预览、配置、显示设置 |
| `frontend/src/lib/api.ts` | API 客户端与类型 |
| `frontend/src/lib/constants.ts` | 版本、分页、序号格式等常量 |

修改交互时，请同时检查：

- 空文件、空 Excel、缺失列等边界。
- 移动端与窄屏布局。
- 执行后结果失效逻辑（输入变化会清空旧结果）。

---

## 5. 测试

### 5.1 后端测试

```bash
cd backend
source .venv/bin/activate
pip install -r requirements-dev.txt
python -m pytest -q
```

只跑某个文件：

```bash
python -m pytest tests/test_run_replace.py -q
```

测试使用独立的 `backend/test_wordreplace.db`，由 `tests/conftest.py` 自动创建和清理。

### 5.2 前端检查

```bash
cd frontend
pnpm install --frozen-lockfile --ignore-scripts
pnpm lint
pnpm build
```

`pnpm build` 使用 Next.js 静态导出，必须成功，否则 Docker 镜像前端阶段会失败。

### 5.3 完整本地校验

```bash
# 版本一致
python3 scripts/sync_version.py --check

# 后端测试
cd backend
python -m pytest -q
cd ..

# 前端
cd frontend
pnpm lint
pnpm build
```

---

## 6. 版本号同步

根目录 `VERSION` 是唯一版本源。同步脚本会更新：

- `backend/app/core/config.py`
- `frontend/package.json`
- `frontend/src/lib/constants.ts`
- `Dockerfile`
- `README.md` 中的镜像标签与示例版本

修改版本：

```bash
# 1. 编辑 VERSION，例如改成 3.0.0
# 2. 同步
python3 scripts/sync_version.py
# 3. 校验
python3 scripts/sync_version.py --check
```

不要把版本号直接手改到各个文件里，否则 CI 会失败。

---

## 7. Git 钩子

仓库自带 `pre-commit` 钩子，用于扫描暂存区里常见的密钥格式。

启用：

```bash
git config core.hooksPath .githooks
```

确认：

```bash
git config --get core.hooksPath
```

输出应为：

```text
.githooks
```

钩子依赖 `rg`（ripgrep）。如果没有安装：

```bash
# Debian / Ubuntu
sudo apt-get install ripgrep

# macOS
brew install ripgrep
```

如果钩子误报，请先确认不是真实密钥，再调整匹配规则或临时跳过；不要直接提交密钥。

---

## 8. 提交前检查清单

- [ ] 代码能本地启动，页面可以完成一次预览、执行、下载。
- [ ] 后端测试通过：`cd backend && python -m pytest -q`。
- [ ] 前端 lint 通过：`cd frontend && pnpm lint`。
- [ ] 前端构建通过：`cd frontend && pnpm build`。
- [ ] 版本一致：`python3 scripts/sync_version.py --check`。
- [ ] 没有提交 `.env`、数据库、日志、证书、密钥或用户文档。
- [ ] 新增或修改接口时，同步更新 `docs/API.md`。
- [ ] 修改部署方式时，同步更新 `docs/DEPLOYMENT.md` 和 `README.md`。
- [ ] 修改用户可见行为时，同步更新 `CHANGELOG.md` 和 `ACCEPTANCE_CHECKLIST.md`。
- [ ] 文档中的命令可以复制执行，内部链接指向真实文件。
- [ ] 没有引入过时的 Compose 顶层 `version` 字段。

---

## 9. 提交信息建议

使用清晰的前缀，例如：

```text
feat: 支持多工作表选择
fix: 修复空单元格保留占位符
docs: 补充反向代理示例
test: 增加分页明细接口测试
chore: 同步版本号到 3.0.0
```

一次提交尽量只做一件事，便于审查和回滚。

---

## 10. 提交与 PR

1. 从 `main` 创建分支：

```bash
git switch -c feat/sheet-selection
```

2. 提交修改：

```bash
git add .
git commit -m "feat: 支持多工作表选择"
```

3. 推送并创建 Pull Request：

```bash
git push -u origin feat/sheet-selection
```

PR 描述中请说明：

- 改了什么。
- 为什么改。
- 如何验证。
- 是否影响配置、接口、数据或部署。

CI 会执行：

- 版本一致性检查。
- 后端语法检查与 pytest。
- 前端 lint 与 build。
- 多架构镜像构建（合并后）。

---

## 11. 相关文档

- [架构说明](docs/ARCHITECTURE.md)
- [API 参考](docs/API.md)
- [部署指南](docs/DEPLOYMENT.md)
- [故障排查](docs/TROUBLESHOOTING.md)
- [升级与回滚](docs/UPGRADE.md)
- [安全说明](SECURITY.md)
- [更新日志](CHANGELOG.md)
