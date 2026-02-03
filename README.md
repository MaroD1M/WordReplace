# Word+Excel 批量替换工具

> 基于 Streamlit 的 Word 模板与 Excel 数据批量替换工具，支持保留格式、批量导出

![Banner](assets/banner.svg)

![Version](https://img.shields.io/badge/version-v1.5.6-blue)   ![Python](https://img.shields.io/badge/python-3.10+-green)   ![License](https://img.shields.io/badge/license-MIT-orange)

## 功能特性

- 批量替换：Word 模板与 Excel 数据批量替换，完美保留格式
- 多种替换模式：支持完整关键词替换和括号内容替换
- 灵活导出：支持 ZIP 压缩包导出、合并导出为单个 Word 文档
- 规则管理：支持规则导入/导出、本地缓存、撤销操作
- 历史记录：自动记录操作历史，便于追踪和回溯
- 高性能：优化的预览机制，支持大文件处理
- 跨平台：支持 Windows、Linux、macOS

## 快速开始

### 方式一：使用 Docker Compose（推荐）

#### 1. Fork 本仓库

点击右上角的 **Fork** 按钮，将仓库 fork 到你的 GitHub 账号下。

#### 2. 克隆你的仓库

```bash
git clone https://github.com/你的用户名/WordReplace.git
cd WordReplace
```

#### 3. 启动服务

```bash
docker-compose up -d
```

服务将在 `http://localhost:12344` 启动。

#### 4. 自定义端口

修改 `docker-compose.yml` 中的端口映射：

```yaml
ports:
  - "你的端口:8501"  # 例如: "8080:8501"
```

### 方式二：使用 Docker 命令行

#### 使用官方镜像

```bash
docker run -d \
  --name WordReplace \
  -p 12344:8501 \
  -e STREAMLIT_SERVER_HEADLESS=true \
  -e STREAMLIT_BROWSER_GATHER_USAGE_STATS=false \
  ghcr.io/MaroD1M/WordReplace:latest
```

#### 使用指定版本镜像

```bash
docker run -d \
  --name WordReplace \
  -p 12344:8501 \
  -e STREAMLIT_SERVER_HEADLESS=true \
  -e STREAMLIT_BROWSER_GATHER_USAGE_STATS=false \
  ghcr.io/MaroD1M/WordReplace:v1.5.6
```

### 方式三：群晖 Docker 部署

#### 1. 安装 Docker 套件

在群晖的 **套件中心** 中安装 **Docker** 套件。

#### 2. 下载镜像

打开 **Docker** -> **注册表**，搜索并下载镜像：

- 镜像名称：`ghcr.io/MaroD1M/WordReplace`
- 标签：`latest` 或 `v1.5.6`

#### 3. 创建容器

1. 在 **映像** 标签页中，右键点击下载的镜像，选择 **启动**
2. 配置容器：
   - **容器名称**：`WordReplace`
   - **端口设置**：本地端口 `12344` 映射到容器端口 `8501`
   - **环境变量**：
     - `STREAMLIT_SERVER_HEADLESS=true`
     - `STREAMLIT_BROWSER_GATHER_USAGE_STATS=false`
3. 点击 **应用** 启动容器

#### 4. 访问应用

在浏览器中访问 `http://群晖IP:12344`

### 方式四：服务器命令行部署

#### Linux/Ubuntu

```bash
# 拉取镜像
docker pull ghcr.io/MaroD1M/WordReplace:latest

# 运行容器
docker run -d \
  --name WordReplace \
  --restart=unless-stopped \
  -p 12344:8501 \
  -e STREAMLIT_SERVER_HEADLESS=true \
  -e STREAMLIT_BROWSER_GATHER_USAGE_STATS=false \
  ghcr.io/MaroD1M/WordReplace:latest

# 查看日志
docker logs -f WordReplace

# 停止容器
docker stop WordReplace

# 启动容器
docker start WordReplace
```

#### 使用 systemd 管理（推荐）

创建服务文件 `/etc/systemd/system/word-replace.service`：

```ini
[Unit]
Description=Word+Excel Batch Replace Tool
After=docker.service
Requires=docker.service

[Service]
Type=oneshot
RemainAfterExit=yes
WorkingDirectory=/path/to/WordReplace
ExecStart=/usr/bin/docker-compose up -d
ExecStop=/usr/bin/docker-compose down
TimeoutStartSec=0

[Install]
WantedBy=multi-user.target
```

启用并启动服务：

```bash
sudo systemctl enable word-replace
sudo systemctl start word-replace
sudo systemctl status word-replace
```

### 方式五：本地运行（开发环境）

#### 1. 安装 Python 3.10+

确保你的系统已安装 Python 3.10 或更高版本。

#### 2. 安装依赖

```bash
pip install -r requirements.txt
```

#### 3. 运行应用

```bash
streamlit run app/main.py
```

应用将在 `http://localhost:8501` 启动。

## 使用方法

### 基本流程

1. **上传 Word 文件**
   - 点击左侧的 **上传 Word 文件** 按钮
   - 选择包含要替换内容的 `.docx` 文件（不支持 `.doc` 格式）

2. **上传 Excel 文件**
   - 点击 **上传 Excel 文件** 按钮
   - 选择包含替换数据的 `.xlsx` 或 `.xls` 文件

3. **预览文件内容**
   - 查看文件预览，确认数据格式正确
   - Excel 预览会显示前 50 行数据

4. **添加替换规则**
   - **新关键字**：从 Word 预览中复制要替换的关键字，如 `【姓名】`、`（部门）`
   - **Excel 列**：选择 Excel 中对应的列
   - 点击 **添加规则** 按钮

5. **设置行范围**
   - **起始行**：从第几行开始处理替换
   - **结束行**：处理到第几行（包括该行），默认到最后一行

6. **执行替换**
   - 点击 **开始替换** 按钮
   - 等待处理完成，查看替换统计信息

7. **导出结果**
   - **导出 ZIP**：将所有替换后的文件保存为一个 ZIP 压缩包
   - **合并导出**：将所有替换后的文件合并为一个 Word 文档，每个文件占一页
   - **导出统计**：导出替换统计数据为 CSV 格式
   - **导出日志**：导出详细的替换操作日志为 TXT 文件

### 高级功能

#### 替换模式选择

- **完整关键词**：直接替换整个关键词，如 `【姓名】` → `张三`
- **括号内容**：只替换括号内的文字，保留括号，如 `（部门）` → `（技术部）`

#### 文件名设置

- **文件名列**：选择 Excel 中的列用于生成文件名，通常选择唯一标识符列
- **文件前缀**：为生成的文件名添加前缀，如 `2024-` 会生成 `2024-文件名.docx`

#### 规则管理

- **导入规则**：从之前导出的 JSON 文件中导入替换规则
- **导出规则**：将当前规则导出为 JSON 文件，可以在其他电脑导入使用
- **保存到缓存**：快速保存规则到本地缓存，下次可以快速加载使用
- **撤销**：撤销最后一次规则操作（添加、删除等）
- **清空规则**：清空所有已添加的替换规则

#### 历史记录

- 自动记录每次操作的历史
- 查看历史记录了解之前的操作
- 清除历史记录释放空间

## 自行编译

如果你希望自行编译 Docker 镜像，可以按照以下步骤操作：

### 前置要求

- Docker 已安装
- Git 已安装

### 编译步骤

#### 1. 克隆仓库

```bash
git clone https://github.com/你的用户名/WordReplace.git
cd WordReplace
```

#### 2. 使用 Docker Buildx 构建多平台镜像

```bash
# 创建并使用 buildx 构建器
docker buildx create --use

# 构建并加载到本地
docker buildx build --platform linux/amd64,linux/arm64 -t word-replace:latest .

# 或者只构建当前平台
docker build -t word-replace:latest .
```

#### 3. 推送到 Docker Hub（可选）

```bash
# 登录 Docker Hub
docker login

# 标记镜像
docker tag word-replace:latest 你的用户名/word-replace:latest

# 推送镜像
docker push 你的用户名/word-replace:latest
```

#### 4. 推送到 GitHub Container Registry（推荐）

```bash
# 登录 GHCR
echo $GITHUB_TOKEN | docker login ghcr.io -u 你的用户名 --password-stdin

# 标记镜像
docker tag word-replace:latest ghcr.io/你的用户名/word-replace:latest

# 推送镜像
docker push ghcr.io/你的用户名/word-replace:latest
```

## 发布新版本

作为项目维护者，你可以通过推送版本标签来自动构建和发布新版本：

### 自动构建流程

当你推送版本标签时，GitHub Actions 会自动执行以下操作：

1. **构建 Docker 镜像**
   - 构建多平台镜像（linux/amd64, linux/arm64）
   - 推送到 GitHub Container Registry

2. **创建 Release**
   - 自动创建 GitHub Release
   - 生成版本说明

### 发布步骤

1. **确保所有更改已提交并推送**

   ```bash
   git add .
   git commit -m "feat: 新功能描述"
   git push origin main
   ```

2. **创建并推送版本标签**

   ```bash
   # 创建标签（使用语义化版本号）
   git tag -a v1.5.6 -m "Release v1.5.6 - 版本描述"

   # 推送标签到远程仓库
   git push origin v1.5.6
   ```

3. **等待自动构建完成**

   - 访问 [Actions](https://github.com/MaroD1M/WordReplace/actions) 页面查看构建进度
   - 构建完成后，访问 [Releases](https://github.com/MaroD1M/WordReplace/releases) 查看新版本

4. **验证发布**

   - 检查 Release 页面是否创建成功
   - 验证 Docker 镜像是否可用

### 版本号规范

请遵循语义化版本号规范（Semantic Versioning）：

- **主版本号（Major）**：不兼容的 API 修改
- **次版本号（Minor）**：向下兼容的功能性新增
- **修订号（Patch）**：向下兼容的问题修正

示例：
- `v1.5.6` - 修订版本（Bug 修复）
- `v1.6.0` - 次版本（新增功能）
- `v2.0.0` - 主版本（重大更新）

### Fork 用户的自动构建

如果你 fork 了本仓库，可以启用 GitHub Actions 自动构建：

1. 进入你的仓库 -> **Settings** -> **Actions** -> **General**
2. 在 **Workflow permissions** 中选择 **Read and write permissions**
3. 提交代码或创建标签（如 `v1.5.6`）会自动触发构建
4. 构建完成后，镜像会自动推送到你的 GitHub Container Registry

使用自动构建的镜像：

```bash
docker pull ghcr.io/你的用户名/WordReplace:latest
```

## 常见问题

### Q: 为什么上传的 Word 文件无法识别？

A: 请确保上传的是 `.docx` 格式文件，本工具不支持 `.doc` 格式。可以使用 Microsoft Word 或 WPS 将 `.doc` 转换为 `.docx`。

### Q: 替换后格式丢失了怎么办？

A: 本工具使用 python-docx 库，会尽量保留原始格式。如果遇到格式问题，建议：
- 在 Word 模板中使用标准格式
- 避免使用复杂的样式和嵌套表格
- 检查替换的关键字是否完整

### Q: 支持哪些 Excel 格式？

A: 支持 `.xlsx` 和 `.xls` 格式。建议使用 `.xlsx` 格式以获得更好的兼容性。

### Q: 如何处理大量数据？

A: 本工具支持大文件处理，但建议：
- 单个 Word 文件不超过 50MB
- Excel 数据行数建议在 10000 行以内
- 如果数据量很大，可以分批处理

### Q: 容器启动失败怎么办？

A: 请检查：
- 端口是否被占用：`netstat -an | grep 12344`
- Docker 服务是否正常运行：`docker ps`
- 查看容器日志：`docker logs WordReplace`

### Q: 如何更新到最新版本？

A: 使用以下命令更新：

```bash
docker-compose pull
docker-compose up -d
```

或使用 Docker 命令：

```bash
docker pull ghcr.io/MaroD1M/WordReplace:latest
docker stop WordReplace
docker rm WordReplace
docker run -d --name WordReplace -p 12344:8501 ghcr.io/MaroD1M/WordReplace:latest
```

### Q: 数据会保存在哪里？

A: 本工具的所有数据都在浏览器本地处理，不会上传到任何服务器。缓存规则和历史记录保存在容器的本地文件系统中，容器删除后数据会丢失。

### Q: 如何备份数据？

A: 建议定期：
- 导出替换规则为 JSON 文件
- 导出操作日志为 TXT 文件
- 将导出的文件保存到本地或云存储

## 技术栈

- **前端**：Streamlit
- **后端**：Python 3.10+
- **数据处理**：Pandas
- **Word 处理**：python-docx
- **Excel 处理**：openpyxl
- **容器化**：Docker
- **CI/CD**：GitHub Actions

## 下载

### Docker 用户

```bash
docker pull ghcr.io/MaroD1M/WordReplace:latest
```

### 源码用户

```bash
git clone https://github.com/MaroD1M/WordReplace.git
```

## 项目结构

```
WordReplace/
├── app/
│   └── main.py              # 主程序文件
├── requirements.txt         # Python 依赖
├── Dockerfile              # Docker 镜像构建文件
├── docker-compose.yml      # Docker Compose 配置
├── .github/
│   └── workflows/
│       └── docker-publish.yml  # GitHub Actions 配置
└── README.md               # 项目说明文档
```

## 版本历史

- **v1.5.6** - 优化部署流程，专注于 Docker 容器化部署
- **v1.5.4** - 最终版：规范的缓存管理、高性能预览、全面 Bug 修复
- **v1.2.4** - 初始版本

## 许可证

本项目采用 MIT 许可证。详见 [LICENSE](LICENSE) 文件。

## 贡献

欢迎提交 Issue 和 Pull Request！

## 联系方式

如有问题或建议，请通过以下方式联系：

- 提交 [Issue](https://github.com/MaroD1M/WordReplace/issues)
- 发送邮件

## 致谢

感谢所有为本项目做出贡献的开发者！

---

**注意**：本工具仅供学习和个人使用，请勿用于商业用途。
