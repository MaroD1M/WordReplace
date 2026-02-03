# Word+Excel 批量替换工具

> 基于 Streamlit 的 Word 模板与 Excel 数据批量替换工具，支持保留格式、批量导出

![Version](https://img.shields.io/badge/version-v1.5.4-blue)
![Python](https://img.shields.io/badge/python-3.10+-green)
![License](https://img.shields.io/badge/license-MIT-orange)

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
  ghcr.io/MaroD1M/WordReplace:v1.5.4
```

### 方式三：群晖 Docker 部署

#### 1. 安装 Docker 套件

在群晖的 **套件中心** 中安装 **Docker** 套件。

#### 2. 下载镜像

打开 **Docker** -> **注册表**，搜索并下载镜像：

- 镜像名称：`ghcr.io/MaroD1M/WordReplace`
- 标签：`latest` 或 `v1.5.4`

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

### 方式五：Windows 可执行文件（推荐给小白用户）

#### 1. 下载预编译的 EXE 文件

从 [Releases](https://github.com/MaroD1M/WordReplace/releases) 页面下载最新的 `WordReplace.exe` 文件。

#### 2. 直接运行

双击 `WordReplace.exe` 即可启动应用，应用会自动在浏览器中打开。

**注意：**
- 首次运行可能需要几分钟时间启动
- Windows Defender 可能会提示安全警告，请选择"仍要运行"
- 不需要安装 Python 或任何依赖

#### 3. 自行编译（可选）

如果你想自己编译 EXE 文件，请按照以下步骤操作：

**前置要求：**
- Windows 操作系统
- Python 3.10 或更高版本

**编译步骤：**

1. 下载项目源码
   ```bash
   git clone https://github.com/MaroD1M/WordReplace.git
   cd WordReplace
   ```

2. 双击运行 `build.ps1` 脚本

   脚本会自动完成以下操作：
   - 检查 Python 环境
   - 安装 PyInstaller 和应用依赖
   - 编译生成 EXE 文件
   - 清理临时文件

   **注意：** 如果 PowerShell 提示执行策略错误，请先运行：
   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```

3. 编译完成后，可执行文件位于 `dist/WordReplace.exe`

4. 你可以将 `dist/WordReplace.exe` 复制到任何位置使用

**手动编译（高级用户）：**

如果 `build.bat` 脚本无法运行，可以手动执行以下命令：

```bash
# 安装 PyInstaller
pip install pyinstaller

# 安装应用依赖
pip install -r requirements.txt

# 编译
pyinstaller --clean WordReplace.spec

# 可执行文件位于 dist/WordReplace.exe
```

### 方式六：本地运行（开发环境）

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

### 使用 GitHub Actions 自动构建

如果你 fork 了本仓库，可以启用 GitHub Actions 自动构建：

1. 进入你的仓库 -> **Settings** -> **Actions** -> **General**
2. 在 **Workflow permissions** 中选择 **Read and write permissions**
3. 提交代码或创建标签（如 `v1.5.5`）会自动触发构建
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

### Q: Windows EXE 文件无法运行怎么办？

A: 请检查：
- 确保下载的是完整的 `WordReplace.exe` 文件
- Windows Defender 可能会拦截，请选择"仍要运行"
- 右键点击文件，选择"属性"，点击"解除锁定"
- 确保你的系统是 Windows 7 或更高版本

### Q: EXE 文件启动很慢正常吗？

A: 是的，首次启动可能需要 2-5 分钟，这是因为：
- 需要解压内置的 Python 环境
- 需要初始化 Streamlit 应用
- 后续启动会快很多

### Q: EXE 文件会被杀毒软件误报吗？

A: 由于 PyInstaller 打包的特性，某些杀毒软件可能会误报。这是正常现象，你可以：
- 将文件添加到杀毒软件的白名单
- 从官方 GitHub Releases 下载以确保安全
- 查看文件哈希值验证完整性

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
- **打包工具**：PyInstaller（用于 Windows EXE）

## 下载

### Windows 用户

从 [Releases](https://github.com/MaroD1M/WordReplace/releases) 页面下载：
- `WordReplace.exe` - Windows 可执行文件（推荐）

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
