# Word+Excel批量替换工具

## 🚀 项目简介

这是一个基于Streamlit的Web应用程序，用于批量处理Word文档和Excel数据，实现自动化文档生成。该工具可以从Word模板文件中提取内容，并使用Excel中的数据批量替换模板中的占位符，从而生成大量个性化的文档。

### ✨ 核心功能

- **批量替换** - 从Word模板中提取占位符，并用Excel中的数据进行批量替换
- **格式保留** - 保留原始Word文档的格式，包括字体、表格、颜色等
- **智能规则管理** - 支持创建、导入、导出和缓存替换规则
- **文件预览** - 在执行替换前预览Word和Excel文件内容
- **多输出选项** - 支持ZIP打包下载、合并为单个文档、导出统计数据等多种输出方式
- **历史记录** - 记录每次操作的历史，方便追踪和回溯
- **统计分析** - 提供替换统计信息和操作成功率
- **高性能处理** - 支持处理大容量文档和数据集

## 🛠️ 部署

### 本地运行

1. 克隆仓库
   ```bash
   git clone https://github.com/MaroD1M/WordReplace.git
   cd WordReplace
   ```

2. 安装依赖
   ```bash
   pip install -r requirements.txt
   ```

3. 运行应用
   ```bash
   streamlit run app/main.py
   ```

### Docker部署

1. 构建镜像
   ```bash
   docker build -t word-replace .
   ```

2. 运行容器
   ```bash
   docker run -d -p 8501:8501 word-replace
   ```

3. 访问应用：http://localhost:8501

### Docker Compose部署

1. 修改环境变量（可选）
   ```bash
   export GITHUB_USERNAME=your_username
   export TAG=latest
   export EXTERNAL_PORT=12344
   ```

2. 启动服务
   ```bash
   docker-compose up -d
   ```

3. 访问应用：http://localhost:8501

### GitHub Container Registry部署

使用预构建的镜像：
```bash
# 拉取最新版本
docker pull ghcr.io/marod1m/wordreplace:latest

# 拉取指定版本
docker pull ghcr.io/marod1m/wordreplace:v1.5.4
```

运行镜像：
```bash
docker run -d -p 8501:8501 ghcr.io/marod1m/wordreplace:latest
```

## 🚀 运行

### 本地运行

1. 确保已安装Python 3.10+
2. 克隆仓库并进入目录
3. 安装依赖：`pip install -r requirements.txt`
4. 运行：`streamlit run app/main.py`
5. 访问：http://localhost:8501

### PyCharm运行

1. 打开项目
2. 配置Python解释器（3.10+）
3. 在Terminal中运行：`streamlit run app/main.py`

### Docker运行

1. 构建镜像：`docker build -t word-replace .`
2. 运行容器：`docker run -d -p 8501:8501 word-replace`
3. 访问：http://localhost:8501

## 📋 功能介绍

### 主要功能

- **Word模板处理**：支持.docx格式的Word文档，保留原有格式
- **Excel数据源**：支持.xlsx/.xls格式的数据表，可预览内容
- **规则管理**：可视化创建替换规则，支持导入导出
- **批量处理**：一次处理多行数据，生成多个文档
- **格式保持**：处理过程中保持原文档的格式不变
- **缓存机制**：支持规则缓存，提高重复操作效率
- **历史记录**：记录操作历史，便于追溯和复用
- **统计功能**：显示处理进度和成功率

### 使用流程

1. **上传文件**：上传Word模板和Excel数据文件
2. **预览内容**：查看文档和数据的预览
3. **创建规则**：从Word预览中选择占位符，关联Excel列
4. **设置参数**：选择起始和结束行，设置文件名前缀
5. **开始替换**：执行批量替换操作
6. **下载结果**：选择ZIP下载、合并文档等方式获取结果

### 高级功能

- **智能替换**：支持多种括号格式（【】、（）、()等）
- **仅替换括号内内容**：可选择只替换括号内的文本
- **文件合并**：可将多个结果文档合并为一个
- **数据统计**：显示替换成功率和处理统计
- **缓存管理**：保存常用替换规则，方便下次使用
- **历史记录**：保留最近的操作记录

## ⚙️ 配置

### 环境变量

- `STREAMLIT_SERVER_HEADLESS` - 设为false禁用统计收集
- `STREAMLIT_BROWSER_GATHER_USAGE_STATS` - 设为false禁用统计收集

### 文件限制

- Word文件最大10MB
- Excel文件最大10MB
- 建议单次处理行数少于1000行

## ❓ 常见问题

### Q: 支持哪些类型的括号？
A: 支持（、）、(、)、【、】等常见括号格式，可选择仅替换括号内内容。

### Q: Docker容器启动时出现ModuleNotFoundError: No module named 'packaging'错误？
A: 这是由于缺少Streamlit的间接依赖packaging模块导致的。已在requirements.txt中添加此依赖，重新构建镜像即可解决：
```bash
docker build -t word-replace .
docker run -d -p 8501:8501 word-replace
```

### Q: 为什么Word文件不支持.doc格式？
A: 由于技术限制，当前版本仅支持.docx格式。请使用Word将.doc文件另存为.docx格式后再使用。

### Q: 如何加快处理速度？
A: 
- 分批处理（每批100-200行）
- 使用SSD硬盘
- 关闭其他程序

### Q: 缓存文件保存在哪里？
A: 
- Windows: %APPDATA%/BatchReplacer
- Mac/Linux: ~/.cache/batch_replacer

## 🛠️ 技术栈

|技术栈|用途|
|---|---|
|🐍 Python 3.10+|后端逻辑处理|
|🌐 Streamlit 1.52.2|Web界面框架|
|📊 Pandas 2.3.3|数据处理|
|📄 python-docx 1.2.0|Word文档处理|
|📊 openpyxl 3.1.5|Excel文件处理|
|🔧 lxml 6.0.2|XML处理|
|📦 packaging 26.0|依赖管理|
|🐳 Docker|容器化部署|


## 💡 使用建议

### 📌 最佳实践：

- 单次处理建议不超过1000行数据
- 文件大小建议控制在10MB以内
- 确保服务器有2GB+可用内存
- 大文件建议分批处理

### 🔒 安全提示：

- 所有处理在内存中进行
- 会话结束自动清理数据
- 建议在可信任网络环境使用

## 🤝 贡献

欢迎提交Issue和Pull Request来改进这个项目！

## 📄 许可证

MIT License

## 🆘 支持

😊 如果这个项目对您有帮助，请给我们一个Star！

📍 访问地址：[http://localhost:8501](http://localhost:8501)（本地运行后）
