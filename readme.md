# Word-Excel 批量替换工具

# 📋 Word+Excel 批量替换工具

🚀 **智能文档批量处理工具** - 让Word和Excel的批量替换变得简单高效！

## ✨ 核心功能

- 🔄 **批量替换** - 基于Excel数据批量替换Word文档内容

- 📊 **表格支持** - 完美支持Word表格内的文字替换

- 🎨 **格式保留** - 替换后保持原有字体、颜色、样式不变

- 🖥️ **可视化界面** - 友好的Web操作界面，无需编程经验

- 📦 **灵活下载** - 支持单个文件下载和ZIP压缩包批量下载

- 💾 **结果持久化** - 替换结果在会话期间持久保存

- 📤📥 **替换规则导入/导出** - 支持JSON格式导入导出替换规则，方便重复使用

- 🔢 **数值精度修复** - 智能修复Excel中公式计算后数值精度问题（如0.48729999999999996被修复为0.4873）
- 📊 **列类型智能识别** - 针对"合计"等特殊列进行专门处理，确保数值格式正确
- 🔍 **高精度计算** - 使用Decimal类型进行精确计算，避免浮点数二进制表示导致的精度损失

## 🚀 快速开始

### 🐳 Docker 一键部署（推荐）

```Bash
version: '3.8'
services:
  word-excel-replace:
    image: ghcr.io/marod1m/wordreplace:latest
    container_name: Wordreplace-tool
    network_mode: bridge
    ports:
      - "12344:8501"
    restart: no
    environment:
      - STREAMLIT_SERVER_HEADLESS=true
      - STREAMLIT_BROWSER_GATHER_USAGE_STATS=false

```

### ⚡ 本地运行

```Bash
git clone https://github.com/你的用户名/wordreplace.git
cd wordreplace
pip install -r requirements.txt
streamlit run app/main.py
```

## 📖 使用流程

1. 📄 **上传文件**

    - Word模板：上传.docx格式的模板文档

    - Excel数据：上传.xlsx/.xls格式的数据文件

2. 👀 **预览文档**

    - 查看Word文档内容（含表格）

    - 选中关键字并按Ctrl+C复制

    - 预览Excel数据结构

3. ⚙️ **设置规则**

    - 粘贴关键字到输入框

    - 选择对应的Excel数据列

    - 添加替换规则（可添加多个）

4. 🚀 **执行替换**

    - 设置文件名格式和前缀

    - 选择替换范围（全部行或指定行）

    - 点击开始批量替换

5. 📥 **下载结果**

    - 单文件下载：分页显示，逐一下载

    - 批量下载：ZIP压缩包一键下载

## 🛠️ 技术栈

|技术|用途|
|---|---|
|🐍 Python 3.10+|后端逻辑处理|
|🎈 Streamlit 1.51.0|Web界面框架|
|📊 Pandas|数据处理|
|📄 python-docx|Word文档处理|
|📊 openpyxl|Excel文件处理|
|🐳 Docker|容器化部署|
## ⚡ 使用建议

### 💡 最佳实践：

- 单次处理建议不超过1000行数据

- 文件大小建议控制在50MB以内

- 确保服务器有2GB+可用内存

- 大文件建议分批次处理

### 🛡️ 安全提示：

- 所有处理在内存中进行

- 会话结束自动清理数据

- 建议在可信网络环境使用

⭐ 如果这个项目对你有帮助，请给我们一个Star！

🌐 访问地址：[http://localhost:12344](http://localhost:12344)（部署后）

## 📝 更新日志

### v1.1.0（数值精度修复增强版）
- **核心功能**：修复Excel数值精度问题，确保正确显示（如0.48729999999999996 → 0.4873）
- **技术改进**：使用Decimal类型进行精确计算，避免浮点数二进制表示导致的精度损失
- **智能适配**：针对"合计"等特殊列进行专门处理，自动优化小数位数
- **兼容性**：保持对文字内容的正确读取和处理，不会影响非数值数据
- **性能优化**：优化数值处理逻辑，提高修复效率

> （文档及代码内容由 AI 生成）