# -*- coding: utf-8 -*-
# Word+Excel 批量替换工具 - 编译脚本

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Word+Excel 批量替换工具 - 编译脚本" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# 检查 Python
Write-Host "[1/5] 检查 Python 环境..." -ForegroundColor Yellow
$pythonCmd = Get-Command python -ErrorAction SilentlyContinue
if (-not $pythonCmd) {
    $pythonCmd = Get-Command py -ErrorAction SilentlyContinue
}

if (-not $pythonCmd) {
    Write-Host "错误: 未找到 Python，请先安装 Python 3.10+" -ForegroundColor Red
    Read-Host "按回车键退出"
    exit 1
}
Write-Host "✓ Python 环境检查通过" -ForegroundColor Green
Write-Host ""

# 安装 PyInstaller
Write-Host "[2/5] 安装编译依赖..." -ForegroundColor Yellow
& $pythonCmd.Source -m pip install pyinstaller -i https://pypi.tuna.tsinghua.edu.cn/simple
if ($LASTEXITCODE -ne 0) {
    Write-Host "错误: 安装 PyInstaller 失败" -ForegroundColor Red
    Read-Host "按回车键退出"
    exit 1
}
Write-Host "✓ PyInstaller 安装完成" -ForegroundColor Green
Write-Host ""

# 安装应用依赖
Write-Host "[3/5] 安装应用依赖..." -ForegroundColor Yellow
& $pythonCmd.Source -m pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
if ($LASTEXITCODE -ne 0) {
    Write-Host "错误: 安装应用依赖失败" -ForegroundColor Red
    Read-Host "按回车键退出"
    exit 1
}
Write-Host "✓ 应用依赖安装完成" -ForegroundColor Green
Write-Host ""

# 编译
Write-Host "[4/5] 开始编译..." -ForegroundColor Yellow
Write-Host "这可能需要几分钟时间，请耐心等待..." -ForegroundColor Gray
Write-Host ""
& $pythonCmd.Source -m PyInstaller --clean WordReplace.spec
if ($LASTEXITCODE -ne 0) {
    Write-Host "错误: 编译失败" -ForegroundColor Red
    Read-Host "按回车键退出"
    exit 1
}
Write-Host "✓ 编译完成" -ForegroundColor Green
Write-Host ""

# 清理临时文件
Write-Host "[5/5] 清理临时文件..." -ForegroundColor Yellow
if (Test-Path "build") {
    Remove-Item -Recurse -Force "build"
}
Write-Host "✓ 清理完成" -ForegroundColor Green
Write-Host ""

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  编译成功！" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "可执行文件位置: dist\WordReplace.exe" -ForegroundColor White
Write-Host ""
Write-Host "你可以:" -ForegroundColor White
Write-Host "  1. 直接运行 dist\WordReplace.exe" -ForegroundColor Gray
Write-Host "  2. 将 dist\WordReplace.exe 复制到其他电脑使用" -ForegroundColor Gray
Write-Host "  3. 将整个 dist 文件夹打包分享给他人" -ForegroundColor Gray
Write-Host ""
Write-Host "注意: 首次运行可能需要几分钟时间启动" -ForegroundColor Yellow
Write-Host ""
Read-Host "按回车键退出"
