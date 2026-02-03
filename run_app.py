import os
import sys
import subprocess
import webbrowser
import time
from threading import Thread

def check_streamlit():
    try:
        import streamlit
        return True
    except ImportError:
        return False

def run_streamlit():
    try:
        cmd = [sys.executable, '-m', 'streamlit', 'run', 'app/main.py', '--server.port=8501', '--server.address=127.0.0.1']
        subprocess.run(cmd, check=True)
    except Exception as e:
        print(f"启动失败: {e}")
        input("按回车键退出...")

def open_browser():
    time.sleep(3)
    webbrowser.open('http://127.0.0.1:8501')

if __name__ == '__main__':
    print("=" * 60)
    print("  Word+Excel 批量替换工具 v1.5.4")
    print("=" * 60)
    print()
    print("正在启动应用...")
    print("应用将在浏览器中打开: http://127.0.0.1:8501")
    print()
    print("按 Ctrl+C 停止应用")
    print("=" * 60)
    print()
    
    if not check_streamlit():
        print("错误: 未找到 Streamlit，请确保已安装所有依赖")
        input("按回车键退出...")
        sys.exit(1)
    
    browser_thread = Thread(target=open_browser, daemon=True)
    browser_thread.start()
    
    try:
        run_streamlit()
    except KeyboardInterrupt:
        print("\n应用已停止")
    except Exception as e:
        print(f"\n发生错误: {e}")
        input("按回车键退出...")
