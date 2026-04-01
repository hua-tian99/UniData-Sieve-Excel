import sys
import threading
import time
import webbrowser
from pathlib import Path
import os

def print_welcome():
    print("======================================================")
    print("               UniData-Sieve 数据提纯引擎               ")
    print("======================================================")
    print("正在为您激活动力核心，挂载本地服务引擎，请稍候...\n")

def open_browser():
    # 延迟等待底层的 Tornado HTTP 服务器完全监听 8501 端口
    time.sleep(4)
    print("✅ 引擎全功率运转，服务已在后台同步启动！")
    print("🌐 正在通过系统默认浏览器唤起可是化操作面板...")
    print("\n⚠️ 警告：请保持黑色窗口运行！若关闭此窗口，本地提纯服务将立刻断开连线！")
    print("======================================================")
    webbrowser.open('http://127.0.0.1:8501')

if __name__ == '__main__':
    try:
        print_welcome()
        
        # 坑点防御 1：必须在进程内启动 Streamlit CLI
        import streamlit.web.cli as stcli
        
        # 坑点防御 2：路径寻址！
        if getattr(sys, 'frozen', False):
            # 兼容不同版本的 PyInstaller onedir 寻址
            base_path = Path(getattr(sys, '_MEIPASS', os.path.dirname(sys.executable)))
            app_path = base_path / "app.py"
            os.chdir(base_path)
            
            # 由于可能出现端口占用引发的新开页端口不为8501的问题，直接让Streamlit输出到控制台即可
            sys.argv = [
                "streamlit", 
                "run", 
                str(app_path), 
                "--global.developmentMode=false",
                "--server.headless=true",
                "--server.address=127.0.0.1",
                "--server.port=8501"
            ]
        else:
            app_path = Path(__file__).parent / "app.py"
            sys.argv = [
                "streamlit", 
                "run", 
                str(app_path), 
                "--global.developmentMode=false",
                "--server.headless=true",
                "--server.address=127.0.0.1",
                "--server.port=8501"
            ]

        print(f"即将启动引擎: {app_path}")
        
        # 打开一个独立的非阻塞监控幽灵线程，倒计时后强制打开网页
        threading.Thread(target=open_browser, daemon=True).start()

        # 移交接管，不直接用 sys.exit 因为如果报错它会跳过最后的 input()
        exit_code = stcli.main()
        print(f"\n🛑 Streamlit 引擎已停止，退出代码: {exit_code}")
        input("按下回车键（Enter）退出控制台...")
        
    except BaseException as e:
        import traceback
        print("\n❌ 启动引擎时发生严重崩溃，以下是崩溃日志：")
        traceback.print_exc()
        print(f"\n⚠️ 前方高能！引擎未能正常启动！异常类型: {type(e)}")
        input("按下回车键（Enter）退出控制台...")
