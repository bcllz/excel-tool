"""
Excel 百宝箱 — 桌面启动器
=========================
双击 exe 后启动 Streamlit 服务并自动打开浏览器。
兼容 PyInstaller 打包和直接运行两种模式。

【重要】Streamlit 1.50 的 server 是在【当前进程】内用 asyncio 跑的，
不需要、也不应该用 subprocess 再起子进程。
打包成 exe 后 sys.executable 指向 exe 自己，subprocess 会变成
「exe 自己启动自己」的无限递归，表现为不停地弹出浏览器标签页。
所以这里直接调用 streamlit.web.bootstrap.run()，在当前进程内启动。
"""

import os
import sys
import time
import socket
import threading
import webbrowser


def find_free_port(start=8501):
    """找一个空闲端口，默认 8501，被占用就 +1。"""
    port = start
    while port < 8520:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            if s.connect_ex(("127.0.0.1", port)) != 0:
                return port
        port += 1
    return 8501  # fallback


def main():
    # ---------- 确定 app.py 位置 ----------
    if getattr(sys, "frozen", False):
        # PyInstaller 打包后：sys._MEIPASS 是解压出来的资源目录
        base_dir = sys._MEIPASS
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))

    app_path = os.path.join(base_dir, "app.py")

    if not os.path.isfile(app_path):
        print(f"❌ 找不到 app.py，路径：{app_path}")
        print("按任意键退出...")
        input()
        return

    port = find_free_port(8501)

    # ---------- 启动前提示 ----------
    print("=" * 50)
    print("  📊 Excel 百宝箱 — 正在启动...")
    print("=" * 50)
    print()
    print(f"  浏览器即将打开 → http://localhost:{port}")
    print()
    print("  ⚠️ 请勿关闭此窗口，关闭即停止程序。")
    print("=" * 50)

    # ---------- 延迟打开浏览器（等服务起来再开，只开一次）----------
    def open_browser_later():
        time.sleep(5)
        webbrowser.open(f"http://localhost:{port}")

    threading.Thread(target=open_browser_later, daemon=True).start()

    # ---------- 直接在当前进程内启动 Streamlit（不要再起子进程）----------
    from streamlit import config as st_config
    from streamlit.web import bootstrap

    # headless=True：不让 streamlit 自己弹浏览器，由上面的线程统一开
    st_config.set_option("server.port", port)
    st_config.set_option("server.headless", True)
    st_config.set_option("browser.serverAddress", "localhost")
    st_config.set_option("server.runOnSave", False)
    st_config.set_option("server.fileWatcherType", "none")  # 避免 watchdog 依赖问题

    bootstrap.run(app_path, is_hello=False, args=[], flag_options={})

    print()
    print("  👋 已停止，感谢使用 Excel 百宝箱！")


if __name__ == "__main__":
    main()
