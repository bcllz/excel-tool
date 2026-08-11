"""
Excel 百宝箱 — 桌面启动器
=========================
双击 exe 后启动 Streamlit 服务并自动打开浏览器。
兼容 PyInstaller 打包和直接运行两种模式。
"""

import os
import sys
import time
import socket
import webbrowser
import subprocess
import threading


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
        # PyInstaller 打包后：sys._MEIPASS 是解压目录
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

    # ---------- 启动 Streamlit ----------
    print("=" * 50)
    print("  📊 Excel 百宝箱 — 正在启动...")
    print("=" * 50)
    print()
    print(f"  浏览器即将打开 → http://localhost:{port}")
    print()
    print("  ⚠️ 请勿关闭此窗口，关闭即停止程序。")
    print("=" * 50)

    # 用线程启动 streamlit
    def run_streamlit():
        # PyInstaller 打包后 sys.executable 指向 exe 自身
        # 需要特殊处理：如果是 frozen，用打包内的 Python
        if getattr(sys, "frozen", False):
            # 打包模式下使用 bootstrap 启动
            streamlit_args = [
                sys.executable,
                "-c",
                f"import sys; sys.argv = ['streamlit', 'run', {app_path!r}, "
                f"'--server.port', '{port}', '--server.headless', 'true', "
                f"'--browser.serverAddress', 'localhost']; "
                "from streamlit.web import cli; cli.main()",
            ]
        else:
            # 开发模式
            streamlit_args = [
                sys.executable, "-m", "streamlit", "run", app_path,
                "--server.port", str(port),
                "--server.headless", "true",
            ]

        subprocess.run(streamlit_args)

    server_thread = threading.Thread(target=run_streamlit, daemon=True)
    server_thread.start()

    # 等 3 秒让服务启动
    time.sleep(3)

    # 打开浏览器
    webbrowser.open(f"http://localhost:{port}")

    # 保持主线程存活
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("\n\n  👋 已停止，感谢使用 Excel 百宝箱！")
        print("  按任意键关闭窗口...")
        input()


if __name__ == "__main__":
    main()
