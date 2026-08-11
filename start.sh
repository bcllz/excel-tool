#!/bin/bash
# =====================================================
# Excel 百宝箱 — 启动脚本 (Mac / Linux)
# 首次运行自动安装依赖，之后直接启动。
# 启动后浏览器打开 http://localhost:8501
# =====================================================

cd "$(dirname "$0")"

# 查找 Python
PYTHON=""
for p in python3 python; do
    if command -v "$p" &>/dev/null; then
        PYTHON="$p"
        break
    fi
done

if [ -z "$PYTHON" ]; then
    echo "❌ 未找到 Python，请先安装 Python 3.9+"
    echo "   下载地址：https://www.python.org/downloads/"
    exit 1
fi

echo "🐍 Python: $($PYTHON --version)"

# 首次运行：创建虚拟环境
if [ ! -d "venv" ]; then
    echo "📦 [1/3] 正在创建虚拟环境..."
    $PYTHON -m venv venv
fi

# 激活
source venv/bin/activate

# 安装依赖（首次）
if [ ! -f "venv/.deps_ok" ]; then
    echo "📥 [2/3] 正在安装依赖（首次需要几分钟）..."
    pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple --trusted-host pypi.tuna.tsinghua.edu.cn 2>/dev/null || \
    pip install -r requirements.txt
    touch venv/.deps_ok
    echo "✅ [2/3] 依赖安装完成"
fi

echo "🚀 [3/3] 正在启动..."
echo ""
echo "============================================"
echo "   浏览器打开 → http://localhost:8501"
echo "   按 Ctrl+C 停止"
echo "============================================"
echo ""

streamlit run app.py --server.port 8501 --server.headless true
