#!/usr/bin/env bash
# run_web.sh for UniData-Sieve
echo "正在启动 UniData-Sieve 数据提纯引擎 Web 界面..."

# 检查虚拟环境是否存在，如果存在则自动激活
if [ -f ".venv/bin/activate" ]; then
    echo "[INFO] 找到虚拟环境，正在激活..."
    source .venv/bin/activate
elif [ -f ".venv/Scripts/activate" ]; then
    echo "[INFO] 找到 Window 格式的虚拟环境，尝试激活..."
    source .venv/Scripts/activate
else
    echo "[WARNING] 未找到 .venv 虚拟环境，使用系统默认 Python。请确保已安装 requirements.txt 依赖。"
fi

echo "[INFO] 正在启动 Streamlit 服务..."
streamlit run app.py
