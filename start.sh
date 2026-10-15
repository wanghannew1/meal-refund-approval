#!/usr/bin/env bash
# Linux launcher for 返餐费审批提交工具 (server deployment).
# Mirrors start.bat logic: venv-first, dep version check, streamlit run.
set -euo pipefail

# Ensure CWD is the project root so .env / .token_cache.json / logs resolve correctly.
cd "$(dirname "$0")"

echo "========================================"
echo "  返餐费审批提交工具"
echo "========================================"
echo ""

# --- Check .env configuration ---
if [ ! -f ".env" ]; then
    echo "[错误] 未找到 .env 文件"
    echo "请复制 .env.example 为 .env 并填入飞书应用凭证："
    echo "  cp .env.example .env && nano .env"
    exit 1
fi

# --- Locate Python (prefer existing venv, else create one) ---
if [ -x "./.venv/bin/python" ]; then
    PY="./.venv/bin/python"
    echo "[venv] 使用已有虚拟环境: .venv"
elif command -v python3 >/dev/null 2>&1; then
    echo "[venv] 未找到虚拟环境，正在创建 .venv ..."
    python3 -m venv .venv
    PY="./.venv/bin/python"
    echo "[venv] 虚拟环境已创建"
else
    echo "[错误] 未找到 python3，请先安装："
    echo "  sudo apt update && sudo apt install -y python3 python3-venv"
    exit 1
fi

# --- Check dependencies (streamlit >= 1.33 required for st.button width) ---
echo "[2/3] 检查依赖..."
if ! "$PY" -c "
import streamlit as s
v = tuple(map(int, s.__version__.split('.')[:2]))
assert v >= (1, 33), f'streamlit {s.__version__} < 1.33'
import requests, dotenv, openpyxl, xlrd
" 2>/dev/null; then
    echo "      正在安装依赖（清华镜像）..."
    "$PY" -m pip install -r requirements.txt \
        --index-url https://pypi.tuna.tsinghua.edu.cn/simple
fi
echo "      依赖就绪"

# --- Prepare log directory ---
mkdir -p logs

# --- Launch Streamlit in server mode ---
echo "[3/3] 启动 Streamlit（服务器模式）..."
echo ""
echo "请访问: http://0.0.0.0:8501"
echo "按 Ctrl+C 可停止服务"
echo ""

exec "$PY" -m streamlit run app/app.py \
    --server.port 8501 \
    --server.address 0.0.0.0 \
    --server.headless true
