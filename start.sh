#!/usr/bin/env bash
# CloudStudio / Lighthouse / 任意 Linux 云主机的启动脚本
set -e
cd "$(dirname "$0")"
PORT="${PORT:-8765}"
export PORT
echo "[start.sh] installing deps..."
pip install -r requirements.txt
echo "[start.sh] starting server on port $PORT ..."
# gunicorn 单进程多线程：Flask + 长任务（Excel 解析）友好；timeout 给到 120s
exec gunicorn -w 1 --threads 4 -b 0.0.0.0:${PORT} --timeout 120 server:app
