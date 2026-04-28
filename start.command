#!/bin/zsh

set -e

APP_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$APP_DIR"

PYTHON="/Users/huhu/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3"
if [[ ! -x "$PYTHON" ]]; then
  PYTHON="$(command -v python3 || true)"
fi

if [[ -z "$PYTHON" ]]; then
  echo "未找到 Python 运行环境，请先安装 Python 3。"
  read -k 1 "?按任意键退出..."
  exit 1
fi

if [[ ! -f "$APP_DIR/app.py" ]]; then
  echo "未找到 app.py，请确认脚本位于报表生成项目目录。"
  read -k 1 "?按任意键退出..."
  exit 1
fi

if [[ ! -f "$APP_DIR/material/报表数据.xlsx" || ! -f "$APP_DIR/material/报表模板.docx" ]]; then
  echo "material 文件夹内缺少 报表数据.xlsx 或 报表模板.docx。"
  read -k 1 "?按任意键退出..."
  exit 1
fi

"$PYTHON" - <<'PY'
import importlib.util
import sys

missing = [name for name in ("openpyxl", "docx", "xlrd") if importlib.util.find_spec(name) is None]
if missing:
    print("缺少 Python 依赖：" + "、".join(missing))
    print("请使用 Codex 自带运行环境，或安装依赖后再启动：pip install openpyxl python-docx xlrd")
    sys.exit(1)
PY

PORT="${PORT:-8001}"
while lsof -iTCP:"$PORT" -sTCP:LISTEN -n -P >/dev/null 2>&1; do
  PORT=$((PORT + 1))
done

wait_until_ready() {
  "$PYTHON" - "$1" <<'PY'
import socket
import sys
import time

port = int(sys.argv[1])
deadline = time.time() + 10
while time.time() < deadline:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.settimeout(0.2)
        if sock.connect_ex(("localhost", port)) == 0:
            sys.exit(0)
    time.sleep(0.1)
sys.exit(1)
PY
}

local_ip() {
  "$PYTHON" - <<'PY'
import socket

ip = ""
try:
    with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as sock:
        sock.connect(("8.8.8.8", 80))
        ip = sock.getsockname()[0]
except OSError:
    ip = ""
print(ip)
PY
}

LOCAL_URL="http://localhost:$PORT"
LAN_IP="$(local_ip)"
LAN_URL=""
if [[ -n "$LAN_IP" ]]; then
  LAN_URL="http://$LAN_IP:$PORT"
fi

echo "本机访问：$LOCAL_URL"
if [[ -n "$LAN_URL" ]]; then
  echo "局域网访问：$LAN_URL"
fi
echo "按 Control + C 停止"

PORT="$PORT" HOST="0.0.0.0" "$PYTHON" app.py &
SERVER_PID=$!
trap 'kill "$SERVER_PID" >/dev/null 2>&1 || true; wait "$SERVER_PID" 2>/dev/null || true' INT TERM EXIT

if ! wait_until_ready "$PORT"; then
  echo "服务启动超时，请查看终端错误信息。"
  exit 1
fi

if [[ "${NO_OPEN:-}" != "1" ]]; then
  open "$LOCAL_URL" >/dev/null 2>&1 || true
fi

wait "$SERVER_PID"
