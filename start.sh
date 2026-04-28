#!/usr/bin/env sh

set -eu

APP_DIR=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
cd "$APP_DIR"

PYTHON=${PYTHON:-}
if [ -z "$PYTHON" ]; then
  if [ -x "$APP_DIR/.venv/bin/python" ]; then
    PYTHON="$APP_DIR/.venv/bin/python"
  elif [ -x "/Users/huhu/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3" ]; then
    PYTHON="/Users/huhu/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3"
  elif command -v python3 >/dev/null 2>&1; then
    PYTHON=$(command -v python3)
  elif command -v python >/dev/null 2>&1; then
    PYTHON=$(command -v python)
  fi
fi

if [ -z "$PYTHON" ]; then
  echo "未找到 Python 运行环境，请先安装 Python 3。"
  exit 1
fi

if [ ! -f "$APP_DIR/app.py" ]; then
  echo "未找到 app.py，请确认脚本位于报表生成项目目录。"
  exit 1
fi

if [ ! -f "$APP_DIR/material/报表数据.xlsx" ] || [ ! -f "$APP_DIR/material/报表模板.docx" ]; then
  echo "material 文件夹内缺少 报表数据.xlsx 或 报表模板.docx。"
  exit 1
fi

"$PYTHON" - <<'PY'
import importlib.util
import sys

missing = [name for name in ("openpyxl", "docx", "xlrd") if importlib.util.find_spec(name) is None]
if missing:
    print("缺少 Python 依赖：" + "、".join(missing))
    print("请安装依赖：pip install openpyxl python-docx xlrd")
    sys.exit(1)
PY

port_is_busy() {
  "$PYTHON" - "$1" <<'PY'
import socket
import sys

port = int(sys.argv[1])
with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
    sock.settimeout(0.2)
    sys.exit(0 if sock.connect_ex(("localhost", port)) == 0 else 1)
PY
}

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

PORT=${PORT:-8001}
while port_is_busy "$PORT"; do
  PORT=$((PORT + 1))
done

LOCAL_URL="http://localhost:$PORT"
LAN_IP=$(local_ip)
LAN_URL=""
if [ -n "$LAN_IP" ]; then
  LAN_URL="http://$LAN_IP:$PORT"
fi

echo "本机访问：$LOCAL_URL"
if [ -n "$LAN_URL" ]; then
  echo "局域网访问：$LAN_URL"
fi
echo "按 Ctrl+C 停止"

PORT="$PORT" HOST="0.0.0.0" "$PYTHON" app.py &
SERVER_PID=$!
trap 'kill "$SERVER_PID" >/dev/null 2>&1 || true; wait "$SERVER_PID" 2>/dev/null || true' INT TERM EXIT

if ! wait_until_ready "$PORT"; then
  echo "服务启动超时，请查看终端错误信息。"
  exit 1
fi

if [ "${NO_OPEN:-}" != "1" ]; then
  if command -v open >/dev/null 2>&1; then
    open "$LOCAL_URL" >/dev/null 2>&1 || true
  elif command -v xdg-open >/dev/null 2>&1; then
    xdg-open "$LOCAL_URL" >/dev/null 2>&1 || true
  elif command -v cmd.exe >/dev/null 2>&1; then
    cmd.exe /c start "$LOCAL_URL" >/dev/null 2>&1 || true
  fi
fi

wait "$SERVER_PID"
