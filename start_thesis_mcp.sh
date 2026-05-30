#!/bin/bash
# Start the Thesis MCP Server for thesis format review.
# Supports both virtual environment and system Python.
#
# Usage:
#   ./start_thesis_mcp.sh              # use default config
#   ./start_thesis_mcp.sh .env         # load config from .env
#
# Setup (first time only):
#   pip install fastmcp python-docx requests boto3

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

# ---------- Load environment ----------
# 1. Try user-specified .env file
if [ -n "$1" ] && [ -f "$1" ]; then
    set -a
    source "$1"
    set +a
    echo "[INFO] Loaded config from: $1"
# 2. Try local .env
elif [ -f "$SCRIPT_DIR/.env" ]; then
    set -a
    source "$SCRIPT_DIR/.env"
    set +a
    echo "[INFO] Loaded local .env"
fi

# ---------- Python environment ----------
# Try virtual environment first (legacy path), then current directory venv
if [ -f "/media/zrway/15e45490-a531-43bb-a299-9498299c5971/LHL/myenv/bin/activate" ]; then
    source "/media/zrway/15e45490-a531-43bb-a299-9498299c5971/LHL/myenv/bin/activate"
    echo "[INFO] Activated virtual environment: myenv"
elif [ -f "$SCRIPT_DIR/venv/bin/activate" ]; then
    source "$SCRIPT_DIR/venv/bin/activate"
    echo "[INFO] Activated virtual environment: venv"
elif [ -f "$SCRIPT_DIR/.venv/bin/activate" ]; then
    source "$SCRIPT_DIR/.venv/bin/activate"
    echo "[INFO] Activated virtual environment: .venv"
fi

# Verify Python and dependencies
python3 -c "import fastmcp, docx, requests" 2>/dev/null || {
    echo "[ERROR] Required packages not found."
    echo "  Run: pip install fastmcp python-docx requests boto3"
    exit 1
}

# ---------- Start / restart ----------
PORT="${THESIS_MCP_PORT:-8899}"

if pgrep -f "thesis_mcp_server.py" > /dev/null 2>&1; then
    echo "[INFO] Stopping existing thesis_mcp_server instance ..."
    pkill -f "thesis_mcp_server.py" 2>/dev/null || true
    sleep 1
fi

nohup python3 "$SCRIPT_DIR/thesis_mcp_server.py" > "$SCRIPT_DIR/thesis_mcp.log" 2>&1 &
PID=$!

sleep 2
if kill -0 "$PID" 2>/dev/null; then
    echo "[OK] Thesis MCP Server started on port $PORT (PID: $PID)"
    echo "[INFO] MCP URL: http://<host-ip>:$PORT/mcp"
    echo "[INFO] Log file: $SCRIPT_DIR/thesis_mcp.log"
else
    echo "[ERROR] Server failed to start. Check $SCRIPT_DIR/thesis_mcp.log"
    exit 1
fi
