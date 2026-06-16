#!/usr/bin/env bash
# start_pptx_server.sh — starts the python-pptx generation server
# Run from anywhere; no need to cd first.

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PORT="${PPTX_SERVER_PORT:-8765}"

# ── Kill any existing instance on the port ───────────────────────────────────
EXISTING=$(lsof -ti :"$PORT" 2>/dev/null)
if [[ -n "$EXISTING" ]]; then
  echo "Stopping existing server on port $PORT..."
  echo "$EXISTING" | xargs kill -9 2>/dev/null || true
  sleep 0.5
fi

# ── Create / reuse a local venv ──────────────────────────────────────────────
VENV_DIR="$SCRIPT_DIR/.venv-pptx-server"
if [[ ! -d "$VENV_DIR" ]]; then
  echo "Creating venv at $VENV_DIR..."
  python3 -m venv "$VENV_DIR"
fi

# ── Install / upgrade dependencies ───────────────────────────────────────────
echo "Installing dependencies..."
"$VENV_DIR/bin/pip" install -q --upgrade pip
"$VENV_DIR/bin/pip" install -q -r "$SCRIPT_DIR/requirements_pptx_server.txt"

# ── Launch ────────────────────────────────────────────────────────────────────
echo ""
echo "Starting PPTX Python Server on port $PORT..."
echo "  Health: http://localhost:$PORT/health"
echo "  Endpoint: POST http://localhost:$PORT/generate-pptx"
echo ""

PPTX_SERVER_PORT="$PORT" "$VENV_DIR/bin/python" "$SCRIPT_DIR/pptx_server.py"
