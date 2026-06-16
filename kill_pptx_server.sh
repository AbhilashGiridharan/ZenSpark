#!/usr/bin/env bash
# kill_pptx_server.sh — stop the python-pptx generation server

PORT="${PPTX_SERVER_PORT:-8765}"

# Kill by port
PIDS=$(lsof -ti :"$PORT" 2>/dev/null)
if [[ -n "$PIDS" ]]; then
  echo "Killing processes on port $PORT: $PIDS"
  echo "$PIDS" | xargs kill -9
else
  echo "No process found on port $PORT"
fi

# Also kill any lingering pptx_server processes
pkill -9 -f "pptx_server.py" 2>/dev/null && echo "Killed pptx_server.py processes" || true

echo "Done."
