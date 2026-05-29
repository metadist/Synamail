#!/usr/bin/env bash
# One-shot starter for the local Synamail dev loop after a reboot.
#
# Brings up the three things Outlook needs that DON'T survive a restart:
#   1. Synaplan Docker stack   (backend :8000 + frontend :5173)
#   2. Synamail taskpane        (Vite, https://localhost:3000)
#   3. HTTPS sign-in bridge     (https://localhost:5174 → :5173)
#
# Idempotent: re-running skips anything already up. dev + bridge run in the
# background; logs go to .dev-logs/. Stop them with:  ./start-dev.sh stop
set -euo pipefail

REPO="$(cd "$(dirname "$0")" && pwd)"
SYNAPLAN_DIR="${SYNAPLAN_DIR:-/wwwroot/synaplan}"
LOG_DIR="$REPO/.dev-logs"
mkdir -p "$LOG_DIR"

listening() { ss -ltn 2>/dev/null | grep -q ":$1\b"; }

wait_for() { # port label timeout
  local port="$1" label="$2" timeout="${3:-30}" i=0
  while ! listening "$port"; do
    i=$((i + 1))
    [ "$i" -ge "$timeout" ] && {
      echo "  ✗ $label (:$port) did not come up in ${timeout}s — see $LOG_DIR" >&2
      return 1
    }
    sleep 1
  done
  echo "  ✓ $label  → :$port"
}

if [ "${1:-start}" = "stop" ]; then
  echo "Stopping dev servers (Synaplan Docker is left running)…"
  pkill -f 'local-ssl-proxy.*5174' 2>/dev/null && echo "  bridge stopped" || echo "  bridge not running"
  pkill -f 'vite.*' 2>/dev/null && echo "  vite stopped" || true
  # Vite is started via `make dev` → `npm run dev`; kill the node process on :3000.
  if listening 3000; then
    fuser -k 3000/tcp 2>/dev/null || true
    echo "  taskpane (:3000) stopped"
  fi
  exit 0
fi

echo "1/3  Synaplan Docker stack…"
(cd "$SYNAPLAN_DIR" && docker compose up -d >/dev/null)
wait_for 8000 "Synaplan backend" 60 || true
wait_for 5173 "Synaplan frontend" 60 || true

echo "2/3  Synamail taskpane (Vite :3000)…"
if listening 3000; then
  echo "  ✓ already running → :3000"
else
  nohup make -C "$REPO" dev >"$LOG_DIR/dev.log" 2>&1 &
  wait_for 3000 "Synamail taskpane" 40
fi

echo "3/3  HTTPS sign-in bridge (:5174)…"
if listening 5174; then
  echo "  ✓ already running → :5174"
else
  nohup make -C "$REPO" bridge >"$LOG_DIR/bridge.log" 2>&1 &
  wait_for 5174 "Sign-in bridge" 40
fi

echo ""
echo "Ready. In Outlook, open Synamail and (if needed) click Retry."
echo "  taskpane : https://localhost:3000"
echo "  sign-in  : https://localhost:5174   (enter this as the self-hosted instance)"
echo "  logs     : $LOG_DIR/{dev,bridge}.log"
echo "  stop     : ./start-dev.sh stop"
