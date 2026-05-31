#!/usr/bin/env bash
# One-shot starter for the local Synamail dev loop after a reboot.
#
# Brings up the three things Outlook needs that DON'T survive a restart:
#   1. Synaplan Docker stack   (backend :8000 + frontend :5173)
#   2. Synamail taskpane        (Vite, https://localhost:3000)
#   3. HTTPS sign-in bridge     (https://localhost:5174 → :5173)
#
# Cross-platform: works on macOS, native Linux, and WSL/Ubuntu — uses `lsof`
# (present on all three) for port + PID lookups instead of `ss` / `fuser`,
# which are Linux-only.
#
# Idempotent: re-running skips anything already up. dev + bridge run in the
# background; logs go to .dev-logs/. Stop them with:  ./start-dev.sh stop
set -euo pipefail

REPO="$(cd "$(dirname "$0")" && pwd)"
LOG_DIR="$REPO/.dev-logs"
mkdir -p "$LOG_DIR"

# --- Cross-platform helpers -------------------------------------------------

# Is something listening on the given TCP port?  Uses lsof (macOS, Linux, WSL,
# BSD all ship it). Falls back to `ss` for the rare Linux box that lacks lsof.
listening() {
  local port="$1"
  if command -v lsof >/dev/null 2>&1; then
    lsof -nP -iTCP:"$port" -sTCP:LISTEN >/dev/null 2>&1
  elif command -v ss >/dev/null 2>&1; then
    ss -ltn 2>/dev/null | grep -q ":$port\b"
  elif command -v netstat >/dev/null 2>&1; then
    netstat -an 2>/dev/null | grep -E "^tcp.*[\.:]$port[[:space:]].*LISTEN" >/dev/null
  else
    echo "  ! cannot determine listeners on :$port (need lsof, ss, or netstat)" >&2
    return 1
  fi
}

# Kill whatever is listening on the given TCP port. Quiet on success/no-op.
kill_port() {
  local port="$1" pids=""
  if command -v lsof >/dev/null 2>&1; then
    pids="$(lsof -ti tcp:"$port" 2>/dev/null || true)"
  fi
  if [ -z "$pids" ] && command -v fuser >/dev/null 2>&1; then
    # Linux fuser supports -k; macOS fuser does not, so this branch only
    # fires on Linux/WSL when lsof is missing.
    fuser -k "$port"/tcp 2>/dev/null && return 0 || true
    return 0
  fi
  if [ -n "$pids" ]; then
    # shellcheck disable=SC2086
    kill $pids 2>/dev/null || true
  fi
}

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

# --- Locate the Synaplan repo ----------------------------------------------
#
# Honours $SYNAPLAN_DIR if set. Otherwise tries, in order:
#   1. sibling of this repo (e.g. /Users/me/wwwroot/synaplan next to Synamail)
#   2. /wwwroot/synaplan        (the WSL convention)
#   3. $HOME/wwwroot/synaplan   (a common native-macOS / native-Linux convention)
#
# A directory counts only if it contains a docker-compose.yml.
resolve_synaplan_dir() {
  local candidate
  if [ -n "${SYNAPLAN_DIR:-}" ]; then
    if [ -f "$SYNAPLAN_DIR/docker-compose.yml" ]; then
      printf '%s' "$SYNAPLAN_DIR"
      return 0
    fi
    echo "error: SYNAPLAN_DIR=$SYNAPLAN_DIR has no docker-compose.yml" >&2
    return 1
  fi
  for candidate in \
    "$(cd "$REPO/.." 2>/dev/null && pwd)/synaplan" \
    "/wwwroot/synaplan" \
    "$HOME/wwwroot/synaplan" \
    ; do
    if [ -f "$candidate/docker-compose.yml" ]; then
      printf '%s' "$candidate"
      return 0
    fi
  done
  echo "error: cannot find the synaplan repo." >&2
  echo "       Set SYNAPLAN_DIR=/path/to/synaplan, or clone it as a sibling of $REPO." >&2
  return 1
}

# --- stop -------------------------------------------------------------------

if [ "${1:-start}" = "stop" ]; then
  echo "Stopping dev servers (Synaplan Docker is left running)…"
  pkill -f 'local-ssl-proxy.*5174' 2>/dev/null && echo "  bridge stopped" || echo "  bridge not running"
  pkill -f 'vite.*' 2>/dev/null && echo "  vite stopped" || true
  if listening 3000; then
    kill_port 3000
    echo "  taskpane (:3000) stopped"
  fi
  exit 0
fi

# --- start ------------------------------------------------------------------

SYNAPLAN_DIR="$(resolve_synaplan_dir)"
echo "Synaplan repo: $SYNAPLAN_DIR"

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
echo "  sign-in  : https://localhost:5174   (local Synaplan, default in dev)"
echo "  logs     : $LOG_DIR/{dev,bridge}.log"
echo "  stop     : ./start-dev.sh stop"
echo "  auth     : real sign-in. In Outlook → SignIn, set the server URL:"
echo "               local : https://localhost:5174  (this dev bridge)"
echo "               live  : https://web.synaplan.com (or your self-hosted host)"
