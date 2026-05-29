#!/usr/bin/env bash
# HTTPS bridge for the local Synaplan dev loop.
#
# The Synamail sign-in dialog goes through Office.context.ui.displayDialogAsync,
# which Office hard-rejects unless the URL is HTTPS. The taskpane is also an
# HTTPS origin (https://localhost:3000), so its fetch() calls to Synaplan are
# blocked as mixed content unless Synaplan is reachable over HTTPS too.
#
# Your local Synaplan Docker exposes the frontend dev server on plain
# http://localhost:5173 (which itself proxies /api → backend on :8000). This
# script puts a one-process HTTPS terminator in front of it on :5174, REUSING
# the office-addin-dev-certs certificate that already fronts Synamail on :3000.
#
# Reusing that cert (instead of a throwaway self-signed one) means a SINGLE
# trusted root in Windows covers BOTH origins, so desktop Outlook's WebView2 —
# not just Outlook on the Web — accepts the dialog without cert errors.
#
# Usage:
#   ./scripts/dev-bridge-proxy.sh            # 5174 → 5173, dev-certs cert
#   SOURCE=6174 TARGET=5173 ./scripts/dev-bridge-proxy.sh
set -euo pipefail

SOURCE="${SOURCE:-5174}"
TARGET="${TARGET:-5173}"
CERT_DIR="${CERT_DIR:-$HOME/.office-addin-dev-certs}"
CERT="$CERT_DIR/localhost.crt"
KEY="$CERT_DIR/localhost.key"

if [[ ! -f "$CERT" || ! -f "$KEY" ]]; then
  echo "error: dev cert not found at $CERT_DIR" >&2
  echo "       run 'npx office-addin-dev-certs install' first (one-time)." >&2
  exit 1
fi

# Already running? Don't crash with EADDRINUSE — one bridge is enough.
if command -v ss >/dev/null 2>&1 && ss -ltn 2>/dev/null | grep -q ":$SOURCE\b"; then
  echo "Bridge already listening on https://localhost:$SOURCE — nothing to do."
  echo "  (Stop it with:  pkill -f 'local-ssl-proxy.*$SOURCE' )"
  exit 0
fi

# Sanity check: is the Synaplan frontend actually up on the target port?
if ! curl -sf -o /dev/null "http://localhost:$TARGET/"; then
  echo "warning: nothing answering on http://localhost:$TARGET/" >&2
  echo "         start Synaplan first:  (cd /wwwroot/synaplan && docker compose up -d)" >&2
fi

echo "Synaplan HTTPS bridge:  https://localhost:$SOURCE  →  http://localhost:$TARGET"
echo "  /addin/connect dialog + /api both reachable over HTTPS for Synamail."
echo "  Cert: $CERT (same trust root as Synamail :3000)."
echo ""

exec npx --yes local-ssl-proxy \
  --source "$SOURCE" \
  --target "$TARGET" \
  --cert "$CERT" \
  --key "$KEY"
