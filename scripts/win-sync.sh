#!/usr/bin/env bash
# Copy the manifest to C:\addin-catalog\manifest.xml with a bumped build number.
# Only needed when manifest.xml itself changed. Vue/TS code changes do NOT
# need this — Vite serves them live on every taskpane open.
set -euo pipefail

REPO="$(cd "$(dirname "$0")/.." && pwd)"
SRC="$REPO/manifest.xml"
OUT_DIR="/mnt/c/addin-catalog"
mkdir -p "$OUT_DIR" 2>/dev/null || OUT_DIR="$(wslpath -u "$(cmd.exe /C 'echo %USERPROFILE%' 2>/dev/null | tr -d '\r\n')")/addin-catalog" && mkdir -p "$OUT_DIR"

CUR="$(grep -oE '<Version>[0-9.]+</Version>' "$SRC" | head -1 | sed 's/<[^>]*>//g')"
COUNT_FILE="$REPO/.vite/build-no"
mkdir -p "$REPO/.vite"
N="$(cat "$COUNT_FILE" 2>/dev/null || echo "${CUR##*.}")"
N=$((N + 1))
echo "$N" > "$COUNT_FILE"
NEW="${CUR%.*}.$N"

sed "s|<Version>${CUR}</Version>|<Version>${NEW}</Version>|" "$SRC" > "$OUT_DIR/manifest.xml"
echo "$NEW → $(wslpath -w "$OUT_DIR")\\manifest.xml"
