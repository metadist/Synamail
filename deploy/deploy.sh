#!/usr/bin/env bash
# =============================================================================
# Manually roll out the latest Synamail add-in image (addin.synaplan.com) on
# the host. CI publishes ghcr.io/metadist/synamail:latest on every push to main; the
# watchguard timer normally rolls it out within ~2 min. Use this to force it.
# =============================================================================
set -euo pipefail

cd "$(dirname "$0")"

echo "==> Pulling latest image..."
docker compose pull

echo "==> Recreating container..."
docker compose up -d --remove-orphans

echo "==> Waiting for health..."
for i in $(seq 1 20); do
  if curl -fsS -m 5 http://127.0.0.1:3001/healthz >/dev/null 2>&1; then
    echo "==> Healthy. Deployed:"
    docker compose ps
    exit 0
  fi
  sleep 3
done

echo "!! Health check did not pass within timeout. Recent logs:" >&2
docker compose logs --tail 40
exit 1
