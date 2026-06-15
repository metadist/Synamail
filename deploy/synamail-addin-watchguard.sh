#!/usr/bin/env bash
# Synamail add-in (addin.synaplan.com) watchguard: poll GHCR for the latest
# image, and if a new digest is available, let docker-compose recreate the
# container. Mirrors the synaplan-org-website watchguard.
#
# Installed at: /usr/local/bin/synamail-addin-watchguard.sh
# Scheduled via: /etc/systemd/system/synamail-addin-watchguard.{service,timer}
# Logs:          journalctl -u synamail-addin-watchguard.service

set -euo pipefail

COMPOSE_DIR="/webroot/synamail-addin"
COMPOSE="docker compose -f ${COMPOSE_DIR}/docker-compose.yml"
SERVICE_IMAGE="ghcr.io/metadist/synamail:latest"
CONTAINER="synamail-addin"

cd "$COMPOSE_DIR"

# Resolve the image currently running for the addin service.
running_digest="$(docker inspect "$CONTAINER" \
  --format "{{index .Image}}" 2>/dev/null || echo "none")"

# Pull silently; capture the digest of whatever is now :latest in GHCR.
docker pull --quiet "$SERVICE_IMAGE" >/dev/null
latest_digest="$(docker image inspect "$SERVICE_IMAGE" --format "{{.Id}}")"

if [ "$running_digest" = "$latest_digest" ]; then
  echo "watchguard: up to date ($latest_digest), no action"
  exit 0
fi

echo "watchguard: new image detected"
echo "  running: $running_digest"
echo "  latest : $latest_digest"
echo "watchguard: recreating stack via docker compose up -d"
$COMPOSE up -d --remove-orphans

# Clean up the previous image to keep disk usage bounded.
docker image prune -f >/dev/null || true
echo "watchguard: redeploy complete"
