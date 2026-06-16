# syntax=docker/dockerfile:1
# =============================================================================
# Synamail add-in static host.
#
# The add-in is a static SPA (Vite multi-page build). This image builds it and
# serves dist/ + the manifest icons over plain HTTP on :80 behind the shared
# Caddy proxy on ch1 (Cloudflare terminates TLS). Published to GHCR as
# ghcr.io/metadist/synamail:latest and rolled out by the watchguard timer.
#
# The runtime stage is caddy:2-alpine — the SAME base image the ch1 proxy
# already runs, so the base layer is reused and the marginal footprint is just
# the ~250 KB static bundle.
# =============================================================================

FROM node:22-alpine AS build
WORKDIR /app
COPY package.json package-lock.json ./
RUN npm ci
COPY . .
# Type-checking is gated in CI; the image build only needs the bundle.
RUN npx vite build

FROM caddy:2-alpine AS runner
# Build provenance, passed by CI (defaults keep local builds working).
ARG GIT_SHA=unknown
ARG BUILD_TIME=unknown
COPY deploy/Caddyfile /etc/caddy/Caddyfile
# Vite output (serves /src/taskpane/taskpane.html, /src/commands/commands.html,
# /src/dialog/auth-relay.html, and the hashed /assets/* JS+CSS).
COPY --from=build /app/dist/ /srv/
# Manifest icons referenced by manifest.prod.xml (/assets/icon-*.png). The glob
# only matches the top-level PNGs, not assets/store/* or assets/source/*.
COPY --from=build /app/assets/*.png /srv/assets/
# Build stamp — curl https://addin.synaplan.com/version.json to confirm what's live.
RUN printf '{"name":"synamail","sha":"%s","built":"%s"}\n' "$GIT_SHA" "$BUILD_TIME" > /srv/version.json
EXPOSE 80
