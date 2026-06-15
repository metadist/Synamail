# Deploying the Synamail add-in host (`addin.synaplan.com`)

The add-in is a static SPA. CI builds it into a tiny `caddy:2-alpine`-based
image and publishes it to GHCR; the deployment host auto-rolls it out.

## Pipeline

```text
push to main ──▶ GitHub Actions (.github/workflows/deploy.yml)
                   └─ build Dockerfile ─▶ ghcr.io/metadist/synamail:latest
host: synamail-addin-watchguard.timer (every ~2 min)
                   └─ docker pull :latest ─▶ digest changed? ─▶ docker compose up -d
host: synaplan-proxy (Caddy) ── addin.synaplan.com ─▶ synamail-addin:80
Cloudflare ── TLS (Flexible) ─▶ host:80
```

No SSH deploy step is needed — pushing to `main` is the whole workflow.

## One-time server setup

Run as `root` on the host that fronts `addin.synaplan.com` (SSH access details
are kept out of this public repo — see the internal ops notes). Paths below use
`$DEPLOY_DIR` (the conventional value on our host is `/webroot/synamail-addin`).

1. **Compose dir** — copy `deploy/docker-compose.yml` + `deploy/deploy.sh` to
   `$DEPLOY_DIR`.

2. **First image** — until CI has pushed `:latest`, load a locally-built image,
   or just `docker compose pull` once the GHCR package exists and is **public**
   (the box pulls GHCR anonymously). Then:

   ```bash
   cd "$DEPLOY_DIR" && docker compose up -d
   ```

3. **Watchguard** — install the auto-updater:

   ```bash
   install -m755 synamail-addin-watchguard.sh /usr/local/bin/
   install -m644 synamail-addin-watchguard.service /etc/systemd/system/
   install -m644 synamail-addin-watchguard.timer   /etc/systemd/system/
   systemctl daemon-reload
   systemctl enable --now synamail-addin-watchguard.timer
   ```

4. **Ingress** — in the shared reverse-proxy compose dir:
   - add the `synamail-addin_default` network to `docker-compose.yml` and the
     proxy service's `networks:` list;
   - add the `addin.synaplan.com` route to the proxy `Caddyfile`;
   - `docker compose up -d` to recreate the proxy on the new network.

## Important: make the GHCR package public

After the first CI push, set the `synamail` package visibility to **public**
(GitHub → Packages → synamail → Package settings). The host pulls GHCR
anonymously, exactly like `synaplan-org-website`.

## DNS / TLS

`addin.synaplan.com` is a Cloudflare-proxied record pointing at the host.
Cloudflare terminates TLS (Flexible) and forwards plain HTTP to the host on :80,
where Caddy routes by Host.
