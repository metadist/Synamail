# Deploying the Synamail add-in host (`addin.synaplan.com`)

The add-in is a static SPA. CI builds it into a tiny `caddy:2-alpine`-based
image and publishes it to GHCR; the **ch1** box auto-rolls it out.

## Pipeline

```text
push to main ──▶ GitHub Actions (.github/workflows/deploy.yml)
                   └─ build Dockerfile ─▶ ghcr.io/metadist/synamail:latest
ch1: synamail-addin-watchguard.timer (every ~2 min)
                   └─ docker pull :latest ─▶ digest changed? ─▶ docker compose up -d
ch1: synaplan-proxy (Caddy) ── addin.synaplan.com ─▶ synamail-addin:80
Cloudflare ── TLS (Flexible) ─▶ ch1:80
```

No SSH deploy step is needed — pushing to `main` is the whole workflow.

## One-time server setup (already done on ch1)

Run as `root` on ch1 (`ssh -p16803 root@ch1.synaplan.com`).

1. **Compose dir** — copy `deploy/docker-compose.yml` + `deploy/deploy.sh` to
   `/webroot/synamail-addin/`.

2. **First image** — until CI has pushed `:latest`, load a locally-built image,
   or just `docker compose pull` once the GHCR package exists and is **public**
   (the box pulls GHCR anonymously). Then:

   ```bash
   cd /webroot/synamail-addin && docker compose up -d
   ```

3. **Watchguard** — install the auto-updater:

   ```bash
   install -m755 synamail-addin-watchguard.sh /usr/local/bin/
   install -m644 synamail-addin-watchguard.service /etc/systemd/system/
   install -m644 synamail-addin-watchguard.timer   /etc/systemd/system/
   systemctl daemon-reload
   systemctl enable --now synamail-addin-watchguard.timer
   ```

4. **Ingress** — in `/webroot/synaplan-proxy/`:
   - add the `synamail-addin_default` network to `docker-compose.yml` and the
     proxy service's `networks:` list;
   - add the `addin.synaplan.com` route to `Caddyfile`;
   - `docker compose up -d` to recreate the proxy on the new network.

## Important: make the GHCR package public

After the first CI push, set the `synamail` package visibility to **public**
(GitHub → Packages → synamail → Package settings). The ch1 watchguard pulls
anonymously, exactly like `synaplan-org-website`.

## DNS / TLS

`addin.synaplan.com` is already a Cloudflare-proxied record pointing at ch1.
Cloudflare terminates TLS (Flexible) and forwards plain HTTP to the box on :80,
where Caddy routes by Host.
