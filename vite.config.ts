import { defineConfig } from 'vite'
import vue from '@vitejs/plugin-vue'
import { fileURLToPath, URL } from 'node:url'
import { resolve, dirname, join } from 'node:path'
import { existsSync, readFileSync } from 'node:fs'
import { homedir } from 'node:os'

// Multi-entry Vite config: one entry per HTML surface in the add-in.
//
// - taskpane:    main UI shown in the Outlook side panel
// - commands:    function-file shell for ribbon button handlers
// - auth-relay:  hosted inside Office.context.ui.displayDialogAsync popup
//
// Office requires HTTPS for the dev source location; certs are provisioned
// by `npx office-addin-dev-certs install` (run once per developer machine).
const __dirname = dirname(fileURLToPath(import.meta.url))

/**
 * Best-effort load of the office-addin-dev-certs HTTPS material.
 *
 * Returns `undefined` when:
 *   - the package isn't installed yet (Sprint 1 pre-bootstrap), or
 *   - the developer hasn't run `npx office-addin-dev-certs install` yet.
 *
 * In those cases Vite falls back to plain HTTP, which is enough for
 * standalone-browser development but **not** enough for Outlook to sideload
 * the manifest. Run the install command and restart `make dev`.
 */
async function loadDevCerts(): Promise<{ key: Buffer; cert: Buffer } | undefined> {
  // Fast path: read the cert files directly from the location that
  // `npx office-addin-dev-certs install` writes to. This avoids the
  // package's CA-trust-store verify step, which can't elevate via sudo
  // on WSL and would otherwise force the dev server down to plain HTTP.
  const certDir = join(homedir(), '.office-addin-dev-certs')
  const keyPath = join(certDir, 'localhost.key')
  const certPath = join(certDir, 'localhost.crt')
  if (existsSync(keyPath) && existsSync(certPath)) {
    return { key: readFileSync(keyPath), cert: readFileSync(certPath) }
  }

  try {
    const devCerts = await import('office-addin-dev-certs')
    const opts = await devCerts.getHttpsServerOptions()
    return { key: opts.key as Buffer, cert: opts.cert as Buffer }
  } catch {
    return undefined
  }
}

export default defineConfig(async ({ mode, command }) => {
  // Only attempt the cert lookup when running `vite` (dev server). The build
  // command does not need HTTPS material.
  const https = command === 'serve' ? await loadDevCerts() : undefined

  if (command === 'serve' && !https) {
    console.warn(
      '\n[vite] office-addin-dev-certs not provisioned — serving plain HTTP on :3000.\n' +
        '       Run `npx office-addin-dev-certs install` then `make dev` again to get HTTPS,\n' +
        '       which Outlook requires to sideload the manifest.\n',
    )
  }

  return {
    plugins: [vue()],
    resolve: {
      alias: {
        '@': fileURLToPath(new URL('./src', import.meta.url)),
        '@shared': fileURLToPath(new URL('./src/shared', import.meta.url)),
      },
    },
    server: {
      port: 3000,
      strictPort: true,
      https,
      // Outlook on the Web loads the taskpane from a cross-origin iframe;
      // allow that during local dev.
      cors: true,
      headers: {
        'Access-Control-Allow-Origin': '*',
      },
    },
    build: {
      // Down-level to a WebKit-safe floor so no ES2022-only syntax reaches
      // Safari or the Outlook-on-Mac WKWebView (both lag Chrome/Edge). An
      // add-in that parses in Chrome but throws a SyntaxError in Safari never
      // mounts and reads as "the pane won't open". The output delta is tiny
      // and stays far under the 2 MiB `make budget` gate.
      target: ['safari15', 'chrome91', 'firefox90', 'edge91'],
      outDir: 'dist',
      emptyOutDir: true,
      sourcemap: mode !== 'production',
      rollupOptions: {
        input: {
          taskpane: resolve(__dirname, 'src/taskpane/taskpane.html'),
          commands: resolve(__dirname, 'src/commands/commands.html'),
          'auth-relay': resolve(__dirname, 'src/dialog/auth-relay.html'),
        },
      },
      chunkSizeWarningLimit: 700,
    },
  }
})
