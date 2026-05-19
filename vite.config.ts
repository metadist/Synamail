import { defineConfig } from 'vite'
import vue from '@vitejs/plugin-vue'
import { fileURLToPath, URL } from 'node:url'
import { resolve, dirname } from 'node:path'

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
      target: 'es2022',
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
