import { defineConfig, mergeConfig } from 'vitest/config'
import viteConfig from './vite.config'

export default defineConfig(async (_env) => {
  const base = await (typeof viteConfig === 'function'
    ? viteConfig({ command: 'serve', mode: 'test', isPreview: false, isSsrBuild: false })
    : Promise.resolve(viteConfig))

  return mergeConfig(base, {
    test: {
      environment: 'jsdom',
      globals: true,
      setupFiles: ['./tests/setup.ts'],
      include: ['tests/unit/**/*.test.ts', 'tests/component/**/*.test.ts'],
      coverage: {
        provider: 'v8',
        reporter: ['text', 'html'],
        include: ['src/**/*.{ts,vue}'],
        exclude: ['src/**/*.d.ts', 'src/**/index.ts'],
      },
    },
  })
})
