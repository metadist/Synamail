/// <reference types="vite/client" />

// TypeScript 6 requires ambient declarations for side-effect CSS imports.
// Vite's `vite/client` reference covers most asset types, but we declare
// `*.css` explicitly so it works without `vite/client` chain regressions.
declare module '*.css'
