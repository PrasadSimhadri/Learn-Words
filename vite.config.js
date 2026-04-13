import { defineConfig } from 'vite';

// https://vitejs.dev/config/
export default defineConfig({
  // Use relative paths for assets so the app works on GitHub Pages / Subfolders
  base: './',
  build: {
    outDir: 'dist',
  },
});
