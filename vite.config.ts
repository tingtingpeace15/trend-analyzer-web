import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

// GitHub Pages 托管在 https://<user>.github.io/trend-analyzer-web/,
// 构建产物要用仓库名做 base;本地 dev 仍是 '/'
export default defineConfig(({ command }) => ({
  base: command === 'build' ? '/trend-analyzer-web/' : '/',
  plugins: [react()],
  worker: {
    format: 'es',
  },
}));
