import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

export default defineConfig({
  // Relative asset URLs also work when Electron loads dist/index.html via file://.
  base: './',
  plugins: [react()],
});
