import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import path from "path";

// https://vite.dev/config/
export default defineConfig({
  plugins: [react()],
  resolve: {
    alias: {
      "@": path.resolve(__dirname, "./src"),
    },
  },
  server: {
    port: 5174,
    strictPort: true,
    proxy: { "/api": "http://127.0.0.1:3001", "/mcp": "http://127.0.0.1:3001" },
    open: false,
    watch: {
      ignored: [
        "**/.verification/**",
        "**/.office-data/**",
        "**/test-results/**",
        "**/playwright-report/**",
      ],
    },
  },
  build: {
    rollupOptions: {
      output: {
        manualChunks: {
          "spreadsheet-engine": ["xlsx"],
          "zip-engine": ["jszip"],
          "react-runtime": ["react", "react-dom"],
        },
      },
    },
  },
  preview: {
    proxy: { "/api": "http://127.0.0.1:3001", "/mcp": "http://127.0.0.1:3001" },
  },
});
