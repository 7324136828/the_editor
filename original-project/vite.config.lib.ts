import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import path from "path";

export default defineConfig({
  plugins: [react()],
  build: {
    lib: {
      entry: path.resolve(__dirname, "src/index.ts"),
      name: "VSCodeOfficeStudio",
      cssFileName: "style",
      fileName: (format) =>
        `vscode-office-studio.${format === "es" ? "js" : "umd.cjs"}`,
      formats: ["es", "umd"],
    },
    rollupOptions: {
      external: ["react", "react-dom"],
      output: {
        exports: "named",
        globals: {
          react: "React",
          "react-dom": "ReactDOM",
        },
      },
    },
    outDir: "dist",
    emptyOutDir: false,
  },
});
