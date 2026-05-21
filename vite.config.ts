/// <reference types="vitest" />
import { resolve } from "node:path";
import basicSsl from "@vitejs/plugin-basic-ssl";
import react from "@vitejs/plugin-react";
import { defineConfig } from "vite";

const publicBase = process.env.PUBLIC_BASE_PATH ?? "/";

export default defineConfig({
  base: publicBase,
  plugins: [react(), basicSsl()],
  server: {
    port: 3000,
    strictPort: true,
  },
  build: {
    outDir: "dist",
    emptyOutDir: true,
    rollupOptions: {
      input: {
        taskpane: resolve(__dirname, "src/taskpane/index.html"),
        commands: resolve(__dirname, "src/commands/commands.html"),
      },
    },
  },
  test: {
    globals: true,
    environment: "node",
  },
});
