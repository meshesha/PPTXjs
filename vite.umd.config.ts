import { defineConfig } from "vite";
import { resolve } from "node:path";

// Single-file browser bundle for <script> tags: window.PPTXjs
export default defineConfig({
  build: {
    outDir: "dist",
    emptyOutDir: false,
    sourcemap: true,
    target: "es2019",
    lib: {
      entry: resolve(__dirname, "src/umd.ts"),
      name: "PPTXjs",
      formats: ["umd"],
      fileName: () => "pptxjs.umd.js",
    },
  },
});
