/// <reference types="vitest/config" />
import { defineConfig, type Plugin } from "vite";
import { resolve } from "node:path";
import { BASE_CSS } from "./src/viewer/styles";
import { VIEWER_CSS } from "./src/viewer/viewer-styles";
import { EDITOR_CSS } from "./src/editor/editor-styles";

/** Emits dist/pptxjs.css (the styles are also injected at runtime by the viewer/editor). */
function cssFile(): Plugin {
  return {
    name: "pptxjs-css",
    generateBundle() {
      const source = ["/* PPTXjs v2 */", BASE_CSS, VIEWER_CSS, EDITOR_CSS].join("\n");
      this.emitFile({ type: "asset", fileName: "pptxjs.css", source });
    },
  };
}

// Library build: ES modules + CommonJS for the core and every framework adapter.
// Framework packages are peer dependencies and stay external.
export default defineConfig({
  plugins: [cssFile()],
  build: {
    outDir: "dist",
    emptyOutDir: true,
    sourcemap: true,
    target: "es2020",
    lib: {
      entry: {
        pptxjs: resolve(__dirname, "src/index.ts"),
        react: resolve(__dirname, "src/adapters/react.ts"),
        vue: resolve(__dirname, "src/adapters/vue.ts"),
        jquery: resolve(__dirname, "src/adapters/jquery.ts"),
        angular: resolve(__dirname, "src/adapters/angular.ts"),
      },
      formats: ["es", "cjs"],
      fileName: (format, name) => `${name}.${format === "es" ? "js" : "cjs"}`,
    },
    rollupOptions: {
      external: ["react", "react/jsx-runtime", "vue", "jquery", "@angular/core", "@angular/common", "rxjs", /^@meshesha\/pptxjs/],
      output: { exports: "named" },
    },
  },
  test: {
    include: ["test/**/*.test.ts"],
    environment: "node",
    testTimeout: 60000,
    hookTimeout: 60000,
    pool: "forks",
    coverage: {
      provider: "v8",
      include: ["src/**/*.ts"],
      exclude: ["src/render/shape.ts", "src/dingbat.ts"],
      reporter: ["text-summary", "html"],
    },
  },
});
