import { defineConfig, type Plugin } from "vite";
import { resolve } from "node:path";
import fs from "node:fs";

const SAMPLE = resolve(__dirname, "../test/fixtures/Sample_12.pptx");

/** Serves (dev) and emits (build) the sample deck without duplicating it in the repo. */
function sampleDeck(): Plugin {
  return {
    name: "pptxjs-sample-deck",
    configureServer(server) {
      server.middlewares.use("/samples/Sample_12.pptx", (_req, res) => {
        res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation");
        fs.createReadStream(SAMPLE).pipe(res);
      });
    },
    generateBundle() {
      if (fs.existsSync(SAMPLE)) this.emitFile({ type: "asset", fileName: "samples/Sample_12.pptx", source: fs.readFileSync(SAMPLE) });
    },
  };
}

export default defineConfig({
  root: __dirname,
  base: "./",
  plugins: [sampleDeck()],
  resolve: {
    alias: {
      "@meshesha/pptxjs/jquery": resolve(__dirname, "../src/adapters/jquery.ts"),
      "@meshesha/pptxjs": resolve(__dirname, "../src/index.ts"),
    },
  },
  build: { outDir: resolve(__dirname, "../dist-demo"), emptyOutDir: true },
  server: { port: 5175, strictPort: true },
});
