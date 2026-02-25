import react from "@vitejs/plugin-react";
import fs from "fs";
import path from "path";
import { defineConfig } from "vite";

const tabName = path.basename(import.meta.dirname);

export default defineConfig({
  plugins: [react()],
  base: `/tabs/${tabName}`,
  root: path.resolve(import.meta.dirname),
  build: {
    outDir: path.resolve(import.meta.dirname, "../../dist", tabName),
    emptyOutDir: true,
  },
  esbuild: {
    tsconfigRaw: fs.readFileSync("./tsconfig.app.json"),
  },
});
