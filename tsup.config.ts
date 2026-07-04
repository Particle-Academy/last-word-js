import { defineConfig } from "tsup";

export default defineConfig({
  entry: ["src/index.ts"],
  format: ["esm", "cjs"],
  dts: true,
  sourcemap: true,
  clean: true,
  // Isomorphic build: keep node builtins external so the dynamic `import("node:fs")`
  // in the file-I/O helpers stays a runtime import (and never ships to browsers).
  platform: "neutral",
  external: ["node:fs", "node:path", "node:os"],
  treeshake: true,
});
