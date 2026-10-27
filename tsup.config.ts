import { defineConfig } from "tsup";

export default defineConfig({
  // src/cli.ts is the Node-only `last-word` bin (package.json "bin" -> dist/cli.js).
  // A separate entry, so the index stays isomorphic.
  entry: ["src/index.ts", "src/cli.ts"],
  format: ["esm", "cjs"],
  dts: true,
  sourcemap: true,
  clean: true,
  // Isomorphic build: keep node builtins external so the dynamic `import("node:fs")`
  // in the file-I/O helpers stays a runtime import (and never ships to browsers).
  platform: "neutral",
  external: ["node:fs", "node:path", "node:os", "node:crypto"],
  treeshake: true,
});
