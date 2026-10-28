import { build } from "esbuild";

// The Console injects React/antd through window.QwenPaw.host, so the bundle has
// no bare imports at all. JSX is compiled to React.createElement calls against
// the locally bound host React instance (classic runtime).
await build({
  entryPoints: ["src/index.tsx"],
  bundle: true,
  format: "esm",
  target: ["es2020"],
  jsx: "transform",
  outfile: "../dist/index.js",
  legalComments: "none",
  minify: true,
  sourcemap: false,
  logLevel: "info",
});
