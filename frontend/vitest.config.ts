import { defineConfig } from "vitest/config";

export default defineConfig({
  esbuild: {
    // The plugin uses the host-supplied React instance; classic JSX keeps the
    // emitted code referring to the local `React` binding instead of importing
    // react/jsx-runtime.
    jsx: "transform",
  },
  test: {
    environment: "jsdom",
    setupFiles: ["test/setup.ts"],
    include: ["test/**/*.test.{ts,tsx}"],
    restoreMocks: true,
  },
});
