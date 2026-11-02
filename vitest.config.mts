import { defineConfig } from "vitest/config";

// Unit tests, set up like the PlaceOS Workplace app's (Vitest with globals and jsdom).
export default defineConfig({
  test: {
    globals: true,
    environment: "jsdom",
    include: ["src/tests/**/*.spec.ts"],
  },
});
