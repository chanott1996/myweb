import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// base: "./" ทำให้ asset ใช้ path แบบ relative → ใช้ได้กับ GitHub Pages subpath (username.github.io/repo/)
export default defineConfig({
  base: "./",
  plugins: [react()],
});
