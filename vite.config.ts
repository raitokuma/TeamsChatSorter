import { defineConfig } from "vite";

export default defineConfig({
  // GitHub Pages 專案網站請設 VITE_BASE=/<repo名稱>/
  base: process.env.VITE_BASE || "/",
  server: {
    port: 5173,
    // 允許透過 dev tunnel / ngrok 網域存取開發伺服器
    allowedHosts: true,
  },
  build: { outDir: "dist", target: "es2021" },
});
