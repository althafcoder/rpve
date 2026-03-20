import { defineConfig } from "vite";
import react from "@vitejs/plugin-react-swc";
import path from "path";
import { componentTagger } from "lovable-tagger";

// https://vitejs.dev/config/
export default defineConfig(({ mode }) => ({
  server: {
    host: "::",
    port: 8080,
    hmr: {
      overlay: false,
    },
    proxy: {
      "/api": {
        target: "http://localhost:8009",
        changeOrigin: true,
      },
      "/docs": {
        target: "http://localhost:8007",
        changeOrigin: true,
      },
      "/openapi.json": {
        target: "http://localhost:8007",
        changeOrigin: true,
      },
      "/redoc": {
        target: "http://localhost:8007",
        changeOrigin: true,
      },
      "/work-comp-docs": {
        target: "http://localhost:8007",
        changeOrigin: true,
      },
      "/cognethro": {
        target: "http://localhost:8007",
        changeOrigin: true,
      },
      "/work-comp": {
        target: "http://localhost:8007",
        changeOrigin: true,
      },
    },
  },
  plugins: [react(), mode === "development" && componentTagger()].filter(Boolean),
  resolve: {
    alias: {
      "@": path.resolve(__dirname, "./src"),
    },
  },
}));
