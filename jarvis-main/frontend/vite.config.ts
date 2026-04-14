import { defineConfig } from "vite"

export default defineConfig({
  server: {
    port: 5173,
    host: "0.0.0.0", // Allows access via your PC's IP address
    proxy: {
      "/ws": {
        // Change to http if you aren't using the .pem certificates
        target: "http://127.0.0.1:8340", 
        ws: true,
        secure: false,
        changeOrigin: true,
      },
      "/api": {
        target: "http://127.0.0.1:8340",
        secure: false,
        changeOrigin: true,
      },
    },
  },
  build: {
    outDir: "dist",
  },
});