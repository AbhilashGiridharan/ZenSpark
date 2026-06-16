import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig(({ command }) => ({
  plugins: [react()],
  base: command === "build" ? "/ZenSpark/" : "/",
  define: {
    "process.env": "{}",
    "process.version": '"v18.0.0"',
    "process.platform": '"browser"',
  },
  build: {
    outDir: "dist",
    rollupOptions: {
      output: {
        manualChunks: {
          pptx: ["pptxgenjs"],
          docx: ["docx"],
          pdf: ["pdfjs-dist"],
        },
      },
    },
  },
  optimizeDeps: {
    include: ["mammoth", "papaparse"],
  },
}));
