import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  plugins: [react()],
  base: "/ZenSpark/",
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
});
