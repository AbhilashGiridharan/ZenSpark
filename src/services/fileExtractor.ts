import * as pdfjsLib from "pdfjs-dist";
import type { InputFile, InputImage } from "../types/document";

// Use CDN worker for PDF.js to avoid bundling the large worker
pdfjsLib.GlobalWorkerOptions.workerSrc = `https://cdnjs.cloudflare.com/ajax/libs/pdf.js/${pdfjsLib.version}/pdf.worker.min.js`;

// ─── Read file as ArrayBuffer ─────────────────────────────────────────────────
function readAsArrayBuffer(file: File): Promise<ArrayBuffer> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result as ArrayBuffer);
    reader.onerror = () => reject(new Error(`Failed to read file: ${file.name}`));
    reader.readAsArrayBuffer(file);
  });
}

// ─── Read file as text ────────────────────────────────────────────────────────
function readAsText(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result as string);
    reader.onerror = () => reject(new Error(`Failed to read file: ${file.name}`));
    reader.readAsText(file, "utf-8");
  });
}

// ─── Read file as base64 ──────────────────────────────────────────────────────
function readAsBase64(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const result = reader.result as string;
      // Strip the data URL prefix
      const base64 = result.split(",")[1] ?? result;
      resolve(base64);
    };
    reader.onerror = () => reject(new Error(`Failed to read image: ${file.name}`));
    reader.readAsDataURL(file);
  });
}

// ─── PDF extraction via PDF.js ────────────────────────────────────────────────
async function extractPdf(file: File): Promise<string> {
  const arrayBuffer = await readAsArrayBuffer(file);
  const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
  const pages: string[] = [];

  for (let i = 1; i <= Math.min(pdf.numPages, 50); i++) {
    const page = await pdf.getPage(i);
    const content = await page.getTextContent();
    const pageText = content.items
      .map((item) => ("str" in item ? item.str : ""))
      .join(" ");
    pages.push(pageText);
  }

  return pages.join("\n\n");
}

// ─── DOCX extraction via mammoth ──────────────────────────────────────────────
async function extractDocx(file: File): Promise<string> {
  const mammoth = await import("mammoth");
  const arrayBuffer = await readAsArrayBuffer(file);
  const result = await mammoth.extractRawText({ arrayBuffer });
  return result.value;
}

// ─── CSV extraction via PapaParse ────────────────────────────────────────────
async function extractCsv(file: File): Promise<string> {
  const Papa = await import("papaparse");
  const text = await readAsText(file);
  const result = Papa.default.parse<Record<string, string>>(text, {
    header: true,
    skipEmptyLines: true,
  });

  if (result.data.length === 0) return text;

  // Format as readable table text
  const headers = result.meta.fields ?? [];
  const rows = result.data.slice(0, 200); // cap at 200 rows
  const lines = [
    headers.join(" | "),
    headers.map(() => "---").join(" | "),
    ...rows.map((row) => headers.map((h) => String(row[h] ?? "")).join(" | ")),
  ];
  return lines.join("\n");
}

// ─── Main extractor ───────────────────────────────────────────────────────────
export async function extractFileText(file: File): Promise<string> {
  const name = file.name.toLowerCase();
  const type = file.type.toLowerCase();

  if (name.endsWith(".pdf") || type === "application/pdf") {
    return extractPdf(file);
  }

  if (
    name.endsWith(".docx") ||
    type === "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
  ) {
    return extractDocx(file);
  }

  if (name.endsWith(".csv") || type === "text/csv") {
    return extractCsv(file);
  }

  if (
    name.endsWith(".txt") ||
    name.endsWith(".md") ||
    name.endsWith(".html") ||
    name.endsWith(".htm") ||
    type.startsWith("text/")
  ) {
    return readAsText(file);
  }

  // Fallback — try reading as text
  try {
    return await readAsText(file);
  } catch {
    return `[Could not extract text from ${file.name}]`;
  }
}

// ─── Convert a File to InputFile ─────────────────────────────────────────────
export async function fileToInputFile(file: File): Promise<InputFile> {
  const extractedText = await extractFileText(file);
  return {
    id: `${file.name}-${Date.now()}`,
    name: file.name,
    size: file.size,
    type: file.type || "unknown",
    extractedText,
  };
}

// ─── Convert a File/Blob to InputImage ───────────────────────────────────────
export async function fileToInputImage(
  file: File,
  captionHint = ""
): Promise<InputImage> {
  const base64 = await readAsBase64(file);
  const preview = `data:${file.type || "image/png"};base64,${base64}`;
  return {
    id: `img-${Date.now()}-${Math.random().toString(36).slice(2)}`,
    name: file.name || "pasted-image.png",
    base64,
    mimeType: file.type || "image/png",
    caption: captionHint,
    preview,
  };
}

// ─── Convert clipboard DataTransfer to InputImage ────────────────────────────
export async function clipboardItemToInputImage(
  item: DataTransferItem
): Promise<InputImage | null> {
  if (!item.type.startsWith("image/")) return null;
  const blob = item.getAsFile();
  if (!blob) return null;
  const file = new File([blob], `screenshot-${Date.now()}.png`, {
    type: blob.type || "image/png",
  });
  return fileToInputImage(file, "");
}

export function formatFileSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}
