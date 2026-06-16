import * as pdfjsLib from "pdfjs-dist";
import PdfWorkerUrl from "pdfjs-dist/build/pdf.worker.min.mjs?url";
import type { InputFile, InputImage } from "../types/document";

// Use the locally-bundled worker — avoids CDN version mismatch (pdf.js v4 uses .mjs)
pdfjsLib.GlobalWorkerOptions.workerSrc = PdfWorkerUrl;

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

// ─── PPTX extraction via JSZip ───────────────────────────────────────────────
// A .pptx file is a ZIP containing XML slide files at ppt/slides/slide*.xml
async function extractPptx(file: File): Promise<string> {
  const JSZip = (await import("jszip")).default;
  const arrayBuffer = await readAsArrayBuffer(file);
  const zip = await JSZip.loadAsync(arrayBuffer);

  // Collect slide XML files in order
  const slideFiles = Object.keys(zip.files)
    .filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
    .sort((a, b) => {
      const numA = parseInt(a.match(/\d+/)?.[0] ?? "0");
      const numB = parseInt(b.match(/\d+/)?.[0] ?? "0");
      return numA - numB;
    });

  const pages: string[] = [];

  for (const slideName of slideFiles) {
    const xml = await zip.files[slideName].async("text");
    // Strip all XML tags, decode entities, collapse whitespace
    const text = xml
      .replace(/<[^>]+>/g, " ")
      .replace(/&amp;/g, "&")
      .replace(/&lt;/g, "<")
      .replace(/&gt;/g, ">")
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'")
      .replace(/\s+/g, " ")
      .trim();
    if (text) pages.push(text);
  }

  return pages
    .map((t, i) => `[Slide ${i + 1}]\n${t}`)
    .join("\n\n");
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
    name.endsWith(".pptx") ||
    type === "application/vnd.openxmlformats-officedocument.presentationml.presentation"
  ) {
    return extractPptx(file);
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

// ─── Folder knowledge base (File System Access API) ──────────────────────────
const KB_SUPPORTED_EXTS = new Set([
  ".pdf", ".pptx", ".docx", ".txt", ".md", ".csv", ".json", ".html", ".htm", ".xml",
]);

async function walkDirectory(
  dirHandle: FileSystemDirectoryHandle,
  path = "",
  files: File[] = []
): Promise<File[]> {
  for await (const [name, handle] of dirHandle as unknown as AsyncIterable<[string, FileSystemHandle]>) {
    if (handle.kind === "file") {
      const ext = name.slice(name.lastIndexOf(".")).toLowerCase();
      if (KB_SUPPORTED_EXTS.has(ext)) {
        const file = await (handle as FileSystemFileHandle).getFile();
        // Attach relative path as a property for display
        Object.defineProperty(file, "_kbPath", { value: path ? `${path}/${name}` : name });
        files.push(file);
      }
    } else if (handle.kind === "directory") {
      await walkDirectory(handle as FileSystemDirectoryHandle, path ? `${path}/${name}` : name, files);
    }
  }
  return files;
}

export interface KnowledgeBaseResult {
  files: InputFile[];
  folderName: string;
  totalFiles: number;
  skippedFiles: number;
}

export async function loadKnowledgeBaseFolder(): Promise<KnowledgeBaseResult | null> {
  if (!("showDirectoryPicker" in window)) {
    alert("Folder picker is only supported in Chrome or Edge (desktop). Please use one of those browsers.");
    return null;
  }
  let dirHandle: FileSystemDirectoryHandle;
  try {
    dirHandle = await (window as unknown as { showDirectoryPicker: (o?: Record<string, unknown>) => Promise<FileSystemDirectoryHandle> })
      .showDirectoryPicker({ mode: "read" });
  } catch {
    // User cancelled
    return null;
  }

  const raw = await walkDirectory(dirHandle);
  const inputFiles: InputFile[] = [];
  let skipped = 0;

  for (const file of raw) {
    try {
      const extracted = await extractFileText(file);
      if (extracted.trim()) {
        inputFiles.push({
          id: `kb-${file.name}-${Date.now()}-${Math.random().toString(36).slice(2)}`,
          name: (file as File & { _kbPath?: string })._kbPath ?? file.name,
          size: file.size,
          type: file.type || "text/plain",
          extractedText: extracted,
        });
      } else {
        skipped++;
      }
    } catch {
      skipped++;
    }
  }

  return {
    files: inputFiles,
    folderName: dirHandle.name,
    totalFiles: raw.length,
    skippedFiles: skipped,
  };
}
