import { useRef, useEffect, useCallback } from "react";
import {
  Send,
  Loader2,
  Paperclip,
  ImageIcon,
  X,
  FileText,
  FileSpreadsheet,
  File as FileIcon,
  Sparkles,
} from "lucide-react";
import type { ChatMessage, InputFile, InputImage } from "../../types/document";
import {
  fileToInputFile,
  fileToInputImage,
  clipboardItemToInputImage,
  formatFileSize,
} from "../../services/fileExtractor";

const DOC_ACCEPT = ".pdf,.docx,.txt,.md,.csv,.html,.htm";
const IMG_ACCEPT = "image/*";

interface Props {
  history: ChatMessage[];
  input: string;
  isGenerating: boolean;
  isRefining: boolean;
  hasDoc: boolean;
  files: InputFile[];
  images: InputImage[];
  onInputChange: (v: string) => void;
  onSend: () => void;
  onAddFiles: (files: InputFile[]) => void;
  onRemoveFile: (id: string) => void;
  onAddImage: (img: InputImage) => void;
  onRemoveImage: (id: string) => void;
}

function fileTypeIcon(type: string) {
  if (type.includes("pdf") || type.includes("word"))
    return <FileText size={12} className="text-blue-400" />;
  if (type.includes("csv") || type.includes("spreadsheet"))
    return <FileSpreadsheet size={12} className="text-green-400" />;
  return <FileIcon size={12} className="text-gray-400" />;
}

export default function ChatRefinement({
  history,
  input,
  isGenerating,
  isRefining,
  hasDoc,
  files,
  images,
  onInputChange,
  onSend,
  onAddFiles,
  onRemoveFile,
  onAddImage,
  onRemoveImage,
}: Props) {
  const bottomRef = useRef<HTMLDivElement>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const imageInputRef = useRef<HTMLInputElement>(null);
  const textareaRef = useRef<HTMLTextAreaElement>(null);
  const busy = isGenerating || isRefining;

  // Auto-scroll on new messages
  useEffect(() => {
    bottomRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [history, isGenerating, isRefining]);

  // Global Ctrl+V → paste images anywhere on page
  useEffect(() => {
    const handlePaste = async (e: ClipboardEvent) => {
      const items = Array.from(e.clipboardData?.items ?? []);
      for (const item of items) {
        if (item.type.startsWith("image/")) {
          const img = await clipboardItemToInputImage(item);
          if (img) onAddImage(img);
        }
      }
    };
    document.addEventListener("paste", handlePaste);
    return () => document.removeEventListener("paste", handlePaste);
  }, [onAddImage]);

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === "Enter" && !e.shiftKey && !busy && input.trim()) {
      e.preventDefault();
      onSend();
    }
  };

  // Auto-grow textarea
  const handleTextareaInput = (e: React.FormEvent<HTMLTextAreaElement>) => {
    const el = e.currentTarget;
    el.style.height = "auto";
    el.style.height = Math.min(el.scrollHeight, 160) + "px";
  };

  const handleDocFiles = useCallback(
    async (e: React.ChangeEvent<HTMLInputElement>) => {
      const picked = Array.from(e.target.files ?? []);
      const results: InputFile[] = [];
      for (const f of picked) {
        try {
          results.push(await fileToInputFile(f));
        } catch {}
      }
      if (results.length) onAddFiles(results);
      e.target.value = "";
    },
    [onAddFiles]
  );

  const handleImageFiles = useCallback(
    async (e: React.ChangeEvent<HTMLInputElement>) => {
      const picked = Array.from(e.target.files ?? []);
      for (const f of picked) {
        onAddImage(await fileToInputImage(f));
      }
      e.target.value = "";
    },
    [onAddImage]
  );

  // Drag files/images directly onto the composer
  const handleDrop = useCallback(
    async (e: React.DragEvent) => {
      e.preventDefault();
      const docs: InputFile[] = [];
      const imgs: InputImage[] = [];
      for (const f of Array.from(e.dataTransfer.files)) {
        if (f.type.startsWith("image/")) {
          imgs.push(await fileToInputImage(f));
        } else {
          try {
            docs.push(await fileToInputFile(f));
          } catch {}
        }
      }
      if (docs.length) onAddFiles(docs);
      imgs.forEach(onAddImage);
    },
    [onAddFiles, onAddImage]
  );

  return (
    <div className="flex h-full flex-col overflow-hidden">
      {/* ── Message history ───────────────────────────────────────── */}
      <div className="flex-1 overflow-y-auto px-4 py-4">
        {history.length === 0 && !isGenerating && (
          <div className="flex h-full flex-col items-center justify-center gap-4 py-10 text-center">
            <div className="flex h-12 w-12 items-center justify-center rounded-xl bg-blue-600/20 text-blue-400">
              <Sparkles size={24} />
            </div>
            <div>
              <p className="text-sm font-medium text-gray-300">
                What would you like to create?
              </p>
              <p className="mt-1 text-xs text-gray-600">
                Describe your document, attach files or paste screenshots — then hit Generate.
              </p>
            </div>
            <div className="flex gap-2 text-xs text-gray-600">
              {["📄 RFP Response", "🤝 Customer Proposal", "🎯 Demo Deck"].map((t) => (
                <div
                  key={t}
                  className="rounded-lg border border-gray-800 px-3 py-2"
                >
                  {t}
                </div>
              ))}
            </div>
          </div>
        )}

        <div className="space-y-3">
          {history.map((msg, i) => (
            <div
              key={i}
              className={`flex ${msg.role === "user" ? "justify-end" : "justify-start"}`}
            >
              <div
                className={`max-w-[80%] rounded-2xl px-4 py-2.5 text-sm leading-relaxed ${
                  msg.role === "user"
                    ? "rounded-br-sm bg-blue-600 text-white"
                    : "rounded-bl-sm bg-gray-800 text-gray-300"
                }`}
              >
                {msg.content}
              </div>
            </div>
          ))}

          {(isGenerating || isRefining) && (
            <div className="flex justify-start">
              <div className="flex items-center gap-2 rounded-2xl rounded-bl-sm bg-gray-800 px-4 py-2.5">
                <Loader2 size={13} className="animate-spin text-blue-400" />
                <span className="text-sm text-gray-400">
                  {isGenerating ? "Generating document…" : "Refining…"}
                </span>
              </div>
            </div>
          )}
          <div ref={bottomRef} />
        </div>
      </div>

      {/* ── Composer ─────────────────────────────────────────────── */}
      <div
        className="flex-shrink-0 border-t border-gray-800 px-4 pb-4 pt-3"
        onDragOver={(e) => e.preventDefault()}
        onDrop={handleDrop}
      >
        {/* Attached document chips */}
        {files.length > 0 && (
          <div className="mb-2 flex flex-wrap gap-1.5">
            {files.map((f) => (
              <span
                key={f.id}
                className="flex items-center gap-1.5 rounded-full border border-gray-700 bg-gray-800/80 px-2.5 py-1 text-xs text-gray-300"
              >
                {fileTypeIcon(f.type)}
                <span className="max-w-[140px] truncate">{f.name}</span>
                <span className="text-gray-600">{formatFileSize(f.size)}</span>
                <button
                  onClick={() => onRemoveFile(f.id)}
                  className="ml-0.5 text-gray-600 hover:text-red-400"
                >
                  <X size={11} />
                </button>
              </span>
            ))}
          </div>
        )}

        {/* Attached image thumbnails */}
        {images.length > 0 && (
          <div className="mb-2 flex flex-wrap gap-1.5">
            {images.map((img) => (
              <div
                key={img.id}
                className="group relative h-14 w-14 flex-shrink-0 overflow-hidden rounded-lg border border-gray-700"
              >
                <img
                  src={img.preview}
                  alt={img.name}
                  className="h-full w-full object-cover"
                />
                <button
                  onClick={() => onRemoveImage(img.id)}
                  className="absolute right-0.5 top-0.5 rounded-full bg-black/70 p-0.5 text-gray-300 opacity-0 transition-opacity group-hover:opacity-100 hover:text-red-400"
                >
                  <X size={10} />
                </button>
              </div>
            ))}
          </div>
        )}

        {/* Input row */}
        <div className="flex items-end gap-2 rounded-xl border border-gray-700 bg-gray-800/40 px-3 py-2 transition-colors focus-within:border-blue-600">
          {/* Attach icons */}
          <div className="flex gap-0.5 pb-0.5">
            <button
              type="button"
              title="Attach document (PDF, DOCX, TXT, CSV…)"
              onClick={() => fileInputRef.current?.click()}
              className="rounded-md p-1.5 text-gray-500 transition-colors hover:bg-gray-700 hover:text-gray-300"
            >
              <Paperclip size={16} />
            </button>
            <button
              type="button"
              title="Add image / screenshot"
              onClick={() => imageInputRef.current?.click()}
              className="rounded-md p-1.5 text-gray-500 transition-colors hover:bg-gray-700 hover:text-gray-300"
            >
              <ImageIcon size={16} />
            </button>
          </div>

          {/* Textarea */}
          <textarea
            ref={textareaRef}
            value={input}
            onChange={(e) => onInputChange(e.target.value)}
            onKeyDown={handleKeyDown}
            onInput={handleTextareaInput}
            placeholder={
              hasDoc
                ? "Ask for changes, e.g. make slide 3 more concise, add a ROI slide..."
                : "Describe what to generate, or paste notes / RFP content here..."
            }
            disabled={busy}
            rows={1}
            className="max-h-40 flex-1 resize-none overflow-y-auto bg-transparent text-sm text-gray-200 placeholder-gray-600 focus:outline-none disabled:opacity-50"
            style={{ lineHeight: "1.5" }}
          />

          {/* Send / Generate button */}
          <button
            type="button"
            onClick={onSend}
            disabled={busy || !input.trim()}
            title={hasDoc ? "Send" : "Generate document"}
            className="mb-0.5 flex-shrink-0 self-end rounded-lg bg-blue-600 p-2 text-white transition-colors hover:bg-blue-500 disabled:cursor-not-allowed disabled:opacity-40"
          >
            {busy ? (
              <Loader2 size={15} className="animate-spin" />
            ) : (
              <Send size={15} />
            )}
          </button>
        </div>

        <p className="mt-1.5 text-center text-xs text-gray-700">
          Enter to {hasDoc ? "send" : "generate"} · Shift+Enter for new line · Ctrl+V to paste screenshot
        </p>
      </div>

      {/* Hidden file inputs */}
      <input
        ref={fileInputRef}
        type="file"
        multiple
        accept={DOC_ACCEPT}
        className="hidden"
        onChange={handleDocFiles}
      />
      <input
        ref={imageInputRef}
        type="file"
        multiple
        accept={IMG_ACCEPT}
        className="hidden"
        onChange={handleImageFiles}
      />
    </div>
  );
}
