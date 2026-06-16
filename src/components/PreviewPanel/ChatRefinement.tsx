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
  FolderOpen,
  HelpCircle,
  CheckCircle2,
  Mic,
  Plus,
  MessageSquare,
  Wand2,
} from "lucide-react";
import type { ChatMessage, ChatAttachment, InputFile, InputImage, Slide } from "../../types/document";
import type { ClarifyingQuestion } from "../../services/azureFoundry";
import {
  fileToInputFile,
  fileToInputImage,
  clipboardItemToInputImage,
  formatFileSize,
} from "../../services/fileExtractor";

const DOC_ACCEPT = ".pdf,.pptx,.docx,.txt,.md,.csv,.html,.htm";
const IMG_ACCEPT = "image/*";

interface Props {
  history: ChatMessage[];
  input: string;
  isGenerating: boolean;
  isRefining: boolean;
  streamingContent?: string;
  hasDoc: boolean;
  files: InputFile[];
  images: InputImage[];
  slides?: Slide[];
  clarifyingQuestions?: ClarifyingQuestion[];
  clarifyAnswers?: Record<string, string>;
  isClarifying?: boolean;
  onInputChange: (v: string) => void;
  onSend: () => void;
  onAddFiles: (files: InputFile[]) => void;
  onRemoveFile: (id: string) => void;
  onAddImage: (img: InputImage) => void;
  onRemoveImage: (id: string) => void;
  onLoadFolder?: () => void;
  onClarifyAnswer?: (id: string, answer: string) => void;
  onSkipClarify?: () => void;
  onSubmitClarify?: () => void;
}

function fileTypeIcon(type: string, name = "") {
  if (name.endsWith(".pptx") || type.includes("presentation"))
    return <FileText size={12} className="text-orange-400" />;
  if (type.includes("pdf") || type.includes("word"))
    return <FileText size={12} className="text-blue-400" />;
  if (type.includes("csv") || type.includes("spreadsheet"))
    return <FileSpreadsheet size={12} className="text-green-400" />;
  return <FileIcon size={12} className="text-gray-500" />;
}

export default function ChatRefinement({
  history,
  input,
  isGenerating,
  isRefining,
  streamingContent = "",
  hasDoc,
  files,
  images,
  slides = [],
  clarifyingQuestions = [],
  clarifyAnswers = {},
  isClarifying = false,
  onInputChange,
  onSend,
  onAddFiles,
  onRemoveFile,
  onAddImage,
  onRemoveImage,
  onLoadFolder,
  onClarifyAnswer,
  onSkipClarify,
  onSubmitClarify,
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
    <div className="flex h-full flex-col overflow-hidden bg-white">

      {/* ── Message history / empty state ─────────────────────────── */}
      <div className="flex-1 overflow-y-auto px-4 py-4">

        {/* Empty state — suggestion cards */}
        {history.length === 0 && !isGenerating && clarifyingQuestions.length === 0 && !isClarifying && (
          <div className="flex h-full flex-col items-start justify-center px-2 pb-24 gap-5">
            <p className="text-[26px] font-semibold text-gray-800 leading-snug px-2">
              What would you like to generate?
            </p>
            <div className="grid grid-cols-2 gap-3 w-full">
              {[
                {
                  title: "PowerPoint presentation",
                  desc: "Slides with layouts, visuals & speaker notes",
                  prompt: "Create a PowerPoint presentation about ",
                  icon: "📊",
                },
                {
                  title: "Word document",
                  desc: "Structured report, proposal or article",
                  prompt: "Write a Word document about ",
                  icon: "📄",
                },
                {
                  title: "Pitch deck",
                  desc: "Investor or sales deck with key slides",
                  prompt: "Create a pitch deck for ",
                  icon: "🚀",
                },
                {
                  title: "Project report",
                  desc: "Status update, findings & next steps",
                  prompt: "Generate a project report for ",
                  icon: "📋",
                },
              ].map(({ title, desc, prompt, icon }) => (
                <button
                  key={title}
                  onClick={() => { onInputChange(prompt); textareaRef.current?.focus(); }}
                  className="flex flex-col items-start gap-1 rounded-2xl border border-gray-200 bg-white px-4 py-3.5 text-left shadow-sm hover:border-gray-300 hover:shadow-md transition-all"
                >
                  <span className="text-xl">{icon}</span>
                  <span className="text-sm font-medium text-gray-800">{title}</span>
                  <span className="text-xs text-gray-400 leading-relaxed">{desc}</span>
                </button>
              ))}
            </div>
          </div>
        )}

        <div className="space-y-4">
          {history.map((msg, i) => (
            <div
              key={i}
              className={`flex ${msg.role === "user" ? "justify-end" : "justify-start"}`}
            >
              <div
                className={`max-w-[82%] rounded-2xl px-4 py-2.5 text-sm leading-relaxed ${
                  msg.role === "user"
                    ? "rounded-br-sm bg-gray-100 text-gray-900"
                    : "text-gray-800"
                }`}
              >
                {msg.attachments && msg.attachments.length > 0 && (
                  <div className="mb-2 flex flex-wrap gap-1.5">
                    {msg.attachments.map((att, ai) => (
                      <AttachmentBadge key={ai} att={att} />
                    ))}
                  </div>
                )}
                {msg.content}
              </div>
            </div>
          ))}

          {/* Clarifying questions */}
          {clarifyingQuestions.length > 0 && !isGenerating && (
            <div className="rounded-2xl border border-gray-200 bg-gray-50 p-4">
              <div className="mb-3 flex items-center gap-2">
                <HelpCircle size={14} className="text-gray-500" />
                <span className="text-sm font-medium text-gray-700">A few quick questions</span>
              </div>
              <div className="space-y-3">
                {clarifyingQuestions.map((q) => (
                  <div key={q.id} className="space-y-1.5">
                    <p className="text-sm text-gray-700">{q.question}</p>
                    {q.hint && <p className="text-xs text-gray-400">{q.hint}</p>}
                    <input
                      type="text"
                      value={clarifyAnswers[q.id] ?? ""}
                      onChange={(e) => onClarifyAnswer?.(q.id, e.target.value)}
                      placeholder="Your answer…"
                      className="w-full rounded-xl border border-gray-200 bg-white px-3 py-2 text-sm text-gray-800 placeholder-gray-400 focus:border-gray-400 focus:outline-none"
                    />
                  </div>
                ))}
              </div>
              <div className="mt-4 flex gap-2">
                <button
                  onClick={onSubmitClarify}
                  className="flex items-center gap-1.5 rounded-full bg-gray-900 px-4 py-2 text-xs font-medium text-white hover:bg-gray-700"
                >
                  <CheckCircle2 size={12} />
                  Generate
                </button>
                <button
                  onClick={onSkipClarify}
                  className="rounded-full border border-gray-200 px-4 py-2 text-xs text-gray-500 hover:border-gray-300 hover:text-gray-700"
                >
                  Skip
                </button>
              </div>
            </div>
          )}

          {isClarifying && (
            <div className="flex justify-start">
              <div className="flex items-center gap-2 text-sm text-gray-400">
                <Loader2 size={13} className="animate-spin" />
                Preparing questions…
              </div>
            </div>
          )}

          {(isGenerating || isRefining) && (
            <div className="flex justify-start">
              {streamingContent ? (
                <div className="max-w-[82%] rounded-2xl px-4 py-2.5 text-sm leading-relaxed text-gray-800 whitespace-pre-wrap">
                  {streamingContent}
                  <span className="inline-block w-1.5 h-3.5 ml-0.5 bg-gray-400 animate-pulse rounded-sm align-text-bottom" />
                </div>
              ) : (
                <div className="flex items-center gap-2 text-sm text-gray-400">
                  <Loader2 size={13} className="animate-spin" />
                  {isGenerating ? "Generating…" : "Refining…"}
                </div>
              )}
            </div>
          )}
          <div ref={bottomRef} />
        </div>
      </div>

      {/* ── Slide chips ───────────────────────────────────────────── */}
      {hasDoc && slides.length > 0 && !busy && (
        <div className="flex-shrink-0 px-4 pb-1">
          <div className="flex flex-wrap gap-1 max-h-14 overflow-y-auto">
            {slides.map((s) => (
              <button
                key={s.slide_number}
                onClick={() => onInputChange(`Update slide ${s.slide_number} (${s.title}): `)}
                title={s.title}
                className="rounded-full border border-gray-200 bg-white px-2.5 py-0.5 text-[10px] text-gray-500 hover:border-gray-400 hover:text-gray-700 transition-colors"
              >
                #{s.slide_number} {s.title.slice(0, 20)}{s.title.length > 20 ? "…" : ""}
              </button>
            ))}
          </div>
        </div>
      )}

      {/* ── Composer ─────────────────────────────────────────────── */}
      <div
        className="flex-shrink-0 px-4 pb-5 pt-2"
        onDragOver={(e) => e.preventDefault()}
        onDrop={handleDrop}
      >
        {/* Attached files */}
        {(files.length > 0 || images.length > 0) && (
          <div className="mb-2 flex flex-wrap gap-1.5 px-1">
            {files.map((f) => (
              <span
                key={f.id}
                className="flex items-center gap-1.5 rounded-full border border-gray-200 bg-gray-50 px-2.5 py-1 text-xs text-gray-600"
              >
                {fileTypeIcon(f.type, f.name)}
                <span className="max-w-[120px] truncate">{f.name}</span>
                <button onClick={() => onRemoveFile(f.id)} className="text-gray-300 hover:text-red-400">
                  <X size={10} />
                </button>
              </span>
            ))}
            {images.map((img) => (
              <div key={img.id} className="group relative h-12 w-12 flex-shrink-0 overflow-hidden rounded-xl border border-gray-200">
                <img src={img.preview} alt={img.name} className="h-full w-full object-cover" />
                <button
                  onClick={() => onRemoveImage(img.id)}
                  className="absolute right-0.5 top-0.5 rounded-full bg-black/60 p-0.5 opacity-0 transition-opacity group-hover:opacity-100 hover:text-red-400 text-white"
                >
                  <X size={9} />
                </button>
              </div>
            ))}
          </div>
        )}

        {/* Pill input */}
        <div className="flex items-center gap-2 rounded-full border border-gray-200 bg-white px-3 py-2.5 shadow-sm transition-shadow focus-within:shadow-md focus-within:border-gray-300">
          {/* + button */}
          <button
            type="button"
            title="Attach file"
            onClick={() => fileInputRef.current?.click()}
            className="flex-shrink-0 rounded-full p-1 text-gray-400 hover:bg-gray-100 hover:text-gray-600 transition-colors"
          >
            <Plus size={17} />
          </button>

          {/* Textarea */}
          <textarea
            ref={textareaRef}
            value={input}
            onChange={(e) => onInputChange(e.target.value)}
            onKeyDown={handleKeyDown}
            onInput={handleTextareaInput}
            placeholder="Message ZenSpark"
            disabled={busy}
            rows={1}
            className="max-h-32 flex-1 resize-none bg-transparent text-sm text-gray-800 placeholder-gray-400 focus:outline-none disabled:opacity-50"
            style={{ lineHeight: "1.5" }}
          />

          {/* Right side: image attach + send/mic */}
          <div className="flex flex-shrink-0 items-center gap-1">
            <button
              type="button"
              title="Add image"
              onClick={() => imageInputRef.current?.click()}
              className="rounded-full p-1 text-gray-400 hover:bg-gray-100 hover:text-gray-600 transition-colors"
            >
              <ImageIcon size={15} />
            </button>
            {onLoadFolder && (
              <button
                type="button"
                title="Load folder"
                onClick={onLoadFolder}
                className="rounded-full p-1 text-gray-400 hover:bg-gray-100 hover:text-gray-600 transition-colors"
              >
                <FolderOpen size={15} />
              </button>
            )}
            {input.trim() ? (
              <button
                type="button"
                onClick={onSend}
                disabled={busy}
                className="rounded-full bg-gray-900 p-1.5 text-white hover:bg-gray-700 disabled:opacity-40 transition-colors"
              >
                {busy ? <Loader2 size={14} className="animate-spin" /> : <Send size={14} />}
              </button>
            ) : (
              <button
                type="button"
                className="rounded-full p-1 text-gray-400 hover:bg-gray-100 hover:text-gray-600 transition-colors"
                onClick={() => textareaRef.current?.focus()}
              >
                <Mic size={16} />
              </button>
            )}
          </div>
        </div>

        {/* Quick-action chips */}
        {history.length === 0 && !busy && (
          <div className="mt-3 flex flex-wrap items-center gap-2 px-1">
            {[
              { label: "Ask", icon: <MessageSquare size={13} /> },
              { label: "Create", icon: <Wand2 size={13} /> },
            ].map(({ label, icon }) => (
              <button
                key={label}
                onClick={() => { onInputChange(label + " "); textareaRef.current?.focus(); }}
                className="flex items-center gap-1.5 rounded-full border border-gray-200 bg-white px-3.5 py-1.5 text-sm text-gray-600 hover:border-gray-300 hover:bg-gray-50 transition-colors shadow-sm"
              >
                {icon}
                {label}
              </button>
            ))}
          </div>
        )}
      </div>

      {/* Hidden file inputs */}
      <input ref={fileInputRef} type="file" multiple accept={DOC_ACCEPT} className="hidden" onChange={handleDocFiles} />
      <input ref={imageInputRef} type="file" multiple accept={IMG_ACCEPT} className="hidden" onChange={handleImageFiles} />
    </div>
  );
}

// ── Attachment badge shown inside a chat bubble ───────────────────────────────
function AttachmentBadge({ att }: { att: ChatAttachment }) {
  if (att.type === "image" && att.preview) {
    return (
      <div className="h-16 w-24 flex-shrink-0 overflow-hidden rounded-lg border border-white/20">
        <img src={att.preview} alt={att.name} className="h-full w-full object-cover" />
      </div>
    );
  }
  // File (PDF, PPTX, DOCX, CSV, etc.)
  const ext = att.name.split(".").pop()?.toUpperCase() ?? "FILE";
  const extColor =
    ext === "PDF" ? "bg-red-500/20 text-red-300 border-red-500/30" :
    ext === "PPTX" || ext === "PPT" ? "bg-orange-500/20 text-orange-300 border-orange-500/30" :
    ext === "DOCX" || ext === "DOC" ? "bg-blue-500/20 text-blue-300 border-blue-500/30" :
    ext === "CSV" || ext === "XLSX" ? "bg-green-500/20 text-green-300 border-green-500/30" :
    "bg-gray-500/20 text-gray-700 border-gray-500/30";

  return (
    <div className={`flex items-center gap-1.5 rounded-lg border px-2 py-1 ${extColor}`}>
      <span className="text-[10px] font-bold">{ext}</span>
      <span className="max-w-[120px] truncate text-[11px]">{att.name}</span>
    </div>
  );
}
