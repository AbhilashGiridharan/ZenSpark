import { useState } from "react";
import { Download, Presentation, FileText } from "lucide-react";
import type {
  DocumentOutput,
  InputImage,
  OutputFormat,
  TokenUsage,
} from "../../types/document";
import { buildAndDownloadPptx, buildAndDownloadVisualPptx } from "../../services/pptxBuilder";
import { buildAndDownloadDocx } from "../../services/docxBuilder";

interface Props {
  doc: DocumentOutput | null;
  images: InputImage[];
  outputFormat: OutputFormat;
  tokenUsage: TokenUsage | null;
}

export default function DownloadButtons({
  doc,
  images,
  outputFormat,
  tokenUsage,
}: Props) {
  const [downloading, setDownloading] = useState<"visual" | "pptx" | "docx" | null>(null);
  const [visualProgress, setVisualProgress] = useState<{ done: number; total: number } | null>(null);

  const hasHtmlSlides = (doc?.slides ?? []).some((s) => s.html);

  const handleDownloadVisual = async () => {
    if (!doc) return;
    setDownloading("visual");
    setVisualProgress({ done: 0, total: doc.slides?.length ?? 0 });
    try {
      await buildAndDownloadVisualPptx(doc, (done, total) => {
        setVisualProgress({ done, total });
      });
    } catch (e) {
      alert(`Visual PPTX failed: ${e instanceof Error ? e.message : String(e)}`);
    } finally {
      setDownloading(null);
      setVisualProgress(null);
    }
  };

  const handleDownloadPptx = async () => {
    if (!doc) return;
    setDownloading("pptx");
    try {
      await buildAndDownloadPptx(doc, images);
    } catch (e) {
      alert(`PPTX generation failed: ${e instanceof Error ? e.message : String(e)}`);
    } finally {
      setDownloading(null);
    }
  };

  const handleDownloadDocx = async () => {
    if (!doc) return;
    setDownloading("docx");
    try {
      await buildAndDownloadDocx(doc, images);
    } catch (e) {
      alert(`DOCX generation failed: ${e instanceof Error ? e.message : String(e)}`);
    } finally {
      setDownloading(null);
    }
  };

  const handleDownloadBoth = async () => {
    if (!doc) return;
    setDownloading("pptx");
    try {
      await buildAndDownloadPptx(doc, images);
      await buildAndDownloadDocx(doc, images);
    } catch (e) {
      alert(`Generation failed: ${e instanceof Error ? e.message : String(e)}`);
    } finally {
      setDownloading(null);
    }
  };

  if (!doc) {
    return (
      <div className="flex h-full items-center justify-center text-center px-4">
        <div>
          <Presentation size={32} strokeWidth={1} className="mx-auto mb-2 text-gray-700" />
          <p className="text-xs text-gray-600">Your slides will appear here after generation</p>
        </div>
      </div>
    );
  }

  return (
    <div className="flex flex-col gap-3">
      {/* Download buttons */}
      <div className="space-y-2">
        {(outputFormat === "pptx" || outputFormat === "both") && (
          <>
            {/* Primary: Editable PPTX — real PowerPoint shapes/text */}
            <DownloadBtn
              label="Download PPTX (Editable)"
              icon={<Presentation size={15} />}
              loading={downloading === "pptx"}
              disabled={!!downloading}
              onClick={handleDownloadPptx}
              color="blue"
              description="Fully editable in PowerPoint"
            />
            {/* Secondary: Visual PPTX — pixel-perfect screenshot, only when HTML exists */}
            {hasHtmlSlides && (
              <button
                onClick={handleDownloadVisual}
                disabled={!!downloading}
                className="w-full text-center text-[11px] text-gray-600 underline-offset-2 hover:text-gray-400 hover:underline disabled:opacity-40"
              >
                {downloading === "visual" && visualProgress
                  ? `Rendering ${visualProgress.done}/${visualProgress.total}…`
                  : "Download visual version (matches preview, image-based)"}
              </button>
            )}
          </>
        )}
        {(outputFormat === "docx" || outputFormat === "both") && (
          <DownloadBtn
            label="Download DOCX"
            icon={<FileText size={15} />}
            loading={downloading === "docx"}
            disabled={!!downloading}
            onClick={handleDownloadDocx}
            color="indigo"
          />
        )}
        {outputFormat === "both" && (
          <DownloadBtn
            label="Download Both"
            icon={<Download size={15} />}
            loading={!!downloading}
            disabled={!!downloading}
            onClick={handleDownloadBoth}
            color="purple"
          />
        )}
      </div>

      {/* Token usage */}
      {tokenUsage && (
        <div className="rounded-lg border border-gray-800 bg-gray-900/40 px-3 py-2 text-xs">
          <p className="mb-1 text-xs font-medium text-gray-500">Token Usage</p>
          <div className="flex justify-between text-gray-600">
            <span>Prompt</span><span>{tokenUsage.promptTokens.toLocaleString()}</span>
          </div>
          <div className="flex justify-between text-gray-600">
            <span>Completion</span><span>{tokenUsage.completionTokens.toLocaleString()}</span>
          </div>
          <div className="mt-1 flex justify-between font-medium text-gray-400">
            <span>Total</span><span>{tokenUsage.totalTokens.toLocaleString()}</span>
          </div>
        </div>
      )}
    </div>
  );
}

function DownloadBtn({
  label,
  icon,
  loading,
  disabled,
  onClick,
  color,
  description,
}: {
  label: string;
  icon: React.ReactNode;
  loading: boolean;
  disabled: boolean;
  onClick: () => void;
  color: "blue" | "indigo" | "purple" | "violet";
  description?: string;
}) {
  const colors = {
    blue: "border-blue-700 text-blue-400 hover:bg-blue-900/30",
    indigo: "border-indigo-700 text-indigo-400 hover:bg-indigo-900/30",
    purple: "border-purple-700 text-purple-400 hover:bg-purple-900/30",
    violet: "border-violet-600 bg-violet-900/20 text-violet-300 hover:bg-violet-900/40",
  };

  return (
    <button
      onClick={onClick}
      disabled={disabled}
      className={`flex w-full flex-col items-center justify-center gap-0.5 rounded-lg border px-3 py-2 text-sm font-medium transition-colors disabled:cursor-not-allowed disabled:opacity-40 ${colors[color]}`}
    >
      <span className="flex items-center gap-2">
        {loading ? <Loader2 size={14} className="animate-spin" /> : icon}
        {label}
      </span>
      {description && (
        <span className="text-[10px] font-normal opacity-60">{description}</span>
      )}
    </button>
  );
}
