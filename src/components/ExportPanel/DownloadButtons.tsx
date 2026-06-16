import { useState } from "react";
import { Download, Presentation, FileText, Code, Loader2 } from "lucide-react";
import type {
  AzureConfig,
  DocumentOutput,
  InputImage,
  OutputFormat,
  TokenUsage,
} from "../../types/document";
import { buildAndDownloadPptx, buildAndDownloadVisualPptx } from "../../services/pptxBuilder";
import { buildAndDownloadDocx } from "../../services/docxBuilder";
import { buildAndDownloadPythonPptx } from "../../services/pythonPptxBuilder";

interface Props {
  doc: DocumentOutput | null;
  images: InputImage[];
  outputFormat: OutputFormat;
  tokenUsage: TokenUsage | null;
  azureConfig: AzureConfig | null;
}

export default function DownloadButtons({
  doc,
  images,
  outputFormat,
  tokenUsage,
  azureConfig,
}: Props) {
  const [downloading, setDownloading] = useState<"visual" | "pptx" | "docx" | "python" | null>(null);
  const [pythonProgress, setPythonProgress] = useState<string | null>(null);
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

  const handleDownloadPython = async () => {
    if (!doc || !azureConfig) return;
    setDownloading("python");
    setPythonProgress("Generating…");
    try {
      await buildAndDownloadPythonPptx(doc, azureConfig, ({ message }) => {
        setPythonProgress(message);
      });
    } catch (e) {
      alert(`Python PPTX failed: ${e instanceof Error ? e.message : String(e)}`);
    } finally {
      setDownloading(null);
      setPythonProgress(null);
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
          <p className="text-xs text-gray-400">Your slides will appear here after generation</p>
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
            {/* Server PPTX is the primary button — always shown when azure config present */}
            {/* Visual PPTX fallback when server not configured */}
            {hasHtmlSlides && !azureConfig && (
              <DownloadBtn
                label={
                  downloading === "visual" && visualProgress
                    ? `Rendering ${visualProgress.done}/${visualProgress.total}…`
                    : "Download PPTX"
                }
                icon={<Presentation size={15} />}
                loading={downloading === "visual"}
                disabled={!!downloading}
                onClick={handleDownloadVisual}
                color="blue"
                description="Looks exactly like the preview"
              />
            )}
            {!hasHtmlSlides && !azureConfig && (
              <DownloadBtn
                label="Download PPTX"
                icon={<Presentation size={15} />}
                loading={downloading === "pptx"}
                disabled={!!downloading}
                onClick={handleDownloadPptx}
                color="blue"
              />
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

        {/* Python PPTX — LLM writes + executes python-pptx script server-side */}
        {azureConfig && (outputFormat === "pptx" || outputFormat === "both") && (
          <DownloadBtn
            label={
              downloading === "python" && pythonProgress
                ? pythonProgress
                : "Server PPTX (Python)"
            }
            icon={<Code size={15} />}
            loading={downloading === "python"}
            disabled={!!downloading}
            onClick={handleDownloadPython}
            color="green"
            description="LLM generates python-pptx code, executed server-side"
          />
        )}
      </div>

      {/* Token usage */}
      {tokenUsage && (
        <div className="rounded-lg border border-gray-200 bg-gray-100/40 px-3 py-2 text-xs">
          <p className="mb-1 text-xs font-medium text-gray-500">Token Usage</p>
          <div className="flex justify-between text-gray-400">
            <span>Prompt</span><span>{tokenUsage.promptTokens.toLocaleString()}</span>
          </div>
          <div className="flex justify-between text-gray-400">
            <span>Completion</span><span>{tokenUsage.completionTokens.toLocaleString()}</span>
          </div>
          <div className="mt-1 flex justify-between font-medium text-gray-500">
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
  color: "blue" | "indigo" | "purple" | "violet" | "green";
  description?: string;
}) {
  const colors = {
    blue:   "border-blue-200 bg-blue-600 text-white hover:bg-blue-700 shadow-sm",
    indigo: "border-indigo-200 bg-indigo-600 text-white hover:bg-indigo-700 shadow-sm",
    purple: "border-purple-200 bg-purple-600 text-white hover:bg-purple-700 shadow-sm",
    violet: "border-violet-200 bg-violet-600 text-white hover:bg-violet-700 shadow-sm",
    green:  "border-emerald-200 bg-emerald-600 text-white hover:bg-emerald-700 shadow-sm",
  };

  return (
    <button
      onClick={onClick}
      disabled={disabled}
      className={`flex w-full flex-col items-center justify-center gap-0.5 rounded-lg border px-3 py-2.5 text-sm font-medium transition-colors disabled:cursor-not-allowed disabled:opacity-40 ${colors[color]}`}
    >
      <span className="flex items-center gap-2">
        {loading ? <Loader2 size={14} className="animate-spin" /> : icon}
        {label}
      </span>
      {description && (
        <span className="text-[10px] font-normal opacity-75">{description}</span>
      )}
    </button>
  );
}
