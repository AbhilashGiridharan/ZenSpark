import { useState } from "react";
import { Download, Presentation, FileText, Layers, Loader2 } from "lucide-react";
import type {
  DocumentOutput,
  InputImage,
  OutputFormat,
  ThemeOption,
  UseCasePreset,
  TokenUsage,
} from "../../types/document";
import {
  OUTPUT_FORMAT_LABELS,
  THEME_LABELS,
  USE_CASE_LABELS,
} from "../../services/promptTemplates";
import { buildAndDownloadPptx } from "../../services/pptxBuilder";
import { buildAndDownloadDocx } from "../../services/docxBuilder";

interface Props {
  doc: DocumentOutput | null;
  images: InputImage[];
  outputFormat: OutputFormat;
  theme: ThemeOption;
  useCase: UseCasePreset;
  tokenUsage: TokenUsage | null;
  onOutputFormatChange: (f: OutputFormat) => void;
  onThemeChange: (t: ThemeOption) => void;
  onUseCaseChange: (uc: UseCasePreset) => void;
  onGenerate: () => void;
  isGenerating: boolean;
  customPrompt: string;
  onCustomPromptChange: (v: string) => void;
}

export default function DownloadButtons({
  doc,
  images,
  outputFormat,
  theme,
  useCase,
  tokenUsage,
  onOutputFormatChange,
  onThemeChange,
  onUseCaseChange,
  onGenerate,
  isGenerating,
  customPrompt,
  onCustomPromptChange,
}: Props) {
  const [downloading, setDownloading] = useState<"pptx" | "docx" | null>(null);

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

  return (
    <div className="flex h-full flex-col gap-4 overflow-y-auto">
      {/* Use Case */}
      <Section title="Use Case">
        <select
          value={useCase}
          onChange={(e) => onUseCaseChange(e.target.value as UseCasePreset)}
          className="w-full rounded-lg border border-gray-700 bg-gray-800/40 px-2 py-1.5 text-xs text-gray-200 focus:border-blue-600 focus:outline-none"
        >
          {(Object.keys(USE_CASE_LABELS) as UseCasePreset[]).map((k) => (
            <option key={k} value={k}>{USE_CASE_LABELS[k]}</option>
          ))}
        </select>
        {useCase === "custom" && (
          <textarea
            value={customPrompt}
            onChange={(e) => onCustomPromptChange(e.target.value)}
            placeholder="Describe what to generate…"
            rows={3}
            className="mt-2 w-full resize-none rounded-lg border border-gray-700 bg-gray-800/40 px-3 py-2 text-xs text-gray-200 placeholder-gray-600 focus:border-blue-600 focus:outline-none"
          />
        )}
      </Section>

      {/* Output Format */}
      <Section title="Output Format">
        <div className="space-y-1">
          {(Object.keys(OUTPUT_FORMAT_LABELS) as OutputFormat[]).map((f) => (
            <label key={f} className="flex cursor-pointer items-center gap-2">
              <input
                type="radio"
                name="outputFormat"
                value={f}
                checked={outputFormat === f}
                onChange={() => onOutputFormatChange(f)}
                className="accent-blue-500"
              />
              <span className="text-xs text-gray-300">{OUTPUT_FORMAT_LABELS[f]}</span>
            </label>
          ))}
        </div>
      </Section>

      {/* Theme */}
      <Section title="Theme">
        <select
          value={theme}
          onChange={(e) => onThemeChange(e.target.value as ThemeOption)}
          className="w-full rounded-lg border border-gray-700 bg-gray-800/40 px-2 py-1.5 text-xs text-gray-200 focus:border-blue-600 focus:outline-none"
        >
          {(Object.keys(THEME_LABELS) as ThemeOption[]).map((t) => (
            <option key={t} value={t}>{THEME_LABELS[t]}</option>
          ))}
        </select>
      </Section>

      {/* Generate button */}
      <button
        onClick={onGenerate}
        disabled={isGenerating}
        className="flex w-full items-center justify-center gap-2 rounded-lg bg-blue-600 py-2.5 text-sm font-semibold text-white transition-colors hover:bg-blue-500 disabled:cursor-not-allowed disabled:opacity-60"
      >
        {isGenerating ? (
          <>
            <Loader2 size={15} className="animate-spin" />
            Generating…
          </>
        ) : (
          <>
            <Layers size={15} />
            Generate
          </>
        )}
      </button>

      {/* Divider */}
      <div className="border-t border-gray-800" />

      {/* Download buttons */}
      <div className="space-y-2">
        {(outputFormat === "pptx" || outputFormat === "both") && (
          <DownloadBtn
            label="Download PPTX"
            icon={<Presentation size={14} />}
            loading={downloading === "pptx"}
            disabled={!doc || !!downloading}
            onClick={handleDownloadPptx}
            color="blue"
          />
        )}
        {(outputFormat === "docx" || outputFormat === "both") && (
          <DownloadBtn
            label="Download DOCX"
            icon={<FileText size={14} />}
            loading={downloading === "docx"}
            disabled={!doc || !!downloading}
            onClick={handleDownloadDocx}
            color="indigo"
          />
        )}
        {outputFormat === "both" && (
          <DownloadBtn
            label="Download Both"
            icon={<Download size={14} />}
            loading={!!downloading}
            disabled={!doc || !!downloading}
            onClick={handleDownloadBoth}
            color="purple"
          />
        )}
      </div>

      {/* Token usage */}
      {tokenUsage && (
        <div className="rounded-lg border border-gray-800 bg-gray-900/40 px-3 py-2 text-xs">
          <p className="font-medium text-gray-400">Token Usage</p>
          <div className="mt-1 space-y-0.5 text-gray-600">
            <p>Prompt: {tokenUsage.promptTokens.toLocaleString()}</p>
            <p>Completion: {tokenUsage.completionTokens.toLocaleString()}</p>
            <p className="text-gray-500">Total: {tokenUsage.totalTokens.toLocaleString()}</p>
          </div>
        </div>
      )}

      {!doc && (
        <p className="text-center text-xs text-gray-700">
          Generate a document to enable downloads
        </p>
      )}
    </div>
  );
}

function Section({ title, children }: { title: string; children: React.ReactNode }) {
  return (
    <div>
      <p className="mb-1.5 text-xs font-medium uppercase tracking-wider text-gray-500">
        {title}
      </p>
      {children}
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
}: {
  label: string;
  icon: React.ReactNode;
  loading: boolean;
  disabled: boolean;
  onClick: () => void;
  color: "blue" | "indigo" | "purple";
}) {
  const colors = {
    blue: "border-blue-700 text-blue-400 hover:bg-blue-900/30",
    indigo: "border-indigo-700 text-indigo-400 hover:bg-indigo-900/30",
    purple: "border-purple-700 text-purple-400 hover:bg-purple-900/30",
  };

  return (
    <button
      onClick={onClick}
      disabled={disabled}
      className={`flex w-full items-center justify-center gap-2 rounded-lg border px-3 py-2 text-sm font-medium transition-colors disabled:cursor-not-allowed disabled:opacity-40 ${colors[color]}`}
    >
      {loading ? <Loader2 size={14} className="animate-spin" /> : icon}
      {label}
    </button>
  );
}
