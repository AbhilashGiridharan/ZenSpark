import { useState } from "react";
import {
  ChevronDown,
  ChevronUp,
  SlidersHorizontal,
  FileText,
} from "lucide-react";
import type { DocumentOutput, Slide } from "../../types/document";

interface Props {
  doc: DocumentOutput | null;
  isGenerating: boolean;
  streamingText: string;
}

const LAYOUT_ICONS: Record<string, string> = {
  title: "🎯",
  bullets: "📋",
  two_column: "◫",
  image_caption: "🖼",
  table: "📊",
  quote: "❝",
  section_divider: "─",
  agenda: "📑",
  closing: "✅",
  stats: "📈",
};

// Extract partial slide info from streaming JSON
function extractStreamingSlides(text: string) {
  const result: Array<{ num: number; title: string; layout: string }> = [];
  const re = /"slide_number"\s*:\s*(\d+)/g;
  let m;
  while ((m = re.exec(text)) !== null) {
    const num = parseInt(m[1]);
    const snippet = text.slice(m.index, Math.min(m.index + 700, text.length));
    const titleM = /"title"\s*:\s*"((?:[^"\\]|\\.)*)"/g.exec(snippet);
    const layoutM = /"layout"\s*:\s*"([a-z_]+)"/g.exec(snippet);
    result.push({
      num,
      title: titleM ? titleM[1].replace(/\\"/g, '"') : `Slide ${num}`,
      layout: layoutM ? layoutM[1] : "bullets",
    });
  }
  return result;
}

export default function SlideOutline({ doc, isGenerating, streamingText }: Props) {
  const [expanded, setExpanded] = useState<number | null>(null);

  // ── Streaming state: show extracted slide cards live ─────────────────────
  if (isGenerating) {
    const partialSlides = extractStreamingSlides(streamingText);
    return (
      <div className="flex h-full flex-col gap-2">
        <div className="flex items-center gap-2">
          <SlidersHorizontal size={14} className="text-blue-400 animate-pulse" />
          <span className="text-sm font-medium text-gray-300">
            Generating
            {partialSlides.length > 0 ? ` — ${partialSlides.length} slides so far…` : "…"}
          </span>
        </div>
        {partialSlides.length === 0 ? (
          <div className="flex flex-1 flex-col items-center justify-center gap-3">
            {[...Array(5)].map((_, i) => (
              <div
                key={i}
                className="w-full h-10 rounded-lg bg-gray-800/60 animate-pulse"
                style={{ opacity: 1 - i * 0.15 }}
              />
            ))}
            <p className="text-xs text-gray-600">Waiting for slides…</p>
          </div>
        ) : (
          <div className="flex-1 overflow-y-auto flex flex-col gap-1.5 pr-1">
            {partialSlides.map((s, i) => (
              <div
                key={s.num}
                className="flex items-center gap-2.5 rounded-lg border border-gray-800 bg-gray-900/60 px-3 py-2"
                style={{ animation: `fadeIn 0.3s ease ${i * 0.05}s both` }}
              >
                <span className="flex-shrink-0 text-sm">
                  {LAYOUT_ICONS[s.layout] ?? "📋"}
                </span>
                <div className="min-w-0 flex-1">
                  <p className="truncate text-xs font-medium text-gray-300">{s.title}</p>
                  <p className="text-[10px] text-gray-600 capitalize">{s.layout.replace("_", " ")}</p>
                </div>
                <span className="flex-shrink-0 text-[10px] text-gray-700">#{s.num}</span>
              </div>
            ))}
            {/* pulse card for next incoming slide */}
            <div className="h-9 w-full animate-pulse rounded-lg bg-gray-800/40" />
          </div>
        )}
      </div>
    );
  }

  // ── Empty state ────────────────────────────────────────────────────────────
  if (!doc) {
    return (
      <div className="flex h-full flex-col items-center justify-center gap-3 text-center text-gray-600">
        <FileText size={40} strokeWidth={1} />
        <div>
          <p className="text-sm font-medium text-gray-500">No document yet</p>
          <p className="mt-1 text-xs">
            Configure inputs on the left, then click Generate
          </p>
        </div>
      </div>
    );
  }

  // ── PPTX slides ────────────────────────────────────────────────────────────
  const slides = doc.slides ?? [];
  const sections = doc.sections ?? [];

  return (
    <div className="flex h-full flex-col gap-1 overflow-y-auto pr-1">
      {/* Document header */}
      <div className="mb-2 rounded-lg border border-blue-900/40 bg-blue-950/20 px-3 py-2">
        <p className="text-sm font-semibold text-blue-300">{doc.title}</p>
        <p className="text-xs text-gray-500">
          {doc.document_type.toUpperCase()} · {doc.theme.replace("_", " ")} · {doc.author} · {doc.date}
        </p>
      </div>

      {/* Slides */}
      {slides.map((s) => (
        <SlideCard
          key={s.slide_number}
          slide={s}
          isExpanded={expanded === s.slide_number}
          onToggle={() =>
            setExpanded((p) => (p === s.slide_number ? null : s.slide_number))
          }
        />
      ))}

      {/* DOCX sections */}
      {sections.map((sec, i) => (
        <div
          key={i}
          className="rounded-lg border border-gray-800 bg-gray-900/40 px-3 py-2"
        >
          <p
            className={`font-medium text-gray-300 ${
              sec.level === 1
                ? "text-sm"
                : sec.level === 2
                ? "text-xs pl-3"
                : "text-xs pl-6 text-gray-400"
            }`}
          >
            {sec.heading}
          </p>
          <p className="mt-0.5 text-xs text-gray-600 line-clamp-2">
            {sec.paragraphs[0] ?? ""}
          </p>
        </div>
      ))}
    </div>
  );
}

function SlideCard({
  slide,
  isExpanded,
  onToggle,
}: {
  slide: Slide;
  isExpanded: boolean;
  onToggle: () => void;
}) {
  const icon = LAYOUT_ICONS[slide.layout] ?? "□";

  return (
    <div className="rounded-lg border border-gray-800 bg-gray-900/50 transition-colors hover:border-gray-700">
      <button
        onClick={onToggle}
        className="flex w-full items-center gap-2 px-3 py-2 text-left"
      >
        <span className="w-5 flex-shrink-0 text-center text-xs text-gray-500">
          {slide.slide_number}
        </span>
        <span className="text-sm">{icon}</span>
        <span className="flex-1 truncate text-sm text-gray-300">{slide.title}</span>
        <span className="text-xs text-gray-600">{slide.layout}</span>
        {isExpanded ? (
          <ChevronUp size={13} className="flex-shrink-0 text-gray-600" />
        ) : (
          <ChevronDown size={13} className="flex-shrink-0 text-gray-600" />
        )}
      </button>

      {isExpanded && (
        <div className="border-t border-gray-800 px-3 pb-3 pt-2 text-xs">
          {slide.subtitle && (
            <p className="mb-1 text-gray-400">{slide.subtitle}</p>
          )}
          {slide.bullets && slide.bullets.length > 0 && (
            <ul className="space-y-0.5">
              {slide.bullets.map((b, i) => (
                <li key={i} className="flex gap-2 text-gray-500">
                  <span className="text-blue-600">•</span>
                  <span>{b}</span>
                </li>
              ))}
            </ul>
          )}
          {slide.quote && (
            <p className="italic text-gray-400">"{slide.quote}"</p>
          )}
          {slide.table && (
            <p className="text-gray-600">
              Table: {slide.table.headers.join(", ")} ({slide.table.rows.length} rows)
            </p>
          )}
          {slide.speaker_notes && (
            <div className="mt-2 rounded border-l-2 border-blue-800 bg-blue-950/20 px-2 py-1">
              <p className="text-gray-500">{slide.speaker_notes}</p>
            </div>
          )}
        </div>
      )}
    </div>
  );
}
