import { useRef, useEffect, useState } from "react";
import { ChevronLeft, ChevronRight } from "lucide-react";
import type { Slide } from "../../types/document";

// Native slide dimensions (16:9)
const SLIDE_W = 960;
const SLIDE_H = 540;

interface Props {
  slides: Slide[];
}

function writeSlideHTML(iframe: HTMLIFrameElement, html: string) {
  const doc = iframe.contentDocument ?? iframe.contentWindow?.document;
  if (!doc) return;
  doc.open();
  doc.write(`<!DOCTYPE html><html><head>
    <meta charset="utf-8"/>
    <style>
      *{margin:0;padding:0;box-sizing:border-box;}
      html,body{width:${SLIDE_W}px;height:${SLIDE_H}px;overflow:hidden;background:#000;}
    </style>
  </head><body>${html}</body></html>`);
  doc.close();
}

export default function HTMLSlidePreview({ slides }: Props) {
  const [current, setCurrent] = useState(0);
  const iframeRef = useRef<HTMLIFrameElement>(null);
  const slideAreaRef = useRef<HTMLDivElement>(null);
  const [scale, setScale] = useState(1);
  const [offset, setOffset] = useState({ x: 0, y: 0 });

  // Scale to fill available area (fit inside, centered)
  useEffect(() => {
    const el = slideAreaRef.current;
    if (!el) return;
    const observer = new ResizeObserver(([entry]) => {
      const { width, height } = entry.contentRect;
      const s = Math.min(width / SLIDE_W, height / SLIDE_H);
      setScale(s);
      setOffset({
        x: Math.round((width - SLIDE_W * s) / 2),
        y: Math.round((height - SLIDE_H * s) / 2),
      });
    });
    observer.observe(el);
    return () => observer.disconnect();
  }, []);

  const slide = slides[current];

  useEffect(() => {
    if (!iframeRef.current || !slide?.html) return;
    writeSlideHTML(iframeRef.current, slide.html);
  }, [slide]);

  if (!slides.length) return null;

  const prev = () => setCurrent((c) => Math.max(0, c - 1));
  const next = () => setCurrent((c) => Math.min(slides.length - 1, c + 1));

  return (
    <div
      className="flex h-full flex-col gap-2 outline-none"
      tabIndex={0}
      onKeyDown={(e) => {
        if (e.key === "ArrowLeft") prev();
        if (e.key === "ArrowRight") next();
      }}
    >
      {/* ── Main slide area — fills remaining height ── */}
      <div
        ref={slideAreaRef}
        className="relative min-h-0 flex-1 overflow-hidden rounded-lg border border-gray-700 bg-gray-950 shadow-xl"
      >
        <iframe
          ref={iframeRef}
          title={`Slide ${current + 1}`}
          sandbox="allow-same-origin"
          style={{
            position: "absolute",
            top: offset.y,
            left: offset.x,
            width: SLIDE_W,
            height: SLIDE_H,
            border: "none",
            transform: `scale(${scale})`,
            transformOrigin: "top left",
            display: "block",
          }}
        />
      </div>

      {/* ── Navigation bar ── */}
      <div className="flex flex-shrink-0 items-center gap-1 px-0.5">
        <button
          onClick={prev}
          disabled={current === 0}
          className="rounded-md p-1 text-gray-500 hover:bg-gray-800 hover:text-gray-200 disabled:opacity-25"
        >
          <ChevronLeft size={16} />
        </button>
        <div className="flex-1 text-center">
          <span className="text-xs font-medium text-gray-400">
            {current + 1} <span className="text-gray-700">/</span> {slides.length}
          </span>
          {slide?.title && (
            <span className="ml-2 max-w-[180px] truncate align-middle text-xs text-gray-600">
              {slide.title}
            </span>
          )}
        </div>
        <button
          onClick={next}
          disabled={current === slides.length - 1}
          className="rounded-md p-1 text-gray-500 hover:bg-gray-800 hover:text-gray-200 disabled:opacity-25"
        >
          <ChevronRight size={16} />
        </button>
      </div>

      {/* ── Thumbnail strip ── */}
      <div className="flex flex-shrink-0 gap-1.5 overflow-x-auto pb-1">
        {slides.map((s, i) => (
          <ThumbButton
            key={s.slide_number}
            slide={s}
            index={i}
            isCurrent={i === current}
            onClick={() => setCurrent(i)}
          />
        ))}
      </div>
    </div>
  );
}

// Thumbnail — properly scaled iframe
function ThumbButton({ slide, index, isCurrent, onClick }: {
  slide: Slide;
  index: number;
  isCurrent: boolean;
  onClick: () => void;
}) {
  const ref = useRef<HTMLIFrameElement>(null);
  const THUMB_W = 112;
  const THUMB_H = 63;
  const thumbScale = THUMB_W / SLIDE_W;

  useEffect(() => {
    if (!ref.current || !slide.html) return;
    writeSlideHTML(ref.current, slide.html);
  }, [slide]);

  return (
    <button
      onClick={onClick}
      title={slide.title}
      className={`relative flex-shrink-0 overflow-hidden rounded-md border-2 transition-all ${
        isCurrent
          ? "border-blue-500 shadow-lg shadow-blue-900/40"
          : "border-gray-700/60 hover:border-gray-500"
      }`}
      style={{ width: THUMB_W, height: THUMB_H }}
    >
      {slide.html ? (
        <iframe
          ref={ref}
          sandbox="allow-same-origin"
          style={{
            width: SLIDE_W,
            height: SLIDE_H,
            border: "none",
            transform: `scale(${thumbScale})`,
            transformOrigin: "top left",
            pointerEvents: "none",
            display: "block",
          }}
        />
      ) : (
        <div className="flex h-full w-full items-center justify-center bg-gray-900 text-[10px] text-gray-600">
          {slide.layout}
        </div>
      )}
      {/* Slide number badge */}
      <span className="absolute bottom-0.5 right-1 rounded bg-black/60 px-1 text-[9px] font-bold text-white/80">
        {index + 1}
      </span>
    </button>
  );
}

