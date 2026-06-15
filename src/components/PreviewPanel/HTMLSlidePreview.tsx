import { useRef, useEffect, useState } from "react";
import { ChevronLeft, ChevronRight } from "lucide-react";
import type { Slide } from "../../types/document";

interface Props {
  slides: Slide[];
}

// Renders LLM-generated HTML slides in a 16:9 viewer with prev/next navigation
export default function HTMLSlidePreview({ slides }: Props) {
  const [current, setCurrent] = useState(0);
  const iframeRef = useRef<HTMLIFrameElement>(null);

  const slide = slides[current];

  // Write HTML into iframe whenever slide changes
  useEffect(() => {
    const iframe = iframeRef.current;
    if (!iframe || !slide?.html) return;
    const doc = iframe.contentDocument ?? iframe.contentWindow?.document;
    if (!doc) return;
    doc.open();
    doc.write(`<!DOCTYPE html><html><head>
      <meta charset="utf-8"/>
      <style>
        * { margin:0; padding:0; box-sizing:border-box; }
        body { width:960px; height:540px; overflow:hidden; }
      </style>
    </head><body>${slide.html}</body></html>`);
    doc.close();
  }, [slide]);

  if (!slides.length) return null;

  const prev = () => setCurrent((c) => Math.max(0, c - 1));
  const next = () => setCurrent((c) => Math.min(slides.length - 1, c + 1));

  return (
    <div className="flex flex-col gap-2">
      {/* 16:9 slide viewer */}
      <div className="relative w-full overflow-hidden rounded-lg border border-gray-700 bg-black shadow-lg"
           style={{ paddingBottom: "56.25%" }}>
        <iframe
          ref={iframeRef}
          title={`Slide ${current + 1}`}
          sandbox="allow-same-origin"
          className="absolute inset-0 h-full w-full border-0"
          style={{ transformOrigin: "top left" }}
        />
      </div>

      {/* Navigation */}
      <div className="flex items-center justify-between px-1">
        <button
          onClick={prev}
          disabled={current === 0}
          className="rounded-md p-1 text-gray-500 hover:bg-gray-800 hover:text-gray-300 disabled:opacity-30"
        >
          <ChevronLeft size={18} />
        </button>
        <span className="text-xs text-gray-500">
          {current + 1} / {slides.length}
          <span className="ml-2 text-gray-700">{slide?.title}</span>
        </span>
        <button
          onClick={next}
          disabled={current === slides.length - 1}
          className="rounded-md p-1 text-gray-500 hover:bg-gray-800 hover:text-gray-300 disabled:opacity-30"
        >
          <ChevronRight size={18} />
        </button>
      </div>

      {/* Thumbnail strip */}
      <div className="flex gap-1.5 overflow-x-auto pb-1">
        {slides.map((s, i) => (
          <button
            key={s.slide_number}
            onClick={() => setCurrent(i)}
            className={`relative flex-shrink-0 overflow-hidden rounded border transition-colors ${
              i === current ? "border-blue-500" : "border-gray-700 hover:border-gray-500"
            }`}
            style={{ width: 96, height: 54 }}
          >
            <SlideThumb slide={s} />
            <span className="absolute bottom-0.5 left-1 text-[9px] font-bold text-white/70 drop-shadow">
              {s.slide_number}
            </span>
          </button>
        ))}
      </div>
    </div>
  );
}

// Mini thumbnail — renders slide HTML in a tiny scaled iframe
function SlideThumb({ slide }: { slide: Slide }) {
  const ref = useRef<HTMLIFrameElement>(null);

  useEffect(() => {
    const iframe = ref.current;
    if (!iframe || !slide.html) return;
    const doc = iframe.contentDocument ?? iframe.contentWindow?.document;
    if (!doc) return;
    doc.open();
    doc.write(`<!DOCTYPE html><html><head>
      <meta charset="utf-8"/>
      <style>
        * { margin:0; padding:0; box-sizing:border-box; }
        body { width:960px; height:540px; overflow:hidden; transform:scale(0.1); transform-origin:top left; }
      </style>
    </head><body>${slide.html}</body></html>`);
    doc.close();
  }, [slide]);

  return (
    <iframe
      ref={ref}
      title={`thumb-${slide.slide_number}`}
      sandbox="allow-same-origin"
      className="border-0"
      style={{ width: 960, height: 540, transform: "scale(0.1)", transformOrigin: "top left", pointerEvents: "none" }}
    />
  );
}
