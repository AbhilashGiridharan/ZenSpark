import { useEffect, useRef } from "react";
import { X, Image as ImageIcon } from "lucide-react";
import type { InputImage } from "../../types/document";
import { clipboardItemToInputImage, fileToInputImage } from "../../services/fileExtractor";

interface Props {
  images: InputImage[];
  onAdd: (image: InputImage) => void;
  onRemove: (id: string) => void;
  onCaptionChange: (id: string, caption: string) => void;
}

export default function ScreenshotPaste({
  images,
  onAdd,
  onRemove,
  onCaptionChange,
}: Props) {
  const dropRef = useRef<HTMLDivElement>(null);

  // Global paste handler (Ctrl+V / Cmd+V anywhere on page)
  useEffect(() => {
    const handlePaste = async (e: ClipboardEvent) => {
      const items = Array.from(e.clipboardData?.items ?? []);
      for (const item of items) {
        if (item.type.startsWith("image/")) {
          const img = await clipboardItemToInputImage(item);
          if (img) onAdd(img);
        }
      }
    };
    document.addEventListener("paste", handlePaste);
    return () => document.removeEventListener("paste", handlePaste);
  }, [onAdd]);

  // Drag-and-drop for images on this panel
  const handleDrop = async (e: React.DragEvent) => {
    e.preventDefault();
    const files = Array.from(e.dataTransfer.files).filter((f) =>
      f.type.startsWith("image/")
    );
    for (const f of files) {
      const img = await fileToInputImage(f);
      onAdd(img);
    }
  };

  const handleFileInput = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = Array.from(e.target.files ?? []);
    for (const f of files) {
      const img = await fileToInputImage(f);
      onAdd(img);
    }
    e.target.value = "";
  };

  return (
    <div>
      {/* Drop zone */}
      <div
        ref={dropRef}
        onDragOver={(e) => e.preventDefault()}
        onDrop={handleDrop}
        className="relative flex cursor-pointer flex-col items-center justify-center rounded-lg border-2 border-dashed border-gray-700 p-3 transition-colors hover:border-gray-500 hover:bg-gray-800/20"
      >
        <label className="flex w-full cursor-pointer flex-col items-center gap-1">
          <ImageIcon size={18} className="text-gray-500" />
          <span className="text-center text-xs text-gray-400">
            Paste screenshot (Ctrl+V) or drag & drop
          </span>
          <span className="text-center text-xs text-gray-600">
            Click to upload · PNG, JPEG, WebP
          </span>
          <input
            type="file"
            accept="image/*"
            multiple
            className="hidden"
            onChange={handleFileInput}
          />
        </label>
      </div>

      {/* Image thumbnails */}
      {images.length > 0 && (
        <div className="mt-2 grid grid-cols-2 gap-2">
          {images.map((img) => (
            <div key={img.id} className="group relative rounded-md border border-gray-700 bg-gray-800/50 overflow-hidden">
              {/* Thumbnail */}
              <div className="relative">
                <img
                  src={img.preview}
                  alt={img.name}
                  className="h-20 w-full object-cover"
                />
                <button
                  onClick={() => onRemove(img.id)}
                  className="absolute right-1 top-1 rounded-full bg-black/60 p-0.5 text-gray-300 opacity-0 transition-opacity group-hover:opacity-100 hover:text-red-400"
                >
                  <X size={11} />
                </button>
                <span className="absolute bottom-1 left-1 rounded bg-black/60 px-1 text-xs text-gray-300">
                  #{images.indexOf(img)}
                </span>
              </div>
              {/* Caption */}
              <input
                type="text"
                value={img.caption}
                onChange={(e) => onCaptionChange(img.id, e.target.value)}
                placeholder="Add caption…"
                className="w-full border-t border-gray-700 bg-transparent px-2 py-1 text-xs text-gray-400 placeholder-gray-600 focus:outline-none focus:ring-0"
              />
            </div>
          ))}
        </div>
      )}
    </div>
  );
}
