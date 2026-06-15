import { useCallback } from "react";
import { useDropzone } from "react-dropzone";
import { Upload, X, FileText, FileSpreadsheet, File } from "lucide-react";
import type { InputFile } from "../../types/document";
import { fileToInputFile, formatFileSize } from "../../services/fileExtractor";

interface Props {
  files: InputFile[];
  onAdd: (files: InputFile[]) => void;
  onRemove: (id: string) => void;
}

const ACCEPT = {
  "application/pdf": [".pdf"],
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document": [".docx"],
  "text/plain": [".txt"],
  "text/markdown": [".md"],
  "text/csv": [".csv"],
  "text/html": [".html", ".htm"],
};

function fileIcon(type: string) {
  if (type.includes("pdf") || type.includes("word")) return <FileText size={14} className="text-blue-400" />;
  if (type.includes("csv") || type.includes("spreadsheet")) return <FileSpreadsheet size={14} className="text-green-400" />;
  return <File size={14} className="text-gray-400" />;
}

export default function FileUpload({ files, onAdd, onRemove }: Props) {
  const onDrop = useCallback(
    async (accepted: File[]) => {
      const results: InputFile[] = [];
      for (const f of accepted) {
        try {
          results.push(await fileToInputFile(f));
        } catch (e) {
          console.warn("Could not extract", f.name, e);
        }
      }
      if (results.length > 0) onAdd(results);
    },
    [onAdd]
  );

  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    onDrop,
    accept: ACCEPT,
    multiple: true,
    maxSize: 50 * 1024 * 1024, // 50 MB
  });

  return (
    <div>
      {/* Drop zone */}
      <div
        {...getRootProps()}
        className={`flex cursor-pointer flex-col items-center justify-center rounded-lg border-2 border-dashed p-4 transition-colors ${
          isDragActive
            ? "border-blue-500 bg-blue-950/30"
            : "border-gray-700 hover:border-gray-500 hover:bg-gray-800/30"
        }`}
      >
        <input {...getInputProps()} />
        <Upload size={20} className={isDragActive ? "text-blue-400" : "text-gray-500"} />
        <p className="mt-2 text-center text-xs text-gray-400">
          {isDragActive ? "Drop files here…" : "Drop files or click to upload"}
        </p>
        <p className="mt-1 text-center text-xs text-gray-600">
          PDF, DOCX, TXT, MD, CSV, HTML
        </p>
      </div>

      {/* File list */}
      {files.length > 0 && (
        <ul className="mt-2 space-y-1">
          {files.map((f) => (
            <li
              key={f.id}
              className="flex items-center justify-between rounded-md bg-gray-800/60 px-2 py-1.5 text-xs"
            >
              <div className="flex items-center gap-2 overflow-hidden">
                {fileIcon(f.type)}
                <span className="truncate text-gray-300">{f.name}</span>
                <span className="flex-shrink-0 text-gray-600">
                  {formatFileSize(f.size)}
                </span>
              </div>
              <button
                onClick={() => onRemove(f.id)}
                className="ml-2 flex-shrink-0 rounded p-0.5 text-gray-600 hover:text-red-400"
              >
                <X size={12} />
              </button>
            </li>
          ))}
        </ul>
      )}
    </div>
  );
}
