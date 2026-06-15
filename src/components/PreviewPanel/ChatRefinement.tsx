import { useRef, useEffect } from "react";
import { Send, Loader2, MessageSquare } from "lucide-react";
import type { ChatMessage } from "../../types/document";

interface Props {
  history: ChatMessage[];
  input: string;
  isRefining: boolean;
  disabled: boolean;
  onInputChange: (v: string) => void;
  onSend: () => void;
}

export default function ChatRefinement({
  history,
  input,
  isRefining,
  disabled,
  onInputChange,
  onSend,
}: Props) {
  const bottomRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    bottomRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [history]);

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === "Enter" && !e.shiftKey && !isRefining && !disabled && input.trim()) {
      e.preventDefault();
      onSend();
    }
  };

  return (
    <div className="flex h-full flex-col">
      <div className="mb-2 flex items-center gap-2">
        <MessageSquare size={13} className="text-blue-400" />
        <span className="text-xs font-medium text-gray-400">
          Refine with Chat
        </span>
      </div>

      {/* Message history */}
      <div className="flex-1 overflow-y-auto rounded-lg border border-gray-800 bg-gray-900/30 p-2">
        {history.length === 0 ? (
          <div className="flex h-full flex-col items-center justify-center gap-1 text-center">
            <p className="text-xs text-gray-600">No messages yet</p>
            <p className="text-xs text-gray-700">
              Generate a document first, then ask for changes
            </p>
          </div>
        ) : (
          <div className="space-y-2">
            {history.map((msg, i) => (
              <div
                key={i}
                className={`flex ${msg.role === "user" ? "justify-end" : "justify-start"}`}
              >
                <div
                  className={`max-w-[85%] rounded-lg px-3 py-2 text-xs leading-relaxed ${
                    msg.role === "user"
                      ? "bg-blue-600 text-white"
                      : "bg-gray-800 text-gray-300"
                  }`}
                >
                  {msg.content}
                </div>
              </div>
            ))}
            {isRefining && (
              <div className="flex justify-start">
                <div className="flex items-center gap-1.5 rounded-lg bg-gray-800 px-3 py-2">
                  <Loader2 size={11} className="animate-spin text-blue-400" />
                  <span className="text-xs text-gray-400">Refining…</span>
                </div>
              </div>
            )}
            <div ref={bottomRef} />
          </div>
        )}
      </div>

      {/* Input row */}
      <div className="mt-2 flex gap-2">
        <textarea
          value={input}
          onChange={(e) => onInputChange(e.target.value)}
          onKeyDown={handleKeyDown}
          placeholder={disabled ? "Generate a document first…" : "e.g. Make slide 3 more concise, add a ROI slide…"}
          disabled={disabled || isRefining}
          rows={2}
          className="flex-1 resize-none rounded-lg border border-gray-700 bg-gray-800/40 px-3 py-2 text-xs text-gray-200 placeholder-gray-600 focus:border-blue-600 focus:outline-none disabled:opacity-40"
        />
        <button
          onClick={onSend}
          disabled={disabled || isRefining || !input.trim()}
          className="flex-shrink-0 self-end rounded-lg bg-blue-600 p-2.5 text-white hover:bg-blue-500 disabled:opacity-40"
        >
          {isRefining ? (
            <Loader2 size={14} className="animate-spin" />
          ) : (
            <Send size={14} />
          )}
        </button>
      </div>
      <p className="mt-1 text-center text-xs text-gray-700">Enter to send · Shift+Enter for new line</p>
    </div>
  );
}
