import { useState, useCallback, useEffect, useRef } from "react";
import { Settings, AlertCircle, X, StopCircle, PlusCircle, ImageIcon } from "lucide-react";
import type {
  AzureConfig,
  InputFile,
  InputImage,
  ChatMessage,
  DocumentOutput,
  TokenUsage,
  OutputFormat,
  ThemeOption,
  UseCasePreset,
} from "./types/document";
import AzureSettings from "./components/ConfigPanel/AzureSettings";
import ChatRefinement from "./components/PreviewPanel/ChatRefinement";
import SlideOutline from "./components/PreviewPanel/SlideOutline";
import HTMLSlidePreview from "./components/PreviewPanel/HTMLSlidePreview";
import DownloadButtons from "./components/ExportPanel/DownloadButtons";
import {
  generateDocumentStream,
  smartChatStream,
  parseDocumentJSON,
} from "./services/azureFoundry";
import {
  buildUserPrompt,
  getSystemPrompt,
} from "./services/promptTemplates";
import { captureBackground } from "./services/pptxBuilder";
import { loadKnowledgeBaseFolder } from "./services/fileExtractor";
import { generateClarifyingQuestions } from "./services/azureFoundry";
import type { ClarifyingQuestion } from "./services/azureFoundry";

const STORAGE_KEY = "ai_doc_azure_config";
const SESSION_KEY = "zenspark_session";

interface PersistedSession {
  chatHistory: ChatMessage[];
  generatedDoc: DocumentOutput | null;
  outputFormat: OutputFormat;
  theme: ThemeOption;
  savedAt: string;
}

function loadConfig(): AzureConfig | null {
  try {
    const raw = sessionStorage.getItem(STORAGE_KEY);
    return raw ? JSON.parse(raw) : null;
  } catch {
    return null;
  }
}

function loadSession(): PersistedSession | null {
  try {
    const raw = localStorage.getItem(SESSION_KEY);
    return raw ? JSON.parse(raw) : null;
  } catch {
    return null;
  }
}

function saveSession(data: PersistedSession) {
  try {
    // Omit html from slides before persisting — it's large and re-generated on refine
    const slim: PersistedSession = {
      ...data,
      generatedDoc: data.generatedDoc ? {
        ...data.generatedDoc,
        slides: data.generatedDoc.slides?.map(({ html: _html, ...rest }) => rest),
      } : null,
    };
    localStorage.setItem(SESSION_KEY, JSON.stringify(slim));
  } catch {
    // localStorage full or unavailable — ignore
  }
}

export default function App() {
  // ── Configuration ──────────────────────────────────────────────────────────
  const [azureConfig, setAzureConfig] = useState<AzureConfig | null>(loadConfig);
  const [showSettings, setShowSettings] = useState(false);

  // ── Inputs ─────────────────────────────────────────────────────────────────
  const [inputFiles, setInputFiles] = useState<InputFile[]>([]);
  const [inputImages, setInputImages] = useState<InputImage[]>([]);

  // ── Output settings ────────────────────────────────────────────────────────
  const [useCase, setUseCase] = useState<UseCasePreset>("customer_proposal");
  const [customPrompt, setCustomPrompt] = useState("");
  const [outputFormat, setOutputFormat] = useState<OutputFormat>("pptx");
  const [theme, setTheme] = useState<ThemeOption>("corporate_blue");

  // ── Right panel resize ────────────────────────────────────────────────────
  const [rightWidth, setRightWidth] = useState(520);
  const resizeDrag = useRef<{ startX: number; startW: number } | null>(null);

  useEffect(() => {
    const onMove = (e: MouseEvent) => {
      if (!resizeDrag.current) return;
      const dx = resizeDrag.current.startX - e.clientX;
      const newW = Math.max(360, Math.min(Math.round(window.innerWidth * 0.78), resizeDrag.current.startW + dx));
      setRightWidth(newW);
    };
    const onUp = () => {
      if (!resizeDrag.current) return;
      resizeDrag.current = null;
      document.body.style.cursor = "";
      document.body.style.userSelect = "";
    };
    document.addEventListener("mousemove", onMove);
    document.addEventListener("mouseup", onUp);
    return () => {
      document.removeEventListener("mousemove", onMove);
      document.removeEventListener("mouseup", onUp);
    };
  }, []);

  const handleResizeMouseDown = (e: React.MouseEvent) => {
    resizeDrag.current = { startX: e.clientX, startW: rightWidth };
    document.body.style.cursor = "ew-resize";
    document.body.style.userSelect = "none";
    e.preventDefault();
  };

  // ── Generation state ───────────────────────────────────────────────────────
  const [isGenerating, setIsGenerating] = useState(false);
  const [isRefining, setIsRefining] = useState(false);
  const abortRef = useRef<AbortController | null>(null);

  const handleStop = () => {
    abortRef.current?.abort();
    abortRef.current = null;
  };
  const [streamingText, setStreamingText] = useState("");
  const [generatedDoc, setGeneratedDoc] = useState<DocumentOutput | null>(null);
  const [chatHistory, setChatHistory] = useState<ChatMessage[]>([]);
  const [chatInput, setChatInput] = useState("");
  const [tokenUsage, setTokenUsage] = useState<TokenUsage | null>(null);
  const [error, setError] = useState<string | null>(null);

  // Pre-captured slide background PNGs (populated after generation, index = slide_number - 1)
  const [slideBackgrounds, setSlideBackgrounds] = useState<(string | null)[]>([]);
  const [bgCapture, setBgCapture] = useState<{ done: number; total: number } | null>(null);

  // Clarifying questions flow
  const [clarifyingQuestions, setClarifyingQuestions] = useState<ClarifyingQuestion[]>([]);
  const [clarifyAnswers, setClarifyAnswers] = useState<Record<string, string>>({});
  const [isClarifying, setIsClarifying] = useState(false);
  const pendingGoalRef = useRef<string>("");

  // Restore last session on mount
  useEffect(() => {
    const session = loadSession();
    if (session) {
      setChatHistory(session.chatHistory ?? []);
      setGeneratedDoc(session.generatedDoc ?? null);
      setOutputFormat(session.outputFormat ?? "pptx");
      setTheme(session.theme ?? "corporate_blue");
    }
    if (!azureConfig) setShowSettings(true);
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Persist session whenever chat or doc changes
  useEffect(() => {
    if (chatHistory.length > 0 || generatedDoc) {
      saveSession({ chatHistory, generatedDoc, outputFormat, theme, savedAt: new Date().toISOString() });
    }
  }, [chatHistory, generatedDoc, outputFormat, theme]);

  // Pre-capture slide backgrounds after generation/refinement (so download is instant)
  useEffect(() => {
    const slides = generatedDoc?.slides ?? [];
    if (!slides.some((s) => s.background_html)) {
      setSlideBackgrounds([]);
      return;
    }
    let cancelled = false;
    const runCapture = async () => {
      setSlideBackgrounds([]);
      setBgCapture({ done: 0, total: slides.length });
      const results: (string | null)[] = [];
      for (const slide of slides) {
        if (cancelled) break;
        results.push(slide.background_html ? await captureBackground(slide.background_html) : null);
        if (!cancelled) setBgCapture({ done: results.length, total: slides.length });
      }
      if (!cancelled) {
        setSlideBackgrounds(results);
        setBgCapture(null);
      }
    };
    runCapture();
    return () => { cancelled = true; setBgCapture(null); };
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [generatedDoc]);

  const handleNewSession = () => {
    setChatHistory([]);
    setGeneratedDoc(null);
    setChatInput("");
    setTokenUsage(null);
    setError(null);
    setInputFiles([]);
    setInputImages([]);
    setSlideBackgrounds([]);
    setBgCapture(null);
    setClarifyingQuestions([]);
    setClarifyAnswers({});
    localStorage.removeItem(SESSION_KEY);
  };

  // ── Handlers ───────────────────────────────────────────────────────────────
  const handleAddFiles = useCallback((newFiles: InputFile[]) => {
    setInputFiles((prev) => [...prev, ...newFiles]);
  }, []);

  // ── Folder Knowledge Base ──────────────────────────────────────────────────
  const handleLoadFolder = useCallback(async () => {
    const result = await loadKnowledgeBaseFolder();
    if (!result) return;
    handleAddFiles(result.files);
    const msg = `📂 Knowledge base loaded: "${result.folderName}" — ${result.files.length} files indexed${result.skippedFiles > 0 ? `, ${result.skippedFiles} skipped` : ""}. The LLM will use this as context.`;
    setChatHistory((prev) => [...prev, { role: "assistant" as const, content: msg }]);
  }, [handleAddFiles]);

  const handleRemoveFile = useCallback((id: string) => {
    setInputFiles((prev) => prev.filter((f) => f.id !== id));
  }, []);

  const handleAddImage = useCallback((img: InputImage) => {
    setInputImages((prev) => [...prev, img]);
  }, []);

  const handleRemoveImage = useCallback((id: string) => {
    setInputImages((prev) => prev.filter((i) => i.id !== id));
  }, []);

  const handleSaveConfig = (cfg: AzureConfig) => {
    setAzureConfig(cfg);
    sessionStorage.setItem(STORAGE_KEY, JSON.stringify(cfg));
    setShowSettings(false);
  };

  // Core generation logic — accepts optional clarify answers appended to goal
  const runGenerate = async (goal: string, answers?: Record<string, string>) => {
    if (!azureConfig) {
      setShowSettings(true);
      return;
    }

    setClarifyingQuestions([]);
    setClarifyAnswers({});
    setError(null);
    setIsGenerating(true);
    setStreamingText("");
    setGeneratedDoc(null);

    // Build goal with answers if present
    let fullGoal = goal;
    if (answers && Object.keys(answers).length > 0) {
      const answersText = Object.entries(answers)
        .map(([id, ans]) => {
          const q = clarifyingQuestions.find((cq) => cq.id === id);
          return q ? `- ${q.question}\n  Answer: ${ans}` : `- ${ans}`;
        })
        .join("\n");
      fullGoal = `${goal}\n\nUser clarifications:\n${answersText}`;
    }

    const userGoal = fullGoal.trim();
    setChatHistory((prev) => {
      // replace last user message if it was set by send handler
      const last = prev[prev.length - 1];
      return last?.role === "user" ? prev : [...prev, { role: "user", content: goal }];
    });
    setChatInput("");

    const fileTexts = inputFiles.map((f) => ({
      name: f.name,
      content: f.extractedText,
    }));

    const userPrompt = buildUserPrompt(
      userGoal,
      fileTexts,
      "",
      outputFormat,
      theme,
      inputImages.length
    );

    const abort = new AbortController();
    abortRef.current = abort;
    let accumulated = "";
    try {
      for await (const chunk of generateDocumentStream(
        azureConfig,
        getSystemPrompt("custom"),
        userPrompt,
        inputImages,
        abort.signal
      )) {
        accumulated += chunk;
        setStreamingText(accumulated);
      }

      const doc = parseDocumentJSON(accumulated);
      setGeneratedDoc(doc);
      setChatHistory((prev) => [
        ...prev,
        { role: "assistant", content: `Document ready: "${doc.title}" — ${(doc.slides ?? doc.sections ?? []).length} ${doc.document_type === "pptx" ? "slides" : "sections"} generated.` },
      ]);

      const promptChars = userPrompt.length + fileTexts.reduce((s, f) => s + f.content.length, 0);
      setTokenUsage({
        promptTokens: Math.round(promptChars / 4),
        completionTokens: Math.round(accumulated.length / 4),
        totalTokens: Math.round((promptChars + accumulated.length) / 4),
      });
    } catch (err) {
      if ((err as Error).name !== "AbortError") {
        setError(err instanceof Error ? err.message : String(err));
      }
    } finally {
      abortRef.current = null;
      setIsGenerating(false);
    }
  };

  const handleGenerate = async () => {
    if (!azureConfig) { setShowSettings(true); return; }

    const goal = chatInput.trim();
    if (!goal) return;

    // Save goal for later use after clarify
    pendingGoalRef.current = goal;
    setChatHistory([{ role: "user", content: goal }]);
    setChatInput("");

    // Ask clarifying questions first
    setIsClarifying(true);
    try {
      const abort = new AbortController();
      abortRef.current = abort;
      const questions = await generateClarifyingQuestions(
        azureConfig,
        goal,
        inputFiles.map((f) => f.name),
        abort.signal
      );
      setClarifyingQuestions(questions);
    } catch {
      // If questions fail, just generate directly
      setClarifyingQuestions([]);
      await runGenerate(goal);
    } finally {
      abortRef.current = null;
      setIsClarifying(false);
    }
  };

  const handleSubmitClarify = () => {
    runGenerate(pendingGoalRef.current, clarifyAnswers);
  };

  const handleSkipClarify = () => {
    runGenerate(pendingGoalRef.current);
  };

  const handleClarifyAnswer = (id: string, answer: string) => {
    setClarifyAnswers((prev) => ({ ...prev, [id]: answer }));
  };

  // Dual-mode chat: conversational OR document editing, decided by the LLM
  const handleChatSend = async (currentDoc: DocumentOutput | null = generatedDoc) => {
    if (!azureConfig || !chatInput.trim()) return;

    const userMsg: ChatMessage = { role: "user", content: chatInput };
    const updatedHistory = [...chatHistory, userMsg];
    setChatHistory(updatedHistory);
    setChatInput("");
    setIsRefining(true);
    setError(null);

    const abort = new AbortController();
    abortRef.current = abort;
    let accumulated = "";
    try {
      for await (const chunk of smartChatStream(
        azureConfig,
        currentDoc,
        updatedHistory,
        userMsg.content,
        inputImages,
        inputFiles.map((f) => ({ name: f.name, content: f.extractedText })),
        abort.signal
      )) {
        accumulated += chunk;
      }

      // Detect if the LLM returned a document JSON or a plain conversational reply
      const trimmed = accumulated.trimStart();
      const looksLikeDocJson =
        trimmed.startsWith("{") &&
        (trimmed.includes('"slides"') || trimmed.includes('"document_type"') || trimmed.includes('"title"'));

      if (looksLikeDocJson && currentDoc) {
        try {
          const refined = parseDocumentJSON(accumulated);
          setGeneratedDoc(refined);
          setChatHistory((prev) => [
            ...prev,
            { role: "assistant", content: "Done — document updated." },
          ]);
        } catch {
          // Parsing failed — show raw response as chat
          setChatHistory((prev) => [...prev, { role: "assistant", content: accumulated }]);
        }
      } else {
        setChatHistory((prev) => [...prev, { role: "assistant", content: accumulated }]);
      }
    } catch (err) {
      if ((err as Error).name !== "AbortError") {
        const msg = err instanceof Error ? err.message : String(err);
        setError(msg);
        setChatHistory((prev) => [
          ...prev,
          { role: "assistant", content: `Error: ${msg}` },
        ]);
      }
    } finally {
      abortRef.current = null;
      setIsRefining(false);
    }
  };

  // Routes: free chat (no doc), generate doc, or dual-mode chat (with doc)
  const isDocCreationIntent = (msg: string) => {
    const m = msg.toLowerCase();
    return /\b(create|make|generate|build|write|draft|design|prepare)\b/.test(m) &&
      /\b(presentation|deck|slides?|pptx|document|report|proposal|whitepaper|pitch)\b/.test(m);
  };

  const handleSend = () => {
    if (!chatInput.trim()) return;
    if (!generatedDoc) {
      if (isDocCreationIntent(chatInput)) {
        handleGenerate();
      } else {
        handleChatSend(null); // free conversational chat, no doc context
      }
    } else {
      handleChatSend();
    }
  };

  // ── Layout ─────────────────────────────────────────────────────────────────
  return (
    <div className="flex h-screen flex-col bg-gray-950 text-gray-200 overflow-hidden">
      {/* Header */}
      <header className="flex flex-shrink-0 items-center justify-between border-b border-gray-800 px-5 py-3">
        <div className="flex items-center gap-3">
          <div className="flex h-7 w-7 items-center justify-center rounded-md bg-blue-600 text-sm font-bold text-white">
            AI
          </div>
          <div>
            <h1 className="text-sm font-semibold leading-none">ZenSpark</h1>
            <p className="text-xs text-gray-500">Powered by Azure AI Foundry</p>
          </div>
        </div>
        <div className="flex items-center gap-2">
          {(isGenerating || isRefining) && (
            <button
              onClick={handleStop}
              className="flex items-center gap-1.5 rounded-lg border border-red-700 px-3 py-1.5 text-xs text-red-400 hover:bg-red-900/20"
            >
              <StopCircle size={13} />
              Stop
            </button>
          )}
          {(chatHistory.length > 0 || generatedDoc) && !isGenerating && !isRefining && (
            <button
              onClick={handleNewSession}
              title="Start a new session"
              className="flex items-center gap-1.5 rounded-lg border border-gray-700 px-3 py-1.5 text-xs text-gray-400 hover:border-gray-500 hover:text-gray-200"
            >
              <PlusCircle size={13} />
              New
            </button>
          )}
          <button
            onClick={() => setShowSettings(true)}
            className={`flex items-center gap-1.5 rounded-lg border px-3 py-1.5 text-xs transition-colors ${
              azureConfig
                ? "border-green-800 text-green-400 hover:border-green-600"
                : "border-amber-600 text-amber-400 hover:border-amber-500"
            }`}
          >
            {azureConfig ? (
              <><span className="h-1.5 w-1.5 rounded-full bg-green-400" /><Settings size={13} /> Connected</>
            ) : (
              <><Settings size={13} /> Configure Azure</>
            )}
          </button>
        </div>
      </header>

      {/* Progress bar — visible whenever generating or refining */}
      {(isGenerating || isRefining) && (
        <div className="relative h-1 w-full flex-shrink-0 overflow-hidden bg-gray-800">
          <div className="absolute inset-y-0 w-1/3 animate-[shimmer_1.2s_linear_infinite] bg-gradient-to-r from-transparent via-blue-500 to-transparent" />
        </div>
      )}

      {/* Error banner */}
      {error && (
        <div className="mx-4 mt-3 flex items-start gap-2 rounded-lg border border-red-800 bg-red-950/40 px-3 py-2 text-xs text-red-300">
          <AlertCircle size={14} className="mt-0.5 flex-shrink-0" />
          <span className="flex-1">{error}</span>
          <button onClick={() => setError(null)} className="flex-shrink-0 hover:text-red-200">
            <X size={13} />
          </button>
        </div>
      )}

      {/* Two-panel body */}
      <div className="flex flex-1 overflow-hidden">
        {/* ── LEFT / MAIN: Chat Panel ───────────────────────────── */}
        <main className="flex min-w-0 flex-1 flex-col overflow-hidden">
          <ChatRefinement
            history={chatHistory}
            input={chatInput}
            isGenerating={isGenerating}
            isRefining={isRefining}
            hasDoc={!!generatedDoc}
            files={inputFiles}
            images={inputImages}
            slides={generatedDoc?.slides ?? []}
            clarifyingQuestions={clarifyingQuestions}
            clarifyAnswers={clarifyAnswers}
            isClarifying={isClarifying}
            onInputChange={setChatInput}
            onSend={handleSend}
            onAddFiles={handleAddFiles}
            onRemoveFile={handleRemoveFile}
            onAddImage={handleAddImage}
            onRemoveImage={handleRemoveImage}
            onLoadFolder={handleLoadFolder}
            onClarifyAnswer={handleClarifyAnswer}
            onSkipClarify={handleSkipClarify}
            onSubmitClarify={handleSubmitClarify}
          />
        </main>

        {/* ── Resize handle ─────────────────────────────────────── */}
        <div
          onMouseDown={handleResizeMouseDown}
          className="group flex w-1 flex-shrink-0 cursor-ew-resize items-center justify-center bg-gray-800 hover:bg-blue-600 transition-colors"
          title="Drag to resize"
        />

        {/* ── RIGHT: Slide Preview + Downloads ──────────────────── */}
        <aside
          className="flex flex-shrink-0 flex-col overflow-hidden"
          style={{ width: rightWidth }}
        >
          {/* Header */}
          <div className="flex flex-shrink-0 items-center justify-between border-b border-gray-800 px-4 py-2.5">
            <p className="text-xs font-semibold uppercase tracking-wider text-gray-500">Slide Preview</p>
            <div className="flex items-center gap-2">
              {/* Background capture progress */}
              {bgCapture && (
                <span className="flex items-center gap-1 rounded-full bg-purple-900/40 px-2 py-0.5 text-xs text-purple-300">
                  <ImageIcon size={10} className="animate-pulse" />
                  Rendering {bgCapture.done}/{bgCapture.total}
                </span>
              )}
              {generatedDoc && !bgCapture && (
                <span className="rounded-full bg-blue-900/40 px-2 py-0.5 text-xs text-blue-400">
                  {(generatedDoc.slides ?? generatedDoc.sections ?? []).length}{" "}
                  {generatedDoc.document_type === "pptx" ? "slides" : "sections"}
                </span>
              )}
            </div>
          </div>

          {/* Slide viewer — fills remaining height */}
          <div className="flex min-h-0 flex-1 flex-col px-3 py-3">
            {generatedDoc?.slides?.some((s) => s.html) ? (
              <HTMLSlidePreview slides={generatedDoc.slides ?? []} />
            ) : (
              <div className="flex-1 overflow-y-auto">
                <SlideOutline
                  doc={generatedDoc}
                  isGenerating={isGenerating}
                  streamingText={streamingText}
                />
              </div>
            )}
          </div>

          {/* Download section — pinned at bottom */}
          {generatedDoc && (
            <div className="flex-shrink-0 border-t border-gray-800 px-4 py-3">
              <DownloadButtons
                doc={generatedDoc}
                images={inputImages}
                outputFormat={outputFormat}
                tokenUsage={tokenUsage}
                slideBackgrounds={slideBackgrounds}
                bgCapturing={!!bgCapture}
              />
            </div>
          )}
        </aside>
      </div>

      {/* Settings Modal */}
      {showSettings && (
        <AzureSettings
          initialConfig={azureConfig}
          onSave={handleSaveConfig}
          onClose={() => {
            if (azureConfig) setShowSettings(false);
          }}
        />
      )}
    </div>
  );
}
