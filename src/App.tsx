import { useState, useCallback, useEffect } from "react";
import { Settings, AlertCircle, X } from "lucide-react";
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
import DownloadButtons from "./components/ExportPanel/DownloadButtons";
import {
  generateDocumentStream,
  refineDocumentStream,
  parseDocumentJSON,
} from "./services/azureFoundry";
import {
  buildUserPrompt,
  getSystemPrompt,
} from "./services/promptTemplates";

const STORAGE_KEY = "ai_doc_azure_config";

function loadConfig(): AzureConfig | null {
  try {
    const raw = sessionStorage.getItem(STORAGE_KEY);
    return raw ? JSON.parse(raw) : null;
  } catch {
    return null;
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

  // ── Generation state ───────────────────────────────────────────────────────
  const [isGenerating, setIsGenerating] = useState(false);
  const [isRefining, setIsRefining] = useState(false);
  const [streamingText, setStreamingText] = useState("");
  const [generatedDoc, setGeneratedDoc] = useState<DocumentOutput | null>(null);
  const [chatHistory, setChatHistory] = useState<ChatMessage[]>([]);
  const [chatInput, setChatInput] = useState("");
  const [tokenUsage, setTokenUsage] = useState<TokenUsage | null>(null);
  const [error, setError] = useState<string | null>(null);

  // Show settings modal on first load if no config saved
  useEffect(() => {
    if (!azureConfig) setShowSettings(true);
  }, []);

  // ── Handlers ───────────────────────────────────────────────────────────────
  const handleAddFiles = useCallback((newFiles: InputFile[]) => {
    setInputFiles((prev) => [...prev, ...newFiles]);
  }, []);

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

  const handleGenerate = async () => {
    if (!azureConfig) {
      setShowSettings(true);
      return;
    }

    setError(null);
    setIsGenerating(true);
    setStreamingText("");
    setGeneratedDoc(null);

    // Show the user's request in chat history
    const userGoal = chatInput.trim();
    setChatHistory([{ role: "user", content: userGoal }]);
    setChatInput("");

    const fileTexts = inputFiles.map((f) => ({
      name: f.name,
      content: f.text,
    }));

    const userPrompt = buildUserPrompt(
      userGoal,
      fileTexts,
      "",
      outputFormat,
      theme,
      inputImages.length
    );

    let accumulated = "";
    try {
      for await (const chunk of generateDocumentStream(
        azureConfig,
        getSystemPrompt(useCase),
        userPrompt,
        inputImages
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
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setIsGenerating(false);
    }
  };

  const handleChatSend = async () => {
    if (!azureConfig || !generatedDoc || !chatInput.trim()) return;

    const userMsg: ChatMessage = { role: "user", content: chatInput };
    const updatedHistory = [...chatHistory, userMsg];
    setChatHistory(updatedHistory);
    setChatInput("");
    setIsRefining(true);
    setError(null);

    let accumulated = "";
    try {
      for await (const chunk of refineDocumentStream(
        azureConfig,
        generatedDoc,
        updatedHistory,
        userMsg.content,
        inputImages
      )) {
        accumulated += chunk;
      }

      const refined = parseDocumentJSON(accumulated);
      setGeneratedDoc(refined);
      setChatHistory((prev) => [
        ...prev,
        { role: "assistant", content: "Done — document updated." },
      ]);
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      setError(msg);
      setChatHistory((prev) => [
        ...prev,
        { role: "assistant", content: `Error: ${msg}` },
      ]);
    } finally {
      setIsRefining(false);
    }
  };

  // Routes to generate (first time) or refine (subsequent)
  const handleSend = () => {
    if (!chatInput.trim()) return;
    if (!generatedDoc) {
      handleGenerate();
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
        <button
          onClick={() => setShowSettings(true)}
          className={`flex items-center gap-1.5 rounded-lg border px-3 py-1.5 text-xs transition-colors ${
            azureConfig
              ? "border-gray-700 text-gray-400 hover:border-gray-500"
              : "border-amber-600 text-amber-400 hover:border-amber-500"
          }`}
        >
          <Settings size={13} />
          {azureConfig ? "Settings" : "Configure Azure"}
        </button>
      </header>

      {/* Progress bar — visible whenever generating or refining */}
      {(isGenerating || isRefining) && (
        <div className="h-0.5 w-full flex-shrink-0 overflow-hidden bg-gray-800">
          <div className="h-full animate-[shimmer_1.5s_ease-in-out_infinite] bg-gradient-to-r from-transparent via-blue-500 to-transparent" />
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
        <main className="flex flex-1 flex-col overflow-hidden">
          <ChatRefinement
            history={chatHistory}
            input={chatInput}
            isGenerating={isGenerating}
            isRefining={isRefining}
            hasDoc={!!generatedDoc}
            files={inputFiles}
            images={inputImages}
            onInputChange={setChatInput}
            onSend={handleSend}
            onAddFiles={handleAddFiles}
            onRemoveFile={handleRemoveFile}
            onAddImage={handleAddImage}
            onRemoveImage={handleRemoveImage}
          />
        </main>

        {/* ── RIGHT: Slide Preview + Export ────────────────────── */}
        <aside className="flex w-72 flex-shrink-0 flex-col overflow-hidden border-l border-gray-800">
          {/* Slide outline — top 55% */}
          <div className="flex flex-col overflow-hidden border-b border-gray-800 px-3 py-3" style={{ flex: "0 0 55%" }}>
            <p className="mb-2 flex-shrink-0 text-xs font-semibold uppercase tracking-wider text-gray-500">
              Preview
            </p>
            <div className="flex-1 overflow-y-auto">
              <SlideOutline
                doc={generatedDoc}
                isGenerating={isGenerating}
                streamingText={streamingText}
              />
            </div>
          </div>

          {/* Export settings — bottom 45% */}
          <div className="flex-1 overflow-y-auto px-3 py-3">
            <DownloadButtons
              doc={generatedDoc}
              images={inputImages}
              outputFormat={outputFormat}
              theme={theme}
              useCase={useCase}
              tokenUsage={tokenUsage}
              onOutputFormatChange={setOutputFormat}
              onThemeChange={setTheme}
              onUseCaseChange={setUseCase}
              onGenerate={handleGenerate}
              isGenerating={isGenerating}
              customPrompt={customPrompt}
              onCustomPromptChange={setCustomPrompt}
            />
          </div>
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

function Section({ title, children }: { title: string; children: React.ReactNode }) {
  return (
    <div>
      <p className="mb-2 text-xs font-semibold uppercase tracking-wider text-gray-500">
        {title}
      </p>
      {children}
    </div>
  );
}
