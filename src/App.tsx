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
import FileUpload from "./components/InputPanel/FileUpload";
import ScreenshotPaste from "./components/InputPanel/ScreenshotPaste";
import TextInput from "./components/InputPanel/TextInput";
import SlideOutline from "./components/PreviewPanel/SlideOutline";
import ChatRefinement from "./components/PreviewPanel/ChatRefinement";
import DownloadButtons from "./components/ExportPanel/DownloadButtons";
import {
  generateDocumentStream,
  refineDocumentStream,
  parseDocumentJSON,
} from "./services/azureFoundry";
import {
  buildUserPrompt,
  getSystemPrompt,
  DEFAULT_SLIDE_COUNT,
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
  const [inputText, setInputText] = useState("");

  // ── Output settings ────────────────────────────────────────────────────────
  const [useCase, setUseCase] = useState<UseCasePreset>("customer_proposal");
  const [customPrompt, setCustomPrompt] = useState("");
  const [outputFormat, setOutputFormat] = useState<OutputFormat>("pptx");
  const [theme, setTheme] = useState<ThemeOption>("corporate_blue");
  const [slideCount, setSlideCount] = useState(DEFAULT_SLIDE_COUNT);

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

  const handleCaptionChange = useCallback((id: string, caption: string) => {
    setInputImages((prev) =>
      prev.map((img) => (img.id === id ? { ...img, caption } : img))
    );
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
    setChatHistory([]);

    const fileTexts = inputFiles.map((f) => ({
      name: f.name,
      content: f.text,
    }));

    const goalText =
      useCase === "custom" && customPrompt.trim()
        ? customPrompt
        : undefined;

    const userPrompt = buildUserPrompt(
      goalText,
      fileTexts,
      inputText.slice(0, 20_000),
      slideCount,
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

      // Approximate token usage (Azure may not stream usage counts)
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
        chatInput,
        inputImages
      )) {
        accumulated += chunk;
      }

      const refined = parseDocumentJSON(accumulated);
      setGeneratedDoc(refined);
      setChatHistory((prev) => [
        ...prev,
        { role: "assistant", content: "Document updated successfully." },
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
            <h1 className="text-sm font-semibold leading-none">AI Doc Generator</h1>
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

      {/* Three-panel body */}
      <div className="flex flex-1 overflow-hidden">
        {/* ── LEFT: Input Panel ────────────────────────────────────────────── */}
        <aside className="flex w-72 flex-shrink-0 flex-col gap-4 overflow-y-auto border-r border-gray-800 p-4">
          <Section title="Upload Documents">
            <FileUpload
              files={inputFiles}
              onAdd={handleAddFiles}
              onRemove={handleRemoveFile}
            />
          </Section>

          <Section title="Screenshots / Images">
            <ScreenshotPaste
              images={inputImages}
              onAdd={handleAddImage}
              onRemove={handleRemoveImage}
              onCaptionChange={handleCaptionChange}
            />
          </Section>

          <Section title="Paste Text / Notes">
            <TextInput
              value={inputText}
              onChange={setInputText}
            />
          </Section>
        </aside>

        {/* ── CENTER: Preview Panel ──────────────────────────────────────── */}
        <main className="flex flex-1 flex-col gap-4 overflow-hidden p-4">
          <div className="flex-1 overflow-hidden">
            <SlideOutline
              doc={generatedDoc}
              isGenerating={isGenerating}
              streamingText={streamingText}
            />
          </div>

          <div className="h-56 flex-shrink-0">
            <ChatRefinement
              history={chatHistory}
              input={chatInput}
              isRefining={isRefining}
              disabled={!generatedDoc}
              onInputChange={setChatInput}
              onSend={handleChatSend}
            />
          </div>
        </main>

        {/* ── RIGHT: Export Panel ───────────────────────────────────────── */}
        <aside className="flex w-64 flex-shrink-0 flex-col overflow-y-auto border-l border-gray-800 p-4">
          <DownloadButtons
            doc={generatedDoc}
            images={inputImages}
            outputFormat={outputFormat}
            theme={theme}
            slideCount={slideCount}
            useCase={useCase}
            tokenUsage={tokenUsage}
            onOutputFormatChange={setOutputFormat}
            onThemeChange={setTheme}
            onSlideCountChange={setSlideCount}
            onUseCaseChange={setUseCase}
            onGenerate={handleGenerate}
            isGenerating={isGenerating}
            customPrompt={customPrompt}
            onCustomPromptChange={setCustomPrompt}
          />
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
