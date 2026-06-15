import { useState, useEffect } from "react";
import { Settings, X, CheckCircle, AlertCircle, Loader2, ExternalLink, Copy } from "lucide-react";
import type { AzureConfig } from "../../types/document";
import { testConnection } from "../../services/azureFoundry";

const SESSION_KEY = "ai_doc_azure_config";
const API_VERSIONS = [
  "2024-12-01-preview",
  "2024-08-01-preview",
  "2024-05-01-preview",
  "2024-02-01",
];

const DEFAULT_CONFIG: AzureConfig = {
  endpoint: "",
  apiKey: "",
  deploymentName: "",
  apiVersion: "2024-12-01-preview",
  maxTokens: 4096,
  temperature: 0.7,
  visionDeploymentName: "",
};

interface Props {
  initialConfig: AzureConfig | null;
  onSave: (config: AzureConfig) => void;
  onClose: () => void;
}

const INPUT_CLS = "w-full rounded-lg border border-gray-700 bg-gray-800/60 px-3 py-2 text-sm text-gray-200 placeholder-gray-600 focus:border-blue-500 focus:outline-none focus:ring-1 focus:ring-blue-500";

function friendlyError(raw: string): string {
  if (raw.includes("404") || raw.toLowerCase().includes("deployment") || raw.toLowerCase().includes("does not exist")) {
    return "404 — Deployment not found. The \"Model Deployment Name\" must exactly match the name you gave your deployment in Azure AI Foundry (e.g. my-claude-deployment). It is NOT the model name.";
  }
  if (raw.includes("401") || raw.toLowerCase().includes("unauthorized") || raw.toLowerCase().includes("api key")) {
    return "401 — Invalid API Key. Copy the key from Azure AI Foundry → your project → Settings → API Keys.";
  }
  if (raw.includes("403")) {
    return "403 — Access denied. Check that your API key has permission to access this resource.";
  }
  if (raw.toLowerCase().includes("failed to fetch") || raw.toLowerCase().includes("networkerror") || raw.toLowerCase().includes("cors")) {
    return "Network error / CORS — Check your Endpoint URL. It must start with https:// and end with .openai.azure.com/";
  }
  return raw;
}

export default function AzureSettings({ initialConfig, onSave, onClose }: Props) {
  const [form, setForm] = useState<AzureConfig>(initialConfig ?? DEFAULT_CONFIG);
  const [testStatus, setTestStatus] = useState<"idle" | "testing" | "ok" | "error">("idle");
  const [testMessage, setTestMessage] = useState("");
  const [copied, setCopied] = useState(false);

  useEffect(() => {
    if (!initialConfig) {
      const saved = sessionStorage.getItem(SESSION_KEY);
      if (saved) {
        try { setForm(JSON.parse(saved) as AzureConfig); } catch { /* ignore */ }
      }
    }
  }, [initialConfig]);

  const set = (field: keyof AzureConfig, value: string | number) => {
    setForm((f) => ({ ...f, [field]: value }));
    setTestStatus("idle");
  };

  // The exact URL that will be called
  const previewUrl = form.endpoint && form.deploymentName
    ? `${form.endpoint.replace(/\/$/, "")}/openai/deployments/${form.deploymentName}/chat/completions?api-version=${form.apiVersion}`
    : "";

  const copyUrl = () => {
    if (previewUrl) {
      navigator.clipboard.writeText(previewUrl);
      setCopied(true);
      setTimeout(() => setCopied(false), 1500);
    }
  };

  const handleTest = async () => {
    if (!form.endpoint || !form.apiKey || !form.deploymentName) {
      setTestStatus("error");
      setTestMessage("Please fill in Endpoint, API Key, and Deployment Name.");
      return;
    }
    setTestStatus("testing");
    setTestMessage("");
    try {
      await testConnection(form);
      setTestStatus("ok");
      setTestMessage("Connected!");
    } catch (e) {
      setTestStatus("error");
      setTestMessage(friendlyError(e instanceof Error ? e.message : String(e)));
    }
  };

  const handleSave = () => {
    sessionStorage.setItem(SESSION_KEY, JSON.stringify(form));
    onSave(form);
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/70 backdrop-blur-sm">
      <div className="w-full max-w-lg rounded-xl border border-gray-700 bg-gray-900 shadow-2xl">
        {/* Header */}
        <div className="flex items-center justify-between border-b border-gray-700 px-6 py-4">
          <div className="flex items-center gap-2">
            <Settings size={18} className="text-blue-400" />
            <h2 className="font-semibold text-white">Azure AI Foundry Settings</h2>
          </div>
          <div className="flex items-center gap-3">
            <a
              href="https://ai.azure.com"
              target="_blank"
              rel="noreferrer"
              className="flex items-center gap-1 text-xs text-blue-400 hover:text-blue-300"
            >
              Open Azure AI Foundry <ExternalLink size={11} />
            </a>
            <button onClick={onClose} className="rounded p-1 text-gray-400 hover:text-white">
              <X size={18} />
            </button>
          </div>
        </div>

        {/* Form */}
        <div className="max-h-[70vh] overflow-y-auto">
          <div className="space-y-4 px-6 py-5">
            <Field label="Endpoint URL" hint='Azure AI Foundry → your project → Overview → "Endpoint"'>
              <input
                type="url"
                value={form.endpoint}
                onChange={(e) => set("endpoint", e.target.value)}
                placeholder="https://your-resource.openai.azure.com/"
                className={INPUT_CLS}
              />
            </Field>

            <Field label="API Key" hint='Azure AI Foundry → your project → Settings → "Keys and Endpoint"'>
              <input
                type="password"
                value={form.apiKey}
                onChange={(e) => set("apiKey", e.target.value)}
                placeholder="••••••••••••••••••••••••••••••••"
                className={INPUT_CLS}
              />
            </Field>

            <div className="grid grid-cols-2 gap-4">
              <Field
                label="Deployment Name"
                hint='The name YOU gave the deployment — not the model name'
              >
                <input
                  type="text"
                  value={form.deploymentName}
                  onChange={(e) => set("deploymentName", e.target.value)}
                  placeholder="e.g. my-claude-deployment"
                  className={INPUT_CLS}
                />
              </Field>

              <Field label="API Version">
                <select
                  value={form.apiVersion}
                  onChange={(e) => set("apiVersion", e.target.value)}
                  className={INPUT_CLS}
                >
                  {API_VERSIONS.map((v) => (
                    <option key={v} value={v}>{v}</option>
                  ))}
                </select>
              </Field>
            </div>

            {/* Live URL preview */}
            {previewUrl && (
              <div className="rounded-lg border border-gray-800 bg-gray-950 px-3 py-2">
                <div className="mb-1 flex items-center justify-between">
                  <span className="text-xs font-medium text-gray-500">API URL that will be called</span>
                  <button onClick={copyUrl} className="flex items-center gap-1 text-xs text-gray-600 hover:text-gray-400">
                    <Copy size={11} />
                    {copied ? "Copied!" : "Copy"}
                  </button>
                </div>
                <p className="break-all font-mono text-xs text-gray-400">{previewUrl}</p>
              </div>
            )}

            <div className="grid grid-cols-2 gap-4">
              <Field label={`Max Tokens: ${form.maxTokens}`}>
                <input
                  type="range" min={512} max={8192} step={256}
                  value={form.maxTokens}
                  onChange={(e) => set("maxTokens", Number(e.target.value))}
                  className="w-full accent-blue-500"
                />
              </Field>
              <Field label={`Temperature: ${form.temperature.toFixed(1)}`}>
                <input
                  type="range" min={0} max={1} step={0.1}
                  value={form.temperature}
                  onChange={(e) => set("temperature", parseFloat(e.target.value))}
                  className="w-full accent-blue-500"
                />
              </Field>
            </div>

            {/* Test result banner */}
            <div className={`rounded-lg border px-4 py-3 text-sm ${
              testStatus === "idle"    ? "border-gray-800 bg-gray-800/20 text-gray-600" :
              testStatus === "testing" ? "border-blue-800 bg-blue-950/40 text-blue-300" :
              testStatus === "ok"      ? "border-green-700 bg-green-950/40 text-green-300" :
                                         "border-red-700 bg-red-950/40 text-red-300"
            }`}>
              {testStatus === "idle" && (
                <p>Click <strong className="text-gray-500">Test Connection</strong> to verify before saving.</p>
              )}
              {testStatus === "testing" && (
                <div className="flex items-center gap-2"><Loader2 size={14} className="animate-spin" /> Connecting to Azure AI Foundry…</div>
              )}
              {testStatus === "ok" && (
                <div className="flex items-center gap-2"><CheckCircle size={14} className="text-green-400" /> <strong>Connected!</strong> Model responded. Ready to save.</div>
              )}
              {testStatus === "error" && (
                <div className="flex items-start gap-2">
                  <AlertCircle size={14} className="mt-0.5 flex-shrink-0 text-red-400" />
                  <span>{testMessage}</span>
                </div>
              )}
            </div>
          </div>
        </div>

        {/* Footer */}
        <div className="flex items-center justify-between border-t border-gray-700 px-6 py-4">
          <button
            onClick={handleTest}
            disabled={testStatus === "testing"}
            className={`flex items-center gap-2 rounded-lg border px-4 py-2 text-sm font-medium transition-colors disabled:opacity-50 ${
              testStatus === "ok"
                ? "border-green-700 text-green-400"
                : "border-gray-600 text-gray-300 hover:border-blue-500 hover:text-blue-400"
            }`}
          >
            {testStatus === "testing" ? <Loader2 size={14} className="animate-spin" /> :
             testStatus === "ok"      ? <CheckCircle size={14} /> : null}
            {testStatus === "testing" ? "Testing…" : testStatus === "ok" ? "Tested ✓" : "Test Connection"}
          </button>

          <div className="flex gap-2">
            <button onClick={onClose} className="rounded-lg px-4 py-2 text-sm text-gray-400 hover:text-white">
              Cancel
            </button>
            <button
              onClick={handleSave}
              disabled={testStatus !== "ok"}
              title={testStatus !== "ok" ? "Test the connection first" : undefined}
              className="flex items-center gap-2 rounded-lg bg-blue-600 px-4 py-2 text-sm font-medium text-white hover:bg-blue-500 disabled:cursor-not-allowed disabled:opacity-40"
            >
              <CheckCircle size={14} />
              Save & Connect
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

function Field({ label, hint, children }: { label: string; hint?: string; children: React.ReactNode }) {
  return (
    <div>
      <label className="mb-1 block text-xs font-medium text-gray-400">{label}</label>
      {children}
      {hint && <p className="mt-1 text-xs text-gray-600">{hint}</p>}
    </div>
  );
}
