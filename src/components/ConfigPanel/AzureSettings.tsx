import { useState, useEffect } from "react";
import { Settings, X, CheckCircle, AlertCircle, Loader2 } from "lucide-react";
import type { AzureConfig } from "../../types/document";
import { testConnection } from "../../services/azureFoundry";

const SESSION_KEY = "ai_doc_azure_config";
const API_VERSIONS = [
  "2024-12-01-preview",
  "2024-08-01-preview",
  "2024-05-01-preview",
  "2024-02-01",
];

interface Props {
  config: AzureConfig | null;
  onSave: (config: AzureConfig) => void;
  onClose: () => void;
}

const DEFAULT_CONFIG: AzureConfig = {
  endpoint: "",
  apiKey: "",
  deploymentName: "claude-3-5-sonnet",
  apiVersion: "2024-12-01-preview",
  maxTokens: 4096,
  temperature: 0.7,
  visionDeploymentName: "",
};

export default function AzureSettings({ config, onSave, onClose }: Props) {
  const [form, setForm] = useState<AzureConfig>(config ?? DEFAULT_CONFIG);
  const [testStatus, setTestStatus] = useState<"idle" | "testing" | "ok" | "error">("idle");
  const [testMessage, setTestMessage] = useState("");

  // Load from sessionStorage on mount
  useEffect(() => {
    if (!config) {
      const saved = sessionStorage.getItem(SESSION_KEY);
      if (saved) {
        try {
          setForm(JSON.parse(saved) as AzureConfig);
        } catch { /* ignore */ }
      }
    }
  }, [config]);

  const set = (field: keyof AzureConfig, value: string | number) => {
    setForm((f) => ({ ...f, [field]: value }));
    setTestStatus("idle");
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
      setTestMessage("Connection successful!");
    } catch (e) {
      setTestStatus("error");
      setTestMessage(e instanceof Error ? e.message : "Connection failed.");
    }
  };

  const handleSave = () => {
    // Store in sessionStorage (cleared on tab close)
    sessionStorage.setItem(SESSION_KEY, JSON.stringify(form));
    onSave(form);
    onClose();
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
          <button onClick={onClose} className="rounded p-1 text-gray-400 hover:text-white">
            <X size={18} />
          </button>
        </div>

        {/* Form */}
        <div className="space-y-4 px-6 py-5">
          <Field label="Endpoint URL" hint="https://your-resource.openai.azure.com/">
            <input
              type="url"
              value={form.endpoint}
              onChange={(e) => set("endpoint", e.target.value)}
              placeholder="https://your-resource.openai.azure.com/"
              className="input"
            />
          </Field>

          <Field label="API Key" hint="Stored in sessionStorage only — cleared on tab close">
            <input
              type="password"
              value={form.apiKey}
              onChange={(e) => set("apiKey", e.target.value)}
              placeholder="••••••••••••••••"
              className="input"
            />
          </Field>

          <div className="grid grid-cols-2 gap-4">
            <Field label="Model Deployment Name">
              <input
                type="text"
                value={form.deploymentName}
                onChange={(e) => set("deploymentName", e.target.value)}
                placeholder="claude-3-5-sonnet"
                className="input"
              />
            </Field>

            <Field label="API Version">
              <select
                value={form.apiVersion}
                onChange={(e) => set("apiVersion", e.target.value)}
                className="input"
              >
                {API_VERSIONS.map((v) => (
                  <option key={v} value={v}>{v}</option>
                ))}
              </select>
            </Field>
          </div>

          <Field label="Vision Deployment (optional)" hint="For image/screenshot analysis; leave blank to use main model">
            <input
              type="text"
              value={form.visionDeploymentName}
              onChange={(e) => set("visionDeploymentName", e.target.value)}
              placeholder="gpt-4o (leave blank to use main model)"
              className="input"
            />
          </Field>

          <div className="grid grid-cols-2 gap-4">
            <Field label={`Max Tokens: ${form.maxTokens}`}>
              <input
                type="range"
                min={512} max={8192} step={256}
                value={form.maxTokens}
                onChange={(e) => set("maxTokens", Number(e.target.value))}
                className="w-full accent-blue-500"
              />
            </Field>
            <Field label={`Temperature: ${form.temperature.toFixed(1)}`}>
              <input
                type="range"
                min={0} max={1} step={0.1}
                value={form.temperature}
                onChange={(e) => set("temperature", parseFloat(e.target.value))}
                className="w-full accent-blue-500"
              />
            </Field>
          </div>

          {/* Test connection */}
          {testMessage && (
            <div className={`flex items-center gap-2 rounded-lg px-3 py-2 text-sm ${testStatus === "ok" ? "bg-green-900/40 text-green-300" : "bg-red-900/40 text-red-300"}`}>
              {testStatus === "ok" ? <CheckCircle size={14} /> : <AlertCircle size={14} />}
              {testMessage}
            </div>
          )}
        </div>

        {/* Footer */}
        <div className="flex items-center justify-between border-t border-gray-700 px-6 py-4">
          <button
            onClick={handleTest}
            disabled={testStatus === "testing"}
            className="flex items-center gap-2 rounded-lg border border-gray-600 px-4 py-2 text-sm text-gray-300 hover:border-blue-500 hover:text-blue-400 disabled:opacity-50"
          >
            {testStatus === "testing" ? <Loader2 size={14} className="animate-spin" /> : null}
            Test Connection
          </button>

          <div className="flex gap-2">
            <button onClick={onClose} className="rounded-lg px-4 py-2 text-sm text-gray-400 hover:text-white">
              Cancel
            </button>
            <button
              onClick={handleSave}
              className="rounded-lg bg-blue-600 px-4 py-2 text-sm font-medium text-white hover:bg-blue-500"
            >
              Save Settings
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
