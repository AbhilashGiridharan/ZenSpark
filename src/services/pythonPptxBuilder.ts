/**
 * pythonPptxBuilder.ts
 *
 * Sends the DocumentOutput JSON to the local FastAPI backend (pptx_server.py).
 * The backend asks the LLM to write a python-pptx script, executes it in a
 * sandboxed subprocess, and streams the resulting .pptx file back.
 *
 * Flow:
 *   DocumentOutput JSON ──► POST /generate-pptx ──► LLM writes code
 *   ──► subprocess executes in /tmp ──► .pptx streamed back ──► browser download
 */

import type { AzureConfig, DocumentOutput } from "../types/document";

const SERVER_URL =
  (import.meta.env.VITE_PPTX_SERVER_URL as string | undefined) ??
  "http://localhost:8765";

export interface PythonPptxProgress {
  stage: "calling-llm" | "executing" | "downloading" | "done" | "error";
  message: string;
}

// ─── Health check ─────────────────────────────────────────────────────────────
export async function checkPythonServerHealth(): Promise<boolean> {
  try {
    const res = await fetch(`${SERVER_URL}/health`, { signal: AbortSignal.timeout(3000) });
    return res.ok;
  } catch {
    return false;
  }
}

// ─── Preview: get the generated Python code without executing it ──────────────
export async function previewGeneratedCode(
  doc: DocumentOutput,
  config: AzureConfig,
): Promise<{ code: string; lines: number; safe: boolean; safetyNote: string }> {
  const res = await fetch(`${SERVER_URL}/preview-code`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      documentJson: doc,
      azureConfig: configPayload(config),
      filename: doc.title || "presentation",
    }),
  });
  if (!res.ok) {
    const err = await res.text();
    throw new Error(`Preview failed (${res.status}): ${err}`);
  }
  return res.json();
}

// ─── Main: generate + execute + download ─────────────────────────────────────
export async function buildAndDownloadPythonPptx(
  doc: DocumentOutput,
  config: AzureConfig,
  onProgress?: (p: PythonPptxProgress) => void,
): Promise<void> {
  const report = (stage: PythonPptxProgress["stage"], message: string) =>
    onProgress?.({ stage, message });

  // 1. Check server is up
  const alive = await checkPythonServerHealth();
  if (!alive) {
    throw new Error(
      `Python PPTX server is not running.\n\n` +
      `Start it with:\n  cd ai-doc-generator && python pptx_server.py\n\n` +
      `Requires: pip install fastapi uvicorn python-pptx openai`,
    );
  }

  // 2. Call backend — parallel LLM per slide (or direct render if pptx_elements present)
  const slideCount = doc.slides?.length ?? 0;
  const hasElements = (doc.slides ?? []).every((s) => s.pptx_elements && s.pptx_elements.length > 0);
  report(
    "calling-llm",
    hasElements
      ? "Rendering presentation from design elements…"
      : `Generating code for ${slideCount} slides in parallel…`,
  );

  const res = await fetch(`${SERVER_URL}/generate-pptx`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      documentJson: doc,
      azureConfig: configPayload(config),
      filename: doc.title || "presentation",
    }),
    signal: AbortSignal.timeout(210_000), // 3.5 min — covers LLM fallback
  });

  if (!res.ok) {
    // Parse FastAPI error detail
    let detail = `HTTP ${res.status}`;
    try {
      const json = await res.json() as { detail?: string };
      detail = json.detail ?? detail;
    } catch {
      detail = await res.text();
    }
    report("error", detail);
    throw new Error(detail);
  }

  // 3. Download the streamed .pptx binary
  report("downloading", "Downloading generated file…");

  const blob = await res.blob();
  if (blob.size === 0) throw new Error("Server returned an empty file.");

  // 4. Trigger browser download
  const safeTitle = (doc.title || "presentation")
    .replace(/[^a-zA-Z0-9 _-]/g, "")
    .trim() || "presentation";

  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `${safeTitle}_python.pptx`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);

  report("done", "Download complete");
}

// ─── Shape the config to match the server's Pydantic model ───────────────────
function configPayload(config: AzureConfig) {
  return {
    endpoint: config.endpoint,
    apiKey: config.apiKey,
    deploymentName: config.deploymentName,
    apiVersion: config.apiVersion ?? null,
    maxTokens: config.maxTokens ?? 4096,
    temperature: config.temperature ?? 0.3,
  };
}
