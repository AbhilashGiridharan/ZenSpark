import OpenAI from "openai";
import type { AzureConfig, DocumentOutput, ChatMessage, InputImage } from "../types/document";

// ─── Endpoint classification (mirrors foundry_client.py) ─────────────────────
function isAzureOpenAI(endpoint: string): boolean {
  return endpoint.toLowerCase().includes(".openai.azure.com") ||
         endpoint.toLowerCase().includes("cognitiveservices.azure.com");
}

function isServicesEndpoint(endpoint: string): boolean {
  return endpoint.toLowerCase().includes("services.ai.azure.com");
}

function isClaudeModel(model: string): boolean {
  return model.toLowerCase().startsWith("claude-");
}

function stripToBase(endpoint: string): string {
  let base = endpoint.replace(/\?.*$/, "").replace(/\/$/, "");
  for (const suffix of [
    "/anthropic/v1/messages",
    "/models/chat/completions",
    "/openai/v1/chat/completions",
    "/openai/v1",
    "/openai/deployments",
    "/v1",
  ]) {
    if (base.endsWith(suffix)) {
      base = base.slice(0, -suffix.length).replace(/\/$/, "");
    }
  }
  return base;
}

// ─── Client factory (OpenAI-compatible endpoints only) ───────────────────────
export function createAzureClient(config: AzureConfig): OpenAI {
  const base = stripToBase(config.endpoint);
  return new OpenAI({
    apiKey: config.apiKey,
    baseURL: isAzureOpenAI(config.endpoint)
      ? `${base}/openai/deployments/${config.deploymentName}`
      : `${base}/v1`,
    ...(config.apiVersion ? { defaultQuery: { "api-version": config.apiVersion } } : {}),
    defaultHeaders: { "api-key": config.apiKey },
    dangerouslyAllowBrowser: true,
  });
}

// ─── Connection test ──────────────────────────────────────────────────────────
export async function testConnection(config: AzureConfig): Promise<void> {
  // For Claude on services.ai.azure.com use the Anthropic Messages API directly
  if (isServicesEndpoint(config.endpoint) && isClaudeModel(config.deploymentName)) {
    const base = stripToBase(config.endpoint);
    const url = `${base}/anthropic/v1/messages`;
    const resp = await fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": config.apiKey,
        "anthropic-version": "2023-06-01",
      },
      body: JSON.stringify({
        model: config.deploymentName,
        max_tokens: 10,
        messages: [{ role: "user", content: 'Reply with the single word "ok".' }],
      }),
    });
    const data = await resp.json() as Record<string, unknown>;
    if (!resp.ok) {
      const err = (data.error as Record<string, string> | undefined);
      throw new Error(`HTTP ${resp.status} — ${err?.message ?? JSON.stringify(data)}`);
    }
    return;
  }

  // OpenAI-compatible path
  const client = createAzureClient(config);
  const response = await client.chat.completions.create({
    model: config.deploymentName,
    messages: [{ role: "user", content: 'Reply with the single word "ok".' }],
    max_tokens: 10,
  });
  if (!response.choices[0]?.message?.content) {
    throw new Error("Empty response from Azure AI Foundry");
  }
}

// ─── Build vision content parts ───────────────────────────────────────────────
function buildImageParts(images: InputImage[]): OpenAI.Chat.ChatCompletionContentPart[] {
  return images.map((img) => ({
    type: "image_url" as const,
    image_url: { url: `data:${img.mimeType};base64,${img.base64}` },
  }));
}

// ─── Anthropic Messages API streaming (for Claude on services.ai.azure.com) ──
async function* streamAnthropic(
  config: AzureConfig,
  systemPrompt: string,
  messages: { role: "user" | "assistant"; content: string }[],
  signal?: AbortSignal
): AsyncGenerator<string> {
  const base = stripToBase(config.endpoint);
  const url = `${base}/anthropic/v1/messages`;

  const resp = await fetch(url, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "x-api-key": config.apiKey,
      "anthropic-version": "2023-06-01",
    },
    body: JSON.stringify({
      model: config.deploymentName,
      max_tokens: config.maxTokens,
      temperature: config.temperature,
      system: systemPrompt,
      messages,
      stream: true,
    }),
    signal,
  });

  if (!resp.ok) {
    const errText = await resp.text();
    let msg = `HTTP ${resp.status}`;
    try { msg += ` — ${(JSON.parse(errText) as { error?: { message?: string } }).error?.message ?? errText}`; } catch { msg += ` — ${errText}`; }
    throw new Error(msg);
  }

  // SSE stream: each line is "data: {...}" or "data: [DONE]"
  const reader = resp.body!.getReader();
  const decoder = new TextDecoder();
  let buf = "";

  while (true) {
    const { done, value } = await reader.read();
    if (done) break;
    buf += decoder.decode(value, { stream: true });
    const lines = buf.split("\n");
    buf = lines.pop() ?? "";
    for (const line of lines) {
      if (!line.startsWith("data:")) continue;
      const json = line.slice(5).trim();
      if (json === "[DONE]") return;
      try {
        const ev = JSON.parse(json) as { type?: string; delta?: { type?: string; text?: string } };
        if (ev.type === "content_block_delta" && ev.delta?.type === "text_delta") {
          yield ev.delta.text ?? "";
        }
      } catch { /* skip malformed */ }
    }
  }
}

// ─── Main streaming generator ─────────────────────────────────────────────────
export async function* generateDocumentStream(
  config: AzureConfig,
  systemPrompt: string,
  userPrompt: string,
  images: InputImage[],
  signal?: AbortSignal
): AsyncGenerator<string> {
  // Claude on services.ai.azure.com → Anthropic native API
  if (isServicesEndpoint(config.endpoint) && isClaudeModel(config.deploymentName)) {
    yield* streamAnthropic(config, systemPrompt, [{ role: "user", content: userPrompt }], signal);
    return;
  }

  // OpenAI-compatible path
  const client = createAzureClient(config);
  const userContent: OpenAI.Chat.ChatCompletionContentPart[] = [
    { type: "text", text: userPrompt },
    ...buildImageParts(images),
  ];
  const stream = await client.chat.completions.create({
    model: config.deploymentName,
    messages: [
      { role: "system", content: systemPrompt },
      { role: "user", content: userContent },
    ],
    stream: true,
    max_tokens: config.maxTokens,
    temperature: config.temperature,
  }, { signal });

  for await (const chunk of stream) {
    if (signal?.aborted) break;
    const delta = chunk.choices[0]?.delta?.content ?? "";
    if (delta) yield delta;
  }
}

// ─── Refinement (chat follow-up) ─────────────────────────────────────────────
export async function* refineDocumentStream(
  config: AzureConfig,
  currentDoc: DocumentOutput,
  chatHistory: ChatMessage[],
  newInstruction: string,
  images: InputImage[],
  signal?: AbortSignal
): AsyncGenerator<string> {
  const systemPrompt = `You are an expert document assistant. The user has an existing document structure in JSON.
They will give you a refinement instruction. Apply the instruction and return the COMPLETE updated document as valid JSON.
Return ONLY valid JSON — no markdown, no code fences, no explanation.`;

  if (isServicesEndpoint(config.endpoint) && isClaudeModel(config.deploymentName)) {
    const msgs: { role: "user" | "assistant"; content: string }[] = [
      ...chatHistory.map((m) => ({ role: m.role as "user" | "assistant", content: m.content })),
      {
        role: "user" as const,
        content: `Current document JSON:\n${JSON.stringify(currentDoc, null, 2)}\n\nInstruction: ${newInstruction}`,
      },
    ];
    yield* streamAnthropic(config, systemPrompt, msgs, signal);
    return;
  }

  const client = createAzureClient(config);
  const messages: OpenAI.Chat.ChatCompletionMessageParam[] = [
    { role: "system", content: systemPrompt },
    ...chatHistory.map((m) => ({
      role: m.role as "user" | "assistant",
      content: m.content,
    })),
    {
      role: "user",
      content: [
        {
          type: "text" as const,
          text: `Current document JSON:\n${JSON.stringify(currentDoc, null, 2)}\n\nInstruction: ${newInstruction}`,
        },
        ...buildImageParts(images),
      ],
    },
  ];

  const stream = await client.chat.completions.create({
    model: config.deploymentName,
    messages,
    stream: true,
    max_tokens: config.maxTokens,
    temperature: config.temperature,
  }, { signal });

  for await (const chunk of stream) {
    if (signal?.aborted) break;
    const delta = chunk.choices[0]?.delta?.content ?? "";
    if (delta) yield delta;
  }
}

// ─── JSON parser with retry ───────────────────────────────────────────────────
export function parseDocumentJSON(raw: string): DocumentOutput {
  const cleaned = raw
    .replace(/^```json\s*/i, "")
    .replace(/^```\s*/i, "")
    .replace(/```\s*$/i, "")
    .trim();

  const parsed = JSON.parse(cleaned) as DocumentOutput;

  if (!parsed.title) throw new Error("Missing 'title' in LLM response");
  if (!parsed.document_type) throw new Error("Missing 'document_type' in LLM response");

  if (parsed.slides) {
    parsed.slides = parsed.slides.map((s, i) => ({
      ...s,
      slide_number: s.slide_number ?? i + 1,
      speaker_notes: s.speaker_notes ?? "",
    }));
  }

  return parsed;
}
