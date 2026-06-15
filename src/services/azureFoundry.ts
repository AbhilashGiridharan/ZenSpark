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
  // Build a compact slide index so the LLM can reason about specific slides
  const slideIndex = (currentDoc.slides ?? [])
    .map((s) => `  Slide ${s.slide_number} [${s.layout}]: "${s.title}"`)
    .join("\n");

  const systemPrompt = `You are an expert document assistant refining an existing presentation.

Current slide index:
${slideIndex || "(no slides yet)"}

Instructions for slide-specific edits:
- If the user references "slide N" or a slide title, update ONLY that slide — keep all other slides exactly as they are.
- If the instruction is global (e.g. "change theme", "make it shorter"), apply it to all slides.
- Return the COMPLETE updated document JSON — all slides must be present.
- Maintain all existing fields (html, background_html, speaker_notes) unless the instruction specifically changes content.
- ONLY return raw JSON. No markdown. No code fences. No explanation. Start with { end with }.`;

  if (isServicesEndpoint(config.endpoint) && isClaudeModel(config.deploymentName)) {
    const msgs: { role: "user" | "assistant"; content: string }[] = [
      ...chatHistory.slice(0, -1).map((m) => ({ role: m.role as "user" | "assistant", content: m.content })),
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
    ...chatHistory.slice(0, -1).map((m) => ({
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

// ─── Clarifying questions (pre-generation) ───────────────────────────────────
export interface ClarifyingQuestion {
  id: string;
  question: string;
  hint?: string;
}

export async function generateClarifyingQuestions(
  config: AzureConfig,
  userGoal: string,
  fileNames: string[],
  signal?: AbortSignal
): Promise<ClarifyingQuestion[]> {
  const systemPrompt = `You are a presentation strategist. The user wants to generate a presentation.
Analyze their goal and uploaded files, then generate 3-5 SHORT clarifying questions that would help you create a much better presentation.
Focus on: audience, key message, specific data to highlight, tone/formality, must-include content.
Return ONLY a JSON array — no markdown, no explanation. Example:
[
  {"id":"q1","question":"Who is the primary audience?","hint":"e.g. C-suite, engineers, investors"},
  {"id":"q2","question":"What is the single most important takeaway?","hint":"The one thing they must remember"}
]`;

  const userMsg = `User goal: ${userGoal}\nUploaded files: ${fileNames.join(", ") || "none"}`;

  let raw = "";

  if (isServicesEndpoint(config.endpoint) && isClaudeModel(config.deploymentName)) {
    for await (const chunk of streamAnthropic(config, systemPrompt, [{ role: "user", content: userMsg }], signal)) {
      raw += chunk;
    }
  } else {
    const client = createAzureClient(config);
    const stream = await client.chat.completions.create({
      model: config.deploymentName,
      messages: [
        { role: "system", content: systemPrompt },
        { role: "user", content: userMsg },
      ],
      stream: true,
      max_tokens: 800,
      temperature: 0.4,
    }, { signal });
    for await (const chunk of stream) {
      raw += chunk.choices[0]?.delta?.content ?? "";
    }
  }

  try {
    const start = raw.indexOf("[");
    const end = raw.lastIndexOf("]");
    if (start !== -1 && end > start) {
      return JSON.parse(raw.slice(start, end + 1)) as ClarifyingQuestion[];
    }
  } catch { /* fall through */ }
  return [];
}

// ─── JSON parser with truncation recovery ───────────────────────────────────
export function parseDocumentJSON(raw: string): DocumentOutput {
  // Aggressively extract just the JSON object — find first { and last }
  // This handles: ```json\n{...}\n```, extra preamble text, trailing commentary
  let cleaned = raw;

  // Strip markdown code fences (any variant: ```json, ```JSON, ``` etc.)
  cleaned = cleaned.replace(/^[\s\S]*?```(?:json)?\s*/i, "").replace(/```[\s\S]*$/i, "");

  // Find the outermost JSON object boundaries
  const firstBrace = cleaned.indexOf("{");
  const lastBrace = cleaned.lastIndexOf("}");
  if (firstBrace !== -1 && lastBrace > firstBrace) {
    cleaned = cleaned.slice(firstBrace, lastBrace + 1);
  } else if (firstBrace !== -1) {
    // Truncated — no closing brace, take everything from first {
    cleaned = cleaned.slice(firstBrace);
  }

  cleaned = cleaned.trim();

  // First try parsing as-is
  try {
    return validateAndNormalize(JSON.parse(cleaned) as DocumentOutput);
  } catch (firstErr) {
    // Attempt structural recovery for truncated responses
    const recovered = attemptRecovery(cleaned);
    try {
      return validateAndNormalize(JSON.parse(recovered) as DocumentOutput);
    } catch {
      // Re-throw original error with context
      throw new Error(
        `JSON parse failed (response may have been truncated). ` +
        `Try increasing Max Tokens in settings. Original error: ${
          firstErr instanceof Error ? firstErr.message : String(firstErr)
        }`
      );
    }
  }
}

function attemptRecovery(json: string): string {
  let s = json;

  // Remove trailing comma before closing (common in truncated arrays)
  s = s.replace(/,\s*$/, "");

  // If there's an unterminated string, close it
  // Count unescaped quotes to detect open string
  let inString = false;
  let i = 0;
  for (; i < s.length; i++) {
    if (s[i] === '"' && (i === 0 || s[i - 1] !== "\\")) inString = !inString;
  }
  if (inString) s += '"'; // close open string

  // Remove any trailing comma again after closing string
  s = s.replace(/,\s*$/, "");

  // Count open brackets/braces and close them
  const stack: string[] = [];
  inString = false;
  for (let j = 0; j < s.length; j++) {
    const ch = s[j];
    if (ch === '"' && (j === 0 || s[j - 1] !== "\\")) { inString = !inString; continue; }
    if (inString) continue;
    if (ch === "{" || ch === "[") stack.push(ch);
    else if (ch === "}" || ch === "]") stack.pop();
  }
  // Close in reverse
  while (stack.length > 0) {
    const open = stack.pop()!;
    s += open === "{" ? "}" : "]";
  }
  return s;
}

function validateAndNormalize(parsed: DocumentOutput): DocumentOutput {
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
