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

function isOpenAIModel(model: string): boolean {
  const m = model.toLowerCase();
  return m.startsWith("gpt-") || m.startsWith("o1") || m.startsWith("o3") || m.startsWith("o4") || m.startsWith("codex");
}

// ─── Client factory (OpenAI-compatible endpoints only) ───────────────────────
export function createAzureClient(config: AzureConfig): OpenAI {
  const base = stripToBase(config.endpoint);

  let baseURL: string;
  if (isAzureOpenAI(config.endpoint)) {
    // Classic Azure OpenAI Service
    baseURL = `${base}/openai/deployments/${config.deploymentName}`;
  } else if (isServicesEndpoint(config.endpoint) && isOpenAIModel(config.deploymentName)) {
    // Foundry unified endpoint + GPT/o-series → Azure OpenAI path
    baseURL = `${base}/openai/deployments/${config.deploymentName}`;
  } else if (isServicesEndpoint(config.endpoint)) {
    // Foundry unified endpoint + non-OpenAI model (Llama, Phi…) → AI Inference path
    baseURL = `${base}/models`;
  } else {
    baseURL = `${base}/v1`;
  }

  const apiVersion = config.apiVersion || (isServicesEndpoint(config.endpoint) ? "2025-04-01-preview" : undefined);

  return new OpenAI({
    apiKey: config.apiKey,
    baseURL,
    ...(apiVersion ? { defaultQuery: { "api-version": apiVersion } } : {}),
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

// Anthropic-native image blocks
type AnthropicBlock =
  | { type: "text"; text: string }
  | { type: "image"; source: { type: "base64"; media_type: string; data: string } };
type AnthropicMessage = { role: "user" | "assistant"; content: string | AnthropicBlock[] };

function buildAnthropicImageParts(images: InputImage[]): AnthropicBlock[] {
  return images.map((img) => ({
    type: "image" as const,
    source: { type: "base64" as const, media_type: img.mimeType, data: img.base64 },
  }));
}

// Attach images to the last user message in an AnthropicMessage array
function attachImagesToLastUserMsg(msgs: AnthropicMessage[], images: InputImage[]): AnthropicMessage[] {
  if (!images.length) return msgs;
  const result = [...msgs];
  const lastIdx = result.length - 1;
  const last = result[lastIdx];
  const textContent = typeof last.content === "string" ? last.content : "";
  result[lastIdx] = {
    ...last,
    content: [
      ...buildAnthropicImageParts(images),
      { type: "text", text: textContent },
    ],
  };
  return result;
}

// ─── Anthropic Messages API streaming (for Claude on services.ai.azure.com) ──
async function* streamAnthropic(
  config: AzureConfig,
  systemPrompt: string,
  messages: AnthropicMessage[],
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
    const msgs: AnthropicMessage[] = [{ role: "user", content: userPrompt }];
    yield* streamAnthropic(config, systemPrompt, attachImagesToLastUserMsg(msgs, images), signal);
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
    const msgs: AnthropicMessage[] = [
      ...chatHistory.slice(0, -1).map((m) => ({ role: m.role as "user" | "assistant", content: m.content })),
      {
        role: "user" as const,
        content: `Current document JSON:\n${JSON.stringify(currentDoc, null, 2)}\n\nInstruction: ${newInstruction}`,
      },
    ];
    yield* streamAnthropic(config, systemPrompt, attachImagesToLastUserMsg(msgs, images), signal);
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
Analyze their goal and uploaded files, then generate 3-5 SHORT clarifying questions that would significantly improve the output.
You MUST always return questions — never return an empty array.
Focus on: target audience, key message/objective, company/client name, specific data or metrics to highlight, tone/formality, number of slides desired.
Return ONLY a JSON array — no markdown, no explanation. Example:
[
  {"id":"q1","question":"Who is the primary audience?","hint":"e.g. C-suite, engineers, investors"},
  {"id":"q2","question":"What is the single most important takeaway?","hint":"The one thing they must remember"},
  {"id":"q3","question":"What is the company or client name?","hint":"Used for branding in the slides"}
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

// ─── Smart dual-mode chat (conversational OR doc-edit) ───────────────────────
// If the user is asking a general question → reply naturally.
// If the user wants to edit the document → return only JSON.
export async function* smartChatStream(
  config: AzureConfig,
  currentDoc: DocumentOutput | null,
  chatHistory: ChatMessage[],
  newMessage: string,
  images: InputImage[],
  fileTexts: { name: string; content: string }[],
  signal?: AbortSignal
): AsyncGenerator<string> {
  const slideIndex = (currentDoc?.slides ?? [])
    .map((s) => `  Slide ${s.slide_number} [${s.layout}]: "${s.title}"`)
    .join("\n");

  const docContext = currentDoc
    ? `\nThe user currently has a presentation open with these slides:\n${slideIndex || "(no slides yet)"}\n`
    : "";

  const systemPrompt = `You are ZenSpark, an intelligent AI assistant that can both converse naturally and create/edit professional presentations.
${docContext}
Decide how to respond based on the user's message:

1. GENERAL CONVERSATION — if the user is asking a question, having a discussion, requesting advice, brainstorming, or anything not directly modifying the presentation:
   → Respond naturally in plain text. Be helpful, concise, and thoughtful.

2. DOCUMENT EDIT or CREATION — if the user wants to change, update, add, remove, restructure slides, OR create/generate a new presentation:
   → Return ONLY valid JSON of the complete document. No text before or after. No markdown. No code fences. Your response must start with { and end with }.
   → The JSON MUST always include ALL top-level fields: "title", "document_type" ("pptx"), "theme", "author", "date", "slides".
   → Each slide MUST include: slide_number, layout, title, html, background_html, speaker_notes, and any content fields for that layout (bullets, stat_cards, etc.).

When in document-edit mode:
- If a specific slide is referenced, update ONLY that slide.
- Keep all other slides exactly as-is.
- Return ALL slides in the JSON — no slides may be omitted.
- Preserve all existing fields (html, background_html, speaker_notes) unless instructed otherwise.`;

  // Build the user text payload — include attached file content
  // Strip html/background_html from doc context to avoid token explosion
  const slimDoc = currentDoc ? {
    ...currentDoc,
    slides: (currentDoc.slides ?? []).map(({ html: _h, background_html: _b, ...rest }) => rest),
  } : null;

  const fileContext = fileTexts.length
    ? `\n\nAttached files:\n${fileTexts.map((f) => `--- ${f.name} ---\n${f.content}`).join("\n\n")}`
    : "";
  const userText = slimDoc
    ? `Current document JSON:\n${JSON.stringify(slimDoc, null, 2)}\n\nUser message: ${newMessage}${fileContext}`
    : `${newMessage}${fileContext}`;

  if (isServicesEndpoint(config.endpoint) && isClaudeModel(config.deploymentName)) {
    const msgs: AnthropicMessage[] = [
      ...chatHistory.slice(0, -1).map((m) => ({
        role: m.role as "user" | "assistant",
        content: m.content,
      })),
      { role: "user" as const, content: userText },
    ];
    yield* streamAnthropic(config, systemPrompt, attachImagesToLastUserMsg(msgs, images), signal);
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
      role: "user" as const,
      content: [
        { type: "text" as const, text: userText },
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
  // Apply defaults rather than throwing — LLM sometimes omits top-level fields
  // (especially in chat-edit mode where it only generates slides array)
  if (!parsed.title) {
    // Try to extract a title from the first slide
    const firstSlideTitle = parsed.slides?.[0]?.title;
    parsed.title = firstSlideTitle ?? "Untitled Presentation";
  }
  if (!parsed.document_type) parsed.document_type = "pptx";
  if (!parsed.theme) parsed.theme = "corporate_blue";

  if (parsed.slides) {
    parsed.slides = parsed.slides.map((s, i) => {
      const slide = {
        ...s,
        slide_number: s.slide_number ?? i + 1,
        layout: s.layout ?? "bullets",
        title: s.title ?? `Slide ${i + 1}`,
        speaker_notes: s.speaker_notes ?? "",
      };
      // Always regenerate html from pptx_elements (never trust LLM-generated html)
      if (slide.pptx_elements && (slide.pptx_elements as unknown[]).length > 0) {
        slide.html = pptxElementsToHtml(slide.pptx_elements as unknown as Array<Record<string, unknown>>);
      } else if (!slide.html) {
        // Fallback: simple white slide with title + content text
        const titleText = String(slide.title ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;");
        const bodyText = String(
          (slide as Record<string, unknown>).subtitle ??
          ((slide as Record<string, unknown>).bullets as string[] | undefined)?.join(" · ") ??
          (slide as Record<string, unknown>).content ?? ""
        ).replace(/&/g, "&amp;").replace(/</g, "&lt;");
        slide.html = `<div style="position:relative;width:960px;height:540px;background:#1a1a2e;font-family:Calibri,Arial,sans-serif;display:flex;flex-direction:column;justify-content:center;padding:60px">` +
          `<div style="color:#ffffff;font-size:36px;font-weight:bold;margin-bottom:20px">${titleText}</div>` +
          `<div style="color:#ccddff;font-size:20px;line-height:1.5">${bodyText}</div></div>`;
      }
      return slide;
    });
  }
  return parsed;
}

// ─── Client-side HTML renderer from pptx_elements (no LLM needed) ─────────────
// Converts pptx_elements (10×7.5" coordinate space) to a 960×540px HTML preview.
// Used when LLM omits html due to token limits (common with 10+ slides).
function pptxElementsToHtml(elements: Array<Record<string, unknown>>): string {
  const W = 960, H = 540;
  const SRC_W = 10, SRC_H = 7.5;
  const sx = W / SRC_W, sy = H / SRC_H;

  const toPx = (v: unknown, axis: "x" | "w" | "y" | "h"): number => {
    const scale = axis === "x" || axis === "w" ? sx : sy;
    const total = axis === "x" || axis === "w" ? W : H;
    if (typeof v === "string" && v.endsWith("%")) return (parseFloat(v) / 100) * total;
    return parseFloat(String(v)) * scale;
  };

  const hex = (v: unknown, fallback = "000000") => {
    const s = String(v ?? "").replace("#", "").trim();
    return /^[0-9a-fA-F]{3,6}$/.test(s) ? `#${s}` : `#${fallback}`;
  };

  const ALIGN_CSS: Record<string, string> = { left: "left", center: "center", right: "right", justify: "justify" };
  const VALIGN_FLEX: Record<string, string> = { top: "flex-start", middle: "center", bottom: "flex-end" };

  let bg = "#111827";
  for (const el of elements) {
    if (el.type === "rect" && toPx(el.w, "w") >= W * 0.95 && toPx(el.h, "h") >= H * 0.90 && el.fill) {
      bg = hex(el.fill);
      break;
    }
  }

  const parts: string[] = [
    `<div style="position:relative;width:${W}px;height:${H}px;overflow:hidden;background:${bg};font-family:Calibri,Arial,sans-serif">`
  ];

  // Render non-text elements first, text on top to avoid being covered by rects
  const sorted = [...elements].sort((a, b) => {
    const order = (t: unknown) => t === "text" ? 1 : 0;
    return order(a.type) - order(b.type);
  });

  for (const el of sorted) {
    const x = Math.round(toPx(el.x ?? 0, "x"));
    const y = Math.round(toPx(el.y ?? 0, "y"));
    const w = Math.round(toPx(el.w ?? 1, "w"));
    const h = Math.round(toPx(el.h ?? 0.5, "h"));
    const fill = el.fill ? hex(el.fill) : "";
    const trans = el.transparency ? `opacity:${1 - Number(el.transparency) / 100};` : "";
    const pos = `position:absolute;left:${x}px;top:${y}px;width:${w}px;height:${h}px;`;

    if (el.type === "rect") {
      parts.push(`<div style="${pos}background:${fill || "transparent"};${trans}"></div>`);
    } else if (el.type === "ellipse") {
      parts.push(`<div style="${pos}background:${fill || "transparent"};border-radius:50%;${trans}"></div>`);
    } else if (el.type === "line") {
      parts.push(`<div style="${pos}background:${fill || hex(el.color, "888888")};"></div>`);
    } else if (el.type === "text" && el.text != null) {
      // pt → CSS px: 1pt = 1/72" × 96px/" = 1.333px
      const fs = Math.round((Number(el.fontSize) || 14) * (96 / 72));
      const fw = el.bold ? "bold" : "normal";
      const fs2 = el.italic ? "italic" : "normal";
      const color = hex(el.color, "000000");
      const align = ALIGN_CSS[String(el.align ?? "left")] ?? "left";
      const valign = VALIGN_FLEX[String(el.valign ?? "top")] ?? "flex-start";
      const wrap = el.wrap === false ? "nowrap" : "normal";
      const fillStyle = fill ? `background:${fill};` : "";
      const text = String(el.text).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/\n/g, "<br>");
      parts.push(
        `<div style="${pos}${fillStyle}display:flex;align-items:${valign};justify-content:${align === "center" ? "center" : align === "right" ? "flex-end" : "flex-start"};overflow:hidden;white-space:${wrap};">` +
        `<span style="font-size:${fs}px;font-weight:${fw};font-style:${fs2};color:${color};text-align:${align};line-height:1.25;word-break:break-word;">${text}</span>` +
        `</div>`
      );
    }
  }

  parts.push("</div>");
  return parts.join("");
}
