import OpenAI from "openai";
import type { AzureConfig, DocumentOutput, ChatMessage, InputImage } from "../types/document";

// ─── Client factory ───────────────────────────────────────────────────────────
export function createAzureClient(config: AzureConfig): OpenAI {
  return new OpenAI({
    apiKey: config.apiKey,
    baseURL: `${config.endpoint.replace(/\/$/, "")}/openai/deployments/${config.deploymentName}`,
    ...(config.apiVersion ? { defaultQuery: { "api-version": config.apiVersion } } : {}),
    defaultHeaders: { "api-key": config.apiKey },
    dangerouslyAllowBrowser: true,
  });
}

// ─── Connection test ──────────────────────────────────────────────────────────
export async function testConnection(config: AzureConfig): Promise<void> {
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

// ─── Main streaming generator ─────────────────────────────────────────────────
export async function* generateDocumentStream(
  config: AzureConfig,
  systemPrompt: string,
  userPrompt: string,
  images: InputImage[],
  signal?: AbortSignal
): AsyncGenerator<string> {
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
  const client = createAzureClient(config);

  const systemPrompt = `You are an expert document assistant. The user has an existing document structure in JSON.
They will give you a refinement instruction. Apply the instruction and return the COMPLETE updated document as valid JSON.
Return ONLY valid JSON — no markdown, no code fences, no explanation.`;

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
  // Strip markdown code fences if LLM added them
  const cleaned = raw
    .replace(/^```json\s*/i, "")
    .replace(/^```\s*/i, "")
    .replace(/```\s*$/i, "")
    .trim();

  const parsed = JSON.parse(cleaned) as DocumentOutput;

  // Validate required fields
  if (!parsed.title) throw new Error("Missing 'title' in LLM response");
  if (!parsed.document_type) throw new Error("Missing 'document_type' in LLM response");

  // Ensure speaker_notes on every slide
  if (parsed.slides) {
    parsed.slides = parsed.slides.map((s, i) => ({
      ...s,
      slide_number: s.slide_number ?? i + 1,
      speaker_notes: s.speaker_notes ?? "",
    }));
  }

  return parsed;
}
