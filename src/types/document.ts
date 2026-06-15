// ─── Azure Config ────────────────────────────────────────────────────────────
export interface AzureConfig {
  endpoint: string;
  apiKey: string;
  deploymentName: string;
  apiVersion?: string;
  maxTokens: number;
  temperature: number;
  visionDeploymentName: string;
}

// ─── Output Format / Theme / Use Case ────────────────────────────────────────
export type ThemeOption =
  | "corporate_blue"
  | "dark_tech"
  | "minimal_white"
  | "green_growth";

export type OutputFormat = "pptx" | "docx" | "both";

export type UseCasePreset =
  | "testing_rfp"
  | "customer_proposal"
  | "demo_prep"
  | "technical_arch"
  | "executive_briefing"
  | "custom";

// ─── Slide types ──────────────────────────────────────────────────────────────
export type SlideLayout =
  | "title"
  | "bullets"
  | "two_column"
  | "image_caption"
  | "table"
  | "quote"
  | "section_divider"
  | "agenda"
  | "stats"
  | "closing";

export interface StatCard {
  value: string;  // e.g. "94%", "$2.4M", "3x"
  label: string;  // e.g. "Customer Satisfaction"
  icon?: string;  // optional emoji e.g. "📊"
}

export interface Slide {
  slide_number: number;
  layout: SlideLayout;
  title: string;
  subtitle?: string;
  bullets?: string[];
  left_column?: string[];
  right_column?: string[];
  left_title?: string;
  right_title?: string;
  table?: { headers: string[]; rows: string[][] };
  quote?: string;
  attribution?: string;
  image_index?: number;
  stat_cards?: StatCard[];
  html?: string;           // LLM-generated HTML — full visual (browser preview only)
  background_html?: string; // LLM-generated HTML — decorative background only (no text), used as PPTX bg image
  speaker_notes: string;
}

// ─── Document section (for DOCX) ─────────────────────────────────────────────
export interface Section {
  level: 1 | 2 | 3;
  heading: string;
  paragraphs: string[];
  bullets?: string[];
  table?: { headers: string[]; rows: string[][] };
  image_index?: number;
}

// ─── Full LLM output schema ───────────────────────────────────────────────────
export interface DocumentOutput {
  title: string;
  document_type: "pptx" | "docx" | "both";
  theme: ThemeOption;
  author: string;
  date: string;
  slides?: Slide[];
  sections?: Section[];
}

// ─── Input types ──────────────────────────────────────────────────────────────
export interface InputFile {
  id: string;
  name: string;
  size: number;
  type: string;
  extractedText: string;
}

export interface InputImage {
  id: string;
  name: string;
  base64: string;
  mimeType: string;
  caption: string;
  preview: string; // data URL for <img> preview
}

// ─── Chat ─────────────────────────────────────────────────────────────────────
export interface ChatMessage {
  role: "user" | "assistant";
  content: string;
  timestamp: Date;
}

// ─── Token usage ──────────────────────────────────────────────────────────────
export interface TokenUsage {
  promptTokens: number;
  completionTokens: number;
  totalTokens: number;
}
