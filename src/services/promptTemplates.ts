import type { UseCasePreset, OutputFormat, ThemeOption } from "../types/document";

// ─── JSON schema embedded in every system prompt ─────────────────────────────
const JSON_SCHEMA = `
IMPORTANT: Return ONLY a raw JSON object. Do NOT use markdown code fences. Do NOT include any explanation. Start with { and end with }.

For each slide you MUST generate an "html" field: a complete self-contained HTML string that visually renders that slide at exactly 960×540px.
HTML rules:
- Use only inline styles (no external CSS, no <link>, no <script>)
- Use web-safe or Google Fonts (@import in a <style> tag at the top is OK)
- The outermost element must be: <div style="width:960px;height:540px;overflow:hidden;position:relative;font-family:...">
- Make it visually rich: use gradients, colored shapes, large bold numbers, icons (use Unicode/emoji), accent bars
- Choose a modern color palette that matches the theme field
- NO images that reference external URLs — use CSS gradients and shapes instead
- Keep all text readable (sufficient contrast)

Schema:
{
  "title": "string",
  "document_type": "pptx" | "docx" | "both",
  "theme": "corporate_blue" | "dark_tech" | "minimal_white" | "green_growth",
  "author": "string",
  "date": "YYYY-MM-DD",
  "slides": [
    {
      "slide_number": number,
      "layout": "title"|"bullets"|"two_column"|"image_caption"|"table"|"quote"|"section_divider"|"agenda"|"stats"|"closing",
      "title": "string",
      "subtitle": "string (optional)",
      "bullets": ["string"] (optional — PREFIX each bullet with a relevant emoji icon),
      "left_title": "string (optional)",
      "left_column": ["string"] (optional),
      "right_title": "string (optional)",
      "right_column": ["string"] (optional),
      "table": { "headers": ["string"], "rows": [["string"]] } (optional),
      "quote": "string (optional)",
      "attribution": "string (optional)",
      "stat_cards": [{ "value": "string", "label": "string", "icon": "string (emoji)" }] (optional),
      "html": "string (REQUIRED — complete self-contained 960×540px HTML for this slide)",
      "speaker_notes": "string (required)"
    }
  ],
  "sections": [
    {
      "level": 1|2|3,
      "heading": "string",
      "paragraphs": ["string"],
      "bullets": ["string"] (optional),
      "table": { "headers": ["string"], "rows": [["string"]] } (optional)
    }
  ]
}
Layout guidance:
- Use "stats" layout for 2-4 key metrics/KPIs — show as large bold numbers in cards in the HTML
- Use "two_column" for comparisons, before/after, pros/cons
- Use "section_divider" between major topics — full-bleed colored background, large title
- Use "quote" for impactful quotes — oversized quotation mark, centered italic text
- Prefix ALL bullet text with a contextually relevant emoji (📊 📈 ✅ 🎯 ⚡ 🔒 💡 🚀 📋 🤝 🏆)
- Every slide MUST have html and speaker_notes
- Include "slides" for document_type "pptx" or "both"
- Include "sections" for document_type "docx" or "both"
`;

// ─── Preset system prompts ────────────────────────────────────────────────────
const PRESET_PROMPTS: Record<UseCasePreset, string> = {
  testing_rfp: `You are a senior QA consulting specialist writing a formal RFP response on behalf of a testing services firm.
Using the provided RFP document and test artifacts as context, generate a structured, professional presentation.
Include slides covering:
1. Executive Summary
2. Our Understanding of Your Requirements
3. Testing Approach & Methodology (unit, integration, E2E, performance, security)
4. Test Automation Strategy & Recommended Tools
5. Team Structure & Roles
6. Quality Metrics & KPIs
7. Sample Test Plan Summary
8. Our Differentiators & Track Record
9. Implementation Timeline
10. Next Steps & Call to Action
Tone: formal, technical, data-driven, consultative.
${JSON_SCHEMA}`,

  customer_proposal: `You are a pre-sales solution architect creating a compelling customer proposal.
Using the provided customer requirements, pain points, and product documentation, generate a persuasive proposal presentation.
Include slides covering:
1. Executive Summary
2. Understanding Your Challenges
3. Our Proposed Solution
4. Solution Architecture & Approach
5. Key Features & Benefits
6. Implementation Roadmap
7. Team & Credentials
8. Commercial Summary
9. Why Us / Differentiators
10. Next Steps
Tone: persuasive, solution-focused, executive-friendly.
${JSON_SCHEMA}`,

  demo_prep: `You are a solutions engineer preparing for a live product demonstration.
Using the provided demo scenario, audience description, and product documentation, generate a demo presentation and script.
Include slides covering:
1. Demo Overview & Objectives
2. Customer Scenario & Pain Points
3. Demo Flow Introduction
4. Step-by-Step Feature Walkthrough (3-5 key scenarios with talking points)
5. Key Differentiating Capabilities
6. Performance & Integration Highlights
7. Anticipated Q&A (top 5 questions + answers)
8. Call to Action & Next Steps
Use image_index to reference any uploaded screenshots in feature walkthrough slides.
Tone: engaging, story-driven, technically credible.
${JSON_SCHEMA}`,

  technical_arch: `You are a senior solution architect creating a technical architecture document.
Using the provided requirements and context, generate a detailed technical presentation.
Include slides covering:
1. Architecture Overview
2. Current State / Problem Statement
3. Proposed Architecture (high-level)
4. Component Deep-Dive
5. Data Flow & Integration Points
6. Security Architecture
7. Scalability & Performance Design
8. Technology Stack Justification
9. Implementation Phases
10. Risk Mitigation
Tone: technical, precise, diagram-oriented (describe diagrams in text where images not available).
${JSON_SCHEMA}`,

  executive_briefing: `You are a C-suite communications specialist creating an executive briefing.
Using the provided context, generate a concise, high-impact executive presentation.
Include slides covering:
1. Situation Overview
2. Key Findings / Insights
3. Strategic Recommendations
4. Business Impact & ROI
5. Implementation Plan (summary)
6. Resource Requirements
7. Risks & Mitigations
8. Decision Required / Next Steps
Tone: concise, data-backed, strategic, non-technical. Maximum 3 bullets per slide.
${JSON_SCHEMA}`,

  custom: `You are an expert presentation and document specialist.
Read the user's goal and any uploaded files carefully. Infer the audience, purpose, and best structure from the content.
Generate a comprehensive, well-structured presentation or document that best serves the user's stated intent.
Choose the most appropriate layouts, number of slides, and level of detail based on what you read.
Tone: professional, clear, and tailored to the inferred audience.
CRITICAL: Your entire response must be a single raw JSON object. Do NOT wrap it in markdown code fences. Do NOT include any text before or after the JSON. Start your response with { and end with }.
${JSON_SCHEMA}`,
};

// ─── Public exports ───────────────────────────────────────────────────────────
export const USE_CASE_LABELS: Record<UseCasePreset, string> = {
  testing_rfp: "Testing RFP Response",
  customer_proposal: "Customer Proposal",
  demo_prep: "Demo Preparation",
  technical_arch: "Technical Architecture",
  executive_briefing: "Executive Briefing",
  custom: "Custom",
};

export const THEME_LABELS: Record<ThemeOption, string> = {
  corporate_blue: "Corporate Blue",
  dark_tech: "Dark Tech",
  minimal_white: "Minimal White",
  green_growth: "Green Growth",
};

export const SLIDE_COUNT_OPTIONS = [8, 10, 12, 16, 20];
export const DEFAULT_SLIDE_COUNT = 12;

export const OUTPUT_FORMAT_LABELS: Record<OutputFormat, string> = {
  pptx: "PowerPoint (.pptx)",
  docx: "Word Document (.docx)",
  both: "Both (PPTX + DOCX)",
};

export function getSystemPrompt(useCase: UseCasePreset): string {
  return PRESET_PROMPTS[useCase];
}

export function buildUserPrompt(
  goal: string,
  fileTexts: { name: string; content: string }[],
  pastedText: string,
  outputFormat: OutputFormat,
  theme: ThemeOption,
  imageCount: number,
  customSystemPrompt?: string
): string {
  const lines: string[] = [];

  lines.push(`## Generation Goal`);
  lines.push(goal || "Generate a professional presentation based on the provided context.");
  lines.push("");

  lines.push(`## Output Settings`);
  // Infer document type from user goal if not explicitly stated
  const goalLower = goal.toLowerCase();
  const inferredFormat = goalLower.includes("word") || goalLower.includes("report") || goalLower.includes("docx")
    ? "docx"
    : goalLower.includes("both") || goalLower.includes("pptx and docx")
    ? "both"
    : outputFormat; // use passed value as fallback
  lines.push(`- Document type: ${inferredFormat}`);
  lines.push(`- Theme: ${theme}`);
  // Slide count: let the LLM decide based on content unless the user stated one explicitly
  const mentionedCount = goal ? goal.match(/(\d+)\s*slide/i) : null;
  if (mentionedCount) {
    lines.push(`- Target slide count: ${mentionedCount[1]} slides (as requested)`);
  } else {
    lines.push(`- Slide count: choose the most appropriate number based on content complexity (typical range 8-20, do not pad with filler slides)`);
  }
  if (imageCount > 0) {
    lines.push(`- Images provided: ${imageCount} screenshot(s) available (use image_index 0-${imageCount - 1} to embed them)`);
  }
  lines.push("");

  if (pastedText.trim()) {
    lines.push("## Pasted Content / Notes");
    lines.push(pastedText.trim());
    lines.push("");
  }

  for (const { name, content } of fileTexts) {
    if (content.trim()) {
      lines.push(`## File: ${name}`);
      lines.push(content.trim().slice(0, 40000)); // ~30k tokens per file, fits Claude's 200k context
      lines.push("");
    }
  }

  if (customSystemPrompt) {
    lines.push("## Additional Instructions");
    lines.push(customSystemPrompt);
  }

  return lines.join("\n");
}
