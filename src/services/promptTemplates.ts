import type { UseCasePreset, OutputFormat, ThemeOption } from "../types/document";

// ─── JSON schema embedded in every system prompt ─────────────────────────────
const JSON_SCHEMA = `
IMPORTANT: Return ONLY a raw JSON object. Do NOT use markdown code fences. Do NOT include any explanation. Start with { and end with }.

CRITICAL — EVERY slide object MUST contain ALL of the following fields. Omitting any of them will break the editable PowerPoint export:
  • slide_number (integer)
  • layout (one of the values listed below)
  • title (string — the slide heading)
  • The content fields that match the layout (bullets, left_column/right_column, table, stat_cards, quote — see below)
  • pptx_elements (array — see rules below)
  • speaker_notes (string)

IMPORTANT: Do NOT include html or background_html fields — these are generated client-side. Focus token budget on pptx_elements, content fields, and speaker_notes.

3. "pptx_elements" — an array of explicit PowerPoint drawing instructions that reproduce the visual design of the slide as editable elements.
   This is used to generate a REAL editable PPTX that looks like your HTML design — every shape, colored bar, text box must be listed here.
   Rules:
   - Coordinate system: slide is 10 inches wide × 7.5 inches tall. x/y/w/h are in inches (decimals OK). Use "100%" for full width/height.
   - Each element has a "type": "rect", "ellipse", "text", or "line"
   - "rect" and "ellipse": must have "fill" (hex color WITHOUT #, e.g. "1A1A2E"). Optional "transparency" (0-100).
   - "text": must have "text" (string), "fontSize" (number), "color" (hex WITHOUT #). Optional: "bold", "italic", "fontFace", "align" ("left"/"center"/"right"), "valign" ("top"/"middle"/"bottom"), "wrap" (true/false).
   - VISUAL DESIGN — use your own professional design judgment:
     • You are acting as a presentation designer. Design each slide to be visually clear and impactful.
     • CRITICAL: The Output Settings section specifies a Theme and Brand. If a specific brand/color palette is listed, it is MANDATORY — use ONLY those colors in pptx_elements fills and text colors. Do not substitute with your own color preferences.
     • If no brand palette is specified, use your own color choices. Ensure readability (body text ≥13pt), strong visual hierarchy, and variety across slides.
     • Do NOT homogenize every slide — use the full range of layout types (stats, quote, two_column, section_divider, etc.) to keep the deck visually engaging.
   - Draw background rects FIRST (lowest z-order), then decorative shapes, then text boxes on top.
   - Font sizes are in pt. Convert from px if needed: pt = px ÷ 1.33.
   - For multi-line bullet lists: create one "text" element per bullet OR a single text element with newlines ("\\n") between items.
   - For tables: create header rect + header text, then alternating row rects + row text elements.
   - Convert ALL readable content (titles, bullets, table cells, stat numbers) into "text" type elements so content is editable.
   Example pptx_elements for a slide with a dark header bar, accent line, and two bullet points:
   [
     {"type":"rect","x":0,"y":0,"w":"100%","h":1.1,"fill":"0D47A1"},
     {"type":"rect","x":0,"y":1.05,"w":"100%","h":0.05,"fill":"42A5F5"},
     {"type":"text","x":0.3,"y":0.1,"w":9.4,"h":0.9,"text":"My Slide Title","fontSize":28,"bold":true,"color":"FFFFFF","fontFace":"Calibri","valign":"middle"},
     {"type":"text","x":0.4,"y":1.25,"w":9.2,"h":0.5,"text":"🔵 First bullet point — concise, scannable, substantive","fontSize":16,"color":"1A2744","fontFace":"Calibri","valign":"middle"},
     {"type":"rect","x":0.3,"y":1.85,"w":9.4,"h":0.45,"fill":"EFF6FF","transparency":20},
     {"type":"text","x":0.4,"y":1.85,"w":9.2,"h":0.45,"text":"🟢 Second bullet with a highlighted background","fontSize":16,"color":"1A2744","fontFace":"Calibri","valign":"middle"}
   ]

For each slide you MUST generate pptx_elements. Do NOT include html or background_html \u2014 these are generated client-side.

Schema:
  "title": "string",
  "document_type": "pptx" | "docx" | "both",
  "theme": "corporate_blue" | "dark_tech" | "minimal_white" | "green_growth" | "zensar_white",
  "author": "string",
  "date": "YYYY-MM-DD",
  "slides": [
    {
      "slide_number": number,                    ← REQUIRED on every slide
      "layout": "title"|"bullets"|"two_column"|"image_caption"|"table"|"quote"|"section_divider"|"agenda"|"stats"|"closing",  ← REQUIRED
      "title": "string",                         ← REQUIRED on every slide
      "subtitle": "string (optional)",
      "bullets": ["string"] (REQUIRED for layout=bullets|agenda — PREFIX each bullet with a relevant emoji icon),
      "left_title": "string (optional)",
      "left_column": ["string"] (REQUIRED for layout=two_column),
      "right_title": "string (optional)",
      "right_column": ["string"] (REQUIRED for layout=two_column),
      "table": { "headers": ["string"], "rows": [["string"]] } (REQUIRED for layout=table),
      "quote": "string (REQUIRED for layout=quote)",
      "attribution": "string (optional)",
      "stat_cards": [{ "value": "string", "label": "string", "icon": "string (emoji)" }] (REQUIRED for layout=stats),
      "pptx_elements": [ ...array of drawing instructions — REQUIRED on every slide, see rules above... ],
      "speaker_notes": "string — REQUIRED"
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
- Every slide MUST have pptx_elements and speaker_notes (html is auto-generated)
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

  custom: `You are an expert presentation designer and content strategist.
Read the user's goal and any uploaded files carefully. Infer the audience, purpose, and best structure from the content.
Generate a visually compelling presentation or document that best serves the user's stated intent.

Design philosophy: make each slide visually distinct and impactful. Use your full creative design judgment — choose colors, typography, layout variety, and visual hierarchy that communicate the content clearly and professionally. A great presentation is not a wall of text — vary layouts, use stat cards, callout boxes, section dividers, and visual accents to keep the reader engaged.

IMPORTANT: If the user asks for "a slide", "one slide", "convert this image", or references a single item — generate exactly 1 slide, not more.
If the Output Settings specify a target slide count, follow it exactly — do not generate more or fewer slides.
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
  zensar_white: "Zensar (White)",
};

export const SLIDE_COUNT_OPTIONS = [8, 10, 12, 16, 20, 25];
export const DEFAULT_SLIDE_COUNT = 15;

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
  // Zensar brand guidance — MANDATORY when zensar_white is selected
  if (theme === "zensar_white") {
    lines.push(`- Brand: Zensar Technologies. The following color palette is MANDATORY — do not deviate:`);
    lines.push(`    Slide background: #FFFFFF (white) — every slide must have a white background`);
    lines.push(`    Header bars / primary filled shapes: #003A70 (Zensar navy)`);
    lines.push(`    Section labels / sub-headings: #9A1F1F (Zensar crimson)`);
    lines.push(`    All body text on white: #1A1F2B (dark navy)`);
    lines.push(`    Muted sub-labels / secondary text: #525A6B (steel gray)`);
    lines.push(`    Positive/growth values: #1F6A3A (forest green)`);
    lines.push(`    Card / panel backgrounds: #F9FDFC (very light off-white)`);
    lines.push(`    Accent line / divider: #D0D9E6`);
    lines.push(`    Font face: Calibri`);
    lines.push(`  STRICT RULES for Zensar theme:`);
    lines.push(`    - NO dark backgrounds. The slide background MUST be white (#FFFFFF) or very light (#F9FDFC).`);
    lines.push(`    - Dark fills are ONLY allowed for header bars (top strip) — color MUST be #003A70.`);
    lines.push(`    - All text on white background MUST use #1A1F2B, #525A6B, #9A1F1F, #003A70, or #1F6A3A.`);
    lines.push(`    - Do NOT use any blue shades other than #003A70 for fills, or #003A70/#525A6B for text.`);
    lines.push(`    - The Zensar logo will be placed top-right automatically — do NOT include it in pptx_elements.`);
  }
  // Slide count: respect explicit numbers, singular intent, or let LLM decide
  const mentionedCount = goal ? goal.match(/(\d+)\s*slide/i) : null;
  const singularSlide = goal ? /\b(a|one|single|1)\s+slide\b/i.test(goal) || /\bconvert\b.*\bimage\b/i.test(goal) || /\bthis\s+(image|photo|screenshot)\b/i.test(goal) : false;
  if (mentionedCount) {
    lines.push(`- Target slide count: ${mentionedCount[1]} slides (as requested)`);
  } else if (singularSlide) {
    lines.push(`- Target slide count: 1 slide (user wants a single slide)`);
  } else {
    lines.push(`- Slide count: choose the most appropriate number based on content complexity. For detailed documents (RFPs, proposals, technical specs) aim for 15-25 slides. For simple topics use 8-12. Do not pad with filler.`);
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
