// eslint-disable-next-line @typescript-eslint/ban-ts-comment
// @ts-ignore — pptxgenjs types
import pptxgen from "pptxgenjs";
import { saveAs } from "file-saver";
import type { DocumentOutput, Slide, ThemeOption, InputImage } from "../types/document";

// ─── Theme definitions ────────────────────────────────────────────────────────
interface ThemeColors {
  background: string;
  headerBg: string;
  headerText: string;
  bodyText: string;
  accent: string;
  tableHeaderBg: string;
  tableHeaderText: string;
  tableRowBg: string;
  divider: string;
}

const THEMES: Record<ThemeOption, ThemeColors> = {
  corporate_blue: {
    background: "FFFFFF",
    headerBg: "1565C0",
    headerText: "FFFFFF",
    bodyText: "2D3748",
    accent: "1565C0",
    tableHeaderBg: "1565C0",
    tableHeaderText: "FFFFFF",
    tableRowBg: "E3F2FD",
    divider: "90CAF9",
  },
  dark_tech: {
    background: "0D1117",
    headerBg: "161B22",
    headerText: "58A6FF",
    bodyText: "C9D1D9",
    accent: "58A6FF",
    tableHeaderBg: "21262D",
    tableHeaderText: "58A6FF",
    tableRowBg: "161B22",
    divider: "30363D",
  },
  minimal_white: {
    background: "FFFFFF",
    headerBg: "FFFFFF",
    headerText: "1A1A1A",
    bodyText: "4A4A4A",
    accent: "E53E3E",
    tableHeaderBg: "F5F5F5",
    tableHeaderText: "1A1A1A",
    tableRowBg: "FAFAFA",
    divider: "E2E8F0",
  },
  green_growth: {
    background: "FFFFFF",
    headerBg: "2E7D32",
    headerText: "FFFFFF",
    bodyText: "1B4332",
    accent: "2E7D32",
    tableHeaderBg: "2E7D32",
    tableHeaderText: "FFFFFF",
    tableRowBg: "E8F5E9",
    divider: "A5D6A7",
  },
};

// ─── Helper: add header bar + slide title ────────────────────────────────────
function addHeaderBar(
  slide: ReturnType<InstanceType<typeof pptxgen>["addSlide"]>,
  title: string,
  tc: ThemeColors
) {
  // Full-width header bar
  slide.addShape("rect", {
    x: 0, y: 0, w: "100%", h: 0.85,
    fill: { color: tc.headerBg },
    line: { color: tc.headerBg },
  });
  // Title in header
  slide.addText(title, {
    x: 0.35, y: 0.05, w: 9.3, h: 0.75,
    fontSize: 22,
    bold: true,
    color: tc.headerText,
    fontFace: "Calibri",
    valign: "middle",
  });
}

// ─── Slide builders ──────────────────────────────────────────────────────────
function buildTitleSlide(
  prs: InstanceType<typeof pptxgen>,
  slide_: ReturnType<InstanceType<typeof pptxgen>["addSlide"]>,
  s: Slide,
  tc: ThemeColors
) {
  // Top half color block
  slide_.addShape("rect", {
    x: 0, y: 0, w: "100%", h: 2.8,
    fill: { color: tc.headerBg },
    line: { color: tc.headerBg },
  });
  // Title
  slide_.addText(s.title, {
    x: 0.5, y: 0.4, w: 9, h: 2.1,
    fontSize: 40,
    bold: true,
    color: tc.headerText,
    fontFace: "Calibri",
    valign: "middle",
    align: "center",
  });
  // Accent divider line
  slide_.addShape("rect", {
    x: 3.5, y: 2.9, w: 3, h: 0.06,
    fill: { color: tc.accent },
    line: { color: tc.accent },
  });
  // Subtitle
  if (s.subtitle) {
    slide_.addText(s.subtitle, {
      x: 0.5, y: 3.1, w: 9, h: 1.0,
      fontSize: 22,
      color: tc.bodyText,
      fontFace: "Calibri",
      align: "center",
      valign: "top",
    });
  }
  void prs;
}

function buildBulletsSlide(
  slide_: ReturnType<InstanceType<typeof pptxgen>["addSlide"]>,
  s: Slide,
  tc: ThemeColors
) {
  addHeaderBar(slide_, s.title, tc);
  const items = (s.bullets ?? []).map((b) => ({
    text: b,
    options: { bullet: { indent: 15 }, breakLine: true, color: tc.bodyText },
  }));
  if (items.length > 0) {
    slide_.addText(items, {
      x: 0.4, y: 1.0, w: 9.2, h: 4.3,
      fontSize: 18,
      fontFace: "Calibri",
      valign: "top",
      color: tc.bodyText,
    });
  }
}

function buildTwoColumnSlide(
  slide_: ReturnType<InstanceType<typeof pptxgen>["addSlide"]>,
  s: Slide,
  tc: ThemeColors
) {
  addHeaderBar(slide_, s.title, tc);
  // Divider line
  slide_.addShape("rect", {
    x: 5.0, y: 0.95, w: 0.02, h: 4.4,
    fill: { color: tc.divider },
    line: { color: tc.divider },
  });
  // Left column title
  if (s.left_title) {
    slide_.addText(s.left_title, {
      x: 0.35, y: 0.95, w: 4.5, h: 0.4,
      fontSize: 14, bold: true, color: tc.accent, fontFace: "Calibri",
    });
  }
  // Right column title
  if (s.right_title) {
    slide_.addText(s.right_title, {
      x: 5.15, y: 0.95, w: 4.5, h: 0.4,
      fontSize: 14, bold: true, color: tc.accent, fontFace: "Calibri",
    });
  }
  const leftY = s.left_title ? 1.45 : 1.0;
  const rightY = s.right_title ? 1.45 : 1.0;
  const colH = 3.9;
  // Left items
  const leftItems = (s.left_column ?? []).map((b) => ({
    text: b,
    options: { bullet: { indent: 10 }, breakLine: true, color: tc.bodyText },
  }));
  if (leftItems.length > 0) {
    slide_.addText(leftItems, { x: 0.35, y: leftY, w: 4.4, h: colH, fontSize: 16, fontFace: "Calibri", valign: "top" });
  }
  // Right items
  const rightItems = (s.right_column ?? []).map((b) => ({
    text: b,
    options: { bullet: { indent: 10 }, breakLine: true, color: tc.bodyText },
  }));
  if (rightItems.length > 0) {
    slide_.addText(rightItems, { x: 5.15, y: rightY, w: 4.5, h: colH, fontSize: 16, fontFace: "Calibri", valign: "top" });
  }
}

function buildImageCaptionSlide(
  slide_: ReturnType<InstanceType<typeof pptxgen>["addSlide"]>,
  s: Slide,
  tc: ThemeColors,
  images: InputImage[]
) {
  addHeaderBar(slide_, s.title, tc);
  const img = s.image_index !== undefined ? images[s.image_index] : undefined;
  if (img) {
    slide_.addImage({
      data: img.preview,
      x: 0.5, y: 1.0, w: 5.5, h: 3.8,
    });
    // Caption next to image
    const captionText = img.caption || s.subtitle || "";
    if (captionText) {
      slide_.addText(captionText, {
        x: 6.2, y: 1.0, w: 3.4, h: 3.8,
        fontSize: 16, color: tc.bodyText, fontFace: "Calibri", valign: "top",
      });
    }
  } else {
    // No image — show as bullets/content
    if (s.subtitle) {
      slide_.addText(s.subtitle, {
        x: 0.4, y: 1.0, w: 9.2, h: 4.3,
        fontSize: 18, color: tc.bodyText, fontFace: "Calibri", valign: "top",
      });
    }
  }
}

function buildTableSlide(
  slide_: ReturnType<InstanceType<typeof pptxgen>["addSlide"]>,
  s: Slide,
  tc: ThemeColors
) {
  addHeaderBar(slide_, s.title, tc);
  if (!s.table) return;

  const headerRow = s.table.headers.map((h) => ({
    text: h,
    options: {
      bold: true,
      color: tc.tableHeaderText,
      fill: { color: tc.tableHeaderBg },
      fontFace: "Calibri",
      fontSize: 14,
      align: "center" as const,
    },
  }));

  const dataRows = s.table.rows.map((row, i) =>
    row.map((cell) => ({
      text: cell,
      options: {
        fill: { color: i % 2 === 0 ? tc.tableRowBg : "FFFFFF" },
        color: tc.bodyText,
        fontFace: "Calibri",
        fontSize: 13,
      },
    }))
  );

  slide_.addTable([headerRow, ...dataRows], {
    x: 0.4, y: 1.0, w: 9.2,
    border: { type: "solid", color: tc.divider, pt: 1 },
    rowH: 0.45,
  });
}

function buildQuoteSlide(
  slide_: ReturnType<InstanceType<typeof pptxgen>["addSlide"]>,
  s: Slide,
  tc: ThemeColors
) {
  slide_.background = { color: tc.background };
  // Large decorative quote mark
  slide_.addText("\u201C", {
    x: 0.3, y: 0.1, w: 1.5, h: 1.5,
    fontSize: 96, color: tc.divider, fontFace: "Georgia", bold: true,
  });
  // Quote text
  slide_.addText(s.quote ?? s.title, {
    x: 0.5, y: 1.2, w: 9, h: 3.0,
    fontSize: 24, color: tc.bodyText, fontFace: "Georgia",
    italic: true, align: "center", valign: "middle",
  });
  // Attribution
  if (s.attribution) {
    slide_.addText(`— ${s.attribution}`, {
      x: 0.5, y: 4.5, w: 9, h: 0.6,
      fontSize: 16, color: tc.accent, fontFace: "Calibri",
      bold: true, align: "right",
    });
  }
  // Bottom accent line
  slide_.addShape("rect", {
    x: 3.5, y: 5.2, w: 3, h: 0.05,
    fill: { color: tc.accent }, line: { color: tc.accent },
  });
}

function buildSectionDividerSlide(
  slide_: ReturnType<InstanceType<typeof pptxgen>["addSlide"]>,
  s: Slide,
  tc: ThemeColors
) {
  // Full background color
  slide_.addShape("rect", {
    x: 0, y: 0, w: "100%", h: "100%",
    fill: { color: tc.headerBg }, line: { color: tc.headerBg },
  });
  slide_.addText(s.title, {
    x: 0.5, y: 1.5, w: 9, h: 2.5,
    fontSize: 44, bold: true, color: tc.headerText,
    fontFace: "Calibri", align: "center", valign: "middle",
  });
  if (s.subtitle) {
    slide_.addText(s.subtitle, {
      x: 0.5, y: 4.0, w: 9, h: 0.8,
      fontSize: 20, color: tc.headerText, fontFace: "Calibri",
      align: "center", valign: "top",
    });
  }
}

function buildAgendaSlide(
  slide_: ReturnType<InstanceType<typeof pptxgen>["addSlide"]>,
  s: Slide,
  tc: ThemeColors
) {
  addHeaderBar(slide_, s.title, tc);
  const items = (s.bullets ?? []).map((b, i) => ({
    text: `${i + 1}.  ${b}`,
    options: { breakLine: true, color: tc.bodyText, fontFace: "Calibri" },
  }));
  if (items.length > 0) {
    slide_.addText(items, {
      x: 0.5, y: 1.0, w: 9, h: 4.3,
      fontSize: 19, valign: "top",
    });
  }
}

function buildClosingSlide(
  slide_: ReturnType<InstanceType<typeof pptxgen>["addSlide"]>,
  s: Slide,
  tc: ThemeColors
) {
  slide_.addShape("rect", {
    x: 0, y: 0, w: "100%", h: 2.5,
    fill: { color: tc.headerBg }, line: { color: tc.headerBg },
  });
  slide_.addText(s.title, {
    x: 0.5, y: 0.4, w: 9, h: 1.8,
    fontSize: 38, bold: true, color: tc.headerText,
    fontFace: "Calibri", align: "center", valign: "middle",
  });
  slide_.addShape("rect", {
    x: 3.5, y: 2.6, w: 3, h: 0.05,
    fill: { color: tc.accent }, line: { color: tc.accent },
  });
  if (s.subtitle) {
    slide_.addText(s.subtitle, {
      x: 0.5, y: 2.8, w: 9, h: 1.2,
      fontSize: 22, color: tc.bodyText, fontFace: "Calibri",
      align: "center",
    });
  }
  if (s.bullets && s.bullets.length > 0) {
    const items = s.bullets.map((b) => ({
      text: b,
      options: { bullet: true, breakLine: true, color: tc.bodyText },
    }));
    slide_.addText(items, {
      x: 1.5, y: 4.0, w: 7, h: 1.2,
      fontSize: 16, fontFace: "Calibri", align: "center", valign: "top",
    });
  }
}

// ─── Main PPTX builder ────────────────────────────────────────────────────────
export async function buildAndDownloadPptx(
  doc: DocumentOutput,
  images: InputImage[]
): Promise<void> {
  const prs = new pptxgen();
  const tc = THEMES[doc.theme] ?? THEMES.corporate_blue;

  prs.layout = "LAYOUT_16x9";
  prs.author = doc.author || "AI Doc Generator";
  prs.title = doc.title;

  const slides = doc.slides ?? [];

  for (const s of slides) {
    const slide_ = prs.addSlide();
    slide_.background = { color: tc.background };

    switch (s.layout) {
      case "title":
        buildTitleSlide(prs, slide_, s, tc);
        break;
      case "bullets":
        buildBulletsSlide(slide_, s, tc);
        break;
      case "two_column":
        buildTwoColumnSlide(slide_, s, tc);
        break;
      case "image_caption":
        buildImageCaptionSlide(slide_, s, tc, images);
        break;
      case "table":
        buildTableSlide(slide_, s, tc);
        break;
      case "quote":
        buildQuoteSlide(slide_, s, tc);
        break;
      case "section_divider":
        buildSectionDividerSlide(slide_, s, tc);
        break;
      case "agenda":
        buildAgendaSlide(slide_, s, tc);
        break;
      case "closing":
        buildClosingSlide(slide_, s, tc);
        break;
      default:
        buildBulletsSlide(slide_, s, tc);
    }

    // Add speaker notes
    if (s.speaker_notes) {
      slide_.addNotes(s.speaker_notes);
    }
  }

  const dateStr = new Date().toISOString().slice(0, 10).replace(/-/g, "");
  const slug = doc.title.replace(/[^a-z0-9]/gi, "_").slice(0, 30);
  const fileName = `${slug}_Presentation_${dateStr}.pptx`;

  const blob = (await prs.write({ outputType: "blob" })) as Blob;
  saveAs(blob, fileName);
}
