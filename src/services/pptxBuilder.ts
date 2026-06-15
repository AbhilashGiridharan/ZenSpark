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
  accent2: string;
  surface: string;
  tableHeaderBg: string;
  tableHeaderText: string;
  tableRowBg: string;
  divider: string;
}

const THEMES: Record<ThemeOption, ThemeColors> = {
  corporate_blue: {
    background: "F4F7FF",
    headerBg: "0D47A1",
    headerText: "FFFFFF",
    bodyText: "1A2744",
    accent: "1565C0",
    accent2: "42A5F5",
    surface: "DBEAFE",
    tableHeaderBg: "0D47A1",
    tableHeaderText: "FFFFFF",
    tableRowBg: "EFF6FF",
    divider: "BFDBFE",
  },
  dark_tech: {
    background: "0D1117",
    headerBg: "161B22",
    headerText: "58A6FF",
    bodyText: "C9D1D9",
    accent: "58A6FF",
    accent2: "3FB950",
    surface: "1C2128",
    tableHeaderBg: "21262D",
    tableHeaderText: "58A6FF",
    tableRowBg: "161B22",
    divider: "30363D",
  },
  minimal_white: {
    background: "FFFFFF",
    headerBg: "1A1A2E",
    headerText: "FFFFFF",
    bodyText: "2D3748",
    accent: "E53E3E",
    accent2: "6B46C1",
    surface: "F7F7F7",
    tableHeaderBg: "1A1A2E",
    tableHeaderText: "FFFFFF",
    tableRowBg: "F9F9F9",
    divider: "E2E8F0",
  },
  green_growth: {
    background: "F6FBF6",
    headerBg: "1B5E20",
    headerText: "FFFFFF",
    bodyText: "1B3A1E",
    accent: "2E7D32",
    accent2: "66BB6A",
    surface: "DCFCE7",
    tableHeaderBg: "1B5E20",
    tableHeaderText: "FFFFFF",
    tableRowBg: "ECFDF5",
    divider: "BBF7D0",
  },
};

type PptxSlide = ReturnType<InstanceType<typeof pptxgen>["addSlide"]>;

// ─── Visual helpers ───────────────────────────────────────────────────────────
/** Circle badge with centered text — numbered bullets, agenda items */
function addBadge(
  slide: PptxSlide,
  x: number, y: number, size: number,
  label: string | number,
  bgColor: string,
  textColor = "FFFFFF",
  fontSize = 11
) {
  slide.addShape("ellipse", {
    x, y, w: size, h: size,
    fill: { color: bgColor },
    line: { color: bgColor },
  });
  slide.addText(String(label), {
    x, y, w: size, h: size,
    fontSize, bold: true, color: textColor,
    fontFace: "Calibri", align: "center", valign: "middle",
  });
}

// ─── Header bar ───────────────────────────────────────────────────────────────
function addHeaderBar(
  slide: PptxSlide,
  title: string,
  slideNum: number,
  tc: ThemeColors
) {
  const H = 0.72;
  slide.addShape("rect", {
    x: 0, y: 0, w: "100%", h: H,
    fill: { color: tc.headerBg },
    line: { color: tc.headerBg },
  });
  // Left accent stripe
  slide.addShape("rect", {
    x: 0, y: 0, w: 0.12, h: H,
    fill: { color: tc.accent2 },
    line: { color: tc.accent2 },
  });
  slide.addText(title, {
    x: 0.25, y: 0, w: 8.85, h: H,
    fontSize: 20, bold: true, color: tc.headerText,
    fontFace: "Calibri", valign: "middle",
  });
  // Slide number badge
  addBadge(slide, 9.42, (H - 0.36) / 2, 0.36, slideNum, tc.accent2, "FFFFFF", 10);
}

// ─── Title slide ──────────────────────────────────────────────────────────────
function buildTitleSlide(
  prs: InstanceType<typeof pptxgen>,
  slide_: PptxSlide,
  s: Slide,
  tc: ThemeColors
) {
  void prs;
  // Top colored block
  slide_.addShape("rect", {
    x: 0, y: 0, w: "100%", h: 4.1,
    fill: { color: tc.headerBg },
    line: { color: tc.headerBg },
  });
  // Decorative large circle (top-right, semi-transparent)
  slide_.addShape("ellipse", {
    x: 6.9, y: -1.3, w: 4.4, h: 4.4,
    fill: { color: tc.accent2, transparency: 80 },
    line: { color: tc.accent2, transparency: 80 },
  });
  // Smaller circle (mid-right)
  slide_.addShape("ellipse", {
    x: 8.5, y: 2.3, w: 1.7, h: 1.7,
    fill: { color: tc.accent, transparency: 65 },
    line: { color: tc.accent, transparency: 65 },
  });
  // Bottom accent bars
  slide_.addShape("rect", {
    x: 0, y: 7.1, w: 3.2, h: 0.4,
    fill: { color: tc.accent2 }, line: { color: tc.accent2 },
  });
  slide_.addShape("rect", {
    x: 9.0, y: 7.1, w: 1.0, h: 0.4,
    fill: { color: tc.accent }, line: { color: tc.accent },
  });
  // Title text
  slide_.addText(s.title, {
    x: 0.55, y: 0.35, w: 7.0, h: 3.3,
    fontSize: 38, bold: true, color: tc.headerText,
    fontFace: "Calibri", valign: "middle", wrap: true,
  });
  // Accent divider line
  slide_.addShape("rect", {
    x: 0.55, y: 4.28, w: 2.8, h: 0.07,
    fill: { color: tc.accent2 }, line: { color: tc.accent2 },
  });
  if (s.subtitle) {
    slide_.addText(s.subtitle, {
      x: 0.55, y: 4.5, w: 9.0, h: 1.7,
      fontSize: 20, color: tc.bodyText,
      fontFace: "Calibri", valign: "top", wrap: true,
    });
  }
}

// ─── Bullets slide — numbered badge circles ───────────────────────────────────
function buildBulletsSlide(
  slide_: PptxSlide,
  s: Slide,
  tc: ThemeColors
) {
  addHeaderBar(slide_, s.title, s.slide_number, tc);
  const bullets = s.bullets ?? [];
  if (bullets.length === 0) return;

  const CONTENT_TOP = 0.88;
  const available = 7.25 - CONTENT_TOP;
  const rowH = Math.min(0.82, available / bullets.length);
  const fontSize = rowH >= 0.70 ? 17 : rowH >= 0.55 ? 15 : 13;
  const badgeSize = Math.min(0.30, rowH * 0.62);

  bullets.forEach((b, i) => {
    const rowY = CONTENT_TOP + i * rowH;
    const badgeY = rowY + rowH / 2 - badgeSize / 2;
    const badgeColor = i % 2 === 0 ? tc.accent : tc.accent2;

    if (i % 2 === 0) {
      slide_.addShape("rect", {
        x: 0.38, y: rowY, w: 9.3, h: rowH,
        fill: { color: tc.surface }, line: { color: tc.surface },
      });
    }
    addBadge(slide_, 0.43, badgeY, badgeSize, i + 1, badgeColor, "FFFFFF",
      Math.max(9, Math.floor(badgeSize * 32)));
    slide_.addText(b, {
      x: 0.86, y: rowY, w: 9.0, h: rowH,
      fontSize, color: tc.bodyText, fontFace: "Calibri", valign: "middle",
    });
  });
}

// ─── Two-column slide — card style ───────────────────────────────────────────
function buildTwoColumnSlide(
  slide_: PptxSlide,
  s: Slide,
  tc: ThemeColors
) {
  addHeaderBar(slide_, s.title, s.slide_number, tc);

  slide_.addShape("rect", {
    x: 0.28, y: 0.86, w: 4.5, h: 5.5,
    fill: { color: tc.surface }, line: { color: tc.divider },
  });
  slide_.addShape("rect", {
    x: 5.22, y: 0.86, w: 4.5, h: 5.5,
    fill: { color: tc.surface }, line: { color: tc.divider },
  });
  // Accent top stripes
  slide_.addShape("rect", {
    x: 0.28, y: 0.86, w: 4.5, h: 0.09,
    fill: { color: tc.accent }, line: { color: tc.accent },
  });
  slide_.addShape("rect", {
    x: 5.22, y: 0.86, w: 4.5, h: 0.09,
    fill: { color: tc.accent2 }, line: { color: tc.accent2 },
  });

  if (s.left_title) {
    slide_.addText(s.left_title, {
      x: 0.42, y: 0.97, w: 4.2, h: 0.44,
      fontSize: 14, bold: true, color: tc.accent, fontFace: "Calibri", valign: "middle",
    });
  }
  if (s.right_title) {
    slide_.addText(s.right_title, {
      x: 5.36, y: 0.97, w: 4.2, h: 0.44,
      fontSize: 14, bold: true, color: tc.accent2, fontFace: "Calibri", valign: "middle",
    });
  }

  const leftY = s.left_title ? 1.48 : 1.02;
  const rightY = s.right_title ? 1.48 : 1.02;

  const leftItems = (s.left_column ?? []).map((b) => ({
    text: b,
    options: { bullet: { indent: 10 }, breakLine: true, color: tc.bodyText },
  }));
  if (leftItems.length > 0) {
    slide_.addText(leftItems, { x: 0.42, y: leftY, w: 4.18, h: 5.0, fontSize: 15, fontFace: "Calibri", valign: "top" });
  }

  const rightItems = (s.right_column ?? []).map((b) => ({
    text: b,
    options: { bullet: { indent: 10 }, breakLine: true, color: tc.bodyText },
  }));
  if (rightItems.length > 0) {
    slide_.addText(rightItems, { x: 5.36, y: rightY, w: 4.18, h: 5.0, fontSize: 15, fontFace: "Calibri", valign: "top" });
  }
}

// ─── Image caption slide ──────────────────────────────────────────────────────
function buildImageCaptionSlide(
  slide_: PptxSlide,
  s: Slide,
  tc: ThemeColors,
  images: InputImage[]
) {
  addHeaderBar(slide_, s.title, s.slide_number, tc);
  const img = s.image_index !== undefined ? images[s.image_index] : undefined;
  if (img) {
    slide_.addImage({ data: img.preview, x: 0.4, y: 0.9, w: 5.6, h: 4.2 });
    const captionText = img.caption || s.subtitle || "";
    if (captionText) {
      slide_.addText(captionText, { x: 6.2, y: 0.9, w: 3.5, h: 4.2, fontSize: 16, color: tc.bodyText, fontFace: "Calibri", valign: "top" });
    }
  } else if (s.subtitle) {
    slide_.addText(s.subtitle, { x: 0.4, y: 0.9, w: 9.2, h: 5.5, fontSize: 18, color: tc.bodyText, fontFace: "Calibri", valign: "top" });
  }
}

// ─── Table slide ──────────────────────────────────────────────────────────────
function buildTableSlide(
  slide_: PptxSlide,
  s: Slide,
  tc: ThemeColors
) {
  addHeaderBar(slide_, s.title, s.slide_number, tc);
  if (!s.table) return;

  const headerRow = s.table.headers.map((h) => ({
    text: h,
    options: {
      bold: true, color: tc.tableHeaderText,
      fill: { color: tc.tableHeaderBg },
      fontFace: "Calibri", fontSize: 14, align: "center" as const,
    },
  }));
  const dataRows = s.table.rows.map((row, i) =>
    row.map((cell) => ({
      text: cell,
      options: {
        fill: { color: i % 2 === 0 ? tc.tableRowBg : "FFFFFF" },
        color: tc.bodyText, fontFace: "Calibri", fontSize: 13,
      },
    }))
  );
  slide_.addTable([headerRow, ...dataRows], {
    x: 0.4, y: 0.88, w: 9.2,
    border: { type: "solid", color: tc.divider, pt: 1 },
    rowH: 0.45,
  });
}

// ─── Quote slide ──────────────────────────────────────────────────────────────
function buildQuoteSlide(
  slide_: PptxSlide,
  s: Slide,
  tc: ThemeColors
) {
  // Left accent bar
  slide_.addShape("rect", {
    x: 0, y: 0, w: 0.45, h: "100%",
    fill: { color: tc.accent }, line: { color: tc.accent },
  });
  slide_.addText("\u201C", {
    x: 0.7, y: 0.1, w: 1.5, h: 1.5,
    fontSize: 96, color: tc.divider, fontFace: "Georgia", bold: true,
  });
  slide_.addText(s.quote ?? s.title, {
    x: 0.7, y: 1.2, w: 8.8, h: 3.5,
    fontSize: 26, color: tc.bodyText, fontFace: "Georgia",
    italic: true, align: "center", valign: "middle",
  });
  if (s.attribution) {
    slide_.addShape("rect", {
      x: 1.0, y: 5.1, w: 8.0, h: 0.05,
      fill: { color: tc.divider }, line: { color: tc.divider },
    });
    slide_.addText("\u2014 " + s.attribution, {
      x: 1.0, y: 5.25, w: 8.0, h: 0.55,
      fontSize: 16, color: tc.accent, fontFace: "Calibri", bold: true, align: "right",
    });
  }
}

// ─── Section divider slide ────────────────────────────────────────────────────
function buildSectionDividerSlide(
  slide_: PptxSlide,
  s: Slide,
  tc: ThemeColors
) {
  slide_.addShape("rect", {
    x: 0, y: 0, w: "100%", h: "100%",
    fill: { color: tc.headerBg }, line: { color: tc.headerBg },
  });
  // Large decorative number in background
  slide_.addText(String(s.slide_number).padStart(2, "0"), {
    x: 5.2, y: -0.6, w: 4.8, h: 5.0,
    fontSize: 190, bold: true, color: tc.accent2,
    transparency: 85,
    fontFace: "Calibri", align: "left", valign: "top",
  });
  // Left accent stripe
  slide_.addShape("rect", {
    x: 0, y: 0, w: 0.5, h: "100%",
    fill: { color: tc.accent2 }, line: { color: tc.accent2 },
  });
  // Bottom strip
  slide_.addShape("rect", {
    x: 0, y: 7.1, w: "100%", h: 0.4,
    fill: { color: tc.accent }, line: { color: tc.accent },
  });
  slide_.addText(s.title, {
    x: 0.75, y: 1.5, w: 8.8, h: 3.0,
    fontSize: 44, bold: true, color: tc.headerText,
    fontFace: "Calibri", align: "left", valign: "middle",
  });
  if (s.subtitle) {
    slide_.addText(s.subtitle, {
      x: 0.75, y: 4.6, w: 8.8, h: 0.9,
      fontSize: 20, color: tc.accent2, fontFace: "Calibri", align: "left",
    });
  }
}

// ─── Agenda slide — numbered circles ─────────────────────────────────────────
function buildAgendaSlide(
  slide_: PptxSlide,
  s: Slide,
  tc: ThemeColors
) {
  addHeaderBar(slide_, s.title, s.slide_number, tc);
  const items = s.bullets ?? [];
  if (items.length === 0) return;

  const CONTENT_TOP = 0.90;
  const available = 7.22 - CONTENT_TOP;
  const rowH = Math.min(0.88, available / items.length);
  const badgeSize = Math.min(0.40, rowH * 0.65);

  items.forEach((item, i) => {
    const rowY = CONTENT_TOP + i * rowH;
    const badgeY = rowY + rowH / 2 - badgeSize / 2;
    const badgeColor = i % 2 === 0 ? tc.accent : tc.accent2;

    if (i % 2 === 0) {
      slide_.addShape("rect", {
        x: 0.38, y: rowY, w: 9.3, h: rowH,
        fill: { color: tc.surface }, line: { color: tc.surface },
      });
    }
    addBadge(slide_, 0.50, badgeY, badgeSize, i + 1, badgeColor, "FFFFFF",
      Math.max(10, Math.floor(badgeSize * 26)));
    slide_.addText(item, {
      x: 1.05, y: rowY, w: 8.7, h: rowH,
      fontSize: rowH >= 0.75 ? 19 : 17,
      color: tc.bodyText, fontFace: "Calibri", valign: "middle",
    });
  });
}

// ─── Stats slide — metric cards ───────────────────────────────────────────────
function buildStatsSlide(
  slide_: PptxSlide,
  s: Slide,
  tc: ThemeColors
) {
  addHeaderBar(slide_, s.title, s.slide_number, tc);
  const cards = s.stat_cards ?? [];
  if (cards.length === 0) {
    const items = (s.bullets ?? []).map((b) => ({
      text: b,
      options: { bullet: { indent: 10 }, breakLine: true, color: tc.bodyText },
    }));
    if (items.length > 0) {
      slide_.addText(items, { x: 0.4, y: 0.9, w: 9.2, h: 5.5, fontSize: 18, fontFace: "Calibri", valign: "top" });
    }
    return;
  }

  type CardPos = { x: number; y: number; w: number; h: number };
  const n = Math.min(cards.length, 4);
  let positions: CardPos[] = [];

  if (n === 1) {
    positions = [{ x: 2.5, y: 1.2, w: 5.0, h: 4.6 }];
  } else if (n === 2) {
    positions = [{ x: 0.40, y: 1.0, w: 4.35, h: 5.3 }, { x: 5.25, y: 1.0, w: 4.35, h: 5.3 }];
  } else if (n === 3) {
    positions = [
      { x: 0.22, y: 1.0, w: 3.0, h: 5.3 },
      { x: 3.50, y: 1.0, w: 3.0, h: 5.3 },
      { x: 6.78, y: 1.0, w: 3.0, h: 5.3 },
    ];
  } else {
    positions = [
      { x: 0.35, y: 0.92, w: 4.38, h: 2.85 },
      { x: 5.27, y: 0.92, w: 4.38, h: 2.85 },
      { x: 0.35, y: 3.95, w: 4.38, h: 2.85 },
      { x: 5.27, y: 3.95, w: 4.38, h: 2.85 },
    ];
  }

  const accentCycle = [tc.accent, tc.accent2, tc.headerBg, tc.accent2];

  cards.slice(0, 4).forEach((card, i) => {
    const { x, y, w, h } = positions[i];
    const aColor = accentCycle[i % accentCycle.length];

    slide_.addShape("rect", { x, y, w, h, fill: { color: tc.surface }, line: { color: tc.divider } });
    slide_.addShape("rect", { x, y, w, h: 0.12, fill: { color: aColor }, line: { color: aColor } });

    const iconH = card.icon ? 0.6 : 0;
    if (card.icon) {
      slide_.addText(card.icon, { x: x + 0.15, y: y + 0.2, w: w - 0.3, h: 0.6, fontSize: 28, align: "center", valign: "middle" });
    }

    const valFontSize = h > 3.0 ? 52 : 38;
    slide_.addText(card.value, {
      x: x + 0.15, y: y + 0.18 + iconH, w: w - 0.3, h: h * 0.52,
      fontSize: valFontSize, bold: true, color: aColor,
      fontFace: "Calibri", align: "center", valign: "middle",
    });
    slide_.addText(card.label, {
      x: x + 0.15, y: y + h * 0.64, w: w - 0.3, h: h * 0.30,
      fontSize: n >= 4 ? 13 : 16, color: tc.bodyText,
      fontFace: "Calibri", align: "center", valign: "top", wrap: true,
    });
  });
}

// ─── Closing slide ────────────────────────────────────────────────────────────
function buildClosingSlide(
  slide_: PptxSlide,
  s: Slide,
  tc: ThemeColors
) {
  slide_.addShape("rect", {
    x: 0, y: 0, w: "100%", h: 2.6,
    fill: { color: tc.headerBg }, line: { color: tc.headerBg },
  });
  slide_.addShape("ellipse", {
    x: 7.5, y: -0.9, w: 3.2, h: 3.2,
    fill: { color: tc.accent2, transparency: 78 },
    line: { color: tc.accent2, transparency: 78 },
  });
  slide_.addShape("ellipse", {
    x: -0.9, y: 1.6, w: 2.0, h: 2.0,
    fill: { color: tc.accent, transparency: 70 },
    line: { color: tc.accent, transparency: 70 },
  });
  slide_.addText(s.title, {
    x: 0.5, y: 0.3, w: 8.8, h: 2.0,
    fontSize: 38, bold: true, color: tc.headerText,
    fontFace: "Calibri", align: "center", valign: "middle",
  });
  slide_.addShape("rect", {
    x: 3.5, y: 2.7, w: 3.0, h: 0.07,
    fill: { color: tc.accent2 }, line: { color: tc.accent2 },
  });
  if (s.subtitle) {
    slide_.addText(s.subtitle, {
      x: 0.5, y: 2.9, w: 9.0, h: 1.4,
      fontSize: 22, color: tc.bodyText, fontFace: "Calibri", align: "center", wrap: true,
    });
  }
  if (s.bullets && s.bullets.length > 0) {
    const items = s.bullets.map((b) => ({
      text: b,
      options: { bullet: true, breakLine: true, color: tc.bodyText },
    }));
    slide_.addText(items, { x: 2.0, y: 4.4, w: 6.0, h: 2.0, fontSize: 16, fontFace: "Calibri", align: "center", valign: "top" });
  }
  slide_.addShape("rect", {
    x: 0, y: 7.1, w: "100%", h: 0.4,
    fill: { color: tc.accent }, line: { color: tc.accent },
  });
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
      case "title":       buildTitleSlide(prs, slide_, s, tc); break;
      case "bullets":     buildBulletsSlide(slide_, s, tc); break;
      case "two_column":  buildTwoColumnSlide(slide_, s, tc); break;
      case "image_caption": buildImageCaptionSlide(slide_, s, tc, images); break;
      case "table":       buildTableSlide(slide_, s, tc); break;
      case "quote":       buildQuoteSlide(slide_, s, tc); break;
      case "section_divider": buildSectionDividerSlide(slide_, s, tc); break;
      case "agenda":      buildAgendaSlide(slide_, s, tc); break;
      case "stats":       buildStatsSlide(slide_, s, tc); break;
      case "closing":     buildClosingSlide(slide_, s, tc); break;
      default:            buildBulletsSlide(slide_, s, tc);
    }

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
