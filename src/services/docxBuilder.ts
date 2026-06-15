import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
  AlignmentType,
  Table,
  TableRow,
  TableCell,
  WidthType,
  BorderStyle,
  ImageRun,
  Header,
  Footer,
  PageNumber,
  UnderlineType,
  TableOfContents,
  StyleLevel,
} from "docx";
import { saveAs } from "file-saver";
import type { DocumentOutput, Section, InputImage } from "../types/document";

// ─── Helpers ─────────────────────────────────────────────────────────────────
function heading(text: string, level: 1 | 2 | 3): Paragraph {
  const levelMap: Record<number, (typeof HeadingLevel)[keyof typeof HeadingLevel]> = {
    1: HeadingLevel.HEADING_1,
    2: HeadingLevel.HEADING_2,
    3: HeadingLevel.HEADING_3,
  };
  return new Paragraph({ text, heading: levelMap[level] ?? HeadingLevel.HEADING_1 });
}

function body(text: string): Paragraph {
  return new Paragraph({
    children: [new TextRun({ text, size: 24, font: "Calibri" })],
    spacing: { after: 160 },
  });
}

function bullet(text: string, level = 0): Paragraph {
  return new Paragraph({
    children: [new TextRun({ text, size: 22, font: "Calibri" })],
    bullet: { level },
    spacing: { after: 80 },
  });
}

function pageBreak(): Paragraph {
  return new Paragraph({ pageBreakBefore: true, children: [] });
}

function buildTable(headers: string[], rows: string[][]): Table {
  const headerRow = new TableRow({
    children: headers.map(
      (h) =>
        new TableCell({
          children: [
            new Paragraph({
              children: [new TextRun({ text: h, bold: true, color: "FFFFFF", font: "Calibri", size: 22 })],
              alignment: AlignmentType.CENTER,
            }),
          ],
          shading: { fill: "1565C0" },
          borders: {
            top: { style: BorderStyle.SINGLE, size: 1, color: "FFFFFF" },
            bottom: { style: BorderStyle.SINGLE, size: 1, color: "FFFFFF" },
            left: { style: BorderStyle.SINGLE, size: 1, color: "FFFFFF" },
            right: { style: BorderStyle.SINGLE, size: 1, color: "FFFFFF" },
          },
        })
    ),
    tableHeader: true,
  });

  const dataRows = rows.map(
    (row, i) =>
      new TableRow({
        children: row.map(
          (cell) =>
            new TableCell({
              children: [
                new Paragraph({
                  children: [new TextRun({ text: cell, font: "Calibri", size: 20 })],
                }),
              ],
              shading: { fill: i % 2 === 0 ? "E3F2FD" : "FFFFFF" },
              borders: {
                top: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
                bottom: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
                left: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
                right: { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" },
              },
            })
        ),
      })
  );

  return new Table({
    rows: [headerRow, ...dataRows],
    width: { size: 100, type: WidthType.PERCENTAGE },
  });
}

async function buildImageParagraph(
  img: InputImage,
  caption?: string
): Promise<Paragraph[]> {
  const binaryStr = atob(img.base64);
  const bytes = new Uint8Array(binaryStr.length);
  for (let i = 0; i < binaryStr.length; i++) {
    bytes[i] = binaryStr.charCodeAt(i);
  }

  const paragraphs: Paragraph[] = [
    new Paragraph({
      children: [
        new ImageRun({
          data: bytes,
          transformation: { width: 500, height: 280 },
          type: img.mimeType === "image/jpeg" ? "jpg" : "png",
        }),
      ],
      alignment: AlignmentType.CENTER,
      spacing: { after: 120 },
    }),
  ];

  if (caption) {
    paragraphs.push(
      new Paragraph({
        children: [
          new TextRun({
            text: caption,
            italics: true,
            size: 18,
            color: "666666",
            font: "Calibri",
          }),
        ],
        alignment: AlignmentType.CENTER,
        spacing: { after: 200 },
      })
    );
  }

  return paragraphs;
}

// ─── Build document children from sections ────────────────────────────────────
async function buildSectionChildren(
  sections: Section[],
  images: InputImage[]
): Promise<(Paragraph | Table)[]> {
  const children: (Paragraph | Table)[] = [];

  for (const sec of sections) {
    children.push(heading(sec.heading, sec.level));

    for (const para of sec.paragraphs) {
      children.push(body(para));
    }

    if (sec.bullets) {
      for (const b of sec.bullets) {
        children.push(bullet(b));
      }
      children.push(new Paragraph({ children: [], spacing: { after: 200 } }));
    }

    if (sec.table) {
      children.push(buildTable(sec.table.headers, sec.table.rows));
      children.push(new Paragraph({ children: [], spacing: { after: 200 } }));
    }

    if (sec.image_index !== undefined) {
      const img = images[sec.image_index];
      if (img) {
        const imgParas = await buildImageParagraph(img, img.caption);
        children.push(...imgParas);
      }
    }
  }

  return children;
}

// ─── Main DOCX builder ────────────────────────────────────────────────────────
export async function buildAndDownloadDocx(
  doc: DocumentOutput,
  images: InputImage[]
): Promise<void> {
  const today = new Date().toLocaleDateString("en-US", {
    year: "numeric",
    month: "long",
    day: "numeric",
  });

  // Cover page
  const coverChildren: (Paragraph | Table)[] = [
    new Paragraph({ children: [], spacing: { after: 2000 } }),
    new Paragraph({
      children: [
        new TextRun({
          text: doc.title,
          bold: true,
          size: 56,
          font: "Calibri",
          color: "1565C0",
        }),
      ],
      alignment: AlignmentType.CENTER,
      spacing: { after: 400 },
    }),
    new Paragraph({
      children: [
        new TextRun({
          text: doc.author || "AI Doc Generator",
          size: 28,
          font: "Calibri",
          color: "555555",
        }),
      ],
      alignment: AlignmentType.CENTER,
      spacing: { after: 200 },
    }),
    new Paragraph({
      children: [
        new TextRun({
          text: today,
          size: 24,
          font: "Calibri",
          color: "888888",
        }),
      ],
      alignment: AlignmentType.CENTER,
      spacing: { after: 200 },
    }),
    new Paragraph({
      border: {
        bottom: { style: BorderStyle.SINGLE, size: 20, color: "1565C0" },
      },
      children: [],
      spacing: { after: 400 },
    }),
    pageBreak(),
  ];

  // TOC page
  const tocChildren: (Paragraph | Table)[] = [];
  try {
    tocChildren.push(
      new TableOfContents("Table of Contents", {
        hyperlink: true,
        headingStyleRange: "1-3",
        stylesWithLevels: [
          new StyleLevel("Heading1", 1),
          new StyleLevel("Heading2", 2),
          new StyleLevel("Heading3", 3),
        ],
      }) as unknown as Paragraph
    );
  } catch {
    // Fallback if TOC isn't available
    tocChildren.push(heading("Table of Contents", 1));
  }
  tocChildren.push(pageBreak());

  // Body content from LLM sections
  const sections = doc.sections ?? [];
  const bodyChildren = await buildSectionChildren(sections, images);

  // Header
  const docHeader = new Header({
    children: [
      new Paragraph({
        children: [
          new TextRun({
            text: doc.title,
            font: "Calibri",
            size: 18,
            color: "888888",
          }),
        ],
        alignment: AlignmentType.RIGHT,
        border: {
          bottom: { style: BorderStyle.SINGLE, size: 6, color: "1565C0" },
        },
      }),
    ],
  });

  // Footer
  const docFooter = new Footer({
    children: [
      new Paragraph({
        children: [
          new TextRun({ text: `${doc.author || "AI Doc Generator"}  |  `, font: "Calibri", size: 16, color: "888888" }),
          new TextRun({
            children: ["Page ", PageNumber.CURRENT, " of ", PageNumber.TOTAL_PAGES],
            font: "Calibri",
            size: 16,
            color: "888888",
          }),
        ],
        alignment: AlignmentType.CENTER,
        border: {
          top: { style: BorderStyle.SINGLE, size: 6, color: "1565C0" },
        },
      }),
    ],
  });

  const document = new Document({
    styles: {
      default: {
        document: {
          run: { font: "Calibri", size: 24 },
        },
        heading1: {
          run: { font: "Calibri", bold: true, size: 32, color: "1565C0" },
          paragraph: { spacing: { before: 400, after: 160 } },
        },
        heading2: {
          run: { font: "Calibri", bold: true, size: 26, color: "1A3A5C" },
          paragraph: { spacing: { before: 300, after: 120 } },
        },
        heading3: {
          run: { font: "Calibri", bold: true, size: 24, color: "2D5986" },
          paragraph: { spacing: { before: 200, after: 80 } },
        },
      },
    },
    sections: [
      {
        headers: { default: docHeader },
        footers: { default: docFooter },
        children: [...coverChildren, ...tocChildren, ...bodyChildren],
      },
    ],
  });

  const blob = await Packer.toBlob(document);
  const dateStr = new Date().toISOString().slice(0, 10).replace(/-/g, "");
  const slug = doc.title.replace(/[^a-z0-9]/gi, "_").slice(0, 30);
  saveAs(blob, `${slug}_Document_${dateStr}.docx`);
}
