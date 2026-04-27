// PDF Resource Generation Helpers
// Theme-agnostic utilities for creating printable A4 worksheets, answer keys,
// graphic organisers, and other companion resources alongside PPTX slide decks.
//
// Uses pdfkit. All colours passed as 6-char hex (no #). Fonts prefer a
// cross-platform sans family with broad Unicode support (□, Δ, ×, ÷, ≥, ≠
// etc.) and fall back to built-in Helvetica if no complete family is found.

"use strict";

const PDFDocument = require("pdfkit");
const fs = require("fs");
const path = require("path");
const {
  isStructuredMockupSpec,
  lightenHex,
  normalizeStructuredMockup,
} = require("./core/mockups");

function normaliseSessionNumber(sessionNumber) {
  const value = Number.parseInt(sessionNumber, 10);
  if (!Number.isInteger(value) || value < 1) {
    throw new Error("Session number must be a positive integer.");
  }
  return value;
}

function cleanResourceLabel(label) {
  const raw = String(label || "")
    .replace(/[_-]+/g, " ")
    .replace(/[<>:\"/\\\\|?*]/g, " ");
  return raw.replace(/\s+/g, " ").trim();
}

function getSessionResourceFolder(sessionNumber) {
  return `resources-session${normaliseSessionNumber(sessionNumber)}`;
}

function formatSessionResourceName(sessionNumber, label) {
  const session = normaliseSessionNumber(sessionNumber);
  const suffix = cleanResourceLabel(label);
  return suffix ? `Session ${session} ${suffix}` : `Session ${session} Resource`;
}

function formatSessionResourceFileName(sessionNumber, label, opts) {
  const o = opts || {};
  const extValue = o.ext || ".pdf";
  const ext = extValue.startsWith(".") ? extValue : `.${extValue}`;
  const folder = o.folder || getSessionResourceFolder(sessionNumber);
  const baseName = cleanResourceLabel(o.baseName || formatSessionResourceName(sessionNumber, label));
  return path.posix.join(folder, `${baseName}${ext}`);
}

function makeSessionResource(sessionNumber, label, description, opts) {
  const o = opts || {};
  return {
    name: cleanResourceLabel(o.name || formatSessionResourceName(sessionNumber, label)),
    fileName: o.fileName || formatSessionResourceFileName(sessionNumber, label, o),
    description: description || "",
  };
}

function resourceNameFromFileName(fileName) {
  if (!fileName) return "Resource";
  const baseName = path.basename(fileName, path.extname(fileName));
  return cleanResourceLabel(baseName) || "Resource";
}

function getPdfFontFamilyCandidates() {
  if (process.platform === "win32") {
    return [
      {
        regular: "C:/Windows/Fonts/arial.ttf",
        bold: "C:/Windows/Fonts/arialbd.ttf",
        italic: "C:/Windows/Fonts/ariali.ttf",
      },
      {
        regular: "C:/Windows/Fonts/calibri.ttf",
        bold: "C:/Windows/Fonts/calibrib.ttf",
        italic: "C:/Windows/Fonts/calibrii.ttf",
      },
    ];
  }

  if (process.platform === "darwin") {
    return [
      {
        regular: "/System/Library/Fonts/Supplemental/Arial.ttf",
        bold: "/System/Library/Fonts/Supplemental/Arial Bold.ttf",
        italic: "/System/Library/Fonts/Supplemental/Arial Italic.ttf",
      },
    ];
  }

  return [
    {
      regular: "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf",
      bold: "/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf",
      italic: "/usr/share/fonts/truetype/liberation2/LiberationSans-Italic.ttf",
    },
    {
      regular: "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
      bold: "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
      italic: "/usr/share/fonts/truetype/liberation/LiberationSans-Italic.ttf",
    },
    {
      regular: "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
      bold: "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
      italic: "/usr/share/fonts/truetype/dejavu/DejaVuSans-Oblique.ttf",
    },
  ];
}

function resolvePdfFontFamily() {
  return getPdfFontFamilyCandidates().find((family) =>
    [family.regular, family.bold, family.italic].every((filePath) => fs.existsSync(filePath))
  ) || null;
}

// ── Page constants (A4 in points: 595.28 x 841.89) ─────────────────────────

const PAGE = {
  W: 595.28,
  H: 841.89,
  MARGIN: 50,
  CONTENT_W: 595.28 - 2 * 50,  // 495.28
  CONTENT_H: 841.89 - 2 * 50,  // 741.89
};

const FOOTER_Y = PAGE.H - PAGE.MARGIN - 10;
const CONTENT_BOTTOM = FOOTER_Y - 10;
const STUDENT_RULE_COLOR = "#000000";
const STUDENT_RULE_WIDTH = 0.9;

// ── Colour helpers ──────────────────────────────────────────────────────────

/** Convert 6-char hex to "#RRGGBB" for pdfkit. */
function hex(color) {
  if (!color) return "#000000";
  return color.startsWith("#") ? color : "#" + color;
}

/** Parse 6-char hex (no #) to {r, g, b}. */
function hexToRgb(c) {
  const h = c.replace("#", "");
  return {
    r: parseInt(h.substring(0, 2), 16),
    g: parseInt(h.substring(2, 4), 16),
    b: parseInt(h.substring(4, 6), 16),
  };
}

/** Lighten a hex colour by mixing with white. Amount 0-1 (0=original, 1=white). */
function lighten(color, amount) {
  const { r, g, b } = hexToRgb(color);
  const nr = Math.round(r + (255 - r) * amount);
  const ng = Math.round(g + (255 - g) * amount);
  const nb = Math.round(b + (255 - b) * amount);
  return "#" + [nr, ng, nb].map(v => v.toString(16).padStart(2, "0")).join("");
}

function ensureBlockFits(doc, y, requiredHeight, opts) {
  const o = opts || {};
  const top = o.pageTop != null ? o.pageTop : PAGE.MARGIN;
  const bottom = o.pageBottom != null ? o.pageBottom : CONTENT_BOTTOM;
  const fullPageHeight = bottom - top;

  // If the block can fit on a fresh page but not in the remaining space,
  // move it before drawing so prompts and answer space stay together.
  if (requiredHeight <= fullPageHeight && y + requiredHeight > bottom) {
    doc.addPage();
    return top;
  }

  return y;
}

function measureBodyTextHeight(doc, text, opts) {
  const o = opts || {};
  const font = o.italic ? "Sans-Italic" : "Sans";
  return doc.fontSize(o.fontSize || 11).font(font).heightOfString(text, {
    width: o.width || PAGE.CONTENT_W,
  });
}

function measureProblemHeight(doc, prompt, opts) {
  const o = opts || {};
  const promptHeight = doc.fontSize(11).font("Sans").heightOfString(prompt, {
    width: PAGE.CONTENT_W - 28,
  });

  let totalHeight = Math.max(20, promptHeight) + 10;

  if (o.headers) {
    totalHeight += (o.hdrH || 26) + (o.valH || 40) + 10;
  }

  if (o.writeLines) {
    totalHeight += o.writeLines.length * 24;
  }

  return totalHeight + 12;
}

// ── Document creation ───────────────────────────────────────────────────────

/**
 * Create a new A4 PDF document.
 * @param {object} opts — { margin }
 * @returns {PDFDocument}
 */
function createPdf(opts) {
  const o = opts || {};
  const doc = new PDFDocument({
    size: "A4",
    margin: o.margin || PAGE.MARGIN,
    info: {
      Title: o.title || "Worksheet",
      Author: o.author || "Generated Resource",
    },
    bufferPages: true,
  });

  const fontFamily = resolvePdfFontFamily();
  try {
    if (fontFamily) {
      doc.registerFont("Sans", fontFamily.regular);
      doc.registerFont("Sans-Bold", fontFamily.bold);
      doc.registerFont("Sans-Italic", fontFamily.italic);
    }
  } catch (_) {
    // Fall back to built-in Helvetica if system fonts unavailable
  }

  return doc;
}

/**
 * Write a PDF document to a file. Returns a promise.
 * Ensures the parent directory exists.
 * @param {PDFDocument} doc
 * @param {string} filePath
 * @returns {Promise<void>}
 */
function writePdf(doc, filePath) {
  fs.mkdirSync(path.dirname(filePath), { recursive: true });

  return new Promise((resolve, reject) => {
    const stream = fs.createWriteStream(filePath);
    doc.pipe(stream);
    doc.end();
    stream.on("finish", resolve);
    stream.on("error", reject);
  });
}

// ── Page elements ───────────────────────────────────────────────────────────

/**
 * Add a worksheet header with coloured title bar, subtitle, and name/date fields.
 * @param {PDFDocument} doc
 * @param {string} title — main title
 * @param {object} opts — { subtitle, color, showNameDate, lessonInfo }
 */
function addPdfHeader(doc, title, opts) {
  const o = opts || {};
  const color = o.color || "1B3A6B";
  const x = PAGE.MARGIN;
  const w = PAGE.CONTENT_W;
  let y = PAGE.MARGIN;

  // Title bar
  doc.save();
  doc.roundedRect(x, y, w, 40, 4).fill(hex(color));
  doc.fontSize(18).font("Sans-Bold").fillColor("#FFFFFF");
  doc.text(title, x + 12, y + 10, { width: w - 24, align: "left" });
  doc.restore();
  y += 48;

  // Subtitle
  if (o.subtitle) {
    doc.fontSize(11).font("Sans").fillColor(hex("6B7280"));
    doc.text(o.subtitle, x, y, { width: w });
    y += 18;
  }

  // Lesson info line
  if (o.lessonInfo) {
    doc.fontSize(9).font("Sans").fillColor(hex("9CA3AF"));
    doc.text(o.lessonInfo, x, y, { width: w });
    y += 16;
  }

  // Name / Date line
  if (o.showNameDate !== false) {
    y += 4;
    doc.fontSize(11).font("Sans").fillColor("#000000");
    doc.text("Name: ", x, y, { continued: true });
    // Underline for name
    const nameLineX = x + 42;
    doc.moveTo(nameLineX, y + 14).lineTo(nameLineX + 200, y + 14)
      .strokeColor(STUDENT_RULE_COLOR).lineWidth(STUDENT_RULE_WIDTH).stroke();
    // Date field
    const dateX = x + 300;
    doc.text("Date: ", dateX, y);
    doc.moveTo(dateX + 38, y + 14).lineTo(dateX + w - 300, y + 14)
      .strokeColor(STUDENT_RULE_COLOR).lineWidth(STUDENT_RULE_WIDTH).stroke();
    y += 28;
  }

  return y; // Return current Y position for content below
}

/**
 * Add a section heading with a thin coloured left bar.
 * @param {PDFDocument} doc
 * @param {string} text
 * @param {number} y — vertical position
 * @param {object} opts — { color, fontSize }
 * @returns {number} y after the heading
 */
function addSectionHeading(doc, text, y, opts) {
  const o = opts || {};
  const color = o.color || "1B3A6B";
  const x = PAGE.MARGIN;
  const fontSize = o.fontSize || 13;
  y = ensureBlockFits(doc, y, fontSize + 14);

  // Left accent bar
  doc.save();
  doc.rect(x, y, 4, fontSize + 6).fill(hex(color));
  doc.restore();

  doc.fontSize(fontSize).font("Sans-Bold").fillColor(hex(color));
  doc.text(text, x + 12, y + 2, { width: PAGE.CONTENT_W - 12 });

  return y + fontSize + 14;
}

/**
 * Add instruction text (regular body text).
 * @param {PDFDocument} doc
 * @param {string} text
 * @param {number} y
 * @param {object} opts — { fontSize, color, italic }
 * @returns {number} y after text
 */
function addBodyText(doc, text, y, opts) {
  const o = opts || {};
  const font = o.italic ? "Sans-Italic" : "Sans";
  y = ensureBlockFits(doc, y, measureBodyTextHeight(doc, text, o) + 8);
  doc.fontSize(o.fontSize || 11).font(font).fillColor(hex(o.color || "2D3142"));
  doc.text(text, PAGE.MARGIN, y, { width: PAGE.CONTENT_W });
  return doc.y + 8;
}

/**
 * Add an empty place value chart for students to fill in.
 * @param {PDFDocument} doc
 * @param {number} y — top of chart
 * @param {string[]} headers — column headers (e.g., ["TTh","Th","H","T","O"])
 * @param {object} opts — { color, cellW, hdrH, valH, values, x }
 * @returns {{ y: number, bottomY: number }} positions for layout
 */
function addPvChartPdf(doc, y, headers, opts) {
  const o = opts || {};
  const n = headers.length;
  const cellW = o.cellW || 70;
  const hdrH = o.hdrH || 26;
  const valH = o.valH || 40;
  const color = o.color || "1B3A6B";
  const totalW = cellW * n;

  // Centre the chart horizontally unless x is explicitly provided
  const x = o.x != null
    ? o.x
    : PAGE.MARGIN + (PAGE.CONTENT_W - totalW) / 2;

  // Header row
  doc.save();
  headers.forEach((h, i) => {
    const cx = x + i * cellW;
    doc.rect(cx, y, cellW, hdrH).fill(hex(color));
    doc.fontSize(9).font("Sans-Bold").fillColor("#FFFFFF");
    doc.text(h, cx, y + 7, { width: cellW, align: "center" });
  });
  doc.restore();

  // Value row
  const values = o.values || [];
  doc.save();
  headers.forEach((_, i) => {
    const cx = x + i * cellW;
    doc.rect(cx, y + hdrH, cellW, valH)
      .lineWidth(STUDENT_RULE_WIDTH).strokeColor(STUDENT_RULE_COLOR).stroke();
    // Fill in values if provided (for answer keys)
    if (values[i] != null && values[i] !== "") {
      doc.fontSize(20).font("Sans-Bold").fillColor(hex("2D3142"));
      doc.text(String(values[i]), cx, y + hdrH + 8, { width: cellW, align: "center" });
    }
  });
  doc.restore();

  const bottomY = y + hdrH + valH;
  return { y, bottomY, x, totalW };
}

/**
 * Add a write-on line (for student answers).
 * @param {PDFDocument} doc
 * @param {string} label — e.g., "Numeral:", "Words:"
 * @param {number} y
 * @param {object} opts — { answer, lineWidth, color }
 * @returns {number} y after the line
 */
function addWriteLine(doc, label, y, opts) {
  const o = opts || {};
  const x = PAGE.MARGIN;
  doc.fontSize(11).font("Sans-Bold").fillColor("#000000");
  doc.text(label, x, y);

  const labelW = doc.widthOfString(label) + 8;
  const lineX = x + labelW;
  const lineY = y + 14;
  const lineW = (o.lineWidth || PAGE.CONTENT_W) - labelW;

  doc.save();
  doc.moveTo(lineX, lineY).lineTo(lineX + lineW, lineY)
    .strokeColor(STUDENT_RULE_COLOR).lineWidth(STUDENT_RULE_WIDTH).stroke();
  doc.restore();

  // Write answer if provided (for answer keys)
  if (o.answer) {
    doc.fontSize(11).font("Sans").fillColor(hex(o.color || "0F7F8C"));
    doc.text(o.answer, lineX + 4, y, { width: lineW - 8 });
  }

  return y + 24;
}

/**
 * Add a numbered problem with optional place value chart and write lines.
 * This is the main building block for PV worksheets.
 * @param {PDFDocument} doc
 * @param {number} num — problem number (1, 2, 3...)
 * @param {string} prompt — the question text
 * @param {number} y — starting y position
 * @param {object} opts — {
 *   headers: string[] — PV chart column headers (omit to skip chart),
 *   chartValues: (string|number)[] — pre-filled values for answer key,
 *   writeLines: { label, answer }[] — lines after chart,
 *   color: string — accent colour,
 *   problemBg: boolean — add light background to problem area
 * }
 * @returns {number} y after the problem
 */
function addProblem(doc, num, prompt, y, opts) {
  const o = opts || {};
  const x = PAGE.MARGIN;
  const color = o.color || "1B3A6B";

  y = ensureBlockFits(doc, y, measureProblemHeight(doc, prompt, o));

  // Problem number badge
  doc.save();
  doc.circle(x + 10, y + 8, 10).fill(hex(color));
  doc.fontSize(11).font("Sans-Bold").fillColor("#FFFFFF");
  doc.text(String(num), x + 2, y + 2, { width: 17, align: "center" });
  doc.restore();

  // Prompt text
  doc.fontSize(11).font("Sans").fillColor("#000000");
  doc.text(prompt, x + 28, y, { width: PAGE.CONTENT_W - 28 });
  y = doc.y + 10;

  // Place value chart
  if (o.headers) {
    const chart = addPvChartPdf(doc, y, o.headers, {
      color,
      values: o.chartValues,
    });
    y = chart.bottomY + 10;
  }

  // Write lines
  if (o.writeLines) {
    o.writeLines.forEach((wl) => {
      y = addWriteLine(doc, wl.label, y, { answer: wl.answer, color });
    });
  }

  // Spacing between problems
  y += 12;

  return y;
}

/**
 * Add a "First... Next... Then..." instruction block.
 * @param {PDFDocument} doc
 * @param {string[]} steps — array of step descriptions
 * @param {number} y
 * @param {object} opts — { color }
 * @returns {number} y after the block
 */
function addStepInstructions(doc, steps, y, opts) {
  const o = opts || {};
  const color = o.color || "1B3A6B";
  const x = PAGE.MARGIN;
  const labels = ["First:", "Next:", "Then:", "Finally:"];

  steps.forEach((step, i) => {
    const label = labels[i] || `Step ${i + 1}:`;
    doc.fontSize(11).font("Sans-Bold").fillColor(hex(color));
    doc.text(label + " ", x + 10, y, { continued: true });
    doc.font("Sans").fillColor("#000000");
    doc.text(step);
    y = doc.y + 6;
  });

  return y + 4;
}

/**
 * Add a light-background tip/instruction box.
 * @param {PDFDocument} doc
 * @param {string} text
 * @param {number} y
 * @param {object} opts — { color, icon }
 * @returns {number} y after the box
 */
function addTipBox(doc, text, y, opts) {
  const o = opts || {};
  const color = o.color || "0F7F8C";
  const x = PAGE.MARGIN;
  const w = PAGE.CONTENT_W;

  // Measure text height (use same font as render: Helvetica-Oblique)
  const textH = doc.fontSize(10).font("Sans-Italic").heightOfString(text, {
    width: w - 30,
  });
  const boxH = textH + 16;
  y = ensureBlockFits(doc, y, boxH + 10);

  doc.save();
  doc.roundedRect(x, y, w, boxH, 3).fill(lighten(color, 0.85));
  // Left accent
  doc.rect(x, y, 4, boxH).fill(hex(color));
  doc.restore();

  doc.fontSize(10).font("Sans-Italic").fillColor(hex("2D3142"));
  doc.text(text, x + 14, y + 8, { width: w - 30 });

  return y + boxH + 10;
}

/**
 * Add a footer to the current page.
 * @param {PDFDocument} doc
 * @param {string} text — footer text
 * @param {object} opts — { color }
 */
function addPdfFooter(doc, text, opts) {
  const o = opts || {};
  doc.fontSize(8).font("Sans").fillColor(hex(o.color || "9CA3AF"));
  doc.text(text, PAGE.MARGIN, FOOTER_Y, {
    width: PAGE.CONTENT_W,
    align: "center",
    lineBreak: false,
  });
}

/**
 * Add a blank lined area for extended writing.
 * @param {PDFDocument} doc
 * @param {number} y — starting y
 * @param {number} lineCount — number of lines
 * @param {object} opts — { lineSpacing }
 * @returns {number} y after the lined area
 */
function addLinedArea(doc, y, lineCount, opts) {
  const o = opts || {};
  const preferredSpacing = o.lineSpacing || 28;
  const minSpacing = o.minLineSpacing || Math.min(preferredSpacing, 20);
  const x = PAGE.MARGIN;
  const w = PAGE.CONTENT_W;
  let spacing = preferredSpacing;
  const available = CONTENT_BOTTOM - y;
  const required = lineCount * preferredSpacing;

  if (required > available) {
    const compactSpacing = Math.floor(available / Math.max(lineCount, 1));
    if (o.compact !== false && compactSpacing >= minSpacing) {
      spacing = compactSpacing;
    } else {
      doc.addPage();
      y = PAGE.MARGIN;
    }
  }

  if (y + lineCount * spacing > CONTENT_BOTTOM) {
    const refitSpacing = Math.floor((CONTENT_BOTTOM - y) / Math.max(lineCount, 1));
    if (o.compact !== false && refitSpacing >= minSpacing) {
      spacing = refitSpacing;
    }
  }

  doc.save();
  doc.strokeColor(STUDENT_RULE_COLOR).lineWidth(STUDENT_RULE_WIDTH);
  for (let i = 0; i < lineCount; i++) {
    const ly = y + i * spacing;
    doc.moveTo(x, ly).lineTo(x + w, ly).stroke();
  }
  doc.restore();

  return y + lineCount * spacing;
}

/**
 * Add a two-column graphic organiser.
 * @param {PDFDocument} doc
 * @param {string} leftHeader
 * @param {string} rightHeader
 * @param {number} y
 * @param {object} opts — { color, rows, rowH, leftContent, rightContent }
 * @returns {number} y after the organiser
 */
function addTwoColumnOrganiser(doc, leftHeader, rightHeader, y, opts) {
  const o = opts || {};
  const color = o.color || "1B3A6B";
  const x = PAGE.MARGIN;
  const colW = PAGE.CONTENT_W / 2;
  const hdrH = 26;
  const rows = o.rows || 4;
  const rowH = o.rowH || 50;

  // Headers
  doc.save();
  [leftHeader, rightHeader].forEach((h, i) => {
    const cx = x + i * colW;
    doc.rect(cx, y, colW, hdrH).fill(hex(color));
    doc.fontSize(10).font("Sans-Bold").fillColor("#FFFFFF");
    doc.text(h, cx, y + 7, { width: colW, align: "center" });
  });
  doc.restore();

  // Rows
  const leftContent = o.leftContent || [];
  const rightContent = o.rightContent || [];
  doc.save();
  for (let r = 0; r < rows; r++) {
    const ry = y + hdrH + r * rowH;
    [0, 1].forEach((i) => {
      const cx = x + i * colW;
      doc.rect(cx, ry, colW, rowH).lineWidth(STUDENT_RULE_WIDTH).strokeColor(STUDENT_RULE_COLOR).stroke();
      const content = i === 0 ? leftContent : rightContent;
      if (content[r]) {
        doc.fontSize(10).font("Sans").fillColor("#000000");
        doc.text(content[r], cx + 6, ry + 6, { width: colW - 12 });
      }
    });
  }
  doc.restore();

  return y + hdrH + rows * rowH + 10;
}

/**
 * Add a clear labelled cycle diagram for science worksheets and scaffolds.
 * Supports 4-stage cycles cleanly.
 *
 * @param {PDFDocument} doc
 * @param {number} y
 * @param {{label:string, color?:string}[]} stages
 * @param {object} opts
 * @returns {number}
 */
function addCycleDiagramPdf(doc, y, stages, opts) {
  const o = opts || {};
  const x = PAGE.MARGIN;
  const w = PAGE.CONTENT_W;
  const h = o.height || 220;
  const stageList = (stages || []).slice(0, 4);
  const centerX = x + w / 2;
  const centerY = y + h / 2 + 4;
  const orbitX = 150;
  const orbitY = 66;
  const nodeW = 112;
  const nodeH = 28;
  const lineColor = hex(o.lineColor || "6B7280");
  const centerLabel = o.centerLabel || "Cycle";

  y = ensureBlockFits(doc, y, h + 12);

  const positions = [
    { x: centerX - nodeW / 2, y: centerY - orbitY - nodeH / 2 - 28, lineX: x + w - 120, lineY: centerY - orbitY - 8 },
    { x: centerX + orbitX - nodeW / 2, y: centerY - nodeH / 2, lineX: x + w - 120, lineY: centerY + 8 },
    { x: centerX - nodeW / 2, y: centerY + orbitY - nodeH / 2 + 28, lineX: x + w - 120, lineY: centerY + orbitY + 36 },
    { x: centerX - orbitX - nodeW / 2, y: centerY - nodeH / 2, lineX: x + 18, lineY: centerY + 8 },
  ];

  function drawArrowWithHead(x1, y1, x2, y2) {
    const angle = Math.atan2(y2 - y1, x2 - x1);
    const head = 7;
    doc.save();
    doc.moveTo(x1, y1).lineTo(x2, y2).lineWidth(1.4).strokeColor(lineColor).stroke();
    doc.moveTo(x2, y2)
      .lineTo(x2 - head * Math.cos(angle - Math.PI / 6), y2 - head * Math.sin(angle - Math.PI / 6))
      .lineTo(x2 - head * Math.cos(angle + Math.PI / 6), y2 - head * Math.sin(angle + Math.PI / 6))
      .lineTo(x2, y2)
      .fill(lineColor);
    doc.restore();
  }

  // Center label
  doc.save();
  doc.roundedRect(centerX - 48, centerY - 18, 96, 36, 10)
    .lineWidth(1.2).strokeColor(hex(o.centerColor || "1B3A6B"))
    .fillAndStroke(lighten(o.centerColor || "1B3A6B", 0.88), hex(o.centerColor || "1B3A6B"));
  doc.fontSize(12).font("Sans-Bold").fillColor(hex(o.centerColor || "1B3A6B"));
  doc.text(centerLabel, centerX - 42, centerY - 7, { width: 84, align: "center" });
  doc.restore();

  // Nodes
  stageList.forEach((stage, index) => {
    const pos = positions[index];
    const color = hex((stage && stage.color) || ["1B3A6B", "0F7F8C", "C2742D", "2D8C4A"][index]);
    const labelText = o.showStageNames === false ? `${index + 1}` : `${index + 1}. ${String((stage && stage.label) || "")}`;

    doc.save();
    doc.roundedRect(pos.x, pos.y, nodeW, nodeH, 12)
      .lineWidth(1.2).strokeColor(color)
      .fillAndStroke(lighten(color, 0.9), color);
    doc.fontSize(9.5).font("Sans-Bold").fillColor(color);
    doc.text(labelText, pos.x + 8, pos.y + 9, { width: nodeW - 16, align: "center" });
    doc.restore();

    if (o.numberedLines !== false) {
      const lineStartX = pos.lineX + 22;
      doc.fontSize(10.5).font("Sans-Bold").fillColor("#000000");
      doc.text(`${index + 1}.`, pos.lineX, pos.lineY - 2, { width: 18, align: "right" });
      doc.moveTo(lineStartX, pos.lineY + 8).lineTo(lineStartX + 105, pos.lineY + 8)
        .strokeColor(STUDENT_RULE_COLOR).lineWidth(STUDENT_RULE_WIDTH).stroke();
    }
  });

  // Arrow loop
  drawArrowWithHead(centerX - 10, centerY - orbitY + 4, centerX + orbitX - 74, centerY - 16);
  drawArrowWithHead(centerX + orbitX - 8, centerY + 22, centerX + 18, centerY + orbitY + 24);
  drawArrowWithHead(centerX - 18, centerY + orbitY + 38, centerX - orbitX + 72, centerY + 22);
  drawArrowWithHead(centerX - orbitX + 8, centerY - 18, centerX - 20, centerY - orbitY + 2);

  return y + h + 10;
}

// ── Resource slide helper (for PPTX) ────────────────────────────────────────

/**
 * Add a "Teacher Resources" slide to a PPTX presentation.
 * Lists companion PDF files with clickable hyperlinks.
 *
 * @param {object} pres — PptxGenJS presentation object
 * @param {object[]} resources — [{ name, fileName, description }]
 * @param {object} theme — { C, FONT_H, FONT_B, addTopBar, addTitle, addFooter, addCard }
 * @param {string} footer — footer text
 * @param {string} notes — teacher notes
 * @returns {object} the slide
 */
function drawMockupTextPdf(doc, x, y, w, text, opts) {
  const o = opts || {};
  if (!text) return;
  doc.fontSize(o.fontSize || 10.5)
    .font(o.font || "Sans")
    .fillColor(hex(o.color || "243142"))
    .text(String(text), x, y, {
      width: w,
      align: o.align || "left",
    });
}

function drawLineSetPdf(doc, x, y, w, h, component, spec) {
  const count = Math.max(1, Number(component.count) || 3);
  const lineGap = component.lineGap != null ? component.lineGap : 4;
  const linePad = component.linePad != null ? component.linePad : 8;
  const lineH = Math.max(3, (h - linePad * 2 - lineGap * (count - 1)) / count);
  const widths = Array.isArray(component.widths) ? component.widths : [];
  const lineColor = hex(component.lineColor || spec.mutedLine || "A0A0A0");
  for (let index = 0; index < count; index += 1) {
    const ratio = widths[index] || (index === count - 1 ? 0.62 : index % 2 === 0 ? 0.92 : 0.82);
    doc.roundedRect(x + 6, y + linePad + index * (lineH + lineGap), Math.max(14, (w - 12) * ratio), lineH, 2).fill(lineColor);
  }
}

function drawPhotoPlaceholderPdf(doc, x, y, w, h, component, spec) {
  const bg = hex(component.fill || lightenHex(spec.accent, 0.9));
  const border = hex(component.border || spec.softBorder || "A0A0A0");
  doc.roundedRect(x, y, w, h, 6).lineWidth(0.8).strokeColor(border).fillAndStroke(bg, border);
  doc.rect(x + 4, y + 4, w - 8, h - 8).lineWidth(0.6).strokeColor(border).fillAndStroke(hex(lightenHex(component.fill || lightenHex(spec.accent, 0.9), 0.05)), border);
  doc.moveTo(x + w * 0.18, y + h * 0.72).lineTo(x + w * 0.4, y + h * 0.52).lineTo(x + w * 0.54, y + h * 0.72)
    .lineWidth(1.1).strokeColor(hex(spec.mutedLine)).stroke();
  doc.moveTo(x + w * 0.5, y + h * 0.72).lineTo(x + w * 0.66, y + h * 0.58).lineTo(x + w * 0.8, y + h * 0.72)
    .lineWidth(1.1).strokeColor(hex(spec.mutedLine)).stroke();
  doc.circle(x + w * 0.72, y + h * 0.2, Math.min(w, h) * 0.06).fill(hex(spec.accent));
}

function drawChartPlaceholderPdf(doc, x, y, w, h, component, spec) {
  const bg = hex(component.fill || "FFFFFF");
  const border = hex(component.border || spec.softBorder);
  doc.roundedRect(x, y, w, h, 6).lineWidth(0.8).strokeColor(border).fillAndStroke(bg, border);
  doc.moveTo(x + 8, y + h - 8).lineTo(x + w - 8, y + h - 8).lineWidth(0.8).strokeColor(hex(spec.mutedLine)).stroke();
  doc.moveTo(x + 8, y + 8).lineTo(x + 8, y + h - 8).lineWidth(0.8).strokeColor(hex(spec.mutedLine)).stroke();
  [0.28, 0.48, 0.68].forEach((pos, index) => {
    const height = h * (0.22 + index * 0.12);
    const fill = index === 1 ? hex(spec.accent) : hex(lightenHex(spec.accent, 0.38 + index * 0.12));
    doc.roundedRect(x + w * pos, y + h - 8 - height, w * 0.1, height, 3).fill(fill);
  });
}

function drawDiagramPlaceholderPdf(doc, x, y, w, h, component, spec) {
  const bg = hex(component.fill || lightenHex(spec.accent, 0.92));
  const border = hex(component.border || spec.softBorder);
  doc.roundedRect(x, y, w, h, 6).lineWidth(0.8).strokeColor(border).fillAndStroke(bg, border);
  const points = [
    { cx: 0.24, cy: 0.58 },
    { cx: 0.5, cy: 0.34 },
    { cx: 0.76, cy: 0.58 },
  ];
  doc.moveTo(x + w * points[0].cx, y + h * points[0].cy).lineTo(x + w * points[1].cx, y + h * points[1].cy)
    .lineWidth(0.9).strokeColor(hex(spec.mutedLine)).stroke();
  doc.moveTo(x + w * points[1].cx, y + h * points[1].cy).lineTo(x + w * points[2].cx, y + h * points[2].cy)
    .lineWidth(0.9).strokeColor(hex(spec.mutedLine)).stroke();
  points.forEach((point, index) => {
    const fill = index === 1 ? hex(spec.accent) : hex(lightenHex(spec.accent, 0.45));
    doc.circle(x + w * point.cx, y + h * point.cy, Math.min(w, h) * 0.06).fill(fill);
  });
}

function drawBrowserFramePlaceholderPdf(doc, x, y, w, h, component, spec) {
  const border = hex(component.border || spec.softBorder);
  doc.roundedRect(x, y, w, h, 6).lineWidth(0.8).strokeColor(border).fillAndStroke(hex(component.fill || "FFFFFF"), border);
  doc.roundedRect(x + 4, y + 4, w - 8, 16, 4).fill(hex(lightenHex(spec.accent, 0.92)));
  [0, 1, 2].forEach((index) => {
    doc.circle(x + 14 + index * 12, y + 12, 3).fill(hex(lightenHex(spec.accent, 0.5)));
  });
  doc.roundedRect(x + 8, y + 28, w * 0.28, h - 36, 4).fill(hex(lightenHex(spec.accent, 0.94)));
  doc.roundedRect(x + w * 0.4, y + 28, w * 0.5, h - 36, 4).lineWidth(0.8).strokeColor(border).fillAndStroke("#FFFFFF", border);
  drawLineSetPdf(doc, x + w * 0.4, y + 28, w * 0.5, h - 36, { count: 4, lineColor: spec.mutedLine }, spec);
}

function drawCardGridPlaceholderPdf(doc, x, y, w, h, component, spec) {
  const border = hex(component.border || spec.softBorder);
  doc.roundedRect(x, y, w, h, 6).lineWidth(0.8).strokeColor(border).fillAndStroke(hex(component.fill || "FFFFFF"), border);
  const rows = Math.max(1, Number(component.rows) || 2);
  const cols = Math.max(1, Number(component.cols) || 2);
  const gap = 4;
  const cardW = (w - gap * (cols + 1)) / cols;
  const cardH = (h - gap * (rows + 1)) / rows;
  for (let row = 0; row < rows; row += 1) {
    for (let col = 0; col < cols; col += 1) {
      const cardX = x + gap + col * (cardW + gap);
      const cardY = y + gap + row * (cardH + gap);
      doc.roundedRect(cardX, cardY, cardW, cardH, 4).lineWidth(0.8).strokeColor(border).fillAndStroke(hex(lightenHex(spec.accent, 0.95)), border);
      doc.roundedRect(cardX + 4, cardY + 4, cardW - 8, 8, 2).fill(hex(lightenHex(spec.accent, 0.25)));
      drawLineSetPdf(doc, cardX, cardY + 12, cardW, cardH - 14, { count: 2, lineColor: spec.mutedLine }, spec);
    }
  }
}

function drawHeroMockupPdf(doc, x, y, w, h, component, spec) {
  const mode = component.mode || "diagram";
  if (mode === "chart") {
    drawChartPlaceholderPdf(doc, x, y, w, h, component, spec);
  } else if (mode === "browserFrame") {
    drawBrowserFramePlaceholderPdf(doc, x, y, w, h, component, spec);
  } else if (mode === "cardGrid") {
    drawCardGridPlaceholderPdf(doc, x, y, w, h, component, spec);
  } else if (mode === "photo") {
    drawPhotoPlaceholderPdf(doc, x, y, w, h, component, spec);
  } else {
    drawDiagramPlaceholderPdf(doc, x, y, w, h, component, spec);
  }

  if (component.overlayText) {
    const overlayFill = hex(component.overlayFill || spec.accent);
    const overlayH = Math.min(18, h * 0.22);
    doc.roundedRect(x + 6, y + 6, w - 12, overlayH, 4).fill(overlayFill);
    drawMockupTextPdf(doc, x + 10, y + 8, w - 20, component.overlayText, {
      fontSize: component.overlayFontSize || 9,
      font: "Sans-Bold",
      color: component.textColor || "FFFFFF",
      align: component.align || "left",
    });
  }
}

function addPosterMockupPdf(doc, x, y, w, h, spec, opts) {
  const o = opts || {};
  if (!isStructuredMockupSpec(spec)) return y + h;

  const normalized = normalizeStructuredMockup(spec);
  const pageFill = hex(normalized.pageFill || o.pageFill || "FFFFFF");
  const pageBorder = hex(normalized.pageBorder || o.pageBorder || "9CA3AF");
  const innerPad = normalized.innerPad != null ? normalized.innerPad : 6;
  const gap = normalized.gap != null ? normalized.gap : 4;

  doc.save();
  doc.roundedRect(x, y, w, h, 8).lineWidth(1).strokeColor(pageBorder).fillAndStroke(pageFill, pageBorder);
  doc.restore();

  const innerX = x + innerPad;
  const innerY = y + innerPad;
  const innerW = w - innerPad * 2;
  const components = normalized.components || [];
  const availableH = h - innerPad * 2 - gap * Math.max(components.length - 1, 0);
  const totalScale = components.reduce((sum, component) => sum + component.scale, 0) || 1;
  let cursorY = innerY;

  components.forEach((component) => {
    const blockH = Math.max(12, availableH * (component.scale / totalScale));
    const textColor = hex(component.textColor || normalized.textColor || "243142");

    if (component.kind === "masthead") {
      const fill = hex(component.fill || normalized.accent);
      doc.roundedRect(innerX, cursorY, innerW, blockH, 5).lineWidth(0.6).strokeColor(fill).fillAndStroke(fill, fill);
      drawMockupTextPdf(doc, innerX + 6, cursorY + Math.max(2, blockH * 0.18), innerW - 12, component.text, {
        fontSize: component.fontSize || (blockH > 22 ? 10.5 : 9),
        font: "Sans-Bold",
        color: component.textColor || "FFFFFF",
        align: component.align || "center",
      });
    } else if (component.kind === "nav") {
      const fill = hex(component.fill || normalized.softFill);
      const border = hex(component.border || normalized.softBorder);
      doc.roundedRect(innerX, cursorY, innerW, blockH, 5).lineWidth(0.8).strokeColor(border).fillAndStroke(fill, border);
      drawMockupTextPdf(doc, innerX + 6, cursorY + Math.max(2, blockH * 0.18), innerW - 12, component.text, {
        fontSize: component.fontSize || 8,
        font: "Sans-Bold",
        color: component.textColor || normalized.textColor,
      });
    } else if (component.kind === "heading" || component.kind === "subheading") {
      const fill = hex(component.fill || "FFFFFF");
      const border = hex(component.border || normalized.softBorder);
      doc.roundedRect(innerX, cursorY, innerW, blockH, 5).lineWidth(0.8).strokeColor(border).fillAndStroke(fill, border);
      drawMockupTextPdf(doc, innerX + 6, cursorY + Math.max(2, blockH * 0.18), innerW - 12, component.text, {
        fontSize: component.fontSize || (component.kind === "heading" ? 10.5 : 9),
        font: component.kind === "heading" ? "Sans-Bold" : "Sans",
        color: component.textColor || (component.kind === "heading" ? normalized.accent : normalized.textColor),
      });
    } else if (component.kind === "hero") {
      drawHeroMockupPdf(doc, innerX, cursorY, innerW, blockH, component, normalized);
    } else if (component.kind === "chart") {
      drawChartPlaceholderPdf(doc, innerX, cursorY, innerW, blockH, component, normalized);
    } else if (component.kind === "stat") {
      const fill = hex(component.fill || lightenHex(normalized.accent, 0.1));
      doc.roundedRect(innerX, cursorY, innerW, blockH, 5).lineWidth(0.6).strokeColor(fill).fillAndStroke(fill, fill);
      drawMockupTextPdf(doc, innerX + 6, cursorY + Math.max(2, blockH * 0.18), innerW - 12, component.text, {
        fontSize: component.fontSize || 10,
        font: "Sans-Bold",
        color: component.textColor || "FFFFFF",
        align: component.align || "center",
      });
    } else if (component.kind === "cta" || component.kind === "footerBand") {
      const outerFill = hex(component.kind === "cta" ? "FFFFFF" : (component.fill || normalized.softFill));
      const outerBorder = hex(component.border || normalized.softBorder);
      doc.roundedRect(innerX, cursorY, innerW, blockH, 5).lineWidth(0.8).strokeColor(outerBorder).fillAndStroke(outerFill, outerBorder);
      const buttonW = component.kind === "cta" ? innerW * 0.64 : innerW;
      const buttonX = component.kind === "cta" ? innerX + (innerW - buttonW) / 2 : innerX;
      const fill = hex(component.fill || normalized.accent);
      doc.roundedRect(buttonX, cursorY + 4, buttonW, Math.max(8, blockH - 8), 4).fill(fill);
      drawMockupTextPdf(doc, buttonX + 6, cursorY + Math.max(6, blockH * 0.24), buttonW - 12, component.text, {
        fontSize: component.fontSize || 9,
        font: "Sans-Bold",
        color: component.textColor || "FFFFFF",
        align: component.align || "center",
      });
    } else if (component.kind === "textBlock") {
      const fill = hex(component.fill || "FFFFFF");
      const border = hex(component.border || normalized.softBorder);
      doc.roundedRect(innerX, cursorY, innerW, blockH, 5).lineWidth(0.8).strokeColor(border).fillAndStroke(fill, border);
      drawLineSetPdf(doc, innerX, cursorY, innerW, blockH, component, normalized);
    } else if (component.kind === "caption") {
      const fill = hex(component.fill || normalized.softFill);
      const border = hex(component.border || normalized.softBorder);
      doc.roundedRect(innerX, cursorY, innerW, blockH, 5).lineWidth(0.8).strokeColor(border).fillAndStroke(fill, border);
      drawMockupTextPdf(doc, innerX + 6, cursorY + Math.max(2, blockH * 0.18), innerW - 12, component.text, {
        fontSize: component.fontSize || 8.2,
        font: "Sans-Italic",
        color: component.textColor || normalized.textColor,
      });
    } else if (component.kind === "quote") {
      const fill = hex(component.fill || "FFFFFF");
      const border = hex(component.border || normalized.softBorder);
      doc.roundedRect(innerX, cursorY, innerW, blockH, 5).lineWidth(0.8).strokeColor(border).fillAndStroke(fill, border);
      doc.rect(innerX + 4, cursorY + 4, 4, blockH - 8).fill(hex(component.accent || normalized.accent));
      drawMockupTextPdf(doc, innerX + 12, cursorY + Math.max(2, blockH * 0.18), innerW - 18, component.text, {
        fontSize: component.fontSize || 8.8,
        font: "Sans-Italic",
        color: component.textColor || normalized.textColor,
      });
    } else if (component.kind === "iconRow") {
      const fill = hex(component.fill || "FFFFFF");
      const border = hex(component.border || normalized.softBorder);
      doc.roundedRect(innerX, cursorY, innerW, blockH, 5).lineWidth(0.8).strokeColor(border).fillAndStroke(fill, border);
      const count = Math.max(3, Number(component.count) || 4);
      const gapW = innerW / (count + 1);
      for (let index = 0; index < count; index += 1) {
        const fillColor = index === 0 ? hex(normalized.accent) : hex(lightenHex(normalized.accent, 0.45));
        doc.circle(innerX + gapW * (index + 0.9), cursorY + blockH * 0.33, 5).fill(fillColor);
        doc.roundedRect(innerX + gapW * (index + 0.62), cursorY + blockH * 0.63, 16, 3, 1).fill(hex(normalized.mutedLine));
      }
    } else if (component.kind === "sidebar") {
      const fill = hex(component.fill || "FFFFFF");
      const border = hex(component.border || normalized.softBorder);
      doc.roundedRect(innerX, cursorY, innerW, blockH, 5).lineWidth(0.8).strokeColor(border).fillAndStroke(fill, border);
      doc.rect(innerX, cursorY, innerW * 0.22, blockH).fill(hex(lightenHex(component.fill || normalized.accent, 0.86)));
      drawLineSetPdf(doc, innerX + innerW * 0.24, cursorY, innerW * 0.72, blockH, { count: 3, lineColor: normalized.mutedLine }, normalized);
    } else {
      const fill = hex(component.fill || normalized.softFill);
      const border = hex(component.border || normalized.softBorder);
      doc.roundedRect(innerX, cursorY, innerW, blockH, 5).lineWidth(0.8).strokeColor(border).fillAndStroke(fill, border);
      drawMockupTextPdf(doc, innerX + 6, cursorY + Math.max(2, blockH * 0.18), innerW - 12, component.text, {
        fontSize: component.fontSize || 9,
        font: "Sans-Bold",
        color: textColor,
      });
    }

    cursorY += blockH + gap;
  });

  return y + h;
}

function addPosterPairPdf(doc, y, leftPoster, rightPoster, opts) {
  const o = opts || {};
  const labelH = 14;
  const posterH = o.posterH || 140;
  const gap = o.gap != null ? o.gap : 18;
  const totalH = labelH + 6 + posterH;
  y = ensureBlockFits(doc, y, totalH + 8);

  const x = PAGE.MARGIN;
  const pairW = PAGE.CONTENT_W;
  const posterW = (pairW - gap) / 2;
  const leftTitle = o.leftTitle || "Poster A";
  const rightTitle = o.rightTitle || "Poster B";
  const color = hex(o.color || "1B3A6B");

  doc.fontSize(11).font("Sans-Bold").fillColor(color);
  doc.text(leftTitle, x, y, { width: posterW });
  doc.text(rightTitle, x + posterW + gap, y, { width: posterW });

  addPosterMockupPdf(doc, x, y + labelH + 6, posterW, posterH, leftPoster, o.leftOpts);
  addPosterMockupPdf(doc, x + posterW + gap, y + labelH + 6, posterW, posterH, rightPoster, o.rightOpts);
  return y + totalH + 10;
}

/**
 * Render the §44 structured Teacher Resources slide. Accepts a config
 * object describing all classroom material types (PDFs, manipulatives,
 * board setup, videos, supplied URLs, OCHRE refs) and renders them in
 * a two-column layout.
 */
function addRichResourceSlide(pres, config, theme, footer, notes) {
  const { C: TC, FONT_H: FH, FONT_B: FB } = theme;
  const s = pres.addSlide();
  const cfg = config || {};
  const resources = Array.isArray(cfg.resources) ? cfg.resources : [];
  const groups = [
    { label: "Manipulatives", items: cfg.manipulatives },
    { label: "Mini-whiteboards & student tools", items: cfg.studentTools },
    { label: "Routine icons used", items: cfg.routineIcons },
    { label: "Teacher board setup", items: cfg.boardSetup },
    { label: "Videos & media", items: cfg.videos },
    { label: "Supplied URLs", items: cfg.urls },
    { label: "OCHRE / source materials", items: cfg.ochre },
  ].filter((g) => Array.isArray(g.items) && g.items.length > 0);

  s.background = { color: TC.CREAM || TC.BG_LIGHT || "F4F7FF" };
  s.addShape("rect", {
    x: 0, y: 0, w: 10, h: 0.06,
    fill: { color: TC.NAVY || TC.PRIMARY || "1B3A6B" },
  });

  const badgeColor = TC.TEAL || TC.SECONDARY || "0F7F8C";
  s.addShape("roundRect", {
    x: 0.5, y: 0.2, w: 2.8, h: 0.36, rectRadius: 0.08,
    fill: { color: badgeColor },
  });
  s.addText("Teacher Resources", {
    x: 0.5, y: 0.2, w: 2.8, h: 0.36,
    fontSize: 11, fontFace: FB, color: TC.WHITE || "FFFFFF",
    align: "center", valign: "middle", bold: true, margin: 0,
  });

  s.addText("Materials & references", {
    x: 0.5, y: 0.65, w: 9, h: 0.55,
    fontSize: 24, fontFace: FH, color: TC.NAVY || TC.PRIMARY || "1B3A6B",
    bold: true, margin: 0,
  });

  if (resources.length === 0) {
    s.addText("No printed student resources required.", {
      x: 0.5, y: 1.25, w: 9, h: 0.3,
      fontSize: 12, fontFace: FB, color: TC.MUTED || "6B7280",
      italic: true, margin: 0,
    });
  } else {
    s.addText("Click any resource to open. Print before the lesson.", {
      x: 0.5, y: 1.25, w: 9, h: 0.3,
      fontSize: 12, fontFace: FB, color: TC.MUTED || "6B7280",
      italic: true, margin: 0,
    });
  }

  const colTop = 1.6;
  const colH = 5.05 - colTop;
  const useTwoCols = resources.length > 0 && groups.length > 0;
  const leftW = useTwoCols ? 4.7 : 9;
  const leftX = 0.5;
  const rightX = leftX + leftW + 0.2;
  const rightW = 9.5 - rightX;

  if (resources.length > 0) {
    const cardH = Math.min(0.62, (colH - 0.18 * (resources.length - 1)) / Math.max(resources.length, 1));
    const gap = 0.16;
    resources.forEach((res, i) => {
      const displayName = cleanResourceLabel(res.name) || resourceNameFromFileName(res.fileName);
      const cy = colTop + i * (cardH + gap);
      s.addShape("roundRect", {
        x: leftX, y: cy, w: leftW, h: cardH, rectRadius: 0.08,
        fill: { color: TC.WHITE || "FFFFFF" },
        shadow: { type: "outer", blur: 4, offset: 1, color: "000000", opacity: 0.10, angle: 135 },
      });
      s.addShape("rect", {
        x: leftX, y: cy, w: 0.06, h: cardH,
        fill: { color: badgeColor },
      });
      const ICON_D = 0.40;
      s.addShape("roundRect", {
        x: leftX + 0.20, y: cy + (cardH - ICON_D) / 2, w: ICON_D, h: ICON_D, rectRadius: ICON_D / 2,
        fill: { color: TC.CORAL || TC.ALERT || "C94030" },
      });
      s.addText("PDF", {
        x: leftX + 0.20, y: cy + (cardH - ICON_D) / 2, w: ICON_D, h: ICON_D,
        fontSize: 9, fontFace: FB, color: TC.WHITE || "FFFFFF",
        align: "center", valign: "middle", bold: true, margin: 0,
      });
      s.addText(String(displayName), {
        x: leftX + 0.70, y: cy + 0.06, w: leftW - 0.85, h: 0.28,
        fontSize: 13, fontFace: FH, color: TC.NAVY || TC.PRIMARY || "1B3A6B",
        bold: true, margin: 0,
        hyperlink: { url: res.fileName, tooltip: "Open " + displayName },
      });
      if (res.description) {
        s.addText(String(res.description), {
          x: leftX + 0.70, y: cy + 0.34, w: leftW - 0.85, h: cardH - 0.40,
          fontSize: 10, fontFace: FB, color: TC.MUTED || "6B7280",
          margin: 0, fit: "shrink", shrinkText: true,
        });
      }
    });
  }

  if (groups.length > 0) {
    const groupX = useTwoCols ? rightX : leftX;
    const groupW = useTwoCols ? rightW : leftW;
    const totalGroups = groups.length;
    const groupH = colH / Math.max(totalGroups, 1);
    groups.forEach((g, i) => {
      const gy = colTop + i * groupH;
      s.addText(g.label, {
        x: groupX, y: gy, w: groupW, h: 0.28,
        fontSize: 12, fontFace: FB, color: badgeColor,
        bold: true, margin: 0,
      });
      const items = g.items.map((entry) => String(entry == null ? "" : entry).trim()).filter(Boolean);
      const bodyH = Math.max(0.4, groupH - 0.36);
      s.addText(items.map((entry, idx) => ({
        text: entry,
        options: {
          bullet: true, breakLine: idx < items.length - 1,
          fontSize: 11, color: TC.CHARCOAL || "1F2937",
        },
      })), {
        x: groupX + 0.06, y: gy + 0.30, w: groupW - 0.06, h: bodyH,
        fontFace: FB, valign: "top", margin: 0,
        fit: "shrink", shrinkText: true,
      });
    });
  }

  if (footer) {
    s.addText(footer, {
      x: 0.5, y: 5.3, w: 9, h: 0.2,
      fontSize: 9, fontFace: FB, color: TC.MUTED || "6B7280", margin: 0,
    });
  }
  if (notes) s.addNotes(notes);
  return s;
}

function addResourceSlide(pres, resources, theme, footer, notes) {
  // Structured §44 config object (manipulatives, board setup, videos, URLs,
  // OCHRE) routes to the rich slide; a flat array uses the legacy renderer.
  if (resources && !Array.isArray(resources) && typeof resources === "object") {
    return addRichResourceSlide(pres, resources, theme, footer, notes);
  }

  const { C: TC, FONT_H: FH, FONT_B: FB } = theme;
  const s = pres.addSlide();

  // Background + top bar
  s.background = { color: TC.CREAM || "F4F7FF" };
  s.addShape("rect", {
    x: 0, y: 0, w: 10, h: 0.06,
    fill: { color: TC.NAVY || "1B3A6B" },
  });

  // Badge
  const badgeColor = TC.TEAL || "0F7F8C";
  s.addShape("roundRect", {
    x: 0.5, y: 0.2, w: 2.8, h: 0.36, rectRadius: 0.08,
    fill: { color: badgeColor },
  });
  s.addText("Printable Resources", {
    x: 0.5, y: 0.2, w: 2.8, h: 0.36,
    fontSize: 11, fontFace: FB, color: TC.WHITE || "FFFFFF",
    align: "center", valign: "middle", bold: true, margin: 0,
  });

  // Title
  s.addText("Teacher Resources", {
    x: 0.5, y: 0.65, w: 9, h: 0.55,
    fontSize: 24, fontFace: FH, color: TC.NAVY || "1B3A6B",
    bold: true, margin: 0,
  });

  // Instruction text
  s.addText("Click any resource below to open the PDF. Print before the lesson.", {
    x: 0.5, y: 1.2, w: 9, h: 0.3,
    fontSize: 12, fontFace: FB, color: TC.MUTED || "6B7280",
    italic: true, margin: 0,
  });

  // Resource cards
  const cardH = 0.7;
  const gap = 0.15;
  const startY = 1.7;

  resources.forEach((res, i) => {
    const displayName = cleanResourceLabel(res.name) || resourceNameFromFileName(res.fileName);
    const cy = startY + i * (cardH + gap);

    // Card background
    s.addShape("roundRect", {
      x: 0.5, y: cy, w: 9, h: cardH, rectRadius: 0.08,
      fill: { color: TC.WHITE || "FFFFFF" },
      shadow: { type: "outer", blur: 4, offset: 1, color: "000000", opacity: 0.10, angle: 135 },
    });

    // Left accent strip
    s.addShape("rect", {
      x: 0.5, y: cy, w: 0.06, h: cardH,
      fill: { color: badgeColor },
    });

    // PDF icon circle
    const ICON_D = 0.46;
    s.addShape("roundRect", {
      x: 0.75, y: cy + 0.12, w: ICON_D, h: ICON_D, rectRadius: ICON_D / 2,
      fill: { color: TC.CORAL || "C94030" },
    });
    s.addText("PDF", {
      x: 0.75, y: cy + 0.12, w: ICON_D, h: ICON_D,
      fontSize: 9, fontFace: FB, color: TC.WHITE || "FFFFFF",
      align: "center", valign: "middle", bold: true, margin: 0,
    });

    // Resource name (clickable)
    s.addText(displayName, {
      x: 1.4, y: cy + 0.08, w: 7.5, h: 0.3,
      fontSize: 14, fontFace: FH, color: TC.NAVY || "1B3A6B",
      bold: true, margin: 0,
      hyperlink: { url: res.fileName, tooltip: "Open " + displayName },
    });

    // Description
    if (res.description) {
      s.addText(res.description, {
        x: 1.4, y: cy + 0.38, w: 7.5, h: 0.25,
        fontSize: 10, fontFace: FB, color: TC.MUTED || "6B7280",
        margin: 0,
      });
    }
  });

  if (footer) {
    s.addText(footer, {
      x: 0.5, y: 5.3, w: 9, h: 0.2,
      fontSize: 9, fontFace: FB, color: TC.MUTED || "6B7280", margin: 0,
    });
  }
  if (notes) s.addNotes(notes);

  return s;
}

module.exports = {
  // Constants
  PAGE,
  // Utilities
  hex, lighten,
  cleanResourceLabel,
  getSessionResourceFolder,
  formatSessionResourceName,
  formatSessionResourceFileName,
  makeSessionResource,
  // Document lifecycle
  createPdf, writePdf,
  // Page elements
  addPdfHeader, addSectionHeading, addBodyText,
  addPvChartPdf, addWriteLine, addProblem,
  addStepInstructions, addTipBox, addPdfFooter,
  addLinedArea, addTwoColumnOrganiser, addCycleDiagramPdf, addPosterMockupPdf, addPosterPairPdf,
  // PPTX integration
  addResourceSlide,
  addRichResourceSlide,
};
