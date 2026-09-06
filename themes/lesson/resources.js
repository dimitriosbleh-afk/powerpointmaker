"use strict";

/**
 * Declarative companion resources for lesson specs (docs/lesson-spec.md).
 *
 * Three resource kinds:
 *   worksheet  numbered items, each with a prompt, an optional paper twin of
 *              a slide visual, and answer lines; an answer key is generated
 *              from the same items unless answerKey: false
 *   page       free layout of blocks (heading, text, tip, steps, visual,
 *              lines, organiser, box) for scaffolds and organisers
 *   cards      cut-out cards in a grid, each with text and/or a visual
 *
 * Every visual is drawn with the pdf_helpers twins or the same pictogram
 * renderer as the slides, so paper shows the representation students met
 * on screen (megaprompt 0a item 16). Sizes follow the year band.
 */

const path = require("path");
const P = require("../pdf_helpers");
const { renderPictogramPng } = require("../core/pictograms");
const { byBand } = require("../core/gradeBand");

const CONTENT_BOTTOM = P.PAGE.H - P.PAGE.MARGIN - 40;

function bandSizes(sz) {
  return {
    body: byBand(sz, 16, 14, 12),
    prompt: byBand(sz, 18, 16, 13),
    heading: byBand(sz, 20, 18, 15),
    lineGap: byBand(sz, 30, 28, 24),
    cell: byBand(sz, 34, 30, 26),
  };
}

function fits(doc, y, needed) {
  if (y + needed > CONTENT_BOTTOM && needed <= CONTENT_BOTTOM - P.PAGE.MARGIN) {
    doc.addPage();
    return P.PAGE.MARGIN;
  }
  return y;
}

/* ── Paper twins of the slide visuals ─────────────────────────────────── */

function drawFrame(doc, x, y, cols, rows, cell, filled, fill, border) {
  for (let r = 0; r < rows; r += 1) {
    for (let c = 0; c < cols; c += 1) {
      const idx = r * cols + c;
      doc.save().lineWidth(1.2).strokeColor(border).rect(x + c * cell, y + r * cell, cell, cell).stroke().restore();
      if (idx < filled) {
        const d = cell * 0.62;
        doc.save().fillColor(fill).circle(x + c * cell + cell / 2, y + r * cell + cell / 2, d / 2).fill().restore();
      }
    }
  }
  return y + rows * cell;
}

function drawDots(doc, x, y, rows, cols, cell, fill) {
  for (let r = 0; r < rows; r += 1) {
    for (let c = 0; c < cols; c += 1) {
      doc.save().fillColor(fill).circle(x + c * cell + cell / 2, y + r * cell + cell / 2, cell * 0.32).fill().restore();
    }
  }
  return y + rows * cell;
}

/**
 * Draw a visual spec on the page at (x, y). Returns the next usable y.
 * `answer` true renders the answer state where the type has one.
 */
function drawVisualPdf(doc, visual, x, y, opts) {
  const o = opts || {};
  const C = o.C;
  const sz = o.sz;
  const S = bandSizes(sz);
  const fill = P.hex(o.fillColor || C.PRIMARY);
  const border = P.hex(C.CHARCOAL);
  const w = o.width || P.PAGE.CONTENT_W;
  const v = visual;
  switch (v.type) {
    case "tensFrame": return drawFrame(doc, x, y, 5, 2, S.cell, v.filled || 0, fill, border) + 8;
    case "fiveFrame": return drawFrame(doc, x, y, 5, 1, S.cell, v.filled || 0, fill, border) + 8;
    case "doubleTensFrame": {
      let yy = drawFrame(doc, x, y, 5, 2, S.cell, v.filledTop || 0, fill, border) + 10;
      yy = drawFrame(doc, x, yy, 5, 2, S.cell, v.filledBottom || 0, P.hex(C.SECONDARY), border) + 8;
      return yy;
    }
    case "dotCard":
    case "dotCards": {
      const counts = v.type === "dotCard" ? [v.count || 1] : (v.counts || [1]);
      const size = S.cell * 2.4;
      const PAT = { 1: [[0.5, 0.5]], 2: [[0.28, 0.28], [0.72, 0.72]], 3: [[0.25, 0.25], [0.5, 0.5], [0.75, 0.75]],
        4: [[0.27, 0.27], [0.73, 0.27], [0.27, 0.73], [0.73, 0.73]], 5: [[0.27, 0.27], [0.73, 0.27], [0.5, 0.5], [0.27, 0.73], [0.73, 0.73]],
        6: [[0.27, 0.2], [0.73, 0.2], [0.27, 0.5], [0.73, 0.5], [0.27, 0.8], [0.73, 0.8]] };
      counts.forEach((n, i) => {
        const cx = x + i * (size + 16);
        doc.save().lineWidth(1.2).strokeColor(border).rect(cx, y, size, size).stroke().restore();
        const safe = Math.max(1, Math.min(10, n));
        const panels = safe <= 6 ? [[cx, size, safe]] : [[cx, size / 2, 6], [cx + size / 2, size / 2, safe - 6]];
        panels.forEach(([px, pw, k]) => (PAT[k] || []).forEach((p) => {
          doc.save().fillColor(fill).circle(px + pw * p[0], y + size * p[1], Math.min(pw, size) * 0.09).fill().restore();
        }));
      });
      return y + size + 8;
    }
    case "numberTrack": {
      const start = v.start != null ? v.start : 1;
      const end = v.end != null ? v.end : 10;
      const n = end - start + 1;
      const cell = Math.min(S.cell, w / n);
      const hi = new Set(v.highlight || []);
      for (let i = 0; i < n; i += 1) {
        const val = start + i;
        doc.save().lineWidth(1).strokeColor(border);
        if (hi.has(val)) doc.fillColor(P.hex(C.ACCENT)).rect(x + i * cell, y, cell, cell).fillAndStroke();
        else doc.rect(x + i * cell, y, cell, cell).stroke();
        doc.restore();
        doc.save().font("Sans-Bold").fontSize(S.body).fillColor(hi.has(val) ? "#FFFFFF" : border)
          .text(String(val), x + i * cell, y + cell * 0.3, { width: cell, align: "center" }).restore();
      }
      return y + cell + 8;
    }
    case "numberLine": {
      const start = v.start || 0;
      const end = v.end != null ? v.end : 10;
      const step = v.step || 1;
      if (start !== 0) {
        // The twin helper starts at 0; draw a simple line for other ranges.
        const count = Math.round((end - start) / step);
        const lineY = y + 14;
        doc.save().lineWidth(2).strokeColor(border).moveTo(x, lineY).lineTo(x + w, lineY).stroke().restore();
        for (let i = 0; i <= count; i += 1) {
          const tx = x + (w / count) * i;
          doc.save().lineWidth(1.5).strokeColor(border).moveTo(tx, lineY - 7).lineTo(tx, lineY + 7).stroke().restore();
          doc.save().font("Sans").fontSize(S.body - 2).fillColor(border)
            .text(String(Math.round((start + i * step) * 1000) / 1000), tx - 20, lineY + 10, { width: 40, align: "center" }).restore();
        }
        return y + 44;
      }
      return P.addNumberLinePdf(doc, y, end, Math.max(1, Math.round(1 / step)), { x, width: w, labelWholes: true }) + 6;
    }
    case "fractionStrips": {
      const strips = v.strips || [{ denom: v.denom || 2, shaded: v.shaded || 0 }];
      const stripH = byBand(sz, 32, 28, 24);
      const gap = 10;
      const stripW = Math.min(w - 70, 380);
      let yy = y;
      strips.forEach((st, i) => {
        const denom = Math.max(1, st.denom || 1);
        const cw = stripW / denom;
        for (let c = 0; c < denom; c += 1) {
          doc.save().lineWidth(1).strokeColor(border);
          if (c < (st.shaded || 0)) doc.fillColor(P.hex(st.color || [C.PRIMARY, C.SECONDARY, C.ACCENT][i % 3])).rect(x + c * cw, yy, cw, stripH).fillAndStroke();
          else doc.rect(x + c * cw, yy, cw, stripH).stroke();
          doc.restore();
        }
        const label = st.label || (st.shaded ? `${st.shaded}/${denom}` : "1 whole");
        doc.save().font("Sans-Bold").fontSize(S.body).fillColor(border).text(label, x + stripW + 8, yy + stripH / 2 - S.body / 2).restore();
        yy += stripH + gap;
      });
      return yy;
    }
    case "array": return drawDots(doc, x, y, v.rows || 1, v.cols || 1, S.cell * 0.8, fill) + 8;
    case "groupedCounters": {
      const per = v.per || 1;
      const groups = v.groups || 1;
      const dot = S.cell * 0.55;
      const groupW = per * dot + 16;
      let cx = x;
      for (let g = 0; g < groups; g += 1) {
        doc.save().lineWidth(1).strokeColor(P.hex(C.SECONDARY)).roundedRect(cx, y, groupW, dot + 16, 6).stroke().restore();
        for (let i = 0; i < per; i += 1) {
          doc.save().fillColor(P.hex(C.ACCENT)).circle(cx + 8 + i * dot + dot / 2, y + 8 + dot / 2, dot * 0.42).fill().restore();
        }
        cx += groupW + 14;
      }
      return y + dot + 24;
    }
    case "ppwMat": return P.addPpwMatPdf(doc, y, { x, width: Math.min(300, w), whole: v.whole, partA: v.partA, partB: v.partB }) + 6;
    case "hundredGrid": return P.addHundredGridPdf(doc, y, v.shaded || 0, { x, label: v.label }) + 6;
    case "pictogram":
    case "pictograms": {
      const items = v.type === "pictogram" ? [{ name: v.name, label: v.label }] : (v.items || []).map((it) => (typeof it === "string" ? { name: it, label: it } : it));
      const showLabels = v.labels !== false;
      const size = v.type === "pictogram" ? S.cell * 2 : Math.min(S.cell * 1.6, (w - 12 * (items.length - 1)) / items.length);
      // Each icon owns a column wide enough for its label on one line.
      const labelFont = S.body - 2;
      doc.save().font("Sans-Bold").fontSize(labelFont);
      const colW = Math.min(w / Math.max(items.length, 1), Math.max(size + 12, ...items.map((it) =>
        (showLabels && it.label !== "") ? doc.widthOfString(String(it.label != null ? it.label : it.name)) + 12 : 0)));
      doc.restore();
      items.forEach((it, i) => {
        const data = renderPictogramPng(it.name, C.CHARCOAL, 256);
        if (!data) return;
        const buf = Buffer.from(data.split(",")[1], "base64");
        const cx = x + i * colW;
        doc.image(buf, cx + (colW - size * 0.8) / 2, y, { width: size * 0.8 });
        if (showLabels && it.label !== "" && (it.label || v.type === "pictograms")) {
          doc.save().font("Sans-Bold").fontSize(labelFont).fillColor(border)
            .text(String(it.label != null ? it.label : it.name), cx, y + size * 0.85, { width: colW, align: "center", lineBreak: false }).restore();
        }
      });
      return y + size + (showLabels ? 6 : 0);
    }
    case "text": {
      const fontSize = v.fontSize || byBand(sz, 40, 34, 28);
      doc.save().font("Sans-Bold").fontSize(fontSize).fillColor(border).text(String(v.text), x, y, { width: w, align: "center" }).restore();
      return y + fontSize * 1.3;
    }
    case "chips": {
      const items = (v.items || []).map(String);
      const chipH = S.cell * 0.9;
      const chipW = (w - 10 * (items.length - 1)) / Math.max(items.length, 1);
      items.forEach((t, i) => {
        const cx = x + i * (chipW + 10);
        doc.save().lineWidth(1.2).strokeColor(fill).roundedRect(cx, y, chipW, chipH, 6).stroke().restore();
        doc.save().font("Sans-Bold").fontSize(S.prompt).fillColor(fill).text(t, cx, y + chipH / 2 - S.prompt / 2, { width: chipW, align: "center" }).restore();
      });
      return y + chipH + 8;
    }
    case "table": {
      const rows = v.rows || [];
      const cols = Math.max(...rows.map((r) => r.length), 1);
      const cw = w / cols;
      const rh = S.body * 2;
      rows.forEach((r, ri) => {
        for (let c = 0; c < cols; c += 1) {
          const isHeader = ri === 0 && v.header !== false;
          doc.save().lineWidth(0.8).strokeColor(P.hex(C.MUTED));
          if (isHeader) doc.fillColor(fill).rect(x + c * cw, y + ri * rh, cw, rh).fillAndStroke();
          else doc.rect(x + c * cw, y + ri * rh, cw, rh).stroke();
          doc.restore();
          const val = r[c] == null ? "" : String(r[c]);
          doc.save().font(isHeader ? "Sans-Bold" : "Sans").fontSize(S.body).fillColor(isHeader ? "#FFFFFF" : border)
            .text(val, x + c * cw + 4, y + ri * rh + rh / 2 - S.body / 2, { width: cw - 8, align: "center" }).restore();
        }
      });
      return y + rows.length * rh + 8;
    }
    default:
      throw new Error(`[resources] visual type "${v.type}" has no paper twin`);
  }
}

function estimateVisualHeight(visual, sz) {
  const S = bandSizes(sz);
  switch (visual.type) {
    case "tensFrame": return S.cell * 2 + 8;
    case "fiveFrame": return S.cell + 8;
    case "doubleTensFrame": return S.cell * 4 + 18;
    case "dotCard": case "dotCards": return S.cell * 2.4 + 8;
    case "numberTrack": return S.cell + 8;
    case "numberLine": return 50;
    case "fractionStrips": return ((visual.strips || [1]).length) * (byBand(sz, 32, 28, 24) + 10);
    case "array": return (visual.rows || 1) * S.cell * 0.8 + 8;
    case "groupedCounters": return S.cell * 0.55 + 24;
    case "ppwMat": return 130;
    case "hundredGrid": return 160;
    case "pictogram": return S.cell * 2 + 6;
    case "pictograms": return S.cell * 1.6 + 6;
    case "text": return (visual.fontSize || byBand(sz, 40, 34, 28)) * 1.3;
    case "chips": return S.cell * 0.9 + 8;
    case "table": return ((visual.rows || []).length) * S.body * 2 + 8;
    default: return 60;
  }
}

/* ── Page composition ─────────────────────────────────────────────────── */

function drawItem(doc, num, item, y, ctx) {
  const { C, sz, isKey } = ctx;
  const S = bandSizes(sz);
  const primary = P.hex(C.PRIMARY);
  const x = P.PAGE.MARGIN;
  const numD = S.prompt + 10;
  const visual = isKey && item.answerVisual ? item.answerVisual : item.visual;
  const lines = item.answerLines != null ? item.answerLines : (item.answer != null || item.lined ? 1 : 0);
  const needed = numD + 8 + (visual ? estimateVisualHeight(visual, sz) + 6 : 0) + lines * S.lineGap + (item.box ? item.box + 8 : 0) + 14;
  y = fits(doc, y, needed);

  // Number in a circle, prompt beside it
  doc.save().fillColor(primary).circle(x + numD / 2, y + numD / 2, numD / 2).fill().restore();
  doc.save().font("Sans-Bold").fontSize(S.prompt - 2).fillColor("#FFFFFF").text(String(num), x, y + 5, { width: numD, align: "center" }).restore();
  doc.save().font("Sans-Bold").fontSize(S.prompt).fillColor(P.hex(C.CHARCOAL))
    .text(String(item.prompt), x + numD + 10, y + 3, { width: P.PAGE.CONTENT_W - numD - 10 }).restore();
  const promptH = doc.heightOfString(String(item.prompt), { width: P.PAGE.CONTENT_W - numD - 10 });
  y += Math.max(numD, promptH + 4) + 6;

  if (visual) {
    y = drawVisualPdf(doc, visual, x + numD + 10, y, { C, sz, width: P.PAGE.CONTENT_W - numD - 10 }) + 6;
  }
  if (item.box) {
    doc.save().lineWidth(1).strokeColor(P.hex(C.MUTED)).dash(3, { space: 3 }).rect(x + numD + 10, y, P.PAGE.CONTENT_W - numD - 10, item.box).stroke().undash().restore();
    if (isKey && item.answer != null && !lines) {
      doc.save().font("Sans-Bold").fontSize(S.body).fillColor(primary).text(String(item.answer), x + numD + 18, y + 8, { width: P.PAGE.CONTENT_W - numD - 26 }).restore();
    }
    y += item.box + 8;
  }
  for (let i = 0; i < lines; i += 1) {
    const label = i === 0 && item.answerLabel ? item.answerLabel : "";
    const lineX = x + numD + 10;
    const lineW = P.PAGE.CONTENT_W - numD - 10;
    const baseline = y + S.lineGap - 8;
    if (label) doc.save().font("Sans").fontSize(S.body).fillColor(P.hex(C.CHARCOAL)).text(label, lineX, baseline - S.body - 2).restore();
    const startX = label ? lineX + doc.widthOfString(label) + 8 : lineX;
    doc.save().lineWidth(1).strokeColor(P.hex(C.MUTED)).moveTo(startX, baseline).lineTo(lineX + lineW, baseline).stroke().restore();
    if (isKey && i === 0 && item.answer != null) {
      doc.save().font("Sans-Bold").fontSize(S.body).fillColor(primary).text(String(item.answer), startX + 6, baseline - S.body - 2, { width: lineW - 12 }).restore();
    }
    y += S.lineGap;
  }
  return y + 4;
}

function drawBlock(doc, block, y, ctx) {
  const { C, sz } = ctx;
  const S = bandSizes(sz);
  const primary = P.hex(C.PRIMARY);
  const x = P.PAGE.MARGIN;
  // A heading needs room for itself AND the start of what follows, so it
  // never sits alone at the foot of a page.
  if (block.heading) { y = fits(doc, y, 110); return P.addSectionHeading(doc, block.heading, y, { color: primary, fontSize: S.heading }); }
  if (block.text) { y = fits(doc, y, 60); return P.addBodyText(doc, block.text, y, { fontSize: S.body }); }
  if (block.tip) { y = fits(doc, y, 70); return P.addTipBox(doc, block.tip, y, { color: primary }); }
  if (block.steps) { y = fits(doc, y, 90); return P.addStepInstructions(doc, block.steps, y, { color: primary }); }
  if (block.visual) {
    y = fits(doc, y, estimateVisualHeight(block.visual, sz) + 10);
    return drawVisualPdf(doc, block.visual, x, y, { C, sz }) + 8;
  }
  if (block.lines) { y = fits(doc, y, block.lines * S.lineGap); return P.addLinedArea(doc, y, block.lines, { lineSpacing: S.lineGap }); }
  if (block.organiser) {
    const org = block.organiser;
    y = fits(doc, y, 200);
    return P.addTwoColumnOrganiser(doc, org.left || "", org.right || "", y, { color: primary, rows: org.rows || 4, leftContent: org.leftContent, rightContent: org.rightContent });
  }
  if (block.box) {
    const h = typeof block.box === "number" ? block.box : 140;
    y = fits(doc, y, h + 30);
    if (block.label) {
      doc.save().font("Sans-Bold").fontSize(S.body).fillColor(primary).text(String(block.label), x, y).restore();
      y += S.body + 6;
    }
    doc.save().lineWidth(1).strokeColor(P.hex(C.MUTED)).dash(3, { space: 3 }).rect(x, y, P.PAGE.CONTENT_W, h).stroke().undash().restore();
    return y + h + 12;
  }
  return y;
}

function drawCards(doc, resource, y, ctx) {
  const { C, sz } = ctx;
  const S = bandSizes(sz);
  const cols = resource.cols || 2;
  const cards = resource.cards || [];
  const gap = 12;
  const cardW = (P.PAGE.CONTENT_W - gap * (cols - 1)) / cols;
  const cardH = resource.cardH || byBand(sz, 170, 150, 130);
  cards.forEach((card, i) => {
    const col = i % cols;
    if (col === 0) y = fits(doc, y, cardH + gap);
    const cx = P.PAGE.MARGIN + col * (cardW + gap);
    doc.save().lineWidth(1).strokeColor(P.hex(C.MUTED)).dash(4, { space: 3 }).rect(cx, y, cardW, cardH).stroke().undash().restore();
    let inner = y + 14;
    if (card.visual) {
      inner = drawVisualPdf(doc, card.visual, cx + 14, inner, { C, sz, width: cardW - 28 }) + 4;
    }
    if (card.text) {
      doc.save().font("Sans-Bold").fontSize(card.visual ? S.body : S.heading + 4).fillColor(P.hex(C.CHARCOAL))
        .text(String(card.text), cx + 10, card.visual ? inner : y + cardH / 2 - S.heading / 2, { width: cardW - 20, align: "center" }).restore();
    }
    if (col === cols - 1 || i === cards.length - 1) y += cardH + gap;
  });
  return y;
}

async function writeResourcePdf(resource, ctx, filePath, opts) {
  const { C, sz, lessonInfo, footer } = ctx;
  const o = opts || {};
  const doc = P.createPdf({ title: o.title || resource.title || resource.label });
  const primary = P.hex(C.PRIMARY);
  let y = P.addPdfHeader(doc, o.title || resource.title || resource.label, {
    subtitle: resource.subtitle, color: primary, showNameDate: resource.showNameDate !== false && !o.isKey, lessonInfo,
  });
  if (o.isKey) y = P.addTipBox(doc, "Teacher answer key. Answers are shown in colour.", y, { color: primary });
  if (resource.instructions) y = P.addBodyText(doc, resource.instructions, y, { fontSize: bandSizes(sz).body });
  if (resource.steps) y = P.addStepInstructions(doc, resource.steps, y, { color: primary });

  if (resource.kind === "worksheet") {
    (resource.items || []).forEach((item, i) => { y = drawItem(doc, i + 1, item, y, { C, sz, isKey: Boolean(o.isKey) }); });
  } else if (resource.kind === "page") {
    (resource.blocks || []).forEach((block) => { y = drawBlock(doc, block, y, { C, sz }); });
  } else if (resource.kind === "cards") {
    y = drawCards(doc, resource, y, { C, sz });
  }
  if (resource.tip && !o.isKey) { y = fits(doc, y, 70); y = P.addTipBox(doc, resource.tip, y, { color: primary }); }
  P.addPdfFooter(doc, footer, { color: P.hex(C.MUTED) });
  await P.writePdf(doc, filePath);
}

/**
 * Build every resource in the spec. Returns the items for the Teacher
 * Resources slide: [{ name, fileName, description }].
 */
async function buildResources(spec, T, outDir, session) {
  const resources = Array.isArray(spec.resources) ? spec.resources : [];
  const items = [];
  const L = spec.lesson;
  const ctx = {
    C: T.C, sz: T.S,
    lessonInfo: L.meta || "",
    footer: spec.lesson.footer || `${L.title} | ${L.meta || ""}`,
  };
  for (const res of resources) {
    const main = P.makeSessionResource(session, res.label, res.description || "");
    await writeResourcePdf(res, ctx, path.join(outDir, main.fileName));
    items.push(main);
    if (res.kind === "worksheet" && res.answerKey !== false) {
      const keyLabel = res.answerKeyLabel || `${res.label.replace(/\bworksheet\b/i, "").trim()} Answer Key`.replace(/\s+/g, " ");
      const key = P.makeSessionResource(session, keyLabel, res.answerKeyDescription || `Teacher answers for the ${main.name}.`);
      await writeResourcePdf(res, ctx, path.join(outDir, key.fileName), { isKey: true, title: `${res.title || res.label} - Answer Key` });
      items.push(key);
    }
  }
  return items;
}

module.exports = { buildResources, drawVisualPdf, estimateVisualHeight };
