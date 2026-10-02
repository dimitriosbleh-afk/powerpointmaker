"use strict";

/**
 * Worksheet layout engine (megaprompt 86).
 *
 * Staff want sheets that look designed, fit one or two pages (one sheet,
 * double-sided, with page 1 able to stand alone), and give students real
 * room to work. This engine:
 *
 *   1. measures every block (header, section heading, worked example, rows of
 *      question cards) at a compact size;
 *   2. paginates at that size and, for two pages, chooses the break that
 *      balances the pages; more than `maxPages` (default 2) is an error that
 *      names what to cut;
 *   3. hands each page's spare space back: diagrams grow to their preferred
 *      size first, then working areas (squared maths paper or ruled lines)
 *      grow, so no page ends in a blank gap.
 *
 * Question kinds (item.kind): "question" (default: prompt, optional diagram,
 * a/b/c `parts`, working area, answer box), "table" (fill the blanks),
 * "sort" (sort cards into groups), "choice" (circle the answer or the odd one
 * out), "mistake" (find and fix someone's working), "open" (make your own).
 * The answer key renders the same layout with answers filled in.
 */

const P = require("../pdf_helpers");
const { lightenHex } = require("../core/mockups");
const { byBand } = require("../core/gradeBand");
const { diagramPng, diagramSvg, DIAGRAM_TYPES, PAPER } = require("../core/diagrams");

const PAGE_W = 595.28;
const PAGE_H = 841.89;
const M = 34;                       // page margin
const W = PAGE_W - 2 * M;           // content width
const FOOT = 22;                    // footer band
const GAP = 10;                     // gap between cards and blocks
const PAD = 9;                      // card padding

function sizes(sz) {
  return {
    prompt: byBand(sz, 16, 14.5, 12.5),
    small: byBand(sz, 13, 12, 10.5),
    heading: byBand(sz, 16, 15, 13.5),
    badge: byBand(sz, 22, 20, 18),
    line: byBand(sz, 30, 26, 22),      // ruled line spacing
    square: 14.17,                     // 5 mm squared paper
    answerH: byBand(sz, 30, 26, 22),
  };
}

function colours(C) {
  return [C.PRIMARY, C.SECONDARY, C.SUCCESS || C.ACCENT, C.ALERT, C.ACCENT].filter(Boolean).map((c) => P.hex(c));
}
const tint = (hex, t) => P.hex(lightenHex(String(hex).replace("#", ""), t));

function textH(doc, text, font, size, width) {
  doc.save().font(font).fontSize(size);
  const h = doc.heightOfString(String(text), { width });
  doc.restore();
  return h;
}

const KINDS = ["question", "table", "sort", "choice", "mistake", "open"];
function kindOf(item) { return item.kind || "question"; }

/** Work units: how much practice an item gives (parts, table rows, cards count). */
function workUnits(item) {
  switch (kindOf(item)) {
    case "table": return Math.max(1, (item.rows || []).length);
    case "sort": return Math.max(1, Math.ceil((item.cards || []).length / 2));
    case "mistake": return 2;
    case "open": return 2;
    default: return Array.isArray(item.parts) && item.parts.length ? item.parts.length : 1;
  }
}

/* ── Diagram sizing ────────────────────────────────────────────────────── */

function diagramBox(visual, maxW, level) {
  const aspect = diagramSvg(visual).aspect;
  const prefH = visual.maxH || (PAPER[visual.type] || { h: 90 }).h;
  const h = level === "min" ? prefH * 0.72 : level === "max" ? prefH * 1.35 : prefH;
  let w = Math.min(maxW, visual.size || 999, h * aspect);
  return { w, h: w / aspect };
}

/** Cards per row in a sort: as many as fit their widest label. */
function sortPer(doc, item, iw, S) {
  const cards = (item.cards || []).map(String);
  doc.save().font("Sans-Bold").fontSize(S.small + 0.5);
  const need = Math.max(62, ...cards.map((c) => doc.widthOfString(c) + 16));
  doc.restore();
  return Math.max(1, Math.min(cards.length || 1, Math.floor((iw + 6) / (need + 6))));
}

/** Options per row in a choice: as many as fit their widest label. */
function choicePer(doc, item, iw, S) {
  const opts = (item.options || []).map(String);
  doc.save().font("Sans-Bold").fontSize(S.prompt);
  const need = Math.max(40, ...opts.map((o) => doc.widthOfString(o) + 20));
  doc.restore();
  return Math.max(1, Math.min(opts.length || 1, Math.floor((iw + 8) / (need + 8))));
}

/** A part's printed prompt, lettered: measured and drawn from the same string. */
function partText(pt, i) { return `${String.fromCharCode(97 + i)})  ${pt.prompt}`; }

/** Width left for a part's prompt beside its answer box (the box takes up to 45%). */
function partW(iw) { return iw - Math.min(110, iw * 0.45) - 26; }

/** A written-explanation question (ruled lines) needs no answer box unless it names one. */
function hasAnswerBox(item) {
  if (item.answerLabel || item.unit) return true;
  if (item.answerBox === false) return false;
  return !(item.lines || item.answerLines > 1);
}

function isLined(item) {
  const k = kindOf(item);
  return k === "choice" || k === "mistake" || !!(item.lines || item.answerLines > 1);
}

function isDiagram(v) { return v && DIAGRAM_TYPES.includes(v.type); }

/* ── Measuring a card ──────────────────────────────────────────────────── */

/**
 * Fixed (non-growing) height of a card's content, and its minimum working
 * area. `level` is "min" or "pref" for the diagram size.
 */
function measureCard(doc, item, cw, ctx, level) {
  const S = sizes(ctx.sz);
  const iw = cw - 2 * PAD;
  const textW = iw - S.badge - 6;
  const k = kindOf(item);
  let h = PAD + Math.max(S.badge, textH(doc, item.prompt || "", "Sans-Bold", S.prompt, textW)) + 6;
  if (item.hint) h += textH(doc, `Hint: ${item.hint}`, "Sans-Italic", S.small, textW) + 4;
  let workMin = 0;
  let diagramH = 0;
  let diagramPref = 0;
  const v = item.visual;
  if (v) {
    if (isDiagram(v)) {
      diagramH = diagramBox(v, iw, level).h;
      diagramPref = diagramBox(v, iw, "pref").h;
    } else {
      diagramH = diagramPref = require("./resources").estimateVisualHeight(v, ctx.sz, iw);
    }
    h += diagramH + 4;
  }
  if (k === "question") {
    if (Array.isArray(item.parts) && item.parts.length) {
      item.parts.forEach((pt, i) => { h += Math.max(S.answerH, textH(doc, partText(pt, i), "Sans", S.prompt, partW(iw))) + 4; });
    } else if (hasAnswerBox(item)) {
      h += S.answerH + 4;
    }
    const lines = item.lines || (item.answerLines > 1 ? item.answerLines : 0);
    workMin = item.working === "none" ? 0 : (lines ? lines * S.line : (item.workMin || item.box || 34));
  } else if (k === "table") {
    const rows = (item.rows || []).length + 1;
    h += rows * (S.answerH + 2) + 4;
    workMin = 0;
  } else if (k === "sort") {
    const chipsH = Math.ceil((item.cards || []).length / sortPer(doc, item, iw, S)) * (S.answerH + 4);
    h += chipsH + 22;
    workMin = 50;
  } else if (k === "choice") {
    const rows = Math.ceil((item.options || []).length / choicePer(doc, item, iw, S));
    h += rows * (S.answerH + 10);
    workMin = item.reason ? 2 * S.line : 0;
  } else if (k === "mistake") {
    const work = (item.work || []).join("\n");
    h += textH(doc, work, "Sans-Italic", S.prompt, iw - 24) + 22 + S.small + 6;
    workMin = 2 * S.line;
  } else if (k === "open") {
    workMin = item.workMin || 80;
  }
  return { fixed: h + PAD, workMin, diagramH, diagramPref, grows: workMin > 0 };
}

/* ── Building blocks ───────────────────────────────────────────────────── */

/**
 * Worked example geometry, shared by measuring and drawing. The steps run
 * across the band as numbered chips (a left-to-right flow reads as "do this,
 * then this"), falling back to a list when a chip would be too narrow.
 */
const EX_VIS_W = 110;
function exampleLayout(doc, ex, ctx) {
  const S = sizes(ctx.sz);
  const hasVis = ex.visual && isDiagram(ex.visual);
  const vis = hasVis ? diagramBox(Object.assign({}, ex.visual, { maxH: 78 }), EX_VIS_W, "pref") : null;
  const tx = M + 16 + (hasVis ? EX_VIS_W + 12 : 0);
  const textW = W - 28 - (hasVis ? EX_VIS_W + 12 : 0);
  const head = `Worked example: ${ex.prompt || ""}`.trim();
  const headH = textH(doc, head, "Sans-Bold", S.small + 0.5, textW);
  const steps = (ex.steps || []).map(String);
  const chipW = steps.length ? (textW - 6 * (steps.length - 1)) / steps.length : textW;
  const across = steps.length > 1 && chipW >= 118;
  let stepsH;
  if (across) stepsH = Math.max(...steps.map((st) => textH(doc, st, "Sans", S.small, chipW - 24))) + 10;
  else stepsH = steps.reduce((t, st) => t + textH(doc, st, "Sans", S.small, textW - 16) + 2, 0);
  const ansH = ex.answer ? textH(doc, `So: ${ex.answer}`, "Sans-Bold", S.small, textW) + 4 : 0;
  const textTotal = 8 + headH + 5 + stepsH + 5 + ansH + 6;
  const h = Math.max(textTotal, hasVis ? vis.h + 16 : 0);
  return { S, hasVis, vis, tx, textW, head, headH, steps, chipW, across, stepsH, h };
}

function drawExample(doc, ex, y, colour, h, ctx) {
  const L = exampleLayout(doc, ex, ctx);
  const S = L.S;
  doc.save().roundedRect(M, y, W, h, 8).fill(tint(colour, 0.9)).restore();
  doc.save().rect(M, y, 4, h).fill(colour).restore();
  if (L.hasVis) drawVisual(doc, Object.assign({}, ex.visual, { maxH: 78 }), M + 12, y + (h - L.vis.h) / 2, EX_VIS_W, "pref", ctx);
  let ty = y + 8;
  doc.save().font("Sans-Bold").fontSize(S.small + 0.5).fillColor(colour).text("Worked example: ", L.tx, ty, { continued: true, width: L.textW })
    .fillColor("#1F2530").text(String(ex.prompt || "")).restore();
  ty += L.headH + 5;
  if (L.across) {
    L.steps.forEach((st, i) => {
      const cx = L.tx + i * (L.chipW + 6);
      doc.save().roundedRect(cx, ty, L.chipW, L.stepsH, 6).fill("#FFFFFF").restore();
      doc.save().fillColor(colour).circle(cx + 10, ty + 10, 7).fill().restore();
      doc.save().font("Sans-Bold").fontSize(9).fillColor("#FFFFFF").text(String(i + 1), cx + 3, ty + 5.5, { width: 14, align: "center", lineBreak: false }).restore();
      doc.save().font("Sans").fontSize(S.small).fillColor("#1F2530").text(st, cx + 21, ty + 5, { width: L.chipW - 24 }).restore();
    });
  } else {
    L.steps.forEach((st, i) => {
      doc.save().font("Sans-Bold").fontSize(S.small).fillColor(colour).text(`${i + 1}`, L.tx, ty, { width: 12, lineBreak: false }).restore();
      doc.save().font("Sans").fontSize(S.small).fillColor("#1F2530").text(st, L.tx + 16, ty, { width: L.textW - 16 }).restore();
      ty += textH(doc, st, "Sans", S.small, L.textW - 16) + 2;
    });
    ty -= L.stepsH;
  }
  ty += L.stepsH + 5;
  if (ex.answer) doc.save().font("Sans-Bold").fontSize(S.small).fillColor(colour).text(`So: ${ex.answer}`, L.tx, ty, { width: L.textW }).restore();
}

function buildBlocks(doc, resource, ctx) {
  const S = sizes(ctx.sz);
  const blocks = [];
  const secs = Array.isArray(resource.sections) ? resource.sections
    : [{ title: null, items: resource.items || [], columns: resource.columns || 2 }];
  let num = 0;
  secs.forEach((sec, si) => {
    const colour = colours(ctx.C)[si % colours(ctx.C).length];
    const secCols = sec.columns || 2;
    const group = [];
    if (sec.title) group.push({ type: "sectionHead", sec, si, colour, h: S.heading + 12, breakBefore: sec.pageBreakBefore });
    if (sec.intro) group.push({ type: "intro", sec, colour, h: textH(doc, sec.intro, "Sans", S.small, W) + 6 });
    if (sec.example) group.push({ type: "example", sec, colour, h: exampleLayout(doc, sec.example, ctx).h + 6 });
    // Tables and sorts need the page width: they take a row of their own
    // (or any item with span: "full"); other items fill the section's columns.
    const items = sec.items || [];
    const fullWidth = (it) => it && (it.span === "full" || (it.span == null && (kindOf(it) === "table" || kindOf(it) === "sort")));
    const rows = [];
    let pending = [];
    items.forEach((it) => {
      if (fullWidth(it) && secCols > 1) {
        if (pending.length) rows.push({ cols: secCols, items: pending });
        pending = [];
        rows.push({ cols: 1, items: [it] });
      } else {
        pending.push(it);
        if (pending.length === secCols) { rows.push({ cols: secCols, items: pending }); pending = []; }
      }
    });
    if (pending.length) rows.push({ cols: secCols, items: pending });
    rows.forEach((r) => {
      const cols = r.cols;
      const cw = (W - GAP * (cols - 1)) / cols;
      const row = r.items.map((it) => { num += 1; return { item: it, num }; });
      const mins = row.map((r) => measureCard(doc, r.item, cw, ctx, "min"));
      const prefs = row.map((r) => measureCard(doc, r.item, cw, ctx, "pref"));
      const maxes = row.map((r) => measureCard(doc, r.item, cw, ctx, "max"));
      const minH = Math.max(...mins.map((m) => m.fixed + m.workMin));
      const prefH = Math.max(...prefs.map((m) => m.fixed + m.workMin));
      const maxH = Math.max(...maxes.map((m) => m.fixed + m.workMin));
      // A row whose only growing space is ruled lines stops at a few extra lines:
      // ten empty lines under "How do you know?" reads as a mistake, not room.
      const lined = row.every((r, j) => !mins[j].grows || isLined(r.item));
      group.push({ type: "row", sec, colour, cols, cw, cells: row, minH: minH + GAP, prefH: prefH + GAP, maxH: maxH + GAP, grows: mins.some((m) => m.grows), workCap: lined ? 4 * S.line : 220 });
    });
    // Keep a heading with its intro, example and first row of questions.
    const lead = group.findIndex((b) => b.type === "row");
    group.forEach((b, i) => { b.keepWithNext = i < lead; if (b.h != null) { b.minH = b.h; b.prefH = b.h; } });
    blocks.push(...group);
  });
  if (resource.tip && !ctx.isKey) {
    const h = textH(doc, resource.tip, "Sans-Bold", S.small, W - 40) + 18;
    blocks.push({ type: "tip", colour: colours(ctx.C)[0], h, minH: h, prefH: h });
  }
  return blocks;
}

/* ── Pagination ────────────────────────────────────────────────────────── */

function headerHeight(page, resource, ctx) {
  const S = sizes(ctx.sz);
  if (page > 0) return 26;
  return 44 + 26 + (resource.subtitle ? S.small + 8 : 0);
}
function capacity(page, resource, ctx) { return PAGE_H - 2 * M - FOOT - headerHeight(page, resource, ctx); }

/** Legal break points: before block i, never splitting a keep-with-next run. */
function legalBreaks(blocks) {
  const out = [];
  for (let i = 1; i < blocks.length; i += 1) if (!blocks[i - 1].keepWithNext) out.push(i);
  return out;
}

function paginate(blocks, resource, ctx) {
  const sum = (a, b) => blocks.slice(a, b).reduce((t, x) => t + x.minH, 0);
  const forced = blocks.map((b, i) => (b.breakBefore && i > 0 ? i : null)).filter((x) => x != null);
  const breaks = legalBreaks(blocks);
  // Greedy at compact size to count pages.
  const pages = [];
  let start = 0;
  while (start < blocks.length) {
    const cap = capacity(pages.length, resource, ctx);
    let end = start;
    let best = null;
    for (let i = start + 1; i <= blocks.length; i += 1) {
      if (i < blocks.length && !breaks.includes(i)) continue;
      if (forced.some((f) => f > start && f < i)) break;
      if (sum(start, i) <= cap) best = i; else break;
    }
    if (best == null) {
      // A single unbreakable run taller than a page: place it and let it be squeezed.
      best = breaks.find((b) => b > start) || blocks.length;
    }
    end = best;
    pages.push([start, end]);
    start = end;
  }
  // Two pages: move the break to balance the pages, so both read as full.
  if (pages.length === 2 && !forced.length) {
    let bestSplit = pages[0][1];
    let bestDiff = Infinity;
    breaks.forEach((b) => {
      const a = sum(0, b);
      const c = sum(b, blocks.length);
      if (a <= capacity(0, resource, ctx) && c <= capacity(1, resource, ctx)) {
        const diff = Math.abs(a / capacity(0, resource, ctx) - c / capacity(1, resource, ctx));
        if (diff < bestDiff) { bestDiff = diff; bestSplit = b; }
      }
    });
    return [[0, bestSplit], [bestSplit, blocks.length]];
  }
  return pages;
}

/** Share a page's spare height: diagrams to preferred size first, then working space. */
function grow(pageBlocks, cap) {
  let spare = cap - pageBlocks.reduce((t, b) => t + b.minH, 0);
  pageBlocks.forEach((b) => { b.useH = b.minH; b.level = "min"; b.extraWork = 0; });
  pageBlocks.filter((b) => b.type === "row").forEach((b) => {
    const need = b.prefH - b.minH;
    if (need > 0 && spare >= need) { b.useH = b.prefH; b.level = "pref"; spare -= need; }
  });
  // Working space: share the spare evenly, re-sharing what capped rows leave.
  let open = pageBlocks.filter((b) => b.type === "row" && b.grows);
  while (open.length && spare > 0.5) {
    const each = spare / open.length;
    open.forEach((b) => {
      const add = Math.min(each, b.workCap - b.extraWork);
      b.extraWork += add; b.useH += add; spare -= add;
    });
    open = open.filter((b) => b.extraWork < b.workCap - 0.5);
  }
  // Still room (a page of measuring, nothing to write): diagrams grow past their usual size.
  pageBlocks.filter((b) => b.type === "row" && b.level === "pref" && b.maxH > b.prefH).forEach((b) => {
    const need = b.maxH - b.prefH;
    if (spare >= need) { b.useH += need; b.level = "max"; spare -= need; }
  });
  // Anything left is air inside the cards, so a page never ends in a blank band.
  // Rows of ruled lines already stopped at their cap, so they take no air.
  const rows = pageBlocks.filter((b) => b.type === "row" && !(b.grows && b.workCap < 220));
  if (rows.length && spare > 0) {
    const air = Math.min(spare / rows.length, 60);
    rows.forEach((b) => { b.useH += air; });
    spare -= air * rows.length;
  }
  return Math.max(0, spare);
}

/* ── Drawing ───────────────────────────────────────────────────────────── */

function drawSquares(doc, x, y, w, h, S) {
  if (h < 8) return;
  doc.save().lineWidth(0.35).strokeColor("#D9DEE7");
  for (let gx = x; gx <= x + w + 0.1; gx += S.square) doc.moveTo(gx, y).lineTo(gx, y + h);
  for (let gy = y; gy <= y + h + 0.1; gy += S.square) doc.moveTo(x, gy).lineTo(x + w, gy);
  doc.stroke().restore();
}
/** The largest size (down to 8pt) at which one line of text fits the width. */
function fitSize(doc, font, text, size, width) {
  let fs = size;
  doc.save().font(font);
  while (fs > 8 && doc.fontSize(fs).widthOfString(String(text)) > width) fs -= 0.5;
  doc.restore();
  return fs;
}

/** Key text written into a working area; on ruled lines each text line sits on a rule. */
function keyText(doc, text, x, y, w, S, colour, ruled, maxH) {
  doc.save().font("Sans-Bold").fontSize(S.small).fillColor(colour);
  if (ruled && maxH != null) {
    const lh = doc.currentLineHeight(true);
    const lines = Math.ceil(doc.heightOfString(String(text), { width: w }) / lh);
    if (S.line * lines > maxH + 2) ruled = false;
  }
  if (ruled) {
    const lh = doc.currentLineHeight(true);
    doc.text(String(text), x, y + S.line - lh - 1, { width: w, lineGap: Math.max(0, S.line - lh) });
  } else {
    doc.text(String(text), x, y + 4, { width: w });
  }
  doc.restore();
}
function drawLines(doc, x, y, w, h, S) {
  doc.save().lineWidth(0.6).strokeColor("#B8C0CC");
  for (let gy = y + S.line; gy <= y + h + 0.1; gy += S.line) doc.moveTo(x, gy).lineTo(x + w, gy);
  doc.stroke().restore();
}

/** True when a key answer is too long for the answer box and belongs in the working area. */
function answerTooLong(doc, answer, boxW, S) {
  if (answer == null) return false;
  doc.save().font("Sans-Bold").fontSize(S.small);
  const long = doc.widthOfString(String(answer)) > boxW - 8;
  doc.restore();
  return long;
}

function answerBox(doc, x, y, w, S, colour, label, unit, answer, isKey) {
  const boxW = Math.min(110, w * 0.45);
  doc.save().font("Sans-Bold").fontSize(S.small).fillColor(colour);
  const labelText = label || "Answer";
  const lw = doc.widthOfString(labelText) + 6;
  const unitW = unit ? doc.widthOfString(unit) + 4 : 0;
  const bx = x + w - boxW - unitW;
  doc.text(labelText, bx - lw, y + (S.answerH - S.small) / 2 - 1, { lineBreak: false });
  doc.restore();
  doc.save().lineWidth(1).strokeColor(colour).roundedRect(bx, y, boxW, S.answerH - 4, 5).stroke().restore();
  if (unit) doc.save().font("Sans-Bold").fontSize(S.small).fillColor("#333333").text(unit, bx + boxW + 3, y + (S.answerH - S.small) / 2 - 1, { lineBreak: false }).restore();
  if (isKey && answer != null && !answerTooLong(doc, answer, boxW, S)) {
    doc.save().font("Sans-Bold").fontSize(S.small).fillColor(colour).text(String(answer), bx + 4, y + (S.answerH - 4 - S.small) / 2 - 1, { width: boxW - 8, align: "center", lineBreak: false }).restore();
  }
}

function drawVisual(doc, v, x, y, maxW, level, ctx) {
  if (isDiagram(v)) {
    const box = diagramBox(v, maxW, level);
    const { buffer } = diagramPng(v, { ink: ctx.C.CHARCOAL, accent: ctx.C.PRIMARY }, 1100);
    doc.image(buffer, x + (maxW - box.w) / 2, y, { width: box.w });
    return box.h;
  }
  const endY = require("./resources").drawVisualPdf(doc, v, x, y, { C: ctx.C, sz: ctx.sz, width: maxW });
  return endY - y;
}

function drawCard(doc, cell, x, y, cw, h, block, ctx) {
  const S = sizes(ctx.sz);
  const { item, num } = cell;
  const colour = block.colour;
  const isKey = ctx.isKey;
  const k = kindOf(item);
  const iw = cw - 2 * PAD;
  doc.save().lineWidth(0.9).strokeColor(tint(colour, 0.55)).roundedRect(x, y, cw, h, 7).stroke().restore();
  // Number badge and prompt.
  doc.save().fillColor(colour).circle(x + PAD + S.badge / 2, y + PAD + S.badge / 2, S.badge / 2).fill().restore();
  doc.save().font("Sans-Bold").fontSize(S.badge * 0.55).fillColor("#FFFFFF").text(String(num), x + PAD, y + PAD + S.badge * 0.2, { width: S.badge, align: "center", lineBreak: false }).restore();
  const tx = x + PAD + S.badge + 6;
  const tw = iw - S.badge - 6;
  doc.save().font("Sans-Bold").fontSize(S.prompt).fillColor("#1F2530").text(String(item.prompt || ""), tx, y + PAD + 1, { width: tw }).restore();
  let cy = y + PAD + Math.max(S.badge, textH(doc, item.prompt || "", "Sans-Bold", S.prompt, tw)) + 6;
  if (item.hint) {
    const hint = `Hint: ${item.hint}`;
    doc.save().font("Sans-Italic").fontSize(S.small).fillColor(colour).text(hint, tx, cy - 2, { width: tw }).restore();
    cy += textH(doc, hint, "Sans-Italic", S.small, tw) + 4;
  }
  const bottom = y + h - PAD;

  const shown = isKey && item.answerVisual ? item.answerVisual : item.visual;
  if (shown) cy += drawVisual(doc, shown, x + PAD, cy, iw, block.level, ctx) + 4;

  if (k === "question") {
    const parts = Array.isArray(item.parts) ? item.parts : [];
    const answerRows = parts.length ? parts.reduce((t, pt, i) => t + Math.max(S.answerH, textH(doc, partText(pt, i), "Sans", S.prompt, partW(iw))) + 4, 0)
      : (hasAnswerBox(item) ? S.answerH + 4 : 0);
    const workH = bottom - cy - answerRows;
    if (item.working !== "none" && workH > 6) {
      if (item.lines || item.answerLines > 1) drawLines(doc, x + PAD, cy, iw, workH, S); else drawSquares(doc, x + PAD, cy, iw, workH, S);
      // Key: an answer with no box, or too long for its box, is written in the working area.
      const boxW = Math.min(110, iw * 0.45);
      if (isKey && item.answer != null && !parts.length && (!hasAnswerBox(item) || answerTooLong(doc, item.answer, boxW, S))) {
        keyText(doc, item.answer, x + PAD + 4, cy, iw - 8, S, colour, !!(item.lines || item.answerLines > 1), workH);
      }
    }
    let ay = bottom - answerRows;
    if (parts.length) {
      parts.forEach((pt, i) => {
        const ph = Math.max(S.answerH, textH(doc, partText(pt, i), "Sans", S.prompt, partW(iw)));
        doc.save().font("Sans").fontSize(S.prompt).fillColor("#1F2530").text(partText(pt, i), x + PAD, ay + 3, { width: partW(iw) }).restore();
        answerBox(doc, x + PAD, ay, iw, S, colour, " ", pt.unit || item.unit, pt.answer, isKey);
        ay += ph + 4;
      });
    } else if (hasAnswerBox(item)) {
      answerBox(doc, x + PAD, ay, iw, S, colour, item.answerLabel, item.unit, item.answer, isKey);
    }
  } else if (k === "table") {
    const cols = item.columns || [];
    const rows = item.rows || [];
    const cwid = iw / Math.max(cols.length, 1);
    const rh = S.answerH + 2;
    cols.forEach((c, j) => {
      doc.save().rect(x + PAD + j * cwid, cy, cwid, rh).fillAndStroke(tint(colour, 0.8), tint(colour, 0.4)).restore();
      const fs = fitSize(doc, "Sans-Bold", c, S.small, cwid - 6);
      doc.save().font("Sans-Bold").fontSize(fs).fillColor("#1F2530").text(String(c), x + PAD + j * cwid + 3, cy + (rh - fs) / 2 - 1, { width: cwid - 6, align: "center", lineBreak: false }).restore();
    });
    rows.forEach((r, i) => {
      const ry = cy + (i + 1) * rh;
      r.forEach((cell, j) => {
        doc.save().lineWidth(0.7).strokeColor(tint(colour, 0.4)).rect(x + PAD + j * cwid, ry, cwid, rh).stroke().restore();
        const given = cell != null && cell !== "";
        const keyVal = item.answers && item.answers[i] ? item.answers[i][j] : null;
        const val = given ? cell : (isKey ? keyVal : null);
        if (val != null && val !== "") {
          const fs = fitSize(doc, given ? "Sans" : "Sans-Bold", val, S.small, cwid - 6);
          doc.save().font(given ? "Sans" : "Sans-Bold").fontSize(fs).fillColor(given ? "#1F2530" : colour)
            .text(String(val), x + PAD + j * cwid + 3, ry + (rh - fs) / 2 - 1, { width: cwid - 6, align: "center", lineBreak: false }).restore();
        }
      });
    });
  } else if (k === "sort") {
    const cards = item.cards || [];
    const per = sortPer(doc, item, iw, S);
    const chipW = (iw + 6) / per - 6;
    cards.forEach((c, i) => {
      const cx = x + PAD + (i % per) * (chipW + 6);
      const ccy = cy + Math.floor(i / per) * (S.answerH + 4);
      doc.save().lineWidth(0.9).strokeColor(colour).fillColor(tint(colour, 0.9)).roundedRect(cx, ccy, chipW, S.answerH, 6).fillAndStroke().restore();
      doc.save().font("Sans-Bold").fontSize(S.small + 0.5).fillColor("#1F2530").text(String(c), cx, ccy + (S.answerH - S.small) / 2 - 1, { width: chipW, align: "center", lineBreak: false }).restore();
    });
    cy += Math.ceil(cards.length / per) * (S.answerH + 4) + 4;
    const groups = item.groups || [];
    const gw = (iw - 6 * (groups.length - 1)) / Math.max(groups.length, 1);
    const gh = bottom - cy;
    groups.forEach((g, j) => {
      const gx = x + PAD + j * (gw + 6);
      doc.save().roundedRect(gx, cy, gw, 18, 5).fill(colour).restore();
      doc.save().font("Sans-Bold").fontSize(S.small).fillColor("#FFFFFF").text(String(g), gx, cy + 4, { width: gw, align: "center", lineBreak: false }).restore();
      doc.save().lineWidth(0.8).dash(3, { space: 3 }).strokeColor(tint(colour, 0.4)).rect(gx, cy + 18, gw, gh - 18).stroke().undash().restore();
      if (isKey && item.answer && item.answer[g]) {
        doc.save().font("Sans-Bold").fontSize(S.small).fillColor(colour).text(item.answer[g].join("\n"), gx + 4, cy + 24, { width: gw - 8, align: "center" }).restore();
      }
    });
  } else if (k === "choice") {
    const opts = item.options || [];
    const per = choicePer(doc, item, iw, S);
    const ow = (iw - 8 * (per - 1)) / per;
    const correct = [].concat(item.answer == null ? [] : item.answer);
    opts.forEach((o, i) => {
      const ox = x + PAD + (i % per) * (ow + 8);
      const oy = cy + Math.floor(i / per) * (S.answerH + 10);
      doc.save().lineWidth(0.9).strokeColor(tint(colour, 0.3)).roundedRect(ox, oy, ow, S.answerH + 4, 8).stroke().restore();
      doc.save().font("Sans-Bold").fontSize(S.prompt).fillColor("#1F2530").text(String(o), ox, oy + (S.answerH + 4 - S.prompt) / 2 - 1, { width: ow, align: "center", lineBreak: false }).restore();
      if (isKey && correct.includes(i)) doc.save().lineWidth(2.2).strokeColor(colour).ellipse(ox + ow / 2, oy + (S.answerH + 4) / 2, ow / 2 - 2, (S.answerH + 4) / 2 + 3).stroke().restore();
    });
    cy += Math.ceil(opts.length / per) * (S.answerH + 10);
    if (item.reason) {
      doc.save().font("Sans-Italic").fontSize(S.small).fillColor("#5B6472").text("Because...", x + PAD, cy + 2, { lineBreak: false }).restore();
      drawLines(doc, x + PAD, cy, iw, bottom - cy, S);
      if (isKey && item.explain) keyText(doc, item.explain, x + PAD + 70, cy, iw - 74, S, colour, true, bottom - cy);
    }
  } else if (k === "mistake") {
    const work = (item.work || []).join("\n");
    const nh = textH(doc, work, "Sans-Italic", S.prompt, iw - 24) + 16;
    doc.save().roundedRect(x + PAD, cy, iw, nh, 6).fill("#FFF7D6").restore();
    doc.save().rect(x + PAD, cy, 4, nh).fill("#E0B400").restore();
    doc.save().font("Sans-Italic").fontSize(S.prompt).fillColor("#3A3320").text(work, x + PAD + 14, cy + 8, { width: iw - 24 }).restore();
    cy += nh + 6;
    doc.save().font("Sans-Bold").fontSize(S.small).fillColor(colour).text("Find it and fix it:", x + PAD, cy, { lineBreak: false }).restore();
    cy += S.small + 4;
    drawLines(doc, x + PAD, cy, iw, bottom - cy, S);
    if (isKey && item.answer) keyText(doc, item.answer, x + PAD + 2, cy, iw - 4, S, colour, true, bottom - cy);
  } else if (k === "open") {
    drawSquares(doc, x + PAD, cy, iw, bottom - cy, S);
    if (isKey && item.answer) doc.save().font("Sans-Bold").fontSize(S.small).fillColor(colour).text(String(item.answer), x + PAD + 4, cy + 4, { width: iw - 8 }).restore();
  }
}

function drawHeader(doc, page, pageCount, resource, ctx) {
  const S = sizes(ctx.sz);
  const c = colours(ctx.C)[0];
  const title = ctx.isKey ? `${resource.title || resource.label} - Answer Key` : (resource.title || resource.label);
  if (page > 0) {
    doc.save().font("Sans-Bold").fontSize(11).fillColor(c).text(title, M, M, { width: W, lineBreak: false }).restore();
    doc.save().lineWidth(1).strokeColor(tint(c, 0.6)).moveTo(M, M + 17).lineTo(M + W, M + 17).stroke().restore();
    return M + 26;
  }
  doc.save().roundedRect(M, M, W, 38, 9).fill(c).restore();
  let fs = 18;
  doc.save().font("Sans-Bold");
  while (fs > 12 && doc.fontSize(fs).widthOfString(title) > W - 164) fs -= 0.5;
  doc.restore();
  doc.save().font("Sans-Bold").fontSize(fs).fillColor("#FFFFFF").text(title, M + 14, M + 19 - fs * 0.6, { lineBreak: false }).restore();
  if (ctx.lessonInfo) doc.save().font("Sans").fontSize(9).fillColor("#FFFFFF").text(ctx.lessonInfo, M + W - 150, M + 14, { width: 136, align: "right", lineBreak: false }).restore();
  let y = M + 44;
  if (!ctx.isKey) {
    doc.save().font("Sans-Bold").fontSize(11.5).fillColor("#1F2530").text("Name", M, y + 6, { lineBreak: false }).restore();
    doc.save().lineWidth(0.8).strokeColor("#9AA3B2").moveTo(M + 40, y + 19).lineTo(M + 300, y + 19).stroke().restore();
    doc.save().font("Sans-Bold").fontSize(11.5).fillColor("#1F2530").text("Date", M + 320, y + 6, { lineBreak: false }).restore();
    doc.save().lineWidth(0.8).strokeColor("#9AA3B2").moveTo(M + 352, y + 19).lineTo(M + W, y + 19).stroke().restore();
  } else {
    doc.save().font("Sans-Bold").fontSize(11).fillColor(c).text("Teacher copy: answers in colour. Accept equivalent answers and working.", M, y + 6, { lineBreak: false }).restore();
  }
  y += 26;
  if (resource.subtitle) {
    doc.save().font("Sans-Italic").fontSize(S.small).fillColor("#4A5260").text(resource.subtitle, M, y, { width: W }).restore();
    y += S.small + 8;
  }
  return y;
}

function drawFooter(doc, page, pageCount, ctx) {
  doc.save().font("Sans").fontSize(8).fillColor("#9AA3B2")
    .text(ctx.footer || "", M, PAGE_H - M - 10, { width: W - 60, lineBreak: false }).restore();
  doc.save().font("Sans").fontSize(8).fillColor("#9AA3B2")
    .text(`Page ${page + 1} of ${pageCount}`, M + W - 60, PAGE_H - M - 10, { width: 60, align: "right", lineBreak: false }).restore();
}

function drawBlock(doc, b, y, ctx) {
  const S = sizes(ctx.sz);
  if (b.type === "sectionHead") {
    doc.save().fillColor(b.colour).circle(M + 10, y + 11, 10).fill().restore();
    doc.save().font("Sans-Bold").fontSize(11).fillColor("#FFFFFF").text(String.fromCharCode(65 + b.si), M, y + 5, { width: 20, align: "center", lineBreak: false }).restore();
    doc.save().font("Sans-Bold").fontSize(S.heading).fillColor(b.colour).text(String(b.sec.title), M + 26, y + 3, { lineBreak: false }).restore();
    doc.save().font("Sans-Bold").fontSize(S.heading);
    const tw = doc.widthOfString(String(b.sec.title));
    doc.restore();
    doc.save().lineWidth(1.2).strokeColor(tint(b.colour, 0.55)).moveTo(M + 34 + tw, y + 11).lineTo(M + W, y + 11).stroke().restore();
  } else if (b.type === "intro") {
    doc.save().font("Sans").fontSize(S.small).fillColor("#3A4250").text(String(b.sec.intro), M, y, { width: W }).restore();
  } else if (b.type === "example") {
    drawExample(doc, b.sec.example, y, b.colour, b.useH - 6, ctx);
  } else if (b.type === "tip") {
    doc.save().roundedRect(M, y + 4, W, b.useH - 8, 8).fill(tint(b.colour, 0.9)).restore();
    doc.save().font("Sans-Bold").fontSize(S.small).fillColor(b.colour).text(String(ctx.tip || ""), M + 14, y + 12, { width: W - 28 }).restore();
  } else if (b.type === "row") {
    const h = b.useH - GAP;
    b.cells.forEach((cell, i) => drawCard(doc, cell, M + i * (b.cw + GAP), y, b.cw, h, b, ctx));
  }
}

/**
 * Lay out and draw a worksheet. Throws when the sheet cannot fit its page
 * limit even at compact sizes. Returns { pages }.
 */
function renderWorksheet(doc, resource, ctx) {
  const blocks = buildBlocks(doc, resource, ctx);
  const pages = paginate(blocks, resource, ctx);
  if (process.env.WS_DEBUG) blocks.forEach((b) => console.log(`[ws] ${b.type} min=${Math.round(b.minH)} pref=${Math.round(b.prefH)}${b.cells ? " cols=" + b.cols : ""}`));
  const maxPages = resource.maxPages || 2;
  if (pages.length > maxPages) {
    const units = blocks.filter((b) => b.type === "row").reduce((t, b) => t + b.cells.reduce((u, c) => u + workUnits(c.item), 0), 0);
    throw new Error(`[worksheet] "${resource.label}" needs ${pages.length} pages even at compact sizes; the limit is ${maxPages} (one sheet, double-sided). It has ${units} work units: cut or combine questions (use parts a/b/c, a table or a sort), use 3 columns for short items, or set maxPages with a reason (megaprompt 86).`);
  }
  pages.forEach(([a, b], p) => {
    if (p > 0) doc.addPage({ size: "A4", margin: 0 });
    let y = drawHeader(doc, p, pages.length, resource, ctx);
    const pageBlocks = blocks.slice(a, b);
    const spare = grow(pageBlocks, capacity(p, resource, ctx));
    // Any spare left after working areas are generous goes between sections.
    const heads = pageBlocks.filter((bl, i) => bl.type === "sectionHead" && i > 0).length;
    const between = heads ? Math.min(18, spare / heads) : 0;
    pageBlocks.forEach((bl, i) => {
      if (bl.type === "sectionHead" && i > 0) y += between;
      drawBlock(doc, bl, y, ctx);
      y += bl.useH;
    });
    drawFooter(doc, p, pages.length, ctx);
  });
  return { pages: pages.length };
}

module.exports = { renderWorksheet, workUnits, KINDS };
