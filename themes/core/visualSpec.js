"use strict";

/**
 * Declarative visuals (megaprompt section 15a / 16a).
 *
 * A visual spec names a representation and its values; the theme decides the
 * size and placement so the representation fills the space it is given:
 *
 *   { type: "tensFrame", filled: 7 }
 *   { type: "numberLine", start: 0, end: 2, step: 1/3, marked: [1] }
 *   { type: "fractionStrips", strips: [{ denom: 4, shaded: 1 }] }
 *   { type: "pictograms", items: ["happy", "sad", "angry"] }
 *   { type: "text", text: "9" }
 *   { type: "custom", draw: (slide, frame) => { ... } }
 *
 * drawVisual(slide, spec, frame) fits the spec into frame {x, y, w, h},
 * centres it, draws it with the shared helpers, and returns the drawn bounds.
 * Builders such as heroVisualSlide, choiceSlide, youDoSlide and the
 * drawRight column of contentSlide / workedExSlide accept a spec directly, so
 * a build script never has to compute coordinates for a representation the
 * theme already knows how to draw.
 */

const { SAFE_BOTTOM } = require("./layout");
const { DEFAULT_SIZES, byBand } = require("./gradeBand");
const { getImageDimensions } = require("./images");

const SUPPORTED_TYPES = [
  "tensFrame", "fiveFrame", "doubleTensFrame", "dotCard", "dotCards", "numberTrack",
  "numberLine", "fractionStrips", "array", "baseTen", "groupedCounters", "ppwMat",
  "chips", "pictogram", "pictograms", "text", "image", "table", "custom",
];

function createVisualSpec(C, FONT_H, FONT_B, el, S, deps) {
  const sz = S || DEFAULT_SIZES;
  const manips = (deps && deps.manips) || {};
  const picto = (deps && deps.picto) || {};

  function centreBox(frame, w, h) {
    return {
      x: frame.x + (frame.w - w) / 2,
      y: frame.y + (frame.h - h) / 2,
      w, h,
    };
  }

  function normaliseFrame(frame) {
    const f = frame || {};
    return {
      x: Number.isFinite(f.x) ? f.x : 0.5,
      y: Number.isFinite(f.y) ? f.y : 1.3,
      w: Number.isFinite(f.w) && f.w > 0 ? f.w : 9,
      h: Number.isFinite(f.h) && f.h > 0 ? f.h : (SAFE_BOTTOM - 1.3),
    };
  }

  function buildNumberLineLabels(spec) {
    if (Array.isArray(spec.labels) && spec.labels.length >= 2) return spec.labels;
    const start = Number(spec.start) || 0;
    const end = Number(spec.end);
    if (!Number.isFinite(end) || end <= start) return ["0", "1"];
    const step = Number(spec.step) > 0 ? Number(spec.step) : 1;
    const labelEvery = Number(spec.labelEvery) > 0 ? Number(spec.labelEvery) : 1;
    const count = Math.round((end - start) / step);
    const labels = [];
    for (let i = 0; i <= count; i += 1) {
      const value = start + i * step;
      const isLabelled = (i % labelEvery === 0) || i === count;
      const rounded = Math.round(value * 1000) / 1000;
      labels.push(isLabelled ? String(rounded) : "");
    }
    return labels;
  }

  /**
   * Draw a themed data table. rows[0] is the header when opts.header !== false.
   * Cells are strings or numbers. Column widths are equal unless opts.colWidths.
   */
  function addDataTable(slide, x, y, w, rows, opts) {
    const o = opts || {};
    const data = (Array.isArray(rows) ? rows : []).map((r) => (Array.isArray(r) ? r : [r]));
    if (!data.length) return { x, y, w, h: 0 };
    const cols = Math.max(...data.map((r) => r.length));
    const hasHeader = o.header !== false;
    const fontSize = o.fontSize || byBand(sz, 24, 20, 16);
    const rowH = o.rowH || Math.max(0.42, fontSize * 0.026 + 0.16);
    const headerColor = o.headerColor || C.PRIMARY;
    const colW = Array.isArray(o.colWidths) && o.colWidths.length === cols
      ? o.colWidths
      : Array.from({ length: cols }, () => w / cols);
    const totalH = rowH * data.length;
    const zebra = o.zebra !== false;

    data.forEach((row, r) => {
      let cx = x;
      const isHeader = hasHeader && r === 0;
      const fill = isHeader
        ? headerColor
        : (zebra && r % 2 === (hasHeader ? 0 : 1) ? (o.zebraFill || C.PRIMARY_SOFT || C.BG_LIGHT) : C.WHITE);
      for (let c = 0; c < cols; c += 1) {
        const cellW = colW[c];
        slide.addShape("rect", {
          x: cx, y: y + r * rowH, w: cellW, h: rowH,
          fill: { color: fill },
          line: { color: o.lineColor || C.PRIMARY_LINE || C.MUTED, width: 0.75 },
        });
        const value = row[c] == null ? "" : String(row[c]);
        if (value) {
          slide.addText(value, {
            x: cx + 0.06, y: y + r * rowH, w: cellW - 0.12, h: rowH,
            fontSize: isHeader ? Math.round(fontSize * 0.9) : fontSize,
            fontFace: isHeader ? FONT_B : (o.fontFace || FONT_B),
            color: isHeader ? C.WHITE : C.CHARCOAL,
            bold: isHeader || (o.boldFirstCol && c === 0),
            align: o.align || "center", valign: "middle", margin: 0,
            fit: "shrink", shrinkText: true,
          });
        }
        cx += cellW;
      }
    });
    return { x, y, w, h: totalH };
  }

  /**
   * Fit and draw a visual spec inside a frame. Returns the drawn bounds
   * ({x, y, w, h}) so a caller can place a label under it.
   */
  function drawVisual(slide, spec, frame, opts) {
    const o = opts || {};
    const f = normaliseFrame(frame);
    const s = spec && typeof spec === "object" ? spec : {};
    const type = String(s.type || "");

    switch (type) {
      case "tensFrame":
      case "fiveFrame": {
        const rows = type === "tensFrame" ? 2 : 1;
        const cell = Math.min(f.w / 5, f.h / rows, s.maxCell || byBand(sz, 1.3, 1.15, 1.0));
        const box = centreBox(f, cell * 5, cell * rows);
        const fn = type === "tensFrame" ? manips.addTensFrame : manips.addFiveFrame;
        fn(slide, box.x, box.y, box.w, s.filled || 0, {
          fillColor: s.fillColor || s.color, cellH: cell, fillFromTop: s.fillFromTop,
        });
        return box;
      }
      case "doubleTensFrame": {
        const gapRatio = 0.35;
        const cell = Math.min(f.w / 5, f.h / (4 + gapRatio), s.maxCell || byBand(sz, 1.0, 0.9, 0.8));
        const box = centreBox(f, cell * 5, cell * (4 + gapRatio));
        manips.addTensFrame(slide, box.x, box.y, box.w, s.filledTop || 0, { fillColor: s.colorTop || s.color, cellH: cell });
        manips.addTensFrame(slide, box.x, box.y + cell * (2 + gapRatio), box.w, s.filledBottom || 0, {
          fillColor: s.colorBottom || s.color2 || C.SECONDARY, cellH: cell,
        });
        return box;
      }
      case "dotCard": {
        const size = Math.min(f.w, f.h, s.maxSize || byBand(sz, 3.4, 3.0, 2.6));
        const box = centreBox(f, size, size);
        manips.addDotCard(slide, box.x, box.y, size, s.count || 1, { dotColor: s.color });
        return box;
      }
      case "dotCards": {
        const counts = Array.isArray(s.counts) ? s.counts : [s.count || 1];
        const gap = 0.3;
        const size = Math.min(f.h, (f.w - gap * (counts.length - 1)) / counts.length, s.maxSize || byBand(sz, 2.6, 2.3, 2.0));
        const box = centreBox(f, size * counts.length + gap * (counts.length - 1), size);
        counts.forEach((n, i) => manips.addDotCard(slide, box.x + i * (size + gap), box.y, size, n, { dotColor: s.color }));
        return box;
      }
      case "numberTrack": {
        const start = Number.isFinite(s.start) ? s.start : 1;
        const end = Number.isFinite(s.end) ? s.end : 10;
        const count = Math.max(1, end - start + 1);
        const cellW = Math.min(f.w / count, s.maxCell || byBand(sz, 1.0, 0.9, 0.8));
        const cellH = Math.min(f.h, cellW * 0.95);
        const box = centreBox(f, cellW * count, cellH);
        manips.addNumberTrack(slide, box.x, box.y, box.w, start, end, s.highlight || [], { cellH, fillColor: s.color });
        return box;
      }
      case "numberLine": {
        const labels = buildNumberLineLabels(s);
        const lineW = Math.min(f.w - 0.6, s.maxW || 8.4);
        const h = byBand(sz, 1.1, 1.0, 0.9);
        const box = centreBox(f, lineW, h);
        const baselineY = box.y + h * 0.42;
        manips.addNumberLine(slide, box.x, baselineY, lineW, labels, s.marked || [], {
          tickH: s.tickH || byBand(sz, 0.24, 0.2, 0.16),
          labelFontSize: s.labelFontSize,
          markColor: s.color,
        });
        return box;
      }
      case "fractionStrips": {
        const strips = Array.isArray(s.strips) ? s.strips : [{ denom: s.denom || 2, shaded: s.shaded || 0 }];
        const stripH = byBand(sz, 0.75, 0.66, 0.56);
        const gap = strips.length > 1 ? 0.14 : 0;
        const h = Math.min(f.h, strips.length * stripH + gap * (strips.length - 1));
        const w = Math.min(f.w, s.maxW || 8.6);
        const box = centreBox(f, w, h);
        manips.addFractionStripSet(slide, box.x, box.y, w, h, strips, {
          labelW: s.labelW != null ? s.labelW : (s.showLabels === false ? 0 : byBand(sz, 1.1, 1.0, 0.9)),
          showLabels: s.showLabels,
          gap,
        });
        return box;
      }
      case "array": {
        const rows = Math.max(1, s.rows || 1);
        const cols = Math.max(1, s.cols || 1);
        const gapRatio = 0.12;
        const cell = Math.min(
          f.w / (cols + (cols - 1) * gapRatio),
          f.h / (rows + (rows - 1) * gapRatio),
          s.maxCell || byBand(sz, 0.85, 0.72, 0.6),
        );
        const gap = cell * gapRatio;
        const box = centreBox(f, cols * cell + (cols - 1) * gap, rows * cell + (rows - 1) * gap);
        manips.addArray(slide, box.x, box.y, rows, cols, { cellSize: cell, gap, color: s.color, withFrame: s.withFrame });
        return box;
      }
      case "baseTen": {
        const hundreds = Math.max(0, s.hundreds || 0);
        const tens = Math.max(0, s.tens || 0);
        const ones = Math.max(0, s.ones || 0);
        // width in units: flats 10u + 0.05", rods u + 0.04", cubes u + 0.04", plus group gaps
        const solveUnit = (u) =>
          hundreds * (10 * u + 0.05) + (hundreds ? 0.13 : 0) +
          tens * (u + 0.04) + (tens ? 0.14 : 0) +
          ones * (u + 0.04);
        let unit = Math.min(f.h / 10, s.maxUnit || byBand(sz, 0.36, 0.32, 0.26));
        while (unit > 0.08 && solveUnit(unit) > f.w) unit -= 0.005;
        const w = solveUnit(unit);
        const box = centreBox(f, w, unit * 10);
        manips.addBaseTenBlocks(slide, box.x, box.y, hundreds, tens, ones, { unit, color: s.color });
        return box;
      }
      case "groupedCounters": {
        const groups = Math.max(1, s.groups || 1);
        const per = Math.max(1, s.per || 1);
        const framePad = 0.10;
        const groupGap = 0.28;
        const solveW = (dot) => groups * (per * dot + (per - 1) * dot * 0.35 + framePad * 2) + (groups - 1) * groupGap;
        let dot = Math.min(f.h - framePad * 2, s.maxDot || byBand(sz, 0.7, 0.6, 0.5));
        while (dot > 0.14 && solveW(dot) > f.w) dot -= 0.01;
        const w = solveW(dot);
        const box = centreBox(f, w, dot + framePad * 2);
        manips.addGroupedCounters(slide, box.x, box.y + framePad, groups, per, {
          dot, gap: dot * 0.35, groupGap, dotColor: s.color,
        });
        return box;
      }
      case "ppwMat": {
        const aspect = 1.75;
        let w = Math.min(f.w, s.maxW || byBand(sz, 6.5, 6.0, 5.5));
        let h = w / aspect;
        if (h > f.h) { h = f.h; w = h * aspect; }
        const box = centreBox(f, w, h);
        manips.addPartPartWholeMat(slide, box.x, box.y, w, h, {
          whole: s.whole, partA: s.partA, partB: s.partB,
        }, { valueFontSize: s.valueFontSize });
        return box;
      }
      case "chips": {
        const items = Array.isArray(s.items) ? s.items : [];
        const chipH = Math.min(f.h, s.chipH || byBand(sz, 1.0, 0.85, 0.7));
        const w = Math.min(f.w, s.maxW || 8.6);
        const box = centreBox(f, w, chipH);
        manips.addChipRow(slide, box.x, box.y, w, items, { chipH, fontSize: s.fontSize, fillColor: s.fillColor, borderColor: s.color, textColor: s.textColor });
        return box;
      }
      case "pictogram": {
        const labelFont = s.labelFontSize || byBand(sz, 22, 19, 15);
        const labelH = s.label ? Math.max(0.3, labelFont * 0.022 + 0.12) + 0.06 : 0;
        const size = Math.min(f.w, f.h - labelH, s.maxSize || byBand(sz, 3.0, 2.7, 2.4));
        const box = centreBox(f, size, size + labelH);
        picto.addPictogram(slide, s.name, box.x, box.y, size, {
          style: s.style, color: s.color, glyphColor: s.glyphColor,
          label: s.label, labelFontSize: s.labelFontSize, labelW: Math.max(size + 1.6, 2.4),
        });
        return box;
      }
      case "pictograms": {
        const items = Array.isArray(s.items) ? s.items : [];
        const labelFont = s.labelFontSize || byBand(sz, 22, 19, 15);
        // Same box the row draws for its labels, so the frame really contains them.
        const labelH = s.labels === false ? 0 : Math.max(0.3, labelFont * 0.022 + 0.12) + 0.06;
        const gap = s.gap != null ? s.gap : 0.35;
        const maxSize = Math.min(f.h - labelH, s.size || byBand(sz, 2.0, 1.8, 1.5));
        const size = Math.min(maxSize, (f.w - gap * (items.length - 1)) / Math.max(items.length, 1));
        const totalW = size * items.length + gap * (items.length - 1);
        const box = centreBox(f, totalW, size + labelH);
        picto.addPictogramRow(slide, box.x, box.y, totalW, items, {
          size, gap, style: s.style, color: s.color, colors: s.colors, labels: s.labels, labelFontSize: s.labelFontSize,
        });
        return box;
      }
      case "text": {
        const text = String(s.text == null ? "" : s.text);
        const fontSize = s.fontSize || byBand(sz, 96, 84, 66);
        const w = Math.min(f.w, s.maxW || 8.6);
        const h = Math.min(f.h, s.maxH || (fontSize * 0.03 + 0.6));
        const box = centreBox(f, w, h);
        if (s.card !== false) {
          slide.addShape("roundRect", {
            x: box.x, y: box.y, w, h, rectRadius: 0.12,
            fill: { color: s.fill || C.WHITE },
            line: { color: s.color || C.PRIMARY, width: 2 },
          });
        }
        slide.addText(text, {
          x: box.x + 0.1, y: box.y + 0.05, w: w - 0.2, h: h - 0.1,
          fontSize, fontFace: s.fontFace || FONT_H,
          color: s.textColor || C.CHARCOAL, bold: s.bold !== false,
          align: "center", valign: "middle", margin: 0,
          fit: "shrink", shrinkText: true,
        });
        return box;
      }
      case "image": {
        const source = s.path || s.data;
        if (!source) { console.warn("WARN [drawVisual] image spec needs a path"); return f; }
        let w = f.w;
        let h = f.h;
        try {
          const dims = getImageDimensions(source);
          const aspect = dims.width / dims.height;
          if (w / h > aspect) w = h * aspect; else h = w / aspect;
        } catch (err) {
          console.warn(`WARN [drawVisual] cannot read image ${source}: ${err.message}`);
        }
        const box = centreBox(f, w, h);
        if (s.frame !== false) {
          slide.addShape("roundRect", {
            x: box.x - 0.06, y: box.y - 0.06, w: w + 0.12, h: h + 0.12, rectRadius: 0.08,
            fill: { color: C.WHITE }, line: { color: C.MUTED, width: 0.6 },
          });
        }
        slide.addImage(s.path ? { path: s.path, x: box.x, y: box.y, w, h } : { data: s.data, x: box.x, y: box.y, w, h });
        return box;
      }
      case "table": {
        const rows = Array.isArray(s.rows) ? s.rows : [];
        const fontSize = s.fontSize || byBand(sz, 24, 20, 16);
        const rowH = Math.min(f.h / Math.max(rows.length, 1), s.rowH || Math.max(0.42, fontSize * 0.026 + 0.16));
        const w = Math.min(f.w, s.maxW || 8.6);
        const box = centreBox(f, w, rowH * rows.length);
        addDataTable(slide, box.x, box.y, w, rows, { ...s, rowH, fontSize });
        return box;
      }
      case "custom": {
        if (typeof s.draw === "function") {
          const result = s.draw(slide, f);
          return result && Number.isFinite(result.x) ? result : f;
        }
        console.warn("WARN [drawVisual] custom spec needs a draw(slide, frame) function");
        return f;
      }
      default:
        console.warn(`WARN [drawVisual] unknown visual type "${type}". Supported: ${SUPPORTED_TYPES.join(", ")}`);
        return f;
    }
  }

  function isVisualSpec(value) {
    return Boolean(value) && typeof value === "object" && !Array.isArray(value) &&
      typeof value.type === "string" && SUPPORTED_TYPES.includes(value.type);
  }

  return { drawVisual, addDataTable, isVisualSpec, VISUAL_TYPES: SUPPORTED_TYPES.slice() };
}

module.exports = { createVisualSpec, SUPPORTED_TYPES };
