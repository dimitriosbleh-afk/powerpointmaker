"use strict";

const { contrastRatio } = require("../core/contrast");
const { SLIDE_H, SAFE_BOTTOM, CONTENT_TOP } = require("../core/layout");
const {
  createStructuredMockupPlan,
  isStructuredMockupSpec,
  lightenHex,
} = require("../core/mockups");
const { normalizeLessonTargets } = require("../core/notes");
const { runSlideDiagnostics } = require("../core/diagnostics");
const { DEFAULT_SIZES, byBand } = require("../core/gradeBand");
const {
  prepareBullets,
  estimateBulletHeight,
  fitBulletFontSize,
  fitTextFontSize,
} = require("../core/bulletFit");

/**
 * Create the 7 universal slide builders bound to a specific palette and
 * grade band. Every subject gets these. They close over C, FONT_H, FONT_B,
 * el, shadowFn and the grade-aware sizing table S.
 *
 * @param {object}   C         - palette colours (semantic keys)
 * @param {string}   FONT_H    - heading font
 * @param {string}   FONT_B    - body font
 * @param {object}   el        - bound element helpers from createElements()
 * @param {Function} shadowFn  - zero-arg shadow factory
 * @param {object}   [S]       - grade-band sizing table; falls back to upper-primary defaults if omitted
 * @param {object}   [defaults] - subject-aware defaults supplied by the factory
 * @returns {object} { titleSlide, liSlide, contentSlide, cfuSlide, closingSlide, annotatedModelSlide, compareVisualSlide }
 */
const SC_TIER_LABELS = ["Everyone", "Most", "Stretch"];

function createBaseBuilders(C, FONT_H, FONT_B, el, shadowFn, S, defaults) {
  const sz = S || DEFAULT_SIZES;
  const builderDefaults = defaults || {};

  const DENSITY = byBand(sz,
    { roomyBulletCount: 3, roomyTotalLines: 5, narrowChars: 28, wideChars: 42 },
    { roomyBulletCount: 4, roomyTotalLines: 6, narrowChars: 32, wideChars: 50 },
    { roomyBulletCount: 4, roomyTotalLines: 7, narrowChars: 36, wideChars: 56 },
  );

  function estimateWrappedLines(text, charsPerLine) {
    const raw = String(text || "");
    const segments = raw.split("\n");
    return segments.reduce((count, segment) => {
      const trimmed = segment.trim();
      if (!trimmed) return count + 1;
      return count + Math.max(1, Math.ceil(trimmed.length / charsPerLine));
    }, 0);
  }

  function getBulletCardMetrics(items, opts) {
    const o = opts || {};
    const narrow = Boolean(o.narrow);
    const charsPerLine = narrow ? DENSITY.narrowChars : DENSITY.wideChars;
    const prepared = prepareBullets(items);
    const totalLines = prepared.reduce(
      (sum, item) => sum + estimateWrappedLines(item.text, charsPerLine),
      0
    );
    const bulletCount = Math.max(prepared.length, 1);
    const roomy = bulletCount <= DENSITY.roomyBulletCount && totalLines <= DENSITY.roomyTotalLines;

    const baseRoomy = narrow ? Math.max(sz.bodyDense, sz.body * 0.92) : sz.body;
    const baseDense = narrow ? Math.max(sz.bodyDense * 0.95, sz.body * 0.85) : sz.bodyDense;
    const baseTight = baseDense * sz._shrink;
    // Floor: never shrink below 14pt for F, 13pt for Y12, 11pt for Y36.
    // Below those, content really should be split across slides instead.
    const fontFloor = byBand(sz, 14, 13, 11);

    const idealFontSize = roomy
      ? baseRoomy
      : totalLines <= (DENSITY.roomyTotalLines + 3)
        ? baseDense
        : baseTight;

    const cardPadding = roomy ? 0.46 : 0.36;
    const topInset = roomy ? 0.20 : 0.15;
    // Available text-box height inside the card (the bullet text frame is
    // inset by topInset top and bottom).
    const maxCardH = SAFE_BOTTOM - CONTENT_TOP;
    const maxTextH = maxCardH - topInset * 2;
    // Pre-compute the largest fontSize at which prepared content fits.
    // PptxGenJS autofit on bullet lists is unreliable — this guarantees fit.
    const fontSize = fitBulletFontSize(prepared, maxTextH, charsPerLine, idealFontSize, fontFloor);
    const bodyH = estimateBulletHeight(prepared, fontSize, charsPerLine);
    const cardH = Math.min(maxCardH, bodyH + cardPadding);
    return {
      fontSize,
      cardH: Math.max(cardH, roomy ? 1.75 : 1.5),
      bodyH,
      topInset,
      prepared,
    };
  }

  function getQuestionCardMetrics(questionText) {
    const charsPerLine = byBand(sz, 36, 44, 54);
    const totalLines = estimateWrappedLines(questionText, charsPerLine);
    const heroSize = sz.heroQuestion;
    const stepDown = heroSize * sz._shrink;
    const tight = stepDown * sz._shrink;
    const idealFontSize = totalLines <= 3 ? heroSize : totalLines <= 6 ? stepDown : tight;
    // Hard cap on card height — the pill above takes 0.56" of vertical space.
    const maxCardH = SAFE_BOTTOM - (CONTENT_TOP + 0.56);
    // 0.36 padding (0.20 top + 0.16 bottom) inside the card for the text box.
    const maxTextH = maxCardH - 0.36;
    const fontFloor = byBand(sz, 22, 20, 18);
    // Pre-compute fontSize that guarantees questionText fits maxTextH —
    // PptxGenJS shrinkText is unreliable on multi-line CFU questions.
    const fontSize = fitTextFontSize(questionText, maxTextH, charsPerLine, idealFontSize, fontFloor);
    const lineHeight = Math.max(0.24, fontSize * 0.020 + 0.06);
    const cardH = Math.min(
      maxCardH,
      Math.max(1.45, totalLines * lineHeight + 0.50)
    );
    return { fontSize, cardH };
  }

  function pickOnDarkColor(preferred, ...fallbacks) {
    const candidates = [preferred, ...fallbacks].filter(Boolean);
    const passing = candidates.find((color) => contrastRatio(color, C.BG_DARK) >= 4.5);
    if (passing) return passing;
    return candidates.reduce((best, color) =>
      contrastRatio(color, C.BG_DARK) > contrastRatio(best, C.BG_DARK) ? color : best
    );
  }

  const subtitleOnDark = pickOnDarkColor(C.SUBTITLE, C.TEXT_ON_DARK, C.WHITE);
  const metaOnDark = pickOnDarkColor(C.MUTED, C.TEXT_ON_DARK, C.WHITE);
  const accentOnDark = pickOnDarkColor(C.ACCENT, C.TEXT_ON_DARK, C.WHITE);

  function addSceneRect(slide, x, y, w, h, fill, lineColor, radius) {
    slide.addShape("roundRect", {
      x,
      y,
      w,
      h,
      rectRadius: radius != null ? radius : 0.04,
      fill: { color: fill },
      line: { color: lineColor || fill, width: 0.4 },
    });
  }

  function addMockupText(slide, x, y, w, h, text, opts) {
    const o = opts || {};
    if (!text) return;
    slide.addText(String(text), {
      x,
      y,
      w,
      h,
      fontSize: o.fontSize || (sz.mockupText + 1),
      fontFace: o.fontFace || FONT_B,
      color: o.color || C.CHARCOAL,
      bold: Boolean(o.bold),
      italic: Boolean(o.italic),
      margin: 0,
      fit: "shrink",
      valign: o.valign || "middle",
      align: o.align || "left",
    });
  }

  function splitMockupSegments(text, maxCount) {
    const limit = Math.max(1, Number(maxCount) || 4);
    return String(text || "")
      .split("|")
      .map((segment) => segment.trim())
      .filter(Boolean)
      .slice(0, limit);
  }

  function drawLineSet(slide, x, y, w, h, component, spec) {
    const count = Math.max(1, Number(component.count) || 3);
    const lineGap = component.lineGap != null ? component.lineGap : 0.04;
    const linePad = component.linePad != null ? component.linePad : 0.08;
    const lineH = Math.max(0.03, (h - linePad * 2 - lineGap * (count - 1)) / count);
    const widths = Array.isArray(component.widths) ? component.widths : [];
    const lineColor = component.lineColor || spec.mutedLine || C.MUTED;
    for (let index = 0; index < count; index += 1) {
      const ratio = widths[index] || (index === count - 1 ? 0.62 : index % 2 === 0 ? 0.92 : 0.82);
      slide.addShape("roundRect", {
        x: x + 0.06,
        y: y + linePad + index * (lineH + lineGap),
        w: Math.max(0.18, (w - 0.12) * ratio),
        h: Math.max(0.02, lineH),
        rectRadius: 0.02,
        fill: { color: lineColor },
        line: { color: lineColor, width: 0.2 },
      });
    }
  }

  function drawPhotoPlaceholder(slide, x, y, w, h, component, spec) {
    const bg = component.fill || lightenHex(spec.accent, 0.9);
    const border = component.border || spec.softBorder || C.MUTED;
    addSceneRect(slide, x, y, w, h, bg, border, 0.05);
    const innerX = x + 0.06;
    const innerY = y + 0.06;
    const innerW = w - 0.12;
    const innerH = h - 0.12;
    slide.addShape("rect", {
      x: innerX, y: innerY, w: innerW, h: innerH,
      fill: { color: lightenHex(bg, 0.04) },
      line: { color: border, width: 0.25 },
    });
    slide.addShape("roundRect", {
      x: innerX + 0.06, y: innerY + 0.06, w: Math.max(0.24, innerW * 0.2), h: 0.08, rectRadius: 0.02,
      fill: { color: lightenHex(spec.accent, 0.25) },
      line: { color: lightenHex(spec.accent, 0.25), width: 0.2 },
    });
    slide.addShape("line", {
      x: innerX + 0.12, y: innerY + 0.14, w: innerW - 0.24, h: innerH - 0.28,
      line: { color: spec.mutedLine, width: 0.9 },
    });
    slide.addShape("roundRect", {
      x: innerX + 0.12, y: innerY + innerH - 0.24, w: Math.max(0.28, innerW - 0.24), h: 0.08, rectRadius: 0.02,
      fill: { color: spec.mutedLine, transparency: 20 },
      line: { color: spec.mutedLine, width: 0.2 },
    });
  }

  function drawNavChips(slide, x, y, w, h, component, spec) {
    const fill = component.fill || spec.softFill;
    addSceneRect(slide, x, y, w, h, fill, component.border || spec.softBorder, 0.03);
    const segments = splitMockupSegments(component.text, 4);
    if (!segments.length) {
      drawLineSet(slide, x, y, w, h, { count: 1, lineColor: spec.mutedLine, widths: [0.78] }, spec);
      return;
    }

    const chipGap = 0.04;
    const chipH = Math.max(0.08, h - 0.12);
    const availableW = Math.max(0.22, w - 0.16 - chipGap * Math.max(segments.length - 1, 0));
    const chipW = Math.max(0.24, availableW / segments.length);
    segments.forEach((segment, index) => {
      const chipX = x + 0.08 + index * (chipW + chipGap);
      slide.addShape("roundRect", {
        x: chipX,
        y: y + (h - chipH) / 2,
        w: chipW,
        h: chipH,
        rectRadius: 0.025,
        fill: { color: "FFFFFF", transparency: 8 },
        line: { color: spec.softBorder, width: 0.25 },
      });
      addMockupText(slide, chipX + 0.04, y + (h - chipH) / 2 + 0.01, chipW - 0.08, chipH - 0.02, segment, {
        fontSize: component.fontSize || 7.6,
        color: component.textColor || spec.textColor || C.CHARCOAL,
        bold: true,
        align: "center",
      });
    });
  }

  function drawChartPlaceholder(slide, x, y, w, h, component, spec) {
    const bg = component.fill || "FFFFFF";
    addSceneRect(slide, x, y, w, h, bg, component.border || spec.softBorder, 0.05);
    slide.addShape("line", {
      x: x + 0.08, y: y + h - 0.08, w: w - 0.16, h: 0,
      line: { color: spec.mutedLine, width: 0.8 },
    });
    slide.addShape("line", {
      x: x + 0.08, y: y + 0.08, w: 0, h: h - 0.16,
      line: { color: spec.mutedLine, width: 0.8 },
    });
    [0.28, 0.48, 0.68].forEach((pos, index) => {
      const height = h * (0.22 + index * 0.12);
      slide.addShape("roundRect", {
        x: x + w * pos, y: y + h - 0.08 - height, w: w * 0.1, h: height, rectRadius: 0.02,
        fill: { color: index === 1 ? spec.accent : lightenHex(spec.accent, 0.38 + index * 0.12) },
        line: { color: index === 1 ? spec.accent : lightenHex(spec.accent, 0.38 + index * 0.12), width: 0.2 },
      });
    });
  }

  function drawDiagramPlaceholder(slide, x, y, w, h, component, spec) {
    const bg = component.fill || lightenHex(spec.accent, 0.92);
    addSceneRect(slide, x, y, w, h, bg, component.border || spec.softBorder, 0.05);
    const points = [
      { cx: 0.24, cy: 0.58 },
      { cx: 0.5, cy: 0.34 },
      { cx: 0.76, cy: 0.58 },
    ];
    slide.addShape("line", {
      x: x + w * points[0].cx, y: y + h * points[0].cy, w: w * (points[1].cx - points[0].cx), h: h * (points[1].cy - points[0].cy),
      line: { color: spec.mutedLine, width: 0.9 },
    });
    slide.addShape("line", {
      x: x + w * points[1].cx, y: y + h * points[1].cy, w: w * (points[2].cx - points[1].cx), h: h * (points[2].cy - points[1].cy),
      line: { color: spec.mutedLine, width: 0.9 },
    });
    points.forEach((point, index) => {
      const fill = index === 1 ? spec.accent : lightenHex(spec.accent, 0.45);
      slide.addShape("roundRect", {
        x: x + w * point.cx - 0.08, y: y + h * point.cy - 0.08, w: 0.16, h: 0.16,
        rectRadius: 0.08,
        fill: { color: fill },
        line: { color: fill, width: 0.2 },
      });
    });
  }

  function drawBrowserFramePlaceholder(slide, x, y, w, h, component, spec) {
    addSceneRect(slide, x, y, w, h, component.fill || "FFFFFF", component.border || spec.softBorder, 0.05);
    slide.addShape("roundRect", {
      x: x + 0.05, y: y + 0.05, w: w - 0.1, h: 0.16, rectRadius: 0.03,
      fill: { color: lightenHex(spec.accent, 0.92) },
      line: { color: spec.softBorder, width: 0.2 },
    });
    [0.1, 0.17, 0.24].forEach((cx) => {
      slide.addShape("roundRect", {
        x: x + w * cx, y: y + 0.09, w: 0.05, h: 0.05,
        rectRadius: 0.025,
        fill: { color: lightenHex(spec.accent, 0.5) },
        line: { color: lightenHex(spec.accent, 0.5), width: 0.2 },
      });
    });
    addSceneRect(slide, x + 0.08, y + 0.28, w * 0.28, h - 0.38, lightenHex(spec.accent, 0.94), spec.softBorder, 0.03);
    addSceneRect(slide, x + w * 0.4, y + 0.28, w * 0.5, h - 0.38, "FFFFFF", spec.softBorder, 0.03);
    drawLineSet(slide, x + w * 0.43, y + 0.34, w * 0.44, h - 0.5, { count: 4, lineColor: spec.mutedLine }, spec);
  }

  function drawCardGridPlaceholder(slide, x, y, w, h, component, spec) {
    addSceneRect(slide, x, y, w, h, component.fill || "FFFFFF", component.border || spec.softBorder, 0.05);
    const rows = Math.max(1, Number(component.rows) || 2);
    const cols = Math.max(1, Number(component.cols) || 2);
    const gap = 0.04;
    const cardW = (w - gap * (cols + 1)) / cols;
    const cardH = (h - gap * (rows + 1)) / rows;
    for (let row = 0; row < rows; row += 1) {
      for (let col = 0; col < cols; col += 1) {
        const cardX = x + gap + col * (cardW + gap);
        const cardY = y + gap + row * (cardH + gap);
        addSceneRect(slide, cardX, cardY, cardW, cardH, lightenHex(spec.accent, 0.95), spec.softBorder, 0.03);
        slide.addShape("roundRect", {
          x: cardX + 0.05, y: cardY + 0.05, w: cardW - 0.1, h: 0.08, rectRadius: 0.02,
          fill: { color: lightenHex(spec.accent, 0.25) },
          line: { color: lightenHex(spec.accent, 0.25), width: 0.2 },
        });
        drawLineSet(slide, cardX, cardY + 0.17, cardW, cardH - 0.22, { count: 2, lineColor: spec.mutedLine }, spec);
      }
    }
  }

  function drawSidebarRegion(slide, x, y, w, h, component, spec) {
    addSceneRect(slide, x, y, w, h, component.fill || "FFFFFF", component.border || spec.softBorder, 0.03);
    const leftW = w * 0.66;
    const railX = x + leftW + 0.05;
    const railW = Math.max(0.18, w - leftW - 0.1);
    drawLineSet(slide, x + 0.03, y + 0.02, leftW - 0.04, h - 0.04, {
      count: Math.max(3, Number(component.count) || 4),
      lineColor: component.lineColor || spec.mutedLine,
      widths: component.widths,
    }, spec);
    addSceneRect(slide, railX, y + 0.03, railW, h - 0.06, lightenHex(spec.accent, 0.94), spec.softBorder, 0.02);
    slide.addShape("roundRect", {
      x: railX + 0.03,
      y: y + 0.08,
      w: railW - 0.06,
      h: 0.08,
      rectRadius: 0.02,
      fill: { color: lightenHex(spec.accent, 0.22) },
      line: { color: lightenHex(spec.accent, 0.22), width: 0.2 },
    });
    [0.22, 0.38, 0.54].forEach((offset) => {
      slide.addShape("roundRect", {
        x: railX + 0.04,
        y: y + h * offset,
        w: railW - 0.08,
        h: 0.06,
        rectRadius: 0.02,
        fill: { color: spec.mutedLine },
        line: { color: spec.mutedLine, width: 0.2 },
      });
    });
  }

  function drawFeatureKeyCard(slide, x, y, w, h, label, detail, color) {
    slide.addShape("roundRect", {
      x,
      y,
      w,
      h,
      rectRadius: 0.05,
      fill: { color: C.WHITE },
      line: { color: C.MUTED, width: 0.45 },
    });
    slide.addShape("roundRect", {
      x: x + 0.08,
      y: y + 0.08,
      w: Math.min(1.12, Math.max(0.92, w * 0.42)),
      h: 0.24,
      rectRadius: 0.04,
      fill: { color },
      line: { color, width: 0.2 },
    });
    slide.addText(String(label || ""), {
      x: x + 0.12,
      y: y + 0.1,
      w: Math.min(1.04, Math.max(0.84, w * 0.38)),
      h: 0.22,
      fontSize: sz.featureLabel,
      fontFace: FONT_B,
      color: C.WHITE,
      bold: true,
      align: "center",
      valign: "middle",
      margin: 0,
      fit: "shrink",
    });
    slide.addText(String(detail || ""), {
      x: x + 0.1,
      y: y + 0.38,
      w: w - 0.2,
      h: h - 0.46,
      fontSize: sz.featureDetail,
      fontFace: FONT_B,
      color: C.CHARCOAL,
      margin: 0,
      fit: "shrink",
      valign: "top",
    });
  }

  function drawHeroMockup(slide, x, y, w, h, component, spec) {
    const mode = component.mode || "diagram";
    if (mode === "chart") {
      drawChartPlaceholder(slide, x, y, w, h, component, spec);
    } else if (mode === "browserFrame") {
      drawBrowserFramePlaceholder(slide, x, y, w, h, component, spec);
    } else if (mode === "cardGrid") {
      drawCardGridPlaceholder(slide, x, y, w, h, component, spec);
    } else if (mode === "photo") {
      drawPhotoPlaceholder(slide, x, y, w, h, component, spec);
    } else {
      drawDiagramPlaceholder(slide, x, y, w, h, component, spec);
    }

    if (component.overlayText) {
      const overlayFill = component.overlayFill || spec.accent;
      slide.addShape("roundRect", {
        x: x + 0.08, y: y + 0.08, w: w - 0.16, h: Math.min(0.18, h * 0.2), rectRadius: 0.04,
        fill: { color: overlayFill },
        line: { color: overlayFill, width: 0.2 },
      });
      addMockupText(slide, x + 0.12, y + 0.09, w - 0.24, Math.min(0.14, h * 0.16), component.overlayText, {
        fontSize: component.overlayFontSize || 9.5,
        fontFace: FONT_H,
        color: component.textColor || C.WHITE,
        bold: true,
        align: component.align || "left",
      });
    }
  }

  function drawStructuredMockup(slide, x, y, w, h, spec, opts) {
    const o = opts || {};
    const plan = createStructuredMockupPlan(spec, { x, y, w, h }, {
      innerPad: o.innerPad != null ? o.innerPad : 0.08,
      gap: o.gap != null ? o.gap : 0.04,
      minBlockH: 0.12,
    });
    const normalized = plan.normalized;
    slide.addShape("roundRect", {
      x,
      y,
      w,
      h,
      rectRadius: 0.05,
      fill: { color: normalized.pageFill || o.fill || C.BG_CARD },
      line: { color: normalized.pageBorder || o.border || C.MUTED, width: 0.6 },
    });

    plan.blocks.forEach(({ component, x: innerX, y: cursorY, w: innerW, h: blockH }) => {
      const baseTextColor = component.textColor || normalized.textColor || C.CHARCOAL;

      if (component.kind === "masthead") {
        const fill = component.fill || normalized.accent;
        addSceneRect(slide, innerX, cursorY, innerW, blockH, fill, component.border || fill, 0.04);
        addMockupText(slide, innerX + 0.08, cursorY + 0.02, innerW - 0.16, blockH - 0.04, component.text, {
          fontSize: component.fontSize || (blockH > 0.22 ? 12 : 10.2),
          fontFace: FONT_H,
          color: component.textColor || C.WHITE,
          bold: true,
          align: component.align || "center",
        });
      } else if (component.kind === "nav") {
        drawNavChips(slide, innerX, cursorY, innerW, blockH, component, normalized);
      } else if (component.kind === "heading" || component.kind === "subheading") {
        const fill = component.fill || "FFFFFF";
        addSceneRect(slide, innerX, cursorY, innerW, blockH, fill, component.border || normalized.softBorder, 0.03);
        addMockupText(slide, innerX + 0.08, cursorY + 0.02, innerW - 0.16, blockH - 0.04, component.text, {
          fontSize: component.fontSize || (component.kind === "heading" ? 11.3 : 9.6),
          fontFace: component.kind === "heading" ? FONT_H : FONT_B,
          color: component.textColor || (component.kind === "heading" ? normalized.accent : baseTextColor),
          bold: component.bold != null ? Boolean(component.bold) : true,
        });
      } else if (component.kind === "hero") {
        drawHeroMockup(slide, innerX, cursorY, innerW, blockH, component, normalized);
      } else if (component.kind === "chart") {
        drawChartPlaceholder(slide, innerX, cursorY, innerW, blockH, component, normalized);
      } else if (component.kind === "stat") {
        const fill = component.fill || lightenHex(normalized.accent, 0.1);
        addSceneRect(slide, innerX, cursorY, innerW, blockH, fill, component.border || fill, 0.04);
        addMockupText(slide, innerX + 0.08, cursorY + 0.02, innerW - 0.16, blockH - 0.04, component.text, {
          fontSize: component.fontSize || 11.4,
          fontFace: FONT_H,
          color: component.textColor || C.WHITE,
          bold: true,
          align: component.align || "center",
        });
      } else if (component.kind === "cta" || component.kind === "footerBand") {
        const outerFill = component.kind === "cta" ? "FFFFFF" : (component.fill || normalized.softFill);
        addSceneRect(slide, innerX, cursorY, innerW, blockH, outerFill, component.border || normalized.softBorder, 0.03);
        const buttonW = component.kind === "cta" ? innerW * 0.64 : innerW;
        const buttonX = component.kind === "cta" ? innerX + (innerW - buttonW) / 2 : innerX;
        const buttonFill = component.fill || normalized.accent;
        addSceneRect(slide, buttonX, cursorY + 0.04, buttonW, Math.max(0.08, blockH - 0.08), buttonFill, component.border || buttonFill, 0.03);
        addMockupText(slide, buttonX + 0.06, cursorY + 0.05, buttonW - 0.12, Math.max(0.04, blockH - 0.1), component.text, {
          fontSize: component.fontSize || 9.8,
          color: component.textColor || C.WHITE,
          bold: true,
          align: component.align || "center",
        });
      } else if (component.kind === "textBlock") {
        const fill = component.fill || "FFFFFF";
        addSceneRect(slide, innerX, cursorY, innerW, blockH, fill, component.border || normalized.softBorder, 0.03);
        drawLineSet(slide, innerX, cursorY, innerW, blockH, component, normalized);
      } else if (component.kind === "caption") {
        const fill = component.fill || normalized.softFill;
        addSceneRect(slide, innerX, cursorY, innerW, blockH, fill, component.border || normalized.softBorder, 0.03);
        addMockupText(slide, innerX + 0.08, cursorY + 0.02, innerW - 0.16, blockH - 0.04, component.text, {
          fontSize: component.fontSize || 8.8,
          color: component.textColor || baseTextColor,
          italic: true,
        });
      } else if (component.kind === "quote") {
        addSceneRect(slide, innerX, cursorY, innerW, blockH, component.fill || "FFFFFF", component.border || normalized.softBorder, 0.03);
        slide.addShape("rect", {
          x: innerX + 0.04, y: cursorY + 0.05, w: 0.05, h: blockH - 0.1,
          fill: { color: component.accent || normalized.accent },
          line: { color: component.accent || normalized.accent, width: 0.2 },
        });
        addMockupText(slide, innerX + 0.14, cursorY + 0.03, innerW - 0.2, blockH - 0.06, component.text, {
          fontSize: component.fontSize || 9.6,
          color: component.textColor || baseTextColor,
          italic: true,
        });
      } else if (component.kind === "iconRow") {
        addSceneRect(slide, innerX, cursorY, innerW, blockH, component.fill || "FFFFFF", component.border || normalized.softBorder, 0.03);
        const count = Math.max(3, Number(component.count) || 4);
        const gapW = innerW / (count + 1);
        for (let index = 0; index < count; index += 1) {
          slide.addShape("roundRect", {
            x: innerX + gapW * (index + 0.65), y: cursorY + blockH * 0.18, w: 0.14, h: 0.14,
            rectRadius: 0.07,
            fill: { color: index === 0 ? normalized.accent : lightenHex(normalized.accent, 0.45) },
            line: { color: index === 0 ? normalized.accent : lightenHex(normalized.accent, 0.45), width: 0.2 },
          });
          slide.addShape("roundRect", {
            x: innerX + gapW * (index + 0.45), y: cursorY + blockH * 0.56, w: 0.26, h: 0.05, rectRadius: 0.02,
            fill: { color: normalized.mutedLine },
            line: { color: normalized.mutedLine, width: 0.2 },
          });
        }
      } else if (component.kind === "sidebar") {
        drawSidebarRegion(slide, innerX, cursorY, innerW, blockH, component, normalized);
      } else {
        const fill = component.fill || normalized.softFill;
        addSceneRect(slide, innerX, cursorY, innerW, blockH, fill, component.border || normalized.softBorder, 0.03);
        addMockupText(slide, innerX + 0.08, cursorY + 0.02, innerW - 0.16, blockH - 0.04, component.text, {
          fontSize: component.fontSize || 10,
          color: component.textColor || baseTextColor,
          bold: true,
        });
      }

    });
  }

  function drawMockupPreview(slide, x, y, w, h, blocks, opts) {
    const o = opts || {};
    slide.addShape("roundRect", {
      x,
      y,
      w,
      h,
      rectRadius: 0.08,
      fill: { color: o.fill || C.BG_CARD },
      line: { color: o.border || C.MUTED, width: 0.8 },
    });

    if (isStructuredMockupSpec(blocks)) {
      drawStructuredMockup(slide, x + 0.06, y + 0.06, w - 0.12, h - 0.12, blocks, {
        fill: o.fill || C.BG_CARD,
        border: o.border || C.MUTED,
        accent: o.accent || C.PRIMARY,
        innerPad: 0.04,
        gap: 0.04,
      });
      return;
    }

    const previewBlocks = Array.isArray(blocks) && blocks.length > 0
      ? blocks.slice(0, 5)
      : ["Headline", "Lead / image / key detail", "Main detail", "Caption / key takeaway"];
    const innerPad = o.innerPad != null ? o.innerPad : 0.1;
    const gap = o.gap != null ? o.gap : 0.05;
    const contentH = h - innerPad * 2 - gap * Math.max(previewBlocks.length - 1, 0);
    const baseBlockH = Math.max(0.12, contentH / Math.max(previewBlocks.length, 1));
    let cursorY = y + innerPad;

    previewBlocks.forEach((block, index) => {
      const blockObj = typeof block === "string" ? { text: block } : (block || {});
      const rawScale = Number(blockObj.scale);
      const scale = Number.isFinite(rawScale) && rawScale > 0 ? rawScale : 1;
      const blockH = Math.max(0.1, baseBlockH * scale);
      const fillColor = blockObj.fill || (index === 0 ? (o.accent || C.PRIMARY) : (index % 2 === 0 ? C.BG_LIGHT : C.WHITE));
      const textColor = blockObj.textColor || (index === 0 ? C.WHITE : C.CHARCOAL);

      slide.addShape("roundRect", {
        x: x + innerPad,
        y: cursorY,
        w: w - innerPad * 2,
        h: blockH,
        rectRadius: 0.05,
        fill: { color: fillColor },
        line: { color: blockObj.lineColor || (index === 0 ? (o.accent || C.PRIMARY) : C.MUTED), width: 0.5 },
      });
      slide.addText(String(blockObj.text || ""), {
        x: x + innerPad + 0.08,
        y: cursorY + 0.02,
        w: w - innerPad * 2 - 0.16,
        h: blockH - 0.04,
        fontSize: blockObj.fontSize || (index === 0 ? 12.5 : 10.5),
        fontFace: blockObj.fontFace || (index === 0 ? FONT_H : FONT_B),
        color: textColor,
        bold: blockObj.bold != null ? Boolean(blockObj.bold) : index === 0,
        margin: 0,
        fit: "shrink",
        valign: "middle",
        align: blockObj.align || "left",
      });
      cursorY += blockH + gap;
    });
  }

  /**
   * addRevealAnswerBar - Large answer overlay sized for back-of-room
   * visibility. Used inside withReveal()'s revealFn for Daily Review,
   * Fluency, We Do, and CFU answer reveals (megaprompt §22 / §23).
   *
   * Renders a coloured bar at the supplied y-position with the answer
   * text large and bold. Optionally adds a small "Tick & fix" cue so the
   * routine is visible to students.
   *
   * @param {object}        slide    PptxGenJS slide object
   * @param {string|string[]} answers  One answer or array of answers (joined)
   * @param {object}        [opts]   { x, y, w, h, color, textColor, label, showTickAndFix, fontSize }
   */
  function addRevealAnswerBar(slide, answers, opts) {
    const o = opts || {};
    const list = Array.isArray(answers) ? answers : [answers];
    const cleaned = list
      .map((entry) => String(entry == null ? "" : entry).trim())
      .filter(Boolean);
    const text = cleaned.join("    ");

    const x = o.x != null ? o.x : 0.5;
    const w = o.w != null ? o.w : 9;
    const h = o.h != null ? o.h : byBand(sz, 1.05, 0.95, 0.8);
    const y = o.y != null ? o.y : (SAFE_BOTTOM - h);
    const fillColor = o.color || C.SUCCESS || C.ACCENT;
    const textColor = o.textColor || C.WHITE;
    const label = o.label != null ? o.label : "Answer";
    const showTick = o.showTickAndFix !== false;
    const fontSize = o.fontSize || byBand(sz, 36, 30, 24);

    slide.addShape("roundRect", {
      x, y, w, h, rectRadius: 0.1,
      fill: { color: fillColor },
    });

    if (label) {
      const tagW = byBand(sz, 1.6, 1.4, 1.2);
      const tagH = byBand(sz, 0.42, 0.38, 0.32);
      el.addTextOnShape(slide, String(label), {
        x: x + 0.18, y: y + 0.16, w: tagW, h: tagH, rectRadius: 0.06,
        fill: { color: C.WHITE },
      }, {
        fontSize: byBand(sz, 16, 14, 12),
        fontFace: FONT_B, color: fillColor,
        bold: true, align: "center", valign: "middle", margin: 0,
      });
    }

    const labelOffset = label ? byBand(sz, 1.85, 1.65, 1.45) : 0.2;
    slide.addText(String(text), {
      x: x + labelOffset, y: y + 0.12,
      w: w - labelOffset - 0.18,
      h: h - 0.24,
      fontSize, fontFace: FONT_H,
      color: textColor, bold: true,
      align: "left", valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    if (showTick) {
      const cueY = y + h - byBand(sz, 0.32, 0.28, 0.24);
      const cueW = byBand(sz, 1.8, 1.6, 1.4);
      slide.addText("Tick & fix", {
        x: x + w - cueW - 0.18, y: cueY,
        w: cueW, h: byBand(sz, 0.24, 0.22, 0.20),
        fontSize: byBand(sz, 12, 11, 10),
        fontFace: FONT_B, color: textColor, italic: true,
        align: "right", valign: "middle", margin: 0,
      });
    }
  }

  /**
   * titleSlide - Dark full-bleed title for lesson start.
   */
  function titleSlide(pres, title, subtitle, meta, notes) {
    const s = pres.addSlide();
    s.background = { color: C.BG_DARK };

    // Vertical accent bar
    s.addShape("rect", { x: 0, y: 0, w: 0.12, h: SLIDE_H, fill: { color: C.ACCENT } });

    // Decorative shapes (large, semi-transparent)
    s.addShape("roundRect", {
      x: 7.5, y: -0.6, w: 3.5, h: 3.5, rectRadius: 1.75,
      fill: { color: C.DECOR_1, transparency: 75 },
    });
    s.addShape("roundRect", {
      x: 8.2, y: 3.8, w: 2.5, h: 2.5, rectRadius: 1.25,
      fill: { color: C.DECOR_2, transparency: 80 },
    });

    const heroH = byBand(sz, 1.6, 1.45, 1.3);
    s.addText(title, {
      x: 0.7, y: 0.9, w: 8.0, h: heroH,
      fontSize: sz.titleHero, fontFace: FONT_H, color: C.WHITE, bold: true, margin: 0,
      fit: "shrink", shrinkText: true,
    });

    // Subtitle
    if (subtitle) {
      const subY = 0.9 + heroH + 0.05;
      s.addText(subtitle, {
        x: 0.7, y: subY, w: 8.0, h: 0.8,
        fontSize: sz.subtitleHero, fontFace: FONT_B, color: subtitleOnDark, margin: 0,
        fit: "shrink", shrinkText: true,
      });
    }

    // Meta line
    if (meta) {
      const metaY = subtitle ? (0.9 + heroH + 0.05 + 0.85) : (0.9 + heroH + 0.1);
      s.addText(meta, {
        x: 0.7, y: metaY, w: 8.0, h: 0.42,
        fontSize: sz.metaHero, fontFace: FONT_B, color: metaOnDark, margin: 0,
      });
    }

    if (notes) s.addNotes(notes);
    return s;
  }

  /**
   * liSlide - Learning Intention + Success Criteria.
   *
   * Default: plain unlabelled "I can..." bullet list (megaprompt §0a item 17 /
   * §14 / §52 — tier labels must NEVER appear on student-facing slides).
   * The internal SC1/SC2/SC3 progression (foundation / core / depth) still
   * drives lesson design but never reaches the slide face.
   *
   * Opt-in tiered preview for non-student-facing review: `opts.tieredSc: true`.
   * `opts.scLabels` and `opts.scColors` are honoured only on that opt-in path.
   *
   * @param {object}            pres
   * @param {string|string[]}   liItems
   * @param {string[]}          scItems
   * @param {string}            notes
   * @param {string}            footer
   * @param {object}            [opts]   { tieredSc, scLabels, scColors }
   */
  function liSlide(pres, liItems, scItems, notes, footer, opts) {
    const s = pres.addSlide();
    const o = opts || {};
    el.addTopBar(s, C.PRIMARY);
    el.addBadge(s, "Learning Intention");
    el.addTitle(s, "Learning Intention & Success Criteria");

    const normalizedTargets = normalizeLessonTargets(liItems, scItems);
    if (normalizedTargets.warnings.length) {
      console.warn(`[liSlide] ${normalizedTargets.warnings.join("; ")}. Extra items will be truncated; only the first LI and first three SC items render on the slide.`);
    }
    liItems = normalizedTargets.liItems;
    scItems = normalizedTargets.scItems;

    // Megaprompt §0a item 17 / §14 / §52: tier labels (Everyone/Most/Stretch)
    // are an internal design tool only and must NEVER appear on student-facing
    // slides. Default to a plain unlabelled "I can..." list. Callers can opt
    // back in with `opts.tieredSc: true` for non-student-facing previews.
    const tiered = o.tieredSc === true;
    const scLabels = Array.isArray(o.scLabels) && o.scLabels.length === 3
      ? o.scLabels
      : SC_TIER_LABELS;
    const scColors = Array.isArray(o.scColors) && o.scColors.length === 3
      ? o.scColors
      : [C.SUCCESS || C.SECONDARY, C.PRIMARY, C.ACCENT];

    const GAP      = 0.16;
    const LI_HDR_H = byBand(sz, 0.54, 0.50, 0.44);
    const SC_HDR_H = byBand(sz, 0.50, 0.46, 0.40);
    const PAD      = 0.14;
    const totalItems = liItems.length + scItems.length;
    const available  = SAFE_BOTTOM - CONTENT_TOP - GAP - LI_HDR_H - SC_HDR_H - PAD * 2;
    const perItemMax = byBand(sz, 0.62, 0.54, 0.42);
    const perItem    = Math.min(perItemMax, available / Math.max(totalItems, 1));
    const dense      = totalItems > 8;
    const fontSize   = dense ? Math.max(sz.liBody * sz._shrink, 10) : sz.liBody;

    // LI card
    const liBodyH = Math.max(liItems.length * perItem, 0.52);
    const liH     = LI_HDR_H + liBodyH + PAD;
    el.addCard(s, 0.5, CONTENT_TOP, 9, liH, { strip: C.PRIMARY });
    s.addText("Learning Intention", {
      x: 0.75, y: CONTENT_TOP + 0.08, w: 5, h: LI_HDR_H - 0.10,
      fontSize: sz.liHeader, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText(liItems[0] || "", {
      x: 0.75, y: CONTENT_TOP + LI_HDR_H, w: 8.5, h: liBodyH,
      fontFace: FONT_B, fontSize, color: C.CHARCOAL, margin: 0, valign: "middle",
      fit: "shrink", shrinkText: true,
    });

    // SC card
    const scY     = CONTENT_TOP + liH + GAP;
    const scBodyH = scItems.length * perItem;
    const scH     = SC_HDR_H + scBodyH + PAD;
    el.addCard(s, 0.5, scY, 9, scH, { strip: C.ACCENT });
    s.addText("Success Criteria — I can…", {
      x: 0.75, y: scY + 0.08, w: 5, h: SC_HDR_H - 0.10,
      fontSize: sz.liHeader, fontFace: FONT_B, color: C.CHARCOAL, bold: true, margin: 0,
    });

    if (tiered) {
      const rowGap = byBand(sz, 0.10, 0.08, 0.06);
      const rowH = Math.max(0.34, (scBodyH - rowGap * Math.max(scItems.length - 1, 0)) / Math.max(scItems.length, 1));
      const chipW = byBand(sz, 1.45, 1.30, 1.10);
      const chipH = Math.min(rowH - 0.04, byBand(sz, 0.48, 0.42, 0.32));
      const chipFontSize = byBand(sz, 14, 12, 10);
      scItems.forEach((text, i) => {
        const rowY = scY + SC_HDR_H + i * (rowH + rowGap);
        const chipColor = scColors[i] || C.PRIMARY;
        const chipLabel = scLabels[i] || "";
        // Chip
        slideAddChipRow(s, {
          x: 0.75, y: rowY,
          chipW, chipH, chipColor, chipLabel,
          chipFontSize,
          textX: 0.75 + chipW + 0.18,
          textY: rowY,
          textW: 8.5 - chipW - 0.18,
          textH: rowH,
          text,
          fontSize,
        });
      });
    } else {
      s.addText(scItems.map((t, i) => ({
        text: t,
        options: { bullet: true, breakLine: i < scItems.length - 1, fontSize, color: C.CHARCOAL },
      })), {
        x: 0.75, y: scY + SC_HDR_H, w: 8.5, h: scBodyH,
        fontFace: FONT_B, margin: 0,
        fit: "shrink", shrinkText: true,
      });
    }

    if (footer) el.addFooter(s, footer);
    if (notes) s.addNotes(notes);
    return s;
  }

  function slideAddChipRow(slide, p) {
    const chipY = p.y + (p.textH - p.chipH) / 2;
    el.addTextOnShape(slide, String(p.chipLabel || ""), {
      x: p.x, y: chipY, w: p.chipW, h: p.chipH, rectRadius: 0.06,
      fill: { color: p.chipColor },
    }, {
      fontSize: p.chipFontSize, fontFace: FONT_B, color: C.WHITE,
      bold: true, align: "center", valign: "middle", margin: 0,
    });
    slide.addText(String(p.text || ""), {
      x: p.textX, y: p.textY,
      w: p.textW, h: p.textH,
      fontSize: p.fontSize, fontFace: FONT_B, color: C.CHARCOAL,
      valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });
  }

  /**
   * contentSlide - Standard content slide with badge, title, and bullet card.
   * Optional drawRight callback for right-column visual content.
   */
  function contentSlide(pres, badgeText, badgeColor, title, bullets, notes, footer, drawRight) {
    const s = pres.addSlide();
    el.addTopBar(s, C.PRIMARY);
    el.addBadge(s, badgeText || "Content", { color: badgeColor || C.PRIMARY });
    el.addTitle(s, title);

    const cardW = drawRight ? 4.5 : 9.0;
    const metrics = bullets && bullets.length ? getBulletCardMetrics(bullets, { narrow: Boolean(drawRight) }) : null;
    const cardH = metrics ? metrics.cardH : (drawRight ? 2.0 : 1.55);
    const contentY = CONTENT_TOP;
    const layoutGuide = {
      titleY: 0.65,
      titleH: 0.62,
      panelTop: contentY,
      panelTopPadded: contentY + 0.08,
      leftCardX: 0.5,
      leftCardY: contentY,
      leftCardW: cardW,
      leftCardH: cardH,
      rightX: 5.2,
      rightW: 4.3,
      safeBottom: SAFE_BOTTOM,
    };

    el.addCard(s, 0.5, contentY, cardW, cardH, {
      strip: badgeColor || C.PRIMARY,
      fill: C.WHITE,
    });

    if (bullets && bullets.length) {
      const prepared = metrics.prepared || prepareBullets(bullets);
      const baseSpacePt = metrics.fontSize >= 16 ? 5 : 3;
      s.addText(prepared.map((item, i) => ({
        text: item.text,
        options: {
          bullet: true,
          breakLine: i < prepared.length - 1,
          fontSize: metrics.fontSize,
          color: C.CHARCOAL,
          // Per-paragraph space-after grows when the build script used an
          // empty-string spacer between bullets. Avoids ugly empty bullet
          // markers while preserving visual grouping.
          paraSpaceAfter: baseSpacePt + (item.extraSpaceAfter || 0) * metrics.fontSize * 0.9,
        },
      })), {
        x: 0.75, y: contentY + metrics.topInset, w: cardW - 0.5, h: cardH - metrics.topInset * 2,
        fontFace: FONT_B, valign: "top", margin: 0,
      });
    }

    if (drawRight) drawRight(s, layoutGuide);
    if (footer) el.addFooter(s, footer);
    if (notes) s.addNotes(notes);
    if (drawRight) runSlideDiagnostics(s, pres);
    return s;
  }

  /**
   * cfuSlide - Check for Understanding slide.
   *
   * Two call shapes are supported:
   *
   *   Legacy positional: cfuSlide(pres, badgeText, title, technique, questionText, notes, footer)
   *   Structured config: cfuSlide(pres, badgeText, title, { technique, question, ... }, notes, footer)
   *
   * The structured form lets callers pass the full §38 CFU pivot object
   * (technique, question, script, scanFor, proceed, pivot) to a single
   * place. This builder uses .technique and .question for the slide face;
   * the remaining fields are intended to be fed to composeNotes() by the
   * build script.
   *
   * A "CHECK" wordmark renders top-right beside the title so the slide's
   * intent is signalled by label as well as by colour (megaprompt §18a:
   * "Do not rely on colour alone").
   */
  function cfuSlide(pres, badgeText, title, cfuOrTechnique, questionTextOrNotes, notesOrFooter, footerOrUndef) {
    let technique, questionText, notes, footer;
    if (cfuOrTechnique && typeof cfuOrTechnique === "object") {
      technique = cfuOrTechnique.technique;
      questionText = cfuOrTechnique.question != null ? cfuOrTechnique.question : cfuOrTechnique.questionText;
      notes = questionTextOrNotes;
      footer = notesOrFooter;
    } else {
      technique = cfuOrTechnique;
      questionText = questionTextOrNotes;
      notes = notesOrFooter;
      footer = footerOrUndef;
    }

    const s = pres.addSlide();
    el.addTopBar(s, C.ALERT);
    el.addBadge(s, badgeText || "CFU", { color: C.ALERT });
    el.addTitle(s, title || "Check for Understanding", { color: C.ALERT });

    // CHECK wordmark (top-right) — §18a accessible signal beyond colour.
    const stampH = byBand(sz, 0.42, 0.38, 0.32);
    const stampW = byBand(sz, 1.7, 1.5, 1.3);
    el.addTextOnShape(s, "✓  CHECK", {
      x: 9.5 - stampW, y: 0.20, w: stampW, h: stampH, rectRadius: 0.08,
      fill: { color: C.WHITE },
      line: { color: C.ALERT, width: 1.5 },
    }, {
      fontSize: byBand(sz, 14, 13, 11), fontFace: FONT_B, color: C.ALERT,
      bold: true, align: "center", valign: "middle", margin: 0,
    });

    const pillH = byBand(sz, 0.50, 0.46, 0.40);
    const pillW = byBand(sz, 3.4, 3.1, 2.8);
    s.addShape("roundRect", {
      x: 0.5, y: CONTENT_TOP, w: pillW, h: pillH, rectRadius: 0.08,
      fill: { color: C.ALERT },
    });
    s.addText(technique || "Show Me Boards", {
      x: 0.5, y: CONTENT_TOP, w: pillW, h: pillH,
      fontSize: sz.chip + 1, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    const qY = CONTENT_TOP + pillH + 0.16;
    const questionMetrics = getQuestionCardMetrics(questionText || "");
    const qH = Math.min(questionMetrics.cardH, SAFE_BOTTOM - qY);
    el.addCard(s, 0.5, qY, 9, qH, { strip: C.ALERT, fill: C.WHITE });
    // valign:top so any residual overflow goes downward (never upward into
    // the technique pill above). fontSize is pre-computed to fit qH - 0.36
    // via fitTextFontSize, so overflow shouldn't happen, but valign:top is
    // the safer default for multi-line CFU questions.
    s.addText(questionText || "", {
      x: 0.75, y: qY + 0.20, w: 8.5, h: qH - 0.36,
      fontSize: questionMetrics.fontSize, fontFace: FONT_B, color: C.CHARCOAL,
      valign: "top", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    if (footer) el.addFooter(s, footer);
    if (notes) s.addNotes(notes);
    return s;
  }

  /**
   * closingSlide - Dark reflection/review slide for lesson close.
   *
   * Two call shapes:
   *
   *   Legacy positional: closingSlide(pres, reflectionPrompt, takeaways, notes)
   *   Structured config: closingSlide(pres, { reflectionPrompt, scItems, selfAssessment, takeaways }, notes)
   *
   * Megaprompt §52 requires the closing slide to (a) show the three success
   * criteria, (b) include a self-assessment routine, (c) include a short
   * reflection prompt, and (d) acknowledge progress. The structured form
   * delivers all four. The legacy form keeps existing builds rendering
   * unchanged.
   *
   * `selfAssessment` accepts a string ("Thumbs up / sideways / down") or
   * an object { prompt, options: [..] } for richer routines (traffic light
   * etc.). When omitted, no self-assessment row renders (legacy path).
   */
  function closingSlide(pres, configOrPrompt, takeawaysOrNotes, notesOrUndef) {
    let reflectionPrompt = "";
    let takeaways = null;
    let scItems = null;
    let selfAssessment = null;
    let notes = "";

    if (configOrPrompt && typeof configOrPrompt === "object" && !Array.isArray(configOrPrompt)) {
      reflectionPrompt = configOrPrompt.reflectionPrompt || "";
      takeaways = configOrPrompt.takeaways || null;
      scItems = Array.isArray(configOrPrompt.scItems) ? configOrPrompt.scItems.slice(0, 3) : null;
      selfAssessment = configOrPrompt.selfAssessment || null;
      notes = takeawaysOrNotes || "";
    } else {
      reflectionPrompt = configOrPrompt || "";
      takeaways = takeawaysOrNotes;
      notes = notesOrUndef || "";
    }

    const s = pres.addSlide();
    s.background = { color: C.BG_DARK };

    // Accent bar
    s.addShape("rect", { x: 0, y: 0, w: 0.12, h: SLIDE_H, fill: { color: C.ACCENT } });

    // Decorative shapes
    s.addShape("roundRect", {
      x: -1.0, y: 3.2, w: 3.5, h: 3.5, rectRadius: 1.75,
      fill: { color: C.DECOR_1, transparency: 75 },
    });
    s.addShape("roundRect", {
      x: 8.5, y: -0.5, w: 2.5, h: 2.5, rectRadius: 1.25,
      fill: { color: C.DECOR_2, transparency: 80 },
    });

    s.addText("Review & Reflect", {
      x: 0.7, y: 0.45, w: 8, h: 0.9,
      fontSize: sz.closingHero, fontFace: FONT_H, color: C.WHITE, bold: true, margin: 0,
      fit: "shrink", shrinkText: true,
    });

    let cursorY = 1.40;
    if (scItems && scItems.length) {
      s.addText("Show what you can do", {
        x: 0.7, y: cursorY, w: 8.0, h: 0.36,
        fontSize: sz.sectionLabel + 2, fontFace: FONT_B, color: accentOnDark, bold: true, margin: 0,
      });
      cursorY += 0.42;

      // Megaprompt §0a item 17 / §52: tier labels (Everyone/Most/Stretch) must
      // NEVER appear on the closing slide. Render the success criteria as a
      // plain unlabelled "I can..." list, sized for the year band.
      const rowH = byBand(sz, 0.46, 0.42, 0.34);
      const rowFontSize = sz.takeaway;
      const dotR = byBand(sz, 0.10, 0.09, 0.07);
      const textX = 0.7 + dotR * 2 + 0.20;
      const textW = 9.0 - textX - 0.2;

      scItems.forEach((text, i) => {
        const rowY = cursorY + i * (rowH + 0.06);
        const dotY = rowY + (rowH / 2) - dotR;
        s.addShape("roundRect", {
          x: 0.7, y: dotY,
          w: dotR * 2, h: dotR * 2,
          rectRadius: dotR,
          fill: { color: accentOnDark },
        });
        s.addText(String(text || ""), {
          x: textX, y: rowY,
          w: textW, h: rowH,
          fontSize: rowFontSize, fontFace: FONT_B, color: C.TEXT_ON_DARK || C.WHITE,
          valign: "middle", margin: 0,
          fit: "shrink", shrinkText: true,
        });
      });

      cursorY += scItems.length * (rowH + 0.06) + 0.08;
    } else {
      // Legacy "Turn & Talk" header preserved for back-compat decks.
      s.addText("Turn & Talk", {
        x: 0.7, y: 1.45, w: 3.0, h: 0.42,
        fontSize: sz.sectionLabel + 2, fontFace: FONT_B, color: accentOnDark, bold: true, margin: 0,
      });
      cursorY = 1.96;
    }

    if (reflectionPrompt) {
      const promptH = scItems ? 0.7 : 1.2;
      s.addText(reflectionPrompt, {
        x: 0.7, y: cursorY, w: 8.5, h: promptH,
        fontSize: sz.closingPrompt, fontFace: FONT_B, color: subtitleOnDark, italic: true, margin: 0,
        fit: "shrink", shrinkText: true,
      });
      cursorY += promptH + 0.08;
    }

    if (selfAssessment) {
      const selfPromptText = typeof selfAssessment === "string"
        ? selfAssessment
        : (selfAssessment.prompt || "Self-assess: thumbs up, sideways, or down.");
      const optionList = (selfAssessment && Array.isArray(selfAssessment.options) && selfAssessment.options.length)
        ? selfAssessment.options
        : ["Got it", "Getting there", "Need more practice"];
      const rowH = byBand(sz, 0.5, 0.46, 0.38);
      if (cursorY + rowH <= SAFE_BOTTOM) {
        const tagW = byBand(sz, 1.7, 1.55, 1.35);
        el.addTextOnShape(s, "Self-check", {
          x: 0.7, y: cursorY, w: tagW, h: rowH, rectRadius: 0.08,
          fill: { color: C.ACCENT },
        }, {
          fontSize: byBand(sz, 14, 13, 11), fontFace: FONT_B, color: C.WHITE,
          bold: true, align: "center", valign: "middle", margin: 0,
        });
        const optionsX = 0.7 + tagW + 0.18;
        const optionsW = 9.0 - optionsX - 0.2;
        const optionGap = 0.18;
        const perOptW = Math.max(1.4, (optionsW - optionGap * (optionList.length - 1)) / optionList.length);
        // Option pills use a transparent fill so the dark BG shows through;
        // addTextOnShape's contrast check would false-warn on the unblended
        // FFFFFF fill, so use raw addShape + addText here.
        optionList.slice(0, 4).forEach((opt, i) => {
          const ox = optionsX + i * (perOptW + optionGap);
          s.addShape("roundRect", {
            x: ox, y: cursorY, w: perOptW, h: rowH, rectRadius: 0.08,
            fill: { color: C.WHITE, transparency: 80 },
            line: { color: accentOnDark, width: 0.6 },
          });
          s.addText(String(opt), {
            x: ox, y: cursorY, w: perOptW, h: rowH,
            fontSize: byBand(sz, 13, 12, 10), fontFace: FONT_B, color: C.TEXT_ON_DARK,
            bold: true, align: "center", valign: "middle", margin: 0,
            fit: "shrink", shrinkText: true,
          });
        });
        if (selfPromptText && cursorY + rowH + 0.30 <= SAFE_BOTTOM && !scItems) {
          s.addText(String(selfPromptText), {
            x: 0.7, y: cursorY + rowH + 0.04, w: 8.5, h: 0.26,
            fontSize: byBand(sz, 13, 12, 10), fontFace: FONT_B, color: subtitleOnDark, italic: true, margin: 0,
          });
        }
        cursorY += rowH + 0.10;
      }
    }

    if (takeaways && takeaways.length) {
      const rowPitch = byBand(sz, 0.44, 0.40, 0.34);
      const takeawayY0 = scItems
        ? cursorY
        : (sz._band === "F" ? 3.30 : 3.20);
      if (takeawayY0 + 0.42 <= SAFE_BOTTOM) {
        s.addText("Key Takeaways", {
          x: 0.7, y: takeawayY0, w: 5, h: 0.36,
          fontSize: sz.sectionLabel + 2, fontFace: FONT_B, color: accentOnDark, bold: true, margin: 0,
        });
        takeaways.forEach((t, i) => {
          const y = takeawayY0 + 0.42 + i * rowPitch;
          if (y + (rowPitch - 0.06) > SAFE_BOTTOM) return;
          s.addText("•  " + t, {
            x: 0.9, y, w: 8.0, h: rowPitch - 0.06,
            fontSize: sz.takeaway, fontFace: FONT_B, color: C.TEXT_ON_DARK, margin: 0,
            fit: "shrink", shrinkText: true,
          });
        });
      }
    }

    if (notes) s.addNotes(notes);
    return s;
  }

  /**
   * annotatedModelSlide - Universal visual-anchor slide for feature spotting,
   * labelled source structure, and "notice this part" teaching.
   */
  function annotatedModelSlide(pres, badgeText, title, prompts, modelTitle, features, notes, footer, opts) {
    const s = pres.addSlide();
    const o = opts || {};
    const promptItems = Array.isArray(prompts) ? prompts : [];
    const featureItems = Array.isArray(features) ? features : [];
    const stripColors = [C.PRIMARY, C.SECONDARY, C.ACCENT, C.ALERT, C.SUCCESS];

    el.addTopBar(s, o.badgeColor || C.SECONDARY);
    el.addBadge(s, badgeText || "Notice", { color: o.badgeColor || C.SECONDARY, w: o.badgeW || 1.95 });
    el.addTitle(s, title);

    const leftX = 0.5;
    const leftY = CONTENT_TOP;
    const leftW = o.leftW || 3.25;
    const gap = 0.2;
    const rightX = leftX + leftW + gap;
    const rightW = 9 - leftW - gap;
    const cardH = SAFE_BOTTOM - CONTENT_TOP;

    const normalizedPrompts = promptItems.map((item, index) => {
      if (typeof item === "string") {
        return { text: item, role: index === 0 ? "header" : "body" };
      }
      return {
        text: String((item && item.text) || ""),
        role: item && item.role ? item.role : (index === 0 ? "header" : "body"),
        bold: Boolean(item && item.bold),
        italic: Boolean(item && item.italic),
        color: item && item.color,
      };
    });

    el.addInstructionCard(s, normalizedPrompts, {
      x: leftX,
      y: leftY,
      w: leftW,
      h: cardH,
      strip: o.promptStrip || C.SECONDARY,
      fill: o.promptFill || C.WHITE,
      headerColor: o.promptHeaderColor || C.SECONDARY,
      emphasisColor: o.promptEmphasisColor || C.ALERT,
    });

    el.addCard(s, rightX, CONTENT_TOP, rightW, cardH, {
      strip: o.modelStrip || C.PRIMARY,
      fill: o.modelFill || C.WHITE,
      shadow: o.shadow || shadowFn(),
    });

    const sourceType = o.sourceType ? String(o.sourceType) : "";
    if (sourceType) {
      s.addText(sourceType, {
        x: rightX + 0.2, y: CONTENT_TOP + 0.08, w: rightW - 0.4, h: 0.22,
        fontSize: sz.caption + 0.5, fontFace: FONT_B, color: C.MUTED, bold: true, margin: 0,
      });
    }

    s.addText(String(modelTitle || "Model"), {
      x: rightX + 0.2, y: CONTENT_TOP + (sourceType ? 0.28 : 0.12), w: rightW - 0.4, h: 0.36,
      fontSize: byBand(sz, 21, 19, 17),
      fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0,
      fit: "shrink", shrinkText: true,
    });

    const subtitle = o.modelSubtitle ? String(o.modelSubtitle) : "";
    if (subtitle) {
      s.addText(subtitle, {
        x: rightX + 0.2, y: CONTENT_TOP + 0.50, w: rightW - 0.4, h: 0.26,
        fontSize: sz.caption + 1, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
        fit: "shrink", shrinkText: true,
      });
    }

    const previewY = CONTENT_TOP + (subtitle ? 0.78 : 0.64);
    const featureCount = Math.min(featureItems.length, 4);
    const featureCols = featureCount <= 1 ? 1 : 2;
    const featureRows = featureCount > 0 ? Math.ceil(featureCount / featureCols) : 0;
    const featureGapX = 0.12;
    const featureGapY = 0.08;
    const featureGridH = featureRows === 0 ? 0 : (featureRows === 1 ? 0.62 : 1.24);
    const previewH = Math.max(1.18, Math.min(1.42, cardH - (previewY - CONTENT_TOP) - featureGridH - (featureCount > 0 ? 0.26 : 0.14)));
    drawMockupPreview(s, rightX + 0.2, previewY, rightW - 0.4, previewH, o.previewSpec || o.previewBlocks, {
      fill: o.previewFill,
      border: o.previewBorder,
      accent: o.previewAccent || C.PRIMARY,
      innerPad: 0.08,
      gap: 0.06,
    });

    if (featureCount > 0) {
      const gridTop = previewY + previewH + 0.12;
      const gridW = rightW - 0.4;
      const cellW = featureCols === 1 ? gridW : (gridW - featureGapX) / 2;
      const cellH = featureRows === 1 ? featureGridH : (featureGridH - featureGapY) / 2;

      featureItems.slice(0, 4).forEach((feature, index) => {
        const row = Math.floor(index / featureCols);
        const col = index % featureCols;
        const cellX = rightX + 0.2 + col * (cellW + featureGapX);
        const cellY = gridTop + row * (cellH + featureGapY);
        const label = typeof feature === "string"
          ? `Feature ${index + 1}`
          : String(feature.label || `Feature ${index + 1}`);
        const detail = typeof feature === "string"
          ? feature
          : String(feature.detail || feature.text || "");
        const stripColor = typeof feature === "object" && feature.color
          ? feature.color
          : stripColors[index % stripColors.length];
        drawFeatureKeyCard(s, cellX, cellY, cellW, cellH, label, detail, stripColor);
      });
    }

    if (footer) el.addFooter(s, footer);
    if (notes) s.addNotes(notes);
    return s;
  }

  /**
   * compareVisualSlide - Side-by-side visual comparison slide for We Do analysis.
   * Keeps the visual objects on screen while fading labels/prompts.
   */
  function compareVisualSlide(pres, badgeText, title, promptText, leftModel, rightModel, notes, footer, opts) {
    const s = pres.addSlide();
    const o = opts || {};
    const left = leftModel || {};
    const right = rightModel || {};

    el.addTopBar(s, o.badgeColor || C.PRIMARY);
    el.addBadge(s, badgeText || "We Do", { color: o.badgeFill || C.SUCCESS, w: o.badgeW || 1.4 });
    el.addTitle(s, title);

    const cardY = CONTENT_TOP;
    const cardH = o.cardH || 2.15;
    const leftX = 0.5;
    const gap = 0.2;
    const leftW = o.leftW || 4.15;
    const rightX = leftX + leftW + gap;
    const rightW = 9 - leftW - gap;

    function drawCompareCard(x, w, model, stripColor) {
      el.addCard(s, x, cardY, w, cardH, { strip: stripColor, fill: C.WHITE });
      s.addText(String(model.panelTitle || "Option"), {
        x: x + 0.25,
        y: cardY + 0.08,
        w: w - 0.5,
        h: 0.24,
        fontSize: sz.sectionLabel + 3,
        fontFace: FONT_H,
        color: stripColor,
        bold: true,
        margin: 0,
        fit: "shrink", shrinkText: true,
      });
      if (model.title) {
        s.addText(String(model.title), {
          x: x + 0.25,
          y: cardY + 0.34,
          w: w - 0.5,
          h: 0.24,
          fontSize: sz.sectionLabel + 1,
          fontFace: FONT_H,
          color: C.CHARCOAL,
          bold: true,
          margin: 0,
          fit: "shrink", shrinkText: true,
        });
      }

      const previewY = cardY + 0.60;
      const previewH = model.previewH || 1.26;
      drawMockupPreview(s, x + 0.22, previewY, w - 0.44, previewH, model.previewSpec || model.previewBlocks, {
        accent: model.previewAccent || stripColor,
        fill: model.previewFill,
        border: model.previewBorder,
        innerPad: 0.08,
        gap: 0.05,
      });
    }

    drawCompareCard(leftX, leftW, left, left.strip || C.SECONDARY);
    drawCompareCard(rightX, rightW, right, right.strip || C.PRIMARY);

    const promptH = byBand(sz, 0.62, 0.56, 0.48);
    const promptY = cardY + cardH + 0.15;
    el.addTextOnShape(s, String(promptText || ""), {
      x: 0.5,
      y: promptY,
      w: 9,
      h: promptH,
      rectRadius: 0.08,
      fill: { color: o.promptFill || C.ALERT },
    }, {
      fontSize: o.promptFontSize || (byBand(sz, 18, 15.5, 12.5)),
      fontFace: FONT_B,
      color: C.WHITE,
      bold: true,
      align: "left",
      valign: "middle",
      margin: 0.08,
    });

    if (footer) el.addFooter(s, footer);
    if (notes) s.addNotes(notes);
    return s;
  }

  /**
   * boardBuildSlide - megaprompt §19 "Build this together" canvas.
   *
   * A short title, an optional visual prompt, and a deliberately blank
   * (or partially blank) build space the teacher fills in live with the
   * class. Use for: number lines, area models, anchor charts, vocabulary
   * webs, sentence construction, sorting examples, misconception fixes.
   *
   * Teacher notes carry the script for what to write/draw. The slide
   * face deliberately stays sparse so attention stays on the live build.
   *
   * @param {object}   pres
   * @param {string}   badgeText      Defaults to "Build Together"
   * @param {string}   title          Slide title (eg. "Build the number line")
   * @param {string}   directive      Short student-facing direction (eg. "Build this together")
   * @param {string}   notes          Teacher notes
   * @param {string}   footer         Footer text
   * @param {object}   [opts]         { promptText, prefilledHints, badgeColor, canvasFill }
   *
   * opts.prefilledHints (string[]) — small grey markers placed inside the
   * canvas so the build has a starting point without giving away the
   * answer (eg. ["0", "?", "10"]).
   */
  function boardBuildSlide(pres, badgeText, title, directive, notes, footer, opts) {
    const s = pres.addSlide();
    const o = opts || {};
    const stripColor = o.badgeColor || C.SECONDARY;
    el.addTopBar(s, stripColor);
    el.addBadge(s, badgeText || "Build Together", { color: stripColor, w: byBand(sz, 2.6, 2.4, 2.2) });
    el.addTitle(s, title);

    const directiveH = byBand(sz, 0.74, 0.66, 0.56);
    const directiveY = CONTENT_TOP;
    el.addTextOnShape(s, String(directive || "Build this together"), {
      x: 0.5, y: directiveY, w: 9, h: directiveH, rectRadius: 0.08,
      fill: { color: stripColor },
    }, {
      fontSize: byBand(sz, 26, 22, 18),
      fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "left", valign: "middle", margin: 0.16,
    });

    const canvasY = directiveY + directiveH + 0.16;
    const canvasH = SAFE_BOTTOM - canvasY;
    s.addShape("roundRect", {
      x: 0.5, y: canvasY, w: 9, h: canvasH, rectRadius: 0.08,
      fill: { color: o.canvasFill || C.WHITE },
      line: { color: C.MUTED, width: 1.2, dashType: "dash" },
    });

    // Optional prompt (small) inside the canvas top-left.
    if (o.promptText) {
      s.addText(String(o.promptText), {
        x: 0.7, y: canvasY + 0.16, w: 8.6, h: 0.36,
        fontSize: byBand(sz, 16, 14, 12),
        fontFace: FONT_B, color: C.MUTED, italic: true,
        align: "left", valign: "top", margin: 0,
      });
    }

    // Optional pre-filled hints — small grey anchors.
    const hints = Array.isArray(o.prefilledHints) ? o.prefilledHints : [];
    if (hints.length) {
      const hintsY = canvasY + canvasH - byBand(sz, 0.7, 0.6, 0.5);
      const hintW = 9 / Math.max(hints.length, 1);
      hints.forEach((hint, i) => {
        const hx = 0.5 + i * hintW + hintW / 2 - 0.4;
        s.addText(String(hint), {
          x: hx, y: hintsY, w: 0.8, h: byBand(sz, 0.5, 0.45, 0.36),
          fontSize: byBand(sz, 22, 19, 16),
          fontFace: FONT_H, color: C.MUTED, bold: true,
          align: "center", valign: "middle", margin: 0,
        });
      });
    }

    if (footer) el.addFooter(s, footer);
    if (notes) s.addNotes(notes);
    return s;
  }

  /**
   * exitTicketSlide (universal) - megaprompt §53 evidence-of-learning slide.
   *
   * Renders 1-3 prompt cards aligned to a Success Criterion (typically SC2).
   * Question numbering is OFF by default — exit tickets are formal
   * assessment items per §15d, so callers may opt into numbering with
   * `opts.numbered: true` when collecting handed-in tickets requires it.
   *
   * @param {object}   pres
   * @param {string|string[]} prompts  one or more exit-ticket prompts
   * @param {string}   notes
   * @param {string}   footer
   * @param {object}   [opts]   { numbered, assessesSc, badgeColor, title }
   */
  function exitTicketSlide(pres, prompts, notes, footer, opts) {
    const s = pres.addSlide();
    const o = opts || {};
    const promptList = Array.isArray(prompts) ? prompts : [prompts];
    const cleaned = promptList
      .map((p) => String(p == null ? "" : p).trim())
      .filter(Boolean)
      .slice(0, sz.maxQuestions || 3);

    const stripColor = o.badgeColor || C.ASSESS || C.ALERT;

    s.background = { color: C.BG_CARD };
    s.addShape("rect", { x: 0, y: 0, w: 10, h: 0.06, fill: { color: stripColor } });

    const badgeH = byBand(sz, 0.42, 0.40, 0.36);
    const badgeW = byBand(sz, 2.1, 1.95, 1.8);
    el.addBadge(s, "Exit Ticket", { color: stripColor, x: 0.5, y: 0.2, w: badgeW, h: badgeH });

    if (o.assessesSc) {
      const scTagW = byBand(sz, 1.6, 1.5, 1.3);
      el.addTextOnShape(s, `Assesses SC${o.assessesSc}`, {
        x: 0.5 + badgeW + 0.16, y: 0.2, w: scTagW, h: badgeH, rectRadius: 0.08,
        fill: { color: C.WHITE },
        line: { color: stripColor, width: 1.2 },
      }, {
        fontSize: byBand(sz, 12, 11, 10),
        fontFace: FONT_B, color: stripColor,
        align: "center", valign: "middle", bold: true, margin: 0,
      });
    }

    const titleY = byBand(sz, 0.72, 0.68, 0.65);
    const titleH = byBand(sz, 0.74, 0.68, 0.62);
    s.addText(o.title || builderDefaults.exitTicketTitle || "Show what you know", {
      x: 0.5, y: titleY, w: 9, h: titleH,
      fontSize: sz.titleH1 - 2, fontFace: FONT_H, color: stripColor, bold: true, margin: 0,
      fit: "shrink", shrinkText: true,
    });

    const startY = titleY + titleH + 0.18;
    if (cleaned.length === 0) {
      console.warn("[exitTicketSlide] no prompts provided; rendered title, footer, and notes only.");
      if (footer) el.addFooter(s, footer);
      if (notes) s.addNotes(notes);
      return s;
    }
    const perH = Math.min(
      byBand(sz, 2.1, 1.5, 1.2),
      (SAFE_BOTTOM - startY) / Math.max(cleaned.length, 1) - 0.12,
    );
    const numbered = Boolean(o.numbered);

    cleaned.forEach((q, i) => {
      const qY = startY + i * (perH + 0.12);
      el.addCard(s, 0.5, qY, 9, perH, { strip: stripColor });
      const display = numbered ? `${i + 1}.  ${q}` : q;
      s.addText(display, {
        x: 0.75, y: qY + 0.08, w: 8.5, h: perH - 0.16,
        fontSize: sz.body, fontFace: FONT_B, color: C.CHARCOAL, margin: 0, valign: "middle",
        fit: "shrink", shrinkText: true,
      });
    });

    if (footer) el.addFooter(s, footer);
    if (notes) s.addNotes(notes);
    return s;
  }

  return {
    titleSlide,
    liSlide,
    contentSlide,
    cfuSlide,
    closingSlide,
    annotatedModelSlide,
    compareVisualSlide,
    boardBuildSlide,
    exitTicketSlide,
    addRevealAnswerBar,
  };
}

module.exports = { createBaseBuilders };
