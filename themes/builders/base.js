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
  estimateTextHeight,
} = require("../core/bulletFit");
const { SUBJECT_PICTOGRAMS } = require("../core/pictograms");

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
  // Pictograms and the declarative visual layer are injected by the factory.
  // Builders degrade gracefully when a caller constructs them without.
  const picto = builderDefaults.picto || null;
  const visual = builderDefaults.visual || null;
  const subjectKey = builderDefaults.subject || "";
  const subjectGlyph = SUBJECT_PICTOGRAMS[subjectKey] || "star";
  const isSpec = (value) => Boolean(visual && visual.isVisualSpec && visual.isVisualSpec(value));

  /**
   * Accept either a drawRight callback or a visual spec for the right column.
   * A spec is fitted into the column frame by drawVisual, so a build script
   * can write `{ type: "tensFrame", filled: 7 }` instead of coordinates.
   */
  function resolveDrawRight(drawRight, frame) {
    if (typeof drawRight === "function") return drawRight;
    if (isSpec(drawRight)) {
      return (slide, guide) => {
        const f = frame || {
          x: guide.rightX, y: guide.panelTop + 0.05,
          w: guide.rightW, h: guide.safeBottom - guide.panelTop - 0.1,
        };
        visual.drawVisual(slide, drawRight, f);
      };
    }
    return null;
  }

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
    // Must READ as a picture at a glance (sun + mountains), not as grey
    // bars - students are told "the image shows X" and need to see an
    // image region, even in a wireframe.
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

    // Sun (top-right)
    const sunD = Math.max(0.1, Math.min(innerW * 0.16, innerH * 0.3));
    slide.addShape("ellipse", {
      x: innerX + innerW * 0.74, y: innerY + innerH * 0.12, w: sunD, h: sunD,
      fill: { color: lightenHex(spec.accent, 0.25) },
      line: { color: lightenHex(spec.accent, 0.25), width: 0.2 },
    });

    // Two overlapping mountains rising from the base
    const baseY = innerY + innerH - 0.03;
    const m1H = innerH * 0.55;
    const m2H = innerH * 0.42;
    slide.addShape("triangle", {
      x: innerX + innerW * 0.06, y: baseY - m1H, w: innerW * 0.52, h: m1H,
      fill: { color: lightenHex(spec.accent, 0.45) },
      line: { color: lightenHex(spec.accent, 0.3), width: 0.4 },
    });
    slide.addShape("triangle", {
      x: innerX + innerW * 0.42, y: baseY - m2H, w: innerW * 0.52, h: m2H,
      fill: { color: lightenHex(spec.accent, 0.3) },
      line: { color: lightenHex(spec.accent, 0.2), width: 0.4 },
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
      } else if (component.kind === "photo" || component.kind === "image") {
        drawPhotoPlaceholder(slide, innerX, cursorY, innerW, blockH, component, normalized);
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
    // Join with a visible separator, never runs of spaces (they render as
    // accidental-looking gaps and are flagged by the slide-text hygiene gate).
    const text = cleaned.join("  |  ");

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
    const tickCueW = byBand(sz, 1.8, 1.6, 1.4);
    slide.addText(String(text), {
      x: x + labelOffset, y: y + 0.12,
      // Stop before the Tick & fix cue so the two text boxes never overlap
      w: w - labelOffset - (showTick ? tickCueW + 0.32 : 0.18),
      h: h - 0.24,
      fontSize, fontFace: FONT_H,
      color: textColor, bold: true,
      align: "left", valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    if (showTick) {
      const cueY = y + h - byBand(sz, 0.32, 0.28, 0.24);
      const cueW = tickCueW;
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
   *
   * Composition: left accent bar, title / subtitle / meta in the left 60%,
   * and on the right a single calm motif: the subject glyph in a soft
   * circle, or the lesson's own visual anchor when `opts.visual` is a
   * visual spec (e.g. { type: "tensFrame", filled: 10 }). No decorative
   * blobs, no gradients: one motif, repeated on the closing slide, is what
   * makes a deck read as designed rather than generated.
   *
   *   titleSlide(pres, title, subtitle, meta, notes, { visual, glyph })
   */
  function titleSlide(pres, title, subtitle, meta, notes, opts) {
    const o = opts || {};
    const s = pres.addSlide();
    s.background = { color: C.BG_DARK };

    // Vertical accent bar
    s.addShape("rect", { x: 0, y: 0, w: 0.12, h: SLIDE_H, fill: { color: C.ACCENT } });

    // Right-hand motif panel
    const panel = { x: 6.35, y: 0.95, w: 3.15, h: 3.7 };
    const hasVisual = isSpec(o.visual);
    if (hasVisual) {
      s.addShape("roundRect", {
        x: panel.x, y: panel.y, w: panel.w, h: panel.h, rectRadius: 0.16,
        fill: { color: C.WHITE },
      });
      visual.drawVisual(s, o.visual, { x: panel.x + 0.22, y: panel.y + 0.22, w: panel.w - 0.44, h: panel.h - 0.44 });
    } else if (picto && o.glyph !== false) {
      const d = 2.5;
      const cx = panel.x + panel.w / 2;
      const cy = panel.y + panel.h / 2;
      s.addShape("roundRect", {
        x: cx - d / 2, y: cy - d / 2, w: d, h: d, rectRadius: d / 2,
        fill: { color: C.BG_DARK_PANEL || C.PRIMARY },
      });
      picto.addPictogram(s, o.glyph || subjectGlyph, cx - d * 0.31, cy - d * 0.31, d * 0.62, {
        style: "flat", color: C.WHITE, glyphColor: C.WHITE,
      });
    }

    const textW = hasVisual || (picto && o.glyph !== false) ? 5.5 : 8.0;
    const heroH = byBand(sz, 1.7, 1.55, 1.4);
    s.addText(title, {
      x: 0.7, y: 0.95, w: textW, h: heroH,
      fontSize: sz.titleHero, fontFace: FONT_H, color: C.WHITE, bold: true, margin: 0,
      valign: "bottom",
      fit: "shrink", shrinkText: true,
    });

    // Subtitle
    if (subtitle) {
      const subY = 0.95 + heroH + 0.12;
      s.addText(subtitle, {
        x: 0.7, y: subY, w: textW, h: 0.95,
        fontSize: sz.subtitleHero, fontFace: FONT_B, color: subtitleOnDark, margin: 0,
        valign: "top",
        fit: "shrink", shrinkText: true,
      });
    }

    // Meta line as a small pill so it reads as a label, not more prose
    if (meta) {
      const metaY = SLIDE_H - 0.95;
      const metaFont = sz.metaHero;
      const metaW = Math.min(textW, Math.max(2.4, String(meta).length * metaFont * 0.0105 + 0.6));
      s.addShape("roundRect", {
        x: 0.7, y: metaY, w: metaW, h: 0.44, rectRadius: 0.22,
        fill: { color: C.BG_DARK_PANEL || C.PRIMARY },
      });
      s.addText(meta, {
        x: 0.7, y: metaY, w: metaW, h: 0.44,
        fontSize: metaFont, fontFace: FONT_B, color: C.WHITE, margin: 0.12,
        align: "center", valign: "middle",
        fit: "shrink", shrinkText: true,
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
    el.addCard(s, 0.5, CONTENT_TOP, 9, liH, { variant: "tint", tone: C.PRIMARY, strip: C.PRIMARY });
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
    s.addText("Success Criteria", {
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
   *
   * `bullets` may be a string (one statement) or an array. `drawRight` may be
   * a callback(slide, layoutGuide) or a visual spec such as
   * { type: "fractionStrips", strips: [...] }, which is fitted into the right
   * column automatically.
   *
   * Density-aware: one or two short lines with no right column render as a
   * hero statement (large type on a soft tint panel, vertically centred)
   * instead of small bullets floating at the top of an empty slide
   * (megaprompt 15b / 15h / 16).
   */
  function contentSlide(pres, badgeText, badgeColor, title, bullets, notes, footer, drawRight) {
    const s = pres.addSlide();
    const tone = badgeColor || C.PRIMARY;
    el.addTopBar(s, C.PRIMARY);
    el.addBadge(s, badgeText || "Content", { color: tone });
    el.addTitle(s, title);

    const rightFn = resolveDrawRight(drawRight);
    const items = bullets == null ? [] : (Array.isArray(bullets) ? bullets : [String(bullets)]);
    const cardW = rightFn ? 4.5 : 9.0;
    const contentY = CONTENT_TOP;
    const layoutGuide = {
      titleY: 0.65,
      titleH: 0.62,
      panelTop: contentY,
      panelTopPadded: contentY + 0.08,
      leftCardX: 0.5,
      leftCardY: contentY,
      leftCardW: cardW,
      leftCardH: 0,
      rightX: 5.2,
      rightW: 4.3,
      safeBottom: SAFE_BOTTOM,
    };

    const prepared = prepareBullets(items);
    const joined = prepared.map((p) => p.text).join(" ");
    const areaH = SAFE_BOTTOM - contentY;
    // Sparse = up to three short statements and no right column. Two
    // launch questions or three You Do steps are the hero of their slide,
    // so they are set large on a tint panel instead of floating as small
    // bullets above an empty half-slide.
    const sparse = !rightFn && prepared.length > 0 && prepared.length <= 3 &&
      joined.length <= byBand(sz, 80, 110, 170);

    if (sparse) {
      const charsPerLine = byBand(sz, 30, 38, 48);
      const ideal = prepared.length === 1 ? sz.heroQuestion : (prepared.length === 2 ? sz.hero : Math.round(sz.hero * 0.9));
      const floor = sz.body;
      // Reveal-safe: many decks pair a contentSlide with a withReveal answer
      // bar placed by the script at y >= 3.9. The hero panel therefore stays
      // inside the top 2.75" of the content area and is centred within that
      // region, so the grown text can never poke out beneath a bar.
      const safeH = areaH - 1.05;
      const maxTextH = safeH - 0.5;
      const textBlock = prepared.map((p) => p.text).join("\n");
      const fontSize = fitTextFontSize(textBlock, maxTextH, charsPerLine, ideal, floor);
      const textH = Math.min(maxTextH, estimateTextHeight(textBlock, fontSize, charsPerLine) + (prepared.length - 1) * 0.2);
      const cardH = Math.max(1.75, Math.min(safeH, textH + 0.7));
      const cardY = contentY + Math.max(0, (safeH - cardH) * 0.5);
      layoutGuide.leftCardY = cardY;
      layoutGuide.leftCardH = cardH;
      el.addCard(s, 0.5, cardY, 9, cardH, { variant: "tint", tone });
      s.addText(prepared.map((item, i) => ({
        text: item.text,
        options: {
          bullet: prepared.length > 1,
          breakLine: i < prepared.length - 1,
          fontSize,
          color: C.CHARCOAL,
          paraSpaceAfter: prepared.length > 1 ? fontSize * 0.45 : 0,
        },
      })), {
        x: 0.85, y: cardY + 0.25, w: 8.3, h: cardH - 0.5,
        fontFace: FONT_B, valign: "middle", margin: 0,
        align: prepared.length === 1 ? "center" : "left",
        fit: "shrink", shrinkText: true,
      });
    } else {
      const metrics = prepared.length ? getBulletCardMetrics(items, { narrow: Boolean(rightFn) }) : null;
      // A visual spec fills the whole right column, so the left card runs the
      // full height too and its few lines sit centred and larger: two
      // balanced panels, not a short card beside a floating model.
      const specRight = isSpec(drawRight);
      const longest = prepared.reduce((m, p) => Math.max(m, p.text.length), 0);
      const fewShort = prepared.length > 0 && prepared.length <= 3 && longest <= 34;
      const cardH = specRight ? areaH : (metrics ? metrics.cardH : (rightFn ? 2.0 : 1.55));
      const fontSize = metrics ? (specRight && fewShort ? Math.max(metrics.fontSize, sz.body) : metrics.fontSize) : sz.body;
      layoutGuide.leftCardH = cardH;

      el.addCard(s, 0.5, contentY, cardW, cardH, {
        strip: tone,
        fill: C.WHITE,
      });

      if (metrics) {
        const baseSpacePt = fontSize >= 16 ? 5 : 3;
        s.addText(metrics.prepared.map((item, i) => ({
          text: item.text,
          options: {
            bullet: true,
            breakLine: i < metrics.prepared.length - 1,
            fontSize,
            color: C.CHARCOAL,
            // Per-paragraph space-after grows when the build script used an
            // empty-string spacer between bullets. Avoids ugly empty bullet
            // markers while preserving visual grouping.
            paraSpaceAfter: baseSpacePt + (item.extraSpaceAfter || 0) * fontSize * 0.9,
          },
        })), {
          x: 0.75, y: contentY + metrics.topInset, w: cardW - 0.5, h: cardH - metrics.topInset * 2,
          fontFace: FONT_B, valign: specRight ? "middle" : "top", margin: 0,
        });
      }
    }

    if (rightFn) rightFn(s, layoutGuide);
    if (footer) el.addFooter(s, footer);
    if (notes) s.addNotes(notes);
    if (rightFn) runSlideDiagnostics(s, pres);
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

    // Density-aware sizing: a short question must not sit in a small
    // top-anchored card with the bottom half of the slide empty
    // (megaprompt 15h - the hero fills the space). Compute the content
    // block first, enlarge sparse questions, then centre the block.
    const availForCard = SAFE_BOTTOM - (CONTENT_TOP + pillH + 0.16);
    const questionMetrics = getQuestionCardMetrics(questionText || "");
    let qFontSize = questionMetrics.fontSize;
    let qH = Math.min(questionMetrics.cardH, availForCard);
    const leftover = availForCard - qH;
    const isSparse = leftover > 0.9;
    if (isSparse) {
      // Grow the hero, but keep the grown card clear of a withReveal answer
      // bar (0.8" bar + 0.15" clearance + margin) whenever possible - CFU
      // slides are the most common reveal-pair base.
      const targetH = Math.min(
        Math.max(availForCard - 1.1, qH),
        Math.max(qH, availForCard * 0.72)
      );
      const growth = targetH / Math.max(qH, 0.1);
      qFontSize = Math.min(
        Math.round(qFontSize * Math.min(growth, 1.4)),
        byBand(sz, 44, 40, 34)
      );
      qH = targetH;
    }

    const blockH = pillH + 0.16 + qH;
    const yOffset = Math.max(0, (SAFE_BOTTOM - CONTENT_TOP - blockH) * 0.4);
    const pillY = CONTENT_TOP + yOffset;

    s.addShape("roundRect", {
      x: 0.5, y: pillY, w: pillW, h: pillH, rectRadius: 0.08,
      fill: { color: C.ALERT },
    });
    s.addText(technique || "Show Me Boards", {
      x: 0.5, y: pillY, w: pillW, h: pillH,
      fontSize: sz.chip + 1, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    const qY = pillY + pillH + 0.16;
    el.addCard(s, 0.5, qY, 9, qH, { variant: "tint", tone: C.ALERT });
    // Sparse questions centre vertically at hero size; longer questions keep
    // valign:top so residual overflow goes downward (never up into the pill).
    // fontSize is pre-computed to fit via fitTextFontSize.
    s.addText(questionText || "", {
      x: 0.75, y: qY + 0.20, w: 8.5, h: qH - 0.36,
      fontSize: qFontSize, fontFace: FONT_B, color: C.CHARCOAL,
      valign: isSparse ? "middle" : "top", margin: 0,
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

    // The same subject motif as the title slide, small, top right: the deck
    // opens and closes on one visual idea.
    if (picto) {
      const d = 0.9;
      s.addShape("roundRect", {
        x: 9.5 - d, y: 0.45, w: d, h: d, rectRadius: d / 2,
        fill: { color: C.BG_DARK_PANEL || C.PRIMARY },
      });
      picto.addPictogram(s, subjectGlyph, 9.5 - d + d * 0.19, 0.45 + d * 0.19, d * 0.62, {
        style: "flat", color: C.WHITE, glyphColor: C.WHITE,
      });
    }

    s.addText("Review & Reflect", {
      x: 0.7, y: 0.45, w: 7.6, h: 0.9,
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
    // Sized for the raised featureLabel/featureDetail floors - annotation
    // text students read must be legible from the back of the room.
    const featureGridH = featureRows === 0 ? 0 : (featureRows === 1 ? 0.74 : 1.5);
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
   * `opts.assessesSc` records which Success Criterion the ticket targets, but
   * the visible "Assesses SC{n}" chip is OFF by default. Megaprompt §0a item
   * 18 / §53: SC1/SC2/SC3 numbering is an internal planning tool and must not
   * appear on any student-facing surface. The SC target belongs in the teacher
   * notes. Render the chip only for a teacher-facing review export by passing
   * `opts.showAssessesTag: true`.
   *
   * @param {object}   pres
   * @param {string|string[]} prompts  one or more exit-ticket prompts
   * @param {string}   notes
   * @param {string}   footer
   * @param {object}   [opts]   { numbered, assessesSc, showAssessesTag, badgeColor, title }
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

    // §0a item 18 / §53: keep internal SC numbering off the student-facing
    // exit ticket. Only render the chip for an explicit teacher-facing review
    // export (opts.showAssessesTag); the SC target still lives in the notes.
    if (o.assessesSc && o.showAssessesTag) {
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
    // Density-aware: with 1-2 prompts, grow the cards and the prompt font
    // (the exit task is the hero), and centre the block vertically so the
    // slide never ends with a dead bottom third (megaprompt 15h).
    const availH = SAFE_BOTTOM - startY;
    const perH = Math.min(
      byBand(sz, 2.1, 1.8, cleaned.length <= 2 ? 1.55 : 1.2),
      availH / Math.max(cleaned.length, 1) - 0.12,
    );
    const promptFont = cleaned.length <= 2 ? byBand(sz, 30, 26, 22) : sz.body;
    const blockH = cleaned.length * perH + (cleaned.length - 1) * 0.12;
    const blockOffset = Math.max(0, (availH - blockH) * 0.35);
    const numbered = Boolean(o.numbered);

    cleaned.forEach((q, i) => {
      const qY = startY + blockOffset + i * (perH + 0.12);
      el.addCard(s, 0.5, qY, 9, perH, { variant: "tint", tone: stripColor });
      const display = numbered ? `${i + 1}.  ${q}` : q;
      s.addText(display, {
        x: 0.75, y: qY + 0.08, w: 8.5, h: perH - 0.16,
        fontSize: promptFont, fontFace: FONT_B, color: C.CHARCOAL, margin: 0, valign: "middle",
        fit: "shrink", shrinkText: true,
      });
    });

    if (footer) el.addFooter(s, footer);
    if (notes) s.addNotes(notes);
    return s;
  }

  /**
   * keyWordSlide - universal visual word card (megaprompt section 29).
   *
   * ONE key word per slide: hero word panel (the word is the largest thing
   * on the slide), a student-friendly meaning, an optional example, and an
   * oral routine chip row. Use one call per word - never render vocabulary
   * as a definition bullet list.
   *
   *   keyWordSlide(pres, { word, meaning, example, routine, color, badgeText, title, pictogram, image }, notes, footer)
   *
   * - word       the key word (required)
   * - meaning    student-friendly meaning (required)
   * - example    optional example sentence or "hear it" line
   * - routine    optional array of routine chips (default ["Say it", "Act it", "Use it"])
   * - pictogram  built-in picture for the word (see listPictograms()), drawn
   *              large under the word so the card carries a graphic
   *              (megaprompt 29: one large word, one meaningful graphic)
   * - image      local image path, used instead of a pictogram when a real
   *              picture is the point (photo, illustration from the text)
   */
  function keyWordSlide(pres, config, notes, footer) {
    const cfg = (config && typeof config === "object") ? config : { word: String(config || "") };
    const word = String(cfg.word || "").trim();
    const meaning = String(cfg.meaning || "").trim();
    const example = cfg.example ? String(cfg.example).trim() : "";
    const routineItems = (Array.isArray(cfg.routine) ? cfg.routine : ["Say it", "Act it", "Use it"])
      .map((r) => String(r).trim()).filter(Boolean).slice(0, 4);
    const accent = cfg.color || C.PRIMARY;

    const s = pres.addSlide();
    el.addTopBar(s, accent);
    el.addBadge(s, cfg.badgeText || "Key Word", { color: accent, w: byBand(sz, 2.0, 1.85, 1.7) });
    el.addTitle(s, cfg.title || "Words for today");

    // Left: hero word panel - the word is the hero
    const panelW = 4.2;
    const panelH = SAFE_BOTTOM - CONTENT_TOP;
    s.addShape("roundRect", {
      x: 0.5, y: CONTENT_TOP, w: panelW, h: panelH, rectRadius: 0.1,
      fill: { color: accent },
    });
    // Pre-compute a font size that keeps the word on ONE line - shrink-fit
    // does not prevent ugly mid-word wrapping for long words.
    const wordFont = Math.max(20, Math.min(
      byBand(sz, 54, 48, 40),
      Math.floor((panelW - 0.3) / (Math.max(word.length, 4) * 0.0118))
    ));
    const sayW = byBand(sz, 2.5, 2.3, 2.1);
    const sayH = byBand(sz, 0.5, 0.46, 0.42);
    const hasPicto = Boolean(cfg.pictogram && picto);
    const hasImage = Boolean(cfg.image);
    const hasGraphic = hasPicto || hasImage;
    if (!hasGraphic) {
      console.log(`ADVISORY keyWordSlide "${word}" has no pictogram or image - megaprompt 29 wants one meaningful graphic per word card. Pass { pictogram: "<name>" } (see listPictograms()) or { image: path }.`);
    }
    // With a graphic the word takes the top third and the picture the middle.
    const wordH = hasGraphic ? byBand(sz, 1.1, 1.0, 0.95) : panelH - 1.05;
    s.addText(word, {
      x: 0.6, y: CONTENT_TOP + 0.15, w: panelW - 0.2, h: wordH,
      fontSize: wordFont, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
    if (hasGraphic) {
      const gTop = CONTENT_TOP + 0.15 + wordH + 0.05;
      const gBottom = CONTENT_TOP + panelH - sayH - 0.3;
      const gH = Math.max(0.8, gBottom - gTop);
      if (hasImage && visual) {
        visual.drawVisual(s, { type: "image", path: cfg.image, frame: true }, {
          x: 0.75, y: gTop, w: panelW - 0.5, h: gH,
        });
      } else if (hasPicto) {
        const d = Math.min(gH, panelW - 1.2);
        picto.addPictogram(s, cfg.pictogram, 0.5 + (panelW - d) / 2, gTop + (gH - d) / 2, d, {
          style: "flat", color: C.WHITE, glyphColor: C.WHITE,
        });
      }
    }
    el.addTextOnShape(s, "Say it with me", {
      x: 0.5 + (panelW - sayW) / 2, y: CONTENT_TOP + panelH - sayH - 0.18,
      w: sayW, h: sayH, rectRadius: 0.08,
      fill: { color: C.WHITE },
    }, {
      fontSize: byBand(sz, 16, 15, 13), fontFace: FONT_B, color: accent,
      bold: true, align: "center", valign: "middle", margin: 0,
    });

    // Right: meaning, optional example, routine chips
    const rx = 5.0;
    const rw = 4.5;
    const chipRowH = byBand(sz, 0.58, 0.52, 0.48);
    const chipRowY = SAFE_BOTTOM - chipRowH;
    const meaningH = example
      ? (chipRowY - CONTENT_TOP - 0.2) * 0.62
      : chipRowY - CONTENT_TOP - 0.2;

    el.addCard(s, rx, CONTENT_TOP, rw, meaningH, { strip: accent, fill: C.WHITE });
    s.addText("It means", {
      x: rx + 0.22, y: CONTENT_TOP + 0.1, w: rw - 0.44, h: 0.26,
      fontSize: byBand(sz, 14, 13, 12), fontFace: FONT_B, color: accent, bold: true, margin: 0,
    });
    s.addText(meaning, {
      x: rx + 0.22, y: CONTENT_TOP + 0.4, w: rw - 0.44, h: meaningH - 0.55,
      fontSize: byBand(sz, 26, 23, 20), fontFace: FONT_B, color: C.CHARCOAL,
      valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
    });

    if (example) {
      const exY = CONTENT_TOP + meaningH + 0.14;
      const exH = chipRowY - exY - 0.14;
      el.addCard(s, rx, exY, rw, exH, { strip: C.SECONDARY, fill: C.WHITE });
      s.addText(example, {
        x: rx + 0.22, y: exY + 0.08, w: rw - 0.44, h: exH - 0.16,
        fontSize: byBand(sz, 18, 16, 14.5), fontFace: FONT_B, color: C.CHARCOAL,
        italic: true, valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
    }

    if (routineItems.length) {
      const chipGap = 0.12;
      const chipW = (rw - chipGap * (routineItems.length - 1)) / routineItems.length;
      routineItems.forEach((label, i) => {
        el.addTextOnShape(s, label, {
          x: rx + i * (chipW + chipGap), y: chipRowY, w: chipW, h: chipRowH,
          rectRadius: 0.08, fill: { color: C.WHITE }, line: { color: accent, width: 1.3 },
        }, {
          fontSize: byBand(sz, 15, 14, 12.5), fontFace: FONT_B, color: accent,
          bold: true, align: "center", valign: "middle", margin: 0,
        });
      });
    }

    if (footer) el.addFooter(s, footer);
    if (notes) s.addNotes(notes);
    runSlideDiagnostics(s, pres);
    return s;
  }

  /**
   * heroVisualSlide - visual-only teaching slide (megaprompt 0a item 20, 15c).
   *
   * The representation IS the slide. The visual spec is fitted to fill the
   * content area on a soft panel and centred, with at most a short label
   * naming the model and an optional one-line prompt bar. Everything the
   * teacher says lives in the notes.
   *
   *   heroVisualSlide(pres, "I Do", "How many counters?",
   *     { type: "tensFrame", filled: 7 }, notes, footer,
   *     { label: "Tens frame", prompt: "Show me on your fingers" })
   *
   * opts: { label, prompt, badgeColor, panel (default true), promptColor }
   */
  function heroVisualSlide(pres, badgeText, title, visualSpec, notes, footer, opts) {
    const s = pres.addSlide();
    const o = opts || {};
    const tone = o.badgeColor || C.PRIMARY;
    el.addTopBar(s, tone);
    el.addBadge(s, badgeText || "Look", { color: tone });
    el.addTitle(s, title);

    const promptH = o.prompt ? byBand(sz, 0.8, 0.72, 0.62) : 0;
    const labelH = o.label ? byBand(sz, 0.5, 0.44, 0.38) : 0;
    const gap = 0.16;
    const panelTop = CONTENT_TOP;
    const panelBottom = SAFE_BOTTOM - (promptH ? promptH + gap : 0);
    const panelH = panelBottom - panelTop;
    const usePanel = o.panel !== false;
    if (usePanel) {
      el.addCard(s, 0.5, panelTop, 9, panelH, { variant: "tint", tone });
    }
    const pad = usePanel ? 0.3 : 0.05;
    const frame = {
      x: 0.5 + pad, y: panelTop + pad, w: 9 - pad * 2,
      h: panelH - pad * 2 - (labelH ? labelH + 0.08 : 0),
    };
    if (!visual) throw new Error("[heroVisualSlide] theme has no visual layer");
    const bounds = visual.drawVisual(s, visualSpec, frame);

    if (o.label) {
      s.addText(String(o.label), {
        x: 0.5, y: panelBottom - pad - labelH + 0.04, w: 9, h: labelH,
        fontSize: byBand(sz, 22, 19, 15), fontFace: FONT_B, color: C.MUTED, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
    }
    if (o.prompt) {
      el.addTextOnShape(s, String(o.prompt), {
        x: 0.5, y: SAFE_BOTTOM - promptH, w: 9, h: promptH, rectRadius: promptH / 2,
        fill: { color: o.promptColor || tone },
      }, {
        fontSize: byBand(sz, 26, 22, 18), fontFace: FONT_H, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0.1,
      });
    }

    if (footer) el.addFooter(s, footer);
    if (notes) s.addNotes(notes);
    runSlideDiagnostics(s, pres, { ignoreUnderfill: true });
    s.heroBounds = bounds;
    return s;
  }

  /**
   * choiceSlide - "Which one?" visual choice (megaprompt 15d, 49, 68j).
   *
   * Two to four large option cards side by side, each holding a visual spec
   * and/or short text, lettered A/B/C by default so students can answer on
   * boards or fingers without anything being numbered. Use for Which one
   * matches?, Same or different?, Example and non-example, odd one out, and
   * A/B/C hinge questions.
   *
   *   const s = choiceSlide(pres, "CFU", "Which shows three quarters?",
   *     "Show me A, B or C on your board", [
   *       { visual: { type: "fractionStrips", strips: [{ denom: 4, shaded: 3 }] } },
   *       { visual: { type: "fractionStrips", strips: [{ denom: 3, shaded: 2 }] } },
   *       { text: "3 out of 5" },
   *     ], notes, footer, { badgeColor: C.ALERT });
   *   clickBuild(s, [() => markChoice(s, 0)]);   // reveal the answer on click
   *
   * option: { visual, text, label, caption }   (label overrides the letter)
   * opts:   { badgeColor, letters (default true), promptColor, captionSize }
   */
  function choiceSlide(pres, badgeText, title, prompt, options, notes, footer, opts) {
    const s = pres.addSlide();
    const o = opts || {};
    const tone = o.badgeColor || C.PRIMARY;
    const list = (Array.isArray(options) ? options : []).slice(0, 4);
    el.addTopBar(s, tone);
    el.addBadge(s, badgeText || "Which one?", { color: tone });
    el.addTitle(s, title);

    const promptH = prompt ? byBand(sz, 0.74, 0.66, 0.56) : 0;
    let cursorY = CONTENT_TOP;
    if (prompt) {
      el.addTextOnShape(s, String(prompt), {
        x: 0.5, y: cursorY, w: 9, h: promptH, rectRadius: promptH / 2,
        fill: { color: o.promptColor || tone },
      }, {
        fontSize: byBand(sz, 24, 20, 16), fontFace: FONT_B, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0.1,
      });
      cursorY += promptH + 0.18;
    }

    const n = Math.max(list.length, 1);
    const gap = 0.25;
    const cardW = (9 - gap * (n - 1)) / n;
    const cardH = SAFE_BOTTOM - cursorY;
    const letters = ["A", "B", "C", "D"];
    const tag = byBand(sz, 0.62, 0.56, 0.48);
    const frames = [];

    list.forEach((raw, i) => {
      const opt = typeof raw === "string" ? { text: raw } : (raw || {});
      const x = 0.5 + i * (cardW + gap);
      el.addCard(s, x, cursorY, cardW, cardH, { variant: "outline", tone });
      const showTag = o.letters !== false || opt.label;
      if (showTag) {
        el.addTextOnShape(s, String(opt.label || letters[i]), {
          x: x + 0.14, y: cursorY + 0.14, w: tag, h: tag, rectRadius: tag / 2,
          fill: { color: tone },
        }, {
          fontSize: byBand(sz, 22, 19, 15), fontFace: FONT_H, color: C.WHITE, bold: true,
        });
      }
      const innerTop = cursorY + (showTag ? tag + 0.26 : 0.2);
      const captionH = opt.caption ? byBand(sz, 0.5, 0.44, 0.38) : 0;
      const textH = opt.text && opt.visual ? byBand(sz, 0.9, 0.8, 0.7) : 0;
      const innerBottom = cursorY + cardH - 0.2 - captionH - textH;
      const frame = { x: x + 0.2, y: innerTop, w: cardW - 0.4, h: Math.max(0.6, innerBottom - innerTop) };

      if (opt.visual && visual) {
        visual.drawVisual(s, opt.visual, frame);
      }
      if (opt.text) {
        const textOnly = !opt.visual;
        const ty = textOnly ? innerTop : innerBottom;
        const th = textOnly ? Math.max(0.6, cursorY + cardH - 0.2 - captionH - innerTop) : textH;
        s.addText(String(opt.text), {
          x: x + 0.2, y: ty, w: cardW - 0.4, h: th,
          fontSize: textOnly ? byBand(sz, 40, 34, 28) : byBand(sz, 22, 19, 16),
          fontFace: textOnly ? FONT_H : FONT_B, color: C.CHARCOAL, bold: textOnly,
          align: "center", valign: "middle", margin: 0,
          fit: "shrink", shrinkText: true,
        });
      }
      if (opt.caption) {
        s.addText(String(opt.caption), {
          x: x + 0.15, y: cursorY + cardH - 0.14 - captionH, w: cardW - 0.3, h: captionH,
          fontSize: o.captionSize || byBand(sz, 18, 16, 13), fontFace: FONT_B, color: C.MUTED,
          align: "center", valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
        });
      }
      frames.push({ x, y: cursorY, w: cardW, h: cardH });
    });

    s.choiceFrames = frames;
    if (footer) el.addFooter(s, footer);
    if (notes) s.addNotes(notes);
    runSlideDiagnostics(s, pres);
    return s;
  }

  /**
   * markChoice - reveal the correct option on a choiceSlide: a thick SUCCESS
   * border plus a tick badge. Call inside a clickBuild step or a withReveal
   * revealFn. `index` is 0-based.
   */
  function markChoice(slide, index, opts) {
    const o = opts || {};
    const frames = (slide && slide.choiceFrames) || [];
    const f = frames[index];
    if (!f) {
      console.warn(`WARN [markChoice] no option ${index} on this slide (${frames.length} options)`);
      return;
    }
    const color = o.color || C.SUCCESS;
    slide.addShape("roundRect", {
      x: f.x - 0.04, y: f.y - 0.04, w: f.w + 0.08, h: f.h + 0.08, rectRadius: 0.14,
      fill: { color: color, transparency: 100 },
      line: { color, width: 4 },
    });
    const d = byBand(sz, 0.7, 0.62, 0.54);
    if (picto) {
      picto.addPictogram(slide, "tick", f.x + f.w - d - 0.1, f.y - d * 0.35, d, { style: "flat", color });
    } else {
      el.addTextOnShape(slide, "Yes", {
        x: f.x + f.w - d - 0.1, y: f.y - d * 0.35, w: d, h: d, rectRadius: d / 2, fill: { color },
      }, { fontSize: byBand(sz, 14, 12, 10), fontFace: FONT_B, color: C.WHITE, bold: true });
    }
  }

  /**
   * youDoSlide - independent task slide (megaprompt 35).
   *
   * The task is the hero. First / Next / Then steps are small numbered chips
   * beneath it (max 3, one action each), an optional visual spec sits on the
   * right as a mini model, and an optional sentence frame or reminder line
   * sits under the task in a dashed pill.
   *
   *   youDoSlide(pres, "Make 10 on your own", "Draw counters to make 10.",
   *     ["Look at the frame", "Draw the counters", "Write how many more"],
   *     notes, footer, {
   *       where: "On your worksheet",
   *       visual: { type: "tensFrame", filled: 6 },
   *       frame: "___ and ___ make 10",
   *     })
   *
   * opts: { where, visual, frame, badgeText, badgeColor, visualLabel }
   */
  function youDoSlide(pres, title, task, steps, notes, footer, opts) {
    const s = pres.addSlide();
    const o = opts || {};
    const isNumeracy = subjectKey === "numeracy";
    const tone = o.badgeColor || (isNumeracy ? C.ALERT : C.SUCCESS);
    const badgeText = o.badgeText || (isNumeracy ? "Stage 4  |  You Do" : "You Do");
    el.addTopBar(s, tone);
    el.addBadge(s, badgeText, { color: tone, w: isNumeracy ? byBand(sz, 3.1, 2.8, 2.4) : undefined });
    el.addTitle(s, title);

    const stepList = (Array.isArray(steps) ? steps : (steps ? [steps] : []))
      .map((t) => String(t == null ? "" : t).trim()).filter(Boolean).slice(0, 3);
    const hasVisual = isSpec(o.visual) && visual;
    const stepsH = stepList.length ? byBand(sz, 1.0, 0.9, 0.8) : 0;
    const frameH = o.frame ? byBand(sz, 0.66, 0.6, 0.52) : 0;
    const gap = 0.16;
    const leftW = hasVisual ? 4.6 : 9;
    const areaH = SAFE_BOTTOM - CONTENT_TOP;
    const taskCardH = areaH - (stepsH ? stepsH + gap : 0) - (frameH ? frameH + gap : 0);

    // Task card (hero)
    el.addCard(s, 0.5, CONTENT_TOP, leftW, taskCardH, { variant: "tint", tone });
    let taskTop = CONTENT_TOP + 0.2;
    if (o.where) {
      const whereH = byBand(sz, 0.42, 0.38, 0.34);
      const whereW = Math.min(leftW - 0.4, Math.max(2.0, String(o.where).length * 0.11 + 0.5));
      el.addTextOnShape(s, String(o.where), {
        x: 0.7, y: taskTop, w: whereW, h: whereH, rectRadius: whereH / 2, fill: { color: tone },
      }, { fontSize: byBand(sz, 15, 14, 12), fontFace: FONT_B, color: C.WHITE, bold: true });
      taskTop += whereH + 0.12;
    }
    const taskText = String(task == null ? "" : task);
    const taskAvailH = CONTENT_TOP + taskCardH - 0.2 - taskTop;
    const charsPerLine = hasVisual ? byBand(sz, 18, 22, 28) : byBand(sz, 32, 40, 50);
    const taskFont = fitTextFontSize(taskText, taskAvailH, charsPerLine, hasVisual ? sz.hero * 0.9 : sz.heroQuestion, sz.body);
    s.addText(taskText, {
      x: 0.75, y: taskTop, w: leftW - 0.5, h: taskAvailH,
      fontSize: taskFont, fontFace: FONT_B, color: C.CHARCOAL, bold: true,
      valign: "middle", align: "left", margin: 0, fit: "shrink", shrinkText: true,
    });

    // Optional mini model on the right
    if (hasVisual) {
      const vx = 0.5 + leftW + 0.2;
      const vw = 9.5 - vx;
      const labelH = o.visualLabel ? byBand(sz, 0.44, 0.4, 0.34) : 0;
      el.addCard(s, vx, CONTENT_TOP, vw, taskCardH, { variant: "outline", tone });
      visual.drawVisual(s, o.visual, { x: vx + 0.25, y: CONTENT_TOP + 0.25, w: vw - 0.5, h: taskCardH - 0.5 - labelH });
      if (o.visualLabel) {
        s.addText(String(o.visualLabel), {
          x: vx, y: CONTENT_TOP + taskCardH - 0.18 - labelH, w: vw, h: labelH,
          fontSize: byBand(sz, 16, 14, 12), fontFace: FONT_B, color: C.MUTED, bold: true,
          align: "center", valign: "middle", margin: 0,
        });
      }
    }

    let cursor = CONTENT_TOP + taskCardH + gap;
    if (o.frame) {
      s.addShape("roundRect", {
        x: 0.5, y: cursor, w: 9, h: frameH, rectRadius: frameH / 2,
        fill: { color: C.WHITE }, line: { color: tone, width: 1.2, dashType: "dash" },
      });
      s.addText(String(o.frame), {
        x: 0.7, y: cursor, w: 8.6, h: frameH,
        fontSize: byBand(sz, 24, 20, 17), fontFace: FONT_B, color: C.CHARCOAL, italic: true,
        align: "center", valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
      cursor += frameH + gap;
    }

    if (stepList.length) {
      const words = ["First", "Next", "Then"];
      const sgap = 0.2;
      const stepW = (9 - sgap * (stepList.length - 1)) / stepList.length;
      const numD = byBand(sz, 0.5, 0.46, 0.4);
      stepList.forEach((text, i) => {
        const x = 0.5 + i * (stepW + sgap);
        el.addCard(s, x, cursor, stepW, stepsH, { variant: "outline", tone });
        el.addTextOnShape(s, String(i + 1), {
          x: x + 0.16, y: cursor + (stepsH - numD) / 2, w: numD, h: numD, rectRadius: numD / 2, fill: { color: tone },
        }, { fontSize: byBand(sz, 18, 16, 13), fontFace: FONT_H, color: C.WHITE, bold: true });
        s.addText([
          { text: words[i] + ": ", options: { bold: true, color: tone } },
          { text, options: { color: C.CHARCOAL } },
        ], {
          x: x + 0.16 + numD + 0.12, y: cursor + 0.08, w: stepW - numD - 0.4, h: stepsH - 0.16,
          fontSize: byBand(sz, 18, 16, 14), fontFace: FONT_B, valign: "middle", margin: 0,
          fit: "shrink", shrinkText: true,
        });
      });
    }

    if (footer) el.addFooter(s, footer);
    if (notes) s.addNotes(notes);
    runSlideDiagnostics(s, pres);
    return s;
  }

  /**
   * textExtractSlide - literacy reading anchor.
   *
   * A large calm reading panel with the extract set in a reading face, with
   * optional highlighted phrases (bold, coloured, marker wash) for
   * find-the-evidence and annotate-the-model work, an optional source line,
   * and an optional prompt bar. Quote text is rendered exactly as passed
   * (megaprompt 5a).
   *
   *   textExtractSlide(pres, "Read", "The storm", extractText, notes, footer, {
   *     highlights: ["gnarled hands", "crept"],
   *     source: "War Horse, chapter 3",
   *     prompt: "Which word shows Joey is afraid?",
   *   })
   *
   * opts: { highlights, source, prompt, badgeColor, fontFace, fontSize }
   */
  function textExtractSlide(pres, badgeText, title, extract, notes, footer, opts) {
    const s = pres.addSlide();
    const o = opts || {};
    const tone = o.badgeColor || C.PRIMARY;
    el.addTopBar(s, tone);
    el.addBadge(s, badgeText || "Read", { color: tone });
    el.addTitle(s, title);

    const promptH = o.prompt ? byBand(sz, 0.78, 0.7, 0.6) : 0;
    const sourceH = o.source ? 0.3 : 0;
    const gap = 0.16;
    const panelH = SAFE_BOTTOM - CONTENT_TOP - (promptH ? promptH + gap : 0) - (sourceH ? sourceH + 0.04 : 0);
    el.addCard(s, 0.5, CONTENT_TOP, 9, panelH, { variant: "outline", tone });
    s.addShape("rect", { x: 0.5, y: CONTENT_TOP, w: 0.09, h: panelH, fill: { color: tone } });

    const text = String(extract == null ? "" : extract);
    const readingFace = o.fontFace || (sz._band === "Y36" ? "Georgia" : FONT_B);
    const charsPerLine = byBand(sz, 34, 42, 58);
    const fontSize = o.fontSize || fitTextFontSize(text, panelH - 0.5, charsPerLine, byBand(sz, 30, 26, 22), byBand(sz, 20, 18, 14));

    // Split into runs so highlighted phrases render bold, coloured and washed.
    const highlights = (Array.isArray(o.highlights) ? o.highlights : []).map((h) => String(h)).filter(Boolean);
    const runs = [];
    if (highlights.length) {
      const pattern = new RegExp(`(${highlights.map((h) => h.replace(/[.*+?^${}()|[\\]\\\\]/g, "\\\\$&")).join("|")})`, "g");
      text.split(pattern).forEach((part) => {
        if (!part) return;
        const hit = highlights.includes(part);
        runs.push({
          text: part,
          options: hit
            ? { bold: true, color: tone, highlight: C.ACCENT_SOFT || el.softOf(C.ACCENT, 0.75) }
            : { color: C.CHARCOAL },
        });
      });
    } else {
      runs.push({ text, options: { color: C.CHARCOAL } });
    }
    s.addText(runs, {
      x: 0.85, y: CONTENT_TOP + 0.22, w: 8.4, h: panelH - 0.44,
      fontSize, fontFace: readingFace, valign: "middle", margin: 0,
      lineSpacingMultiple: 1.15,
      fit: "shrink", shrinkText: true,
    });

    let cursor = CONTENT_TOP + panelH;
    if (o.source) {
      s.addText(String(o.source), {
        x: 0.5, y: cursor + 0.02, w: 9, h: sourceH,
        fontSize: byBand(sz, 13, 12, 11), fontFace: FONT_B, color: C.MUTED, italic: true,
        align: "right", valign: "middle", margin: 0,
      });
      cursor += sourceH + 0.04;
    }
    if (o.prompt) {
      el.addTextOnShape(s, String(o.prompt), {
        x: 0.5, y: SAFE_BOTTOM - promptH, w: 9, h: promptH, rectRadius: promptH / 2,
        fill: { color: o.promptColor || tone },
      }, {
        fontSize: byBand(sz, 24, 20, 17), fontFace: FONT_B, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0.1,
      });
    }

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
    keyWordSlide,
    annotatedModelSlide,
    compareVisualSlide,
    boardBuildSlide,
    exitTicketSlide,
    addRevealAnswerBar,
    heroVisualSlide,
    choiceSlide,
    markChoice,
    youDoSlide,
    textExtractSlide,
  };
}

module.exports = { createBaseBuilders };
