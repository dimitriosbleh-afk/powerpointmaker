"use strict";

const { SAFE_BOTTOM, CONTENT_TOP, SLIDE_W, validateBounds } = require("../core/layout");
const { runSlideDiagnostics } = require("../core/diagnostics");
const { DEFAULT_SIZES, byBand } = require("../core/gradeBand");
const { createManipulatives } = require("../core/manipulatives");
const { prepareBullets, fitBulletFontSize } = require("../core/bulletFit");

/**
 * Factory that returns numeracy-specific slide builders and maths visual
 * helpers bound to a given palette, fonts, and element helpers.
 *
 * @param {object} C       Semantic palette colours (PRIMARY, SECONDARY, ACCENT, ALERT, SUCCESS, ASSESS, BG_DARK, BG_LIGHT, BG_CARD, WHITE, CHARCOAL, MUTED, TEXT_ON_DARK, SUBTITLE, DECOR_1, DECOR_2)
 * @param {string} FONT_H  Heading font name
 * @param {string} FONT_B  Body font name
 * @param {object} el       Bound element helpers: addTopBar, addBadge, addTitle, addCard, addFooter, addIconCircle, addTextOnShape
 * @returns {object}        { STAGE_COLORS, addStageBadge, workedExSlide, dailyReviewSlide, fluencySlide, addPlaceValueChart, addTenthsStrip, addAreaModel, addDecimalDot } plus every shared manipulative (addNumberLine lives in core/manipulatives.js)
 */
function createNumeracyBuilders(C, FONT_H, FONT_B, el, S, defaults) {
  const sz = S || DEFAULT_SIZES;
  const manipulatives = createManipulatives(C, FONT_B, sz);
  const visual = (defaults && defaults.visual) || null;
  const isSpec = (value) => Boolean(visual && visual.isVisualSpec && visual.isVisualSpec(value));

  /** Accept a callback or a visual spec for the right column (see base.js). */
  function resolveDrawRight(drawRight) {
    if (typeof drawRight === "function") return drawRight;
    if (isSpec(drawRight)) {
      return (slide, guide) => visual.drawVisual(slide, drawRight, {
        x: guide.rightX, y: guide.panelTop + 0.05,
        w: guide.rightW, h: guide.safeBottom - guide.panelTop - 0.1,
      });
    }
    return null;
  }

  /* ------------------------------------------------------------------ */
  /*  STAGE_COLORS                                                       */
  /* ------------------------------------------------------------------ */

  const STAGE_COLORS = {
    "1": C.ACCENT,                   // Daily Review / Fluency
    "2": C.PRIMARY,                  // I Do
    "3": C.SECONDARY,               // We Do
    "4": C.ALERT,                   // You Do
    "5": C.ASSESS || C.ALERT,       // Exit Ticket
  };

  const STAGE_LABELS = {
    "1": "Daily Review",
    "2": "I Do",
    "3": "We Do",
    "4": "You Do",
    "5": "Exit Ticket",
  };

  /* ------------------------------------------------------------------ */
  /*  addStageBadge                                                      */
  /* ------------------------------------------------------------------ */

  /**
   * Stage-coloured badge positioned at the top-left of the slide.
   *
   * @param {object} slide     PptxGenJS slide object
   * @param {number} stageNum  Stage number (1-5)
   * @param {string} label     Stage label text
   */
  function addStageBadge(slide, stageNum, label) {
    const resolvedLabel = label || STAGE_LABELS[String(stageNum)] || "Stage " + stageNum;
    const color = STAGE_COLORS[String(stageNum)] || C.PRIMARY;
    // Wider/taller badge for F/Y12 to fit larger badge text without clipping.
    const baseW = resolvedLabel.length > 20 ? 3.2 : 2.4;
    const w = byBand(sz, baseW + 0.7, baseW + 0.4, baseW);
    const h = byBand(sz, 0.42, 0.40, 0.36);
    slide.addShape("roundRect", {
      x: 0.5, y: 0.2, w, h, rectRadius: h / 2,
      fill: { color },
    });
    slide.addText("Stage " + stageNum + "  |  " + resolvedLabel, {
      x: 0.5, y: 0.2, w, h,
      fontSize: sz.badge, fontFace: FONT_B, color: C.WHITE,
      align: "center", valign: "middle", bold: true, margin: 0,
    });
  }

  /* ------------------------------------------------------------------ */
  /*  workedExSlide                                                      */
  /* ------------------------------------------------------------------ */

  /**
   * Worked example slide with stage badge, steps on left card, and an
   * optional right-side callback for visual content.
   *
   * @param {object}   pres        PptxGenJS presentation instance
   * @param {number}   stageNum    Stage number (1-5)
   * @param {string}   stageLabel  Stage label text
   * @param {string}   title       Slide title
   * @param {string[]} steps       Bullet-point steps for the worked example
   * @param {string}   notes       Teacher notes
   * @param {string}   footer      Footer text
   * @param {Function|object} drawRight   Optional callback(slide, layoutGuide) for right-column
   *                                      visuals, OR a visual spec such as
   *                                      { type: "fractionStrips", strips: [...] } which is
   *                                      fitted into the right column automatically
   * @returns {object}             The slide object
   */
  function workedExSlide(pres, stageNum, stageLabel, title, steps, notes, footer, drawRightArg) {
    const s = pres.addSlide();
    const drawRight = resolveDrawRight(drawRightArg);
    const stageColor = STAGE_COLORS[String(stageNum)] || C.PRIMARY;
    el.addTopBar(s, stageColor);
    addStageBadge(s, stageNum, stageLabel);
    // Slightly smaller than the slide H1 so the stage badge above it stays
    // visually dominant; still scales by band.
    el.addTitle(s, title, { fontSize: sz.titleH1 - 2, color: stageColor });

    const cardW = drawRight ? 4.5 : 9;
    const contentY = CONTENT_TOP;
    const layoutGuide = {
      titleY: 0.65,
      titleH: 0.62,
      panelTop: contentY,
      panelTopPadded: contentY + 0.08,
      leftCardX: 0.5,
      leftCardY: contentY,
      leftCardW: cardW,
      leftCardH: SAFE_BOTTOM - contentY,
      rightX: 5.3,
      rightW: 4.2,
      safeBottom: SAFE_BOTTOM,
    };

    el.addCard(s, 0.5, contentY, cardW, SAFE_BOTTOM - contentY, { strip: stageColor });

    // Filter empty-string spacers from build-script step lists; they
    // become paraSpaceAfter boosts on the preceding bullet so the slide
    // gets clean grouping instead of empty bullet markers.
    const prepared = prepareBullets(steps);
    // Available text-frame height inside the step card.
    const textY = contentY + 0.14;
    const textH = SAFE_BOTTOM - contentY - 0.24;
    // Ideal step font size; narrow column drops one step. fitBulletFontSize
    // shrinks deterministically when content would overflow the card —
    // PptxGenJS shrinkText on bullet lists is unreliable, so we pre-compute.
    // Few short steps get hero-ish type: an I Do with three lines beside a
    // model should not read as small print (megaprompt 15b / 16).
    const stepCount = prepared.length;
    const longest = prepared.reduce((m, p) => Math.max(m, p.text.length), 0);
    const sparseSteps = stepCount <= 3 && longest <= (drawRight ? 34 : 60);
    const idealStepFontSize = sparseSteps
      ? (drawRight ? sz.body : Math.round(sz.body * 1.2))
      : (drawRight ? sz.bodyDense : sz.body);
    const fontFloor = byBand(sz, 14, 13, 11);
    const charsPerLine = drawRight ? byBand(sz, 28, 32, 36) : byBand(sz, 42, 50, 56);
    const stepFontSize = fitBulletFontSize(prepared, textH, charsPerLine, idealStepFontSize, fontFloor);
    const baseSpacePt = stepFontSize >= 16 ? 5 : 3;
    const stepTexts = prepared.map((item, i) => ({
      text: item.text,
      options: {
        bullet: true,
        breakLine: i < prepared.length - 1,
        fontSize: stepFontSize,
        color: C.CHARCOAL,
        paraSpaceAfter: baseSpacePt + (item.extraSpaceAfter || 0) * stepFontSize * 0.9,
      },
    }));
    s.addText(stepTexts, {
      x: 0.75, y: textY, w: cardW - 0.4, h: textH,
      fontFace: FONT_B, margin: 0, valign: "top",
    });

    if (drawRight) drawRight(s, layoutGuide);
    if (footer) el.addFooter(s, footer);
    if (notes) s.addNotes(notes);
    if (drawRight) runSlideDiagnostics(s, pres);
    return s;
  }

  /* ================================================================== */
  /*  Maths Visual Helpers                                               */
  /* ================================================================== */

  /* ------------------------------------------------------------------ */
  /*  addPlaceValueChart                                                 */
  /* ------------------------------------------------------------------ */

  /**
   * Place value chart with auto-sizing. Draws header cells (filled with
   * PRIMARY) and value cells (white with PRIMARY border). Auto-scales
   * font sizes for narrow cells. Inserts a thousands-gap marker using
   * ACCENT when ten-thousands header is present.
   *
   * @param {object}   slide    PptxGenJS slide object
   * @param {number}   x        Left edge x (inches)
   * @param {number}   y        Top edge y (inches)
   * @param {string[]} headers  Column header labels
   * @param {Array}    values   Values to display (may contain nulls for empty cells)
   * @param {object}   opts     Options: totalW, w, cellW, hdrH, valH, headerColor
   * @returns {object}          Geometry: { cellW, totalW, hdrH, valH, n, x, y }
   */
  function addPlaceValueChart(slide, x, y, headers, values, opts) {
    const o = opts || {};
    const n = headers.length;

    if (n === 0) {
      console.warn("[addPlaceValueChart] empty headers array — skipping");
      return { cellW: 0, totalW: 0, hdrH: 0, valH: 0, n: 0, x, y };
    }

    let cellW;
    if (o.totalW != null) { cellW = o.totalW / n; }
    else if (o.w != null) { cellW = o.w / n; }
    else { cellW = o.cellW || 1.2; }

    const hdrH = o.hdrH || 0.52;
    const valH = o.valH || 0.7;
    const totalW = cellW * n;
    const headerColor = o.headerColor || C.PRIMARY;

    validateBounds("addPlaceValueChart", x, y, totalW, hdrH + valH);

    const hasTenThousands = headers.some(h => h.toLowerCase().includes("ten thousand"));
    const hdrFontSize = cellW < 0.7 ? 7 : cellW < 0.9 ? 8 : 9;
    const valFontSize = cellW < 0.7 ? 18 : cellW < 0.9 ? 22 : 26;

    headers.forEach((h, i) => {
      const cx = x + i * cellW;

      // Header cell
      slide.addShape("rect", {
        x: cx, y, w: cellW, h: hdrH,
        fill: { color: headerColor },
        line: { color: C.WHITE, width: 1 },
      });
      slide.addText(h, {
        x: cx, y, w: cellW, h: hdrH,
        fontSize: hdrFontSize, fontFace: FONT_B, color: C.WHITE,
        align: "center", valign: "middle", bold: true, margin: 0,
      });

      // Value cell
      slide.addShape("rect", {
        x: cx, y: y + hdrH, w: cellW, h: valH,
        fill: { color: C.WHITE },
        line: { color: headerColor, width: 1 },
      });

      const val = values && values[i] != null ? String(values[i]) : "";
      if (val !== "") {
        slide.addText(val, {
          x: cx, y: y + hdrH, w: cellW, h: valH,
          fontSize: valFontSize, fontFace: FONT_H, color: C.CHARCOAL,
          align: "center", valign: "middle", bold: true, margin: 0,
        });
      }
    });

    // Thousands gap marker
    if (hasTenThousands && n >= 5) {
      const gapX = x + 2 * cellW;
      slide.addShape("rect", {
        x: gapX - 0.03, y: y + hdrH, w: 0.06, h: valH,
        fill: { color: C.ACCENT },
      });
    }

    return { cellW, totalW, hdrH, valH, n, x, y };
  }

  /* ------------------------------------------------------------------ */
  /*  addTenthsStrip                                                     */
  /* ------------------------------------------------------------------ */

  /**
   * Horizontal strip of 10 equal segments representing tenths.
   * Filled segments use SECONDARY, empty use WHITE, borders use PRIMARY.
   *
   * @param {object} slide   PptxGenJS slide object
   * @param {number} x       Left edge x (inches)
   * @param {number} y       Top edge y (inches)
   * @param {number} w       Total strip width (inches)
   * @param {number} filled  Number of filled segments (0-10)
   * @param {object} opts    Options: h, fillColor, emptyColor
   */
  function addTenthsStrip(slide, x, y, w, filled, opts) {
    const o = opts || {};
    const h = o.h || 0.45;
    const segW = w / 10;
    const fillColor = o.fillColor || C.SECONDARY;
    const emptyColor = o.emptyColor || C.WHITE;

    validateBounds("addTenthsStrip", x, y, w + 0.6, h);

    for (let i = 0; i < 10; i++) {
      slide.addShape("rect", {
        x: x + i * segW, y, w: segW, h,
        fill: { color: i < filled ? fillColor : emptyColor },
        line: { color: C.PRIMARY, width: 1 },
      });
    }

    // Fraction label
    slide.addText(filled + "/10", {
      x: x + w + 0.1, y: y, w: 0.5, h,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL,
      align: "left", valign: "middle", margin: 0,
    });
  }

  /* ------------------------------------------------------------------ */
  /*  addAreaModel                                                       */
  /* ------------------------------------------------------------------ */

  /**
   * 10x10 grid area model for representing decimals (tenths and hundredths).
   * Full columns use SECONDARY fill, extra hundredths cells use SECONDARY,
   * empty cells use WHITE. All borders use PRIMARY.
   *
   * @param {object} slide             PptxGenJS slide object
   * @param {number} x                 Left edge x (inches)
   * @param {number} y                 Top edge y (inches)
   * @param {number} sizeIn            Side length of the grid (inches)
   * @param {number} filledTenths      Number of fully filled columns (0-10)
   * @param {number} extraHundredths   Extra cells in the next column (0-9)
   * @param {object} opts              Options: fillColor, extraColor
   */
  function addAreaModel(slide, x, y, sizeIn, filledTenths, extraHundredths, opts) {
    const o = opts || {};
    const cellSize = sizeIn / 10;
    const fillColor = o.fillColor || C.SECONDARY;
    const extraColor = o.extraColor || C.SECONDARY;

    validateBounds("addAreaModel", x, y, sizeIn, sizeIn);

    for (let col = 0; col < 10; col++) {
      for (let row = 0; row < 10; row++) {
        const isFullCol = col < filledTenths;
        const isExtraCell = col === filledTenths && row < extraHundredths;

        slide.addShape("rect", {
          x: x + col * cellSize,
          y: y + row * cellSize,
          w: cellSize, h: cellSize,
          fill: { color: isFullCol ? fillColor : (isExtraCell ? extraColor : C.WHITE) },
          line: { color: C.PRIMARY, width: 0.5 },
        });
      }
    }
  }

  /* ------------------------------------------------------------------ */
  /*  addDecimalDot                                                      */
  /* ------------------------------------------------------------------ */

  /**
   * Decimal dot positioned relative to place value chart geometry.
   * Placed between two columns (after afterCol).
   *
   * @param {object} slide     PptxGenJS slide object
   * @param {object} chartGeo  Geometry object returned by addPlaceValueChart
   * @param {number} afterCol  Column index after which to place the dot (0-based)
   * @param {object} opts      Options: dotSize, color, position ("baseline" | "center")
   */
  function addDecimalDot(slide, chartGeo, afterCol, opts) {
    const o = opts || {};
    const dotSize = o.dotSize || 0.14;
    const color = o.color || C.ALERT;
    const position = o.position || "baseline";

    const dotX = chartGeo.x + (afterCol + 1) * chartGeo.cellW - dotSize / 2;
    let dotY;
    if (position === "center") {
      dotY = chartGeo.y + chartGeo.hdrH + chartGeo.valH / 2 - dotSize / 2;
    } else {
      dotY = chartGeo.y + chartGeo.hdrH + chartGeo.valH * 0.75 - dotSize / 2;
    }

    slide.addShape("roundRect", {
      x: dotX, y: dotY, w: dotSize, h: dotSize, rectRadius: dotSize / 2,
      fill: { color },
    });
  }

  /* ------------------------------------------------------------------ */
  /*  dailyReviewSlide                                                    */
  /* ------------------------------------------------------------------ */

  /**
   * Daily Review slide. Reviews PRIOR learning only — must not teach
   * today's new content. Honour the user-provided Daily Review focus
   * exactly. Use 1-3 prompts (capped by band). No question numbers.
   *
   * Pair with withReveal() + addRevealAnswerBar() for the answer-reveal
   * slide that follows.
   *
   * @param {object}            pres
   * @param {string}            title       eg. "Daily Review: Coordinates & area"
   * @param {string|string[]}   prompts     1-3 prompts (capped by band)
   * @param {string}            notes
   * @param {string}            footer
   * @param {Function}          [drawRight] Optional callback(slide, layoutGuide) for a right-column representation (grid, table, number line, etc.)
   * @returns {object}                       The slide object
   */
  function dailyReviewSlide(pres, title, prompts, notes, footer, drawRightArg) {
    const s = pres.addSlide();
    const promptCountEarly = (Array.isArray(prompts) ? prompts : [prompts])
      .filter((p) => String(p == null ? "" : p).trim()).length;
    // A visual spec is fitted beside the prompt cards, level with them. With
    // no prompts at all the representation IS the review, so it takes the
    // full width (use heroVisualSlide for that case where possible).
    const drawRight = typeof drawRightArg === "function"
      ? drawRightArg
      : (isSpec(drawRightArg)
        ? (slide, guide) => visual.drawVisual(slide, drawRightArg, promptCountEarly === 0
          ? { x: 0.5, y: guide.panelTop, w: 9, h: guide.leftCardH }
          : { x: guide.rightX, y: guide.panelTop, w: guide.rightW, h: guide.leftCardH })
        : null);
    el.addTopBar(s, C.ACCENT);
    addStageBadge(s, 1, "Daily Review");
    el.addTitle(s, title || "Daily Review", { color: C.ACCENT });

    const list = Array.isArray(prompts) ? prompts : [prompts];
    const cleaned = list
      .map((p) => String(p == null ? "" : p).trim())
      .filter(Boolean)
      .slice(0, sz.maxQuestions || 3);

    const startY = CONTENT_TOP;
    const cardsAvailH = SAFE_BOTTOM - startY - byBand(sz, 1.1, 1.0, 0.9);
    const promptCount = Math.max(cleaned.length, 1);
    const cardW = drawRight ? 4.5 : 9;
    const cardH = Math.max(0.6, (cardsAvailH - 0.10 * (promptCount - 1)) / promptCount);
    const fontSize = byBand(sz, 28, 24, 20);

    cleaned.forEach((q, i) => {
      const y = startY + i * (cardH + 0.10);
      el.addCard(s, 0.5, y, cardW, cardH, { variant: "tint", tone: C.ACCENT });
      s.addText(String(q), {
        x: 0.7, y: y + 0.08, w: cardW - 0.4, h: cardH - 0.16,
        fontSize, fontFace: FONT_H, color: C.CHARCOAL,
        bold: true, valign: "middle", margin: 0,
        fit: "shrink", shrinkText: true,
      });
    });

    if (drawRight) {
      const layoutGuide = {
        leftCardX: 0.5, leftCardY: startY,
        leftCardW: cardW, leftCardH: cardsAvailH,
        rightX: 5.2, rightW: 4.3,
        panelTop: startY, panelTopPadded: startY + 0.08,
        safeBottom: SAFE_BOTTOM,
      };
      drawRight(s, layoutGuide);
    }

    if (footer) el.addFooter(s, footer);
    if (notes) s.addNotes(notes);
    return s;
  }

  /* ------------------------------------------------------------------ */
  /*  fluencySlide                                                        */
  /* ------------------------------------------------------------------ */

  /**
   * Fluency slide — brisk Number-and-Algebra automaticity routine.
   * Sits between Daily Review and LI/SC. Honour the user-provided Number
   * Fluency Focus exactly. Use very large numerals; no instructions.
   *
   * Multiple prompts render as side-by-side cards (1-3, capped by band).
   *
   * @param {object}            pres
   * @param {string}            title       eg. "Fluency: Division facts"
   * @param {string|string[]}   prompts     1-3 short prompts (eg. "72 / 8", "56 / 7")
   * @param {string}            notes
   * @param {string}            footer
   * @returns {object}                       The slide object
   */
  function fluencySlide(pres, title, prompts, notes, footer) {
    const s = pres.addSlide();
    el.addTopBar(s, C.ACCENT);
    addStageBadge(s, 1, "Fluency");
    el.addTitle(s, title || "Fluency", { color: C.ACCENT });

    const list = Array.isArray(prompts) ? prompts : [prompts];
    const cleaned = list
      .map((p) => String(p == null ? "" : p).trim())
      .filter(Boolean)
      .slice(0, sz.maxQuestions || 3);

    if (cleaned.length === 0) {
      if (footer) el.addFooter(s, footer);
      if (notes) s.addNotes(notes);
      return s;
    }

    // Side-by-side cards. Each card is dominated by the prompt itself, and
    // the cards stop short of the bottom so a click-revealed answer bar has
    // its full height (megaprompt 23, 68d).
    const startY = CONTENT_TOP;
    const cardH = SAFE_BOTTOM - startY - byBand(sz, 1.15, 1.05, 0.9);
    const cardGap = 0.18;
    const totalW = 9;
    const cardW = (totalW - cardGap * (cleaned.length - 1)) / cleaned.length;
    // A lone numeral or short fact is the whole slide: set it huge.
    const longest = cleaned.reduce((m, q) => Math.max(m, q.length), 0);
    const promptFontSize = cleaned.length === 1 && longest <= 3
      ? byBand(sz, 150, 130, 110)
      : (cleaned.length === 1 && longest <= 8 ? byBand(sz, 96, 84, 72) : byBand(sz, 60, 52, 42));

    cleaned.forEach((q, i) => {
      const x = 0.5 + i * (cardW + cardGap);
      el.addCard(s, x, startY, cardW, cardH, { variant: "outline", tone: C.ACCENT });
      s.addText(String(q), {
        x: x + 0.16, y: startY + 0.16, w: cardW - 0.32, h: cardH - 0.32,
        fontSize: promptFontSize, fontFace: FONT_H, color: C.CHARCOAL,
        bold: true, align: "center", valign: "middle", margin: 0,
        fit: "shrink", shrinkText: true,
      });
    });

    if (footer) el.addFooter(s, footer);
    if (notes) s.addNotes(notes);
    return s;
  }

  /* ------------------------------------------------------------------ */
  /*  Return all builders and helpers                                    */
  /* ------------------------------------------------------------------ */

  return {
    STAGE_COLORS,
    STAGE_LABELS,
    addStageBadge,
    workedExSlide,
    dailyReviewSlide,
    fluencySlide,
    addPlaceValueChart,
    addTenthsStrip,
    addAreaModel,
    addDecimalDot,
    ...manipulatives,
  };
}

module.exports = { createNumeracyBuilders };
