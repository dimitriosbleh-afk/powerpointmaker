"use strict";

const { SAFE_BOTTOM, CONTENT_TOP } = require("../core/layout");
const { DEFAULT_SIZES, byBand } = require("../core/gradeBand");

/**
 * Factory that returns science-specific slide builders bound to a given
 * palette, fonts, and element helpers.
 *
 * @param {object} C       Semantic palette colours (PRIMARY, SECONDARY, ACCENT, ALERT, SUCCESS, BG_DARK, BG_LIGHT, BG_CARD, WHITE, CHARCOAL, MUTED, TEXT_ON_DARK, SUBTITLE, DECOR_1, DECOR_2)
 * @param {string} FONT_H  Heading font name
 * @param {string} FONT_B  Body font name
 * @param {object} el      Bound element helpers: addTopBar, addBadge, addTitle, addCard, addFooter, addIconCircle, addTextOnShape
 * @returns {object}        { experimentSlide, observationSlide, conclusionSlide, processFlowSlide, cycleDiagramSlide }
 */
function createScienceBuilders(C, FONT_H, FONT_B, el, S, defaults) {
  const sz = S || DEFAULT_SIZES;
  const picto = (defaults && defaults.picto) || null;

  /**
   * Stage chip with an optional pictogram at its left edge. `step.icon` is a
   * pictogram name (see listPictograms()); the chip text is "n. Label".
   */
  function drawStageChip(slide, x, y, w, h, text, color, icon, fontSize) {
    slide.addShape("roundRect", { x, y, w, h, rectRadius: h / 2, fill: { color } });
    const hasIcon = Boolean(icon && picto && picto.hasPictogram(icon));
    if (icon && picto && !hasIcon) picto.addPictogram(slide, icon, x, y, h); // emits the WARN
    if (hasIcon) {
      const d = h * 0.72;
      picto.addPictogram(slide, icon, x + 0.08, y + (h - d) / 2, d, { style: "flat", color: C.WHITE, glyphColor: C.WHITE });
    }
    const textX = hasIcon ? x + h * 0.8 + 0.06 : x + 0.08;
    slide.addText(String(text), {
      x: textX, y, w: x + w - textX - 0.08, h,
      fontSize, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: hasIcon ? "left" : "center", valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });
  }
  function drawHorizontalArrow(slide, x, y, w, color, direction) {
    slide.addShape("line", {
      x, y, w: Math.max(0.05, w), h: 0,
      line: {
        color,
        width: 1.4,
        beginArrowType: direction === "left" ? "triangle" : "none",
        endArrowType: direction === "right" ? "triangle" : "none",
      },
    });
  }

  /**
   * Straight arrow from (x1, y1) to (x2, y2) with the head at the end point.
   * PptxGenJS lines run top-left to bottom-right inside their box, so other
   * directions use flipH / flipV (negative sizes corrupt the file).
   */
  function drawArrowBetween(slide, x1, y1, x2, y2, color) {
    const w = Math.max(0.05, Math.abs(x2 - x1));
    const h = Math.max(0.05, Math.abs(y2 - y1));
    slide.addShape("line", {
      x: Math.min(x1, x2), y: Math.min(y1, y2), w, h,
      flipH: x2 < x1,
      flipV: y2 < y1,
      line: { color, width: 1.6, beginArrowType: "none", endArrowType: "triangle" },
    });
  }

  function drawVerticalArrow(slide, x, y, h, color, direction) {
    slide.addShape("line", {
      x, y, w: 0, h: Math.max(0.05, h),
      line: {
        color,
        width: 1.4,
        beginArrowType: direction === "up" ? "triangle" : "none",
        endArrowType: direction === "down" ? "triangle" : "none",
      },
    });
  }

  /* ------------------------------------------------------------------ */
  /*  experimentSlide                                                    */
  /* ------------------------------------------------------------------ */

  /**
   * Scientific method slide — two-column layout with hypothesis + materials
   * on the left and method steps on the right.
   *
   * @param {object}   pres        PptxGenJS presentation instance
   * @param {string}   badgeText   Badge label (customisable)
   * @param {string}   title       Slide title
   * @param {string}   hypothesis  The hypothesis statement
   * @param {string[]} materials   Array of material/equipment strings
   * @param {string[]} method      Array of method step strings
   * @param {string}   notes       Teacher notes
   * @param {string}   footer      Footer text
   * @returns {object}             The slide object
   */
  function experimentSlide(pres, badgeText, title, hypothesis, materials, method, notes, footer) {
    const s = pres.addSlide();
    el.addTopBar(s, C.PRIMARY);
    el.addBadge(s, badgeText || "Experiment", { color: C.PRIMARY });
    el.addTitle(s, title);

    const cardH = SAFE_BOTTOM - CONTENT_TOP;
    const HDR_PAD = 0.32;
    const LEFT_W = 4.3;
    const RIGHT_W = 4.5;
    const RIGHT_X = 0.5 + LEFT_W + 0.2;

    // --- Left card: Hypothesis + Materials ---
    el.addCard(s, 0.5, CONTENT_TOP, LEFT_W, cardH, { strip: C.PRIMARY, fill: C.WHITE });

    // Hypothesis section
    s.addText("Hypothesis", {
      x: 0.75, y: CONTENT_TOP + 0.08, w: LEFT_W - 0.50, h: 0.28,
      fontSize: sz.sectionLabel, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });

    const hypoH = Math.min(1.2, cardH * 0.35);
    s.addText(hypothesis || "", {
      x: 0.75, y: CONTENT_TOP + HDR_PAD, w: LEFT_W - 0.50, h: hypoH,
      fontSize: sz.bodyDense, fontFace: FONT_B, color: C.CHARCOAL, valign: "top", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    // Materials section
    const matY = CONTENT_TOP + HDR_PAD + hypoH + 0.12;
    s.addText("Materials", {
      x: 0.75, y: matY, w: LEFT_W - 0.50, h: 0.28,
      fontSize: sz.sectionLabel, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });

    if (materials && materials.length) {
      const matBodyY = matY + 0.30;
      const matBodyH = CONTENT_TOP + cardH - matBodyY - 0.08;
      const fs = materials.length > 8 ? Math.max(sz.bodyDense * sz._shrink, 10) : sz.bodyDense;
      s.addText(materials.map((m, i) => ({
        text: m,
        options: {
          bullet: true,
          breakLine: i < materials.length - 1,
          fontSize: fs,
          color: C.CHARCOAL,
        },
      })), {
        x: 0.75, y: matBodyY, w: LEFT_W - 0.50, h: matBodyH,
        fontFace: FONT_B, valign: "top", margin: 0,
        fit: "shrink", shrinkText: true,
      });
    }

    // --- Right card: Method ---
    el.addCard(s, RIGHT_X, CONTENT_TOP, RIGHT_W, cardH, { strip: C.SECONDARY, fill: C.WHITE });

    s.addText("Method", {
      x: RIGHT_X + 0.20, y: CONTENT_TOP + 0.08, w: RIGHT_W - 0.45, h: 0.28,
      fontSize: sz.sectionLabel, fontFace: FONT_B, color: C.SECONDARY, bold: true, margin: 0,
    });

    if (method && method.length) {
      const methBodyH = cardH - HDR_PAD - 0.08;
      const fs = method.length > 8 ? Math.max(sz.bodyDense * sz._shrink, 10) : sz.bodyDense + 0.5;
      s.addText(method.map((step, i) => ({
        text: (i + 1) + ".  " + step,
        options: {
          breakLine: i < method.length - 1,
          fontSize: fs,
          color: C.CHARCOAL,
        },
      })), {
        x: RIGHT_X + 0.20, y: CONTENT_TOP + HDR_PAD, w: RIGHT_W - 0.45, h: methBodyH,
        fontFace: FONT_B, valign: "top", margin: 0,
        fit: "shrink", shrinkText: true,
      });
    }

    if (footer) el.addFooter(s, footer);
    if (notes) s.addNotes(notes);
    return s;
  }

  /* ------------------------------------------------------------------ */
  /*  observationSlide                                                   */
  /* ------------------------------------------------------------------ */

  /**
   * Recording observations slide — prompt cards with alternating
   * strip colours.
   *
   * Question numbers are OFF by default (megaprompt §15d forbids numbered
   * Q1/Q2/Q3 on student-facing teaching slides). Pass
   * `opts.numbered: true` only when the slide is acting as a formal
   * assessment item.
   *
   * @param {object}   pres       PptxGenJS presentation instance
   * @param {string}   badgeText  Badge label (customisable)
   * @param {string}   title      Slide title (defaults to "What Did You Observe?")
   * @param {string[]} prompts    Array of observation prompt strings
   * @param {string}   notes      Teacher notes
   * @param {string}   footer     Footer text
   * @param {object}   [opts]     { numbered }
   * @returns {object}            The slide object
   */
  function observationSlide(pres, badgeText, title, prompts, notes, footer, opts) {
    const s = pres.addSlide();
    const o = opts || {};
    el.addTopBar(s, C.SECONDARY);
    el.addBadge(s, badgeText || "Observe", { color: C.SECONDARY });
    el.addTitle(s, title || "What Did You Observe?");

    const availH = SAFE_BOTTOM - CONTENT_TOP;
    const gap = 0.10;
    const pCount = Math.max(prompts.length, 1);
    const cardHMax = byBand(sz, 1.6, 1.2, 0.95);
    const cardH = Math.min(cardHMax, (availH - gap * (pCount - 1)) / pCount);
    const promptFontSize = sz.body;
    const numFontSize = sz.body + 1;
    const numbered = Boolean(o.numbered);

    prompts.forEach((p, i) => {
      const y = CONTENT_TOP + i * (cardH + gap);
      if (y + cardH > SAFE_BOTTOM) return;

      el.addCard(s, 0.5, y, 9, cardH, {
        strip: i % 2 === 0 ? C.PRIMARY : C.SECONDARY,
        fill: C.WHITE,
      });

      if (numbered) {
        const numStr = String(i + 1) + ".  ";
        s.addText([
          { text: numStr, options: { bold: true, fontSize: numFontSize, color: C.PRIMARY } },
          { text: String(p), options: { fontSize: promptFontSize, color: C.CHARCOAL } },
        ], {
          x: 0.75, y: y + 0.10, w: 8.5, h: cardH - 0.20,
          fontFace: FONT_B, valign: "middle", margin: 0,
          fit: "shrink", shrinkText: true,
        });
      } else {
        s.addText(String(p), {
          x: 0.75, y: y + 0.10, w: 8.5, h: cardH - 0.20,
          fontSize: promptFontSize, fontFace: FONT_B, color: C.CHARCOAL,
          valign: "middle", margin: 0,
          fit: "shrink", shrinkText: true,
        });
      }
    });

    if (footer) el.addFooter(s, footer);
    if (notes) s.addNotes(notes);
    return s;
  }

  /* ------------------------------------------------------------------ */
  /*  conclusionSlide                                                    */
  /* ------------------------------------------------------------------ */

  /**
   * Drawing conclusions slide — three stacked cards: question, evidence
   * (auto-sized), and conclusion.
   *
   * @param {object}   pres        PptxGenJS presentation instance
   * @param {string}   badgeText   Badge label (customisable)
   * @param {string}   title       Slide title
   * @param {string}   question    The original inquiry question
   * @param {string[]} evidence    Array of evidence/observation strings
   * @param {string}   conclusion  Conclusion statement text
   * @param {string}   notes       Teacher notes
   * @param {string}   footer      Footer text
   * @returns {object}             The slide object
   */
  function conclusionSlide(pres, badgeText, title, question, evidence, conclusion, notes, footer) {
    const s = pres.addSlide();
    el.addTopBar(s, C.ACCENT);
    el.addBadge(s, badgeText || "Conclude", { color: C.ACCENT });
    el.addTitle(s, title);

    const GAP = 0.12;
    const HDR_PAD = 0.32;
    const totalAvail = SAFE_BOTTOM - CONTENT_TOP;

    // Fixed heights for question and conclusion; evidence gets the remainder
    const qH = 0.8;
    const concH = 1.0;
    const evidH = Math.max(totalAvail - qH - concH - GAP * 2, 0.6);

    let curY = CONTENT_TOP;

    // Question card
    el.addCard(s, 0.5, curY, 9, qH, { strip: C.PRIMARY, fill: C.WHITE });
    s.addText("Our Question", {
      x: 0.75, y: curY + 0.08, w: 5, h: 0.28,
      fontSize: sz.sectionLabel, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText(question || "", {
      x: 0.75, y: curY + HDR_PAD, w: 8.5, h: qH - HDR_PAD - 0.06,
      fontSize: sz.body, fontFace: FONT_B, color: C.CHARCOAL, valign: "top", margin: 0,
      fit: "shrink", shrinkText: true,
    });
    curY += qH + GAP;

    // Evidence card (auto-sized to fill)
    el.addCard(s, 0.5, curY, 9, evidH, { strip: C.SECONDARY, fill: C.WHITE });
    s.addText("Evidence", {
      x: 0.75, y: curY + 0.08, w: 5, h: 0.28,
      fontSize: sz.sectionLabel, fontFace: FONT_B, color: C.SECONDARY, bold: true, margin: 0,
    });

    if (evidence && evidence.length) {
      const evidBodyH = evidH - HDR_PAD - 0.06;
      const fs = evidence.length > 6 ? Math.max(sz.bodyDense * sz._shrink, 10) : sz.bodyDense;
      s.addText(evidence.map((e, i) => ({
        text: e,
        options: {
          bullet: true,
          breakLine: i < evidence.length - 1,
          fontSize: fs,
          color: C.CHARCOAL,
        },
      })), {
        x: 0.75, y: curY + HDR_PAD, w: 8.5, h: evidBodyH,
        fontFace: FONT_B, valign: "top", margin: 0,
        fit: "shrink", shrinkText: true,
      });
    }
    curY += evidH + GAP;

    // Conclusion card
    el.addCard(s, 0.5, curY, 9, concH, { strip: C.ACCENT, fill: C.WHITE });
    s.addText("Conclusion", {
      x: 0.75, y: curY + 0.08, w: 5, h: 0.28,
      fontSize: sz.sectionLabel, fontFace: FONT_B, color: C.ACCENT, bold: true, margin: 0,
    });
    s.addText(conclusion || "", {
      x: 0.75, y: curY + HDR_PAD, w: 8.5, h: concH - HDR_PAD - 0.06,
      fontSize: sz.body, fontFace: FONT_B, color: C.CHARCOAL, valign: "top", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    if (footer) el.addFooter(s, footer);
    if (notes) s.addNotes(notes);
    return s;
  }

  /* ------------------------------------------------------------------ */
  /*  processFlowSlide                                                   */
  /* ------------------------------------------------------------------ */

  /**
   * Ordered process / system slide with an instruction card on the left
   * and a numbered flow on the right. Designed for journeys, cycles,
   * organs in order, life cycles, and other science sequences where a
   * visual anchor should carry the concept.
   *
   * @param {object} pres
   * @param {string} badgeText
   * @param {string} title
   * @param {string} promptTitle
   * @param {string[]} promptLines
   * @param {{label:string, detail:string, color?:string}[]} steps
   * @param {string} notes
   * @param {string} footer
   * @returns {object}
   */
  function processFlowSlide(pres, badgeText, title, promptTitle, promptLines, steps, notes, footer) {
    const s = pres.addSlide();
    el.addTopBar(s, C.SECONDARY);
    el.addBadge(s, badgeText || "Process", { color: C.SUCCESS });
    el.addTitle(s, title);

    const promptItems = [{ text: promptTitle || "Think together", role: "header" }];
    (promptLines || []).forEach((line, index) => {
      promptItems.push({ text: "", role: "spacer" });
      promptItems.push({
        text: line,
        role: index === (promptLines || []).length - 1 && /seconds|minutes|now/i.test(String(line || "")) ? "emphasis" : "body",
      });
    });

    el.addInstructionCard(s, promptItems, {
      x: 0.5, y: CONTENT_TOP, w: 4.2, h: 2.55,
      strip: C.SECONDARY, fill: C.WHITE,
    });

    const flowX = 5.0;
    const flowY = CONTENT_TOP;
    const flowW = 4.5;
    const flowH = 3.55;
    el.addCard(s, flowX, flowY, flowW, flowH, { strip: C.PRIMARY, fill: C.WHITE });
    s.addText("Process flow", {
      x: flowX + 0.22, y: flowY + 0.08, w: 2.6, h: 0.26,
      fontSize: sz.sectionLabel + 1, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0,
    });

    const safeSteps = (steps || []).slice(0, 6);
    const chipPalette = [C.PRIMARY, C.SECONDARY, C.ACCENT, C.ALERT, C.SUCCESS, C.PRIMARY];
    const rowGap = safeSteps.length > 4 ? 0.10 : 0.16;
    const rowH = Math.min(0.62, (flowH - 0.56 - rowGap * Math.max(safeSteps.length - 1, 0)) / Math.max(safeSteps.length, 1));
    const chipH = Math.min(0.46, rowH - 0.04);
    const chipW = 1.85;
    const chipFont = Math.min(sz.chip + 2, Math.round(chipH * 30));
    const detailFont = Math.min(sz.bodyDense, Math.max(sz.chip, Math.round(rowH * 26)));

    safeSteps.forEach((step, index) => {
      const rowY = flowY + 0.42 + index * (rowH + rowGap);
      const chipColor = step && step.color ? step.color : chipPalette[index];
      drawStageChip(s, flowX + 0.20, rowY + (rowH - chipH) / 2, chipW, chipH,
        `${index + 1}. ${String((step && step.label) || "")}`, chipColor, step && step.icon, chipFont);
      s.addText(String((step && step.detail) || ""), {
        x: flowX + 0.20 + chipW + 0.12, y: rowY, w: flowW - chipW - 0.52, h: rowH,
        fontSize: detailFont, fontFace: FONT_B, color: C.CHARCOAL,
        margin: 0, valign: "middle", fit: "shrink", shrinkText: true,
      });
      if (index < safeSteps.length - 1) {
        s.addShape("line", {
          x: flowX + 0.20 + chipW / 2, y: rowY + (rowH + chipH) / 2, w: 0, h: rowGap + (rowH - chipH) / 2 + 0.02,
          line: { color: C.MUTED, width: 1.2, beginArrowType: "none", endArrowType: "triangle" },
        });
      }
    });

    if (footer) el.addFooter(s, footer);
    if (notes) s.addNotes(notes);
    return s;
  }

  /* ------------------------------------------------------------------ */
  /*  cycleDiagramSlide                                                  */
  /* ------------------------------------------------------------------ */

  /**
   * Cycle diagram slide with a prompt card on the left and a proper
   * labelled cycle on the right. Designed for water cycles, rock cycles,
   * life cycles, seasons, and any science content where the return-loop
   * matters conceptually.
   *
   * @param {object} pres
   * @param {string} badgeText
   * @param {string} title
   * @param {string} promptTitle
   * @param {string[]} promptLines
   * @param {string} centerLabel
   * @param {{label:string, detail:string, color?:string}[]} steps
   * @param {string} notes
   * @param {string} footer
   * @returns {object}
   */
  function cycleDiagramSlide(pres, badgeText, title, promptTitle, promptLines, centerLabel, steps, notes, footer) {
    const s = pres.addSlide();
    el.addTopBar(s, C.SECONDARY);
    el.addBadge(s, badgeText || "Cycle", { color: C.SUCCESS });
    el.addTitle(s, title);

    const promptItems = [{ text: promptTitle || "With your partner", role: "header" }];
    (promptLines || []).forEach((line, index) => {
      promptItems.push({ text: "", role: "spacer" });
      promptItems.push({
        text: line,
        role: index === (promptLines || []).length - 1 && /seconds|minutes|now/i.test(String(line || "")) ? "emphasis" : "body",
      });
    });

    // Narrower prompt card so the cycle itself gets the room (it is the
    // visual anchor the megaprompt requires for cycle content).
    el.addInstructionCard(s, promptItems, {
      x: 0.5, y: CONTENT_TOP, w: 3.0, h: SAFE_BOTTOM - CONTENT_TOP,
      strip: C.SECONDARY, fill: C.WHITE,
    });

    const cardX = 3.7;
    const cardY = CONTENT_TOP;
    const cardW = 5.8;
    const cardH = SAFE_BOTTOM - CONTENT_TOP;
    el.addCard(s, cardX, cardY, cardW, cardH, { variant: "tint", tone: C.PRIMARY });

    const safeSteps = (steps || []).slice(0, 4);
    const palette = [C.PRIMARY, C.SECONDARY, C.ACCENT, C.SUCCESS];
    const hasDetail = safeSteps.some((st) => st && st.detail);
    const legendH = hasDetail ? 0.5 : 0;
    const legendGapY = 0.1;
    const legendRows = hasDetail ? Math.ceil(safeSteps.length / 2) : 0;
    const legendBlockH = legendRows * legendH + Math.max(legendRows - 1, 0) * legendGapY;
    const diagramH = cardH - 0.3 - (legendBlockH ? legendBlockH + 0.2 : 0);
    const chipW = 1.95;
    const chipH = 0.46;
    const cx = cardX + cardW / 2;
    const cy = cardY + 0.15 + diagramH / 2;
    const orbitX = (cardW - chipW) / 2 - 0.35;
    const orbitY = (diagramH - chipH) / 2 - 0.1;
    const positions = [
      { x: cx, y: cy - orbitY },
      { x: cx + orbitX, y: cy },
      { x: cx, y: cy + orbitY },
      { x: cx - orbitX, y: cy },
    ];

    // Faint orbit ring so the loop reads as a loop before the arrows do.
    s.addShape("ellipse", {
      x: cx - orbitX, y: cy - orbitY, w: orbitX * 2, h: orbitY * 2,
      fill: { color: C.WHITE, transparency: 100 },
      line: { color: C.PRIMARY_LINE || C.MUTED, width: 1.5, dashType: "dash" },
    });

    const centreW = Math.min(1.9, orbitX * 2 - chipW - 0.3);
    el.addTextOnShape(s, centerLabel || "Cycle", {
      x: cx - centreW / 2, y: cy - 0.36, w: centreW, h: 0.72, rectRadius: 0.14,
      fill: { color: C.WHITE },
      line: { color: C.PRIMARY, width: 1.4 },
    }, {
      fontSize: sz.sectionLabel + 3, fontFace: FONT_H, color: C.PRIMARY, bold: true,
    });

    safeSteps.forEach((step, index) => {
      const pos = positions[index];
      const color = step && step.color ? step.color : palette[index];
      drawStageChip(s, pos.x - chipW / 2, pos.y - chipH / 2, chipW, chipH,
        `${index + 1}. ${String((step && step.label) || "")}`, color, step && step.icon, sz.chip);
    });

    if (safeSteps.length >= 4) {
      // Clockwise arrows between neighbouring chips, drawn as diagonals that
      // follow the ring, so no arrow ever crosses a chip label.
      const gapX = chipW / 2 + 0.14;
      const gapY = chipH / 2 + 0.12;
      drawArrowBetween(s, cx + gapX, cy - orbitY + 0.06, cx + orbitX - 0.12, cy - gapY, C.CHARCOAL);       // top -> right
      drawArrowBetween(s, cx + orbitX - 0.12, cy + gapY, cx + gapX, cy + orbitY - 0.06, C.CHARCOAL);      // right -> bottom
      drawArrowBetween(s, cx - gapX, cy + orbitY - 0.06, cx - orbitX + 0.12, cy + gapY, C.CHARCOAL);      // bottom -> left
      drawArrowBetween(s, cx - orbitX + 0.12, cy - gapY, cx - gapX, cy - orbitY + 0.06, C.CHARCOAL);      // left -> top
    }

    if (hasDetail) {
      const legendY = cardY + cardH - 0.15 - legendBlockH;
      const legendGapX = 0.16;
      const legendW = (cardW - 0.4 - legendGapX) / 2;
      safeSteps.forEach((step, index) => {
        const row = Math.floor(index / 2);
        const col = index % 2;
        const lx = cardX + 0.2 + col * (legendW + legendGapX);
        const ly = legendY + row * (legendH + legendGapY);
        const color = step && step.color ? step.color : palette[index];
        s.addShape("roundRect", {
          x: lx, y: ly, w: legendW, h: legendH, rectRadius: 0.08,
          fill: { color: C.WHITE },
          line: { color, width: 1.2 },
        });
        s.addShape("rect", { x: lx, y: ly, w: 0.08, h: legendH, fill: { color } });
        s.addText(`${index + 1}. ${String((step && step.detail) || "")}`, {
          x: lx + 0.16, y: ly + 0.04, w: legendW - 0.24, h: legendH - 0.08,
          fontSize: Math.max(sz.caption + 2, sz.chip), fontFace: FONT_B, color: C.CHARCOAL,
          margin: 0, align: "left", valign: "middle", fit: "shrink", shrinkText: true,
        });
      });
    }

    if (footer) el.addFooter(s, footer);
    if (notes) s.addNotes(notes);
    return s;
  }

  /* ------------------------------------------------------------------ */
  /*  Return all builders                                                */
  /* ------------------------------------------------------------------ */

  return { experimentSlide, observationSlide, conclusionSlide, processFlowSlide, cycleDiagramSlide };
}

module.exports = { createScienceBuilders };
