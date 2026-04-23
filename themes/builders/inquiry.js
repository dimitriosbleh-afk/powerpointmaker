"use strict";

const { SAFE_BOTTOM, CONTENT_TOP } = require("../core/layout");
const { DEFAULT_SIZES, byBand } = require("../core/gradeBand");

/**
 * Factory that returns inquiry-specific slide builders bound to a given
 * palette, fonts, and element helpers.
 *
 * @param {object} C       Semantic palette colours (PRIMARY, SECONDARY, ACCENT, ALERT, SUCCESS, BG_DARK, BG_LIGHT, BG_CARD, WHITE, CHARCOAL, MUTED, TEXT_ON_DARK, SUBTITLE, DECOR_1, DECOR_2)
 * @param {string} FONT_H  Heading font name
 * @param {string} FONT_B  Body font name
 * @param {object} el      Bound element helpers: addTopBar, addBadge, addTitle, addCard, addFooter, addIconCircle, addTextOnShape
 * @returns {object}        { investigationSlide, findingsSlide, pairShareSlide }
 */
function createInquiryBuilders(C, FONT_H, FONT_B, el, S) {
  const sz = S || DEFAULT_SIZES;

  /* ------------------------------------------------------------------ */
  /*  investigationSlide                                                 */
  /* ------------------------------------------------------------------ */

  /**
   * Inquiry question framing slide — question card, optional hypothesis card,
   * optional numbered steps card. Auto-sizes to fit within SAFE_BOTTOM.
   *
   * @param {object}   pres        PptxGenJS presentation instance
   * @param {string}   badgeText   Badge label (customisable)
   * @param {string}   title       Slide title
   * @param {string}   question    The inquiry question text
   * @param {string}   hypothesis  Optional hypothesis statement
   * @param {string[]} steps       Optional array of investigation steps
   * @param {string}   notes       Teacher notes
   * @param {string}   footer      Footer text
   * @returns {object}             The slide object
   */
  function investigationSlide(pres, badgeText, title, question, hypothesis, steps, notes, footer) {
    const s = pres.addSlide();
    el.addTopBar(s, C.PRIMARY);
    el.addBadge(s, badgeText || "Investigate", { color: C.PRIMARY });
    el.addTitle(s, title);

    const GAP = 0.12;
    const HDR_PAD = 0.32; // space for section header inside each card

    // Count sections and calculate available height
    const sectionCount = 1 + (hypothesis ? 1 : 0) + (steps && steps.length ? 1 : 0);
    const totalGaps = (sectionCount - 1) * GAP;
    const totalAvail = SAFE_BOTTOM - CONTENT_TOP - totalGaps;

    // Distribute heights proportionally
    let qH, hypoH, stepsH;
    if (sectionCount === 1) {
      qH = totalAvail;
    } else if (sectionCount === 2) {
      if (hypothesis) {
        qH = totalAvail * 0.55;
        hypoH = totalAvail * 0.45;
      } else {
        qH = totalAvail * 0.35;
        stepsH = totalAvail * 0.65;
      }
    } else {
      qH = totalAvail * 0.30;
      hypoH = totalAvail * 0.25;
      stepsH = totalAvail * 0.45;
    }

    // Clamp question card
    qH = Math.max(qH, 0.6);
    qH = Math.min(qH, 1.5);

    let curY = CONTENT_TOP;

    // Inquiry Question card
    el.addCard(s, 0.5, curY, 9, qH, { strip: C.PRIMARY, fill: C.WHITE });
    s.addText("Inquiry Question", {
      x: 0.75, y: curY + 0.08, w: 5, h: 0.26,
      fontSize: sz.sectionLabel, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText(question, {
      x: 0.75, y: curY + HDR_PAD, w: 8.5, h: qH - HDR_PAD - 0.08,
      fontSize: sz.body + 1, fontFace: FONT_B, color: C.CHARCOAL, valign: "top", margin: 0,
      fit: "shrink", shrinkText: true,
    });
    curY += qH + GAP;

    // Hypothesis card (optional)
    if (hypothesis) {
      hypoH = Math.max(hypoH, 0.5);
      hypoH = Math.min(hypoH, 1.2);
      el.addCard(s, 0.5, curY, 9, hypoH, { strip: C.ACCENT, fill: C.WHITE });
      s.addText("Our Hypothesis", {
        x: 0.75, y: curY + 0.08, w: 5, h: 0.26,
        fontSize: sz.sectionLabel, fontFace: FONT_B, color: C.ACCENT, bold: true, margin: 0,
      });
      s.addText(hypothesis, {
        x: 0.75, y: curY + HDR_PAD, w: 8.5, h: hypoH - HDR_PAD - 0.08,
        fontSize: sz.body, fontFace: FONT_B, color: C.CHARCOAL, valign: "top", margin: 0,
        fit: "shrink", shrinkText: true,
      });
      curY += hypoH + GAP;
    }

    // Steps card (optional)
    if (steps && steps.length) {
      // Use whatever remains down to SAFE_BOTTOM
      stepsH = Math.max(SAFE_BOTTOM - curY, 0.5);
      el.addCard(s, 0.5, curY, 9, stepsH, { strip: C.SECONDARY, fill: C.WHITE });
      s.addText("Investigation Steps", {
        x: 0.75, y: curY + 0.08, w: 5, h: 0.26,
        fontSize: sz.sectionLabel, fontFace: FONT_B, color: C.SECONDARY, bold: true, margin: 0,
      });

      const stepsBodyH = stepsH - HDR_PAD - 0.08;
      const fs = steps.length > 6 ? Math.max(sz.bodyDense * sz._shrink, 10) : sz.bodyDense;
      s.addText(steps.map((step, i) => ({
        text: (i + 1) + ".  " + step,
        options: {
          breakLine: i < steps.length - 1,
          fontSize: fs,
          color: C.CHARCOAL,
        },
      })), {
        x: 0.75, y: curY + HDR_PAD, w: 8.5, h: stepsBodyH,
        fontFace: FONT_B, valign: "top", margin: 0,
        fit: "shrink", shrinkText: true,
      });
    }

    if (footer) el.addFooter(s, footer);
    if (notes) s.addNotes(notes);
    return s;
  }

  /* ------------------------------------------------------------------ */
  /*  findingsSlide                                                      */
  /* ------------------------------------------------------------------ */

  /**
   * Evidence/findings display slide — findings bullet card with optional
   * conclusion card below.
   *
   * @param {object}   pres        PptxGenJS presentation instance
   * @param {string}   badgeText   Badge label (customisable)
   * @param {string}   title       Slide title
   * @param {string[]} findings    Array of finding/evidence strings
   * @param {string}   conclusion  Optional conclusion text
   * @param {string}   notes       Teacher notes
   * @param {string}   footer      Footer text
   * @returns {object}             The slide object
   */
  function findingsSlide(pres, badgeText, title, findings, conclusion, notes, footer) {
    const s = pres.addSlide();
    el.addTopBar(s, C.SECONDARY);
    el.addBadge(s, badgeText || "Findings", { color: C.SECONDARY });
    el.addTitle(s, title);

    const GAP = 0.12;
    const HDR_PAD = 0.32;
    const totalAvail = SAFE_BOTTOM - CONTENT_TOP;

    // Calculate card heights
    let findH, concH;
    if (conclusion) {
      concH = Math.min(1.2, totalAvail * 0.30);
      findH = totalAvail - GAP - concH;
    } else {
      findH = totalAvail;
    }

    // Findings card
    el.addCard(s, 0.5, CONTENT_TOP, 9, findH, { strip: C.SECONDARY, fill: C.WHITE });
    s.addText("What We Found", {
      x: 0.75, y: CONTENT_TOP + 0.08, w: 5, h: 0.26,
      fontSize: sz.sectionLabel, fontFace: FONT_B, color: C.SECONDARY, bold: true, margin: 0,
    });

    if (findings && findings.length) {
      const bodyH = findH - HDR_PAD - 0.08;
      const fs = findings.length > 6 ? Math.max(sz.bodyDense * sz._shrink, 10) : sz.body;
      s.addText(findings.map((f, i) => ({
        text: f,
        options: {
          bullet: true,
          breakLine: i < findings.length - 1,
          fontSize: fs,
          color: C.CHARCOAL,
        },
      })), {
        x: 0.75, y: CONTENT_TOP + HDR_PAD, w: 8.5, h: bodyH,
        fontFace: FONT_B, valign: "top", margin: 0,
        fit: "shrink", shrinkText: true,
      });
    }

    // Conclusion card (optional)
    if (conclusion) {
      const concY = CONTENT_TOP + findH + GAP;
      el.addCard(s, 0.5, concY, 9, concH, { strip: C.ACCENT, fill: C.WHITE });
      s.addText("Our Conclusion", {
        x: 0.75, y: concY + 0.08, w: 5, h: 0.26,
        fontSize: sz.sectionLabel, fontFace: FONT_B, color: C.ACCENT, bold: true, margin: 0,
      });
      s.addText(conclusion, {
        x: 0.75, y: concY + HDR_PAD, w: 8.5, h: concH - HDR_PAD - 0.08,
        fontSize: sz.body, fontFace: FONT_B, color: C.CHARCOAL, valign: "top", margin: 0,
        fit: "shrink", shrinkText: true,
      });
    }

    if (footer) el.addFooter(s, footer);
    if (notes) s.addNotes(notes);
    return s;
  }

  /* ------------------------------------------------------------------ */
  /*  pairShareSlide                                                     */
  /* ------------------------------------------------------------------ */

  /**
   * Discussion prompt slide with alternating-colour question cards.
   * Same pattern as the wellbeing version.
   *
   * @param {object}   pres       PptxGenJS presentation instance
   * @param {string}   title      Slide title (defaults to "Discuss with Your Partner")
   * @param {string[]} questions  Array of discussion question strings
   * @param {string}   notes      Teacher notes
   * @param {string}   footer     Footer text
   * @returns {object}            The slide object
   */
  function pairShareSlide(pres, title, questions, notes, footer) {
    const s = pres.addSlide();
    el.addTopBar(s, C.SECONDARY);
    el.addBadge(s, "Pair-Share", { color: C.SECONDARY });
    el.addTitle(s, title || "Discuss with Your Partner");

    const _qs = questions || [];
    const cappedQs = _qs.slice(0, sz.maxQuestions || _qs.length);
    const availH = SAFE_BOTTOM - CONTENT_TOP;
    const gap = 0.10;
    const qCount = Math.max(cappedQs.length, 1);
    const qHMax = byBand(sz, 2.6, 1.8, 0.95);
    const qH = Math.min(qHMax, (availH - gap * (qCount - 1)) / qCount);
    const fontSize = cappedQs.length >= 5 ? sz.bodyDense : sz.body + 1;

    cappedQs.forEach((q, i) => {
      const y = CONTENT_TOP + i * (qH + gap);
      if (y + qH > SAFE_BOTTOM) return;
      el.addCard(s, 0.5, y, 9, qH, {
        strip: i % 2 === 0 ? C.PRIMARY : C.SECONDARY,
        fill: C.WHITE,
      });
      s.addText(q, {
        x: 0.75, y: y + 0.10, w: 8.5, h: qH - 0.20,
        fontSize, fontFace: FONT_B, color: C.CHARCOAL, valign: "middle", margin: 0,
        fit: "shrink", shrinkText: true,
      });
    });

    if (footer) el.addFooter(s, footer);
    if (notes) s.addNotes(notes);
    return s;
  }

  /* ------------------------------------------------------------------ */
  /*  Return all builders                                                */
  /* ------------------------------------------------------------------ */

  return { investigationSlide, findingsSlide, pairShareSlide };
}

module.exports = { createInquiryBuilders };
