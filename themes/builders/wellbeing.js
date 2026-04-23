"use strict";

const { SAFE_BOTTOM, CONTENT_TOP } = require("../core/layout");
const { DEFAULT_SIZES, byBand } = require("../core/gradeBand");

/**
 * Factory that returns wellbeing-specific slide builders bound to a given
 * palette, fonts, and element helpers.
 *
 * @param {object} C       Semantic palette colours (PRIMARY, SECONDARY, ACCENT, ALERT, SUCCESS, BG_DARK, BG_LIGHT, BG_CARD, WHITE, CHARCOAL, MUTED, TEXT_ON_DARK, SUBTITLE, DECOR_1, DECOR_2)
 * @param {string} FONT_H  Heading font name
 * @param {string} FONT_B  Body font name
 * @param {object} el      Bound element helpers: addTopBar, addBadge, addTitle, addCard, addFooter, addIconCircle, addTextOnShape
 * @returns {object}        { pairShareSlide, scenarioSlide, reflectionSlide }
 */
function createWellbeingBuilders(C, FONT_H, FONT_B, el, S) {
  const sz = S || DEFAULT_SIZES;

  /* ------------------------------------------------------------------ */
  /*  pairShareSlide                                                     */
  /* ------------------------------------------------------------------ */

  /**
   * Discussion prompt slide with alternating-colour question cards.
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
  /*  scenarioSlide                                                      */
  /* ------------------------------------------------------------------ */

  /**
   * Social scenario with discussion questions — scenario card at top,
   * question cards below.
   *
   * @param {object}   pres       PptxGenJS presentation instance
   * @param {string}   badgeText  Badge label (customisable)
   * @param {string}   title      Slide title
   * @param {string}   scenario   Scenario description text
   * @param {string[]} questions  Array of discussion question strings
   * @param {string}   notes      Teacher notes
   * @param {string}   footer     Footer text
   * @returns {object}            The slide object
   */
  function scenarioSlide(pres, badgeText, title, scenario, questions, notes, footer) {
    const s = pres.addSlide();
    el.addTopBar(s, C.PRIMARY);
    el.addBadge(s, badgeText || "Scenario", { color: C.PRIMARY });
    el.addTitle(s, title);

    // Scenario card — taller for F/Y12 to fit larger body text.
    const scenH = byBand(sz, 1.95, 1.75, 1.6);
    el.addCard(s, 0.5, CONTENT_TOP, 9, scenH, { strip: C.PRIMARY, fill: C.WHITE });
    s.addText(scenario, {
      x: 0.75, y: CONTENT_TOP + 0.12, w: 8.5, h: scenH - 0.24,
      fontSize: sz.body, fontFace: FONT_B, color: C.CHARCOAL, valign: "top", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    // Question cards — capped to band's maxQuestions.
    const _qs = questions || [];
    const cappedScenQs = _qs.slice(0, sz.maxQuestions || _qs.length);
    const qStartY = CONTENT_TOP + scenH + 0.12;
    const qAvail = SAFE_BOTTOM - qStartY;
    const qCount = Math.max(cappedScenQs.length, 1);
    const gap = 0.10;
    const qHMax = byBand(sz, 1.6, 1.2, 0.95);
    const qH = Math.min(qHMax, (qAvail - gap * (qCount - 1)) / qCount);

    cappedScenQs.forEach((q, i) => {
      const y = qStartY + i * (qH + gap);
      if (y + qH > SAFE_BOTTOM) return;
      el.addCard(s, 0.5, y, 9, qH, { strip: C.ACCENT, fill: C.WHITE });
      s.addText(q, {
        x: 0.75, y: y + 0.10, w: 8.5, h: qH - 0.20,
        fontSize: sz.bodyDense, fontFace: FONT_B, color: C.CHARCOAL, valign: "middle", margin: 0,
        fit: "shrink", shrinkText: true,
      });
    });

    if (footer) el.addFooter(s, footer);
    if (notes) s.addNotes(notes);
    return s;
  }

  /* ------------------------------------------------------------------ */
  /*  reflectionSlide                                                    */
  /* ------------------------------------------------------------------ */

  /**
   * Personal reflection slide with numbered prompt cards.
   *
   * @param {object}   pres     PptxGenJS presentation instance
   * @param {string}   title    Slide title (defaults to "Time to Reflect")
   * @param {string[]} prompts  Array of reflection prompt strings
   * @param {string}   notes    Teacher notes
   * @param {string}   footer   Footer text
   * @returns {object}          The slide object
   */
  function reflectionSlide(pres, title, prompts, notes, footer) {
    const s = pres.addSlide();
    el.addTopBar(s, C.SUCCESS);
    el.addBadge(s, "Reflect", { color: C.SUCCESS });
    el.addTitle(s, title || "Time to Reflect");

    const availH = SAFE_BOTTOM - CONTENT_TOP;
    const gap = 0.10;
    const pCount = Math.max(prompts.length, 1);
    const cardHMax = byBand(sz, 1.5, 1.15, 0.95);
    const cardH = Math.min(cardHMax, (availH - gap * (pCount - 1)) / pCount);
    const numBoxW = byBand(sz, 1.2, 1.1, 1.0);
    const numFontSize = byBand(sz, 32, 28, 22);

    prompts.forEach((p, i) => {
      const y = CONTENT_TOP + i * (cardH + gap);
      if (y + cardH > SAFE_BOTTOM) return;

      // Card background
      el.addCard(s, 0.5, y, 9, cardH, { strip: C.SUCCESS, fill: C.WHITE });

      // Numbered accent box on the left
      s.addShape("rect", {
        x: 0.5, y: y, w: numBoxW, h: cardH,
        fill: { color: C.SUCCESS },
      });
      s.addText(String(i + 1), {
        x: 0.5, y: y, w: numBoxW, h: cardH,
        fontSize: numFontSize, fontFace: FONT_H, color: C.WHITE,
        bold: true, align: "center", valign: "middle", margin: 0,
      });

      // Prompt text
      s.addText(p, {
        x: 0.5 + numBoxW + 0.15, y: y + 0.10, w: 9 - numBoxW - 0.40, h: cardH - 0.20,
        fontSize: sz.body, fontFace: FONT_B, color: C.CHARCOAL, valign: "middle", margin: 0,
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

  return { pairShareSlide, scenarioSlide, reflectionSlide };
}

module.exports = { createWellbeingBuilders };
