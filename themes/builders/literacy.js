"use strict";

const { SAFE_BOTTOM, CONTENT_TOP } = require("../core/layout");
const { DEFAULT_SIZES, byBand } = require("../core/gradeBand");
const { contrastRatio } = require("../core/contrast");

/**
 * Factory that returns literacy-specific slide builders bound to a given
 * palette, fonts, element helpers, and grade-band sizing table S.
 *
 * @param {object} C       Semantic palette colours (PRIMARY, SECONDARY, ACCENT, ALERT, SUCCESS, BG_DARK, BG_LIGHT, BG_CARD, WHITE, CHARCOAL, MUTED, TEXT_ON_DARK, SUBTITLE, DECOR_1, DECOR_2)
 * @param {string} FONT_H  Heading font name
 * @param {string} FONT_B  Body font name
 * @param {object} el       Bound element helpers: addTopBar, addBadge, addTitle, addCard, addFooter, addIconCircle, addTextOnShape
 * @param {object} [S]     Grade-band sizing table (from getGradeSizes)
 * @returns {object}        { vocabSlide, quoteSlide, modellingSlide, pairShareSlide }
 */
function createLiteracyBuilders(C, FONT_H, FONT_B, el, S) {
  const sz = S || DEFAULT_SIZES;
  const readableOn = (fill) => (
    contrastRatio(C.WHITE, fill) >= contrastRatio(C.CHARCOAL, fill) ? C.WHITE : C.CHARCOAL
  );

  /* ------------------------------------------------------------------ */
  /*  vocabSlide                                                         */
  /* ------------------------------------------------------------------ */

  /**
   * Vocabulary focus slide — word, part of speech badge, definition, example sentence.
   *
   * Megaprompt §29 requires "one large word, one meaningful graphic, a
   * short student-friendly meaning, a quick oral or physical practice
   * routine". Pass `opts.image` (local file path or data URI) to add the
   * graphic; without an image the slide falls back to the legacy three-
   * card text layout for back-compat.
   *
   * @param {object} pres              PptxGenJS presentation instance
   * @param {string} word              The vocabulary word
   * @param {string} partOfSpeech      Part of speech label (e.g. "noun")
   * @param {string} definition        Definition text
   * @param {string} exampleSentence   Example sentence (displayed in quotes)
   * @param {string} notes             Teacher notes
   * @param {string} footer            Footer text
   * @param {object} [opts]            { image, routine, sourceLink }
   *   - image: path or data URI to an instructional image
   *   - routine: short student-facing routine cue (eg. "Say it together. Use it in a sentence.")
   *   - sourceLink: object { sentence, source } — only set when the source text is supplied
   * @returns {object}                 The slide object
   */
  function vocabSlide(pres, word, partOfSpeech, definition, exampleSentence, notes, footer, opts) {
    const s = pres.addSlide();
    const o = opts || {};
    const hasImage = Boolean(o.image);
    const routine = o.routine ? String(o.routine) : "";

    el.addTopBar(s, C.SECONDARY);
    el.addBadge(s, "Vocabulary", { color: C.SECONDARY });
    el.addTitle(s, "Word Study");

    // Word banner — taller for F/Y12 to fit the larger hero size.
    const bannerH = byBand(sz, 1.35, 1.20, 1.1);
    el.addCard(s, 0.5, CONTENT_TOP, 9, bannerH, { fill: C.PRIMARY });
    s.addText(word, {
      x: 0.7, y: CONTENT_TOP + 0.10, w: 6.5, h: bannerH - 0.20,
      fontSize: byBand(sz, 48, 42, 34),
      fontFace: FONT_H, color: C.WHITE, bold: true, margin: 0,
      fit: "shrink", shrinkText: true,
    });

    // Part-of-speech pill
    if (partOfSpeech) {
      const pillH = byBand(sz, 0.46, 0.42, 0.38);
      const pillW = byBand(sz, 1.7, 1.6, 1.5);
      const pillY = CONTENT_TOP + (bannerH - pillH) / 2;
      const pillFill = C.ACCENT;
      s.addShape("roundRect", {
        x: 9.3 - pillW, y: pillY, w: pillW, h: pillH, rectRadius: 0.08,
        fill: { color: pillFill },
      });
      s.addText(partOfSpeech, {
        x: 9.3 - pillW, y: pillY, w: pillW, h: pillH,
        fontSize: sz.chip, fontFace: FONT_B, color: readableOn(pillFill), bold: true,
        align: "center", valign: "middle", margin: 0,
      });
    }

    if (hasImage) {
      const colY = CONTENT_TOP + bannerH + 0.14;
      const colH = SAFE_BOTTOM - colY;
      const imgW = 4.2;
      const imgX = 0.5;
      const textX = imgX + imgW + 0.2;
      const textW = 9.5 - textX;

      // Image container card with light border
      s.addShape("roundRect", {
        x: imgX, y: colY, w: imgW, h: colH, rectRadius: 0.08,
        fill: { color: C.BG_CARD },
        line: { color: C.MUTED, width: 0.6 },
      });
      // PptxGenJS resolves image paths at writeFile time, not addImage time,
      // so a sync try/catch here would never catch a missing-file error.
      // Let writeFile fail loudly if the supplied path is bad.
      s.addImage({
        path: o.image,
        x: imgX + 0.08, y: colY + 0.08,
        w: imgW - 0.16, h: colH - 0.16,
        sizing: { type: "contain", w: imgW - 0.16, h: colH - 0.16 },
      });

      // Right column: Definition (top), Example (mid), Routine (bottom)
      const routineH = routine ? byBand(sz, 0.6, 0.55, 0.46) : 0;
      const remainingH = colH - routineH - (routine ? 0.10 : 0);
      const defH = remainingH * 0.50;
      const exH = remainingH * 0.50 - 0.10;
      const defHdrH = byBand(sz, 0.36, 0.32, 0.28);

      el.addCard(s, textX, colY, textW, defH, { strip: C.SECONDARY, fill: C.WHITE });
      s.addText("Definition", {
        x: textX + 0.18, y: colY + 0.06, w: textW - 0.4, h: defHdrH,
        fontSize: sz.sectionLabel - 1, fontFace: FONT_B, color: C.SECONDARY, bold: true, margin: 0,
      });
      s.addText(String(definition || ""), {
        x: textX + 0.18, y: colY + defHdrH + 0.04, w: textW - 0.36, h: defH - defHdrH - 0.10,
        fontSize: sz.body - 2, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
        valign: "top", fit: "shrink", shrinkText: true,
      });

      const exY = colY + defH + 0.10;
      el.addCard(s, textX, exY, textW, exH, { strip: C.ACCENT, fill: C.BG_CARD });
      s.addText("Example", {
        x: textX + 0.18, y: exY + 0.06, w: textW - 0.4, h: defHdrH,
        fontSize: sz.sectionLabel - 1, fontFace: FONT_B, color: C.CHARCOAL, bold: true, margin: 0,
      });
      s.addText("“" + String(exampleSentence || "") + "”", {
        x: textX + 0.18, y: exY + defHdrH + 0.04, w: textW - 0.36, h: exH - defHdrH - 0.10,
        fontSize: sz.quote - 2, fontFace: FONT_H, color: C.CHARCOAL, italic: true, margin: 0,
        valign: "top", fit: "shrink", shrinkText: true,
      });

      if (routine) {
        const routineY = exY + exH + 0.10;
        el.addTextOnShape(s, routine, {
          x: textX, y: routineY, w: textW, h: routineH, rectRadius: 0.06,
          fill: { color: C.SECONDARY },
        }, {
          fontSize: sz.sectionLabel - 1, fontFace: FONT_B, color: C.WHITE, bold: true,
          align: "center", valign: "middle", margin: 0.1,
        });
      }
    } else {
      const defY = CONTENT_TOP + bannerH + 0.14;
      const defHdrH = byBand(sz, 0.40, 0.36, 0.32);
      const defH = byBand(sz, 1.65, 1.55, 1.5);
      el.addCard(s, 0.5, defY, 9, defH, { strip: C.SECONDARY, fill: C.WHITE });
      s.addText("Definition", {
        x: 0.75, y: defY + 0.08, w: 4, h: defHdrH - 0.04,
        fontSize: sz.sectionLabel, fontFace: FONT_B, color: C.SECONDARY, bold: true, margin: 0,
      });
      s.addText(definition, {
        x: 0.75, y: defY + defHdrH, w: 8.4, h: defH - defHdrH - 0.10,
        fontSize: sz.body, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
        fit: "shrink", shrinkText: true,
      });

      // Example sentence card (uses remaining space)
      const exY = defY + defH + 0.14;
      const exH = SAFE_BOTTOM - exY;
      if (exH > 0.3) {
        el.addCard(s, 0.5, exY, 9, exH, { strip: C.ACCENT, fill: C.BG_CARD });
        s.addText("Example", {
          x: 0.75, y: exY + 0.08, w: 4, h: 0.32,
          fontSize: sz.sectionLabel, fontFace: FONT_B, color: C.CHARCOAL, bold: true, margin: 0,
        });
        s.addText("“" + exampleSentence + "”", {
          x: 0.75, y: exY + 0.42, w: 8.4, h: exH - 0.54,
          fontSize: sz.quote, fontFace: FONT_H, color: C.CHARCOAL, italic: true, margin: 0,
          fit: "shrink", shrinkText: true,
        });
      }
    }

    if (footer) el.addFooter(s, footer);
    if (notes) s.addNotes(notes);
    return s;
  }

  /* ------------------------------------------------------------------ */
  /*  quoteSlide                                                         */
  /* ------------------------------------------------------------------ */

  /**
   * Quote / read-aloud slide — large quote block with optional discussion question.
   *
   * @param {object} pres       PptxGenJS presentation instance
   * @param {string} badgeText  Badge label (defaults to "Read Aloud")
   * @param {string} chapter    Chapter or section title
   * @param {string} quote      The quoted text
   * @param {string} pageRef    Page reference (e.g. "p. 12")
   * @param {string} question   Discussion question shown below the quote
   * @param {string} notes      Teacher notes
   * @param {string} footer     Footer text
   * @returns {object}          The slide object
   */
  function quoteSlide(pres, badgeText, chapter, quote, pageRef, question, notes, footer) {
    const s = pres.addSlide();
    el.addTopBar(s, C.PRIMARY);
    el.addBadge(s, badgeText || "Read Aloud", { color: C.PRIMARY });
    el.addTitle(s, chapter || "Chapter");

    // Quote card (dark background) — taller for F/Y12 to fit larger quote text.
    const quoteCardH = byBand(sz, 2.55, 2.35, 2.1);
    el.addCard(s, 0.5, CONTENT_TOP, 9, quoteCardH, { fill: C.PRIMARY });
    s.addText("“", {
      x: 0.6, y: CONTENT_TOP + 0.05, w: 0.7, h: 0.85,
      fontSize: byBand(sz, 64, 58, 52),
      fontFace: FONT_H, color: C.ACCENT, margin: 0,
    });
    s.addText(quote, {
      x: 1.2, y: CONTENT_TOP + 0.20, w: 7.6, h: quoteCardH - 0.40,
      fontSize: sz.quote, fontFace: FONT_H, color: C.WHITE, italic: true, margin: 0,
      fit: "shrink", shrinkText: true,
    });
    if (pageRef) {
      s.addText(pageRef, {
        x: 8.4, y: CONTENT_TOP + quoteCardH - 0.35, w: 1.0, h: 0.26,
        fontSize: sz.caption, fontFace: FONT_B, color: readableOn(C.PRIMARY), align: "right", margin: 0,
      });
    }

    // Discussion card (uses remaining space)
    const discY = CONTENT_TOP + quoteCardH + 0.14;
    const discH = SAFE_BOTTOM - discY;
    const discHdrH = byBand(sz, 0.36, 0.32, 0.28);
    if (question && discH > 0.3) {
      el.addCard(s, 0.5, discY, 9, discH, { strip: C.ACCENT, fill: C.WHITE });
      s.addText("Discussion", {
        x: 0.75, y: discY + 0.08, w: 4, h: discHdrH,
        fontSize: sz.sectionLabel, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
      });
      s.addText(question, {
        x: 0.75, y: discY + discHdrH + 0.10, w: 8.5, h: discH - discHdrH - 0.20,
        fontSize: sz.body, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
        fit: "shrink", shrinkText: true,
      });
    }

    if (footer) el.addFooter(s, footer);
    if (notes) s.addNotes(notes);
    return s;
  }

  /* ------------------------------------------------------------------ */
  /*  modellingSlide                                                     */
  /* ------------------------------------------------------------------ */

  /**
   * Modelling / "I Do" slide — teacher demonstration with one or two content panels.
   * When both leftContent and rightContent are provided, renders as a two-column layout.
   * When only one is provided, renders as a single full-width card.
   *
   * @param {object} pres          PptxGenJS presentation instance
   * @param {string} badgeText     Badge label (defaults to "I Do")
   * @param {string} title         Slide title
   * @param {string} leftContent   Left column text (or single-panel text)
   * @param {string} rightContent  Right column text (optional)
   * @param {string} notes         Teacher notes
   * @param {string} footer        Footer text
   * @returns {object}             The slide object
   */
  function modellingSlide(pres, badgeText, title, leftContent, rightContent, notes, footer) {
    const s = pres.addSlide();
    el.addTopBar(s, C.PRIMARY);
    el.addBadge(s, badgeText || "I Do", { color: C.PRIMARY, w: 2.2 });
    el.addTitle(s, title);

    const cardH = SAFE_BOTTOM - CONTENT_TOP;

    if (leftContent && rightContent) {
      // Two-column layout — narrow column gets bodyDense; right column same.
      el.addCard(s, 0.5, CONTENT_TOP, 4.3, cardH, { strip: C.PRIMARY, fill: C.WHITE });
      s.addText(leftContent, {
        x: 0.75, y: CONTENT_TOP + 0.14, w: 3.8, h: cardH - 0.24,
        fontSize: sz.bodyDense, fontFace: FONT_B, color: C.CHARCOAL, valign: "top", margin: 0,
        fit: "shrink", shrinkText: true,
      });

      el.addCard(s, 5.0, CONTENT_TOP, 4.5, cardH, { strip: C.ACCENT, fill: C.BG_CARD });
      s.addText(rightContent, {
        x: 5.2, y: CONTENT_TOP + 0.14, w: 4.1, h: cardH - 0.24,
        fontSize: sz.bodyDense, fontFace: FONT_H, color: C.CHARCOAL, valign: "top", italic: true, margin: 0,
        fit: "shrink", shrinkText: true,
      });
    } else {
      // Single full-width card
      const content = leftContent || rightContent || "";
      el.addCard(s, 0.5, CONTENT_TOP, 9, cardH, { strip: C.PRIMARY, fill: C.WHITE });
      s.addText(content, {
        x: 0.75, y: CONTENT_TOP + 0.14, w: 8.5, h: cardH - 0.24,
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
   * Pair-Share discussion slide — evenly spaced question cards with alternating strip colours.
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

    // F/Y12 caps the displayed question count to honour the megaprompt's
    // "1 question per slide" rule for younger students.
    const _qs = questions || [];
    const cappedQuestions = _qs.slice(0, sz.maxQuestions || _qs.length);
    const availH = SAFE_BOTTOM - CONTENT_TOP;
    const qHMax = byBand(sz, 2.6, 1.8, 0.95);
    const qH = Math.min(qHMax, (availH - 0.1) / Math.max(cappedQuestions.length, 1));
    const fontSize = cappedQuestions.length > (sz.maxQuestions || 3) ? sz.bodyDense : sz.body + 1;

    cappedQuestions.forEach((q, i) => {
      const y = CONTENT_TOP + i * (qH + 0.10);
      if (y + qH > SAFE_BOTTOM) return;
      el.addCard(s, 0.5, y, 9, qH, { strip: i % 2 === 0 ? C.PRIMARY : C.SECONDARY, fill: C.WHITE });
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

  // NOTE: annotatedModelSlide is intentionally NOT overridden here.
  // The base builder's version (from base.js) supports both previewBlocks
  // and structured previewSpec (poster specs via drawMockupPreview/drawPosterSpec).
  // A stale literacy fork that only supported previewBlocks was removed in
  // March 2026. If literacy needs genuinely different annotatedModelSlide
  // behaviour in the future, add it here with a comment explaining why
  // the base version is insufficient.

  /* ------------------------------------------------------------------ */
  /*  Return all builders                                                */
  /* ------------------------------------------------------------------ */

  return { vocabSlide, quoteSlide, modellingSlide, pairShareSlide };
}

module.exports = { createLiteracyBuilders };
