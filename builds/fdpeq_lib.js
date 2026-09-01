"use strict";

// Shared visual layer for the Year 5/6 "Fractions, Decimals & Percentages"
// 2-session unit (build_fdpeq_lesson1.js, build_fdpeq_lesson2.js).
//
// It exists so the UNIT ANCHOR is literally the same object in both sessions
// (megaprompt section 79): one whole, one line, three ways to name the same
// point. A helper defined once cannot drift between Session 1 and Session 2.
//
// UNIT ANCHOR (locked wording, used verbatim on slides, in notes and in PDFs):
//   "Same whole. Same line. Then compare."
//     1. Same whole   - check both amounts are parts of the SAME whole.
//     2. Same line    - put them on one line, or rename to the same bottom number.
//     3. Then compare - and say how you know.

const ANCHOR_MOVES = ["1.  Same whole", "2.  Same line", "3.  Then compare"];
const ANCHOR_PHRASE = "Same whole. Same line. Then compare.";

function createLib(T, FOOTER) {
  const {
    C, FONT_H, FONT_B,
    SAFE_BOTTOM, CONTENT_TOP,
    addCard, addFooter, addTopBar, addTitle, addBadge, addTextOnShape,
    addRevealAnswerBar, clickBuild, runSlideDiagnostics,
  } = T;

  /* -- The unit anchor: three move chips ---------------------------------- */

  // Drawn identically in both sessions. Session 1 introduces it, Session 2
  // restates it, and every I Do says the phrase before extending it.
  function anchorMoves(slide, y, opts) {
    const o = opts || {};
    const h = o.h || 0.55;
    ANCHOR_MOVES.forEach((m, i) => {
      addTextOnShape(slide, m, {
        x: 0.5 + i * 3.05, y, w: 2.9, h, rectRadius: 0.08,
        fill: { color: o.color || C.PRIMARY },
      }, { fontSize: o.fontSize || 16, fontFace: FONT_H, color: C.WHITE, bold: true });
    });
    return y + h;
  }

  /* -- The unit anchor visual: one bar, three ways to name each point ------ */

  // A single whole split into quarters, with the percent, decimal and fraction
  // name of the same five points stacked underneath. This is the "same line"
  // half of the anchor and it reappears on paper in the worksheets.
  function tripleScale(slide, x, y, w, opts) {
    const o = opts || {};
    const barH = o.barH || 0.60;
    const cells = 4;
    const cellW = w / cells;
    const shaded = Math.max(0, Math.min(cells, o.shaded || 0));

    for (let i = 0; i < cells; i += 1) {
      slide.addShape("rect", {
        x: x + i * cellW, y, w: cellW, h: barH,
        fill: { color: i < shaded ? C.SECONDARY : C.WHITE },
        line: { color: C.CHARCOAL, width: 1.25 },
      });
    }

    const stacks = [
      { pct: "0%", dec: "0", frac: "0" },
      { pct: "25%", dec: "0.25", frac: "1/4" },
      { pct: "50%", dec: "0.5", frac: "1/2" },
      { pct: "75%", dec: "0.75", frac: "3/4" },
      { pct: "100%", dec: "1", frac: "1 whole" },
    ];
    const labelW = o.labelW || 1.55;
    const labelY = y + barH + 0.12;
    const labelH = o.labelH || 0.92;

    stacks.forEach((st, i) => {
      const cx = x + i * cellW;
      // Tick sits in the 0.12" band between the bar and the label stack.
      slide.addShape("rect", {
        x: cx - 0.012, y: y + barH, w: 0.024, h: 0.09,
        fill: { color: C.CHARCOAL },
      });
      slide.addText([
        { text: st.pct, options: { color: C.PRIMARY, bold: true, fontSize: 17, breakLine: true } },
        { text: st.dec, options: { color: C.CHARCOAL, bold: true, fontSize: 15, breakLine: true } },
        { text: st.frac, options: { color: C.SECONDARY, bold: true, fontSize: 15 } },
      ], {
        x: cx - labelW / 2, y: labelY, w: labelW, h: labelH,
        fontFace: FONT_B, align: "center", valign: "top", margin: 0,
      });
    });

    return labelY + labelH;
  }

  /* -- Fluency slide with a click-revealed answer bar ---------------------- */

  // The shared fluencySlide grows its number cards to y 4.60, which the answer
  // bar (top 4.30) would then sit on top of. Same look, cards stopped at 4.15
  // so the reveal has real clearance.
  function fluencySlide(pres, title, prompts, answers, notes) {
    const s = pres.addSlide();
    addTopBar(s, C.ACCENT);
    addBadge(s, "Fluency", { color: C.ACCENT });
    addTitle(s, title, { color: C.ACCENT });

    const cardTop = CONTENT_TOP;
    const cardBottom = 4.15;
    const gap = 0.18;
    const cardW = (9 - gap * (prompts.length - 1)) / prompts.length;

    prompts.forEach((q, i) => {
      const x = 0.5 + i * (cardW + gap);
      addCard(s, x, cardTop, cardW, cardBottom - cardTop, { strip: C.ACCENT });
      s.addText(String(q), {
        x: x + 0.16, y: cardTop + 0.16, w: cardW - 0.32, h: cardBottom - cardTop - 0.32,
        fontSize: 40, fontFace: FONT_H, color: C.CHARCOAL,
        bold: true, align: "center", valign: "middle", margin: 0,
        fit: "shrink", shrinkText: true,
      });
    });

    addFooter(s, FOOTER);
    s.addNotes(notes);
    clickBuild(s, [
      () => { addRevealAnswerBar(s, answers, { color: C.SUCCESS, y: 4.30, h: 0.78, fontSize: 26 }); },
    ]);
    runSlideDiagnostics(s, pres);
    return s;
  }

  /* -- Check for understanding, with room for the revealed answer ---------- */

  // The shared cfuSlide sizes its question card to fill the content area, so a
  // click-revealed answer bar always lands on top of it. This keeps the same
  // visual language and reserves the bottom band.
  function checkSlide(pres, cfg) {
    const s = pres.addSlide();
    addTopBar(s, C.ALERT);
    addBadge(s, "CFU", { color: C.ALERT });
    addTitle(s, cfg.title, { color: C.ALERT });

    addTextOnShape(s, "Check", {
      x: 8.35, y: 0.20, w: 1.15, h: 0.36, rectRadius: 0.08,
      fill: { color: C.WHITE }, line: { color: C.ALERT, width: 1.5 },
    }, {
      fontSize: 12, fontFace: FONT_B, color: C.ALERT,
      bold: true, align: "center", valign: "middle", margin: 0,
    });

    addTextOnShape(s, cfg.technique || "Show Me Boards", {
      x: 0.5, y: CONTENT_TOP, w: 2.9, h: 0.44, rectRadius: 0.08,
      fill: { color: C.ALERT },
    }, { fontSize: 13.5, fontFace: FONT_B, color: C.WHITE, bold: true });

    const qY = CONTENT_TOP + 0.58;
    const options = cfg.options || [];
    // With no options the question card grows into the space the option row
    // would have used, so the slide never ends on a dead band.
    const qH = options.length ? 1.20 : 2.30;
    addCard(s, 0.5, qY, 9, qH, { strip: C.ALERT });
    if (cfg.lead) {
      s.addText(cfg.lead, {
        x: 0.75, y: qY + 0.10, w: 8.5, h: 0.28,
        fontSize: 14, fontFace: FONT_B, color: C.MUTED, bold: true, margin: 0,
      });
    }
    s.addText(cfg.question, {
      x: 0.75, y: qY + (cfg.lead ? 0.40 : 0.14), w: 8.5, h: qH - (cfg.lead ? 0.52 : 0.28),
      fontSize: cfg.questionSize || 30, fontFace: FONT_B, color: C.CHARCOAL,
      bold: true, align: "center", valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    if (options.length) {
      const optY = qY + qH + 0.20;
      const optGap = 0.16;
      const optW = (9 - optGap * (options.length - 1)) / options.length;
      options.forEach((opt, i) => {
        addTextOnShape(s, opt, {
          x: 0.5 + i * (optW + optGap), y: optY, w: optW, h: 0.62, rectRadius: 0.08,
          fill: { color: C.WHITE }, line: { color: C.PRIMARY, width: 1.5 },
        }, {
          fontSize: 15, fontFace: FONT_B, color: C.PRIMARY, bold: true,
        });
      });
    }

    if (cfg.drawExtra) cfg.drawExtra(s);

    addFooter(s, FOOTER);
    s.addNotes(cfg.notes);
    clickBuild(s, [
      () => {
        addRevealAnswerBar(s, cfg.answer, {
          color: cfg.answerColor || C.SUCCESS, label: "Answer",
          fontSize: cfg.answerSize || 22, y: 4.28, h: 0.80,
        });
      },
    ]);
    runSlideDiagnostics(s, pres);
    return s;
  }

  /* -- You Do task-setup slide -------------------------------------------- */

  function youDoSlide(pres, cfg) {
    const stageColor = T.STAGE_COLORS["4"];
    const s = pres.addSlide();
    addTopBar(s, stageColor);
    T.addStageBadge(s, 4, "You Do");
    addTitle(s, cfg.title, { color: stageColor });

    addCard(s, 0.5, CONTENT_TOP, 9.0, 1.15, { strip: stageColor });
    const runs = [];
    cfg.steps.forEach((st) => {
      runs.push({ text: st.label + " ", options: { fontSize: 19, color: C.ALERT, bold: true } });
      runs.push({ text: st.text + "  ", options: { fontSize: 16.5, color: C.CHARCOAL } });
    });
    s.addText(runs, {
      x: 0.75, y: CONTENT_TOP + 0.14, w: 8.5, h: 0.87,
      fontFace: FONT_B, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    const panelY = CONTENT_TOP + 1.32;
    addCard(s, 0.5, panelY, 9.0, SAFE_BOTTOM - panelY, { strip: C.SECONDARY });
    s.addText(cfg.panelTitle, {
      x: 0.7, y: panelY + 0.14, w: 8.6, h: 0.30,
      fontSize: 16, fontFace: FONT_H, color: C.SECONDARY, bold: true,
      align: "center", margin: 0,
    });
    cfg.reminders.forEach((r, i) => {
      addTextOnShape(s, r.text, {
        x: 1.0, y: panelY + 0.60 + i * 0.58, w: 8.0, h: 0.46, rectRadius: 0.08,
        fill: { color: r.color },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
    });

    addFooter(s, FOOTER);
    s.addNotes(cfg.notes);
    runSlideDiagnostics(s, pres);
    return s;
  }

  return { anchorMoves, tripleScale, fluencySlide, checkSlide, youDoSlide };
}

module.exports = { createLib, ANCHOR_MOVES, ANCHOR_PHRASE };
