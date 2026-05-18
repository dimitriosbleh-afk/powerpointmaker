"use strict";

/**
 * Sentence Combining Warm-Up - Year 5/6 Enrichment, Term 2 Week 5, Day 2 (Tue)
 * Subordinating conjunction: WHEN (time / sequence).
 * Topic: specialist subjects (music, art, PE).
 *
 * Day 2 is content-only - no title or framing slides (those live in Lesson 1).
 *
 * Daily routine (15 min):
 *   Review BECAUSE (2)  -> Teach WHEN (3) -> Guided reveal pair (4)
 *   -> Independent reveal pair, stem hidden (4) -> Exit check with QR (2).
 */

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant, composeNotes } = require("../themes/factory");

const T = createTheme("literacy", "grade56", weekToVariant(4));
const {
  C, FONT_H, FONT_B,
  SLIDE_W, SLIDE_H, SAFE_BOTTOM, CONTENT_TOP,
  addCard, addFooter, addTopBar, addBadge, addTitle, addTextOnShape,
  withReveal, runSlideDiagnostics,
} = T;

const FOOTER = "Sentence Combining Warm-Up | Year 5/6 | Term 2 Week 5 | Day 2 Tue";
const OUT_DIR = "output/SC_T2W5_Lesson2_Tue_WHEN";
fs.mkdirSync(OUT_DIR, { recursive: true });

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Sentence Combining Warm-Up - Year 5/6 - Term 2 Week 5 Day 2 (Tue)";
pres.author = "Year 5/6 Enrichment PLC";

const SC_BASE = (joiningWord) => ([
  "I can find two short ideas that go together.",
  `I can use ${joiningWord} to join them into ONE sentence.`,
  "I can write a sentence that makes sense.",
]);

const MARK_CRITERIA = (joiningWord) => ([
  "Two ideas combined into ONE sentence",
  `Used ${joiningWord} where it fits the meaning`,
  "Sentence is grammatical (no run-on, no comma splice)",
]);

function badgeStack(s, phaseText, phaseColor, dayText, dayColor) {
  addBadge(s, phaseText, { color: phaseColor, w: 1.6, x: 0.5, y: 0.20 });
  addBadge(s, dayText, { color: dayColor || C.SECONDARY, w: 1.05, x: 2.18, y: 0.20 });
}

function highlightSentence(sentence, joiningWord, baseSize, baseColor, highlightColor) {
  const re = new RegExp(`(\\s|^)(${joiningWord})(\\s|[.,?!])`, "i");
  const m = sentence.match(re);
  if (!m) {
    return [{ text: sentence, options: { fontSize: baseSize, color: baseColor } }];
  }
  const before = sentence.slice(0, m.index + m[1].length);
  const word   = m[2];
  const after  = sentence.slice(m.index + m[1].length + word.length);
  return [
    { text: before, options: { fontSize: baseSize, color: baseColor } },
    { text: word,   options: { fontSize: baseSize, color: highlightColor, bold: true } },
    { text: after,  options: { fontSize: baseSize, color: baseColor } },
  ];
}

function dayIntroSlide({ dayText, dayColor, joiningWord, hookHeader, hookBody, scItems, notes }) {
  const s = pres.addSlide();
  addTopBar(s, C.PRIMARY);
  badgeStack(s, "Day start", C.PRIMARY, dayText, dayColor);
  addTitle(s, `${dayText} - Joining word: ${joiningWord}`);

  const cardH = SAFE_BOTTOM - CONTENT_TOP;
  const leftW = 4.6;
  const rightW = 4.4;
  const rightX = 5.2;

  addCard(s, 0.5, CONTENT_TOP, leftW, cardH, { strip: C.SECONDARY, fill: C.WHITE });
  s.addText(hookHeader, {
    x: 0.7, y: CONTENT_TOP + 0.14, w: leftW - 0.4, h: 0.34,
    fontSize: 13, fontFace: FONT_B, color: C.SECONDARY, bold: true, margin: 0,
  });
  s.addText(hookBody, {
    x: 0.7, y: CONTENT_TOP + 0.52, w: leftW - 0.4, h: cardH - 0.66,
    fontSize: 17, fontFace: FONT_H, color: C.CHARCOAL, valign: "top", margin: 0,
    fit: "shrink", shrinkText: true,
  });

  addCard(s, rightX, CONTENT_TOP, rightW, cardH, { strip: C.ACCENT, fill: C.BG_CARD });
  s.addText("Today's success criteria - I can...", {
    x: rightX + 0.2, y: CONTENT_TOP + 0.14, w: rightW - 0.4, h: 0.34,
    fontSize: 13, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
  });
  s.addText(scItems.map((t, i) => ({
    text: t,
    options: {
      bullet: true,
      breakLine: i < scItems.length - 1,
      fontSize: 15, color: C.CHARCOAL, paraSpaceAfter: 6,
    },
  })), {
    x: rightX + 0.2, y: CONTENT_TOP + 0.54, w: rightW - 0.4, h: cardH - 0.68,
    fontFace: FONT_B, valign: "top", margin: 0,
    fit: "shrink", shrinkText: true,
  });

  addFooter(s, FOOTER);
  if (notes) s.addNotes(notes);
  runSlideDiagnostics(s, pres, { respectSafeBottom: false });
  return s;
}

function teachSlide({ dayText, dayColor, joiningWord, meaning, twoIdeas, combined, whyItWorks, nonExample, notes }) {
  const s = pres.addSlide();
  addTopBar(s, C.PRIMARY);
  badgeStack(s, "Teach", C.PRIMARY, dayText, dayColor);
  addTitle(s, `How "${joiningWord}" works - ${meaning}`);

  const workedY = CONTENT_TOP;
  const workedH = 2.55;
  addCard(s, 0.5, workedY, 9, workedH, { strip: C.PRIMARY, fill: C.WHITE });

  s.addText("Worked example", {
    x: 0.75, y: workedY + 0.10, w: 4, h: 0.28,
    fontSize: 12, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
  });
  s.addText("Two short ideas:", {
    x: 0.75, y: workedY + 0.42, w: 4, h: 0.26,
    fontSize: 11, fontFace: FONT_B, color: C.MUTED, bold: true, margin: 0,
  });
  const ideaY = workedY + 0.70;
  s.addText(`"${twoIdeas[0]}"`, {
    x: 0.85, y: ideaY, w: 8.0, h: 0.34,
    fontSize: 17, fontFace: FONT_H, color: C.CHARCOAL, italic: true, margin: 0,
    fit: "shrink", shrinkText: true,
  });
  s.addText(`"${twoIdeas[1]}"`, {
    x: 0.85, y: ideaY + 0.36, w: 8.0, h: 0.34,
    fontSize: 17, fontFace: FONT_H, color: C.CHARCOAL, italic: true, margin: 0,
    fit: "shrink", shrinkText: true,
  });
  s.addText("v", {
    x: 0.75, y: ideaY + 0.74, w: 0.6, h: 0.32,
    fontSize: 18, fontFace: FONT_H, color: C.ACCENT, bold: true, margin: 0,
  });
  s.addText(`Joined with "${joiningWord}":`, {
    x: 1.45, y: ideaY + 0.74, w: 4.9, h: 0.32,
    fontSize: 11, fontFace: FONT_B, color: C.MUTED, bold: true, margin: 0,
  });
  const combinedRuns = highlightSentence(combined, joiningWord, 19, C.CHARCOAL, C.ACCENT);
  s.addText(combinedRuns.map((run, i) => ({
    text: `${i === 0 ? '"' : ""}${run.text}${i === combinedRuns.length - 1 ? '"' : ""}`,
    options: { ...run.options, fontFace: FONT_H, breakLine: false },
  })), {
    x: 0.85, y: ideaY + 1.10, w: 8.0, h: 0.42,
    valign: "middle", margin: 0,
    fit: "shrink", shrinkText: true,
  });

  const whyY = workedY + workedH - 0.34;
  s.addShape("rect", {
    x: 0.5, y: whyY - 0.02, w: 9, h: 0.36, fill: { color: C.BG_CARD },
  });
  s.addText("Why this works: " + whyItWorks, {
    x: 0.75, y: whyY, w: 8.5, h: 0.32,
    fontSize: 12, fontFace: FONT_B, color: C.SECONDARY, bold: true, margin: 0,
    fit: "shrink", shrinkText: true,
  });

  const neY = workedY + workedH + 0.14;
  const neH = SAFE_BOTTOM - neY;
  addCard(s, 0.5, neY, 9, neH, { strip: C.ALERT, fill: C.WHITE });
  s.addText("Watch out - common mistake", {
    x: 0.75, y: neY + 0.10, w: 5, h: 0.28,
    fontSize: 12, fontFace: FONT_B, color: C.ALERT, bold: true, margin: 0,
  });
  s.addText(`Wrong: "${nonExample.wrong}"`, {
    x: 0.75, y: neY + 0.40, w: 8.5, h: 0.34,
    fontSize: 14.5, fontFace: FONT_H, color: C.CHARCOAL, italic: true, margin: 0,
    fit: "shrink", shrinkText: true,
  });
  s.addText("Why it's wrong: " + nonExample.why, {
    x: 0.75, y: neY + 0.74, w: 8.5, h: neH - 0.84,
    fontSize: 12, fontFace: FONT_B, color: C.MUTED, valign: "top", margin: 0,
    fit: "shrink", shrinkText: true,
  });

  addFooter(s, FOOTER);
  if (notes) s.addNotes(notes);
  runSlideDiagnostics(s, pres, { respectSafeBottom: false });
  return s;
}

function guidedSlide({ dayText, dayColor, joiningWord, pairs, answers, notesBuild, notesReveal }) {
  return withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.SECONDARY);
      badgeStack(s, "Practise together", C.SECONDARY, dayText, dayColor);
      addTitle(s, "Combine each pair into ONE sentence");

      addTextOnShape(s, "Mini-whiteboards", {
        x: 7.4, y: 0.20, w: 2.1, h: 0.36, rectRadius: 0.06,
        fill: { color: C.WHITE },
        line: { color: C.SECONDARY, width: 1.0 },
      }, {
        fontSize: 11, fontFace: FONT_B, color: C.SECONDARY, bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      const cardW = 9;
      const gap = 0.16;
      const pairsArea = 2.55;
      const cardH = (pairsArea - gap) / 2;
      pairs.forEach((pair, i) => {
        const y = CONTENT_TOP + i * (cardH + gap);
        addCard(s, 0.5, y, cardW, cardH, {
          strip: i === 0 ? C.PRIMARY : C.ACCENT,
          fill: C.WHITE,
        });
        s.addText(`Try this - pair ${i + 1}`, {
          x: 0.75, y: y + 0.08, w: 5, h: 0.26,
          fontSize: 11, fontFace: FONT_B,
          color: i === 0 ? C.PRIMARY : C.ACCENT, bold: true, margin: 0,
        });
        s.addText(`"${pair[0]}"`, {
          x: 0.75, y: y + 0.34, w: 8.4, h: 0.34,
          fontSize: 16, fontFace: FONT_H, color: C.CHARCOAL, italic: true, margin: 0,
          fit: "shrink", shrinkText: true,
        });
        s.addText(`"${pair[1]}"`, {
          x: 0.75, y: y + 0.70, w: 8.4, h: 0.34,
          fontSize: 16, fontFace: FONT_H, color: C.CHARCOAL, italic: true, margin: 0,
          fit: "shrink", shrinkText: true,
        });
      });

      const stubY = CONTENT_TOP + pairsArea + 0.10;
      const stubH = SAFE_BOTTOM - stubY;
      s.addShape("roundRect", {
        x: 0.5, y: stubY, w: 9, h: stubH, rectRadius: 0.10,
        fill: { color: C.BG_CARD },
        line: { color: C.MUTED, width: 0.8, dashType: "dash" },
      });
      s.addText(`Whiteboards FIRST, then hold up, then click to reveal "${joiningWord}" answers.`, {
        x: 0.75, y: stubY, w: 8.5, h: stubH,
        fontSize: 13, fontFace: FONT_B, color: C.MUTED, italic: true,
        align: "center", valign: "middle", margin: 0,
      });

      addFooter(s, FOOTER);
      if (notesBuild) s.addNotes(notesBuild);
      runSlideDiagnostics(s, pres, { respectSafeBottom: false });
      return s;
    },
    (s) => {
      const stubY = CONTENT_TOP + 2.55 + 0.10;
      const stubH = SAFE_BOTTOM - stubY;
      s.addShape("roundRect", {
        x: 0.5, y: stubY, w: 9, h: stubH, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      });
      s.addText("Answers", {
        x: 0.75, y: stubY + 0.06, w: 3, h: 0.24,
        fontSize: 11, fontFace: FONT_B, color: C.WHITE, bold: true, margin: 0,
      });
      const ansListY = stubY + 0.30;
      const perAnsH = (stubH - 0.36) / 2;
      answers.forEach((ans, i) => {
        const y = ansListY + i * perAnsH;
        const runs = highlightSentence(ans, joiningWord, 14, C.WHITE, C.BG_CARD);
        s.addText(runs.map((run, j) => ({
          text: `${j === 0 ? `${i + 1}.  "` : ""}${run.text}${j === runs.length - 1 ? '"' : ""}`,
          options: { ...run.options, fontFace: FONT_H, bold: true, breakLine: false },
        })), {
          x: 0.75, y, w: 8.5, h: perAnsH - 0.05,
          valign: "middle", margin: 0,
          fit: "shrink", shrinkText: true,
        });
      });
      if (notesReveal) s.addNotes(notesReveal);
    }
  );
}

function independentSlide({ dayText, dayColor, joiningWord, prompt, sentenceStarter, notesBuild, notesReveal }) {
  return withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ACCENT);
      badgeStack(s, "Your turn", C.ACCENT, dayText, dayColor);
      addTitle(s, `Write ONE sentence using "${joiningWord}"`);

      addTextOnShape(s, "In your warm-up book", {
        x: 6.85, y: 0.20, w: 2.65, h: 0.36, rectRadius: 0.06,
        fill: { color: C.WHITE },
        line: { color: C.ACCENT, width: 1.0 },
      }, {
        fontSize: 11, fontFace: FONT_B, color: C.ACCENT, bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      const promptY = CONTENT_TOP;
      const promptH = 2.55;
      addCard(s, 0.5, promptY, 9, promptH, { strip: C.ACCENT, fill: C.WHITE });
      s.addText("Topic prompt", {
        x: 0.75, y: promptY + 0.12, w: 5, h: 0.28,
        fontSize: 12, fontFace: FONT_B, color: C.ACCENT, bold: true, margin: 0,
      });
      s.addText(prompt, {
        x: 0.75, y: promptY + 0.46, w: 8.5, h: promptH - 0.60,
        fontSize: 22, fontFace: FONT_H, color: C.CHARCOAL, valign: "middle", margin: 0,
        fit: "shrink", shrinkText: true,
      });

      const remindY = promptY + promptH + 0.14;
      const remindH = SAFE_BOTTOM - remindY;
      s.addShape("roundRect", {
        x: 0.5, y: remindY, w: 9, h: remindH, rectRadius: 0.08,
        fill: { color: C.PRIMARY },
      });
      s.addText(`Remember: ONE sentence, TWO ideas joined by "${joiningWord}". Write in your warm-up book.`, {
        x: 0.75, y: remindY, w: 8.5, h: remindH,
        fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      addFooter(s, FOOTER);
      if (notesBuild) s.addNotes(notesBuild);
      runSlideDiagnostics(s, pres, { respectSafeBottom: false });
      return s;
    },
    (s) => {
      const promptY = CONTENT_TOP;
      const promptH = 1.95;
      const stemY = promptY + promptH + 0.14;
      const stemH = 0.62;
      s.addShape("roundRect", {
        x: 0.5, y: stemY, w: 9, h: stemH, rectRadius: 0.08,
        fill: { color: C.BG_CARD },
        line: { color: C.SECONDARY, width: 0.9 },
      });
      s.addText("Need a start? Try: ", {
        x: 0.75, y: stemY, w: 2.0, h: stemH,
        fontSize: 12, fontFace: FONT_B, color: C.SECONDARY, bold: true,
        valign: "middle", margin: 0,
      });
      s.addText(`"${sentenceStarter}"`, {
        x: 2.75, y: stemY, w: 6.55, h: stemH,
        fontSize: 14.5, fontFace: FONT_H, color: C.CHARCOAL, italic: true,
        valign: "middle", margin: 0, fit: "shrink",
      });
      if (notesReveal) s.addNotes(notesReveal);
    }
  );
}

function exitSlide({ dayText, dayColor, joiningWord, taskHeader, taskIdeas, markCriteria, qrLabel, notes }) {
  const s = pres.addSlide();
  addTopBar(s, C.ALERT);
  badgeStack(s, "Exit check", C.ALERT, dayText, dayColor);

  addTextOnShape(s, "EXIT", {
    x: 8.0, y: 0.20, w: 1.5, h: 0.36, rectRadius: 0.06,
    fill: { color: C.WHITE },
    line: { color: C.ALERT, width: 1.5 },
  }, {
    fontSize: 11, fontFace: FONT_B, color: C.ALERT, bold: true,
    align: "center", valign: "middle", margin: 0,
  });

  addTitle(s, "Show what you know", { w: 7.4 });

  const cardH = SAFE_BOTTOM - CONTENT_TOP;
  const leftW = 3.7;
  const midX = 4.4;
  const midW = 2.0;
  const rightX = 6.55;
  const rightW = 2.95;

  addCard(s, 0.5, CONTENT_TOP, leftW, cardH, { strip: C.ALERT, fill: C.WHITE });
  s.addText(taskHeader, {
    x: 0.7, y: CONTENT_TOP + 0.12, w: leftW - 0.4, h: 0.30,
    fontSize: 12, fontFace: FONT_B, color: C.ALERT, bold: true, margin: 0,
  });
  const ideaY = CONTENT_TOP + 0.46;
  s.addText(`"${taskIdeas[0]}"`, {
    x: 0.7, y: ideaY, w: leftW - 0.4, h: 0.46,
    fontSize: 15.5, fontFace: FONT_H, color: C.CHARCOAL, italic: true, margin: 0,
    fit: "shrink", shrinkText: true,
  });
  s.addText(`"${taskIdeas[1]}"`, {
    x: 0.7, y: ideaY + 0.52, w: leftW - 0.4, h: 0.46,
    fontSize: 15.5, fontFace: FONT_H, color: C.CHARCOAL, italic: true, margin: 0,
    fit: "shrink", shrinkText: true,
  });
  s.addText(`Combine into ONE sentence using "${joiningWord}".`, {
    x: 0.7, y: ideaY + 1.10, w: leftW - 0.4, h: 0.40,
    fontSize: 12.5, fontFace: FONT_B, color: C.SECONDARY, bold: true, margin: 0,
    fit: "shrink", shrinkText: true,
  });
  s.addText("Submit via QR (or paper slip).", {
    x: 0.7, y: CONTENT_TOP + cardH - 0.36, w: leftW - 0.4, h: 0.26,
    fontSize: 10.5, fontFace: FONT_B, color: C.MUTED, italic: true, margin: 0,
  });

  addCard(s, midX, CONTENT_TOP, midW, cardH, { strip: C.PRIMARY, fill: C.BG_CARD });
  s.addText("Scan to submit", {
    x: midX + 0.12, y: CONTENT_TOP + 0.12, w: midW - 0.24, h: 0.28,
    fontSize: 11, fontFace: FONT_B, color: C.PRIMARY, bold: true,
    align: "center", margin: 0,
  });
  const qrSize = 1.35;
  const qrX = midX + (midW - qrSize) / 2;
  const qrY = CONTENT_TOP + 0.46;
  s.addShape("rect", {
    x: qrX, y: qrY, w: qrSize, h: qrSize,
    fill: { color: C.WHITE },
    line: { color: C.CHARCOAL, width: 1.0 },
  });
  const fp = 0.28;
  const inset = 0.10;
  [[qrX + inset, qrY + inset], [qrX + qrSize - inset - fp, qrY + inset], [qrX + inset, qrY + qrSize - inset - fp]].forEach(([fx, fy]) => {
    s.addShape("rect", { x: fx, y: fy, w: fp, h: fp, fill: { color: C.CHARCOAL } });
    s.addShape("rect", { x: fx + 0.06, y: fy + 0.06, w: fp - 0.12, h: fp - 0.12, fill: { color: C.WHITE } });
    s.addShape("rect", { x: fx + 0.10, y: fy + 0.10, w: fp - 0.20, h: fp - 0.20, fill: { color: C.CHARCOAL } });
  });
  s.addText("QR", {
    x: qrX + qrSize - 0.45, y: qrY + qrSize - 0.40, w: 0.40, h: 0.32,
    fontSize: 11, fontFace: FONT_B, color: C.CHARCOAL, bold: true,
    align: "center", valign: "middle", margin: 0,
  });
  s.addText(qrLabel, {
    x: midX + 0.10, y: qrY + qrSize + 0.10, w: midW - 0.20, h: 0.34,
    fontSize: 10.5, fontFace: FONT_B, color: C.CHARCOAL,
    align: "center", margin: 0, fit: "shrink",
  });
  s.addText("Paper backup in basket.", {
    x: midX + 0.10, y: CONTENT_TOP + cardH - 0.36, w: midW - 0.20, h: 0.26,
    fontSize: 9.5, fontFace: FONT_B, color: C.MUTED, italic: true,
    align: "center", margin: 0,
  });

  addCard(s, rightX, CONTENT_TOP, rightW, cardH, { strip: C.SUCCESS, fill: C.BG_CARD });
  s.addText("Tick if students...", {
    x: rightX + 0.16, y: CONTENT_TOP + 0.12, w: rightW - 0.32, h: 0.30,
    fontSize: 12, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
  });
  const rowStart = CONTENT_TOP + 0.46;
  const rowH = 0.66;
  markCriteria.forEach((crit, i) => {
    const y = rowStart + i * (rowH + 0.08);
    s.addShape("roundRect", {
      x: rightX + 0.16, y: y + 0.10, w: 0.30, h: 0.30, rectRadius: 0.04,
      fill: { color: C.WHITE },
      line: { color: C.SUCCESS, width: 1.2 },
    });
    s.addText(crit, {
      x: rightX + 0.52, y, w: rightW - 0.66, h: rowH,
      fontSize: 11, fontFace: FONT_B, color: C.CHARCOAL, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });
  });
  s.addText("Tracker: ticks + which conjunction.", {
    x: rightX + 0.16, y: CONTENT_TOP + cardH - 0.40, w: rightW - 0.32, h: 0.30,
    fontSize: 10, fontFace: FONT_B, color: C.MUTED, italic: true, margin: 0,
    fit: "shrink",
  });

  addFooter(s, FOOTER);
  if (notes) s.addNotes(notes);
  runSlideDiagnostics(s, pres, { respectSafeBottom: false });
  return s;
}

/* =========================================================================
 *  DAY 2 - TUESDAY - WHEN (time / sequence) - topic: specialist subjects
 * ========================================================================= */

const TUE = { dayText: "Tue", dayColor: C.SECONDARY };

dayIntroSlide({
  dayText: TUE.dayText, dayColor: TUE.dayColor,
  joiningWord: "WHEN",
  hookHeader: "Hook - 30 seconds, then a recall",
  hookBody:
    "On your whiteboard, write TWO short ideas about music, art or PE class. " +
    "Two short sentences with full stops. Hold onto these - we will combine them at the end of the session. " +
    "Then quick recall: tell your partner ONE sentence using BECAUSE.",
  scItems: SC_BASE("WHEN"),
  notes: composeNotes({
    say: [
      "Two jobs to start. First - on your whiteboards, write TWO short ideas about a specialist class. Music, art or PE. Two short sentences. We will come back to these at the end.",
      "Second - quick recall. Tell your partner ONE sentence using yesterday's word, BECAUSE.",
      "Today's new word is WHEN. WHEN tells us the TIME - when something happens.",
    ],
    do: [
      "Set a 90-second timer covering both jobs - 60 seconds for the two-idea hook, 30 seconds for the partner recall.",
      "Scan whiteboards as you walk - keep the boards on desks for the rest of the lesson.",
      "Pick one strong BECAUSE example from the partner share to read aloud, then introduce WHEN.",
    ],
    teacherNotes: [
      "Bookend pattern (Teacher 2's idea from Week 4) carried into Tuesday: hook writes two short ideas; independent later combines those same ideas with WHEN. Students see the 15-minute growth on their own board.",
      "Topic is specialist subjects (music, art, PE) - every student attends specialists, universal access, level playing field.",
      "Today supports rubric criterion 3 (appropriate clause linking) and criterion 5 (no comma splice).",
    ],
    enabling: [
      "FOCUS GROUP: if they cannot generate two ideas, give them oral options - 'we sit on the floor' / 'the music teacher claps three times' - and just ask them to write those.",
      "EXTENSION: ask higher-ability students to predict what 'time' means in a sentence and which other words (WHILE, AS, AFTER, BEFORE) signal a similar meaning.",
    ],
    watchFor: [
      "Students writing ONE long sentence instead of TWO short ones - redirect: 'two short ideas, two full stops'.",
      "Comma splice creeping back in - frame as developmental progress, not failure.",
    ],
    tag: "[Day 2 | Tue | Hook + recall + LI/SC for WHEN]",
  }, { requireSay: false, requireDo: false }),
});

teachSlide({
  dayText: TUE.dayText, dayColor: TUE.dayColor,
  joiningWord: "WHEN",
  meaning: "joining a time",
  twoIdeas: [
    "We sing the warm-up song.",
    "Music class begins.",
  ],
  combined: "We sing the warm-up song WHEN music class begins.",
  whyItWorks:
    "WHEN tells us the TIME. The first idea is what we do; the second idea is when. Different subjects ('we' / 'music class') - keep both.",
  nonExample: {
    wrong: "We sing the warm-up song, music class begins.",
    why:
      "Two full ideas joined by only a comma - a comma splice. A comma cannot join two full ideas on its own. Use WHEN (or a full stop) to fix it.",
  },
  notes: composeNotes({
    say: [
      "WHEN tells us the TIME or the SEQUENCE. Listen: we sing the warm-up song WHEN music class begins.",
      "The first idea is what we do. The second idea is when. The two ideas happen together.",
      "Now the wrong example. Just a comma between two full ideas - that's a comma splice. The fix is to put WHEN between them.",
    ],
    do: [
      "Read both ideas aloud, then read the combined sentence.",
      "Highlight WHEN in the combined sentence.",
      "Read the non-example aloud and ask 'what is the only thing joining these two ideas?' Take one response - a comma is not enough.",
      "Quick note: WHEN can also go at the START of a sentence with a comma after the time clause (e.g. 'When music class begins, we sing the warm-up song.'). Both are correct. Default to WHEN in the middle this week.",
    ],
    teacherNotes: [
      "Comma splice is the target error today. Naming it for students is fine but not essential - focus on the fix: WHEN (or a full stop).",
      "Today's example has DIFFERENT subjects, so we keep both. The subject-repetition rule from Monday only fires when subjects match.",
    ],
    watchFor: [
      "Students using AND instead of WHEN - AND can work but WHEN is more precise here because the two events are connected by time, not just sequence.",
    ],
    tag: "[Day 2 | Tue | Teach WHEN]",
  }, { requireSay: false, requireDo: false }),
});

guidedSlide({
  dayText: TUE.dayText, dayColor: TUE.dayColor,
  joiningWord: "WHEN",
  pairs: [
    ["The art teacher claps three times.", "We stop and listen."],
    ["I always feel warmer.", "I have been running in PE."],
  ],
  answers: [
    "We stop and listen WHEN the art teacher claps three times.",
    "I always feel warmer WHEN I have been running in PE.",
  ],
  notesBuild: composeNotes({
    say: [
      "Together. Two pairs to combine using WHEN.",
      "Try pair 1 on your whiteboard - 60 seconds.",
      "Think: which idea tells us the TIME? Put WHEN before that idea.",
    ],
    do: [
      "PROTOCOL - whiteboards FIRST, hold up, THEN click reveal. Do not click reveal until every student has attempted.",
      "Read pair 1 aloud, pause 60 seconds, scan boards.",
      "Then pair 2, same routine.",
    ],
    teacherNotes: [
      "Pair 1: the cue is 'the art teacher claps three times'. The action is 'we stop and listen'. So 'We stop and listen WHEN the art teacher claps three times.' Different subjects, keep both.",
      "Pair 2: SAME subject ('I'). Accept both 'I always feel warmer WHEN I have been running in PE' (unreduced) and 'I always feel warmer after running in PE' (reduced, but uses AFTER - not the day's target word). Stick with the unreduced WHEN version.",
      "Both pairs support criterion 3 (clause linking).",
    ],
    watchFor: [
      "Students putting WHEN in front of the action instead of the time - reread aloud: which one tells you when?",
      "Comma splice - frame as progress, redirect to WHEN.",
    ],
    tag: "[Day 2 | Tue | Guided WHEN - build]",
  }, { requireSay: false, requireDo: false }),
  notesReveal: composeNotes({
    say: [
      "Here are the answers. Tick if yours matches the meaning.",
      "WHEN always sits before the TIME idea.",
    ],
    do: [
      "Read both answers, highlight WHEN.",
      "If 50%+ got both right, move on. If not, do one more quick pair: 'The bell rings.' + 'We line up.' -> 'We line up WHEN the bell rings.'",
    ],
    tag: "[Day 2 | Tue | Guided WHEN - reveal]",
  }, { requireSay: false, requireDo: false }),
});

independentSlide({
  dayText: TUE.dayText, dayColor: TUE.dayColor,
  joiningWord: "WHEN",
  prompt:
    "Go back to your hook whiteboard - the TWO ideas you wrote at the start about music, art or PE. " +
    "Combine them into ONE sentence using WHEN. Write the sentence in your warm-up book.",
  sentenceStarter: "I ___ WHEN ___.",
  notesBuild: composeNotes({
    say: [
      "Your turn. Find your hook whiteboard - the two short ideas you wrote at the start about a specialist class.",
      "Look at them. One of those ideas tells you WHEN. Combine the two into ONE sentence using WHEN.",
      "Write the sentence in your warm-up book, not your whiteboard. Four minutes.",
    ],
    do: [
      "Set a 4-minute timer.",
      "Circulate. For stuck students, read their two hook ideas aloud and ask 'which one tells you WHEN?'",
      "Note 1-2 strong sentences to share - especially any where you can compare the two-sentence hook side by side with the one-sentence combined version (visible 15-minute growth).",
      "Reveal the sentence stem ONLY for individual focus students who are stuck - do not announce it to the class.",
    ],
    enabling: [
      "FOCUS GROUP: if their hook ideas don't connect with a time relationship, give them a fresh pair orally (e.g. 'pack up my paints' / 'the music teacher claps') and let them combine that.",
      "Click reveal for the stem - walk to the student, point at the screen, walk away.",
    ],
    extension: [
      "HIGHER-ABILITY EXTENSION: choose ONE and write it in your book under your main sentence:",
      "  (a) Write TWO versions - one with WHEN in the middle and one with WHEN at the start (with a comma after the time clause). Write a line saying which sounds clearer and why.",
      "  (b) Add a WHERE detail to one of your clauses (e.g. 'I always sit on the mat in the art room WHEN we start a new project.').",
      "  (c) Try a rarer time word - WHILE, AS, AFTER or BEFORE - and check the sentence still says the same kind of thing.",
    ],
    watchFor: [
      "Students using AND where WHEN fits better - prompt them to identify which idea is the time.",
      "Comma splice - mark as progress and prompt the fix.",
      "Sentences where WHEN sits in front of the action by mistake (e.g. 'WHEN I sit on the mat, art class begins.' meaning that art begins because they sat down) - reread aloud and check the meaning.",
    ],
    tag: "[Day 2 | Tue | Independent WHEN - build, stem hidden]",
  }, { requireSay: false, requireDo: false }),
  notesReveal: composeNotes({
    say: [
      "If you need a start, here's a sentence frame.",
    ],
    do: [
      "Use this reveal for individual focus students only.",
      "Walk to them, point at the screen, then walk away. Do not read aloud to the class.",
    ],
    tag: "[Day 2 | Tue | Independent WHEN - stem reveal]",
  }, { requireSay: false, requireDo: false }),
});

exitSlide({
  dayText: TUE.dayText, dayColor: TUE.dayColor,
  joiningWord: "WHEN",
  taskHeader: "Today's exit check",
  taskIdeas: [
    "We get ready for PE.",
    "Mr Davies blows the whistle.",
  ],
  markCriteria: MARK_CRITERIA("WHEN"),
  qrLabel: "Tue WHEN form",
  notes: composeNotes({
    say: [
      "Last task. Combine these two ideas into ONE sentence using WHEN.",
      "Type the sentence into the Google Form on your iPad. Two minutes.",
      "If your iPad isn't working, write the sentence on a paper slip and drop it in the basket.",
    ],
    do: [
      "Timer 2 minutes.",
      "Use the marking criteria. Watch responses populate in the sheet view.",
      "Quick share before pack-up.",
    ],
    teacherNotes: [
      "Expected answer: 'We get ready for PE WHEN Mr Davies blows the whistle.' Or the reverse meaning: 'We get ready for PE, then Mr Davies blows the whistle' - but that doesn't use WHEN. Mark on meaning + grammar.",
      "Today's exit primarily checks criterion 3 (appropriate clause linking) and criterion 5 (no comma splice).",
    ],
    watchFor: [
      "Comma splice - cross criterion 5, frame as progress.",
      "Students who put WHEN before the wrong clause (e.g. 'Mr Davies blows the whistle WHEN we get ready for PE') - cross criterion 3 because the cause is wrong way round; reteach quickly.",
      "Students who use AND - cross criterion 3 (meaning is less precise) but tick criteria 1 and 5 if grammar is sound.",
    ],
    tag: "[Day 2 | Tue | Exit WHEN - QR submit]",
  }, { requireSay: false, requireDo: false }),
});

const outFile = path.join(OUT_DIR, "SC_T2W5_Lesson2_Tue_WHEN.pptx");
pres.writeFile({ fileName: outFile }).then(() => {
  console.log("PPTX written to " + outFile);
}).catch((err) => {
  console.error("Build failed:", err);
  process.exitCode = 1;
});
