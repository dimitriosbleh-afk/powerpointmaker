"use strict";

/**
 * Sentence Combining Warm-Up - Year 5/6 Enrichment, Term 2 Week 5, Day 3 (Wed)
 * Subordinating conjunction: IF (condition / hypothetical).
 * Topic: specialist subjects continued (different angle from Tue).
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

const FOOTER = "Sentence Combining Warm-Up | Year 5/6 | Term 2 Week 5 | Day 3 Wed";
const OUT_DIR = "output/SC_T2W5_Lesson3_Wed_IF";
fs.mkdirSync(OUT_DIR, { recursive: true });

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Sentence Combining Warm-Up - Year 5/6 - Term 2 Week 5 Day 3 (Wed)";
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
 *  DAY 3 - WEDNESDAY - IF (condition / hypothetical) - topic: specialists
 * ========================================================================= */

const WED = { dayText: "Wed", dayColor: C.ACCENT };

dayIntroSlide({
  dayText: WED.dayText, dayColor: WED.dayColor,
  joiningWord: "IF",
  hookHeader: "Hook - 30 seconds, then a recall",
  hookBody:
    "On your whiteboard, write TWO short ideas about an after-school club or something you do at recess. " +
    "Two short sentences with full stops. Hold onto these - we will combine them at the end. " +
    "Then quick recall: tell your partner ONE sentence using WHEN.",
  scItems: SC_BASE("IF"),
  notes: composeNotes({
    say: [
      "Two jobs to start. First - on your whiteboards, write TWO short ideas about an after-school club, lunchtime, or recess. Two short sentences. We will come back to these at the end.",
      "Second - quick recall. Tell your partner ONE sentence using yesterday's word, WHEN.",
      "Today's new word is IF. IF is for things that MIGHT happen - the second idea is a condition.",
    ],
    do: [
      "Set a 90-second timer covering both jobs - 60 seconds for the two-idea hook, 30 seconds for the partner recall.",
      "Scan whiteboards - keep them on desks for the rest of the lesson.",
      "Pick one strong WHEN example to read aloud, then introduce IF.",
    ],
    teacherNotes: [
      "Bookend pattern carried into Wednesday too: hook writes two short ideas; independent later combines them with IF.",
      "Topic broadens to club/recess so students think about choice and condition - sets up IF naturally.",
      "Today supports rubric criteria 3 (clause linking) and 5 (no comma splice).",
      "By end of today, students have all three subordinators (BECAUSE, WHEN, IF) ready for Thursday's mix.",
    ],
    enabling: [
      "FOCUS GROUP: if they cannot generate two ideas, give them oral options - 'we get to play handball' / 'the courts are dry' - and ask them to write those.",
      "EXTENSION: ask higher-ability students to think of an example sentence that uses IF before it's introduced.",
    ],
    watchFor: [
      "Students writing ONE long sentence instead of TWO short ones - redirect.",
      "Students still using AND where a subordinator was taught - remind them: WHEN tells us time precisely; AND just sequences.",
    ],
    tag: "[Day 3 | Wed | Hook + recall + LI/SC for IF]",
  }, { requireSay: false, requireDo: false }),
});

// Custom Wednesday teach slide: shows IF in TWO positions (middle, then start
// with comma) so the position/comma rule is taught explicitly. Steph and Bek
// flagged Week 4 that students need this distinction before Week 6.
(function teachIF() {
  const s = pres.addSlide();
  addTopBar(s, C.PRIMARY);
  badgeStack(s, "Teach", C.PRIMARY, WED.dayText, WED.dayColor);
  addTitle(s, `How "IF" works - joining a condition`);

  // Two worked examples side by side - same meaning, different position
  const workedY = CONTENT_TOP;
  const workedH = 2.20;
  const colW = 4.45;
  const gap = 0.10;

  const examples = [
    {
      label: "Position 1 - IF in the middle",
      subRule: "No comma needed.",
      sentence: "We get to choose our art project IF we finish the warm-up quickly.",
      color: C.PRIMARY,
    },
    {
      label: "Position 2 - IF at the start",
      subRule: "Add a comma after the condition.",
      sentence: "IF we finish the warm-up quickly, we get to choose our art project.",
      color: C.SECONDARY,
    },
  ];

  examples.forEach((ex, i) => {
    const x = 0.5 + i * (colW + gap);
    addCard(s, x, workedY, colW, workedH, { strip: ex.color, fill: C.WHITE });
    s.addText(ex.label, {
      x: x + 0.16, y: workedY + 0.10, w: colW - 0.32, h: 0.30,
      fontSize: 12, fontFace: FONT_B, color: ex.color, bold: true, margin: 0,
    });
    s.addText(ex.subRule, {
      x: x + 0.16, y: workedY + 0.40, w: colW - 0.32, h: 0.26,
      fontSize: 11, fontFace: FONT_B, color: C.MUTED, italic: true, margin: 0,
    });
    const runs = highlightSentence(ex.sentence, "IF", 14, C.CHARCOAL, ex.color);
    s.addText(runs.map((run, j) => ({
      text: `${j === 0 ? '"' : ""}${run.text}${j === runs.length - 1 ? '"' : ""}`,
      options: { ...run.options, fontFace: FONT_H, breakLine: false },
    })), {
      x: x + 0.16, y: workedY + 0.72, w: colW - 0.32, h: workedH - 0.82,
      italic: true, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });
  });

  // Why-it-works strip
  const whyY = workedY + workedH + 0.10;
  s.addShape("rect", {
    x: 0.5, y: whyY, w: 9, h: 0.36, fill: { color: C.BG_CARD },
  });
  s.addText(
    "Why this works: IF tells us the CONDITION. Same meaning in both - it's only the position that changes the comma.",
    {
      x: 0.75, y: whyY + 0.02, w: 8.5, h: 0.32,
      fontSize: 12, fontFace: FONT_B, color: C.SECONDARY, bold: true, margin: 0,
      fit: "shrink", shrinkText: true,
    },
  );

  // Watch-out: fragment
  const neY = whyY + 0.36 + 0.14;
  const neH = SAFE_BOTTOM - neY;
  addCard(s, 0.5, neY, 9, neH, { strip: C.ALERT, fill: C.WHITE });
  s.addText("Watch out - common mistake", {
    x: 0.75, y: neY + 0.06, w: 5, h: 0.26,
    fontSize: 12, fontFace: FONT_B, color: C.ALERT, bold: true, margin: 0,
  });
  s.addText(`Wrong: "IF we finish the warm-up quickly."`, {
    x: 0.75, y: neY + 0.32, w: 8.5, h: 0.30,
    fontSize: 13.5, fontFace: FONT_H, color: C.CHARCOAL, italic: true, margin: 0,
    fit: "shrink", shrinkText: true,
  });
  s.addText(
    "An IF clause on its own is a fragment - it needs a main idea attached. The reader is left asking 'IF we finish, then what?'",
    {
      x: 0.75, y: neY + 0.62, w: 8.5, h: neH - 0.72,
      fontSize: 11.5, fontFace: FONT_B, color: C.MUTED, valign: "top", margin: 0,
      fit: "shrink", shrinkText: true,
    },
  );

  addFooter(s, FOOTER);
  s.addNotes(composeNotes({
    say: [
      "IF tells us the CONDITION. The second idea has to happen first for the first idea to be true.",
      "Today there are TWO positions. Listen.",
      "Position 1, IF in the middle - 'We get to choose our art project IF we finish the warm-up quickly.' No comma.",
      "Position 2, IF at the start - 'IF we finish the warm-up quickly, we get to choose our art project.' Comma after the condition.",
      "Same meaning. Only the position changes the comma rule.",
      "Now the wrong example. An IF clause on its own is a fragment - not a full sentence. The reader is left hanging.",
    ],
    do: [
      "Read both worked examples aloud, with a clear pause at the comma in position 2.",
      "Highlight IF in each.",
      "Read the wrong example aloud and ask 'what's missing here?' - a main idea.",
    ],
    teacherNotes: [
      "Position + comma rule is the heart of today. Spelling it out now saves a comma-splice confusion later. Students who learned not to use comma splices in Week 4 may otherwise think the comma after 'IF we finish quickly,' is a contradiction.",
      "The same position/comma rule applies to WHEN and BECAUSE too - just less critical because students do not flip those positions as often.",
      "Subject-rule check: both worked examples have the SAME subject ('we') in both clauses. Keeping both reads more naturally in Year 5/6 - drop is optional, not required.",
      "Fragment is the secondary target error today. Naming the fragment is fine; the fix matters more: attach the main idea.",
    ],
    watchFor: [
      "Students who put a comma after IF in the MIDDLE position - prompt 'where is IF here? Middle - so no comma needed.'",
      "Students who DROP the comma when IF is at the start - prompt 'where is IF here? Start - so comma after the condition.'",
      "Students using WHEN where IF fits better - 'we get to choose art WHEN we finish quickly' suggests it always happens; IF suggests a condition. Subtle but worth noticing.",
    ],
    tag: "[Day 3 | Wed | Teach IF - position + comma rule]",
  }, { requireSay: false, requireDo: false }));
  runSlideDiagnostics(s, pres, { respectSafeBottom: false });
})();

guidedSlide({
  dayText: WED.dayText, dayColor: WED.dayColor,
  joiningWord: "IF",
  pairs: [
    ["You can borrow the good colour pencils.", "You promise to put them back."],
    ["Mrs Patel lets us play a music game.", "Everyone tidies their instruments."],
  ],
  answers: [
    "You can borrow the good colour pencils IF you promise to put them back.",
    "Mrs Patel lets us play a music game IF everyone tidies their instruments.",
  ],
  notesBuild: composeNotes({
    say: [
      "Together. Two pairs to combine using IF.",
      "Try pair 1 on your whiteboard. About 60 seconds. Then pair 2.",
      "Think: which idea is the condition? Put IF before it.",
    ],
    do: [
      "PROTOCOL - whiteboards FIRST, hold up, THEN click reveal. Do not click reveal until every student has attempted.",
      "Read pair 1 aloud, pause 60 seconds, scan.",
      "Then pair 2, same routine.",
    ],
    teacherNotes: [
      "Pair 1: the condition is the promise. So 'You can borrow the good colour pencils IF you promise to put them back.' Same subject ('you'), but keeping both reads more naturally in Year 5/6 - accept either.",
      "Pair 2: the condition is everyone tidying. Different subjects ('Mrs Patel' / 'everyone'), so keep both.",
      "Today's pairs support criterion 3 (clause linking).",
    ],
    watchFor: [
      "Students putting IF before the wrong clause - reread aloud and ask 'which one is the condition? Which one has to happen first?'",
      "Comma splice - frame as developmental progress, redirect to IF.",
    ],
    tag: "[Day 3 | Wed | Guided IF - build]",
  }, { requireSay: false, requireDo: false }),
  notesReveal: composeNotes({
    say: [
      "Here are the answers. Tick yours if the meaning matches.",
      "IF always sits before the CONDITION.",
    ],
    do: [
      "Read both answers, highlight IF.",
      "If 50%+ got both right, move on. If not, do a fresh quick pair from the board: 'I will pack up.' + 'Mrs Patel asks.' -> 'I will pack up IF Mrs Patel asks.'",
    ],
    tag: "[Day 3 | Wed | Guided IF - reveal]",
  }, { requireSay: false, requireDo: false }),
});

independentSlide({
  dayText: WED.dayText, dayColor: WED.dayColor,
  joiningWord: "IF",
  prompt:
    "Go back to your hook whiteboard - the TWO ideas you wrote at the start about after-school or recess. " +
    "Combine them into ONE sentence using IF. Write the sentence in your warm-up book.",
  sentenceStarter: "We get to ___ IF ___.",
  notesBuild: composeNotes({
    say: [
      "Your turn. Find your hook whiteboard - the two short ideas from the start about an after-school club or recess.",
      "Look at them. One of those ideas is the CONDITION - the thing that has to happen first. Combine them into ONE sentence using IF.",
      "Write the sentence in your warm-up book. Four minutes.",
    ],
    do: [
      "Set a 4-minute timer.",
      "Circulate. For stuck students, read their two hook ideas aloud and ask 'which one has to happen first for the other one to work?'",
      "Note 1-2 strong sentences to share - especially any where the hook board and the warm-up book version sit side by side and show the 15-minute growth.",
      "Reveal the stem ONLY for individual focus students who are stuck.",
    ],
    enabling: [
      "FOCUS GROUP: if their hook ideas don't have a clear condition, give them an oral pair (e.g. 'we get to play handball' / 'the courts are dry') and ask them to put IF between them. Click reveal to show the stem and point to it.",
    ],
    extension: [
      "HIGHER-ABILITY EXTENSION: choose ONE and write under your main sentence:",
      "  (a) Write TWO versions - IF in the middle, and IF at the start (with a comma after the condition). Write a line saying which is clearer.",
      "  (b) Add a WHEN detail to your condition (e.g. 'We get to play a music game IF everyone packs up before the bell rings.').",
      "  (c) Use a rarer condition word - UNLESS or PROVIDED - and check the meaning is similar.",
    ],
    watchFor: [
      "Students writing two separate sentences - redirect to ONE sentence joined by IF.",
      "Students using IF where WHEN fits better (when the event always happens) - ask 'does this always happen, or only sometimes?'",
      "Comma splice - mark as developmental progress.",
    ],
    tag: "[Day 3 | Wed | Independent IF - build, stem hidden]",
  }, { requireSay: false, requireDo: false }),
  notesReveal: composeNotes({
    say: [
      "If you need a start, here's a sentence frame.",
    ],
    do: [
      "Reveal for individual focus students only. Walk to them, point at the screen, walk away.",
    ],
    tag: "[Day 3 | Wed | Independent IF - stem reveal]",
  }, { requireSay: false, requireDo: false }),
});

exitSlide({
  dayText: WED.dayText, dayColor: WED.dayColor,
  joiningWord: "IF",
  taskHeader: "Today's exit check",
  taskIdeas: [
    "We can paint outside today.",
    "The weather stays dry.",
  ],
  markCriteria: MARK_CRITERIA("IF"),
  qrLabel: "Wed IF form",
  notes: composeNotes({
    say: [
      "Last task. Combine these two ideas into ONE sentence using IF.",
      "Type the sentence into the Google Form on your iPad. Two minutes.",
      "If your iPad isn't working, write the sentence on a paper slip and drop it in the basket.",
    ],
    do: [
      "Timer 2 minutes.",
      "Use marking criteria. Watch responses populate live.",
      "Quick share before pack-up.",
    ],
    teacherNotes: [
      "Expected answer: 'We can paint outside today IF the weather stays dry.' Or the IF-first version: 'IF the weather stays dry, we can paint outside today.' Both fine.",
      "Today's exit primarily checks criterion 3 (appropriate clause linking) and criterion 5 (no comma splice).",
    ],
    watchFor: [
      "Comma splice - cross criterion 5, frame as progress.",
      "Students putting IF in front of the wrong clause (e.g. 'The weather stays dry IF we can paint outside today') - cross criterion 3 and reteach quickly.",
      "Students using WHEN - cross criterion 3 (meaning is different; condition vs time).",
    ],
    tag: "[Day 3 | Wed | Exit IF - QR submit]",
  }, { requireSay: false, requireDo: false }),
});

const outFile = path.join(OUT_DIR, "SC_T2W5_Lesson3_Wed_IF.pptx");
pres.writeFile({ fileName: outFile }).then(() => {
  console.log("PPTX written to " + outFile);
}).catch((err) => {
  console.error("Build failed:", err);
  process.exitCode = 1;
});
