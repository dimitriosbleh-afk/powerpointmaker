"use strict";

/**
 * Sentence Combining Warm-Up - Year 5/6 Enrichment, Term 2 Week 5, Day 1 (Mon)
 *
 * Subordinating conjunctions: BECAUSE.
 *
 * Day 1 carries the front matter for the whole week:
 *   - Title slide
 *   - Teacher Resources slide
 *   - Week framing slide (PLC week 2 of 6, links to 5-criteria post-test rubric)
 *   - Monday's 5-slide daily routine (Hook + LI/SC, Teach, Guided (reveal pair),
 *     Independent (reveal pair, stem hidden by default), Exit check with QR).
 *
 * Daily routine (15 min total): review/hook (2) -> teach (3) -> guided (4)
 *  -> independent in workbook (4) -> exit (2 - QR + paper backup).
 *
 * Week 5 carry-forward from Week 4 feedback:
 *   - Production happens in workbooks, NOT whiteboards.
 *   - Sentence stem on the Your Turn slide is HIDDEN by default; reveal only
 *     for focus students who need it.
 *   - Guided practice: students write FIRST, hold up, THEN reveal. Notes say so.
 *   - Extension prompts are genuinely harder: two valid combinations + explain
 *     which is clearer; add where/when detail; use a rarer subordinator
 *     (since, although, while).
 *   - Comma splice marked as developmental progress, not failure.
 *
 * Monday-specific moves:
 *   - Hook (2 ideas about inter-school sports) is referenced again at the
 *     exit so students see their 15-minute growth.
 *   - Subject-repetition rule taught explicitly:
 *       same subject in both ideas -> can drop the repeat after BECAUSE
 *       different subjects -> keep both.
 */

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant, composeNotes } = require("../themes/factory");

// Keep variant 3 (Week 4's variant) for visual continuity across the 6-week
// PLC cycle. Students recognise the routine; this is Week 2 of 6 of the same
// warm-up sequence.
const T = createTheme("literacy", "grade56", weekToVariant(4));
const {
  C, FONT_H, FONT_B,
  SLIDE_W, SLIDE_H, SAFE_BOTTOM, CONTENT_TOP,
  addCard, addFooter, addTopBar, addBadge, addTitle, addTextOnShape,
  titleSlide, withReveal, runSlideDiagnostics,
} = T;

const FOOTER = "Sentence Combining Warm-Up | Year 5/6 | Term 2 Week 5 | Day 1 Mon";
const OUT_DIR = "output/SC_T2W5_Lesson1_Mon_BECAUSE";
fs.mkdirSync(OUT_DIR, { recursive: true });

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Sentence Combining Warm-Up - Year 5/6 - Term 2 Week 5 Day 1 (Mon)";
pres.author = "Year 5/6 Enrichment PLC";

/* Shared success-criteria builder (no tier labels on student-facing slides). */
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

/* ─────────────────────────────────────────────────────────────────────────
 *  Slide builders (Day intro, Teach, Guided, Independent, Exit)
 *  Independent is a reveal pair: build = clean prompt, reveal = adds stem.
 * ───────────────────────────────────────────────────────────────────────── */

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
      fontSize: 15,
      color: C.CHARCOAL,
      paraSpaceAfter: 6,
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

/* Independent (reveal pair): build slide shows the prompt only.
 * Reveal slide adds the sentence stem strip - teacher uses judgement on
 * whether to click. Production happens in workbooks, not whiteboards. */
function independentSlide({ dayText, dayColor, joiningWord, prompt, sentenceStarter, notesBuild, notesReveal }) {
  return withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ACCENT);
      badgeStack(s, "Your turn", C.ACCENT, dayText, dayColor);
      addTitle(s, `Write ONE sentence using "${joiningWord}"`);

      // Workbook cue (top-right) - clarifies that this is NOT a whiteboard task
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

      // Reminder strip (full width below prompt)
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
      // Stash the prompt geometry so the reveal can size the stem strip.
      s._scLayout = { promptY, promptH, remindY, remindH };
      return s;
    },
    (s) => {
      // Reveal: shrink the prompt card, slot a sentence-stem strip beneath it
      // before the reminder strip. The whole frame stays inside the safe zone.
      const promptY = CONTENT_TOP;
      const promptH = 1.95; // shrunk from 2.55 to make room for the stem strip
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

/* Exit check with QR placeholder + marking criteria.
 * Left: task ideas. Middle: QR + Google Form panel. Right: marking criteria. */
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
  // Three-panel layout: task | QR | marking
  const leftW = 3.7;
  const midX = 4.4;
  const midW = 2.0;
  const rightX = 6.55;
  const rightW = 2.95;

  // Left: task ideas
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

  // Middle: QR panel
  addCard(s, midX, CONTENT_TOP, midW, cardH, { strip: C.PRIMARY, fill: C.BG_CARD });
  s.addText("Scan to submit", {
    x: midX + 0.12, y: CONTENT_TOP + 0.12, w: midW - 0.24, h: 0.28,
    fontSize: 11, fontFace: FONT_B, color: C.PRIMARY, bold: true,
    align: "center", margin: 0,
  });
  // QR placeholder square - teacher replaces with their school's Google Form QR
  const qrSize = 1.35;
  const qrX = midX + (midW - qrSize) / 2;
  const qrY = CONTENT_TOP + 0.46;
  s.addShape("rect", {
    x: qrX, y: qrY, w: qrSize, h: qrSize,
    fill: { color: C.WHITE },
    line: { color: C.CHARCOAL, width: 1.0 },
  });
  // Draw the three QR finder squares so the placeholder reads as a QR pattern
  const fp = 0.28;
  const inset = 0.10;
  [[qrX + inset, qrY + inset], [qrX + qrSize - inset - fp, qrY + inset], [qrX + inset, qrY + qrSize - inset - fp]].forEach(([fx, fy]) => {
    s.addShape("rect", {
      x: fx, y: fy, w: fp, h: fp,
      fill: { color: C.CHARCOAL },
    });
    s.addShape("rect", {
      x: fx + 0.06, y: fy + 0.06, w: fp - 0.12, h: fp - 0.12,
      fill: { color: C.WHITE },
    });
    s.addShape("rect", {
      x: fx + 0.10, y: fy + 0.10, w: fp - 0.20, h: fp - 0.20,
      fill: { color: C.CHARCOAL },
    });
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

  // Right: marking criteria
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

/* ─────────────────────────────────────────────────────────────────────────
 *  Front matter (title, resources, framing) - lesson 1 only
 * ───────────────────────────────────────────────────────────────────────── */

titleSlide(
  pres,
  "Sentence Combining Warm-Up",
  "5 days x 15 minutes - subordinating conjunctions: because, when, if",
  "Year 5/6 Enrichment  |  Term 2 Week 5  |  PLC Cycle Week 2 of 6",
  composeNotes({
    say: [
      "This week we're learning to join two short ideas using subordinating conjunctions.",
      "Same 15-minute routine as last week: review, teach, practise together, write your own, exit check.",
      "Monday is BECAUSE, Tuesday is WHEN, Wednesday is IF. Thursday mixes all three. Friday contrasts these new words with last week's AND, BUT, SO, OR.",
    ],
    do: [
      "Use the day labels at the top of each slide to navigate.",
      "Have warm-up books and mini-whiteboards ready before students arrive.",
      "Open the cohort tracker (and Google Form sheet view) before Day 1.",
    ],
    teacherNotes: [
      "Goal this week: 60% of focus students achieving 60%+ on the 5-criteria post-test (3 out of 5 criteria).",
      "Production happens in warm-up books, not whiteboards. Whiteboards are for guided practice only.",
      "Australian curriculum 2.0 alignment: Year 5 and 6 English, Language strand (text structure and organisation; expressing and developing ideas). Specifically VCELA323 - main and subordinate clauses.",
      "HITS being targeted: setting goals, worked examples, explicit teaching.",
    ],
    tag: "[Sentence Combining | Year 5/6 | T2 W5]",
  }, { requireSay: false, requireDo: false })
);

function teacherResourcesSlide() {
  const s = pres.addSlide();
  addTopBar(s, C.SECONDARY);
  addBadge(s, "Teacher resources", { color: C.SECONDARY, w: 2.4 });
  addTitle(s, "What you need before you start");

  const cardH = SAFE_BOTTOM - CONTENT_TOP;
  const leftW = 4.4;
  const rightW = 4.6;
  const rightX = 5.0;

  addCard(s, 0.5, CONTENT_TOP, leftW, cardH, { strip: C.PRIMARY, fill: C.WHITE });
  s.addText("In the room", {
    x: 0.7, y: CONTENT_TOP + 0.14, w: leftW - 0.4, h: 0.32,
    fontSize: 13, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
  });
  const materials = [
    "Warm-up book per student (independent + exit writing)",
    "Mini-whiteboards + markers (guided practice only)",
    "Projector with this deck open",
    "iPad per student for QR exit submission",
    "Paper exit slips + basket as iPad backup",
    "Cohort tracker open (Google Sheet view of the form)",
  ];
  s.addText(materials.map((t, i) => ({
    text: t,
    options: {
      bullet: true,
      breakLine: i < materials.length - 1,
      fontSize: 13.5, color: C.CHARCOAL,
      paraSpaceAfter: 5,
    },
  })), {
    x: 0.7, y: CONTENT_TOP + 0.50, w: leftW - 0.4, h: cardH - 0.66,
    fontFace: FONT_B, valign: "top", margin: 0,
    fit: "shrink", shrinkText: true,
  });

  addCard(s, rightX, CONTENT_TOP, rightW, cardH, { strip: C.ACCENT, fill: C.BG_CARD });
  s.addText("No printed student resources", {
    x: rightX + 0.20, y: CONTENT_TOP + 0.14, w: rightW - 0.4, h: 0.32,
    fontSize: 13, fontFace: FONT_B, color: C.ACCENT, bold: true, margin: 0,
  });
  const teacherInfo = [
    "All worked examples and prompts are on the slides.",
    "Exit check uses a QR code per day linked to a Google Form (fields: name, combined sentence).",
    "Replace the QR placeholder on each Exit slide with your school's actual QR for that day's form.",
    "Differentiation in speaker notes: focus group (15 students, 6-12 months below) + extension.",
    "Sentence stem on the Your Turn slide is HIDDEN by default - reveal only for focus students.",
    "Guided practice: students attempt on whiteboards FIRST, then click reveal.",
  ];
  s.addText(teacherInfo.map((t, i) => ({
    text: t,
    options: {
      bullet: true,
      breakLine: i < teacherInfo.length - 1,
      fontSize: 12.5, color: C.CHARCOAL,
      paraSpaceAfter: 4,
    },
  })), {
    x: rightX + 0.20, y: CONTENT_TOP + 0.50, w: rightW - 0.4, h: cardH - 0.66,
    fontFace: FONT_B, valign: "top", margin: 0,
    fit: "shrink", shrinkText: true,
  });

  addFooter(s, FOOTER);
  s.addNotes(composeNotes({
    say: [
      "Quick check before Day 1: warm-up books, whiteboards, iPads, paper slip basket, cohort tracker.",
      "Workbooks are the permanent record this week. Whiteboards are only for the guided practice attempt.",
    ],
    do: [
      "Distribute whiteboards and markers as students enter.",
      "Open the Google Form sheet view so you can scan responses live during the exit check.",
      "Have the paper backup slip basket on a side bench in case iPads are down.",
    ],
    teacherNotes: [
      "This is the only setup slide for the whole week. Same materials apply Mon-Fri.",
      "If a teacher in the team has not been briefed, this slide plus the framing slide is enough to teach Day 1.",
    ],
    tag: "[Sentence Combining | Resources | T2 W5]",
  }, { requireSay: false, requireDo: false }));
  runSlideDiagnostics(s, pres, { respectSafeBottom: false });
  return s;
}

teacherResourcesSlide();

function weekFramingSlide() {
  const s = pres.addSlide();
  addTopBar(s, C.PRIMARY);
  addBadge(s, "PLC framing", { color: C.PRIMARY, w: 2.0 });
  addTitle(s, "Where Week 5 fits");

  const cardH = SAFE_BOTTOM - CONTENT_TOP;
  const leftW = 4.6;
  const rightW = 4.4;
  const rightX = 5.2;

  // Left: this week's focus + carry-forward
  addCard(s, 0.5, CONTENT_TOP, leftW, cardH, { strip: C.ACCENT, fill: C.WHITE });
  s.addText("This week", {
    x: 0.7, y: CONTENT_TOP + 0.14, w: leftW - 0.4, h: 0.30,
    fontSize: 13, fontFace: FONT_B, color: C.ACCENT, bold: true, margin: 0,
  });
  s.addText(
    "Week 2 of 6 in the PLC cycle. Subordinating conjunctions: BECAUSE (Mon), WHEN (Tue), IF (Wed). " +
    "Thursday mixes all three. Friday contrasts subordinating with last week's coordinating words.",
    {
      x: 0.7, y: CONTENT_TOP + 0.46, w: leftW - 0.4, h: 1.15,
      fontSize: 13.5, fontFace: FONT_B, color: C.CHARCOAL, valign: "top", margin: 0,
    },
  );
  const carryY = CONTENT_TOP + 1.70;
  s.addText("Carry-forward from Week 4", {
    x: 0.7, y: carryY, w: leftW - 0.4, h: 0.28,
    fontSize: 12, fontFace: FONT_B, color: C.SECONDARY, bold: true, margin: 0,
  });
  const carry = [
    "Production in workbooks, not whiteboards.",
    "Sentence stem hidden by default - reveal only if a focus student is stuck.",
    "Comma splice marked as developmental progress, not failure.",
    "Extension prompts are genuinely harder this week.",
  ];
  s.addText(carry.map((t, i) => ({
    text: t,
    options: {
      bullet: true,
      breakLine: i < carry.length - 1,
      fontSize: 12, color: C.CHARCOAL, paraSpaceAfter: 3,
    },
  })), {
    x: 0.7, y: carryY + 0.30, w: leftW - 0.4, h: cardH - carryY + CONTENT_TOP - 0.40,
    fontFace: FONT_B, valign: "top", margin: 0,
    fit: "shrink", shrinkText: true,
  });

  // Right: the 5-criteria post-test rubric
  addCard(s, rightX, CONTENT_TOP, rightW, cardH, { strip: C.PRIMARY, fill: C.BG_CARD });
  s.addText("Post-test rubric (5 criteria)", {
    x: rightX + 0.20, y: CONTENT_TOP + 0.14, w: rightW - 0.4, h: 0.30,
    fontSize: 13, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
  });
  s.addText("SMART goal: 60% of focus students achieving 60%+ (3 of 5).", {
    x: rightX + 0.20, y: CONTENT_TOP + 0.44, w: rightW - 0.4, h: 0.34,
    fontSize: 11, fontFace: FONT_B, color: C.SECONDARY, italic: true, margin: 0,
  });
  const criteria = [
    "All ideas retained",
    "One grammatical sentence",
    "Appropriate clause linking",
    "No unnecessary repetition",
    "No fragment, run-on or comma splice",
  ];
  const rowsTop = CONTENT_TOP + 0.86;
  const rowH = (cardH - 1.00) / criteria.length;
  criteria.forEach((c, i) => {
    const y = rowsTop + i * rowH;
    s.addShape("rect", {
      x: rightX + 0.20, y: y + 0.04, w: 0.10, h: rowH - 0.10,
      fill: { color: C.PRIMARY },
    });
    s.addText(`${i + 1}.  ${c}`, {
      x: rightX + 0.36, y: y, w: rightW - 0.56, h: rowH,
      fontSize: 12.5, fontFace: FONT_B, color: C.CHARCOAL, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });
  });

  addFooter(s, FOOTER);
  s.addNotes(composeNotes({
    say: [
      "Quick framing - this is Week 2 of 6 in our PLC cycle.",
      "Last week was coordinating conjunctions: AND, BUT, SO, OR. This week is subordinating: BECAUSE, WHEN, IF.",
      "By Week 8 students sit a post-test scored on these 5 criteria.",
    ],
    do: [
      "Read this slide once before Day 1. It is for you, not students.",
      "Point any teacher in the team who has not been briefed to this slide first.",
    ],
    teacherNotes: [
      "Each day's lesson primarily supports criterion 3 (appropriate clause linking) and criterion 5 (no comma splice).",
      "Friday is the day to check criterion 4 (no unnecessary repetition) because the contrast with coordinating words exposes whether students are repeating clauses.",
      "Cohort tracker fills automatically from the daily Google Form. Sheet view shows you who has not submitted.",
    ],
    tag: "[PLC framing | Week 5 of T2 | Cycle Week 2 of 6]",
  }, { requireSay: false, requireDo: false }));
  runSlideDiagnostics(s, pres, { respectSafeBottom: false });
  return s;
}

weekFramingSlide();

/* =========================================================================
 *  DAY 1 - MONDAY - BECAUSE (cause/reason) - topic: inter-school sports
 * ========================================================================= */

const MON = { dayText: "Mon", dayColor: C.PRIMARY };

dayIntroSlide({
  dayText: MON.dayText, dayColor: MON.dayColor,
  joiningWord: "BECAUSE",
  hookHeader: "Hook - quick start (Day 1)",
  hookBody:
    "On your whiteboard, write TWO short ideas about an inter-school sports game you have played this year. " +
    "Two short sentences with full stops. We will come back to these at the end of the session.",
  scItems: SC_BASE("BECAUSE"),
  notes: composeNotes({
    say: [
      "Welcome back to our 15-minute warm-up. Last week we joined ideas with AND, BUT, SO and OR.",
      "This week we move to another set of joining words. These are called subordinating conjunctions. They tell us WHY, WHEN, or under what condition.",
      "Today's word is BECAUSE. BECAUSE tells us the reason.",
      "First job - hook. On your whiteboards, write TWO short ideas about an inter-school sports game you have played. Two short sentences. Hold onto them - we will use them at the end.",
    ],
    do: [
      "Hand out whiteboards as students arrive.",
      "Set a 90-second timer for the hook. Scan boards as you walk - note one or two pairs you might revisit at the exit.",
      "Keep boards on desks until end of session.",
    ],
    teacherNotes: [
      "Inter-school sports works as a topic because every student in the year level has played a round before Week 5. Universal access, low background-knowledge cost.",
      "Day 1 specifically uses a hook (not a same-day review). From Tuesday onwards this slide reviews yesterday's joining word.",
      "Today primarily supports rubric criteria 3 (appropriate clause linking) and 5 (no comma splice).",
    ],
    enabling: [
      "FOCUS GROUP (around 15 students, 6-12 months below grade level): if they're stuck, give them oral sentence options - 'we played netball' / 'I scored a goal' - and just ask them to write those.",
      "EXTENSION: ask higher-ability students to write two ideas that are clearly linked by cause and effect, so they're set up for the BECAUSE move on the next slide.",
    ],
    watchFor: [
      "Students writing ONE long sentence instead of two short ones - redirect: 'two short ideas, two full stops'.",
    ],
    tag: "[Day 1 | Mon | Hook + LI/SC]",
  }, { requireSay: false, requireDo: false }),
});

teachSlide({
  dayText: MON.dayText, dayColor: MON.dayColor,
  joiningWord: "BECAUSE",
  meaning: "joining a reason",
  twoIdeas: [
    "We lost the netball game.",
    "Our goal shooter was away sick.",
  ],
  // Same-subject vs different-subject reduction rule is the secondary move
  // today. Both ideas here have DIFFERENT subjects ('we' / 'our goal shooter'),
  // so we keep both. The Watch Out below shows the same-subject case.
  combined: "We lost the netball game BECAUSE our goal shooter was away sick.",
  whyItWorks:
    "BECAUSE adds the REASON. The first idea is what happened; the second idea is why. Different subjects ('we' / 'our goal shooter') so we keep both.",
  nonExample: {
    wrong: "Because our goal shooter was away sick. We lost the netball game.",
    why:
      "A clause starting with BECAUSE on its own is a fragment - it is not a full sentence. " +
      "Either join it to the main idea (no full stop in the middle) or remove BECAUSE.",
  },
  notes: composeNotes({
    say: [
      "BECAUSE adds the REASON. Listen: We lost the netball game BECAUSE our goal shooter was away sick.",
      "The first idea is what happened. The second idea is why.",
      "Subject rule - really important. These two ideas have DIFFERENT subjects: 'we' and 'our goal shooter'. So we keep both. If both ideas had the SAME subject - 'we lost the game' and 'we played without our shooter' - you could drop the second 'we' after BECAUSE.",
      "Now the wrong example. Starting a sentence with BECAUSE and then putting a full stop after the reason leaves the reason floating on its own. That is called a fragment. The fix is to join the two ideas without the full stop in the middle.",
    ],
    do: [
      "Track with your finger from idea 1 to idea 2, then point at the combined sentence.",
      "Highlight BECAUSE in the combined sentence.",
      "Show one quick same-subject example live on the board: 'We trained hard.' + 'We won.' -> 'We won BECAUSE we trained hard.' Then cross out the second 'we' to show the drop.",
      "Read the non-example aloud and ask 'what is wrong here?' Take one quick response.",
    ],
    teacherNotes: [
      "Two moves today: (1) JOIN with BECAUSE for a reason. (2) Apply the SUBJECT-REPETITION rule - same subject means you can drop the second one; different subjects means keep both.",
      "Students do not need to name 'subordinate clause' - they need to use BECAUSE accurately. Save the formal term for Week 6 if it comes up.",
    ],
    watchFor: [
      "Students confusing cause and effect (joining the wrong way round) - reread aloud and ask 'which one is the reason?'",
      "Students who try to start the sentence with BECAUSE - acceptable but only if the comma and second clause are correct (a Year 7 move). Default to BECAUSE in the middle for this week.",
    ],
    tag: "[Day 1 | Mon | Teach BECAUSE + subject rule]",
  }, { requireSay: false, requireDo: false }),
});

guidedSlide({
  dayText: MON.dayText, dayColor: MON.dayColor,
  joiningWord: "BECAUSE",
  pairs: [
    ["We arrived late to the carnival.", "The bus had a flat tyre."],
    ["I was nervous before my race.", "I had never run a relay before."],
  ],
  answers: [
    "We arrived late to the carnival BECAUSE the bus had a flat tyre.",
    "I was nervous before my race BECAUSE I had never run a relay before.",
  ],
  notesBuild: composeNotes({
    say: [
      "Together. Two pairs to combine using BECAUSE.",
      "Try pair 1 on your whiteboard. About 60 seconds. Then pair 2.",
      "Think: which idea is the reason? Put BECAUSE before the reason.",
    ],
    do: [
      "PROTOCOL - students write on whiteboards FIRST, hold up, THEN click reveal. Do not click reveal until every student has attempted. If kids learn to wait, the practice drifts to recognition instead of production.",
      "Read pair 1 aloud, pause 60 seconds, scan boards.",
      "Then pair 2, same routine.",
      "Click reveal once everyone has had a real attempt.",
    ],
    teacherNotes: [
      "Pair 1 has DIFFERENT subjects ('we' / 'the bus'), so we keep both - no reduction.",
      "Pair 2 has the SAME subject ('I'). Both 'I was nervous BECAUSE I had never run a relay before' and the reduced 'I was nervous BECAUSE never having run a relay before' are too advanced - the reduced version for BECAUSE is not natural Year 5/6 writing. Accept both 'I' subjects.",
      "Today's pairs support criteria 3 (clause linking) and 5 (no comma splice).",
    ],
    watchFor: [
      "Students using a comma instead of BECAUSE - that is a comma splice. Frame as developmental progress: 'you've recognised the two ideas connect, that's progress, now let's add the joining word'.",
      "Students putting BECAUSE in front of the wrong clause - reread aloud and ask 'which one is the reason?'",
    ],
    tag: "[Day 1 | Mon | Guided BECAUSE - build]",
  }, { requireSay: false, requireDo: false }),
  notesReveal: composeNotes({
    say: [
      "Here are the answers. Tick yours if the meaning matches, even if your wording is slightly different.",
      "BECAUSE always comes before the reason.",
    ],
    do: [
      "Read both answers aloud. Highlight BECAUSE in each.",
      "If 50%+ got both right, move to independent. If less, do one more pair from the board with a fresh pair: 'We were tired.' + 'We had run a long way.'",
    ],
    tag: "[Day 1 | Mon | Guided BECAUSE - reveal]",
  }, { requireSay: false, requireDo: false }),
});

independentSlide({
  dayText: MON.dayText, dayColor: MON.dayColor,
  joiningWord: "BECAUSE",
  prompt:
    "Go back to your hook whiteboard - the TWO ideas you wrote about inter-school sports. " +
    "Join them into ONE sentence using BECAUSE. Write the sentence in your warm-up book.",
  sentenceStarter: "I ___ BECAUSE ___.",
  notesBuild: composeNotes({
    say: [
      "Now your turn. Find your hook whiteboard - the two short ideas you wrote at the start.",
      "Look at them. One of those ideas is the action, the other is the reason. Combine them into ONE sentence using BECAUSE.",
      "Write the final sentence in your warm-up book, not your whiteboard. Four minutes.",
    ],
    do: [
      "Set a timer for 4 minutes.",
      "Circulate. For stuck students, read both their hook ideas aloud and ask 'which one is the reason?'",
      "Note 1-2 strong sentences to share before the exit.",
      "Reveal the sentence stem ONLY if a focus student is stuck. Do not reveal to the class by default.",
    ],
    enabling: [
      "FOCUS GROUP: if their hook ideas don't connect with cause-and-effect, give them a fresh pair orally - 'we played our first game' / 'I was excited' - and let them combine that with BECAUSE.",
      "Click reveal to show the sentence stem for individual focus students - point to it on the screen for them, don't announce it to the class.",
    ],
    extension: [
      "HIGHER-ABILITY EXTENSION: choose ONE of these and write it in your book under your main sentence:",
      "  (a) Write TWO different valid versions of your sentence (e.g. BECAUSE in the middle, then BECAUSE at the start with a comma). Write one line saying which is clearer and why.",
      "  (b) Add a WHEN or WHERE detail to one of your clauses (e.g. 'We lost the netball game on Friday BECAUSE our goal shooter was away sick.').",
      "  (c) Use a rarer subordinator instead of BECAUSE - try SINCE or AS - and check it still makes the same kind of sense.",
    ],
    watchFor: [
      "Students writing two separate sentences with a full stop in the middle - redirect: ONE sentence, BECAUSE in the middle.",
      "Students with hook ideas that don't have a clear cause-effect connection - give them the fallback pair quickly so they still write something.",
      "Comma splice (comma instead of BECAUSE) - mark as developmental progress, not failure: 'you've spotted the link, now add the joining word'.",
    ],
    tag: "[Day 1 | Mon | Independent BECAUSE - build, stem hidden]",
  }, { requireSay: false, requireDo: false }),
  notesReveal: composeNotes({
    say: [
      "If you need a start, here's a sentence frame. Fill in the blanks.",
    ],
    do: [
      "Use this reveal for individual focus students. Walk to them, point to the screen, then walk away.",
      "Do not read the stem aloud to the whole class - it lowers expectations for the rest.",
    ],
    tag: "[Day 1 | Mon | Independent BECAUSE - stem reveal]",
  }, { requireSay: false, requireDo: false }),
});

exitSlide({
  dayText: MON.dayText, dayColor: MON.dayColor,
  joiningWord: "BECAUSE",
  taskHeader: "Today's exit check",
  taskIdeas: [
    "Our team won the relay.",
    "We practised the changeovers all term.",
  ],
  markCriteria: MARK_CRITERIA("BECAUSE"),
  qrLabel: "Mon BECAUSE form",
  notes: composeNotes({
    say: [
      "Last task. Combine these two ideas into ONE sentence using BECAUSE.",
      "Type your sentence into the Google Form on your iPad - scan the QR. Two minutes.",
      "If your iPad isn't working, write your sentence on a paper slip and drop it in the basket.",
    ],
    do: [
      "Set a timer for 2 minutes.",
      "Walk the room. Use the marking criteria on the right.",
      "Open the Google Form sheet view on your laptop - watch responses populate live.",
      "Quick share before pack-up: read out one strong example.",
    ],
    teacherNotes: [
      "Expected answer: 'Our team won the relay BECAUSE we practised the changeovers all term.' Reduction would say: 'Our team won the relay BECAUSE we practised changeovers all term' (drop 'all term' or 'the') - both fine. Mark on meaning + grammar, not exact wording.",
      "Today's exit primarily checks criterion 3 (appropriate clause linking) and criterion 5 (no comma splice).",
      "Backup: if the Google Form goes down, paper slips in the basket. Mark them in the same way and add to the sheet view manually at lunch.",
    ],
    watchFor: [
      "Comma instead of BECAUSE - comma splice. Cross criterion 5, frame as developmental progress.",
      "Two separate sentences with a full stop - cross criteria 2 and 3, but note that the student has split the ideas correctly.",
      "Students who reverse the cause and effect (e.g. 'We practised the changeovers all term BECAUSE our team won the relay') - cross criterion 3, reteach quickly.",
    ],
    tag: "[Day 1 | Mon | Exit BECAUSE - QR submit]",
  }, { requireSay: false, requireDo: false }),
});

/* ─────────────────────────────────────────────────────────────────────────
 *  Write file
 * ───────────────────────────────────────────────────────────────────────── */

const outFile = path.join(OUT_DIR, "SC_T2W5_Lesson1_Mon_BECAUSE.pptx");
pres.writeFile({ fileName: outFile }).then(() => {
  console.log("PPTX written to " + outFile);
}).catch((err) => {
  console.error("Build failed:", err);
  process.exitCode = 1;
});
