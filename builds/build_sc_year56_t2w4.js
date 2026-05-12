"use strict";

/**
 * Sentence Combining Warm-Up — Year 5/6 Enrichment, Term 2 Week 4
 *
 * 5-day, 15-minute daily warm-up sequence introducing coordinating
 * conjunctions (AND, BUT, SO, OR) one per day, with mixed practice on Friday.
 *
 * Daily routine: Day intro / Review (2 min) -> Teach worked + non-example (3 min)
 *  -> Guided practice with reveal (4 min) -> Independent production (4 min)
 *  -> Exit check (2 min).
 *
 * 6 slides per day x 5 = 30, plus title + resources = 32. Reveal pairs count
 * as two physical slides.
 */

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant, composeNotes } = require("../themes/factory");

// Term 2 Week 4 -> variant 3
const T = createTheme("literacy", "grade56", weekToVariant(4));
const {
  C, FONT_H, FONT_B,
  SLIDE_W, SLIDE_H, SAFE_BOTTOM, CONTENT_TOP,
  addCard, addFooter, addTopBar, addBadge, addTitle, addTextOnShape,
  titleSlide, withReveal, runSlideDiagnostics,
} = T;

const FOOTER = "Sentence Combining Warm-Up | Year 5/6 | Term 2 Week 4";
const OUT_DIR = "output/Sentence_Combining_Year56_T2W4";
fs.mkdirSync(OUT_DIR, { recursive: true });

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Sentence Combining Warm-Up — Year 5/6 — Term 2 Week 4";
pres.author = "Year 5/6 Enrichment PLC";

/* ─────────────────────────────────────────────────────────────────────────
 *  Shared success criteria (rotates per day with the joining word slotted in)
 * ───────────────────────────────────────────────────────────────────────── */

const SC_BASE = (joiningWord) => ([
  "I can find two short ideas that go together.",
  `I can use ${joiningWord} to join them into ONE sentence.`,
  "I can write a sentence that makes sense.",
]);

/* Shared marking strip — what teachers tick on the exit check */
const MARK_CRITERIA = (joiningWord) => ([
  "Two ideas combined into ONE sentence",
  `Used ${joiningWord} (correct joining word)`,
  "Sentence is grammatical (no run-on, no comma splice)",
]);

/* ─────────────────────────────────────────────────────────────────────────
 *  Phase + day badge stack (top-left of every slide)
 * ───────────────────────────────────────────────────────────────────────── */

function badgeStack(s, phaseText, phaseColor, dayText, dayColor) {
  // phase badge first
  addBadge(s, phaseText, { color: phaseColor, w: 1.6, x: 0.5, y: 0.20 });
  // day badge second (e.g. "Mon")
  addBadge(s, dayText, { color: dayColor || C.SECONDARY, w: 1.05, x: 2.18, y: 0.20 });
}

/* ─────────────────────────────────────────────────────────────────────────
 *  Helper: highlight a joining word inside a sentence.
 *  Returns a mixed text-runs array for PptxGenJS.
 * ───────────────────────────────────────────────────────────────────────── */

function highlightSentence(sentence, joiningWord, baseSize, baseColor, highlightColor) {
  // Wrap joining word with spaces to match it cleanly when adjacent to other words.
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
 *  Builder: Day intro / Review
 *  Left card: hook (Day 1) or yesterday's joining word review (Days 2–5)
 *  Right card: today's success criteria (plain unlabelled "I can…" list)
 * ───────────────────────────────────────────────────────────────────────── */

function dayIntroSlide({ dayText, dayColor, joiningWord, hookHeader, hookBody, scItems, notes }) {
  const s = pres.addSlide();
  addTopBar(s, C.PRIMARY);
  badgeStack(s, "Day Start", C.PRIMARY, dayText, dayColor);
  addTitle(s, `${dayText} — Joining word: ${joiningWord}`);

  const cardH = SAFE_BOTTOM - CONTENT_TOP;
  const leftW = 4.6;
  const rightW = 4.4;
  const rightX = 5.2;

  // Left: hook / review
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

  // Right: SC card
  addCard(s, rightX, CONTENT_TOP, rightW, cardH, { strip: C.ACCENT, fill: C.BG_CARD });
  s.addText("Today's success criteria — I can…", {
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

/* ─────────────────────────────────────────────────────────────────────────
 *  Builder: Teach (worked example on top, non-example beneath)
 *  Worked example shows: "Two short ideas" -> arrow -> "One combined sentence"
 *  with the joining word highlighted, plus a one-line "why it works".
 * ───────────────────────────────────────────────────────────────────────── */

function teachSlide({ dayText, dayColor, joiningWord, meaning, twoIdeas, combined, whyItWorks, nonExample, notes }) {
  const s = pres.addSlide();
  addTopBar(s, C.PRIMARY);
  badgeStack(s, "Teach", C.PRIMARY, dayText, dayColor);
  addTitle(s, `How "${joiningWord}" works — ${meaning}`);

  // ── Worked example card ───────────────────────────────────────────────
  const workedY = CONTENT_TOP;
  const workedH = 2.55;
  addCard(s, 0.5, workedY, 9, workedH, { strip: C.PRIMARY, fill: C.WHITE });

  // Header row inside worked card
  s.addText("Worked example", {
    x: 0.75, y: workedY + 0.10, w: 4, h: 0.28,
    fontSize: 12, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
  });

  // "Two short ideas" subhead
  s.addText("Two short ideas:", {
    x: 0.75, y: workedY + 0.42, w: 4, h: 0.26,
    fontSize: 11, fontFace: FONT_B, color: C.MUTED, bold: true, margin: 0,
  });

  // Two ideas stacked
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

  // Arrow (small chevron-style) — text arrow keeps it editable
  s.addText("⬇", {
    x: 0.75, y: ideaY + 0.74, w: 0.6, h: 0.32,
    fontSize: 18, fontFace: FONT_H, color: C.ACCENT, bold: true, margin: 0,
  });
  s.addText(`Joined with "${joiningWord}":`, {
    x: 1.45, y: ideaY + 0.74, w: 4.9, h: 0.32,
    fontSize: 11, fontFace: FONT_B, color: C.MUTED, bold: true, margin: 0,
  });

  // Combined sentence with joining word highlighted
  const combinedRuns = highlightSentence(combined, joiningWord, 19, C.CHARCOAL, C.ACCENT);
  s.addText(combinedRuns.map((run, i) => ({
    text: `${i === 0 ? '"' : ""}${run.text}${i === combinedRuns.length - 1 ? '"' : ""}`,
    options: { ...run.options, fontFace: FONT_H, breakLine: false },
  })), {
    x: 0.85, y: ideaY + 1.10, w: 8.0, h: 0.42,
    valign: "middle", margin: 0,
    fit: "shrink", shrinkText: true,
  });

  // Why it works strip
  const whyY = workedY + workedH - 0.34;
  s.addShape("rect", {
    x: 0.5, y: whyY - 0.02, w: 9, h: 0.36, fill: { color: C.BG_CARD },
  });
  s.addText("Why this works: " + whyItWorks, {
    x: 0.75, y: whyY, w: 8.5, h: 0.32,
    fontSize: 12, fontFace: FONT_B, color: C.SECONDARY, bold: true, margin: 0,
    fit: "shrink", shrinkText: true,
  });

  // ── Non-example card ─────────────────────────────────────────────────
  const neY = workedY + workedH + 0.14;
  const neH = SAFE_BOTTOM - neY;
  addCard(s, 0.5, neY, 9, neH, { strip: C.ALERT, fill: C.WHITE });

  s.addText("Watch out — common mistake", {
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

/* ─────────────────────────────────────────────────────────────────────────
 *  Builder: Guided practice with reveal pair
 *  Two sentence pairs to combine; reveal shows both combined answers.
 * ───────────────────────────────────────────────────────────────────────── */

function guidedSlide({ dayText, dayColor, joiningWord, pairs, answers, notesBuild, notesReveal }) {
  return withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.SECONDARY);
      badgeStack(s, "Practise together", C.SECONDARY, dayText, dayColor);
      addTitle(s, "Combine each pair into ONE sentence");

      // Whiteboard cue (small, top-right)
      addTextOnShape(s, "Mini-whiteboards", {
        x: 7.4, y: 0.20, w: 2.1, h: 0.36, rectRadius: 0.06,
        fill: { color: C.WHITE },
        line: { color: C.SECONDARY, width: 1.0 },
      }, {
        fontSize: 11, fontFace: FONT_B, color: C.SECONDARY, bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      // Two practice cards stacked
      const cardW = 9;
      const gap = 0.16;
      const pairsArea = 2.55; // leave room for reveal bar
      const cardH = (pairsArea - gap) / 2;

      pairs.forEach((pair, i) => {
        const y = CONTENT_TOP + i * (cardH + gap);
        addCard(s, 0.5, y, cardW, cardH, {
          strip: i === 0 ? C.PRIMARY : C.ACCENT,
          fill: C.WHITE,
        });
        s.addText(`Try this — pair ${i + 1}`, {
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

      // Reveal stub — placeholder bar so the build slide stays balanced
      // even before the click-to-reveal happens
      const stubY = CONTENT_TOP + pairsArea + 0.10;
      const stubH = SAFE_BOTTOM - stubY;
      s.addShape("roundRect", {
        x: 0.5, y: stubY, w: 9, h: stubH, rectRadius: 0.10,
        fill: { color: C.BG_CARD },
        line: { color: C.MUTED, width: 0.8, dashType: "dash" },
      });
      s.addText(`Try first on whiteboards. Click to reveal answers using "${joiningWord}".`, {
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
      // Reveal: replace the placeholder strip with the answer bar showing both combined sentences
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

      // Two answers stacked inside the reveal bar
      const ansListY = stubY + 0.30;
      const perAnsH = (stubH - 0.36) / 2;
      answers.forEach((ans, i) => {
        const y = ansListY + i * perAnsH;
        const runs = highlightSentence(ans, joiningWord, 14, C.WHITE, C.BG_CARD);
        s.addText(runs.map((run, j) => ({
          text: `${j === 0 ? `${i + 1}.  "` : ""}${run.text}${j === runs.length - 1 ? '"' : ""}`,
          options: { ...run.options, fontFace: FONT_H, bold: j === 1 ? true : true, breakLine: false },
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

/* ─────────────────────────────────────────────────────────────────────────
 *  Builder: Independent production
 *  Single hero prompt — students write ONE original sentence.
 * ───────────────────────────────────────────────────────────────────────── */

function independentSlide({ dayText, dayColor, joiningWord, prompt, sentenceStarter, notes }) {
  const s = pres.addSlide();
  addTopBar(s, C.ACCENT);
  badgeStack(s, "Your turn", C.ACCENT, dayText, dayColor);
  addTitle(s, `Write ONE sentence using "${joiningWord}"`);

  // Hero prompt card
  const promptY = CONTENT_TOP;
  const promptH = 2.30;
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

  // Sentence starter strip
  const starterY = promptY + promptH + 0.14;
  const starterH = 0.62;
  s.addShape("roundRect", {
    x: 0.5, y: starterY, w: 9, h: starterH, rectRadius: 0.08,
    fill: { color: C.BG_CARD },
    line: { color: C.SECONDARY, width: 0.8 },
  });
  s.addText("Need a start? Try: ", {
    x: 0.75, y: starterY, w: 1.85, h: starterH,
    fontSize: 12, fontFace: FONT_B, color: C.SECONDARY, bold: true,
    valign: "middle", margin: 0,
  });
  s.addText(`"${sentenceStarter}"`, {
    x: 2.6, y: starterY, w: 6.7, h: starterH,
    fontSize: 14.5, fontFace: FONT_H, color: C.CHARCOAL, italic: true,
    valign: "middle", margin: 0, fit: "shrink",
  });

  // Reminder strip — must combine TWO ideas using the joining word
  const remindY = starterY + starterH + 0.12;
  const remindH = SAFE_BOTTOM - remindY;
  s.addShape("roundRect", {
    x: 0.5, y: remindY, w: 9, h: remindH, rectRadius: 0.08,
    fill: { color: C.PRIMARY },
  });
  s.addText(`Remember: ONE sentence, with TWO ideas joined by "${joiningWord}".`, {
    x: 0.75, y: remindY, w: 8.5, h: remindH,
    fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true,
    align: "center", valign: "middle", margin: 0,
  });

  addFooter(s, FOOTER);
  if (notes) s.addNotes(notes);
  runSlideDiagnostics(s, pres, { respectSafeBottom: false });
  return s;
}

/* ─────────────────────────────────────────────────────────────────────────
 *  Builder: Exit check
 *  Left card: the exit task. Right card: marking criteria (what teachers tick).
 * ───────────────────────────────────────────────────────────────────────── */

function exitSlide({ dayText, dayColor, joiningWord, taskHeader, taskIdeas, markCriteria, notes }) {
  const s = pres.addSlide();
  addTopBar(s, C.ALERT);
  badgeStack(s, "Exit check", C.ALERT, dayText, dayColor);

  // CHECK wordmark (top-right) — accessibility signal beyond colour
  addTextOnShape(s, "✓  EXIT", {
    x: 8.0, y: 0.20, w: 1.5, h: 0.36, rectRadius: 0.06,
    fill: { color: C.WHITE },
    line: { color: C.ALERT, width: 1.5 },
  }, {
    fontSize: 11, fontFace: FONT_B, color: C.ALERT, bold: true,
    align: "center", valign: "middle", margin: 0,
  });

  addTitle(s, "Show what you know");

  const cardH = SAFE_BOTTOM - CONTENT_TOP;
  const leftW = 5.6;
  const rightX = 6.3;
  const rightW = 3.2;

  // Left: task
  addCard(s, 0.5, CONTENT_TOP, leftW, cardH, { strip: C.ALERT, fill: C.WHITE });
  s.addText(taskHeader, {
    x: 0.75, y: CONTENT_TOP + 0.14, w: leftW - 0.4, h: 0.32,
    fontSize: 12, fontFace: FONT_B, color: C.ALERT, bold: true, margin: 0,
  });

  // Two short ideas stacked — students combine these into one
  const ideaY = CONTENT_TOP + 0.50;
  s.addText(`"${taskIdeas[0]}"`, {
    x: 0.75, y: ideaY, w: leftW - 0.4, h: 0.42,
    fontSize: 18, fontFace: FONT_H, color: C.CHARCOAL, italic: true, margin: 0,
    fit: "shrink", shrinkText: true,
  });
  s.addText(`"${taskIdeas[1]}"`, {
    x: 0.75, y: ideaY + 0.46, w: leftW - 0.4, h: 0.42,
    fontSize: 18, fontFace: FONT_H, color: C.CHARCOAL, italic: true, margin: 0,
    fit: "shrink", shrinkText: true,
  });

  // Cue strip
  s.addText(`Combine into ONE sentence using "${joiningWord}".`, {
    x: 0.75, y: ideaY + 1.05, w: leftW - 0.4, h: 0.36,
    fontSize: 14, fontFace: FONT_B, color: C.SECONDARY, bold: true, margin: 0,
  });

  // Workspace hint
  s.addText("Write your sentence in your warm-up book.", {
    x: 0.75, y: CONTENT_TOP + cardH - 0.40, w: leftW - 0.4, h: 0.28,
    fontSize: 11, fontFace: FONT_B, color: C.MUTED, italic: true, margin: 0,
  });

  // Right: marking criteria
  addCard(s, rightX, CONTENT_TOP, rightW, cardH, { strip: C.SUCCESS, fill: C.BG_CARD });
  s.addText("Tick if students…", {
    x: rightX + 0.18, y: CONTENT_TOP + 0.14, w: rightW - 0.36, h: 0.32,
    fontSize: 12, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
  });

  // Render checkbox + criterion rows
  const rowStart = CONTENT_TOP + 0.50;
  const rowH = 0.66;
  markCriteria.forEach((crit, i) => {
    const y = rowStart + i * (rowH + 0.08);
    // checkbox
    s.addShape("roundRect", {
      x: rightX + 0.18, y: y + 0.10, w: 0.32, h: 0.32, rectRadius: 0.04,
      fill: { color: C.WHITE },
      line: { color: C.SUCCESS, width: 1.2 },
    });
    s.addText(crit, {
      x: rightX + 0.58, y, w: rightW - 0.76, h: rowH,
      fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });
  });

  // Tracker reminder
  s.addText("Record on the cohort tracker.", {
    x: rightX + 0.18, y: CONTENT_TOP + cardH - 0.38, w: rightW - 0.36, h: 0.26,
    fontSize: 10.5, fontFace: FONT_B, color: C.MUTED, italic: true, margin: 0,
  });

  addFooter(s, FOOTER);
  if (notes) s.addNotes(notes);
  runSlideDiagnostics(s, pres, { respectSafeBottom: false });
  return s;
}

/* ─────────────────────────────────────────────────────────────────────────
 *  Custom: Title slide (uses base titleSlide builder)
 * ───────────────────────────────────────────────────────────────────────── */

titleSlide(
  pres,
  "Sentence Combining Warm-Up",
  "5 days x 15 minutes — joining two short ideas into one sentence",
  "Year 5/6 Enrichment  |  Term 2 Week 4  |  PLC Cycle Week 1 of 6",
  composeNotes({
    say: [
      "This week we're learning to join two short ideas into ONE sentence.",
      "Each day takes 15 minutes. Same routine: review, teach, practise together, write your own, exit check.",
      "We'll use a different joining word each day — AND, BUT, SO, OR — and a mixed practice on Friday.",
    ],
    do: [
      "Open the deck on Day 1 and use the day labels at the top of each slide to navigate.",
      "Have warm-up books and mini-whiteboards ready before students arrive.",
      "Have the cohort tracker open to record exit-check ticks daily.",
    ],
    teacherNotes: [
      "Goal: lift the cohort from 63% to 70% combining ideas into ONE grammatical sentence by Week 8.",
      "Same 5-day routine repeats across the 6-week PLC cycle. The 4 conjunctions cycle each week.",
    ],
    tag: "[Sentence Combining | Year 5/6 | T2 W4]",
  }, { requireSay: false, requireDo: false })
);

/* ─────────────────────────────────────────────────────────────────────────
 *  Teacher Resources slide (immediately after title — megaprompt §0a item 19)
 * ───────────────────────────────────────────────────────────────────────── */

function teacherResourcesSlide() {
  const s = pres.addSlide();
  addTopBar(s, C.SECONDARY);
  addBadge(s, "Teacher Resources", { color: C.SECONDARY, w: 2.4 });
  addTitle(s, "What you need before you start");

  const cardH = SAFE_BOTTOM - CONTENT_TOP;
  const leftW = 4.4;
  const rightW = 4.6;
  const rightX = 5.0;

  // Left: classroom materials
  addCard(s, 0.5, CONTENT_TOP, leftW, cardH, { strip: C.PRIMARY, fill: C.WHITE });
  s.addText("In the room", {
    x: 0.7, y: CONTENT_TOP + 0.14, w: leftW - 0.4, h: 0.32,
    fontSize: 13, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
  });
  const materials = [
    "Mini-whiteboards + markers (one per student)",
    "Warm-up book / writing book per student",
    "Cohort exit-check tracker (digital or printed)",
    "This deck open on the projector",
  ];
  s.addText(materials.map((t, i) => ({
    text: t,
    options: {
      bullet: true,
      breakLine: i < materials.length - 1,
      fontSize: 14, color: C.CHARCOAL,
      paraSpaceAfter: 6,
    },
  })), {
    x: 0.7, y: CONTENT_TOP + 0.50, w: leftW - 0.4, h: cardH - 0.66,
    fontFace: FONT_B, valign: "top", margin: 0,
    fit: "shrink", shrinkText: true,
  });

  // Right: no printed worksheet
  addCard(s, rightX, CONTENT_TOP, rightW, cardH, { strip: C.ACCENT, fill: C.BG_CARD });
  s.addText("No printed student resources required", {
    x: rightX + 0.20, y: CONTENT_TOP + 0.14, w: rightW - 0.4, h: 0.32,
    fontSize: 13, fontFace: FONT_B, color: C.ACCENT, bold: true, margin: 0,
  });
  const teacherInfo = [
    "All worked examples and prompts are on the slides.",
    "Students respond on whiteboards (guided practice) and in their warm-up book (independent + exit).",
    "No printing needed for daily warm-up.",
    "Differentiation: focus group prompt and extension prompt are in the speaker notes for each day.",
    "The exit check slide for each day shows the marking criteria — tick on the cohort tracker.",
  ];
  s.addText(teacherInfo.map((t, i) => ({
    text: t,
    options: {
      bullet: true,
      breakLine: i < teacherInfo.length - 1,
      fontSize: 13, color: C.CHARCOAL,
      paraSpaceAfter: 5,
    },
  })), {
    x: rightX + 0.20, y: CONTENT_TOP + 0.50, w: rightW - 0.4, h: cardH - 0.66,
    fontFace: FONT_B, valign: "top", margin: 0,
    fit: "shrink", shrinkText: true,
  });

  addFooter(s, FOOTER);
  s.addNotes(composeNotes({
    say: [
      "This is the warm-up resources slide. Quick check: whiteboards, warm-up books, tracker.",
      "Same routine every day — teachers can swap with each other if needed.",
    ],
    do: [
      "Distribute whiteboards and markers as students enter.",
      "Open the cohort tracker before starting Day 1.",
    ],
    teacherNotes: [
      "Same 5 minutes of setup applies all week. The deck is self-contained — no printing or pre-reading needed.",
    ],
    tag: "[Sentence Combining | Resources]",
  }, { requireSay: false, requireDo: false }));
  runSlideDiagnostics(s, pres, { respectSafeBottom: false });
  return s;
}

teacherResourcesSlide();

/* ─────────────────────────────────────────────────────────────────────────
 *  Framing slide: "What this week builds" — placed between Resources and
 *  Monday so any teacher in the team understands the destination before
 *  teaching Day 1. Maps Week 4 -> Week 8 progression.
 * ───────────────────────────────────────────────────────────────────────── */

function weekFramingSlide() {
  const s = pres.addSlide();
  addTopBar(s, C.PRIMARY);
  addBadge(s, "PLC framing", { color: C.PRIMARY, w: 2.0 });
  addTitle(s, "What this week builds");

  const cardH = SAFE_BOTTOM - CONTENT_TOP;
  const leftW = 4.6;
  const rightW = 4.4;
  const rightX = 5.2;

  // Left card: the destination + the why
  addCard(s, 0.5, CONTENT_TOP, leftW, cardH, { strip: C.ACCENT, fill: C.WHITE });
  s.addText("The destination", {
    x: 0.7, y: CONTENT_TOP + 0.14, w: leftW - 0.4, h: 0.32,
    fontSize: 13, fontFace: FONT_B, color: C.ACCENT, bold: true, margin: 0,
  });
  s.addText(
    "Week 1 of 6. The destination is the Q3 post-test in Week 8: " +
    "students combining 5 short ideas into ONE grammatical sentence.",
    {
      x: 0.7, y: CONTENT_TOP + 0.50, w: leftW - 0.4, h: 1.30,
      fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, valign: "top", margin: 0,
      paraSpaceAfter: 4,
    },
  );

  // "If a student asks why" strip
  const whyY = CONTENT_TOP + cardH - 1.50;
  s.addText("If a student asks 'why are we doing this?'", {
    x: 0.7, y: whyY, w: leftW - 0.4, h: 0.30,
    fontSize: 12, fontFace: FONT_B, color: C.SECONDARY, bold: true, margin: 0,
  });
  s.addText(
    "Better writers control their sentences. We're building one piece of that " +
    "control each week.",
    {
      x: 0.7, y: whyY + 0.34, w: leftW - 0.4, h: 1.06,
      fontSize: 13, fontFace: FONT_H, color: C.CHARCOAL, italic: true,
      valign: "top", margin: 0,
    },
  );

  // Right card: 5-week progression as labelled rows
  addCard(s, rightX, CONTENT_TOP, rightW, cardH, { strip: C.PRIMARY, fill: C.BG_CARD });
  s.addText("The 5-week build", {
    x: rightX + 0.20, y: CONTENT_TOP + 0.14, w: rightW - 0.4, h: 0.32,
    fontSize: 13, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
  });

  const weeks = [
    { label: "Week 4 (now)", text: "Vocabulary: AND, BUT, SO, OR.", strip: C.PRIMARY },
    { label: "Week 5",       text: "Adds: because, when, if.",      strip: C.SECONDARY },
    { label: "Week 6",       text: "Full Q3 task — combining 3-5 sentences with deletion and embedding.", strip: C.ACCENT },
    { label: "Week 7",       text: "Applies it to students' own writing.", strip: C.ALERT },
    { label: "Week 8",       text: "Q3 post-test.",                 strip: C.SUCCESS },
  ];

  const rowsTop = CONTENT_TOP + 0.50;
  const rowH = (cardH - 0.62) / weeks.length;
  weeks.forEach((w, i) => {
    const y = rowsTop + i * rowH;
    // Tiny coloured strip + week label
    s.addShape("rect", {
      x: rightX + 0.18, y: y + 0.04, w: 0.10, h: rowH - 0.10,
      fill: { color: w.strip },
    });
    s.addText(w.label, {
      x: rightX + 0.36, y: y + 0.02, w: rightW - 0.56, h: 0.24,
      fontSize: 11, fontFace: FONT_B, color: w.strip, bold: true, margin: 0,
    });
    s.addText(w.text, {
      x: rightX + 0.36, y: y + 0.26, w: rightW - 0.56, h: rowH - 0.32,
      fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, valign: "top", margin: 0,
      fit: "shrink", shrinkText: true,
    });
  });

  addFooter(s, FOOTER);
  s.addNotes(composeNotes({
    say: [
      "Quick framing before we start — this is a 6-week build, not a one-week activity.",
      "Right now, Week 4, we're teaching the four basic joining words: AND, BUT, SO, OR.",
      "Week 5 adds because, when, if. Week 6 is the full Q3 task — combining 3-5 sentences. Week 7 applies it to students' own writing. Week 8 is the post-test.",
      "If a student asks 'why are we doing this?' — better writers control their sentences. We're building one piece of that control each week.",
    ],
    do: [
      "Read this slide once before Day 1. It is for you, not students.",
      "If a teacher in the team is unclear on the destination, point them here.",
    ],
    teacherNotes: [
      "Pre-test data: 63% of cohort can combine ideas into ONE grammatical sentence. Goal by Week 8: 70%.",
      "This warm-up alone won't get the cohort there — Weeks 6 and 7 do the heavier lifting on deletion and embedding. Week 4 builds the joining-word vocabulary that those weeks need.",
    ],
    tag: "[PLC framing | Week 4 of 6]",
  }, { requireSay: false, requireDo: false }));
  runSlideDiagnostics(s, pres, { respectSafeBottom: false });
  return s;
}

weekFramingSlide();

/* =========================================================================
 *  DAY 1 — MONDAY — AND (joining two related ideas) — topic: camp
 * ========================================================================= */

const MON = { dayText: "Mon", dayColor: C.PRIMARY };

dayIntroSlide({
  dayText: MON.dayText, dayColor: MON.dayColor,
  joiningWord: "AND",
  hookHeader: "Hook — quick start (Day 1)",
  hookBody:
    "On your whiteboards, write TWO short things you remember about camp. " +
    "Just two short ideas, like 'We swam in the river.' and 'I shared a tent with Sam.' " +
    "We will join them later.",
  scItems: SC_BASE("AND"),
  notes: composeNotes({
    say: [
      "Welcome to our new warm-up. For one week we're going to practise joining two short ideas into ONE sentence.",
      "Today's joining word is AND. We use AND when two ideas go together.",
      "On your whiteboards, write TWO short things you remember about camp. Two short ideas, full stops, that's it.",
    ],
    do: [
      "Hand out whiteboards as students arrive.",
      "Time the hook: 90 seconds maximum, then move on.",
      "Scan boards as you walk — note one or two pairs to use later if you want.",
    ],
    teacherNotes: [
      "This is the FIRST day of the cycle, so use a hook instead of a review. From Tuesday onwards this slide reviews yesterday's joining word.",
      "Topic 'camp' lowers cognitive load so students think about the sentence structure, not the content.",
    ],
    enabling: [
      "FOCUS GROUP (15 students, 6-12 months below): give them two starter sentences orally — 'Camp is fun.' / 'I swam in the river.' — and just ask them to write those down.",
      "EXTENSION: ask higher-ability students to add an adjective or specific detail to each short idea before the next slide.",
    ],
    watchFor: [
      "Students writing one long sentence — redirect to TWO short ideas with full stops.",
    ],
    tag: "[Day 1 | Mon | Hook + LI/SC]",
  }, { requireSay: false, requireDo: false }),
});

teachSlide({
  dayText: MON.dayText, dayColor: MON.dayColor,
  joiningWord: "AND",
  meaning: "joining ideas",
  twoIdeas: [
    "We pitched the tents.",
    "We collected wood for the fire.",
  ],
  // Combined demonstrates REDUCTION: same subject ('we'), so we drop the
  // second one after AND. Models the move students need for the Q3 task.
  combined: "We pitched the tents AND collected wood for the fire.",
  whyItWorks:
    "Both ideas share the SAME subject ('we') — drop the repeat after AND. ONE shorter, stronger sentence.",
  nonExample: {
    wrong: "We pitched the tents AND.",
    why:
      "AND must join TWO ideas. There's only one idea here, so AND has nothing to join. " +
      "Make sure each side of AND is a full idea (subject + verb).",
  },
  notes: composeNotes({
    say: [
      "Here are TWO short ideas. Both are about setting up camp — they go together.",
      "Watch how I join them: I drop the full stop, lowercase the next word, and put AND in between.",
      "Both ideas start with 'we'. I can drop the second 'we' to make it shorter and tighter: 'We pitched the tents AND collected wood for the fire.' ONE subject, two actions.",
      "Now the wrong example. AND has nothing to join — it's left dangling.",
    ],
    do: [
      "Track with your finger from idea 1 down to idea 2, then point at the combined sentence.",
      "Highlight AND in the combined sentence and physically cross out the second 'we' on the board so the reduction move is visible.",
      "Read the non-example aloud and ask: 'What's missing after AND?'",
    ],
    teacherNotes: [
      "Two moves today: (1) JOIN with AND. (2) REDUCE by dropping the repeated subject. The reduction is what students need for the Q3-style task.",
      "Both sides of AND must still be a full idea OR be the second half of a same-subject pair where the subject has been dropped.",
    ],
    watchFor: [
      "Students who keep both 'we's — accept it (still grammatical), but model the reduced version once more.",
      "Students who drop AND as well as the second subject — push them to keep AND as the joiner.",
    ],
    tag: "[Day 1 | Mon | Teach AND - join + reduce]",
  }, { requireSay: false, requireDo: false }),
});

guidedSlide({
  dayText: MON.dayText, dayColor: MON.dayColor,
  joiningWord: "AND",
  pairs: [
    ["The bus left at 8 am.", "We arrived at camp by lunch."],
    ["I packed my torch.", "I packed my raincoat."],
  ],
  answers: [
    "The bus left at 8 am AND we arrived at camp by lunch.",
    "I packed my torch AND my raincoat.",
  ],
  notesBuild: composeNotes({
    say: [
      "Together. Two pairs to combine using AND.",
      "Try pair 1 on your whiteboard. Then pair 2. You have about 60 seconds for each.",
      "Think: are both ideas full ideas? Then put AND between them.",
    ],
    do: [
      "PROTOCOL: students write on whiteboards FIRST, hold up, THEN click to reveal. Do not click reveal until every student has attempted. If kids learn to wait, the practice becomes recognition not production.",
      "Read pair 1 aloud, pause for 60 seconds, scan boards.",
      "Then pair 2, same routine.",
      "Click to reveal both answers together so students can self-check.",
    ],
    teacherNotes: [
      "Pair 1 has DIFFERENT subjects ('the bus' / 'we'), so we keep both. Reduction won't fire here.",
      "Pair 2 has the SAME subject ('I') — it's the reduction case from the Teach slide. Both 'I packed my torch AND my raincoat' (reduced) and 'I packed my torch AND I packed my raincoat' (unreduced) are acceptable, but call out the reduced version as the stronger move.",
    ],
    watchFor: [
      "Comma splice creeping in — gently redirect to AND with no comma needed for these short pairs.",
      "Students treating Pair 2 like Pair 1 and keeping the second 'I' — accept it, but model the reduction once more.",
    ],
    tag: "[Day 1 | Mon | Guided AND - build + reduce]",
  }, { requireSay: false, requireDo: false }),
  notesReveal: composeNotes({
    say: [
      "Here are the answers. Tick if yours matches the meaning, even if your wording is slightly different.",
      "Look at pair 2 — the answer drops the second 'I packed' because the subject is the same. That's the reduction move from the Teach slide.",
    ],
    do: [
      "Read both answers aloud. Highlight AND in each.",
      "On pair 2, point at the missing 'I packed' and name it: 'we dropped this because both ideas had the same subject'.",
      "If 50% or more got both right, move to independent. If less, do one more pair from the board.",
    ],
    tag: "[Day 1 | Mon | Guided AND - reveal + reduce]",
  }, { requireSay: false, requireDo: false }),
});

independentSlide({
  dayText: MON.dayText, dayColor: MON.dayColor,
  joiningWord: "AND",
  prompt:
    "Write ONE sentence about your favourite thing about camp. " +
    "Use AND to join two ideas.",
  sentenceStarter: "My favourite thing about camp was ___ AND ___.",
  notes: composeNotes({
    say: [
      "Now your turn. ONE sentence about your favourite thing about camp.",
      "Use AND to join two ideas. The sentence starter is on the screen if you need it.",
      "You have four minutes. Write neatly in your warm-up book.",
    ],
    do: [
      "Set a timer for 4 minutes.",
      "Circulate. Tap students who are stuck and read the starter aloud with them.",
      "Note 1-2 strong sentences to share before the exit check.",
    ],
    enabling: [
      "FOCUS GROUP: give them the sentence starter to copy and complete in two blanks. Allow them to use a class shared experience like 'the games night' if they can't remember a personal one.",
      "EXTENSION: challenge higher-ability students to write a sentence with TWO subjects joined by AND (e.g. 'My friend AND I both...') or to add a where/when detail in each half.",
    ],
    watchFor: [
      "Students writing two separate sentences — redirect: ONE sentence, joined by AND.",
      "Students using AND to join an adjective list ('hot and sunny') — this still counts as ONE idea. Push them to combine TWO full ideas instead.",
    ],
    tag: "[Day 1 | Mon | Independent AND]",
  }, { requireSay: false, requireDo: false }),
});

exitSlide({
  dayText: MON.dayText, dayColor: MON.dayColor,
  joiningWord: "AND",
  taskHeader: "Today's exit check",
  taskIdeas: [
    "We hiked up the hill.",
    "We saw the view from the top.",
  ],
  markCriteria: MARK_CRITERIA("AND"),
  notes: composeNotes({
    say: [
      "Last task. In your warm-up book, write these two ideas as ONE sentence using AND.",
      "Two minutes.",
    ],
    do: [
      "Set a timer for 2 minutes.",
      "Walk the room. Use the marking criteria on the right. Tick or cross on the cohort tracker as you go.",
      "Quick share: read out one strong example before students pack up.",
    ],
    teacherNotes: [
      "Expected answer: 'We hiked up the hill AND we saw the view from the top.' Or 'We hiked up the hill AND saw the view from the top.' Both count.",
    ],
    watchFor: [
      "Comma instead of AND — that's a comma splice, mark as cross.",
      "Two separate sentences with a full stop — mark as cross, but note that the student split correctly so they can join tomorrow.",
    ],
    tag: "[Day 1 | Mon | Exit AND]",
  }, { requireSay: false, requireDo: false }),
});

/* =========================================================================
 *  DAY 2 — TUESDAY — BUT (contrasting two ideas) — topic: school holidays
 * ========================================================================= */

const TUE = { dayText: "Tue", dayColor: C.SECONDARY };

dayIntroSlide({
  dayText: TUE.dayText, dayColor: TUE.dayColor,
  joiningWord: "BUT",
  hookHeader: "Review — yesterday's joining word",
  hookBody:
    "Yesterday we used AND to join two ideas that GO TOGETHER. " +
    "On your whiteboard, write ONE sentence about school using AND. " +
    "Quick share, then we move on.",
  scItems: SC_BASE("BUT"),
  notes: composeNotes({
    say: [
      "Yesterday's joining word was AND. We used AND when two ideas go together.",
      "Quick recall: write one sentence about school using AND. 60 seconds.",
      "Today we move to BUT. BUT is for two ideas that contrast — they don't quite agree.",
    ],
    do: [
      "Set a 60-second timer for the AND recall.",
      "Pick one strong example to read aloud, then introduce BUT.",
    ],
    teacherNotes: [
      "Today's topic is school holidays — relatable for every student.",
      "Most likely SC2 reach today is using BUT correctly without slipping into AND.",
    ],
    enabling: [
      "FOCUS GROUP: orally generate the AND recall sentence as a pair before writing.",
      "EXTENSION: ask higher-ability students to predict what 'contrast' means before BUT is taught.",
    ],
    watchFor: [
      "Students still writing two separate sentences — remind them: ONE sentence, joined by AND.",
    ],
    tag: "[Day 2 | Tue | Review + LI/SC]",
  }, { requireSay: false, requireDo: false }),
});

teachSlide({
  dayText: TUE.dayText, dayColor: TUE.dayColor,
  joiningWord: "BUT",
  meaning: "contrasting ideas",
  twoIdeas: [
    "I wanted to sleep in.",
    "My little sister woke me up early.",
  ],
  combined: "I wanted to sleep in BUT my little sister woke me up early.",
  whyItWorks:
    "Two ideas push against each other — wanting to sleep, then being woken up. BUT signals the contrast. Different subjects ('I' / 'sister'), so we keep both — no reduction here.",
  nonExample: {
    wrong: "I wanted to sleep in, my little sister woke me up early.",
    why:
      "This is a comma splice — two full ideas joined by only a comma. " +
      "A comma alone cannot join two full ideas. Use BUT (or a full stop) to fix it.",
  },
  notes: composeNotes({
    say: [
      "BUT joins two ideas that contrast — they push against each other.",
      "Listen: I wanted to sleep in BUT my little sister woke me up early. The first idea is what I wanted; the second idea is what actually happened. They don't agree.",
      "Now the wrong example. Just a comma between two full ideas — that's called a comma splice. It's a really common mistake.",
    ],
    do: [
      "Read both ideas aloud with strong contrast in your voice.",
      "Highlight BUT in the combined sentence.",
      "Read the non-example aloud and ask 'What's the only thing joining these two ideas?' (a comma — not enough).",
    ],
    teacherNotes: [
      "Comma splice is the target error today. Naming it for students is fine but not essential — focus on the fix: BUT (or a full stop).",
      "Today's example has DIFFERENT subjects ('I' / 'my little sister'), so we keep both. Reduction (Mon's move) only works when subjects match.",
    ],
    watchFor: [
      "Students who write 'and but' or 'but and' — pick one joining word per pair.",
      "Students who try to drop a subject when the subjects are different — that breaks the meaning. Push them to keep both subjects.",
    ],
    tag: "[Day 2 | Tue | Teach BUT]",
  }, { requireSay: false, requireDo: false }),
});

guidedSlide({
  dayText: TUE.dayText, dayColor: TUE.dayColor,
  joiningWord: "BUT",
  pairs: [
    ["I planned to read every day.", "I only read on the first weekend."],
    ["The forecast said rain.", "The sun came out by lunchtime."],
  ],
  answers: [
    "I planned to read every day BUT I only read on the first weekend.",
    "The forecast said rain BUT the sun came out by lunchtime.",
  ],
  notesBuild: composeNotes({
    say: [
      "Two pairs. Both have a contrast in them — find the contrast, then join with BUT.",
      "Pair 1 first. 60 seconds. Then pair 2.",
    ],
    do: [
      "PROTOCOL: students write on whiteboards FIRST, hold up, THEN click to reveal. Do not click reveal until every student has attempted. If kids learn to wait, the practice becomes recognition not production.",
      "Read pair 1 aloud. Pause. Scan boards.",
      "Then pair 2 — same routine.",
      "Click to reveal answers.",
    ],
    teacherNotes: [
      "Pair 1: contrast is between planning and actually doing. Pair 2: contrast is between forecast and reality.",
    ],
    watchFor: [
      "Students using AND instead of BUT — ask 'do these ideas agree, or disagree?'",
      "Students adding a comma before BUT — accept it as fine for these short sentences.",
    ],
    tag: "[Day 2 | Tue | Guided BUT - build]",
  }, { requireSay: false, requireDo: false }),
  notesReveal: composeNotes({
    say: [
      "Here are the answers. The contrast is the key — that's what BUT signals.",
    ],
    do: [
      "Read both answers aloud with the contrast voice.",
      "Quick check: thumbs up if both right; sideways if one right; down if neither.",
    ],
    tag: "[Day 2 | Tue | Guided BUT - reveal]",
  }, { requireSay: false, requireDo: false }),
});

independentSlide({
  dayText: TUE.dayText, dayColor: TUE.dayColor,
  joiningWord: "BUT",
  prompt:
    "Write ONE sentence about something you wanted to do these school holidays " +
    "and what actually happened. Use BUT to join the two ideas.",
  sentenceStarter: "I wanted to ___ BUT ___.",
  notes: composeNotes({
    say: [
      "Your turn. Think of ONE thing you wanted to do these holidays, and what actually happened instead.",
      "Use BUT to join the two ideas. ONE sentence in your warm-up book.",
      "Four minutes — go.",
    ],
    do: [
      "Timer 4 minutes.",
      "Circulate. Read the starter aloud with stuck students.",
      "Note one strong example to share.",
    ],
    enabling: [
      "FOCUS GROUP: offer a personal scaffold orally — 'I wanted pizza for dinner BUT we had pasta.' — then ask them to write their own version with the same shape.",
      "EXTENSION: ask higher-ability students to extend their sentence with WHEN or WHERE detail in each half (e.g. 'I wanted to go skating on Saturday BUT it rained all afternoon.').",
    ],
    watchFor: [
      "Students writing two separate sentences with a full stop — redirect: ONE sentence with BUT in the middle.",
      "Comma splice — the same error from the non-example. Reteach BUT, not just a comma.",
    ],
    tag: "[Day 2 | Tue | Independent BUT]",
  }, { requireSay: false, requireDo: false }),
});

exitSlide({
  dayText: TUE.dayText, dayColor: TUE.dayColor,
  joiningWord: "BUT",
  taskHeader: "Today's exit check",
  taskIdeas: [
    "I wanted to stay up late.",
    "I fell asleep on the couch.",
  ],
  markCriteria: MARK_CRITERIA("BUT"),
  notes: composeNotes({
    say: [
      "Final task. Combine these two ideas into ONE sentence using BUT. Two minutes.",
    ],
    do: [
      "Set a 2-minute timer.",
      "Walk and tick on the cohort tracker using the right-hand criteria.",
      "Read out one strong sentence before students pack up.",
    ],
    teacherNotes: [
      "Expected: 'I wanted to stay up late BUT I fell asleep on the couch.' Comma before BUT is acceptable but not required.",
    ],
    watchFor: [
      "Comma splice (no BUT) — mark as cross and reteach tomorrow if it shows up across the cohort.",
      "Use of AND instead of BUT — mark as cross because the meaning changes.",
    ],
    tag: "[Day 2 | Tue | Exit BUT]",
  }, { requireSay: false, requireDo: false }),
});

/* =========================================================================
 *  DAY 3 — WEDNESDAY — SO (cause and effect) — topic: sport
 * ========================================================================= */

const WED = { dayText: "Wed", dayColor: C.ACCENT };

dayIntroSlide({
  dayText: WED.dayText, dayColor: WED.dayColor,
  joiningWord: "SO",
  hookHeader: "Review — yesterday's joining word",
  hookBody:
    "Yesterday we used BUT for two ideas that CONTRAST. " +
    "On your whiteboard, write ONE sentence about your morning using BUT. " +
    "Quick share, then we move on.",
  scItems: SC_BASE("SO"),
  notes: composeNotes({
    say: [
      "Yesterday's joining word was BUT — for two ideas that contrast.",
      "Quick recall on whiteboards: one sentence about your morning using BUT. 60 seconds.",
      "Today's joining word is SO. SO shows cause and effect — one thing makes the other happen.",
    ],
    do: [
      "60-second timer for BUT recall.",
      "Pick one strong example to share, then introduce SO.",
    ],
    teacherNotes: [
      "Today's topic is sport. SO is the cause-effect joiner — most students confuse it with BECAUSE, but the order is reversed.",
    ],
    enabling: [
      "FOCUS GROUP: rehearse the BUT recall orally with a partner before writing.",
      "EXTENSION: ask higher-ability students 'what is the difference between SO and BECAUSE?' as a quick warm-up question.",
    ],
    watchFor: [
      "Students who use BUT to mean 'and' (no real contrast) — flag for the next teach slide.",
    ],
    tag: "[Day 3 | Wed | Review + LI/SC]",
  }, { requireSay: false, requireDo: false }),
});

teachSlide({
  dayText: WED.dayText, dayColor: WED.dayColor,
  joiningWord: "SO",
  meaning: "cause and effect",
  twoIdeas: [
    "It was raining hard.",
    "We moved training inside.",
  ],
  combined: "It was raining hard SO we moved training inside.",
  whyItWorks:
    "The first idea (rain) is the CAUSE. The second idea (moving inside) is the EFFECT. SO connects cause to effect, in that order.",
  nonExample: {
    wrong: "I was tired so I went to bed so I slept all night so I missed breakfast.",
    why:
      "Run-on sentence — four SOs strung together. SO joins ONE pair (cause and effect). " +
      "Use a full stop after one SO and start a new sentence: 'I was tired SO I went to bed. I slept all night and missed breakfast.'",
  },
  notes: composeNotes({
    say: [
      "SO is for cause and effect. The first idea makes the second idea happen.",
      "It was raining hard SO we moved training inside. The rain caused us to move. Cause first, effect second.",
      "Now the wrong example — four SOs in a row. That's a run-on. SO joins one pair, not a chain.",
    ],
    do: [
      "Point to the first idea and say 'cause'. Point to the second and say 'effect'.",
      "Read the run-on aloud — students hear it become breathless.",
      "Show the fix: first SO is fine, then a full stop and a new sentence.",
    ],
    teacherNotes: [
      "Common conflation: students mix up SO and BECAUSE. SO = cause then effect; BECAUSE = effect then cause. Today, focus only on SO.",
      "Different subjects ('It' / 'we'), so we keep both — no reduction here. Reduction (Mon's move) only fires when the subjects match.",
    ],
    watchFor: [
      "Run-on sentences chaining multiple SOs — the most common error today.",
    ],
    tag: "[Day 3 | Wed | Teach SO]",
  }, { requireSay: false, requireDo: false }),
});

guidedSlide({
  dayText: WED.dayText, dayColor: WED.dayColor,
  joiningWord: "SO",
  pairs: [
    ["My boots were too tight.", "I asked Mum for new ones."],
    ["The whistle blew.", "Everyone ran onto the field."],
  ],
  answers: [
    "My boots were too tight SO I asked Mum for new ones.",
    "The whistle blew SO everyone ran onto the field.",
  ],
  notesBuild: composeNotes({
    say: [
      "Two pairs. For each pair, find the cause and the effect. Then join with SO.",
      "Pair 1, then pair 2. About 60 seconds each.",
    ],
    do: [
      "PROTOCOL: students write on whiteboards FIRST, hold up, THEN click to reveal. Do not click reveal until every student has attempted. If kids learn to wait, the practice becomes recognition not production.",
      "Read pair 1 aloud, pause, scan boards.",
      "Pair 2 — same routine.",
      "Click to reveal answers.",
    ],
    teacherNotes: [
      "Pair 1: cause = tight boots, effect = asking Mum. Pair 2: cause = whistle, effect = running. Both fit SO neatly.",
    ],
    watchFor: [
      "Students reversing cause and effect (e.g. 'I asked Mum SO my boots were too tight') — that's the order error. Reteach: cause first.",
    ],
    tag: "[Day 3 | Wed | Guided SO - build]",
  }, { requireSay: false, requireDo: false }),
  notesReveal: composeNotes({
    say: [
      "Here are the answers. Cause first, then SO, then effect.",
    ],
    do: [
      "Read both. Point at cause and effect as you read.",
      "Thumbs check.",
    ],
    tag: "[Day 3 | Wed | Guided SO - reveal]",
  }, { requireSay: false, requireDo: false }),
});

independentSlide({
  dayText: WED.dayText, dayColor: WED.dayColor,
  joiningWord: "SO",
  prompt:
    "Write ONE sentence about a sport you play or watch. " +
    "Use SO to join a cause and an effect.",
  sentenceStarter: "I had ___ SO I ___.",
  notes: composeNotes({
    say: [
      "Your turn. Think of a sport you play or watch.",
      "Write ONE sentence with a cause and an effect, joined by SO.",
      "Four minutes.",
    ],
    do: [
      "Timer 4 minutes.",
      "Circulate. For students who can't think of a sport, suggest 'PE this week'.",
      "Note one example to share.",
    ],
    enabling: [
      "FOCUS GROUP: scaffold orally — 'I was tired SO I rested at half time.' — and ask them to swap one detail (the cause OR the effect) to make it their own.",
      "EXTENSION: ask higher-ability students to write a 2-sentence pair where the first uses SO (cause -> effect) and the second uses BUT (contrast). Models how joining words combine in real writing.",
    ],
    watchFor: [
      "Run-ons (more than one SO) — redirect to ONE pair only.",
      "Effect-then-cause order — ask 'which one happened first?'",
    ],
    tag: "[Day 3 | Wed | Independent SO]",
  }, { requireSay: false, requireDo: false }),
});

exitSlide({
  dayText: WED.dayText, dayColor: WED.dayColor,
  joiningWord: "SO",
  taskHeader: "Today's exit check",
  taskIdeas: [
    "It was a hot day.",
    "We brought extra water bottles.",
  ],
  markCriteria: MARK_CRITERIA("SO"),
  notes: composeNotes({
    say: [
      "Final task. Combine these two ideas into ONE sentence using SO. Two minutes.",
    ],
    do: [
      "Timer 2 minutes.",
      "Walk and tick using the right-hand criteria.",
      "Share one strong example before pack-up.",
    ],
    teacherNotes: [
      "Expected: 'It was a hot day SO we brought extra water bottles.' Cause = heat, effect = bringing water.",
    ],
    watchFor: [
      "Reversed cause-effect order (water -> heat) — mark as cross because the meaning is wrong.",
      "Run-ons — mark as cross.",
    ],
    tag: "[Day 3 | Wed | Exit SO]",
  }, { requireSay: false, requireDo: false }),
});

/* =========================================================================
 *  DAY 4 — THURSDAY — OR (choice between two ideas) — topic: weekend / concert
 * ========================================================================= */

const THU = { dayText: "Thu", dayColor: C.ALERT };

dayIntroSlide({
  dayText: THU.dayText, dayColor: THU.dayColor,
  joiningWord: "OR",
  hookHeader: "Review — yesterday's joining word",
  hookBody:
    "Yesterday we used SO for cause and effect. " +
    "On your whiteboard, write ONE sentence about today using SO. " +
    "Quick share, then we move on.",
  scItems: SC_BASE("OR"),
  notes: composeNotes({
    say: [
      "Yesterday's joining word was SO — for cause and effect.",
      "Quick recall: one sentence about today using SO. 60 seconds.",
      "Today's joining word is OR. OR is for a CHOICE between two things.",
    ],
    do: [
      "60-second timer for SO recall.",
      "Pick one strong example, then introduce OR.",
    ],
    teacherNotes: [
      "Today's topic is weekend plans or going to a concert — both have natural choice scenarios.",
      "OR can join two nouns, two verb phrases, or two full ideas. Today we focus on joining two ideas.",
    ],
    enabling: [
      "FOCUS GROUP: scaffold the SO recall orally with a partner before writing.",
      "EXTENSION: ask higher-ability students 'what's the difference between OR and AND when there are two options?'",
    ],
    watchFor: [
      "Run-ons from yesterday's lesson — quick reminder to use a full stop after one SO.",
    ],
    tag: "[Day 4 | Thu | Review + LI/SC]",
  }, { requireSay: false, requireDo: false }),
});

teachSlide({
  dayText: THU.dayText, dayColor: THU.dayColor,
  joiningWord: "OR",
  meaning: "showing a choice",
  twoIdeas: [
    "We could line up for the concert.",
    "We could grab dinner first.",
  ],
  // Combined demonstrates REDUCTION: same subject + modal ('we could'),
  // so we drop the second 'we could' after OR. Shorter, stronger, and
  // matches the Q3-style reduction pattern.
  combined: "We could line up for the concert OR grab dinner first.",
  whyItWorks:
    "Both options share 'we could' — drop the repeat after OR. Two options, ONE shorter sentence.",
  nonExample: {
    wrong: "We could line up for the concert OR.",
    why:
      "OR must offer TWO options. There's only one option here, so OR has nothing to compare it to. " +
      "Make sure each side of OR is a full idea or a clean second option.",
  },
  notes: composeNotes({
    say: [
      "OR is for a choice. We're picking between two options.",
      "Two ideas: 'We could line up for the concert.' / 'We could grab dinner first.'",
      "Both ideas start with 'we could' — same trick as Monday with AND. I drop the second 'we could' after OR. 'We could line up for the concert OR grab dinner first.'",
      "Now the wrong example. OR has nothing on the other side.",
    ],
    do: [
      "Read with a 'pick one' tone — slight upward inflection on the first idea.",
      "Highlight OR in the combined sentence and physically cross out the second 'we could' on the board so the reduction move is visible.",
      "Read the non-example aloud and ask 'what's the second option?' (there isn't one).",
    ],
    teacherNotes: [
      "OR works on the same structural principle as AND — both sides need a full idea. The MEANING is what's different (choice vs. addition).",
      "Reduction (Mon's move) fires again here: same subject ('we could') means we can drop the repeat after OR.",
    ],
    watchFor: [
      "Students who use OR for two things that aren't actually options ('I like soccer OR I like cricket' is fine; 'I like soccer OR I have a sister' is not).",
      "Students who keep both 'we could's — accept it (still grammatical), but model the reduced version once more.",
    ],
    tag: "[Day 4 | Thu | Teach OR - join + reduce]",
  }, { requireSay: false, requireDo: false }),
});

guidedSlide({
  dayText: THU.dayText, dayColor: THU.dayColor,
  joiningWord: "OR",
  pairs: [
    ["We could go to the beach.", "We could stay home and watch movies."],
    ["I could save my money for the concert.", "I could spend it on snacks."],
  ],
  answers: [
    "We could go to the beach OR we could stay home and watch movies.",
    "I could save my money for the concert OR I could spend it on snacks.",
  ],
  notesBuild: composeNotes({
    say: [
      "Two pairs. Each pair gives two options — join them with OR.",
      "Pair 1, then pair 2. 60 seconds each.",
    ],
    do: [
      "PROTOCOL: students write on whiteboards FIRST, hold up, THEN click to reveal. Do not click reveal until every student has attempted. If kids learn to wait, the practice becomes recognition not production.",
      "Read pair 1, pause, scan boards.",
      "Pair 2 — same routine.",
      "Click to reveal.",
    ],
    teacherNotes: [
      "Pair 1: choice between activity types. Pair 2: choice about money.",
    ],
    watchFor: [
      "Students who try to drop 'we could' or 'I could' — that's fine for fluency, just ensure both sides are still readable as options.",
    ],
    tag: "[Day 4 | Thu | Guided OR - build]",
  }, { requireSay: false, requireDo: false }),
  notesReveal: composeNotes({
    say: [
      "Here are the answers. Two real options on each side of OR.",
    ],
    do: [
      "Read both. Highlight OR.",
      "Thumbs check.",
    ],
    tag: "[Day 4 | Thu | Guided OR - reveal]",
  }, { requireSay: false, requireDo: false }),
});

independentSlide({
  dayText: THU.dayText, dayColor: THU.dayColor,
  joiningWord: "OR",
  prompt:
    "Write ONE sentence about your weekend plans or going to a concert. " +
    "Use OR to join two options.",
  sentenceStarter: "I might ___ OR ___.",
  notes: composeNotes({
    say: [
      "Your turn. Think about your weekend plans or about going to a concert.",
      "Write ONE sentence with TWO options, joined by OR.",
      "Four minutes.",
    ],
    do: [
      "Timer 4 minutes.",
      "Circulate. For stuck students, suggest 'this Saturday' as the topic.",
      "Note one example to share.",
    ],
    enabling: [
      "FOCUS GROUP: scaffold orally — 'I might watch a movie OR play with my dog.' — and ask them to swap one option to make it their own.",
      "EXTENSION: ask higher-ability students to write a sentence with THREE options joined by OR (using commas — e.g. 'I might swim, ride my bike, OR read.'). Introduces the comma-then-OR list pattern.",
    ],
    watchFor: [
      "Students using OR for non-options ('I might swim OR my brother can ride bikes') — meaning is broken.",
      "Two separate sentences — redirect: ONE sentence with OR in the middle.",
    ],
    tag: "[Day 4 | Thu | Independent OR]",
  }, { requireSay: false, requireDo: false }),
});

exitSlide({
  dayText: THU.dayText, dayColor: THU.dayColor,
  joiningWord: "OR",
  taskHeader: "Today's exit check",
  taskIdeas: [
    "We could walk to the park.",
    "We could ride our bikes there.",
  ],
  markCriteria: MARK_CRITERIA("OR"),
  notes: composeNotes({
    say: [
      "Last task. Combine these two ideas into ONE sentence using OR. Two minutes.",
    ],
    do: [
      "Timer 2 minutes.",
      "Walk and tick using the right-hand criteria.",
      "Share one strong example before pack-up.",
    ],
    teacherNotes: [
      "Expected: 'We could walk to the park OR we could ride our bikes there.' Both options are about getting to the park — that's a clean choice.",
    ],
    watchFor: [
      "Use of AND instead of OR — meaning shifts from 'choose' to 'do both'. Mark as cross.",
      "Comma splice — mark as cross.",
    ],
    tag: "[Day 4 | Thu | Exit OR]",
  }, { requireSay: false, requireDo: false }),
});

/* =========================================================================
 *  DAY 5 — FRIDAY — MIXED PRACTICE — choose AND/BUT/SO/OR
 * ========================================================================= */

const FRI = { dayText: "Fri", dayColor: C.SUCCESS };

dayIntroSlide({
  dayText: FRI.dayText, dayColor: FRI.dayColor,
  joiningWord: "AND / BUT / SO / OR",
  hookHeader: "Review — this week's four joining words",
  hookBody:
    "AND - ideas go TOGETHER.\n" +
    "BUT - ideas CONTRAST.\n" +
    "SO - one idea CAUSES the other.\n" +
    "OR - a CHOICE between ideas.\n" +
    "Today: you choose the right one.",
  scItems: [
    "I can spot the meaning between two ideas.",
    "I can choose the right joining word (AND / BUT / SO / OR).",
    "I can write ONE grammatical sentence using my choice.",
  ],
  notes: composeNotes({
    say: [
      "Friday is mixed practice. We've done AND, BUT, SO, OR.",
      "Today YOU choose. We won't tell you which joining word — you have to read the two ideas and pick.",
      "Here's the meaning of each one — read it through with me.",
    ],
    do: [
      "Read the four meanings aloud, pointing to each.",
      "Frame the difficulty: today is harder because you don't get the joining word handed to you.",
    ],
    teacherNotes: [
      "Friday is the cumulative retrieval slide. The harder transfer is choosing, not joining.",
      "If the cohort tracker shows weakness in one joining word from Mon-Thu, flag that pattern in your notes for Week 5.",
    ],
    enabling: [
      "FOCUS GROUP: pre-print or write on the side board the four joining words with their meanings so students can refer to them all session.",
      "EXTENSION: ask higher-ability students to write a personal one-sentence summary of when each joining word is the best fit.",
    ],
    watchFor: [
      "Students who default to AND for everything — push them to read for meaning first.",
    ],
    tag: "[Day 5 | Fri | Review all four + LI/SC]",
  }, { requireSay: false, requireDo: false }),
});

/* Friday's "teach" slide is a decision-rule reminder rather than a new word.
 * Use a custom 4-card layout instead of the standard teachSlide. */

function fridayDecisionSlide() {
  const s = pres.addSlide();
  addTopBar(s, C.PRIMARY);
  badgeStack(s, "Decision rule", C.PRIMARY, FRI.dayText, FRI.dayColor);
  addTitle(s, "How to choose the right joining word");

  // 2 x 2 grid of mini-cards, one per joining word
  const cardW = 4.4;
  const cardH = 1.65;
  const gapX = 0.20;
  const gapY = 0.16;
  const startX = 0.5;
  const startY = CONTENT_TOP;

  const cards = [
    { word: "AND", meaning: "ideas go TOGETHER", example: '"I packed my torch AND my raincoat."', strip: C.PRIMARY },
    { word: "BUT", meaning: "ideas CONTRAST",     example: '"It was sunny BUT cold."',          strip: C.SECONDARY },
    { word: "SO",  meaning: "one idea CAUSES the other", example: '"It rained SO we stayed inside."', strip: C.ACCENT },
    { word: "OR",  meaning: "a CHOICE between ideas", example: '"We could read OR play a game."', strip: C.ALERT },
  ];

  cards.forEach((c, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = startX + col * (cardW + gapX);
    const y = startY + row * (cardH + gapY);
    addCard(s, x, y, cardW, cardH, { strip: c.strip, fill: C.WHITE });

    // Big joining word
    s.addText(c.word, {
      x: x + 0.18, y: y + 0.10, w: 1.5, h: 0.70,
      fontSize: 32, fontFace: FONT_H, color: c.strip, bold: true, margin: 0,
    });
    // Meaning
    s.addText(c.meaning, {
      x: x + 1.75, y: y + 0.18, w: cardW - 1.85, h: 0.62,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });
    // Example
    s.addText(c.example, {
      x: x + 0.18, y: y + 0.85, w: cardW - 0.36, h: cardH - 0.95,
      fontSize: 12, fontFace: FONT_H, color: C.MUTED, italic: true, valign: "top", margin: 0,
      fit: "shrink", shrinkText: true,
    });
  });

  addFooter(s, FOOTER);
  s.addNotes(composeNotes({
    say: [
      "Quick recap of the four. Read each card with me.",
      "Step 1: read the two ideas and ask 'what's the meaning between them?'",
      "Step 2: pick the joining word that matches that meaning.",
    ],
    do: [
      "Walk the four cards with the class. 60-90 seconds total.",
      "Don't drag this — it's a refresher, not a re-teach.",
    ],
    teacherNotes: [
      "Keep this slide tight (under 2 minutes). The harder thinking happens on the next slide.",
      "Reminder: when both ideas share the same subject (Mon and Thu examples — 'we', 'I'), drop the repeat after the joining word. That reduction move is what gets students to the Q3 standard.",
    ],
    watchFor: [
      "Students who can't recall what one of the joining words means — flag that one for extra attention this week.",
    ],
    tag: "[Day 5 | Fri | Decision rule]",
  }, { requireSay: false, requireDo: false }));
  runSlideDiagnostics(s, pres, { respectSafeBottom: false });
  return s;
}

fridayDecisionSlide();

/* Friday's guided practice — students must CHOOSE the joining word.
 * Reuse guidedSlide but the "joiningWord" highlight will be different in
 * each answer, so we pass a generic placeholder and rely on the answers'
 * literal text to show the chosen word. */

function fridayGuidedSlide() {
  return withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.SECONDARY);
      badgeStack(s, "Practise together", C.SECONDARY, FRI.dayText, FRI.dayColor);
      addTitle(s, "Choose the right joining word");

      // Mini whiteboard cue
      addTextOnShape(s, "Mini-whiteboards", {
        x: 7.4, y: 0.20, w: 2.1, h: 0.36, rectRadius: 0.06,
        fill: { color: C.WHITE },
        line: { color: C.SECONDARY, width: 1.0 },
      }, {
        fontSize: 11, fontFace: FONT_B, color: C.SECONDARY, bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      const pairs = [
        ["The bell rang.", "We packed up our books."],          // SO
        ["I love pancakes.", "I don't like syrup on them."],    // BUT
      ];
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
        s.addText(`Pair ${i + 1} — pick AND, BUT, SO, or OR`, {
          x: 0.75, y: y + 0.08, w: 8, h: 0.26,
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
      s.addText("Try first on whiteboards. Click to reveal which joining word fits.", {
        x: 0.75, y: stubY, w: 8.5, h: stubH,
        fontSize: 13, fontFace: FONT_B, color: C.MUTED, italic: true,
        align: "center", valign: "middle", margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(composeNotes({
        say: [
          "Two pairs. For each one, decide which joining word fits — AND, BUT, SO, or OR — and write the combined sentence.",
          "Pair 1 first. Pair 2 second. About 60 seconds each.",
        ],
        do: [
          "Read pair 1 aloud. Pause for 60 seconds. Scan boards.",
          "Pair 2 — same routine.",
          "Click to reveal both answers and the joining word used.",
        ],
        teacherNotes: [
          "Pair 1 is cause-effect (SO). Pair 2 is contrast (BUT). Both are clear cases — students should not have to wrestle with ambiguity.",
        ],
        watchFor: [
          "Students using AND for both pairs — push them back to read the meaning first.",
        ],
        tag: "[Day 5 | Fri | Guided mixed - build]",
      }, { requireSay: false, requireDo: false }));
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

      const answers = [
        { text: 'The bell rang SO we packed up our books.', word: "SO" },
        { text: "I love pancakes BUT I don't like syrup on them.", word: "BUT" },
      ];
      const ansListY = stubY + 0.30;
      const perAnsH = (stubH - 0.36) / 2;
      answers.forEach((ans, i) => {
        const y = ansListY + i * perAnsH;
        const runs = highlightSentence(ans.text, ans.word, 14, C.WHITE, C.BG_CARD);
        s.addText(runs.map((run, j) => ({
          text: `${j === 0 ? `${i + 1}.  "` : ""}${run.text}${j === runs.length - 1 ? '"' : ""}`,
          options: { ...run.options, fontFace: FONT_H, bold: true, breakLine: false },
        })), {
          x: 0.75, y, w: 8.5, h: perAnsH - 0.05,
          valign: "middle", margin: 0,
          fit: "shrink", shrinkText: true,
        });
      });

      s.addNotes(composeNotes({
        say: [
          "Pair 1: cause and effect — the bell caused us to pack up. SO is the right joining word.",
          "Pair 2: I love them, but there's a contrast — I don't like syrup. BUT is the right joining word.",
        ],
        do: [
          "Read both answers. Name the meaning for each (cause-effect / contrast).",
          "Thumbs check before independent.",
        ],
        tag: "[Day 5 | Fri | Guided mixed - reveal]",
      }, { requireSay: false, requireDo: false }));
    }
  );
}

fridayGuidedSlide();

/* Friday independent — student picks the joining word themselves */

function fridayIndependentSlide() {
  const s = pres.addSlide();
  addTopBar(s, C.ACCENT);
  badgeStack(s, "Your turn", C.ACCENT, FRI.dayText, FRI.dayColor);
  addTitle(s, "Write ONE sentence — your choice");

  // Hero prompt card
  const promptY = CONTENT_TOP;
  const promptH = 2.30;
  addCard(s, 0.5, promptY, 9, promptH, { strip: C.ACCENT, fill: C.WHITE });
  s.addText("Topic prompt", {
    x: 0.75, y: promptY + 0.12, w: 5, h: 0.28,
    fontSize: 12, fontFace: FONT_B, color: C.ACCENT, bold: true, margin: 0,
  });
  s.addText(
    "Write ONE sentence about something that happened this week (at school, at home, or at sport). " +
    "Choose the best joining word: AND, BUT, SO, or OR.",
    {
      x: 0.75, y: promptY + 0.46, w: 8.5, h: promptH - 0.60,
      fontSize: 20, fontFace: FONT_H, color: C.CHARCOAL, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    }
  );

  // Reference strip — shows the four options
  const refY = promptY + promptH + 0.14;
  const refH = 0.62;
  s.addShape("roundRect", {
    x: 0.5, y: refY, w: 9, h: refH, rectRadius: 0.08,
    fill: { color: C.BG_CARD },
    line: { color: C.SECONDARY, width: 0.8 },
  });
  s.addText("Choose: ", {
    x: 0.75, y: refY, w: 1.1, h: refH,
    fontSize: 12, fontFace: FONT_B, color: C.SECONDARY, bold: true,
    valign: "middle", margin: 0,
  });
  // Option chips
  const chips = [
    { word: "AND", color: C.PRIMARY },
    { word: "BUT", color: C.SECONDARY },
    { word: "SO",  color: C.ACCENT },
    { word: "OR",  color: C.ALERT },
  ];
  const chipW = 1.5;
  const chipGap = 0.15;
  const chipsStart = 1.95;
  chips.forEach((c, i) => {
    const cx = chipsStart + i * (chipW + chipGap);
    addTextOnShape(s, c.word, {
      x: cx, y: refY + 0.08, w: chipW, h: refH - 0.16, rectRadius: 0.06,
      fill: { color: c.color },
    }, {
      fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
  });

  // Reminder strip
  const remindY = refY + refH + 0.12;
  const remindH = SAFE_BOTTOM - remindY;
  s.addShape("roundRect", {
    x: 0.5, y: remindY, w: 9, h: remindH, rectRadius: 0.08,
    fill: { color: C.PRIMARY },
  });
  s.addText("Remember: ONE sentence, with TWO ideas joined by your chosen word.", {
    x: 0.75, y: remindY, w: 8.5, h: remindH,
    fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true,
    align: "center", valign: "middle", margin: 0,
  });

  addFooter(s, FOOTER);
  s.addNotes(composeNotes({
    say: [
      "Your turn. Think of one thing that happened this week.",
      "Two short ideas, then choose your joining word — AND, BUT, SO, or OR.",
      "Four minutes.",
    ],
    do: [
      "Timer 4 minutes.",
      "Circulate. For stuck students, prompt 'what happened today, and what happened next?'",
      "Note one example per joining word if you can — useful for share.",
    ],
    enabling: [
      "FOCUS GROUP: ask which joining word they want BEFORE they write, so they only have to combine, not choose.",
      "EXTENSION: challenge higher-ability students to write a sentence where they REDUCE — both ideas share the same subject, and they drop the repeat after the joining word (e.g. 'I rode my bike to school AND played soccer at lunch.'). This is the Q3-task move and a stretch this week.",
    ],
    watchFor: [
      "Default to AND when a different word fits better — push them to defend the choice.",
      "Choosing a word that doesn't match the meaning (e.g. SO when the ideas just go together) — mark and reteach.",
      "Students repeating the subject when they could reduce — accept it but call out the stronger reduced version when sharing.",
    ],
    tag: "[Day 5 | Fri | Independent mixed + reduce]",
  }, { requireSay: false, requireDo: false }));
  runSlideDiagnostics(s, pres, { respectSafeBottom: false });
  return s;
}

fridayIndependentSlide();

/* Friday exit check — same shape as days 1-4 but no joining word given. */

function fridayExitSlide() {
  const s = pres.addSlide();
  addTopBar(s, C.ALERT);
  badgeStack(s, "Exit check", C.ALERT, FRI.dayText, FRI.dayColor);

  addTextOnShape(s, "✓  EXIT", {
    x: 8.0, y: 0.20, w: 1.5, h: 0.36, rectRadius: 0.06,
    fill: { color: C.WHITE },
    line: { color: C.ALERT, width: 1.5 },
  }, {
    fontSize: 11, fontFace: FONT_B, color: C.ALERT, bold: true,
    align: "center", valign: "middle", margin: 0,
  });

  addTitle(s, "Show what you know", { w: 7.4 });

  const cardH = SAFE_BOTTOM - CONTENT_TOP;
  const leftW = 5.6;
  const rightX = 6.3;
  const rightW = 3.2;

  addCard(s, 0.5, CONTENT_TOP, leftW, cardH, { strip: C.ALERT, fill: C.WHITE });
  s.addText("Today's exit check", {
    x: 0.75, y: CONTENT_TOP + 0.14, w: leftW - 0.4, h: 0.32,
    fontSize: 12, fontFace: FONT_B, color: C.ALERT, bold: true, margin: 0,
  });

  const ideaY = CONTENT_TOP + 0.50;
  s.addText('"My alarm did not go off."', {
    x: 0.75, y: ideaY, w: leftW - 0.4, h: 0.42,
    fontSize: 18, fontFace: FONT_H, color: C.CHARCOAL, italic: true, margin: 0,
    fit: "shrink", shrinkText: true,
  });
  s.addText('"I was late to school."', {
    x: 0.75, y: ideaY + 0.46, w: leftW - 0.4, h: 0.42,
    fontSize: 18, fontFace: FONT_H, color: C.CHARCOAL, italic: true, margin: 0,
    fit: "shrink", shrinkText: true,
  });

  s.addText("Combine into ONE sentence. Choose the best joining word.", {
    x: 0.75, y: ideaY + 1.05, w: leftW - 0.4, h: 0.36,
    fontSize: 14, fontFace: FONT_B, color: C.SECONDARY, bold: true, margin: 0,
  });

  s.addText("Write your sentence in your warm-up book.", {
    x: 0.75, y: CONTENT_TOP + cardH - 0.40, w: leftW - 0.4, h: 0.28,
    fontSize: 11, fontFace: FONT_B, color: C.MUTED, italic: true, margin: 0,
  });

  // Right: marking criteria for Friday
  addCard(s, rightX, CONTENT_TOP, rightW, cardH, { strip: C.SUCCESS, fill: C.BG_CARD });
  s.addText("Tick if students…", {
    x: rightX + 0.18, y: CONTENT_TOP + 0.14, w: rightW - 0.36, h: 0.32,
    fontSize: 12, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
  });

  const friMark = [
    "Two ideas combined into ONE sentence",
    "Chose the joining word that matches the meaning",
    "Sentence is grammatical (no run-on, no comma splice)",
  ];

  const rowStart = CONTENT_TOP + 0.50;
  const rowH = 0.66;
  friMark.forEach((crit, i) => {
    const y = rowStart + i * (rowH + 0.08);
    s.addShape("roundRect", {
      x: rightX + 0.18, y: y + 0.10, w: 0.32, h: 0.32, rectRadius: 0.04,
      fill: { color: C.WHITE },
      line: { color: C.SUCCESS, width: 1.2 },
    });
    s.addText(crit, {
      x: rightX + 0.58, y, w: rightW - 0.76, h: rowH,
      fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });
  });

  s.addText("Record on the cohort tracker (note the joining word chosen).", {
    x: rightX + 0.18, y: CONTENT_TOP + cardH - 0.42, w: rightW - 0.36, h: 0.30,
    fontSize: 10.5, fontFace: FONT_B, color: C.MUTED, italic: true, margin: 0,
    fit: "shrink",
  });

  addFooter(s, FOOTER);
  s.addNotes(composeNotes({
    say: [
      "Final exit. Combine these two ideas into ONE sentence — and YOU pick the joining word.",
      "Two minutes.",
    ],
    do: [
      "Timer 2 minutes.",
      "Tick on the cohort tracker. Note WHICH joining word each student chose if you have time.",
      "Quick share: read one strong example before pack-up.",
    ],
    teacherNotes: [
      "SCORING RULE — all 5 teachers score the same way to keep cohort tracker data valid.",
      "Tick if the sentence is ONE grammatical sentence with a defensible joining word. SO is the strongest match (cause-effect). AND is also acceptable (sequential).",
      "BUT and OR are not defensible here. Cross the meaning criterion for those.",
      "Tick on either SO or AND. The point of this exit is grammatical control and meaning-match, not picking the 'right' answer.",
      "Tick the FIRST criterion (two ideas combined into ONE) for any grammatical combination.",
      "Tick the SECOND criterion (meaning match) for SO or AND.",
      "Tick the THIRD criterion (grammatical) for any run-on-free, comma-splice-free sentence.",
    ],
    watchFor: [
      "Run-ons or comma splices — cross the grammatical criterion.",
      "Use of BUT or OR — cross the meaning criterion (no contrast, no choice present).",
      "Use of AND — tick on all three criteria (sequential is defensible). Note in the cohort tracker how many students chose SO vs AND so the team can decide whether Week 5 needs more meaning-match work.",
    ],
    tag: "[Day 5 | Fri | Exit mixed]",
  }, { requireSay: false, requireDo: false }));
  runSlideDiagnostics(s, pres, { respectSafeBottom: false });
  return s;
}

fridayExitSlide();

/* ─────────────────────────────────────────────────────────────────────────
 *  Write file
 * ───────────────────────────────────────────────────────────────────────── */

const outFile = path.join(OUT_DIR, "Sentence_Combining_Year56_T2W4.pptx");
pres.writeFile({ fileName: outFile }).then(() => {
  console.log("PPTX written to " + outFile);
});
