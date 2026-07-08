"use strict";

// Sentence Combining — Year 5/6 Daily Warm-Up (Wed, Thu, Fri)
// 3 days x ~6 teaching slides per day in a single PPTX.
// 15-minute warm-ups: Hook -> Teach -> Practise Together -> Your Turn -> Exit Check.
// No printed resources — all production in student workbooks.

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");
const { createTheme } = require("../themes/factory");

const T = createTheme("literacy", "grade56", 0);
const {
  C, FONT_H, FONT_B, S,
  SLIDE_W, SLIDE_H, SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, contentSlide, closingSlide,
  addTopBar, addBadge, addTitle, addCard, addFooter, addTextOnShape,
  addInstructionCard, withReveal, addRevealAnswerBar,
  runSlideDiagnostics, composeNotes,
  makeShadow,
} = T;

const OUT_DIR = path.join("output", "SC_Year56_Sentence_Combining");
const FOOTER = "Sentence Combining | Year 5/6 | 15-min warm-up";

// ──────────────────────────────────────────────────────────────────────────────
//  Shared helpers
// ──────────────────────────────────────────────────────────────────────────────

function strikeRuns(parts, fontSize) {
  const sz = fontSize || 17;
  return parts.map((p) => ({
    text: String(p.text),
    options: {
      fontSize: sz,
      fontFace: FONT_B,
      color: p.strike ? "B85450" : C.CHARCOAL,
      strike: Boolean(p.strike),
    },
  }));
}

function sentenceListText(sentences, fontSize) {
  const sz = fontSize || S.body;
  const runs = [];
  sentences.forEach((sent, i) => {
    runs.push({
      text: String(sent),
      options: { fontSize: sz, fontFace: FONT_B, color: C.CHARCOAL, breakLine: i < sentences.length - 1 },
    });
  });
  return runs;
}

function dayDivider(pres, dayLabel, dayNumber, sc) {
  const s = pres.addSlide();
  s.background = { color: C.BG_DARK };
  s.addShape("rect", { x: 0, y: 0, w: 0.10, h: SLIDE_H, fill: { color: C.ACCENT } });

  s.addText(dayLabel.toUpperCase(), {
    x: 0.7, y: 1.0, w: 8, h: 0.42,
    fontSize: 18, fontFace: FONT_B, color: C.ACCENT, bold: true, margin: 0,
  });
  s.addText(`Day ${dayNumber} of 3`, {
    x: 0.7, y: 1.48, w: 8.5, h: 0.9,
    fontSize: 38, fontFace: FONT_H, color: C.WHITE, bold: true, margin: 0,
  });

  addCard(s, 0.7, 2.7, 8.6, 1.55, { strip: C.ACCENT, fill: C.WHITE });
  s.addText("Today's goal", {
    x: 1.0, y: 2.82, w: 7.8, h: 0.28,
    fontSize: 12, fontFace: FONT_B, color: C.ACCENT, bold: true, margin: 0,
  });
  s.addText(String(sc), {
    x: 1.0, y: 3.15, w: 8.0, h: 0.95,
    fontSize: 20, fontFace: FONT_H, color: C.CHARCOAL, margin: 0, valign: "top",
    fit: "shrink", shrinkText: true,
  });

  s.addText("15-minute warm-up  |  Hook -> Teach -> Practise -> Your Turn -> Exit Check", {
    x: 0.7, y: 4.7, w: 8.5, h: 0.28,
    fontSize: 11, fontFace: FONT_B, color: C.ACCENT, margin: 0,
  });
  return s;
}

function hookSlide(pres, sentences, notes) {
  const s = pres.addSlide();
  addTopBar(s, C.PRIMARY);
  addBadge(s, "HOOK", { color: C.PRIMARY });
  addTitle(s, "Copy these sentences");

  const lineH = 0.34;
  const cardH = 0.38 + sentences.length * lineH + 0.15;
  addCard(s, 0.5, CONTENT_TOP, 9.0, cardH, { strip: C.PRIMARY, fill: C.WHITE });

  sentences.forEach((sent, i) => {
    s.addText(String(sent), {
      x: 0.80, y: CONTENT_TOP + 0.22 + i * lineH, w: 8.4, h: lineH,
      fontSize: S.body, fontFace: FONT_B, color: C.CHARCOAL, margin: 0, valign: "middle",
    });
  });

  const instrY = CONTENT_TOP + cardH + 0.2;
  addTextOnShape(s, "Copy into your workbook. We will come back to them.", {
    x: 0.5, y: instrY, w: 9.0, h: 0.48, rectRadius: 0.08,
    fill: { color: C.BG_LIGHT },
  }, {
    fontSize: 16, fontFace: FONT_B, color: C.MUTED, bold: true,
    align: "center", valign: "middle", margin: 0,
  });

  addFooter(s, FOOTER);
  if (notes) s.addNotes(notes);
  runSlideDiagnostics(s, pres, { respectSafeBottom: false });
  return s;
}

function practiceSlides(pres, set1, set2, model1, model2, qNotes, aNotes) {
  const created = [];
  const maxLines = Math.max(set1.length, set2.length);
  const sentFontSize = maxLines >= 5 ? 15 : 17;
  const lineH = maxLines >= 5 ? 0.28 : 0.32;
  const cardH = 0.42 + maxLines * lineH + 0.10;

  const buildFn = () => {
    const s = pres.addSlide();
    created.push(s);
    addTopBar(s, C.SECONDARY);
    addBadge(s, "PRACTISE TOGETHER", { color: C.SECONDARY, w: 2.6 });
    addTitle(s, "Combine each set into one sentence");

    addCard(s, 0.5, CONTENT_TOP, 4.5, cardH, { strip: C.SECONDARY, fill: C.WHITE });
    s.addText("Set 1", {
      x: 0.75, y: CONTENT_TOP + 0.10, w: 3.5, h: 0.24,
      fontSize: 11, fontFace: FONT_B, color: C.SECONDARY, bold: true, margin: 0,
    });
    s.addText(set1.map((sent, i) => ({
      text: String(sent),
      options: { fontSize: sentFontSize, fontFace: FONT_B, color: C.CHARCOAL, breakLine: i < set1.length - 1 },
    })), {
      x: 0.75, y: CONTENT_TOP + 0.38, w: 4.0, h: cardH - 0.48,
      margin: 0, valign: "top", paraSpaceAfter: 2,
      fit: "shrink", shrinkText: true,
    });

    addCard(s, 5.2, CONTENT_TOP, 4.3, cardH, { strip: C.ACCENT, fill: C.WHITE });
    s.addText("Set 2", {
      x: 5.45, y: CONTENT_TOP + 0.10, w: 3.5, h: 0.24,
      fontSize: 11, fontFace: FONT_B, color: C.ACCENT, bold: true, margin: 0,
    });
    s.addText(set2.map((sent, i) => ({
      text: String(sent),
      options: { fontSize: sentFontSize, fontFace: FONT_B, color: C.CHARCOAL, breakLine: i < set2.length - 1 },
    })), {
      x: 5.45, y: CONTENT_TOP + 0.38, w: 3.8, h: cardH - 0.48,
      margin: 0, valign: "top", paraSpaceAfter: 2,
      fit: "shrink", shrinkText: true,
    });

    const instrY = CONTENT_TOP + cardH + 0.15;
    addTextOnShape(s, "Write your combined sentence in your workbook first.", {
      x: 0.5, y: instrY, w: 9.0, h: 0.42, rectRadius: 0.08,
      fill: { color: C.BG_LIGHT },
    }, {
      fontSize: 15, fontFace: FONT_B, color: C.MUTED, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    return s;
  };

  const answerSlide = withReveal(buildFn, (s) => {
    const barY1 = CONTENT_TOP + cardH + 0.65;
    const barH = 0.72;
    addCard(s, 0.5, barY1, 4.5, barH, { strip: C.SUCCESS, fill: C.WHITE });
    s.addText("Set 1", {
      x: 0.75, y: barY1 + 0.06, w: 1.0, h: 0.22,
      fontSize: 10, fontFace: FONT_B, color: C.SUCCESS, bold: true, margin: 0,
    });
    s.addText(String(model1), {
      x: 0.75, y: barY1 + 0.26, w: 4.0, h: barH - 0.32,
      fontSize: 14, fontFace: FONT_H, color: C.CHARCOAL, margin: 0, valign: "top",
      fit: "shrink", shrinkText: true,
    });

    addCard(s, 5.2, barY1, 4.3, barH, { strip: C.SUCCESS, fill: C.WHITE });
    s.addText("Set 2", {
      x: 5.45, y: barY1 + 0.06, w: 1.0, h: 0.22,
      fontSize: 10, fontFace: FONT_B, color: C.SUCCESS, bold: true, margin: 0,
    });
    s.addText(String(model2), {
      x: 5.45, y: barY1 + 0.26, w: 3.8, h: barH - 0.32,
      fontSize: 14, fontFace: FONT_H, color: C.CHARCOAL, margin: 0, valign: "top",
      fit: "shrink", shrinkText: true,
    });
  });
  if (qNotes) created[0].addNotes(qNotes);
  if (aNotes) answerSlide.addNotes(aNotes);
}

function yourTurnSlides(pres, sentences, scaffold, qNotes, sNotes) {
  const created = [];
  const lineH = 0.34;
  const cardH = 0.32 + sentences.length * lineH + 0.12;

  const buildFn = () => {
    const s = pres.addSlide();
    created.push(s);
    addTopBar(s, C.ACCENT);
    addBadge(s, "YOUR TURN", { color: C.ACCENT });
    addTitle(s, "On your own");

    addCard(s, 0.5, CONTENT_TOP, 9.0, cardH, { strip: C.ACCENT, fill: C.WHITE });

    sentences.forEach((sent, i) => {
      s.addText(String(sent), {
        x: 0.80, y: CONTENT_TOP + 0.22 + i * lineH, w: 8.4, h: lineH,
        fontSize: S.body, fontFace: FONT_B, color: C.CHARCOAL, margin: 0, valign: "middle",
      });
    });

    const instrY = CONTENT_TOP + cardH + 0.15;
    addTextOnShape(s, "Combine these into one sentence in your workbook.", {
      x: 0.5, y: instrY, w: 9.0, h: 0.42, rectRadius: 0.08,
      fill: { color: C.BG_LIGHT },
    }, {
      fontSize: 15, fontFace: FONT_B, color: C.MUTED, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    return s;
  };

  const scaffoldSlide = withReveal(buildFn, (s) => {
    const scafY = CONTENT_TOP + cardH + 0.65;
    addCard(s, 1.0, scafY, 8.0, 0.75, { strip: C.MUTED, fill: C.BG_LIGHT });
    s.addText("Scaffold (show only if needed)", {
      x: 1.25, y: scafY + 0.06, w: 7.5, h: 0.22,
      fontSize: 10, fontFace: FONT_B, color: C.MUTED, bold: true, margin: 0,
    });
    s.addText(String(scaffold), {
      x: 1.25, y: scafY + 0.28, w: 7.5, h: 0.40,
      fontSize: 16, fontFace: FONT_H, color: C.CHARCOAL, italic: true, margin: 0,
      valign: "middle", fit: "shrink", shrinkText: true,
    });
  });
  if (qNotes) created[0].addNotes(qNotes);
  if (sNotes) scaffoldSlide.addNotes(sNotes);
}

function exitCheckSlide(pres, hookSentences, notes) {
  const s = pres.addSlide();
  addTopBar(s, C.ALERT);
  addBadge(s, "EXIT CHECK", { color: C.ALERT });
  addTitle(s, "Combine them now");

  const lineH = 0.34;
  const cardH = 0.38 + hookSentences.length * lineH + 0.15;
  addCard(s, 0.5, CONTENT_TOP, 9.0, cardH, { strip: C.ALERT, fill: C.WHITE });

  hookSentences.forEach((sent, i) => {
    s.addText(String(sent), {
      x: 0.80, y: CONTENT_TOP + 0.22 + i * lineH, w: 8.4, h: lineH,
      fontSize: S.body, fontFace: FONT_B, color: C.CHARCOAL, margin: 0, valign: "middle",
    });
  });

  const instrY = CONTENT_TOP + cardH + 0.2;
  addTextOnShape(s, "Combine these into one sentence in your workbook.", {
    x: 0.5, y: instrY, w: 9.0, h: 0.48, rectRadius: 0.08,
    fill: { color: C.BG_LIGHT },
  }, {
    fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL, bold: true,
    align: "center", valign: "middle", margin: 0,
  });

  addFooter(s, FOOTER);
  if (notes) s.addNotes(notes);
  runSlideDiagnostics(s, pres, { respectSafeBottom: false });
  return s;
}

// ──────────────────────────────────────────────────────────────────────────────
//  Teach slide builder — custom per day
// ──────────────────────────────────────────────────────────────────────────────

function teachSlideCustom(pres, strikeSentences, combinedText, watchOutText, watchOutReason, notes, extraRight) {
  const s = pres.addSlide();
  addTopBar(s, C.ACCENT);
  addBadge(s, "TEACH", { color: C.ACCENT });
  addTitle(s, "Combine into one sentence");

  const perLine = strikeSentences.length <= 3 ? 0.32 : 0.28;
  const leftH = 0.40 + strikeSentences.length * perLine + 0.18;
  addCard(s, 0.5, CONTENT_TOP, 4.5, leftH, { strip: C.PRIMARY, fill: C.WHITE });
  s.addText("Before", {
    x: 0.75, y: CONTENT_TOP + 0.08, w: 3.5, h: 0.26,
    fontSize: 12, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
  });
  const allRuns = [];
  strikeSentences.forEach((runs, i) => {
    runs.forEach((run, j) => {
      const isLastRun = j === runs.length - 1;
      allRuns.push({
        text: run.text,
        options: {
          ...run.options,
          breakLine: isLastRun && i < strikeSentences.length - 1,
        },
      });
    });
  });
  s.addText(allRuns, {
    x: 0.75, y: CONTENT_TOP + 0.36,
    w: 4.0, h: leftH - 0.46,
    fontFace: FONT_B, margin: 0, valign: "top",
    paraSpaceAfter: 3,
    fit: "shrink", shrinkText: true,
  });

  const rightX = 5.2;
  const rightW = 4.3;
  const afterH = extraRight ? 1.35 : Math.max(1.5, leftH);
  addCard(s, rightX, CONTENT_TOP, rightW, afterH, { strip: C.SUCCESS, fill: C.WHITE });
  s.addText("After", {
    x: rightX + 0.2, y: CONTENT_TOP + 0.08, w: 3.5, h: 0.26,
    fontSize: 12, fontFace: FONT_B, color: C.SUCCESS, bold: true, margin: 0,
  });
  s.addText(String(combinedText), {
    x: rightX + 0.2, y: CONTENT_TOP + 0.36, w: rightW - 0.4, h: afterH - 0.46,
    fontSize: 17, fontFace: FONT_H, color: C.CHARCOAL, margin: 0, valign: "top",
    fit: "shrink", shrinkText: true,
  });

  if (extraRight) extraRight(s, rightX, rightW, CONTENT_TOP + afterH + 0.12);

  const woY = Math.max(CONTENT_TOP + leftH, CONTENT_TOP + afterH + (extraRight ? 1.15 : 0)) + 0.15;
  const woH = 0.78;
  s.addShape("roundRect", {
    x: 0.5, y: woY, w: 9.0, h: woH, rectRadius: 0.08,
    fill: { color: "FDF0EF" },
    line: { color: C.ALERT, width: 1.2 },
  });
  s.addText("Watch out", {
    x: 0.75, y: woY + 0.06, w: 2.0, h: 0.22,
    fontSize: 11, fontFace: FONT_B, color: C.ALERT, bold: true, margin: 0,
  });
  s.addText(String(watchOutText), {
    x: 0.75, y: woY + 0.26, w: 5.8, h: woH - 0.34,
    fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, italic: true, margin: 0,
    valign: "top", fit: "shrink", shrinkText: true,
  });
  s.addText(String(watchOutReason), {
    x: 6.7, y: woY + 0.10, w: 2.5, h: woH - 0.18,
    fontSize: 12, fontFace: FONT_B, color: C.ALERT, margin: 0, valign: "middle",
    fit: "shrink", shrinkText: true,
  });

  addFooter(s, FOOTER);
  if (notes) s.addNotes(notes);
  runSlideDiagnostics(s, pres, { respectSafeBottom: false });
  return s;
}

// ──────────────────────────────────────────────────────────────────────────────
//  Rubric text (teacher notes only — never on slides)
// ──────────────────────────────────────────────────────────────────────────────

const RUBRIC_BLOCK = [
  "5-CRITERIA RUBRIC (mark each 0 or 1, total /5):",
  "  (1) All ideas retained",
  "  (2) One grammatical sentence",
  "  (3) Appropriate clause linking",
  "  (4) No unnecessary repetition",
  "  (5) No fragment, run-on, or comma splice",
  "",
  "MARKING NOTE: If a student writes a comma splice, acknowledge the connection attempt: 'You've recognised these ideas connect - that's a step forward - now add a joining word.'",
].join("\n");

function rubricWithSamples(samples) {
  return RUBRIC_BLOCK + "\n\nSAMPLE STUDENT RESPONSES:\n" + samples.join("\n\n");
}

// ──────────────────────────────────────────────────────────────────────────────
//  Build
// ──────────────────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.title = "Sentence Combining - Year 5/6 Daily Warm-Up";

  // ── Title slide ────────────────────────────────────────────────────────────

  titleSlide(pres, "Sentence Combining", "Year 5/6 Daily Warm-Up",
    "Wednesday, Thursday, Friday  |  15 minutes per session",
    composeNotes({
    say: ["This deck covers three 15-minute warm-up sessions on combining short sentences into one."],
    do: ["Use one day's slides per session. Each day follows the same routine."],
    teacherNotes: "Students have already learned 7 joining words over the previous two weeks: and, but, so, or (coordinating) and because, when, if (subordinating). This week they apply those conjunctions to the new skill of combining multiple short sentences into one.",
    watchFor: ["Students who have not internalised the 7 conjunctions may need a reference card on their desk."],
  }));

  // ── Teacher overview / resources ───────────────────────────────────────────

  contentSlide(pres, "TEACHER OVERVIEW", C.PRIMARY,
    "3-Day Warm-Up Structure",
    [
      "Each day: Hook (2 min) -> Teach (3 min) -> Practise Together (4 min) -> Your Turn (4 min) -> Exit Check (2 min)",
      "Wednesday: combine 3 sentences (pets) + comma rule reinforcement",
      "Thursday: combine 4 sentences (recess) - focus on identifying and deleting repetition",
      "Friday: combine 5 sentences (specialist lessons) - full destination skill",
      "Materials: student workbooks only. No printed resources, no whiteboards.",
      "Scaffolds on slides are hidden by default. Reveal only for focus students who are stuck.",
    ],
    composeNotes({
      say: ["This is a teacher-facing overview slide. Do not display to students."],
      do: ["Read before your first session. Each day is self-contained."],
      teacherNotes: [
        "Pre-assessment: students can use 7 conjunctions in their own sentences.",
        "New skill this week: combining MULTIPLE short sentences into one by (a) choosing the main idea, (b) deleting repeated subjects/nouns, (c) selecting conjunctions.",
        "Friday's exit check is the closest rehearsal for the post-test in two weeks. Mark strictly against the 5-criteria rubric on Friday.",
        "Victorian Curriculum 2.0 alignment: Year 5/6 English, Language strand - Text structure and organisation; Expressing and developing ideas (VCELA323 - main and subordinate clauses).",
        "HITS targeted: Setting goals, Worked examples, Explicit teaching, Feedback.",
      ],
      watchFor: ["Teachers across all 5 classes should mark exit checks consistently using the rubric in the exit check speaker notes."],
    }),
    FOOTER);

  // ════════════════════════════════════════════════════════════════════════════
  //  WEDNESDAY — Combine 3 sentences (pets) + comma rule
  // ════════════════════════════════════════════════════════════════════════════

  const wedHook = [
    "My cat is playful.",
    "My cat chases butterflies.",
    "My cat always comes home for dinner.",
  ];

  dayDivider(pres, "Wednesday", 1,
    "I can combine 3 short sentences into one sentence by deleting repetition and using joining words.");

  hookSlide(pres, wedHook, composeNotes({
    say: [
      "You have three short sentences on the screen.",
      "Copy them into your workbook exactly as they are.",
      "At the end of today's warm-up, you will combine them into one sentence.",
    ],
    do: [
      "Display slide. Give 90 seconds for copying.",
      "Circulate to check all students have copied all three sentences.",
    ],
    teacherNotes: "This is the 'before' snapshot. Students return to these sentences for the exit check. (2 min)",
    watchFor: [
      "Students who start combining early - redirect them to copy first.",
      "Students who skip a sentence.",
    ],
    tag: "Hook | 2 min",
  }));

  // ── Wednesday Teach ────────────────────────────────────────────────────────

  teachSlideCustom(pres,
    [
      strikeRuns([
        { text: "The dog", strike: true },
        { text: " barked at the postman.", strike: false },
      ]),
      strikeRuns([
        { text: "The dog", strike: true },
        { text: " was always excited.", strike: false },
      ]),
      strikeRuns([
        { text: "The dog", strike: true },
        { text: " ran to the fence.", strike: false },
      ]),
    ],
    "Because the dog was always excited, it barked at the postman and ran to the fence.",
    "The dog barked at the postman and the dog was always excited and the dog ran to the fence.",
    "Repeats 'the dog' three times. Uses 'and' for every join. Runs on without proper punctuation.",
    composeNotes({
      say: [
        "Watch how I turn three short sentences into one.",
        "First I find the repeated words. 'The dog' appears three times. I only need it once.",
        "Ask: Which idea should be the main part of my sentence? I'll pick 'barked at the postman' as my main action.",
        "Now I choose joining words. 'Because' tells my reader the reason. 'And' links two actions.",
        "Notice the comma after 'Because the dog was always excited' - that's our comma rule. When a joining word like 'because' starts the sentence, put a comma after the first part.",
        "If I flip it around - 'The dog barked at the postman because it was always excited' - no comma needed in the middle.",
        "Now look at the Watch Out example. What went wrong?",
      ],
      do: [
        "Point to each crossed-out 'The dog' as you explain the deletion.",
        "Point to the conjunctions in the combined sentence.",
        "Tap the comma rule card on the right to draw attention to the two examples.",
      ],
      teacherNotes: "Key teaching move: show that deletion of repetition and conjunction choice work together. The comma rule reinforcement is secondary but important. (3 min)",
      watchFor: [
        "Students who think every 'and' chain is wrong - clarify that 'and' is fine, but repeating the subject AND using only 'and' is the problem.",
        "Watch for comma splices in student practice - two ideas joined with only a comma and no joining word. Example: 'The dog barked, it ran to the fence.' Acknowledge the connection attempt, then prompt: 'You need a joining word after that comma. Which one fits?'",
      ],
      tag: "Teach | 3 min | Worked example + comma rule",
    }),
    (s, rx, rw, extraY) => {
      addCard(s, rx, extraY, rw, 1.0, { strip: C.ACCENT, fill: C.BG_LIGHT });
      s.addText("Comma rule", {
        x: rx + 0.2, y: extraY + 0.06, w: 3.0, h: 0.22,
        fontSize: 11, fontFace: FONT_B, color: C.ACCENT, bold: true, margin: 0,
      });
      s.addText([
        { text: "Because the dog was excited", options: { fontSize: 12, color: C.CHARCOAL, fontFace: FONT_B } },
        { text: ",", options: { fontSize: 12, color: C.SUCCESS, bold: true, fontFace: FONT_B } },
        { text: " it barked...  ", options: { fontSize: 12, color: C.CHARCOAL, fontFace: FONT_B } },
        { text: "comma", options: { fontSize: 11, color: C.SUCCESS, bold: true, fontFace: FONT_B } },
      ], {
        x: rx + 0.2, y: extraY + 0.28, w: rw - 0.4, h: 0.28, margin: 0, valign: "middle",
      });
      s.addText([
        { text: "It barked... because the dog was excited.  ", options: { fontSize: 12, color: C.CHARCOAL, fontFace: FONT_B } },
        { text: "no comma", options: { fontSize: 11, color: C.ALERT, bold: true, fontFace: FONT_B } },
      ], {
        x: rx + 0.2, y: extraY + 0.58, w: rw - 0.4, h: 0.28, margin: 0, valign: "middle",
      });
    }
  );

  // ── Wednesday Practise Together ────────────────────────────────────────────

  practiceSlides(pres,
    [
      "The rabbit hopped around the garden.",
      "The rabbit was small and fluffy.",
      "The rabbit found a carrot.",
    ],
    [
      "The kitten climbed the curtains.",
      "The kitten was tiny.",
      "The kitten got stuck at the top.",
    ],
    "The small and fluffy rabbit hopped around the garden and found a carrot.",
    "The tiny kitten climbed the curtains but got stuck at the top.",
    composeNotes({
      say: [
        "Two sets of sentences on the screen. Pick either set - or try both if you are quick.",
        "Write your combined sentence in your workbook now. I will wait for everyone before I show the model.",
      ],
      do: [
        "Students write FIRST, teacher reveals SECOND. Do not reveal before every student has attempted.",
        "Circulate while students write. Note which students default to and-and-and chains.",
      ],
      teacherNotes: "4 min total for this slide. Allow 2-2.5 min writing, then reveal models. (4 min)",
      enabling: "Focus students: circle the repeated words first with a coloured pencil. Cross them out. Then try joining what is left.",
      watchFor: [
        "Students who write two separate sentences instead of one.",
        "Students who keep every 'the rabbit' / 'the kitten' without deleting.",
      ],
      tag: "Guided practice | 4 min",
    }),
    composeNotes({
      say: [
        "Here are two ways I could have combined them. Your version might be different and still correct.",
        "Check: did you delete the repeated subject? Did you use a joining word?",
      ],
      do: [
        "Read both model answers aloud.",
        "Ask 2-3 students to share a version that is different from the model.",
      ],
      teacherNotes: "There is no single correct answer. Accept any grammatically correct sentence that retains all ideas, deletes repetition, and uses appropriate conjunctions.",
      watchFor: ["Students whose version is correct but different from the model - affirm their work."],
      tag: "Guided practice reveal | included in 4 min",
    })
  );

  // ── Wednesday Your Turn ────────────────────────────────────────────────────

  yourTurnSlides(pres,
    [
      "The puppy chewed a shoe.",
      "The puppy was bored.",
      "The puppy wagged its tail when we got home.",
    ],
    "The puppy chewed a shoe because _______, but _______.",
    composeNotes({
      say: [
        "Fresh set of sentences. Combine them into one sentence on your own.",
        "If you finish quickly, try combining them a second way and decide which version sounds better.",
      ],
      do: [
        "Circulate. Scaffold is hidden on the next slide - reveal only for focus students who are stuck after 1 minute.",
        "Do not reveal the scaffold to the whole class.",
      ],
      teacherNotes: "4 min. The scaffold is a sentence stem on the next slide, visible only to the teacher until revealed. (4 min)",
      enabling: "Focus students: reveal the scaffold slide. Point to the sentence stem and read it aloud with the student.",
      extending: [
        "Combine the same three sentences TWO different valid ways. Which version is clearer or has a better rhythm? Write both and put a star next to your preferred version.",
      ],
      watchFor: [
        "Students who default to 'and' for every join.",
        "Students who write two sentences instead of one.",
      ],
      tag: "Independent practice | 4 min",
    }),
    composeNotes({
      say: ["If you need help getting started, here is a sentence stem to guide you."],
      do: ["Reveal this slide ONLY for focus students who are stuck. Do not show the whole class."],
      teacherNotes: "This scaffold reduces the task from full composition to gap-filling. Use only when a student cannot start after 1 minute of independent effort.",
      watchFor: ["Students who copy the scaffold exactly without thinking about what goes in the gaps."],
      tag: "Scaffold (hidden by default)",
    })
  );

  // ── Wednesday Exit Check ───────────────────────────────────────────────────

  exitCheckSlide(pres, wedHook, composeNotes({
    say: [
      "Remember those sentences you copied at the start? Now combine them into one sentence in your workbook.",
      "You have about 90 seconds. Show me what you have learned today.",
    ],
    do: [
      "Students write independently in workbooks.",
      "Mark each focus student's exit check against the 5-criteria rubric below.",
    ],
    teacherNotes: [
      "2 min. This is the 'after' picture. Students see their own 15-minute growth.",
      "",
      rubricWithSamples([
        "Sample 1 (5/5): 'My playful cat chases butterflies but always comes home for dinner.'\n  (1) All ideas: 1 | (2) One sentence: 1 | (3) Linking: 1 | (4) No repetition: 1 | (5) No errors: 1",
        "Sample 2 (3/5): 'My cat is playful and my cat chases butterflies and my cat always comes home for dinner.'\n  (1) All ideas: 1 | (2) One sentence: 1 | (3) Linking: 0 (and-and-and) | (4) No repetition: 0 (my cat x3) | (5) No errors: 1",
        "Sample 3 (2/5): 'My cat is playful, chases butterflies. Always comes home for dinner.'\n  (1) All ideas: 1 | (2) One sentence: 0 (fragment) | (3) Linking: 0 | (4) No repetition: 1 | (5) No errors: 0 (fragment)",
      ]),
    ],
    watchFor: [
      "Comma splices - acknowledge the connection attempt, then prompt for a joining word.",
      "Students who write multiple sentences instead of one.",
    ],
    tag: "Exit check | 2 min",
  }));

  // ════════════════════════════════════════════════════════════════════════════
  //  THURSDAY — Combine 4 sentences (recess/lunch) — deletion focus
  // ════════════════════════════════════════════════════════════════════════════

  const thuHook = [
    "The students played handball at recess.",
    "The students argued about the rules.",
    "The handball game lasted the whole break.",
    "The students were exhausted afterwards.",
  ];

  dayDivider(pres, "Thursday", 2,
    "I can combine 4 short sentences into one sentence by identifying and deleting repeated words.");

  hookSlide(pres, thuHook, composeNotes({
    say: [
      "Four sentences today. Copy them into your workbook exactly.",
      "You will combine them into one sentence at the end.",
    ],
    do: [
      "Display slide. Give 2 minutes for copying.",
      "Circulate to check all four sentences are copied.",
    ],
    teacherNotes: "Same routine as yesterday. Students should recognise the pattern. (2 min)",
    watchFor: ["Students who miss the fourth sentence."],
    tag: "Hook | 2 min",
  }));

  // ── Thursday Teach ─────────────────────────────────────────────────────────

  teachSlideCustom(pres,
    [
      strikeRuns([
        { text: "The children", strike: true },
        { text: " sat on the grass.", strike: false },
      ]),
      strikeRuns([
        { text: "The children", strike: true },
        { text: " ate their lunch.", strike: false },
      ]),
      strikeRuns([
        { text: "The ", strike: true },
        { text: "lunch", strike: true },
        { text: " was delicious.", strike: false },
      ]),
      strikeRuns([
        { text: "The children", strike: true },
        { text: " talked about their weekend.", strike: false },
      ]),
    ],
    "The children sat on the grass, ate their delicious lunch, and talked about their weekend.",
    "The children sat on the grass and the children ate their lunch and the lunch was delicious and the children talked about their weekend.",
    "Repeats 'the children' 3 times, 'the lunch' twice. Four 'and' connectors. No deletion.",
    composeNotes({
      say: [
        "Today we have four sentences to combine. The big skill today is spotting and crossing out the repeated words.",
        "How many times does 'the children' appear? Three times. I only need it once.",
        "And 'the lunch' appears twice. I can fold 'delicious' into the sentence instead of giving 'lunch' its own sentence.",
        "Watch: I cross out the copies, keep the actions, and join them with commas and 'and'.",
        "Ask: What did I delete? What did I keep? [Expected: deleted repeated subjects, kept the actions and details]",
      ],
      do: [
        "Point to each crossed-out phrase as you say it.",
        "Use a finger or pointer to trace from the 'before' sentences to the 'after' sentence.",
      ],
      teacherNotes: [
        "Research shows deletion of repetition is the hardest sub-skill. Today's teach deliberately foregrounds the visual strikethrough. (3 min)",
        "Explicitly name the 'delicious' fold: we kept the IDEA from 'The lunch was delicious' but folded it into 'their delicious lunch'. The whole sentence was deleted but its meaning was kept as a single word. Without this callout, some students think they must delete entire ideas, which contradicts criterion 1 (all ideas retained).",
      ],
      watchFor: [
        "Students who think they must keep every word from every sentence.",
        "Students who delete important ideas along with the repeated subjects.",
        "Students who think folding an idea into an adjective means they have lost the idea - reassure them.",
      ],
      tag: "Teach | 3 min | Deletion focus",
    })
  );

  // ── Thursday Practise Together ─────────────────────────────────────────────

  practiceSlides(pres,
    [
      "The friends played soccer on the oval.",
      "The soccer ball went over the fence.",
      "The teacher helped them get it back.",
      "The friends started a new game.",
    ],
    [
      "The girls sat under a tree.",
      "The tree gave them shade.",
      "They shared their fruit.",
      "They planned what to do after school.",
    ],
    "The friends played soccer on the oval, but the ball went over the fence, so the teacher helped them get it back and they started a new game.",
    "The girls sat under a tree that gave them shade, shared their fruit, and planned what to do after school.",
    composeNotes({
      say: [
        "Two sets of four sentences. Write your combined sentence in your workbook first.",
        "Before you start writing, quickly circle any repeated words you can see.",
      ],
      do: [
        "Students write FIRST, teacher reveals SECOND. Do not reveal before every student has attempted.",
        "Circulate. Prompt students who are stuck: 'Which words appear more than once? Cross out the copies first.'",
      ],
      teacherNotes: "4 min total. Allow 2.5 min writing, then reveal. (4 min)",
      enabling: "Focus students: use coloured pencils to underline repeated words across the four sentences. Cross out the copies. Then join what is left.",
      watchFor: [
        "Students who delete important ideas along with the repeated subject.",
        "Students who keep all four subjects without any deletion.",
      ],
      tag: "Guided practice | 4 min",
    }),
    composeNotes({
      say: [
        "Compare your sentence with the model. Your version might be different and still correct.",
        "The key check: did you delete the repeated words? Did you keep all the ideas?",
      ],
      do: [
        "Read both models aloud. Ask students to put a tick if they deleted the repeated subject, a cross if they kept it.",
      ],
      teacherNotes: "Accept any grammatically correct version that retains all four ideas and deletes repetition.",
      watchFor: ["Students whose sentence is correct but uses only 'and' - note this for Friday."],
      tag: "Guided practice reveal | included in 4 min",
    })
  );

  // ── Thursday Your Turn ─────────────────────────────────────────────────────

  yourTurnSlides(pres,
    [
      "The boys raced to the drinking fountain.",
      "The drinking fountain was near the library.",
      "They were hot from playing.",
      "They splashed water on their faces.",
    ],
    "The boys raced to the drinking fountain near _______ because _______, and they _______.",
    composeNotes({
      say: [
        "Four sentences. Combine them on your own.",
        "Start by circling repeated words, then cross them out, then write your combined sentence.",
      ],
      do: [
        "Circulate. Scaffold is hidden on the next slide for focus students only.",
      ],
      teacherNotes: "4 min. Focus on whether students are actively deleting repetition before writing. (4 min)",
      enabling: "Focus students: reveal the scaffold. Read it aloud with the student and point to where each gap matches an idea from the original sentences.",
      extending: [
        "Add a subordinate clause that tells WHERE, WHEN, or WHY to your combined sentence. Use a less common joining word such as although, since, while, or until.",
      ],
      watchFor: [
        "Students who keep 'the drinking fountain' twice without folding the detail in.",
        "Students who lose the 'hot from playing' idea.",
      ],
      tag: "Independent practice | 4 min",
    }),
    composeNotes({
      say: ["Here is a sentence stem if you need help getting started."],
      do: ["Reveal ONLY for focus students who cannot start after 1 minute."],
      teacherNotes: "Same protocol as Wednesday. Scaffold visible only to struggling students.",
      watchFor: ["Students who copy the scaffold word-for-word without filling the gaps meaningfully."],
      tag: "Scaffold (hidden by default)",
    })
  );

  // ── Thursday Exit Check ────────────────────────────────────────────────────

  exitCheckSlide(pres, thuHook, composeNotes({
    say: [
      "Back to your sentences from the start. Combine all four into one sentence.",
      "Remember: circle the repeated words, cross them out, choose your joining words, write.",
    ],
    do: [
      "Students write independently. Mark focus students against the 5-criteria rubric.",
    ],
    teacherNotes: [
      "2 min. Compare today's exit check with yesterday's to track growth in deletion skill.",
      "",
      rubricWithSamples([
        "Sample 1 (5/5): 'The students played handball at recess and argued about the rules, and although the game lasted the whole break, they were exhausted afterwards.'\n  (1) All ideas: 1 | (2) One sentence: 1 | (3) Linking: 1 | (4) No repetition: 1 | (5) No errors: 1",
        "Sample 2 (3/5): 'The students played handball at recess and the students argued about the rules and the handball game lasted the whole break and the students were exhausted.'\n  (1) All ideas: 1 | (2) One sentence: 1 | (3) Linking: 0 (and-and-and) | (4) No repetition: 0 (the students x3) | (5) No errors: 1",
        "Sample 3 (1/5): 'The students played handball. They argued, the game lasted the whole break.'\n  (1) All ideas: 0 (missing exhaustion) | (2) One sentence: 0 (two sentences) | (3) Linking: 0 | (4) No repetition: 1 | (5) No errors: 0 (comma splice)",
      ]),
    ],
    watchFor: [
      "Students who still keep all three instances of 'the students' without deletion.",
      "Comma splices between ideas.",
    ],
    tag: "Exit check | 2 min",
  }));

  // ════════════════════════════════════════════════════════════════════════════
  //  FRIDAY — Combine 5 sentences (specialist lessons) — destination skill
  // ════════════════════════════════════════════════════════════════════════════

  const friHook = [
    "The students went to music class.",
    "They learned a new song.",
    "The song had a tricky rhythm.",
    "The students clapped along.",
    "Most students could sing it by the end.",
  ];

  dayDivider(pres, "Friday", 3,
    "I can combine 5 short sentences into one sentence using all three combining moves.");

  hookSlide(pres, friHook, composeNotes({
    say: [
      "Five sentences today. This is the destination skill - the full task.",
      "Copy all five into your workbook. You will combine them at the end.",
    ],
    do: [
      "Display slide. Give 2 minutes for copying.",
      "Circulate to check all five sentences are copied.",
    ],
    teacherNotes: "Same routine. By now students should recognise the hook-to-exit-check structure. (2 min)",
    watchFor: ["Students who rush and skip the fourth or fifth sentence."],
    tag: "Hook | 2 min",
  }));

  // ── Friday Teach ───────────────────────────────────────────────────────────

  teachSlideCustom(pres,
    [
      strikeRuns([
        { text: "The class", strike: true },
        { text: " walked to the art room.", strike: false },
      ]),
      strikeRuns([
        { text: "The art room", strike: true },
        { text: " was bright and colourful.", strike: false },
      ]),
      strikeRuns([
        { text: "The ", strike: true },
        { text: "teacher showed ", strike: false },
        { text: "them", strike: true },
        { text: " how to mix paint.", strike: false },
      ]),
      strikeRuns([
        { text: "The students", strike: true },
        { text: " painted a landscape.", strike: false },
      ]),
      strikeRuns([
        { text: "The landscape", strike: true },
        { text: " included mountains and a river.", strike: false },
      ]),
    ],
    "The class walked to the bright and colourful art room, where the teacher showed them how to mix paint, and they painted a landscape with mountains and a river.",
    "The class walked to the art room and the art room was bright and colourful and the teacher showed them how to mix paint and the students painted a landscape and the landscape included mountains and a river.",
    "Five 'and' connectors. No deletion. No variety in joining words.",
    composeNotes({
      say: [
        "Five sentences today - this is the full skill.",
        "Same three moves: find repetition, delete it, choose joining words.",
        "Look at the repeated nouns. 'The art room' appears twice. 'The landscape' appears twice. 'The class' and 'the students' are the same group. I only need each one once.",
        "Watch how 'bright and colourful' folds into the sentence as a describing phrase. 'Mountains and a river' folds in with 'that'.",
        "I used 'where', 'and', 'that' - three different joining words. Much better than five 'ands'.",
        "Now look at the Watch Out. Five ands and zero deletions.",
      ],
      do: [
        "Point to each deletion. Trace the flow from before to after.",
        "Count the 'ands' in the Watch Out aloud with students.",
      ],
      teacherNotes: "This is the destination skill. Emphasise that all three moves (identify main idea, delete repetition, choose conjunctions) work together. (3 min)",
      watchFor: [
        "Students who feel overwhelmed by five sentences - reassure them the process is the same as three.",
      ],
      tag: "Teach | 3 min | Full destination skill",
    })
  );

  // ── Friday Practise Together ───────────────────────────────────────────────

  practiceSlides(pres,
    [
      "The students had PE on the oval.",
      "They practised sprinting in lanes.",
      "The sprints were timed.",
      "Some students beat their personal best.",
      "The teacher cheered for everyone.",
    ],
    [
      "The class visited the library.",
      "The library was quiet and cool.",
      "They returned their old books.",
      "They chose new ones.",
      "The librarian recommended a mystery series.",
    ],
    "The students had PE on the oval, where they practised timed sprinting in lanes, and some students beat their personal best, so the teacher cheered for everyone.",
    "The class visited the quiet and cool library, returned their old books, chose new ones, and the librarian recommended a mystery series.",
    composeNotes({
      say: [
        "Five sentences in each set. Circle the repeated words, cross them out, then write one combined sentence.",
        "Write in your workbook first. I will wait for everyone.",
      ],
      do: [
        "Students write FIRST, teacher reveals SECOND.",
        "Circulate. This is the hardest practice yet - give an extra 30 seconds if needed.",
      ],
      teacherNotes: "4 min. May need slightly more writing time given 5 sentences. (4 min)",
      enabling: "Focus students: start by combining just three of the five sentences. Then try to add the other two.",
      watchFor: [
        "Students who lose ideas when combining five sentences - prompt them to check each original idea is present.",
        "Students who write correct but very long run-on sentences without commas.",
      ],
      tag: "Guided practice | 4 min",
    }),
    composeNotes({
      say: [
        "Here are two model answers. Compare with yours.",
        "Check: did you keep all five ideas? Did you delete the repeated words? Did you use different joining words?",
      ],
      do: [
        "Read both models aloud. Ask students to count how many original ideas they kept.",
      ],
      teacherNotes: "Accept any grammatically correct version that retains all five ideas.",
      watchFor: ["Students who dropped one or two ideas - prompt them to recount."],
      tag: "Guided practice reveal | included in 4 min",
    })
  );

  // ── Friday Your Turn ───────────────────────────────────────────────────────

  yourTurnSlides(pres,
    [
      "The students walked to the STEM room.",
      "The STEM room had new equipment.",
      "They worked in pairs.",
      "They designed a bridge out of straws.",
      "The bridge had to hold a heavy book.",
    ],
    "The students walked to the STEM room, which had _______, and worked in pairs to _______ that _______.",
    composeNotes({
      say: [
        "Five sentences. This is your chance to show me you can do the full skill.",
        "Combine all five into one sentence on your own.",
      ],
      do: [
        "Circulate. Scaffold on next slide for focus students only.",
      ],
      teacherNotes: "4 min. This is the last independent practice before the exit check. Note students who are still struggling - you may need to plan a small-group reteach next week. (4 min)",
      enabling: "Focus students: reveal the scaffold. Help them map each gap to an idea from the original sentences.",
      extending: [
        "Combine all 5 original sentences PLUS this 6th sentence: 'The winning bridge held three heavy books.' Write one sentence using at least three different joining words.",
      ],
      watchFor: [
        "Students who write two sentences instead of one.",
        "Students who use only 'and' for all joins.",
      ],
      tag: "Independent practice | 4 min",
    }),
    composeNotes({
      say: ["Here is a sentence stem if you need help."],
      do: ["Reveal ONLY for focus students who cannot start."],
      teacherNotes: "Same protocol. Friday's scaffold is the last supported attempt before the post-test.",
      watchFor: ["Students who need the scaffold on Friday may need targeted small-group practice next week."],
      tag: "Scaffold (hidden by default)",
    })
  );

  // ── Friday Exit Check ──────────────────────────────────────────────────────

  exitCheckSlide(pres, friHook, composeNotes({
    say: [
      "Back to your music sentences. Combine all five into one sentence.",
      "This is the same task you will do in your post-test in two weeks. Show me your best work.",
    ],
    do: [
      "Students write independently. Mark EVERY focus student against the 5-criteria rubric.",
      "This is the strictest marking day. Record scores to compare with the post-test.",
    ],
    teacherNotes: [
      "2 min. This exit check most closely mirrors the post-test format. Mark strictly so you have a true read on where students sit.",
      "",
      rubricWithSamples([
        "Sample 1a (5/5 - complex): 'The students went to music class, where they learned a new song with a tricky rhythm, and they clapped along until most of them could sing it by the end.'\n  (1) All ideas: 1 | (2) One sentence: 1 | (3) Linking: 1 | (4) No repetition: 1 | (5) No errors: 1",
        "Sample 1b (5/5 - simpler): 'In music class, the students learned a new song with a tricky rhythm, clapped along, and most could sing it by the end.'\n  (1) All ideas: 1 | (2) One sentence: 1 | (3) Linking: 1 | (4) No repetition: 1 | (5) No errors: 1\n  NOTE: Both versions score 5/5. Students do not need to write 25-word sentences to get full marks. Either level of complexity is acceptable.",
        "Sample 2 (3/5): 'The students went to music class and they learned a new song and the song had a tricky rhythm and the students clapped along and most students could sing it by the end.'\n  (1) All ideas: 1 | (2) One sentence: 1 | (3) Linking: 0 (and-and-and) | (4) No repetition: 0 (the students/the song) | (5) No errors: 1",
        "Sample 3 (2/5): 'The students went to music class, they learned a tricky song and clapped along and could sing it.'\n  (1) All ideas: 0 (missing 'most') | (2) One sentence: 0 (comma splice) | (3) Linking: 1 | (4) No repetition: 1 | (5) No errors: 0 (comma splice)",
      ]),
    ],
    watchFor: [
      "Students who drop one or two ideas from the five sentences.",
      "Comma splices - the most common error at this level.",
      "Students who still cannot delete repeated subjects after three days - flag for small-group reteach.",
    ],
    tag: "Exit check | 2 min | POST-TEST REHEARSAL",
  }));

  // ── Write file ─────────────────────────────────────────────────────────────

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "Sentence Combining Year 56 Warm-Up.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);
}

build().catch((err) => {
  console.error(err);
  process.exit(1);
});
