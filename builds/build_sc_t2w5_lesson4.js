"use strict";

/**
 * Sentence Combining Warm-Up - Year 5/6 Enrichment, Term 2 Week 5, Day 4 (Thu)
 * MIX: BECAUSE / WHEN / IF. Choose the right subordinator from context.
 * Topic: a school day from morning to home time.
 *
 * Builders adapted for mixed practice:
 *   - Teach slide shows the three meanings side by side as a quick reference.
 *   - Guided slide: each pair takes a different subordinator; reveal highlights
 *     the chosen word per pair (not one fixed word for the whole slide).
 *   - Independent: student picks the subordinator that fits the meaning.
 *   - Exit: student picks BECAUSE / WHEN / IF for the given ideas.
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

const FOOTER = "Sentence Combining Warm-Up | Year 5/6 | Term 2 Week 5 | Day 4 Thu";
const OUT_DIR = "output/SC_T2W5_Lesson4_Thu_MIX";
fs.mkdirSync(OUT_DIR, { recursive: true });

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Sentence Combining Warm-Up - Year 5/6 - Term 2 Week 5 Day 4 (Thu)";
pres.author = "Year 5/6 Enrichment PLC";

function badgeStack(s, phaseText, phaseColor, dayText, dayColor) {
  addBadge(s, phaseText, { color: phaseColor, w: 1.6, x: 0.5, y: 0.20 });
  addBadge(s, dayText, { color: dayColor || C.SECONDARY, w: 1.05, x: 2.18, y: 0.20 });
}

function highlightWord(sentence, word, baseSize, baseColor, highlightColor) {
  const re = new RegExp(`(\\s|^)(${word})(\\s|[.,?!])`, "i");
  const m = sentence.match(re);
  if (!m) {
    return [{ text: sentence, options: { fontSize: baseSize, color: baseColor } }];
  }
  const before = sentence.slice(0, m.index + m[1].length);
  const w = m[2];
  const after = sentence.slice(m.index + m[1].length + w.length);
  return [
    { text: before, options: { fontSize: baseSize, color: baseColor } },
    { text: w,      options: { fontSize: baseSize, color: highlightColor, bold: true } },
    { text: after,  options: { fontSize: baseSize, color: baseColor } },
  ];
}

const THU = { dayText: "Thu", dayColor: C.ALERT };

const SC_THU = [
  "I can find two short ideas that go together.",
  "I can choose the right joining word: BECAUSE, WHEN or IF.",
  "I can write a sentence that makes sense.",
];

const MARK_THU = [
  "Two ideas combined into ONE sentence",
  "Chose the joining word that matches the meaning",
  "Sentence is grammatical (no run-on, no comma splice)",
];

/* ─────────────────────────────────────────────────────────────────────────
 *  Slide 1: Day intro - review all three meanings
 * ───────────────────────────────────────────────────────────────────────── */

(function dayIntro() {
  const s = pres.addSlide();
  addTopBar(s, C.PRIMARY);
  badgeStack(s, "Day start", C.PRIMARY, THU.dayText, THU.dayColor);
  addTitle(s, "Thu - Choose the right joining word");

  const cardH = SAFE_BOTTOM - CONTENT_TOP;
  const leftW = 4.6;
  const rightW = 4.4;
  const rightX = 5.2;

  addCard(s, 0.5, CONTENT_TOP, leftW, cardH, { strip: C.SECONDARY, fill: C.WHITE });
  s.addText("Review - this week's three joining words", {
    x: 0.7, y: CONTENT_TOP + 0.14, w: leftW - 0.4, h: 0.34,
    fontSize: 13, fontFace: FONT_B, color: C.SECONDARY, bold: true, margin: 0,
  });

  // Three small word + meaning rows
  const rows = [
    { word: "BECAUSE", meaning: "joins a REASON", color: C.PRIMARY },
    { word: "WHEN",    meaning: "joins a TIME",   color: C.SECONDARY },
    { word: "IF",      meaning: "joins a CONDITION", color: C.ACCENT },
  ];
  const rowsY = CONTENT_TOP + 0.56;
  const rowH = (cardH - 0.72) / rows.length;
  rows.forEach((r, i) => {
    const y = rowsY + i * rowH;
    addTextOnShape(s, r.word, {
      x: 0.7, y: y + 0.06, w: 1.8, h: rowH - 0.18, rectRadius: 0.06,
      fill: { color: r.color },
    }, {
      fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
    s.addText(r.meaning, {
      x: 2.6, y: y, w: leftW - 2.2, h: rowH,
      fontSize: 15, fontFace: FONT_B, color: C.CHARCOAL, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });
  });

  addCard(s, rightX, CONTENT_TOP, rightW, cardH, { strip: C.ACCENT, fill: C.BG_CARD });
  s.addText("Today's success criteria - I can...", {
    x: rightX + 0.2, y: CONTENT_TOP + 0.14, w: rightW - 0.4, h: 0.34,
    fontSize: 13, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
  });
  s.addText(SC_THU.map((t, i) => ({
    text: t,
    options: {
      bullet: true,
      breakLine: i < SC_THU.length - 1,
      fontSize: 15, color: C.CHARCOAL, paraSpaceAfter: 6,
    },
  })), {
    x: rightX + 0.2, y: CONTENT_TOP + 0.54, w: rightW - 0.4, h: cardH - 0.68,
    fontFace: FONT_B, valign: "top", margin: 0,
    fit: "shrink", shrinkText: true,
  });

  addFooter(s, FOOTER);
  s.addNotes(composeNotes({
    say: [
      "Quick review of this week's three joining words.",
      "BECAUSE adds the REASON. WHEN adds the TIME. IF adds the CONDITION.",
      "Today you mix them. Today's job is to choose the RIGHT one for the meaning.",
      "Topic - a school day from morning to home time.",
    ],
    do: [
      "Point to each word and read the meaning aloud.",
      "Quick choral check - 'BECAUSE means...', 'WHEN means...', 'IF means...'",
    ],
    teacherNotes: [
      "Choosing the right subordinator is the new move today. Up until now the joining word was given.",
      "Today supports rubric criteria 2 (one grammatical sentence), 3 (appropriate clause linking) and 5 (no comma splice).",
    ],
    enabling: [
      "FOCUS GROUP: keep the three-meaning card visible by leaving the slide up for the whole lesson, or jot the three words on the side board.",
      "EXTENSION: ask higher-ability students if they remember WHILE, AS, UNLESS or SINCE - rarer subordinators they could use today.",
    ],
    watchFor: [
      "Students defaulting to BECAUSE for every pair - prompt them to read aloud and ask 'is this a reason, a time, or a condition?'",
    ],
    tag: "[Day 4 | Thu | Review all + LI/SC]",
  }, { requireSay: false, requireDo: false }));
  runSlideDiagnostics(s, pres, { respectSafeBottom: false });
})();

/* ─────────────────────────────────────────────────────────────────────────
 *  Slide 2: Teach - three mini-examples side by side, then non-example
 * ───────────────────────────────────────────────────────────────────────── */

(function teachMix() {
  const s = pres.addSlide();
  addTopBar(s, C.PRIMARY);
  badgeStack(s, "Teach", C.PRIMARY, THU.dayText, THU.dayColor);
  addTitle(s, "How to choose - meaning decides the word");

  // Three side-by-side mini-cards (each = one worked example)
  const examples = [
    {
      word: "BECAUSE", meaning: "reason", color: C.PRIMARY,
      sentence: "I was tired BECAUSE I had stayed up late.",
    },
    {
      word: "WHEN", meaning: "time", color: C.SECONDARY,
      sentence: "The siren rang WHEN it was time for lunch.",
    },
    {
      word: "IF", meaning: "condition", color: C.ACCENT,
      sentence: "We can play handball IF it isn't raining.",
    },
  ];

  const topRowY = CONTENT_TOP;
  const topRowH = 2.30;
  const colW = 2.93;
  const gap = 0.10;

  examples.forEach((ex, i) => {
    const x = 0.5 + i * (colW + gap);
    addCard(s, x, topRowY, colW, topRowH, { strip: ex.color, fill: C.WHITE });
    s.addText(ex.word, {
      x: x + 0.16, y: topRowY + 0.12, w: colW - 0.32, h: 0.34,
      fontSize: 18, fontFace: FONT_H, color: ex.color, bold: true,
      align: "center", margin: 0,
    });
    s.addText(`= ${ex.meaning}`, {
      x: x + 0.16, y: topRowY + 0.48, w: colW - 0.32, h: 0.28,
      fontSize: 12, fontFace: FONT_B, color: C.MUTED, italic: true,
      align: "center", margin: 0,
    });
    const runs = highlightWord(ex.sentence, ex.word, 13, C.CHARCOAL, ex.color);
    s.addText(runs.map((run, j) => ({
      text: `${j === 0 ? '"' : ""}${run.text}${j === runs.length - 1 ? '"' : ""}`,
      options: { ...run.options, fontFace: FONT_H, breakLine: false },
    })), {
      x: x + 0.16, y: topRowY + 0.84, w: colW - 0.32, h: topRowH - 0.96,
      italic: true, valign: "top", margin: 0,
      fit: "shrink", shrinkText: true,
    });
  });

  // Non-example: wrong subordinator chosen
  const neY = topRowY + topRowH + 0.16;
  const neH = SAFE_BOTTOM - neY;
  addCard(s, 0.5, neY, 9, neH, { strip: C.ALERT, fill: C.WHITE });
  s.addText("Watch out - wrong joining word changes the meaning", {
    x: 0.75, y: neY + 0.10, w: 7, h: 0.28,
    fontSize: 12, fontFace: FONT_B, color: C.ALERT, bold: true, margin: 0,
  });
  s.addText(`Wrong: "I was tired WHEN I had stayed up late."`, {
    x: 0.75, y: neY + 0.40, w: 8.5, h: 0.34,
    fontSize: 14.5, fontFace: FONT_H, color: C.CHARCOAL, italic: true, margin: 0,
    fit: "shrink", shrinkText: true,
  });
  s.addText(
    "WHEN says the tired feeling and the staying up late happened at the same time - that's not the meaning. BECAUSE shows that the staying up late CAUSED the tired feeling. Choose the word that matches the meaning.",
    {
      x: 0.75, y: neY + 0.74, w: 8.5, h: neH - 0.84,
      fontSize: 11.5, fontFace: FONT_B, color: C.MUTED, valign: "top", margin: 0,
      fit: "shrink", shrinkText: true,
    }
  );

  addFooter(s, FOOTER);
  s.addNotes(composeNotes({
    say: [
      "Three examples - one for each joining word.",
      "BECAUSE example - I was tired because I had stayed up late. The staying up late is the REASON.",
      "WHEN example - the siren rang when it was time for lunch. Two things happening at the same TIME.",
      "IF example - we can play handball if it isn't raining. The not-raining is the CONDITION.",
      "Now the wrong example. Read it aloud - I was tired WHEN I had stayed up late. Does that say the same thing as the BECAUSE version? No - it says they happened at the same time, not that one caused the other.",
      "Today's job is to choose the right word for the meaning. Read your sentence aloud and listen.",
    ],
    do: [
      "Point at each example as you read it. Highlight the joining word.",
      "Read the wrong example with emphasis on WHEN to show how it changes the meaning.",
    ],
    teacherNotes: [
      "The think-aloud is 'read it aloud and listen for the meaning'. That's the strategy students need on the exit.",
      "Today supports criteria 2 (one grammatical sentence) and 3 (appropriate clause linking).",
    ],
    watchFor: [
      "Students looking at the words and trying to memorise rules instead of listening for the meaning - keep the prompt 'read it aloud and listen' visible.",
    ],
    tag: "[Day 4 | Thu | Teach MIX - choose by meaning]",
  }, { requireSay: false, requireDo: false }));
  runSlideDiagnostics(s, pres, { respectSafeBottom: false });
})();

/* ─────────────────────────────────────────────────────────────────────────
 *  Slide 3 + reveal: Guided practice - mixed subordinators, choose per pair
 * ───────────────────────────────────────────────────────────────────────── */

const mixedPairs = [
  {
    ideas: ["We had to wait outside the classroom.", "Our teacher was still setting up."],
    word: "BECAUSE",
    answer: "We had to wait outside the classroom BECAUSE our teacher was still setting up.",
    color: C.PRIMARY,
  },
  {
    ideas: ["I always wash my hands.", "I get back from lunch."],
    word: "WHEN",
    answer: "I always wash my hands WHEN I get back from lunch.",
    color: C.SECONDARY,
  },
];

withReveal(
  () => {
    const s = pres.addSlide();
    addTopBar(s, C.SECONDARY);
    badgeStack(s, "Practise together", C.SECONDARY, THU.dayText, THU.dayColor);
    addTitle(s, "Combine - choose BECAUSE, WHEN or IF");

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
    mixedPairs.forEach((pair, i) => {
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
      s.addText(`"${pair.ideas[0]}"`, {
        x: 0.75, y: y + 0.34, w: 8.4, h: 0.34,
        fontSize: 16, fontFace: FONT_H, color: C.CHARCOAL, italic: true, margin: 0,
        fit: "shrink", shrinkText: true,
      });
      s.addText(`"${pair.ideas[1]}"`, {
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
    s.addText("Whiteboards FIRST. Read the pair aloud, decide BECAUSE / WHEN / IF, then click to reveal.", {
      x: 0.75, y: stubY, w: 8.5, h: stubH,
      fontSize: 13, fontFace: FONT_B, color: C.MUTED, italic: true,
      align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(composeNotes({
      say: [
        "Two pairs. For each pair, decide BECAUSE, WHEN or IF, then write the combined sentence on your whiteboard.",
        "Read the pair aloud first. Listen for reason, time or condition.",
        "Try pair 1 - about 60 seconds. Then pair 2.",
      ],
      do: [
        "PROTOCOL - whiteboards FIRST, hold up, THEN click reveal. Do not click reveal until every student has attempted.",
        "Read pair 1 aloud, pause 60 seconds, scan.",
        "Then pair 2.",
        "Click reveal once everyone has attempted.",
      ],
      teacherNotes: [
        "Pair 1 takes BECAUSE - the teacher still setting up is the REASON we waited.",
        "Pair 2 takes WHEN - washing hands and getting back from lunch happen at the same TIME.",
        "If a student wrote 'we had to wait outside WHEN our teacher was still setting up' it's defensible but less precise. Mark as criterion 1 and 5 yes, criterion 3 lower - reteach the meaning distinction.",
      ],
      watchFor: [
        "Students who default to BECAUSE for both - prompt them to read each pair aloud and ask 'is this a reason or a time?'",
        "Comma splice - frame as developmental progress.",
      ],
      tag: "[Day 4 | Thu | Guided MIX - build]",
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
    const ansListY = stubY + 0.30;
    const perAnsH = (stubH - 0.36) / 2;
    mixedPairs.forEach((pair, i) => {
      const y = ansListY + i * perAnsH;
      const runs = highlightWord(pair.answer, pair.word, 14, C.WHITE, C.BG_CARD);
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
        "Pair 1 - BECAUSE. The teacher still setting up is the REASON we waited.",
        "Pair 2 - WHEN. Washing hands and getting back from lunch happen at the same TIME.",
        "Tick yours if your meaning matches.",
      ],
      do: [
        "Read both answers, highlight the joining word in each.",
        "Name the meaning for each - reason / time.",
        "Thumbs check before moving to independent.",
      ],
      tag: "[Day 4 | Thu | Guided MIX - reveal]",
    }, { requireSay: false, requireDo: false }));
  }
);

/* ─────────────────────────────────────────────────────────────────────────
 *  Slide 4 + reveal: Independent - student picks the subordinator
 * ───────────────────────────────────────────────────────────────────────── */

withReveal(
  () => {
    const s = pres.addSlide();
    addTopBar(s, C.ACCENT);
    badgeStack(s, "Your turn", C.ACCENT, THU.dayText, THU.dayColor);
    addTitle(s, "Write ONE sentence - choose the joining word");

    addTextOnShape(s, "In your warm-up book", {
      x: 6.85, y: 0.20, w: 2.65, h: 0.36, rectRadius: 0.06,
      fill: { color: C.WHITE },
      line: { color: C.ACCENT, width: 1.0 },
    }, {
      fontSize: 11, fontFace: FONT_B, color: C.ACCENT, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    const promptY = CONTENT_TOP;
    const promptH = 1.85;
    addCard(s, 0.5, promptY, 9, promptH, { strip: C.ACCENT, fill: C.WHITE });
    s.addText("Topic prompt", {
      x: 0.75, y: promptY + 0.12, w: 5, h: 0.28,
      fontSize: 12, fontFace: FONT_B, color: C.ACCENT, bold: true, margin: 0,
    });
    s.addText(
      "Write ONE sentence about something that happened today, from morning to home time. Use BECAUSE, WHEN or IF - whichever fits your meaning.",
      {
        x: 0.75, y: promptY + 0.46, w: 8.5, h: promptH - 0.60,
        fontSize: 19, fontFace: FONT_H, color: C.CHARCOAL, valign: "middle", margin: 0,
        fit: "shrink", shrinkText: true,
      }
    );

    // Reference chip strip - three options
    const refY = promptY + promptH + 0.14;
    const refH = 0.62;
    s.addShape("roundRect", {
      x: 0.5, y: refY, w: 9, h: refH, rectRadius: 0.08,
      fill: { color: C.BG_CARD },
      line: { color: C.SECONDARY, width: 0.8 },
    });
    s.addText("Choose:", {
      x: 0.75, y: refY, w: 1.0, h: refH,
      fontSize: 12, fontFace: FONT_B, color: C.SECONDARY, bold: true,
      valign: "middle", margin: 0,
    });
    const chips = [
      { word: "BECAUSE", color: C.PRIMARY },
      { word: "WHEN",    color: C.SECONDARY },
      { word: "IF",      color: C.ACCENT },
    ];
    const chipW = 2.10;
    const chipGap = 0.15;
    const chipsStart = 1.85;
    chips.forEach((c, i) => {
      const cx = chipsStart + i * (chipW + chipGap);
      addTextOnShape(s, c.word, {
        x: cx, y: refY + 0.08, w: chipW, h: refH - 0.16, rectRadius: 0.06,
        fill: { color: c.color },
      }, {
        fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
    });

    const remindY = refY + refH + 0.12;
    const remindH = SAFE_BOTTOM - remindY;
    s.addShape("roundRect", {
      x: 0.5, y: remindY, w: 9, h: remindH, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    });
    s.addText("ONE sentence, TWO ideas, joined by your chosen word. Read aloud to check the meaning.", {
      x: 0.75, y: remindY, w: 8.5, h: remindH,
      fontSize: 13.5, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(composeNotes({
      say: [
        "Your turn. Think of one thing that happened today - from morning to home time.",
        "Two short ideas, then choose your joining word. BECAUSE for a reason, WHEN for a time, IF for a condition.",
        "Write the sentence in your warm-up book. Four minutes.",
      ],
      do: [
        "Set a 4-minute timer.",
        "Circulate. For stuck students, ask 'what happened today? Then what?'",
        "Reveal the stem ONLY for individual focus students who are stuck.",
        "Note 1-2 strong sentences to share.",
      ],
      enabling: [
        "FOCUS GROUP: ask which joining word they want BEFORE they write. That way they only have to combine, not choose.",
        "Click reveal for the stem - point to it for the student, walk away.",
      ],
      extension: [
        "HIGHER-ABILITY EXTENSION: choose ONE and write under your main sentence:",
        "  (a) Write TWO different valid sentences about the same moment - one with BECAUSE and one with WHEN (or IF). Write a line saying which fits the meaning better.",
        "  (b) Add a WHERE detail to one of your clauses.",
        "  (c) Use a rarer subordinator - SINCE (like BECAUSE), AS / WHILE (like WHEN), UNLESS (the opposite of IF) - and check the meaning still works.",
      ],
      watchFor: [
        "Default to BECAUSE for everything - prompt them to read aloud and ask 'is this a reason, a time, or a condition?'",
        "Wrong joining word for the meaning (e.g. WHEN instead of BECAUSE) - cross criterion 3 and reteach the meaning distinction.",
        "Comma splice - frame as developmental progress.",
      ],
      tag: "[Day 4 | Thu | Independent MIX - build, stem hidden]",
    }, { requireSay: false, requireDo: false }));
    runSlideDiagnostics(s, pres, { respectSafeBottom: false });
    return s;
  },
  (s) => {
    // Reveal: add a sentence stem strip under the chip strip.
    const stemY = CONTENT_TOP + 1.85 + 0.14 + 0.62 + 0.12 + 0.78; // below remind strip - tight
    // Simpler: replace the reminder strip area with the stem? Actually we just
    // need to ADD the stem - put it below the remind strip if room, else
    // overlap. With remindY = CONTENT_TOP + 1.85 + 0.14 + 0.62 + 0.12 = ~ 3.13
    // and SAFE_BOTTOM = 5.1, remind strip height ~ 1.97. Plenty of room INSIDE
    // the remind strip for an inline stem - but that changes contrast. Instead,
    // shrink approach: don't add inside; render a separate stem strip on top
    // of the reminder area's bottom half.
    const stripY = SAFE_BOTTOM - 0.62;
    const stripH = 0.62;
    s.addShape("roundRect", {
      x: 0.5, y: stripY, w: 9, h: stripH, rectRadius: 0.08,
      fill: { color: C.WHITE },
      line: { color: C.SECONDARY, width: 1.0 },
    });
    s.addText("Need a start? Try: ", {
      x: 0.75, y: stripY, w: 2.0, h: stripH,
      fontSize: 12, fontFace: FONT_B, color: C.SECONDARY, bold: true,
      valign: "middle", margin: 0,
    });
    s.addText(`"Today I ___ BECAUSE / WHEN / IF ___."`, {
      x: 2.75, y: stripY, w: 6.55, h: stripH,
      fontSize: 14.5, fontFace: FONT_H, color: C.CHARCOAL, italic: true,
      valign: "middle", margin: 0, fit: "shrink",
    });
    s.addNotes(composeNotes({
      say: [
        "If you need a start, here's a frame. Pick ONE joining word.",
      ],
      do: [
        "Reveal for individual focus students only.",
      ],
      tag: "[Day 4 | Thu | Independent MIX - stem reveal]",
    }, { requireSay: false, requireDo: false }));
  }
);

/* ─────────────────────────────────────────────────────────────────────────
 *  Slide 5: Exit check - student chooses BECAUSE / WHEN / IF
 * ───────────────────────────────────────────────────────────────────────── */

(function exitMix() {
  const s = pres.addSlide();
  addTopBar(s, C.ALERT);
  badgeStack(s, "Exit check", C.ALERT, THU.dayText, THU.dayColor);
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

  // Left: task ideas
  addCard(s, 0.5, CONTENT_TOP, leftW, cardH, { strip: C.ALERT, fill: C.WHITE });
  s.addText("Today's exit check", {
    x: 0.7, y: CONTENT_TOP + 0.12, w: leftW - 0.4, h: 0.30,
    fontSize: 12, fontFace: FONT_B, color: C.ALERT, bold: true, margin: 0,
  });
  const ideaY = CONTENT_TOP + 0.46;
  s.addText('"I packed my bag for home."', {
    x: 0.7, y: ideaY, w: leftW - 0.4, h: 0.46,
    fontSize: 15.5, fontFace: FONT_H, color: C.CHARCOAL, italic: true, margin: 0,
    fit: "shrink", shrinkText: true,
  });
  s.addText('"The bell went."', {
    x: 0.7, y: ideaY + 0.52, w: leftW - 0.4, h: 0.46,
    fontSize: 15.5, fontFace: FONT_H, color: C.CHARCOAL, italic: true, margin: 0,
    fit: "shrink", shrinkText: true,
  });
  s.addText("Combine into ONE sentence. Choose BECAUSE, WHEN or IF.", {
    x: 0.7, y: ideaY + 1.10, w: leftW - 0.4, h: 0.46,
    fontSize: 12.5, fontFace: FONT_B, color: C.SECONDARY, bold: true, margin: 0,
    fit: "shrink", shrinkText: true,
  });
  s.addText("Submit via QR (or paper slip).", {
    x: 0.7, y: CONTENT_TOP + cardH - 0.36, w: leftW - 0.4, h: 0.26,
    fontSize: 10.5, fontFace: FONT_B, color: C.MUTED, italic: true, margin: 0,
  });

  // Middle: QR
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
    fill: { color: C.WHITE }, line: { color: C.CHARCOAL, width: 1.0 },
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
  s.addText("Thu MIX form", {
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
  MARK_THU.forEach((crit, i) => {
    const y = rowStart + i * (rowH + 0.08);
    s.addShape("roundRect", {
      x: rightX + 0.16, y: y + 0.10, w: 0.30, h: 0.30, rectRadius: 0.04,
      fill: { color: C.WHITE }, line: { color: C.SUCCESS, width: 1.2 },
    });
    s.addText(crit, {
      x: rightX + 0.52, y, w: rightW - 0.66, h: rowH,
      fontSize: 11, fontFace: FONT_B, color: C.CHARCOAL, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });
  });
  s.addText("Note which conjunction was chosen.", {
    x: rightX + 0.16, y: CONTENT_TOP + cardH - 0.40, w: rightW - 0.32, h: 0.30,
    fontSize: 10, fontFace: FONT_B, color: C.MUTED, italic: true, margin: 0,
    fit: "shrink",
  });

  addFooter(s, FOOTER);
  s.addNotes(composeNotes({
    say: [
      "Last task. Combine these two ideas - 'I packed my bag for home.' and 'The bell went.' - into ONE sentence.",
      "Choose BECAUSE, WHEN or IF. Read aloud to check the meaning.",
      "Type into the Google Form on your iPad. Two minutes. Paper slip if iPad isn't working.",
    ],
    do: [
      "Timer 2 minutes.",
      "Use the marking criteria. Open the Google Form sheet view.",
      "Note the joining word each student chose on the cohort tracker.",
    ],
    teacherNotes: [
      "Defensible answers: 'I packed my bag for home WHEN the bell went.' is the strongest meaning match (time). 'I packed my bag for home BECAUSE the bell went' is also defensible (reason / cue) - mark criterion 3 as yes.",
      "'I packed my bag for home IF the bell went' is NOT defensible - IF makes it conditional, which doesn't fit. Cross criterion 3.",
      "Today's exit primarily checks criterion 2 (one grammatical sentence), 3 (clause linking) and 5 (no comma splice).",
    ],
    watchFor: [
      "IF chosen here - cross criterion 3, reteach the condition meaning.",
      "Comma splice - cross criterion 5, frame as progress.",
      "Two separate sentences - cross criterion 2.",
    ],
    tag: "[Day 4 | Thu | Exit MIX - QR submit]",
  }, { requireSay: false, requireDo: false }));
  runSlideDiagnostics(s, pres, { respectSafeBottom: false });
})();

const outFile = path.join(OUT_DIR, "SC_T2W5_Lesson4_Thu_MIX.pptx");
pres.writeFile({ fileName: outFile }).then(() => {
  console.log("PPTX written to " + outFile);
}).catch((err) => {
  console.error("Build failed:", err);
  process.exitCode = 1;
});
