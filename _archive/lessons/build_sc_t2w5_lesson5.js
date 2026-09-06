"use strict";

/**
 * Sentence Combining Warm-Up - Year 5/6 Enrichment, Term 2 Week 5, Day 5 (Fri)
 * Contrast subordinating (BECAUSE / WHEN / IF) with coordinating (AND / BUT
 * / SO / OR). Students choose from all seven.
 * Topic: weekend plans / what they're looking forward to.
 *
 * Friday includes the unit's closing reflection slide at the end so the
 * combined deck wraps up cleanly.
 *
 * Scoring rule (in speaker notes for the exit slide): defensible joining
 * word for the meaning. Multiple answers count.
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

const FOOTER = "Sentence Combining Warm-Up | Year 5/6 | Term 2 Week 5 | Day 5 Fri";
const OUT_DIR = "output/SC_T2W5_Lesson5_Fri_MIX_COORD_SUBORD";
fs.mkdirSync(OUT_DIR, { recursive: true });

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Sentence Combining Warm-Up - Year 5/6 - Term 2 Week 5 Day 5 (Fri)";
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

const FRI = { dayText: "Fri", dayColor: C.SUCCESS };

const SC_FRI = [
  "I can find two short ideas that go together.",
  "I can choose a coordinating OR subordinating word that fits the meaning.",
  "I can write a sentence that makes sense.",
];

const MARK_FRI = [
  "Two ideas combined into ONE sentence",
  "Chose a defensible joining word for the meaning",
  "Sentence is grammatical (no run-on, no comma splice)",
];

/* ─────────────────────────────────────────────────────────────────────────
 *  Slide 1: Day intro - review both groups
 * ───────────────────────────────────────────────────────────────────────── */

(function dayIntro() {
  const s = pres.addSlide();
  addTopBar(s, C.PRIMARY);
  badgeStack(s, "Day start", C.PRIMARY, FRI.dayText, FRI.dayColor);
  addTitle(s, "Fri - Coordinating or subordinating?");

  const cardH = SAFE_BOTTOM - CONTENT_TOP;
  const leftW = 4.6;
  const rightW = 4.4;
  const rightX = 5.2;

  addCard(s, 0.5, CONTENT_TOP, leftW, cardH, { strip: C.SECONDARY, fill: C.WHITE });
  s.addText("Two groups of joining words", {
    x: 0.7, y: CONTENT_TOP + 0.14, w: leftW - 0.4, h: 0.34,
    fontSize: 13, fontFace: FONT_B, color: C.SECONDARY, bold: true, margin: 0,
  });

  // Coordinating row
  const coordY = CONTENT_TOP + 0.56;
  s.addText("Coordinating (Week 4)", {
    x: 0.7, y: coordY, w: leftW - 0.4, h: 0.30,
    fontSize: 12, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
  });
  const coordChips = ["AND", "BUT", "SO", "OR"];
  const coordChipW = 0.85;
  const coordChipGap = 0.10;
  coordChips.forEach((w, i) => {
    const cx = 0.7 + i * (coordChipW + coordChipGap);
    addTextOnShape(s, w, {
      x: cx, y: coordY + 0.32, w: coordChipW, h: 0.42, rectRadius: 0.06,
      fill: { color: C.PRIMARY },
    }, {
      fontSize: 13, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
  });
  s.addText("Join two equal ideas - the meaning sits side by side.", {
    x: 0.7, y: coordY + 0.80, w: leftW - 0.4, h: 0.30,
    fontSize: 11.5, fontFace: FONT_B, color: C.MUTED, italic: true, margin: 0,
    fit: "shrink",
  });

  // Subordinating row
  const subY = coordY + 1.20;
  s.addText("Subordinating (Week 5)", {
    x: 0.7, y: subY, w: leftW - 0.4, h: 0.30,
    fontSize: 12, fontFace: FONT_B, color: C.ACCENT, bold: true, margin: 0,
  });
  const subChips = ["BECAUSE", "WHEN", "IF"];
  const subChipW = 1.20;
  subChips.forEach((w, i) => {
    const cx = 0.7 + i * (subChipW + coordChipGap);
    addTextOnShape(s, w, {
      x: cx, y: subY + 0.32, w: subChipW, h: 0.42, rectRadius: 0.06,
      fill: { color: C.ACCENT },
    }, {
      fontSize: 13, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
  });
  s.addText("Make one idea SUPPORT the other - reason, time or condition.", {
    x: 0.7, y: subY + 0.80, w: leftW - 0.4, h: 0.30,
    fontSize: 11.5, fontFace: FONT_B, color: C.MUTED, italic: true, margin: 0,
    fit: "shrink",
  });

  addCard(s, rightX, CONTENT_TOP, rightW, cardH, { strip: C.ACCENT, fill: C.BG_CARD });
  s.addText("Today's success criteria - I can...", {
    x: rightX + 0.2, y: CONTENT_TOP + 0.14, w: rightW - 0.4, h: 0.34,
    fontSize: 13, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
  });
  s.addText(SC_FRI.map((t, i) => ({
    text: t,
    options: {
      bullet: true,
      breakLine: i < SC_FRI.length - 1,
      fontSize: 14.5, color: C.CHARCOAL, paraSpaceAfter: 6,
    },
  })), {
    x: rightX + 0.2, y: CONTENT_TOP + 0.54, w: rightW - 0.4, h: cardH - 0.68,
    fontFace: FONT_B, valign: "top", margin: 0,
    fit: "shrink", shrinkText: true,
  });

  addFooter(s, FOOTER);
  s.addNotes(composeNotes({
    say: [
      "Last day this week. Today we put both groups together.",
      "Coordinating - AND, BUT, SO, OR - join two ideas that sit side by side. Equal weight.",
      "Subordinating - BECAUSE, WHEN, IF - make one idea SUPPORT the other. Reason, time or condition.",
      "Today's job is to pick the word that fits your meaning. Sometimes more than one word works.",
      "Topic - weekend plans, or one thing you are looking forward to.",
    ],
    do: [
      "Point at each group as you describe it.",
      "Quick choral check - 'AND, BUT, SO, OR are...', 'BECAUSE, WHEN, IF are...'",
    ],
    teacherNotes: [
      "Friday is the integration day for the cycle so far. By the end of today, students have practised 7 joining words.",
      "Today supports rubric criteria 2 (one grammatical sentence), 3 (appropriate clause linking), 4 (no unnecessary repetition) and 5 (no comma splice).",
    ],
    enabling: [
      "FOCUS GROUP: keep this two-group reference up as you teach. Walk the focus group through the two groups physically (point to each chip) before guided practice.",
      "EXTENSION: ask higher-ability students which rarer words belong in each group - AND/BUT/SO/OR plus FOR/YET/NOR (coord); BECAUSE/WHEN/IF plus SINCE/AS/WHILE/AFTER/BEFORE/UNTIL/ALTHOUGH/UNLESS (subord).",
    ],
    watchFor: [
      "Students defaulting to AND for everything - prompt them to read aloud and check for reason, time, condition or contrast.",
    ],
    tag: "[Day 5 | Fri | Review both groups + LI/SC]",
  }, { requireSay: false, requireDo: false }));
  runSlideDiagnostics(s, pres, { respectSafeBottom: false });
})();

/* ─────────────────────────────────────────────────────────────────────────
 *  Slide 2: Teach - contrast meanings with two examples
 * ───────────────────────────────────────────────────────────────────────── */

(function teachContrast() {
  const s = pres.addSlide();
  addTopBar(s, C.PRIMARY);
  badgeStack(s, "Teach", C.PRIMARY, FRI.dayText, FRI.dayColor);
  addTitle(s, "Same two ideas - different joining words");

  // Two side-by-side cards comparing coord and subord
  const examples = [
    {
      label: "Coordinating",
      sub: "AND - two ideas side by side",
      sentence: "I am going to the pool on Saturday AND my cousins are coming over on Sunday.",
      word: "AND",
      color: C.PRIMARY,
    },
    {
      label: "Subordinating",
      sub: "BECAUSE - one idea is the REASON",
      sentence: "I am going to the pool on Saturday BECAUSE the weather is going to be hot.",
      word: "BECAUSE",
      color: C.ACCENT,
    },
  ];

  const topY = CONTENT_TOP;
  const topH = 1.85;
  const colW = 4.45;
  const gap = 0.10;
  examples.forEach((ex, i) => {
    const x = 0.5 + i * (colW + gap);
    addCard(s, x, topY, colW, topH, { strip: ex.color, fill: C.WHITE });
    s.addText(ex.label, {
      x: x + 0.16, y: topY + 0.10, w: colW - 0.32, h: 0.28,
      fontSize: 12.5, fontFace: FONT_B, color: ex.color, bold: true,
      align: "center", margin: 0,
    });
    s.addText(ex.sub, {
      x: x + 0.16, y: topY + 0.40, w: colW - 0.32, h: 0.24,
      fontSize: 11, fontFace: FONT_B, color: C.MUTED, italic: true,
      align: "center", margin: 0,
    });
    const runs = highlightWord(ex.sentence, ex.word, 13, C.CHARCOAL, ex.color);
    s.addText(runs.map((run, j) => ({
      text: `${j === 0 ? '"' : ""}${run.text}${j === runs.length - 1 ? '"' : ""}`,
      options: { ...run.options, fontFace: FONT_H, breakLine: false },
    })), {
      x: x + 0.16, y: topY + 0.70, w: colW - 0.32, h: topH - 0.80,
      italic: true, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });
  });

  // Middle strip: "How to choose" rule
  const ruleY = topY + topH + 0.12;
  const ruleH = 0.62;
  s.addShape("roundRect", {
    x: 0.5, y: ruleY, w: 9, h: ruleH, rectRadius: 0.06,
    fill: { color: C.BG_CARD },
    line: { color: C.SECONDARY, width: 0.8 },
  });
  s.addText("How to choose:", {
    x: 0.75, y: ruleY + 0.04, w: 1.6, h: ruleH - 0.08,
    fontSize: 11.5, fontFace: FONT_B, color: C.SECONDARY, bold: true,
    valign: "middle", margin: 0,
  });
  s.addText(
    "Two equal ideas -> coordinating (AND, BUT, SO, OR). One supports the other (reason, time, condition) -> subordinating (BECAUSE, WHEN, IF).",
    {
      x: 2.35, y: ruleY + 0.04, w: 6.95, h: ruleH - 0.08,
      fontSize: 11.5, fontFace: FONT_B, color: C.CHARCOAL, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    },
  );

  // Non-example: comma splice that should have used SO or BECAUSE
  const neY = ruleY + ruleH + 0.14;
  const neH = SAFE_BOTTOM - neY;
  addCard(s, 0.5, neY, 9, neH, { strip: C.ALERT, fill: C.WHITE });
  s.addText("Watch out - comma alone is not a joining word", {
    x: 0.75, y: neY + 0.06, w: 7, h: 0.26,
    fontSize: 12, fontFace: FONT_B, color: C.ALERT, bold: true, margin: 0,
  });
  s.addText(
    `Wrong: "I am going to the pool on Saturday, the weather is going to be hot."`,
    {
      x: 0.75, y: neY + 0.32, w: 8.5, h: 0.30,
      fontSize: 13.5, fontFace: FONT_H, color: C.CHARCOAL, italic: true, margin: 0,
      fit: "shrink", shrinkText: true,
    },
  );
  s.addText(
    "Two full ideas joined by only a comma. Fix with SO or BECAUSE: 'going to the pool SO the weather will be hot' is wrong meaning; 'going to the pool BECAUSE the weather is going to be hot' is the meaning. Comma alone never joins two full ideas.",
    {
      x: 0.75, y: neY + 0.62, w: 8.5, h: neH - 0.72,
      fontSize: 11, fontFace: FONT_B, color: C.MUTED, valign: "top", margin: 0,
      fit: "shrink", shrinkText: true,
    },
  );

  addFooter(s, FOOTER);
  s.addNotes(composeNotes({
    say: [
      "Same kind of weekend ideas, two different joining words.",
      "Left - AND. Just two facts about my weekend, side by side. No reason, no time, no condition - they just go together.",
      "Right - BECAUSE. Now the second idea is the REASON for the first. The hot weather is why I'm going to the pool.",
      "How to choose? Read the two ideas. If they are equal, use a coordinating word - AND, BUT, SO, OR. If one supports the other - reason, time or condition - use a subordinating word - BECAUSE, WHEN, IF.",
      "Now the wrong example. Two full ideas joined by only a comma - a comma splice. The fix is a joining word that matches the meaning. SO would say 'the pool caused the hot weather' - wrong. BECAUSE works - 'I am going to the pool BECAUSE the weather is going to be hot.'",
    ],
    do: [
      "Read the left sentence with neutral stress.",
      "Read the right sentence with a small pause before BECAUSE to make the reason audible.",
      "Point at the 'how to choose' rule as you say it.",
      "Read the wrong example aloud, then read the BECAUSE fix.",
    ],
    teacherNotes: [
      "Both worked examples are grammatically correct. The choice is meaning, not rule.",
      "Today supports criterion 3 (clause linking), criterion 4 (no unnecessary repetition - subordinating words let you tighten) and criterion 5 (no comma splice).",
      "Today is the first day students retrieve from the FULL set of seven joining words. The pre-reading (page 8) calls this interleaving and recommends it from Week 4 onwards. We sequenced (coord -> subord -> mix) for cognitive load reasons; Friday is the interleaving day before Week 6's full Q3 task.",
    ],
    watchFor: [
      "Students treating coordinating and subordinating as interchangeable - keep the meaning question front and centre.",
      "Students who pick SO because cause-and-effect 'sounds right' here - read aloud and ask which event caused which.",
    ],
    tag: "[Day 5 | Fri | Teach contrast + non-example]",
  }, { requireSay: false, requireDo: false }));
  runSlideDiagnostics(s, pres, { respectSafeBottom: false });
})();

/* ─────────────────────────────────────────────────────────────────────────
 *  Slide 3 + reveal: Guided practice - mixed coord/subord
 * ───────────────────────────────────────────────────────────────────────── */

const friPairs = [
  {
    ideas: ["I might go to the skate park.", "It depends on the weather."],
    answer: "I might go to the skate park IF the weather is good.",
    bestWord: "IF",
    color: C.ACCENT,
    note: "Best fit: IF (condition). 'SO it depends on the weather' is grammatical but less precise.",
  },
  {
    ideas: ["My nan is making a roast on Sunday.", "She makes the best gravy."],
    answer: "My nan is making a roast on Sunday AND she makes the best gravy.",
    bestWord: "AND",
    color: C.PRIMARY,
    note: "Best fit: AND (two equal ideas about my nan's roast). BECAUSE could fit but slightly changes meaning - 'AND' is cleaner here.",
  },
];

withReveal(
  () => {
    const s = pres.addSlide();
    addTopBar(s, C.SECONDARY);
    badgeStack(s, "Practise together", C.SECONDARY, FRI.dayText, FRI.dayColor);
    addTitle(s, "Combine - choose any of the seven words");

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
    friPairs.forEach((pair, i) => {
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
    s.addText("Whiteboards FIRST. Pick from AND, BUT, SO, OR, BECAUSE, WHEN, IF. Then click reveal.", {
      x: 0.75, y: stubY, w: 8.5, h: stubH,
      fontSize: 12.5, fontFace: FONT_B, color: C.MUTED, italic: true,
      align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(composeNotes({
      say: [
        "Two pairs. Pick any of the seven joining words from this week and last week.",
        "Read the pair aloud. Listen for the meaning. Pick the word that fits.",
        "60 seconds for pair 1. Then pair 2.",
      ],
      do: [
        "PROTOCOL - whiteboards FIRST, hold up, THEN click reveal.",
        "Read pair 1 aloud, pause 60 seconds, scan.",
        "Then pair 2.",
        "Click reveal once everyone has attempted.",
      ],
      teacherNotes: [
        "Pair 1 best fit is IF (the weather is the condition). 'SO it depends on the weather' is grammatical but stretches the meaning. Acceptable.",
        "Pair 2 best fit is AND. Two equal ideas about my nan's roast. BECAUSE could work but slightly twists the meaning.",
        "Today's pairs support criterion 3 (clause linking) and 4 (no unnecessary repetition).",
        "If a student picked a defensible alternative, that's a tick for criteria 1, 2 and 5 - even if criterion 3 is borderline.",
      ],
      watchFor: [
        "Default AND for both pairs - prompt them to listen for whether the second idea supports the first.",
        "Comma splice - frame as developmental progress.",
      ],
      tag: "[Day 5 | Fri | Guided contrast - build]",
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
    s.addText("Answers (best fit)", {
      x: 0.75, y: stubY + 0.06, w: 4, h: 0.24,
      fontSize: 11, fontFace: FONT_B, color: C.WHITE, bold: true, margin: 0,
    });
    const ansListY = stubY + 0.30;
    const perAnsH = (stubH - 0.36) / 2;
    friPairs.forEach((pair, i) => {
      const y = ansListY + i * perAnsH;
      const runs = highlightWord(pair.answer, pair.bestWord, 14, C.WHITE, C.BG_CARD);
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
        "Pair 1 - IF. The weather is the CONDITION. SO is also defensible but a bit less precise.",
        "Pair 2 - AND. Two equal ideas about the roast. The gravy is part of the same nice-weekend picture.",
        "Tick yours if your joining word fits the meaning.",
      ],
      do: [
        "Read both answers, highlight the chosen word.",
        "Name the meaning - condition / side-by-side.",
        "Thumbs check before moving to independent.",
      ],
      tag: "[Day 5 | Fri | Guided contrast - reveal]",
    }, { requireSay: false, requireDo: false }));
  }
);

/* ─────────────────────────────────────────────────────────────────────────
 *  Slide 4 + reveal: Independent - any of the seven
 * ───────────────────────────────────────────────────────────────────────── */

withReveal(
  () => {
    const s = pres.addSlide();
    addTopBar(s, C.ACCENT);
    badgeStack(s, "Your turn", C.ACCENT, FRI.dayText, FRI.dayColor);
    addTitle(s, "Write ONE sentence - any of the seven");

    addTextOnShape(s, "In your warm-up book", {
      x: 6.85, y: 0.20, w: 2.65, h: 0.36, rectRadius: 0.06,
      fill: { color: C.WHITE },
      line: { color: C.ACCENT, width: 1.0 },
    }, {
      fontSize: 11, fontFace: FONT_B, color: C.ACCENT, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    const promptY = CONTENT_TOP;
    const promptH = 1.75;
    addCard(s, 0.5, promptY, 9, promptH, { strip: C.ACCENT, fill: C.WHITE });
    s.addText("Topic prompt", {
      x: 0.75, y: promptY + 0.12, w: 5, h: 0.28,
      fontSize: 12, fontFace: FONT_B, color: C.ACCENT, bold: true, margin: 0,
    });
    s.addText(
      "Write ONE sentence about your weekend plans or something you are looking forward to. Choose the joining word that fits your meaning.",
      {
        x: 0.75, y: promptY + 0.46, w: 8.5, h: promptH - 0.60,
        fontSize: 18, fontFace: FONT_H, color: C.CHARCOAL, valign: "middle", margin: 0,
        fit: "shrink", shrinkText: true,
      }
    );

    // Reference chip strip - all seven options
    const refY = promptY + promptH + 0.14;
    const refH = 0.62;
    s.addShape("roundRect", {
      x: 0.5, y: refY, w: 9, h: refH, rectRadius: 0.08,
      fill: { color: C.BG_CARD },
      line: { color: C.SECONDARY, width: 0.8 },
    });
    s.addText("Choose:", {
      x: 0.65, y: refY, w: 0.9, h: refH,
      fontSize: 11.5, fontFace: FONT_B, color: C.SECONDARY, bold: true,
      valign: "middle", margin: 0,
    });
    const chips = [
      { word: "AND", color: C.PRIMARY },
      { word: "BUT", color: C.PRIMARY },
      { word: "SO",  color: C.PRIMARY },
      { word: "OR",  color: C.PRIMARY },
      { word: "BECAUSE", color: C.ACCENT },
      { word: "WHEN", color: C.ACCENT },
      { word: "IF",  color: C.ACCENT },
    ];
    const chipsStart = 1.55;
    const chipsEnd = 9.30;
    const chipGap = 0.08;
    const totalChipsW = chipsEnd - chipsStart - chipGap * (chips.length - 1);
    const chipW = totalChipsW / chips.length;
    chips.forEach((c, i) => {
      const cx = chipsStart + i * (chipW + chipGap);
      addTextOnShape(s, c.word, {
        x: cx, y: refY + 0.08, w: chipW, h: refH - 0.16, rectRadius: 0.06,
        fill: { color: c.color },
      }, {
        fontSize: 11.5, fontFace: FONT_H, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
    });

    const remindY = refY + refH + 0.12;
    const remindH = SAFE_BOTTOM - remindY;
    s.addShape("roundRect", {
      x: 0.5, y: remindY, w: 9, h: remindH, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    });
    s.addText(
      "Coordinating (left): two ideas side by side.  |  Subordinating (right): one idea SUPPORTS the other.",
      {
        x: 0.75, y: remindY + 0.04, w: 8.5, h: 0.40,
        fontSize: 11.5, fontFace: FONT_B, color: C.WHITE, italic: true,
        align: "center", valign: "top", margin: 0,
      },
    );
    s.addText("ONE sentence, TWO ideas, joined by your chosen word. Read aloud to check the meaning.", {
      x: 0.75, y: remindY + 0.42, w: 8.5, h: remindH - 0.50,
      fontSize: 13.5, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(composeNotes({
      say: [
        "Your turn. Think of one weekend plan or one thing you're looking forward to.",
        "Two short ideas, then choose the joining word that fits your meaning. Any of the seven.",
        "Write the sentence in your warm-up book. Four minutes.",
      ],
      do: [
        "Timer 4 minutes.",
        "Circulate. For stuck students, ask 'what are you looking forward to? Why? When?'",
        "Reveal the stem ONLY for individual focus students who are stuck.",
        "Note 1-2 strong sentences to share - same workload as other days.",
        "As you circulate, observe whether focus students are reaching for a coordinating or a subordinating word. You do not need to record this - just hold a rough sense of the split. It is useful Week 6 planning data, and observing isn't the same as recording.",
      ],
      enabling: [
        "FOCUS GROUP: ask which group of words they want to use BEFORE they write. That cuts the choice from 7 down to 3 or 4.",
        "Click reveal for the stem - point to it for the student, walk away.",
      ],
      extension: [
        "HIGHER-ABILITY EXTENSION: choose ONE and write under your main sentence:",
        "  (a) Write TWO versions of your sentence using DIFFERENT joining words (e.g. one with AND, one with BECAUSE). Write a line saying which is clearer or sounds better and why.",
        "  (b) Add a WHEN or WHERE detail to one of your clauses.",
        "  (c) Use a rarer joining word from either group - SINCE, AS, WHILE, ALTHOUGH, UNLESS, BEFORE, AFTER - and check the meaning still works.",
      ],
      watchFor: [
        "Default AND for everything - prompt them to listen for reason, time, condition or contrast.",
        "Comma splice - mark as developmental progress.",
        "Two separate sentences - redirect to ONE sentence joined by their chosen word.",
      ],
      tag: "[Day 5 | Fri | Independent contrast - build, stem hidden]",
    }, { requireSay: false, requireDo: false }));
    runSlideDiagnostics(s, pres, { respectSafeBottom: false });
    return s;
  },
  (s) => {
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
    s.addText(`"This weekend I am ___ ___ ___."   (fill the middle with your joining word)`, {
      x: 2.75, y: stripY, w: 6.55, h: stripH,
      fontSize: 13, fontFace: FONT_H, color: C.CHARCOAL, italic: true,
      valign: "middle", margin: 0, fit: "shrink",
    });
    s.addNotes(composeNotes({
      say: [
        "If you need a start, here is a frame. Pick ONE joining word for the middle.",
      ],
      do: [
        "Reveal for individual focus students only.",
      ],
      tag: "[Day 5 | Fri | Independent contrast - stem reveal]",
    }, { requireSay: false, requireDo: false }));
  }
);

/* ─────────────────────────────────────────────────────────────────────────
 *  Slide 5: Exit check - student picks any of the seven
 * ───────────────────────────────────────────────────────────────────────── */

(function exitFri() {
  const s = pres.addSlide();
  addTopBar(s, C.ALERT);
  badgeStack(s, "Exit check", C.ALERT, FRI.dayText, FRI.dayColor);
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
  s.addText("Today's exit check", {
    x: 0.7, y: CONTENT_TOP + 0.12, w: leftW - 0.4, h: 0.30,
    fontSize: 12, fontFace: FONT_B, color: C.ALERT, bold: true, margin: 0,
  });
  const ideaY = CONTENT_TOP + 0.46;
  s.addText('"The match was tied at full time."', {
    x: 0.7, y: ideaY, w: leftW - 0.4, h: 0.60,
    fontSize: 14, fontFace: FONT_H, color: C.CHARCOAL, italic: true, margin: 0,
    fit: "shrink", shrinkText: true,
  });
  s.addText('"The crowd was on the edge of their seats."', {
    x: 0.7, y: ideaY + 0.64, w: leftW - 0.4, h: 0.60,
    fontSize: 14, fontFace: FONT_H, color: C.CHARCOAL, italic: true, margin: 0,
    fit: "shrink", shrinkText: true,
  });
  s.addText("Combine into ONE sentence. Choose the joining word that fits.", {
    x: 0.7, y: ideaY + 1.30, w: leftW - 0.4, h: 0.46,
    fontSize: 11.5, fontFace: FONT_B, color: C.SECONDARY, bold: true, margin: 0,
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
  s.addText("Fri MIX form", {
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
  MARK_FRI.forEach((crit, i) => {
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
  s.addText("Defensible answers tick.", {
    x: rightX + 0.16, y: CONTENT_TOP + cardH - 0.40, w: rightW - 0.32, h: 0.30,
    fontSize: 10, fontFace: FONT_B, color: C.MUTED, italic: true, margin: 0,
    fit: "shrink",
  });

  addFooter(s, FOOTER);
  s.addNotes(composeNotes({
    say: [
      "Last exit of the week. Combine these two ideas into ONE sentence using ANY of the seven joining words.",
      "Read aloud to check the meaning. Pick the word that fits.",
      "Type into the Google Form on your iPad. Two minutes. Paper slip if iPad isn't working.",
    ],
    do: [
      "Timer 2 minutes.",
      "Use the marking criteria. Watch responses populate live in the Google Form sheet view.",
      "Tick on the cohort tracker.",
    ],
    teacherNotes: [
      "SCORING RULE - all 5 teachers score the same way to keep the cohort tracker valid.",
      "Tick if the sentence is ONE grammatical sentence with a defensible joining word for the meaning. Mark on meaning-match and grammar, NOT on picking the 'right' answer.",
      "Defensible for THIS pair: SO (cause-effect - the tied score caused the edge-of-seats crowd) is the strongest. AND (sequential / side by side) is also defensible. BECAUSE is also acceptable if the student flips the order - 'The crowd was on the edge of their seats BECAUSE the match was tied at full time.' All three tick.",
      "Not defensible for this pair: OR (no choice), IF (no condition), WHEN (the times are the same so it's wooden but not wrong - mark soft). BUT (contrast) - if used the meaning gets weird; mark soft.",
      "Don't penalise a student for choosing AND over SO if the meaning still works. Sarah ticks, Stacey ticks, Bek ticks - same criteria, same result.",
      "FRAGMENTS vs COMMA SPLICES - students don't need to know that subordinating conjunctions create dependent clauses, but you do. If a focus student writes a fragment like 'Because the match was tied at full time.' as a complete sentence, the correction is 'attach it to a main clause' - DO NOT mark it as a comma splice. A comma splice is two full ideas joined by only a comma; a fragment is one half of a sentence on its own.",
      "INTERLEAVING - today is the first day students retrieve from the full 7-word set. The pre-reading (page 8) calls this interleaving and recommends it from Week 4 onwards. Today's exit data tells us whether students can switch between coordinating and subordinating words on demand. That answer drives Week 6 planning.",
      "Today's exit primarily checks criterion 2 (one grammatical sentence), 3 (clause linking) and 5 (no comma splice).",
    ],
    watchFor: [
      "Comma splice - cross criterion 5, frame as developmental progress.",
      "Two separate sentences - cross criterion 1.",
      "Fragment (e.g. 'Because the match was tied at full time.') - cross criteria 1 and 2; correction is to attach the main clause, NOT to call it a comma splice.",
      "Joining word that doesn't fit the meaning (OR, IF) - cross criterion 2, but tick 1 and 3 if grammar is sound.",
    ],
    tag: "[Day 5 | Fri | Exit contrast - QR submit + scoring rule]",
  }, { requireSay: false, requireDo: false }));
  runSlideDiagnostics(s, pres, { respectSafeBottom: false });
})();

/* ─────────────────────────────────────────────────────────────────────────
 *  Closing reflection - end of the week
 * ───────────────────────────────────────────────────────────────────────── */

(function closingReflection() {
  const s = pres.addSlide();
  addTopBar(s, C.SUCCESS);
  addBadge(s, "Closing reflection", { color: C.SUCCESS, w: 2.4 });
  addTitle(s, "Look back at this week", { w: 7.4 });

  const cardH = SAFE_BOTTOM - CONTENT_TOP;
  const leftW = 4.6;
  const rightW = 4.4;
  const rightX = 5.2;

  // Left: the three week-5 SCs as a single I can list (no tier labels)
  addCard(s, 0.5, CONTENT_TOP, leftW, cardH, { strip: C.SUCCESS, fill: C.WHITE });
  s.addText("By the end of this week - I can...", {
    x: 0.7, y: CONTENT_TOP + 0.14, w: leftW - 0.4, h: 0.32,
    fontSize: 13, fontFace: FONT_B, color: C.SUCCESS, bold: true, margin: 0,
  });
  const weekSc = [
    "I can find two short ideas that go together.",
    "I can use BECAUSE, WHEN or IF to join them into ONE sentence.",
    "I can write a sentence that makes sense.",
  ];
  s.addText(weekSc.map((t, i) => ({
    text: t,
    options: {
      bullet: true,
      breakLine: i < weekSc.length - 1,
      fontSize: 15, color: C.CHARCOAL, paraSpaceAfter: 8,
    },
  })), {
    x: 0.7, y: CONTENT_TOP + 0.54, w: leftW - 0.4, h: cardH - 0.68,
    fontFace: FONT_B, valign: "top", margin: 0,
    fit: "shrink", shrinkText: true,
  });

  // Right: self-check prompt
  addCard(s, rightX, CONTENT_TOP, rightW, cardH, { strip: C.ACCENT, fill: C.BG_CARD });
  s.addText("Thumb check on each line", {
    x: rightX + 0.2, y: CONTENT_TOP + 0.14, w: rightW - 0.4, h: 0.32,
    fontSize: 13, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
  });
  const thumbs = [
    "Thumbs up - I can do this on my own.",
    "Thumbs sideways - I can do this with help.",
    "Thumbs down - I still need practice.",
  ];
  s.addText(thumbs.map((t, i) => ({
    text: t,
    options: {
      bullet: true,
      breakLine: i < thumbs.length - 1,
      fontSize: 13.5, color: C.CHARCOAL, paraSpaceAfter: 6,
    },
  })), {
    x: rightX + 0.2, y: CONTENT_TOP + 0.54, w: rightW - 0.4, h: 1.20,
    fontFace: FONT_B, valign: "top", margin: 0,
    fit: "shrink", shrinkText: true,
  });

  const askY = CONTENT_TOP + 1.84;
  s.addText("Quick share", {
    x: rightX + 0.2, y: askY, w: rightW - 0.4, h: 0.28,
    fontSize: 12, fontFace: FONT_B, color: C.SECONDARY, bold: true, margin: 0,
  });
  s.addText(
    "One thing you got better at this week. One thing to keep practising next week.",
    {
      x: rightX + 0.2, y: askY + 0.32, w: rightW - 0.4, h: cardH - askY + CONTENT_TOP - 0.40,
      fontSize: 13, fontFace: FONT_H, color: C.CHARCOAL, italic: true, valign: "top", margin: 0,
      fit: "shrink", shrinkText: true,
    },
  );

  addFooter(s, FOOTER);
  s.addNotes(composeNotes({
    say: [
      "Quick reflection. Read the three I-can statements with me.",
      "Now thumb-check each one. Thumbs up if you can do it on your own. Sideways if you can do it with help. Down if you still need practice.",
      "One quick share - tell your partner one thing you got better at this week, and one thing to keep practising next week.",
    ],
    do: [
      "Read each I-can statement aloud, pause, scan thumbs.",
      "Note which criterion the class is least confident on - that's the focus for early next week.",
      "Two-minute partner share, then pack up.",
    ],
    teacherNotes: [
      "This slide closes the week. Use the thumb data to plan the start of Week 6 - if criterion 3 (clause linking) is still weak, lead Monday Week 6 with a quick reteach of choosing the right joining word.",
      "Week 6 is the full Q3 task - combining 3-5 sentences with deletion and embedding. Today's mixing of coord and subord is the bridge.",
    ],
    watchFor: [
      "Students who all thumbs-up immediately - prompt them with one quick prove-it: 'tell me when you'd use BECAUSE vs WHEN'.",
      "Students who all thumbs-down - reassure: comma splices and wrong joining words are developmental progress, not failure.",
    ],
    tag: "[Day 5 | Fri | Closing reflection - end of cycle Week 2 of 6]",
  }, { requireSay: false, requireDo: false }));
  runSlideDiagnostics(s, pres, { respectSafeBottom: false });
})();

const outFile = path.join(OUT_DIR, "SC_T2W5_Lesson5_Fri_MIX_COORD_SUBORD.pptx");
pres.writeFile({ fileName: outFile }).then(() => {
  console.log("PPTX written to " + outFile);
}).catch((err) => {
  console.error("Build failed:", err);
  process.exitCode = 1;
});
