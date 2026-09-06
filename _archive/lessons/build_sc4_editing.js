"use strict";

// =============================================================================
//  Sentence Combining - Week 4 (Editing & Revising real writing)
//  Year 5/6 | Daily 15-minute warm-ups | Monday to Friday
//
//  ONE deck, all five days. The week shifts from combining sentences GIVEN to
//  students toward EDITING choppy writing that already exists. Final teaching
//  week before the post-test.
//
//  Per-day routine (identical for student recognition):
//    1. Goal            -> liSlide (today's goal + 3 "I can")
//    2. Review          -> 7 joining words (do NOT prescribe one)
//    3. Watch me edit   -> before/after worked example + "Watch out" non-example
//    4. Practise together-> withReveal: choppy paragraph -> model answer
//    5. Your turn       -> withReveal: fresh paragraph -> hidden sentence-stem scaffold
//    6. Exit check      -> task only (NO rubric on screen; rubric in notes)
//
//  Production happens in workbooks. Whiteboards only for the tiered check.
//  Scaffolds and model answers are hidden by default (on the next click).
//  No QR codes / digital forms. Australian spelling throughout.
// =============================================================================

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme } = require("../themes/factory");

// Variant 0 "Midnight Scholar": navy PRIMARY (1B2A3B), amber/gold ACCENT
// (8B7328), cream BG_LIGHT (F7F4EE). One variant for the whole week = unit
// cohesion. (weekToVariant(4) would give 3 "Ink & Paper"; 0 matches the
// requested navy + cream/amber palette more precisely.)
const T = createTheme("literacy", "grade56", 0);
const {
  C, FONT_H, FONT_B,
  SLIDE_W, SLIDE_H, SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, contentSlide, closingSlide,
  addCard, addFooter, addTopBar, addBadge, addTitle, addTextOnShape,
  withReveal, runSlideDiagnostics, composeNotes,
} = T;

const OUT_DIR = "output/Sentence Combining Week 4";
fs.mkdirSync(OUT_DIR, { recursive: true });

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Sentence Combining - Week 4 (Editing & Revising)";
pres.author = "James Hooke";

function footerFor(day) {
  return `Sentence Combining  |  Week 4  |  ${day}  |  Year 5/6`;
}

// ---------------------------------------------------------------------------
//  Shared note helpers
// ---------------------------------------------------------------------------

function dayBadges(s, typeText, typeColor, typeW, day) {
  addBadge(s, typeText, { color: typeColor, w: typeW });
  addBadge(s, day, { color: C.CHARCOAL, w: 1.5, x: 0.5 + typeW + 0.15 });
}

// Compose a standard note block in megaprompt section 45 order.
function notes({ time, move, say, doo, model, scaffoldReminder, extra, enabling, watchFor, tag }) {
  const tn = { text: `Time: ${time}. ${move}`, bullets: [] };
  if (model) tn.bullets.push(`Expected model answer: ${model}`);
  if (scaffoldReminder) tn.bullets.push(scaffoldReminder);
  if (extra) extra.forEach((e) => tn.bullets.push(e));
  return composeNotes({
    say,
    do: doo,
    teacherNotes: tn,
    enabling,
    watchFor,
    tag,
  });
}

// ---------------------------------------------------------------------------
//  before / after rich-text renderers
//  before parts: { t, del }  -> repeated/deletable words struck in ALERT
//  after  parts: { t, join, fold } -> joining words SUCCESS bold, folded
//                                     describing words SECONDARY bold
// ---------------------------------------------------------------------------

function beforeRuns(parts, fontSize) {
  return parts.map((p) => ({
    text: String(p.t),
    options: {
      fontSize,
      fontFace: FONT_B,
      color: p.del ? C.ALERT : C.CHARCOAL,
      strike: p.del ? true : false,
      breakLine: false,
    },
  }));
}

function afterRuns(parts, fontSize) {
  return parts.map((p) => ({
    text: String(p.t),
    options: {
      fontSize,
      fontFace: FONT_B,
      color: p.join ? C.SUCCESS : p.fold ? C.SECONDARY : C.CHARCOAL,
      bold: Boolean(p.join || p.fold),
      breakLine: false,
    },
  }));
}

// ===========================================================================
//  FRONT MATTER
// ===========================================================================

titleSlide(
  pres,
  "Sentence Combining",
  "Week 4: editing real writing from choppy to strong",
  "Year 5/6  |  Daily 15-minute warm-ups  |  Monday to Friday",
  composeNotes({
    say: [
      "This deck runs five 15-minute editing warm-ups, Monday to Friday.",
      "Every day follows the same six-slide routine so students know exactly what to do.",
      "This is the last teaching week before the post-test - keep the routine tight.",
    ],
    do: [
      "Navigate to today's Goal slide to begin.",
      "Have student workbooks open; mini-whiteboards ready for the Practise check.",
    ],
    teacherNotes: {
      text: "The week shifts from combining sentences we GIVE students toward EDITING choppy writing that already exists. Students must first find which sentences are worth combining, then combine them - the closest skill to real writing and to next week's post-test.",
      bullets: [
        "Model answers (Practise) and sentence-stem scaffolds (Your turn) are HIDDEN on the next click - reveal only when needed.",
        "Exit checks are marked from workbooks against a 5-criteria rubric held in the speaker notes. The rubric is never shown to students.",
        "Any of five teachers can run this deck as-is: every worked example is on the slide, all guidance is in the notes.",
      ],
    },
    tag: "[Sentence Combining | Week 4 | Title | VTLM 2.0 Explicit Teaching]",
  })
);

// Teacher Resources slide (immediately after title - megaprompt section 19/44)
contentSlide(
  pres,
  "Before You Teach", C.PRIMARY,
  "Teacher Resources & Materials",
  [
    "Student workbooks - all editing and every exit check is written here.",
    "Mini-whiteboards + markers - used ONLY for the quick tiered check in Practise Together.",
    "Projector + this deck - every worked example is on the slide; no extra prep.",
    "Hidden by default: the model answer (Practise) and the sentence stems (Your turn) are on the NEXT click - reveal only if students are stuck.",
    "Exit checks: mark from workbooks against the 5-criteria rubric in each Exit Check slide's notes. NEVER put the rubric on screen.",
    "No printed worksheets, QR codes or digital forms - everything runs from the board and books.",
  ],
  composeNotes({
    say: [
      "Today you will edit a choppy paragraph into stronger sentences - just like a real writer.",
      "We say this at the start every day so you can feel your own progress by the exit check.",
    ],
    do: [
      "Before the lesson: workbooks out, whiteboards stacked at the front row.",
      "Decide your 2-3 focus students for the exit check before you begin.",
    ],
    teacherNotes: {
      text: "This slide prepares any teacher to run the day with no further briefing.",
      bullets: [
        "Production is in workbooks. Whiteboards are only for the tiered 'front row show me' check in Practise Together.",
        "Reveals are deliberate: the goal is to see what students can do with reduced support before you scaffold.",
      ],
    },
    tag: "[Sentence Combining | Week 4 | Teacher Resources | VTLM 2.0]",
  }),
  footerFor("All week")
);

// Week overview (teacher-facing - megaprompt section 68b)
(function weekOverview() {
  const s = pres.addSlide();
  addTopBar(s, C.PRIMARY);
  addBadge(s, "For the Teacher", { color: C.PRIMARY, w: 2.2 });
  addTitle(s, "Week 4 Overview - the shift to editing");

  const cardY = CONTENT_TOP;
  const cardH = SAFE_BOTTOM - CONTENT_TOP;

  // Left: the shift + canonical before/after example
  const lx = 0.5, lw = 4.55;
  addCard(s, lx, cardY, lw, cardH, { strip: C.PRIMARY, fill: C.WHITE });
  s.addText("The shift this week", {
    x: lx + 0.2, y: cardY + 0.12, w: lw - 0.4, h: 0.3,
    fontSize: 14, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
  });
  s.addText(
    "Earlier weeks gave students clean sets of short sentences to combine. This week they EDIT real choppy writing: they must first find the sentences worth combining, then combine them - the closest skill to next week's post-test.",
    {
      x: lx + 0.2, y: cardY + 0.46, w: lw - 0.4, h: 1.20,
      fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, margin: 0, valign: "top",
      paraSpaceAfter: 3, fit: "shrink", shrinkText: true,
    }
  );
  s.addText("Before (choppy):", {
    x: lx + 0.2, y: cardY + 1.72, w: lw - 0.4, h: 0.24,
    fontSize: 11, fontFace: FONT_B, color: C.ALERT, bold: true, margin: 0,
  });
  s.addText("I went to the beach. The beach was hot. I built a sandcastle. The sandcastle was big. A wave came. The wave knocked it down.", {
    x: lx + 0.2, y: cardY + 1.98, w: lw - 0.4, h: 0.62,
    fontSize: 11, fontFace: FONT_B, color: C.CHARCOAL, italic: true, margin: 0, valign: "top",
    fit: "shrink", shrinkText: true,
  });
  s.addText("After (strong):", {
    x: lx + 0.2, y: cardY + 2.64, w: lw - 0.4, h: 0.24,
    fontSize: 11, fontFace: FONT_B, color: C.SUCCESS, bold: true, margin: 0,
  });
  s.addText("I went to the hot beach and built a big sandcastle, but a wave came and knocked it down.", {
    x: lx + 0.2, y: cardY + 2.90, w: lw - 0.4, h: 0.82,
    fontSize: 11.5, fontFace: FONT_H, color: C.CHARCOAL, bold: true, margin: 0, valign: "top",
    fit: "shrink", shrinkText: true,
  });

  // Right: daily focus
  const rx = 5.2, rw = 4.3;
  addCard(s, rx, cardY, rw, cardH, { strip: C.ACCENT, fill: C.BG_CARD });
  s.addText("Daily focus (same routine each day)", {
    x: rx + 0.2, y: cardY + 0.12, w: rw - 0.4, h: 0.3,
    fontSize: 14, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
  });
  const focusRows = [
    ["Mon", "Spot which sentences to combine; delete repeated words."],
    ["Tue", "Fold describing words into the main sentence."],
    ["Wed", "Use a mix of joining words, not just 'and'."],
    ["Thu", "Decide the main idea; make the rest supporting detail."],
    ["Fri", "Edit a whole paragraph alone - rehearsal for the post-test."],
  ];
  const rowH = 0.62;
  const rowY0 = cardY + 0.56;
  focusRows.forEach((row, i) => {
    const ry = rowY0 + i * rowH;
    addTextOnShape(s, row[0], {
      x: rx + 0.2, y: ry, w: 0.7, h: 0.42, rectRadius: 0.06,
      fill: { color: C.PRIMARY },
    }, { fontSize: 12, fontFace: FONT_B, color: C.WHITE, bold: true });
    s.addText(row[1], {
      x: rx + 1.0, y: ry, w: rw - 1.2, h: 0.5,
      fontSize: 11.5, fontFace: FONT_B, color: C.CHARCOAL, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });
  });

  addFooter(s, footerFor("All week"));
  s.addNotes(composeNotes({
    say: ["For your planning only - this slide is not for students."],
    do: ["Skim before the week so the routine and daily focus are clear."],
    teacherNotes: {
      text: "Each 15-minute session: Goal (30 sec) -> Review joining words (2 min) -> Watch me edit (3 min) -> Practise together (4 min) -> Your turn (4 min) -> Exit check (2 min).",
      bullets: [
        "HITS targeted across the week: setting goals, worked examples, explicit teaching, feedback.",
        "Australian Curriculum 2.0 English - Language (Text structure & organisation; Expressing & developing ideas) and Literacy (Writing); main and subordinate clauses (VCELA323).",
        "Comma rule reinforced Wed and Fri: a fronted 'when/because/if' clause takes a comma after it.",
      ],
    },
    tag: "[Sentence Combining | Week 4 | Overview | VTLM 2.0]",
  }, { requireSay: false, requireDo: false }));
  runSlideDiagnostics(s, pres, { respectSafeBottom: false });
})();

// ===========================================================================
//  PER-DAY BUILDERS
// ===========================================================================

// Review: the seven joining words (do NOT prescribe one).
function reviewSlide(day) {
  const s = pres.addSlide();
  addTopBar(s, C.SECONDARY);
  dayBadges(s, "Review", C.SECONDARY, 1.5, day);
  addTitle(s, "Joining words we know");

  const left = [
    ["and", C.PRIMARY, "joins ideas that go together"],
    ["but", C.ALERT, "shows a difference or surprise"],
    ["so", C.SUCCESS, "shows a result"],
    ["or", C.SECONDARY, "shows a choice"],
  ];
  const right = [
    ["because", C.PRIMARY, "gives a reason"],
    ["when", C.SECONDARY, "tells the time"],
    ["if", C.SUCCESS, "sets a condition"],
  ];

  function column(items, colX, pillW, pitch) {
    const y0 = CONTENT_TOP;
    items.forEach((it, i) => {
      const y = y0 + i * pitch;
      addTextOnShape(s, it[0], {
        x: colX, y, w: pillW, h: 0.5, rectRadius: 0.08,
        fill: { color: it[1] },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
      s.addText(it[2], {
        x: colX + pillW + 0.18, y, w: (colX < 5 ? 4.85 : 9.5) - (colX + pillW + 0.18), h: 0.5,
        fontSize: 14.5, fontFace: FONT_B, color: C.CHARCOAL, valign: "middle", margin: 0,
        fit: "shrink", shrinkText: true,
      });
    });
  }

  column(left, 0.5, 1.45, 0.61);   // 4 rows: 1.30 .. 3.13
  column(right, 5.1, 1.75, 0.82);  // 3 rows: 1.30 .. 2.94

  // One quick example (full width, bottom). Does NOT tell students which to use.
  const exY = 3.95;
  addCard(s, 0.5, exY, 9, 0.95, { strip: C.ACCENT, fill: C.BG_CARD });
  s.addText("One quick example", {
    x: 0.72, y: exY + 0.1, w: 3, h: 0.26,
    fontSize: 12, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
  });
  s.addText([
    { text: "Two short ideas: ", options: { color: C.CHARCOAL } },
    { text: "I was tired. I kept reading.", options: { color: C.CHARCOAL, italic: true } },
    { text: "   becomes   ", options: { color: C.MUTED } },
    { text: "I was tired, ", options: { color: C.CHARCOAL } },
    { text: "but", options: { color: C.SUCCESS, bold: true } },
    { text: " I kept reading.", options: { color: C.CHARCOAL } },
  ], {
    x: 0.72, y: exY + 0.4, w: 8.5, h: 0.45,
    fontSize: 15, fontFace: FONT_B, margin: 0, valign: "middle", fit: "shrink", shrinkText: true,
  });

  addFooter(s, footerFor(day));
  s.addNotes(notes({
    time: "2 min",
    move: "Quick retrieval so the joining words are warm before students choose their own. Do NOT tell students which word to use in their editing today.",
    say: [
      "Read the joining words with me. What job does each one do?",
      "Watch this: 'I was tired. I kept reading.' How could one joining word link them?",
      "Today you choose the joining word that fits your meaning - I will not choose for you.",
    ],
    doo: [
      "Point to each word; choral read the word and its job (about 90 seconds).",
      "Take one or two quick suggestions for the example, then move on.",
    ],
    scaffoldReminder: "Keep this brisk - it is a warm-up, not new teaching.",
    enabling: [
      "ENABLING (focus students 6-12 months below): pre-read 'and / but / so' only and have them point to each as you say it.",
      "EXTENDING (genuine challenge): ask a student to give a sentence where 'but' and 'because' would both fit and explain the difference in meaning.",
    ],
    watchFor: [
      "Students treating 'and' as the only option - affirm the others are available.",
    ],
    tag: `[Sentence Combining | Week 4 | ${day} Review | Retrieval]`,
  }));
  runSlideDiagnostics(s, pres, { respectSafeBottom: false });
  return s;
}

// Watch me edit: before / after worked example + "Watch out" non-example.
function teachSlide(day, t) {
  const s = pres.addSlide();
  addTopBar(s, C.PRIMARY);
  dayBadges(s, "Watch me", C.PRIMARY, 1.7, day);
  addTitle(s, t.title);

  const topY = CONTENT_TOP, topH = 2.05;
  // Before card
  addCard(s, 0.5, topY, 4.55, topH, { strip: C.ALERT, fill: C.WHITE });
  s.addText("Before - choppy (cross out repeats)", {
    x: 0.7, y: topY + 0.1, w: 4.2, h: 0.26,
    fontSize: 11.5, fontFace: FONT_B, color: C.ALERT, bold: true, margin: 0,
  });
  s.addText(beforeRuns(t.before, 17), {
    x: 0.7, y: topY + 0.42, w: 4.15, h: topH - 0.54,
    valign: "top", margin: 0, fit: "shrink", shrinkText: true,
  });

  // After card
  addCard(s, 5.15, topY, 4.35, topH, { strip: C.SUCCESS, fill: C.BG_CARD });
  s.addText("After - one strong sentence", {
    x: 5.35, y: topY + 0.1, w: 4.0, h: 0.26,
    fontSize: 11.5, fontFace: FONT_B, color: C.SUCCESS, bold: true, margin: 0,
  });
  const afterH = t.commaTip ? 1.05 : (topH - 0.54);
  s.addText(afterRuns(t.after, 17), {
    x: 5.35, y: topY + 0.42, w: 3.95, h: afterH,
    valign: "top", margin: 0, fit: "shrink", shrinkText: true,
  });
  if (t.commaTip) {
    s.addText(t.commaTip, {
      x: 5.35, y: topY + 1.55, w: 3.95, h: 0.42,
      fontSize: 11, fontFace: FONT_B, color: C.MUTED, italic: true, valign: "top", margin: 0,
    });
  }

  // Watch out non-example
  const wY = 3.5, wH = 1.4;
  addCard(s, 0.5, wY, 9, wH, { strip: C.ALERT, fill: C.WHITE });
  s.addText("Watch out", {
    x: 0.7, y: wY + 0.08, w: 3, h: 0.26,
    fontSize: 12, fontFace: FONT_B, color: C.ALERT, bold: true, margin: 0,
  });
  s.addText(`"${t.watchOut}"`, {
    x: 0.7, y: wY + 0.36, w: 8.6, h: 0.52,
    fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, italic: true, valign: "top", margin: 0,
    fit: "shrink", shrinkText: true,
  });
  s.addText([
    { text: "Why it is not stronger: ", options: { color: C.ALERT, bold: true } },
    { text: t.watchWhy, options: { color: C.CHARCOAL } },
  ], {
    x: 0.7, y: wY + 0.9, w: 8.6, h: 0.44,
    fontSize: 12.5, fontFace: FONT_B, valign: "top", margin: 0, fit: "shrink", shrinkText: true,
  });

  addFooter(s, footerFor(day));
  s.addNotes(notes({
    time: "3 min",
    move: t.move,
    say: t.say,
    doo: t.doo,
    model: t.model,
    extra: t.extra,
    scaffoldReminder: "Think aloud as you cross out repeats and fold in the describing words - show the decision, do not just present the answer.",
    enabling: [
      `ENABLING (focus students 6-12 months below): ${t.enabling}`,
      `EXTENDING (genuine challenge): ${t.extending}`,
    ],
    watchFor: t.watchFor,
    tag: `[Sentence Combining | Week 4 | ${day} Watch me | Worked example]`,
  }));
  runSlideDiagnostics(s, pres, { respectSafeBottom: false });
  return s;
}

// Practise together: choppy paragraph -> model answer (reveal). Write FIRST.
function practiseSlide(day, p) {
  const noteText = notes({
    time: "4 min",
    move: "Guided practice. Students write their revised version FIRST in their books, then you click to reveal the model. Do not reveal before every student has attempted.",
    say: [
      "Edit this paragraph in your book. Write your own version first.",
      "Front row, show me on whiteboards... now the next row. Who deleted a repeated word? Who folded a describing word in?",
      "Now let's compare with one strong version.",
    ],
    doo: [
      "Students write in books (about 2 minutes). Circulate.",
      "Tiered whiteboard check: front row show me, then next row - spot who is struggling.",
      "Only AFTER every student has attempted, click once to reveal the model.",
    ],
    model: p.model,
    scaffoldReminder: "The model answer is HIDDEN until the next click. There are valid alternatives - accept any version that keeps all ideas, links clauses sensibly and avoids repetition.",
    extra: [
      "One well-placed joining word can link three ideas - students do NOT need 'and' between every pair. Discourage 'and...and...and' while still affirming the attempt to join ideas.",
    ],
    enabling: [
      `ENABLING (focus students 6-12 months below): ${p.enabling}`,
      `EXTENDING (genuine challenge): ${p.extending}`,
    ],
    watchFor: [
      "Comma splices (ideas joined with only a comma). Acknowledge the link, then prompt for a joining word.",
      "Students who wait passively - everyone writes before any reveal.",
    ],
    tag: `[Sentence Combining | Week 4 | ${day} Practise | Guided + reveal]`,
  });

  return withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.SUCCESS);
      dayBadges(s, "Practise together", C.SUCCESS, 2.7, day);
      addTitle(s, "Practise together: edit this paragraph");

      addTextOnShape(s, "Edit this in your book. Write your version first.", {
        x: 0.5, y: CONTENT_TOP, w: 9, h: 0.6, rectRadius: 0.08,
        fill: { color: C.SECONDARY },
      }, { fontSize: 16, fontFace: FONT_B, color: C.WHITE, bold: true, align: "left", margin: 0.12 });

      addCard(s, 0.5, 2.05, 9, 1.5, { strip: C.PRIMARY, fill: C.WHITE });
      s.addText(p.before, {
        x: 0.75, y: 2.18, w: 8.5, h: 1.24,
        fontSize: 22, fontFace: FONT_B, color: C.CHARCOAL, valign: "middle", margin: 0,
        fit: "shrink", shrinkText: true,
      });

      addFooter(s, footerFor(day));
      s.addNotes(noteText);
      runSlideDiagnostics(s, pres, { respectSafeBottom: false });
      return s;
    },
    (s) => {
      const aY = 3.75, aH = 1.2;
      addCard(s, 0.5, aY, 9, aH, { strip: C.SUCCESS, fill: C.SUCCESS });
      s.addText("One strong version", {
        x: 0.75, y: aY + 0.1, w: 4, h: 0.26,
        fontSize: 12, fontFace: FONT_B, color: C.WHITE, bold: true, margin: 0,
      });
      s.addText(p.model, {
        x: 0.75, y: aY + 0.4, w: 8.5, h: aH - 0.5,
        fontSize: 19, fontFace: FONT_H, color: C.WHITE, bold: true, valign: "middle", margin: 0,
        fit: "shrink", shrinkText: true,
      });
      runSlideDiagnostics(s, pres, { respectSafeBottom: false });
    }
  );
}

// Your turn: fresh paragraph -> hidden sentence-stem scaffold (reveal).
function yourTurnSlide(day, y) {
  const noteText = notes({
    time: "4 min",
    move: "Independent production in workbooks. The sentence-stem scaffold is HIDDEN. Reveal it only if a focus student cannot start after about one minute.",
    say: [
      "Your turn. Edit this paragraph on your own, in your book.",
      "Find the sentences worth combining first, then choose your joining words.",
    ],
    doo: [
      "Students edit independently (about 3 minutes). Circulate to your focus students.",
      "If a student is stuck after about 1 minute, click once to reveal the sentence stems - then click back or leave them up for that student only.",
    ],
    model: y.model,
    scaffoldReminder: "The sentence stems are HIDDEN by default (next click). The aim is to see what students can do with reduced support - reveal sparingly.",
    extra: y.teacherTip ? [y.teacherTip] : undefined,
    enabling: [
      `ENABLING (focus students 6-12 months below): ${y.enabling}`,
      `EXTENDING (genuine challenge): ${y.extending}`,
    ],
    watchFor: [
      "Students combining everything into one over-long sentence - two strong sentences are fine.",
      "Dropped ideas - remind them every idea must survive the edit.",
    ],
    tag: `[Sentence Combining | Week 4 | ${day} Your turn | Independent + scaffold]`,
  });

  return withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.SECONDARY);
      dayBadges(s, "Your turn", C.SECONDARY, 1.8, day);
      addTitle(s, "Your turn: edit on your own");

      addTextOnShape(s, "Edit this paragraph on your own, in your workbook.", {
        x: 0.5, y: CONTENT_TOP, w: 9, h: 0.6, rectRadius: 0.08,
        fill: { color: C.PRIMARY },
      }, { fontSize: 16, fontFace: FONT_B, color: C.WHITE, bold: true, align: "left", margin: 0.12 });

      addCard(s, 0.5, 2.05, 9, 1.5, { strip: C.SECONDARY, fill: C.WHITE });
      s.addText(y.before, {
        x: 0.75, y: 2.18, w: 8.5, h: 1.24,
        fontSize: 22, fontFace: FONT_B, color: C.CHARCOAL, valign: "middle", margin: 0,
        fit: "shrink", shrinkText: true,
      });

      addFooter(s, footerFor(day));
      s.addNotes(noteText);
      runSlideDiagnostics(s, pres, { respectSafeBottom: false });
      return s;
    },
    (s) => {
      const aY = 3.75, aH = 1.2;
      addCard(s, 0.5, aY, 9, aH, { strip: C.ACCENT, fill: C.BG_CARD });
      s.addText("Stuck? Try a sentence starter", {
        x: 0.75, y: aY + 0.1, w: 5, h: 0.26,
        fontSize: 12, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
      });
      s.addText(y.stems.map((stem, i) => ({
        text: stem,
        options: { bullet: true, breakLine: i < y.stems.length - 1, color: C.CHARCOAL, fontSize: 15 },
      })), {
        x: 0.85, y: aY + 0.4, w: 8.4, h: aH - 0.5,
        fontFace: FONT_B, valign: "top", margin: 0, fit: "shrink", shrinkText: true,
      });
      runSlideDiagnostics(s, pres, { respectSafeBottom: false });
    }
  );
}

// Exit check: task only. NO rubric on screen. Rubric + samples in notes.
function exitSlide(day, e, exitNoteText) {
  const s = pres.addSlide();
  addTopBar(s, C.ASSESS);
  dayBadges(s, "Exit check", C.ASSESS, 1.9, day);
  addTitle(s, "Exit check: edit in your workbook", { color: C.ASSESS });

  addTextOnShape(s, "Edit this into stronger sentences in your workbook.", {
    x: 0.5, y: CONTENT_TOP, w: 9, h: 0.6, rectRadius: 0.08,
    fill: { color: C.ASSESS },
  }, { fontSize: 16, fontFace: FONT_B, color: C.WHITE, bold: true, align: "left", margin: 0.12 });

  addCard(s, 0.5, 2.05, 9, 2.6, { strip: C.ASSESS, fill: C.WHITE });
  s.addText(e.before, {
    x: 0.75, y: 2.25, w: 8.5, h: 2.2,
    fontSize: 24, fontFace: FONT_B, color: C.CHARCOAL, valign: "middle", margin: 0,
    fit: "shrink", shrinkText: true,
  });

  addFooter(s, footerFor(day));
  s.addNotes(exitNoteText);
  runSlideDiagnostics(s, pres, { respectSafeBottom: false });
  return s;
}

// Build the exit-check note (rubric + scored samples) as one plain-text block.
const RUBRIC_HEADER =
  "MARKING RUBRIC - teacher only. DO NOT show this to students.\n" +
  "Score each criterion 0 or 1 (total out of 5):\n" +
  "(1) All ideas retained\n" +
  "(2) One strong sentence, or a small number of strong sentences\n" +
  "(3) Appropriate clause linking - the right joining word for the meaning\n" +
  "(4) No unnecessary repetition\n" +
  "(5) No fragment, run-on or comma splice\n\n" +
  "CONSISTENCY RULE (apply identically in all five classes): an unbroken 'and...and...and' chain is technically grammatical, so it PASSES criterion 5 - it is not a run-on or comma splice. It still loses criterion 3 (no variety in linking) and usually criterion 4 (repetition). Marking this the same way everywhere is what keeps the post-test marks aligned across classes.\n\n" +
  "SERIAL COMMA (Week 6 PLC decision): a comma before the final 'and' in a list of three or more items (the Oxford comma) is acceptable in Australian English but not required. Do not mark students up or down for using or omitting it.";

function exitNotes(day, e) {
  const samplesText = e.samples
    .map((sm) => `  - "${sm.response}"  ->  ${sm.score}\n    ${sm.why}`)
    .join("\n");
  const tnText =
    `Time: 2 min. Students edit this short paragraph independently in their workbooks. The slide shows ONLY the task.\n\n` +
    `Expected model answer: ${e.model}\n\n` +
    RUBRIC_HEADER + "\n\n" +
    "SAMPLE RESPONSES (so all five classes mark consistently):\n" +
    samplesText + "\n\n" +
    "MARKING PHILOSOPHY: frame errors as progress. If a student writes a comma splice, say 'You have seen that these ideas connect - that is a step forward - now add a joining word.' No red ink." +
    (e.postTest ? "\n\nPOST-TEST NOTE: Friday's exit check most closely mirrors the post-test format next week. Marking strictness matters - apply the rubric exactly as you will next week." : "");

  return composeNotes({
    say: [
      "On your own now. Edit this paragraph into stronger sentences in your workbook.",
      "Use everything from this week - delete repeats, fold in describing words, choose joining words that fit.",
    ],
    do: [
      "Students write independently in workbooks (about 2 minutes). No talking.",
      "Mark your focus students against the 5-criteria rubric below. Keep the rubric off the screen.",
    ],
    teacherNotes: { text: tnText },
    enabling: [
      `ENABLING (focus students 6-12 months below): ${e.enabling}`,
      `EXTENDING (genuine challenge): ${e.extending}`,
    ],
    watchFor: [
      "Students copying the structure of an earlier slide - this paragraph is new and must be read first.",
      "Dropped ideas and comma splices - the two most common marks lost.",
    ],
    tag: `[Sentence Combining | Week 4 | ${day} Exit check | Assesses the week's skill]`,
  });
}

// ===========================================================================
//  DAY CONTENT
// ===========================================================================

const DAYS = [
  // ----- MONDAY -----------------------------------------------------------
  {
    name: "Monday",
    li: "I can edit a choppy paragraph into stronger sentences by deleting repeated words and using joining words.",
    sc: [
      "I can find words that are repeated and cross them out.",
      "I can join two short sentences into one using a joining word.",
      "I can read my new sentence and check it sounds smooth.",
    ],
    goalNote: notes({
      time: "30 sec",
      move: "Set today's goal and the growth frame: students will edit a paragraph at the end, so they can feel their progress across the 15 minutes.",
      say: [
        "By the end of today you will edit a paragraph on your own.",
        "Today we hunt for repeated words and cross them out, then join the ideas.",
        "If this feels new, that is okay - we will build it together.",
      ],
      doo: ["Read the goal and the three 'I can' statements aloud (about 30 seconds).", "Point out the exit check is coming - this is their chance to show growth."],
      scaffoldReminder: "Keep this to about 30 seconds - it frames the lesson, it is not a teaching slide.",
      enabling: ["ENABLING: read just the first 'I can' and make sure every student knows what 'repeated' means with a quick example."],
      watchFor: ["Students who think they must be perfect now - reassure them this is practice for the exit check."],
      tag: "[Sentence Combining | Week 4 | Monday Goal | Setting goals]",
    }),
    teach: {
      title: "Watch me edit: cross out repeats, then join",
      before: [
        { t: "I went to the park. " },
        { t: "The park", del: true },
        { t: " was muddy. I kicked the ball. " },
        { t: "The ball", del: true },
        { t: " was flat." },
      ],
      after: [
        { t: "I went to the " },
        { t: "muddy ", fold: true },
        { t: "park " },
        { t: "and", join: true },
        { t: " kicked the " },
        { t: "flat ", fold: true },
        { t: "ball." },
      ],
      commaTip: null,
      watchOut: "I went to the park and the park was muddy and I kicked the ball and the ball was flat.",
      watchWhy: "Everything is joined with 'and' and the repeated words are still there. It is longer, but it is not stronger.",
      model: "I went to the muddy park and kicked the flat ball.",
      move: "Show how to spot a repeated word (the park, the ball), cross it out, fold the describing word into the first mention, then join with one joining word.",
      say: [
        "Watch me read the choppy version. I can hear the same words coming back: 'the park', 'the ball'.",
        "I cross out the repeat and move the describing word up: 'the muddy park'.",
        "Now I join the two ideas with one joining word: 'and'. Six little sentences become one strong one.",
      ],
      doo: [
        "Read the before version aloud so students hear the choppiness.",
        "Strike the repeated words on the board; draw an arrow folding 'muddy' and 'flat' up.",
        "Read the after version - notice it is smoother.",
      ],
      extra: ["One joining word ('and') links several ideas here - you do not need 'and' between every pair."],
      enabling: "underline the two repeated phrases for them before you start, so the deletion is visible.",
      extending: "ask them to edit the same example a second valid way and say which reads better.",
      watchFor: [
        "Students who keep both copies of the repeated noun - point to the strike-through.",
        "Adjectives left as their own sentence instead of folded in.",
      ],
    },
    practise: {
      before: "I made a sandwich. The sandwich was huge. I ate it outside. The garden was sunny.",
      model: "I made a huge sandwich and ate it outside in the sunny garden.",
      enabling: "tell them to combine just the first two sentences, then the last two.",
      extending: "ask them to add one new detail of their own that fits the scene.",
    },
    yourTurn: {
      before: "I rode my scooter. My scooter was fast. I went down the hill. The hill was steep.",
      stems: [
        "I rode my ___ scooter and ___ .",
        "Cross out the second 'scooter'. Move 'fast' up next to the first one.",
      ],
      model: "I rode my fast scooter down the steep hill.",
      enabling: "reveal the first sentence stem and underline the repeated word 'scooter'.",
      extending: "ask them to write the edit two ways - one sentence, then two - and choose the better.",
    },
    exit: {
      before: "I woke up early. The morning was cold. I made some toast. The toast was warm.",
      model: "I woke up early on the cold morning and made some warm toast.",
      enabling: "accept a correct combination of just the first two and last two sentences.",
      extending: "ask them to add a 'because' or 'so' idea that still keeps all the original ideas.",
      postTest: false,
      samples: [
        { response: "I woke up early on the cold morning and made some warm toast.", score: "5/5", why: "All ideas kept; one strong sentence; 'and' fits; repeats deleted; no errors." },
        { response: "I woke up early and the morning was cold and I made some toast and the toast was warm.", score: "3/5", why: "Ideas kept (1); one grammatical sentence (1); under the and-chain rule above it passes criterion 5 (1). It loses criterion 3 - the repeated 'and' is not the best linking (0) - and criterion 4 - 'the toast' is still repeated (0). This is the same call as Friday's Sample 2." },
        { response: "I woke up early, the morning was cold, I made warm toast.", score: "2/5", why: "Ideas kept (1) and 'warm toast' folded with no repeat (1), but it is a comma splice with no joining word (0,0,0). Praise the link, then coach: add a joining word." },
      ],
    },
  },

  // ----- TUESDAY ----------------------------------------------------------
  {
    name: "Tuesday",
    li: "I can edit a choppy paragraph by folding describing words into the main sentence.",
    sc: [
      "I can underline the describing words in a paragraph.",
      "I can fold a describing word into the main sentence, like 'a fast game'.",
      "I can combine the sentences and keep all the details.",
    ],
    goalNote: notes({
      time: "30 sec",
      move: "Set today's goal: folding describing words in instead of leaving them in their own little sentences. Remind students the exit check is coming.",
      say: [
        "By the end of today you will edit a paragraph on your own.",
        "Today's move: when a sentence just describes something, fold that describing word into the main sentence.",
        "Some of you may remember doing this with single sentences - now we do it inside real writing.",
      ],
      doo: ["Read the goal and the three 'I can' statements (about 30 seconds).", "Name the describing-word move clearly."],
      scaffoldReminder: "About 30 seconds - this frames the session.",
      enabling: ["ENABLING: check every student can point to a describing word in a quick spoken example ('a red ball')."],
      watchFor: ["Students unsure what 'describing word' means - give one fast example before moving on."],
      tag: "[Sentence Combining | Week 4 | Tuesday Goal | Setting goals]",
    }),
    teach: {
      title: "Watch me edit: fold the describing words in",
      before: [
        { t: "We played a game. " },
        { t: "The game", del: true },
        { t: " was fast. I passed the ball. " },
        { t: "The ball", del: true },
        { t: " was slippery. We won." },
      ],
      after: [
        { t: "We played a " },
        { t: "fast ", fold: true },
        { t: "game " },
        { t: "and", join: true },
        { t: " passed the " },
        { t: "slippery ", fold: true },
        { t: "ball, " },
        { t: "so", join: true },
        { t: " we won." },
      ],
      commaTip: null,
      watchOut: "We played a game and the game was fast and I passed the ball and the ball was slippery and we won.",
      watchWhy: "The describing words (fast, slippery) are still floating in their own sentences. Fold them next to the thing they describe: 'a fast game', 'the slippery ball'.",
      model: "We played a fast game and passed the slippery ball, so we won.",
      move: "Show how a describing sentence ('The game was fast') becomes a describing word inside the main sentence ('a fast game'). Then join the ideas.",
      say: [
        "Listen: 'We played a game. The game was fast.' The second sentence only describes the game.",
        "I fold 'fast' straight into the first sentence: 'a fast game'. The repeated 'the game' goes.",
        "I do the same with 'slippery ball', then join with 'and' and 'so'.",
      ],
      doo: [
        "Underline the describing words (fast, slippery) on the board.",
        "Draw an arrow folding each describing word up into the main sentence.",
        "Read the after version and count: five sentences became one.",
      ],
      extra: ["Two joining words ('and', 'so') do the work here - not an 'and' between every idea."],
      enabling: "underline 'fast' and 'slippery' for them and show the first fold.",
      extending: "ask them to add one more describing word of their own that fits, without making the sentence clumsy.",
      watchFor: [
        "Describing words left as separate sentences.",
        "Students dropping a describing word entirely - it must survive, just in a new place.",
      ],
    },
    practise: {
      before: "We had a swimming lesson. The pool was freezing. I swam four laps. The laps were slow. I felt proud.",
      model: "We had a swimming lesson in the freezing pool. I swam four slow laps, so I felt proud.",
      enabling: "tell them to fold just 'freezing' and 'slow' in, then join with one word.",
      extending: "ask them to use a 'because' idea to explain why they felt proud, keeping all ideas.",
    },
    yourTurn: {
      before: "I went skating. The rink was crowded. I wore new skates. The skates were tight. I had fun.",
      stems: [
        "I went skating at the ___ rink, and ___ .",
        "Fold the describing words (crowded, new, tight) next to the thing they describe.",
      ],
      model: "I went skating at the crowded rink, and I wore tight new skates, but I had fun.",
      enabling: "reveal the first stem and underline 'crowded' for them.",
      extending: "ask for two versions and a sentence explaining which folds read more smoothly.",
    },
    exit: {
      before: "I played soccer. The field was muddy. I scored a goal. The goal was amazing. We cheered.",
      model: "I played soccer on the muddy field, and I scored an amazing goal, so we cheered.",
      enabling: "accept a version that folds 'muddy' and 'amazing' and joins with one word.",
      extending: "ask them to keep all ideas but use a different joining word from the one they reached for first.",
      postTest: false,
      samples: [
        { response: "I played soccer on the muddy field, and I scored an amazing goal, so we cheered.", score: "5/5", why: "All ideas kept; describing words folded; 'and'/'so' fit; no repeats; no errors." },
        { response: "I played soccer and the field was muddy and I scored a goal and the goal was amazing and we cheered.", score: "3/5", why: "Ideas kept (1), one grammatical sentence (1), no run-on (1), but the 'and' chain is weak linking (0) and the describing words were not folded - 'the goal' repeats (0)." },
        { response: "I played soccer on the muddy field and scored an amazing goal.", score: "4/5", why: "Strong, smooth sentence (2,3,4,5 all met) but 'we cheered' was dropped - criterion 1 fails (0). Coach: keep every idea." },
      ],
    },
  },

  // ----- WEDNESDAY --------------------------------------------------------
  {
    name: "Wednesday",
    li: "I can edit a choppy paragraph using a mix of joining words, not just 'and'.",
    sc: [
      "I can read the joining words and pick one that fits.",
      "I can use more than one kind of joining word in my editing.",
      "I can use a comma when a sentence starts with 'when', 'because' or 'if'.",
    ],
    goalNote: notes({
      time: "30 sec",
      move: "Set today's goal: variety. Strong editing uses different joining words for different meanings, not 'and' every time. Flag the comma rule for fronted clauses.",
      say: [
        "By the end of today you will edit a paragraph on your own.",
        "Today's move: choose different joining words - 'but' for a surprise, 'so' for a result, 'when' for time.",
        "We will also watch one comma rule when a sentence starts with 'when' or 'because'.",
      ],
      doo: ["Read the goal and 'I can' statements (about 30 seconds).", "Foreshadow the comma rule you will model."],
      scaffoldReminder: "About 30 seconds - frame and move on.",
      enabling: ["ENABLING: focus them on just swapping one 'and' for a 'but' or 'so' today."],
      watchFor: ["Students who default to 'and' for everything - that is exactly today's target."],
      tag: "[Sentence Combining | Week 4 | Wednesday Goal | Setting goals]",
    }),
    teach: {
      title: "Watch me edit: mix the joining words",
      before: [
        { t: "We caught the bus to the zoo. The bus was noisy. We saw the lions. " },
        { t: "The lions", del: true },
        { t: " were asleep. The monkeys were hiding." },
      ],
      after: [
        { t: "When", join: true },
        { t: " we caught the " },
        { t: "noisy ", fold: true },
        { t: "bus to the zoo, we saw the " },
        { t: "sleeping ", fold: true },
        { t: "lions, " },
        { t: "but", join: true },
        { t: " the monkeys were hiding." },
      ],
      commaTip: "Comma tip: 'When ...' starts the sentence, so a comma goes after the first part (... to the zoo, we saw ...).",
      watchOut: "We caught the bus and the bus was noisy and we went to the zoo and we saw the lions and the lions were asleep and the monkeys were hiding.",
      watchWhy: "Every idea is joined with 'and'. Use different joining words - 'but' shows the surprise that the monkeys were hiding, and the writing sounds stronger.",
      model: "When we caught the noisy bus to the zoo, we saw the sleeping lions, but the monkeys were hiding.",
      move: "Show how choosing 'when' and 'but' (instead of 'and' four times) makes the writing stronger, and model the comma after a fronted 'When ...' clause.",
      say: [
        "If I join everything with 'and', it just lists. Let me choose words that show meaning.",
        "I start with 'When we caught the noisy bus to the zoo' - because it starts with 'When', I put a comma after it.",
        "Then 'but' shows the surprise: the monkeys were hiding. Different joining words, stronger writing.",
      ],
      doo: [
        "On the board, write the 'and...and...and' version, then improve it live with 'when' and 'but'.",
        "Circle the comma after the fronted 'When ...' clause and say the rule.",
        "Read both versions so students hear the difference.",
      ],
      extra: [
        "Comma rule: when 'when/because/if' starts the sentence, put a comma after the first clause. When it sits in the middle, no comma is needed.",
        "Team decision (Week 6 PLC): the serial / Oxford comma (a comma before the final 'and' in a list of three or more) is acceptable in Australian English but not required. Teach it as an option and do not mark students up or down for using or omitting it.",
      ],
      enabling: "give them the joining words 'when' and 'but' on a card and let them slot them in.",
      extending: "ask them to rewrite the paragraph using at least three different joining words.",
      watchFor: [
        "The 'and...and...and' chain - prompt for a 'but' or 'so'.",
        "A missing comma after a fronted 'When ...' clause.",
      ],
    },
    practise: {
      before: "We visited the farm. The barn was huge. We fed the goats. The goats were hungry. It started to rain. We ran inside.",
      model: "When we visited the farm, we fed the hungry goats in the huge barn, but it started to rain, so we ran inside.",
      enabling: "tell them to use at least one word that is not 'and'.",
      extending: "challenge them to use 'when', 'but' and 'so' all correctly in one edit.",
    },
    yourTurn: {
      before: "We went to the beach. The sand was hot. We swam in the sea. The water was cold. We built a sandcastle. We went home.",
      stems: [
        "We went to the beach, ___ the sand was hot.",
        "Use one joining word for the surprise and a different one for the result.",
      ],
      model: "We went to the beach, but the sand was hot, so we swam in the cold sea. We built a sandcastle and went home.",
      teacherTip: "The model deliberately breaks into two sentences - that is a strong, deliberate choice that stops the writing becoming a run-on. Accept either one strong sentence or two; a student may reasonably ask why the model stopped joining and started a new sentence.",
      enabling: "reveal the first stem so they only have to choose the joining word.",
      extending: "ask them to use a less common joining word (although, since, while) correctly.",
    },
    exit: {
      before: "We went to the park. The park was crowded. We found a swing. The swing was free. We played.",
      model: "When we went to the crowded park, we found a free swing and played.",
      enabling: "accept any version that uses at least one joining word that is not 'and' and keeps all ideas.",
      extending: "ask for a version using a fronted 'When ...' clause with the comma placed correctly.",
      postTest: false,
      samples: [
        { response: "When we went to the crowded park, we found a free swing and played.", score: "5/5", why: "All ideas; comma after the fronted 'When'; describing words folded; mix of structure; no errors." },
        { response: "We went to the crowded park and we found a free swing and we played.", score: "3/5", why: "Ideas kept (1), one grammatical sentence (1), no splice (1), but only 'and' is used - no variety (0) and 'we ... we ... we' repeats (0)." },
        { response: "When we went to the park the park was crowded, we found a free swing, we played.", score: "1/5", why: "All ideas just kept (1), but 'the park' repeats (0), no comma after the fronted clause and comma splices follow (0,0,0). Lots to build on - praise the 'When' start, then add joining words and the comma." },
      ],
    },
  },

  // ----- THURSDAY ---------------------------------------------------------
  {
    name: "Thursday",
    li: "I can edit a choppy paragraph by choosing the main idea and making the rest supporting detail.",
    sc: [
      "I can choose the most important idea in a paragraph.",
      "I can make that idea my main sentence and add the rest as detail.",
      "I can avoid comma splices by using a joining word.",
    ],
    goalNote: notes({
      time: "30 sec",
      move: "Set today's goal: not every idea is equally important. Choose the main idea for the main sentence; the rest becomes supporting detail.",
      say: [
        "By the end of today you will edit a paragraph on your own.",
        "Today's move: pick the most important idea - that becomes your main sentence. The rest gives extra detail.",
        "We will also watch out for comma splices and fix them with a joining word.",
      ],
      doo: ["Read the goal and 'I can' statements (about 30 seconds).", "Define 'main idea' simply: the thing the writing is really about."],
      scaffoldReminder: "About 30 seconds - frame the session.",
      enabling: ["ENABLING: practise choosing the main idea aloud from a two-sentence example first."],
      watchFor: ["Students treating every sentence as equally important - that is today's focus."],
      tag: "[Sentence Combining | Week 4 | Thursday Goal | Setting goals]",
    }),
    teach: {
      title: "Watch me edit: choose the main idea",
      before: [
        { t: "It was sports day. I ran in the race. I felt nervous. I came first. Everyone cheered." },
      ],
      after: [
        { t: "On sports day I ran in the race, " },
        { t: "and", join: true },
        { t: " I felt nervous, " },
        { t: "but", join: true },
        { t: " I came first, " },
        { t: "so", join: true },
        { t: " everyone cheered." },
      ],
      commaTip: "The big idea here is 'I came first'. 'I felt nervous' is supporting detail, joined with 'but'.",
      watchOut: "I ran in the race, I felt nervous, I came first, everyone cheered.",
      watchWhy: "These whole ideas are joined with only commas - that is a comma splice. Commas cannot join complete ideas on their own. Choose a main idea and use joining words (and, but, so).",
      model: "On sports day I ran in the race, and I felt nervous, but I came first, so everyone cheered.",
      move: "Show how to decide which idea matters most ('I came first'), make it the main clause, and turn the others into supporting detail with joining words - and why comma-only joins are a splice.",
      say: [
        "Which idea matters most here? 'I came first.' That is my main idea.",
        "The other ideas support it: 'I felt nervous' is a contrast, so I use 'but'. 'Everyone cheered' is a result, so I use 'so'.",
        "If I only used commas between these ideas, that would be a comma splice - watch out for that.",
      ],
      doo: [
        "Ask the class to point to the most important idea before you combine.",
        "Show the comma-splice version, then fix it live with joining words.",
        "Read the strong version aloud.",
      ],
      extra: ["Comma splice = two whole ideas joined with only a comma. Fix it with a joining word or a full stop."],
      enabling: "tell them which idea is the main one, then they add the detail.",
      extending: "ask them to use a less common joining word (although, since, while) to add the supporting detail.",
      watchFor: [
        "Comma splices - the key error this week.",
        "Students who make a trivial idea the main sentence - prompt 'what is the most important thing that happened?'",
      ],
    },
    practise: {
      before: "It was book week. I dressed up. My costume was a dragon. We had a parade. I won a prize. I was excited.",
      model: "For book week I dressed up in my dragon costume, and when we had a parade, I won a prize, so I was excited.",
      enabling: "tell them the main idea is 'I won a prize'; they add the rest as detail.",
      extending: "ask them to start the sentence with a 'When ...' clause and place the comma correctly.",
    },
    yourTurn: {
      before: "We had an assembly. I read a poem. My voice was shaky. People listened. I finished. They clapped.",
      stems: [
        "Choose your main idea (the most important thing): ___ .",
        "Turn 'my voice was shaky' into supporting detail using 'but' or 'when'.",
      ],
      model: "At assembly I read a poem, and my voice was shaky, but people listened, so when I finished they clapped.",
      enabling: "reveal the first stem and suggest 'I read a poem' as the main idea.",
      extending: "ask them to edit it two ways with different main ideas and say which version is stronger.",
    },
    exit: {
      before: "It was our class concert. I played the drums. I made a mistake. The crowd clapped. I smiled.",
      model: "At our class concert I played the drums, and I made a mistake, but the crowd clapped, so I smiled.",
      enabling: "accept a version that keeps all ideas and uses 'but' to add the mistake as detail.",
      extending: "ask them to use three different joining words correctly with no comma splice.",
      postTest: false,
      samples: [
        { response: "At our class concert I played the drums, and I made a mistake, but the crowd clapped, so I smiled.", score: "5/5", why: "All ideas; main idea chosen; and/but/so all fit; no repeats; no splice." },
        { response: "I played the drums at our class concert, I made a mistake, but the crowd clapped, I smiled.", score: "2/5", why: "All ideas kept (1); no repetition (1). The middle 'but' is a correct join, but criterion 3 still fails because two of the three joins are comma-only (0), and those comma splices fail criteria 2 and 5 (0,0). Score 2/5. Praise the 'but', then replace the two comma-only joins with joining words." },
        { response: "I played the drums and made a mistake, but the crowd clapped, so I smiled.", score: "4/5", why: "Strong combine, good joining words, no errors (2,3,4,5 met) but 'our class concert' was dropped - criterion 1 fails (0)." },
      ],
    },
  },

  // ----- FRIDAY -----------------------------------------------------------
  {
    name: "Friday",
    li: "I can edit a whole choppy paragraph into a few strong sentences on my own.",
    sc: [
      "I can find the sentences that are worth combining.",
      "I can edit a whole paragraph into a few strong sentences on my own.",
      "I can use different joining words and correct punctuation.",
    ],
    goalNote: notes({
      time: "30 sec",
      move: "Set today's goal: pull the whole week together. Today is the rehearsal for next week's post-test - everything they have learned, on their own.",
      say: [
        "By the end of today you will edit a paragraph on your own - this is your rehearsal for next week.",
        "Use every move: find what to combine, delete repeats, fold describing words, choose joining words, watch your punctuation.",
        "Show yourself how far you have come since Monday.",
      ],
      doo: ["Read the goal and 'I can' statements (about 30 seconds).", "Name today as the post-test rehearsal."],
      scaffoldReminder: "About 30 seconds - then straight into the routine.",
      enabling: ["ENABLING: remind your focus students of the one move they found most useful this week."],
      watchFor: ["Nerves about the post-test - frame today as a friendly practice run."],
      tag: "[Sentence Combining | Week 4 | Friday Goal | Setting goals]",
    }),
    teach: {
      title: "Watch me edit: put it all together",
      before: [
        { t: "My family went camping. We put up the tent. " },
        { t: "The tent", del: true },
        { t: " was small. It started to rain. We cooked dinner. " },
        { t: "The dinner", del: true },
        { t: " was hot. We told stories." },
      ],
      after: [
        { t: "My family went camping, " },
        { t: "and", join: true },
        { t: " we put up the " },
        { t: "small ", fold: true },
        { t: "tent. " },
        { t: "When", join: true },
        { t: " it started to rain, we cooked a " },
        { t: "hot ", fold: true },
        { t: "dinner and told stories." },
      ],
      commaTip: "Two strong sentences are fine. 'When it started to rain' starts a sentence, so a comma follows it.",
      watchOut: "My family went camping we put up the small tent it started to rain we cooked a hot dinner and told stories.",
      watchWhy: "This is a run-on. The ideas are jammed together with no joining words or full stops, so the reader cannot tell where one idea ends and the next begins.",
      model: "My family went camping, and we put up the small tent. When it started to rain, we cooked a hot dinner and told stories.",
      move: "Pull every move together: delete repeats, fold describing words, choose a mix of joining words, and break into two strong sentences rather than one giant run-on.",
      say: [
        "This is everything from the week. I find the repeats ('the tent', 'the dinner') and cross them out.",
        "I fold 'small' and 'hot' in, then choose joining words - 'and', then 'When' for the rain.",
        "Two strong sentences read better than one giant run-on. Watch the comma after 'When it started to rain'.",
      ],
      doo: [
        "Model the whole edit step by step on the board.",
        "Show the run-on version and explain why it fails the reader.",
        "Read the final two-sentence version aloud.",
      ],
      extra: ["Splitting into two strong sentences is a valid, strong choice - it is not all-or-nothing into one sentence."],
      enabling: "give them the two repeated phrases underlined and one joining word to start.",
      extending: "ask them to edit it as one long sentence and as two sentences, then say which suits the writing better.",
      watchFor: [
        "Run-ons - no joining words and no full stops.",
        "Over-long single sentences - two strong sentences are encouraged.",
      ],
    },
    practise: {
      before: "We had a sleepover. My friends came. We watched a movie. The movie was funny. We ate popcorn. We stayed up late. We fell asleep.",
      model: "My friends came over for a sleepover, and we watched a funny movie and ate popcorn. We stayed up late, but we finally fell asleep.",
      enabling: "let them aim for two or three sentences rather than one.",
      extending: "ask for at least three different joining words used correctly.",
    },
    yourTurn: {
      before: "It was the school holidays. We drove to the city. The traffic was slow. We visited a museum. The museum was huge. We bought souvenirs. We were tired.",
      stems: [
        "Plan your main ideas first, then join them.",
        "Use at least two different joining words (try 'but' and 'when').",
      ],
      model: "In the school holidays we drove to the city, but the traffic was slow. When we visited the huge museum, we bought souvenirs, so we were tired.",
      enabling: "reveal the stems and suggest they break the paragraph into two sentences.",
      extending: "ask them to add one new detail of their own that fits, keeping all the original ideas.",
    },
    exit: {
      before: "I had a birthday party. My friends arrived. We played games. The games were fun. We ate cake. The cake was chocolate. We opened presents.",
      model: "When my friends arrived at my birthday party, we played fun games and ate chocolate cake, and then we opened presents.",
      enabling: "accept two or three strong sentences that keep all ideas and avoid splices.",
      extending: "ask for a version using a fronted clause with a correct comma and at least two different joining words.",
      postTest: true,
      samples: [
        { response: "When my friends arrived at my birthday party, we played fun games and ate chocolate cake, and then we opened presents.", score: "5/5", why: "All ideas; comma after the fronted clause; describing words folded; varied linking; no errors." },
        { response: "I had a birthday party and my friends arrived and we played games and the games were fun and we ate cake and the cake was chocolate and we opened presents.", score: "3/5", why: "All ideas (1); one grammatical sentence (1); under the and-chain rule above it passes criterion 5 (1). It loses criterion 3 - pure 'and' chain, no variety (0) - and criterion 4 - 'the games'/'the cake' repeat with nothing folded (0). Same call as Monday's Sample 2: an and-chain passes 5 but loses 3 and 4. Tell the student it is drifting toward a run-on even though it scores the point." },
        { response: "I had a birthday party when my friends arrived, we played fun games, we ate chocolate cake, we opened presents.", score: "2/5", why: "All ideas kept (1) and describing words folded with no repeats (1), but the list is held together by comma splices (criteria 2,3,5 fail). Praise the 'when' opening and the folding, then replace the commas with 'and' or full stops. Mirror post-test strictness here." },
      ],
    },
  },
];

// ===========================================================================
//  BUILD EACH DAY IN ORDER
// ===========================================================================

DAYS.forEach((day) => {
  // 1. Goal (Learning Intention + 3 "I can")
  liSlide(pres, day.li, day.sc, day.goalNote, footerFor(day.name));
  // 2. Review joining words
  reviewSlide(day.name);
  // 3. Watch me edit (worked example + non-example)
  teachSlide(day.name, day.teach);
  // 4. Practise together (reveal: model answer)
  practiseSlide(day.name, day.practise);
  // 5. Your turn (reveal: hidden scaffold)
  yourTurnSlide(day.name, day.yourTurn);
  // 6. Exit check (task only; rubric + samples in notes)
  exitSlide(day.name, day.exit, exitNotes(day.name, day.exit));
});

// ===========================================================================
//  CLOSING REFLECTION (whole-week)
// ===========================================================================

closingSlide(
  pres,
  {
    reflectionPrompt: "Think back to Monday. What can you do now with choppy writing that you could not do before?",
    scItems: [
      "I can find the sentences worth combining.",
      "I can delete repeats and fold in describing words.",
      "I can choose joining words that fit and punctuate them correctly.",
    ],
    selfAssessment: {
      prompt: "How ready do you feel to edit a paragraph on your own next week?",
      options: ["Confident", "Getting there", "Need more practice"],
    },
  },
  composeNotes({
    say: [
      "Look how far you have come this week - from choppy lists to strong, smooth sentences.",
      "Next week you will do this on your own in the post-test. You are ready to practise it.",
    ],
    do: [
      "Run a quick thumbs self-check against the three 'I can' statements.",
      "Note which students still need the comma-splice fix before the post-test.",
    ],
    teacherNotes: {
      text: "Use the self-check to decide who needs a short top-up before next week's post-test. Friday's exit check is your best predictor of post-test performance.",
      bullets: [
        "Acknowledge progress explicitly - this is the feedback that builds confidence.",
        "No new content here; this is reflection and goal-setting only.",
      ],
    },
    tag: "[Sentence Combining | Week 4 | Closing | Reflection]",
  })
);

// ===========================================================================
//  WRITE
// ===========================================================================

(async () => {
  const outFile = path.join(OUT_DIR, "Sentence Combining Week 4 - Editing and Revising.pptx");
  await pres.writeFile({ fileName: outFile });
  console.log(`PPTX written to ${outFile}`);
})();
