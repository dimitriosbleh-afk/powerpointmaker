"use strict";

// Persuasive Writing - Term 3 Week 6, Session 4 (Year 5/6 Literacy)
// "The Mighty Middle: one reason, proved"
//
// Victorian Curriculum 2.0: English, Language and Literacy, Levels 5-6 - using
// modality to adjust certainty, and structuring a paragraph that supports one
// reason with evidence.
//
// UNIT ANCHOR (locked): Bold Beginning. Mighty Middle. Excellent Ending.
// Voice all the way through. Today is the MIGHTY MIDDLE band.
//
// The model paragraph is the ARC Level 5-6 sample text, quoted exactly. Its
// topic is cyberbullying, so the I Do carries a CARE line and the topic is
// named as the example, not opened as a discussion.

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme } = require("../themes/factory");
const { composeGlanceNotes } = require("../themes/core/composeNotes");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addLinedArea,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
  PAGE, hex,
} = require("../themes/pdf_helpers");

const P = require("./persuasive_lib");

const UNIT_VARIANT = 5;
const T = createTheme("literacy", "grade56", UNIT_VARIANT);
const {
  C, FONT_H, FONT_B, SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, keyWordSlide, closingSlide,
  addCard, addFooter, addTextOnShape, addRevealAnswerBar,
  clickBuild, runSlideDiagnostics,
} = T;

const SESSION = 4;
const FOOTER = "Persuasive Writing | Week 6 Session 4 | Year 5/6 Literacy";
const OUT_DIR = "output/Persuasive_W6_S4";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const BUILDER_RES = makeSessionResource(SESSION,
  "Mighty Middle Builder",
  "The model paragraph annotated, then a four-step frame to write your own.");
const RESOURCE_ITEMS = [BUILDER_RES];
fs.mkdirSync(RES_DIR, { recursive: true });

// The four moves of the model paragraph, in order, quoted exactly.
const MODEL_MOVES = [
  { label: "Reason", color: "PRIMARY", text: P.ARC.body1Reason },
  { label: "Evidence", color: "SECONDARY", text: P.ARC.body1Evidence },
  { label: "Explain", color: "ASSESS", text: P.ARC.body1Explain },
  { label: "Punch", color: "ALERT", text: "No child should have to feel afraid every time they check their messages." },
  { label: "Link", color: "SUCCESS", text: P.ARC.body1Link },
];

// ─── Teacher notes ──────────────────────────────────────────────────────────

const NOTES_TITLE =
  "Session 4. One line: we have the shape and the words, today we build the part that does the arguing.";

const NOTES_RESOURCES =
  "Prep slide. Print the Mighty Middle Builder. Students need their Word Bank and Structure Card.\n" +
  "Choose and cue up one BTN story before the lesson. Any current issue with two sides works. It is the launch and the You Do topic option.\n" +
  "CATCH-UP: missed a session? Give a spare Word Bank, point at the Mighty Middle band on the Structure Card, and start at the launch. Nothing today depends on Sessions 2 or 3.";

const NOTES_OVERVIEW =
  "For you, not the class. Skip past it when projecting.\n" +
  "Today is the Mighty Middle band: one reason, proved, in four moves. Modality is taught first because it is what makes the reason sound certain.\n" +
  "Decision points: the modality hinge, the boards check in the We Do, and the exit ticket. Between them keep the pace brisk.";

const NOTES_LAUNCH = composeGlanceNotes({
  answer: "open - listen for a clear issue and two possible sides.",
  beats: [
    [
      "PLAY the BTN story you chose. Two to three minutes, no pausing.",
      "SAY: Watch for the argument, not just the facts.",
    ],
    [
      "ASK: What is the issue, and what are the two sides?",
      "40 sec. Turn and tell. Partner B first.",
      "EXPECT: the issue named, plus a for and an against.",
    ],
    [
      "SCAN the room as pairs talk.",
      "80%+ naming two sides -> next slide.",
      "Less -> replay the last thirty seconds, name one side yourself, re-ask.",
    ],
  ],
  trap: [
    "retelling the story instead of finding the argument.",
    "Fix: ask what someone might disagree with, student names it.",
  ],
  prep: "Low-coupling launch: any BTN story works and no earlier session is assumed. Whole block under 8 minutes.",
  sources: "Teacher-selected BTN story from ABC Behind the News.",
  tag: "[Launch | Retention and recall | HITS 6]",
});

const NOTES_LI = composeGlanceNotes({
  beats: [
    "POINT to the intention. SAY: Today we build the part that does the arguing.",
    "SAY: Read the I can statements with me. Everyone, together, on three.",
    "SAY: Everyone leaves able to sort words by how certain they sound.",
  ],
  prep: "SC1 reachable by all; SC2 is the core target the exit ticket assesses; SC3 stretches into explaining the effect of modality.",
  tag: "[LI/SC | Planning made visible | HITS 1]",
});

const NOTES_VOCAB = composeGlanceNotes({
  answer: "modality is how certain a word makes you sound.",
  beats: [
    [
      "SAY: Modality is how sure you sound.",
      "Might is unsure. Should is likely. Must leaves no room.",
    ],
    [
      "ASK: Which sounds most certain, might, should or must?",
      "10 sec. Choral response: Everyone, together, on three.",
      "EXPECT: must.",
    ],
    "SAY: In persuasive writing you usually want high modality. Sound sure.",
  ],
  trap: [
    "thinking high modality means shouting or exclamation marks.",
    "Fix: say a must sentence calmly, student hears the certainty.",
  ],
  prep: "The full three columns are on the printed Word Bank from Session 2. Three words on screen is enough here.",
  sources: "Persuasive word banks supplied by the teacher.",
  tag: "[Key Vocabulary | Knowledge and memory | HITS 6]",
});

const NOTES_MODALITY = composeGlanceNotes({
  answer: "the same idea, moved from unsure to certain.",
  beats: [
    [
      "SAY: One sentence, three ways. Watch how the certainty climbs.",
      "CLICK through low, then medium, then high.",
    ],
    [
      "ASK: Which one would convince a principal to act?",
      "15 sec. Fingers at your chest. One, two or three.",
      "EXPECT: three.",
    ],
    [
      "SCAN the fingers.",
      "80%+ on three -> cold call one: which word did that?",
      "Less -> read one and three back to back, re-ask.",
    ],
  ],
  trap: [
    "thinking a longer sentence sounds more certain.",
    "Fix: point at the one changed word, student names it.",
  ],
  stretch: "Rewrite the high version using a different high modality word.",
  help: "Give the three modality columns from the Word Bank, open.",
  prep: "Teach modality before the paragraph, because the reason sentence needs it in the very first move.",
  sources: "Persuasive word banks supplied by the teacher.",
  tag: "[I Do | Explicit teaching | SC1 | HITS 4]",
});

const NOTES_MODEL = composeGlanceNotes({
  answer: "reason, evidence, explain, punch, link.",
  beats: [
    [
      "SAY: One paragraph from the sample text. One reason only.",
      "READ it aloud straight through before labelling anything.",
    ],
    [
      "CLICK through the five labels in order.",
      "SAY: Every mighty middle paragraph does these jobs, in this order.",
    ],
    [
      "ASK: Which move is the proof?",
      "10 sec. Choral response: Everyone, together, on three.",
      "EXPECT: evidence.",
    ],
  ],
  trap: [
    "treating the punch line as the evidence.",
    "Fix: ask which line has a fact in it, student points to it.",
  ],
  stretch: "Find the high modality words in this paragraph.",
  help: "Cover all but two moves, name reason and evidence only.",
  care: "cyberbullying is the sample topic. Name it, do not open it up.",
  prep: "Quoted exactly from the ARC sample. The moves are the same four on the Structure Card, plus the emotive punch from Session 2.",
  sources: P.ARC_SOURCE_LINE,
  tag: "[I Do | Explicit teaching | SC2 | HITS 4]",
});

const NOTES_CFU = composeGlanceNotes({
  answer: "B. Must leaves the reader no room.",
  beats: [
    [
      "SAY: Three sentences, same idea. One sounds certain.",
      "Do not call out. Boards for this one.",
    ],
    [
      "ASK: Which is high modality, A, B or C?",
      "30 sec. Write it... chin it... show me.",
      "EXPECT: B.",
    ],
    [
      "SCAN boards, back row first.",
      "80%+ -> cold call one B: which word? Then reveal.",
      "Less -> read A and B back to back, re-ask.",
    ],
    "REVEAL after boards are scanned. SAY: One word. Must.",
  ],
  trap: [
    "picking C because should sounds polite and grown up.",
    "Fix: ask which one a principal could ignore, student re-tests.",
  ],
  prep: "The hinge. A is low, C is medium. Both wrong answers show a student not yet hearing certainty in a single word.",
  tag: "[CFU hinge | Supported application | SC1 | HITS 7, 8]",
});

const NOTES_WEDO = composeGlanceNotes({
  answer: "open - a reason sentence with because, then a fact.",
  beats: [
    "SAY: I take the reason. You take the evidence.",
    [
      "ASK: Give me the evidence sentence. Start with For example.",
      "60 sec. Write it... chin it... show me.",
      "EXPECT: a countable fact about paper in our room.",
    ],
    [
      "SCAN boards for a fact, not an opinion.",
      "80%+ -> cold call one strong, one shaky: is that a fact?",
      "Less -> count one bin together, re-ask.",
    ],
  ],
  trap: [
    "writing a second opinion instead of evidence.",
    "Fix: ask how we could check it, student rewrites it.",
  ],
  stretch: "Add the explain move, starting This means that.",
  help: "Give the starter and a counted number to use.",
  prep: "Evidence is the move students skip. Insist on something checkable, even a rough count from your own bin.",
  tag: "[We Do | Supported application | SC2 | HITS 4, 5]",
});

const NOTES_YOUDO = composeGlanceNotes({
  answer: "open - one paragraph doing all four moves in order.",
  beats: [
    [
      "SAY: Your own topic, or the BTN issue from the start of the lesson.",
      "SAY: One reason only. All four moves, in order.",
    ],
    [
      "POINT to the Builder. SAY: One move per line. Do not skip evidence.",
      "TIME: fifteen minutes.",
    ],
    [
      "CIRCULATE. Check move two is a fact before you read anything else.",
      "COLLECT two paragraphs to read at the closing.",
    ],
  ],
  trap: [
    "writing three reasons in one paragraph.",
    "Fix: ring the second reason, student saves it for a new paragraph.",
  ],
  stretch: "Write a second paragraph for a different reason.",
  help: "Give the Builder with the reason sentence already started.",
  prep: "Own topic or BTN, not the paper bin. Different content from the We Do on purpose.",
  tag: "[You Do | Mastery and application | SC2 | HITS 10]",
});

const NOTES_EXIT = composeGlanceNotes({
  answer: "a because sentence, then a checkable fact.",
  beats: [
    [
      "SAY: Two sentences only. The reason, then the evidence.",
      "SAY: Topic is on the board. Trees.",
    ],
    [
      "TIME: four minutes. Sheet on the desk when you finish.",
      "COLLECT by walking the rows.",
    ],
    [
      "SORT as you collect: has a fact, or two opinions.",
      "A thick second pile -> reteach evidence at the top of Week 7.",
    ],
  ],
  prep: "Assesses the core target: a reason with evidence. The fact is the evidence, not the length.",
  tag: "[Exit Ticket | Mastery and application | SC2 | HITS 8]",
});

const NOTES_CLOSING = composeGlanceNotes({
  answer: "open - most students should show two.",
  beats: [
    "READ the two collected paragraphs aloud. SAY: Listen for the evidence move.",
    [
      "ASK: Rate yourself on the three statements.",
      "15 sec. Fingers at your chest. One, two or three. Show me.",
      "EXPECT: most on two.",
    ],
    "SAY: Next week you pick one topic and publish it three different ways.",
  ],
  prep: "Last session of Week 6. Anyone showing one starts Week 7 with the Builder beside them.",
  tag: "[Closing | Retention and recall | HITS 1, 9]",
});

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  titleSlide(pres, "The Mighty Middle",
    "One reason, proved, in four moves",
    "Week 6 Session 4 | Year 5/6 Literacy", NOTES_TITLE);

  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  (() => {
    const s = P.customSlide(pres, T, "For the teacher", C.MUTED, "Session 4 at a glance", { badgeW: 2.4 });
    const cardY = CONTENT_TOP;
    const cardH = SAFE_BOTTOM - CONTENT_TOP;

    addCard(s, 0.5, cardY, 4.4, cardH, { strip: C.PRIMARY });
    s.addText("Where this sits", {
      x: 0.7, y: cardY + 0.12, w: 4.0, h: 0.28,
      fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText([
      { text: "Today is the MIGHTY MIDDLE band of the anchor.", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: P.ANCHOR_PHRASE, options: { color: C.PRIMARY, bold: true, breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "The four moves, held for both weeks:", options: { bold: true, breakLine: true } },
      { text: "Reason: [position] because [reason]", options: { breakLine: true } },
      { text: "Evidence: According to... For example...", options: { breakLine: true } },
      { text: "Explain: This means that...", options: { breakLine: true } },
      { text: "Link: Therefore, ...", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "Evidence is the move students skip. Insist on something checkable.", options: { italic: true, color: C.MUTED } },
    ], {
      x: 0.7, y: cardY + 0.46, w: 4.0, h: cardH - 0.60,
      fontSize: 12.5, fontFace: FONT_B, color: C.CHARCOAL,
      valign: "top", margin: 0, paraSpaceAfter: 3,
    });

    addCard(s, 5.1, cardY, 4.4, cardH, { strip: C.ACCENT });
    s.addText("How today works", {
      x: 5.3, y: cardY + 0.12, w: 4.0, h: 0.28,
      fontSize: 14, fontFace: FONT_H, color: C.ACCENT, bold: true, margin: 0,
    });
    s.addText([
      { text: "Shape: example-first. Modality is taught before the paragraph, because the reason sentence needs it immediately.", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "Before the lesson:", options: { bold: true, color: C.ACCENT, breakLine: true } },
      { text: "choose one BTN story with two clear sides. It is the launch and a You Do topic option.", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "Sensitive content:", options: { bold: true, color: C.ACCENT, breakLine: true } },
      { text: "the model paragraph is about cyberbullying. Name it as the sample topic, do not open a discussion.", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "Decision points: the modality hinge, the We Do boards check, the exit ticket.", options: { italic: true, color: C.MUTED } },
    ], {
      x: 5.3, y: cardY + 0.46, w: 4.0, h: cardH - 0.60,
      fontSize: 12.5, fontFace: FONT_B, color: C.CHARCOAL,
      valign: "top", margin: 0, paraSpaceAfter: 3,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_OVERVIEW);
    runSlideDiagnostics(s, pres);
  })();

  // Launch - BTN
  (() => {
    const s = P.customSlide(pres, T, "Launch", C.SECONDARY, "Today's news, two sides");
    const y0 = CONTENT_TOP;

    addCard(s, 0.5, y0, 9, 1.30, { strip: C.SECONDARY });
    s.addText("Teacher-selected BTN story", {
      x: 0.75, y: y0 + 0.14, w: 8.5, h: 0.38,
      fontSize: 20, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0,
    });
    s.addText("Watch for the argument, not just the facts.", {
      x: 0.75, y: y0 + 0.58, w: 8.5, h: 0.56,
      fontSize: 22, fontFace: FONT_B, color: C.CHARCOAL, margin: 0, valign: "middle",
    });

    const qy = y0 + 1.50;
    s.addText("What is the issue? What are the two sides?", {
      x: 0.5, y: qy, w: 9, h: 1.05,
      fontSize: 34, fontFace: FONT_H, color: C.PRIMARY, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    addTextOnShape(s, "Turn and tell. Partner B first.", {
      x: 2.95, y: qy + 1.20, w: 4.1, h: 0.66, rectRadius: 0.08,
      fill: { color: C.WHITE }, line: { color: C.SECONDARY, width: 1.5 },
    }, {
      fontSize: 17, fontFace: FONT_B, color: C.SECONDARY, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_LAUNCH);
    runSlideDiagnostics(s, pres);
  })();

  liSlide(pres,
    "We are learning to build a paragraph that proves one reason.",
    [
      "I can sort words into low, medium and high modality.",
      "I can write a mighty middle paragraph with a reason and evidence.",
      "I can explain why high modality makes my writing stronger.",
    ],
    NOTES_LI, FOOTER);

  keyWordSlide(pres, {
    word: "modality",
    meaning: "How certain a word makes you sound.",
    example: "It might help. It should help. It must happen.",
    routine: ["Say it", "Rank it", "Use it"],
    color: C.SECONDARY,
    title: "The word for today",
  }, NOTES_VOCAB, FOOTER);

  // I Do A - the modality ladder
  (() => {
    const s = P.customSlide(pres, T, "I Do", C.PRIMARY, "One sentence, three levels of sure");
    const y0 = CONTENT_TOP;
    const rows = [
      { level: "Low", color: C.MUTED, text: "Our school could maybe put a bin in some classrooms.", words: P.MODALITY.low.slice(0, 5) },
      { level: "Medium", color: C.SECONDARY, text: "Our school should probably put a bin in each classroom.", words: P.MODALITY.medium.slice(0, 5) },
      { level: "High", color: C.ALERT, text: "Our school must put a bin in every classroom.", words: ["must", "always", "never", "definitely", "certainly"] },
    ];
    const rh = 0.90;
    const gap = 0.13;

    clickBuild(s, rows.map((row, i) => () => {
      const ry = y0 + i * (rh + gap);
      s.addShape("roundRect", {
        x: 0.5, y: ry, w: 1.45, h: rh, rectRadius: 0.08,
        fill: { color: row.color },
      });
      s.addText(row.level, {
        x: 0.5, y: ry, w: 1.45, h: rh,
        fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      s.addShape("roundRect", {
        x: 2.05, y: ry, w: 7.45, h: rh, rectRadius: 0.08,
        fill: { color: C.WHITE }, line: { color: row.color, width: 1.5 },
      });
      s.addText(row.text, {
        x: 2.25, y: ry + 0.06, w: 7.05, h: 0.46,
        fontSize: 18, fontFace: FONT_B, color: C.CHARCOAL,
        valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
      s.addText(row.words.join(" | "), {
        x: 2.25, y: ry + 0.52, w: 7.05, h: 0.32,
        fontSize: 12.5, fontFace: FONT_B, color: row.color, bold: true,
        valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
    }));

    addTextOnShape(s, "Persuasive writing lives at the bottom of this ladder.", {
      x: 0.5, y: y0 + 3.10, w: 9, h: 0.68, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, {
      fontSize: 19, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_MODALITY);
    runSlideDiagnostics(s, pres);
  })();

  // I Do B - the model paragraph, five moves
  (() => {
    const s = P.customSlide(pres, T, "I Do", C.PRIMARY, "One reason, proved: the model");
    const y0 = CONTENT_TOP;
    const rh = 0.68;
    const gap = 0.09;

    // The paragraph is on screen from the start; the LABELS are what build,
    // so the teacher reads it whole before naming any move.
    MODEL_MOVES.forEach((move, i) => {
      const ry = y0 + i * (rh + gap);
      s.addShape("roundRect", {
        x: 2.15, y: ry, w: 7.35, h: rh, rectRadius: 0.07,
        fill: { color: C.WHITE }, line: { color: C.MUTED, width: 1 },
      });
      s.addText(move.text, {
        x: 2.32, y: ry, w: 7.01, h: rh,
        fontSize: 12.5, fontFace: FONT_B, color: C.CHARCOAL,
        valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
    });

    clickBuild(s, MODEL_MOVES.map((move, i) => () => {
      const ry = y0 + i * (rh + gap);
      s.addShape("roundRect", {
        x: 0.5, y: ry, w: 1.50, h: rh, rectRadius: 0.07,
        fill: { color: C[move.color] },
      });
      s.addText(move.label, {
        x: 0.5, y: ry, w: 1.50, h: rh,
        fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
    }));

    addFooter(s, FOOTER);
    s.addNotes(NOTES_MODEL);
    runSlideDiagnostics(s, pres);
  })();

  // CFU hinge - modality
  (() => {
    const s = P.customSlide(pres, T, "CFU", C.ALERT, "Which one sounds certain?");
    const y0 = CONTENT_TOP;

    addTextOnShape(s, "Topic: our school should plant more trees", {
      x: 0.5, y: y0, w: 9, h: 0.48, rectRadius: 0.08,
      fill: { color: C.ALERT },
    }, {
      fontSize: 16, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    const options = [
      { key: "A", text: "Our school could possibly plant a few more trees sometime." },
      { key: "B", text: "Our school must plant more trees, and it must happen this year." },
      { key: "C", text: "Our school should probably plant more trees fairly soon." },
    ];
    const oh = 0.66;
    const og = 0.12;
    options.forEach((opt, i) => {
      const oy = y0 + 0.60 + i * (oh + og);
      s.addShape("roundRect", {
        x: 0.5, y: oy, w: 9, h: oh, rectRadius: 0.08,
        fill: { color: C.WHITE }, line: { color: C.PRIMARY, width: 1.5 },
      });
      s.addShape("roundRect", {
        x: 0.68, y: oy + 0.11, w: 0.44, h: 0.44, rectRadius: 0.22,
        fill: { color: C.PRIMARY },
      });
      s.addText(opt.key, {
        x: 0.68, y: oy + 0.11, w: 0.44, h: 0.44,
        fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      s.addText(opt.text, {
        x: 1.28, y: oy, w: 7.97, h: oh,
        fontSize: 16.5, fontFace: FONT_B, color: C.CHARCOAL,
        valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
    });

    clickBuild(s, [
      () => {
        addRevealAnswerBar(s, ["B - must, twice"], {
          y: P.REVEAL_Y, h: P.REVEAL_H, fontSize: 24, color: C.SUCCESS,
        });
      },
    ]);

    addFooter(s, FOOTER);
    s.addNotes(NOTES_CFU);
    runSlideDiagnostics(s, pres);
  })();

  // We Do
  (() => {
    const s = P.customSlide(pres, T, "We Do", C.SUCCESS, "I take the reason. You take the proof.");
    const y0 = CONTENT_TOP;

    addCard(s, 0.5, y0, 9, 1.10, { strip: C.PRIMARY });
    s.addText("Reason (mine)", {
      x: 0.75, y: y0 + 0.09, w: 4, h: 0.28,
      fontSize: 14, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText([
      { text: "A scrap paper bin ", options: { color: C.CHARCOAL } },
      { text: "must", options: { color: C.ALERT, bold: true } },
      { text: " be in every room ", options: { color: C.CHARCOAL } },
      { text: "because", options: { color: C.PRIMARY, bold: true } },
      { text: " we waste far too much.", options: { color: C.CHARCOAL } },
    ], {
      x: 0.75, y: y0 + 0.40, w: 8.5, h: 0.60,
      fontSize: 21, fontFace: FONT_B, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    const ey = y0 + 1.26;
    addCard(s, 0.5, ey, 9, 1.10, { strip: C.SECONDARY });
    s.addText("Evidence (yours)", {
      x: 0.75, y: ey + 0.09, w: 4, h: 0.28,
      fontSize: 14, fontFace: FONT_B, color: C.SECONDARY, bold: true, margin: 0,
    });
    s.addText([
      { text: "For example, ", options: { color: C.SECONDARY, bold: true } },
      { text: "[something we can actually count]", options: { color: C.MUTED, italic: true } },
    ], {
      x: 0.75, y: ey + 0.40, w: 8.5, h: 0.60,
      fontSize: 21, fontFace: FONT_B, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    addTextOnShape(s, "A fact, not another opinion. Write it... chin it... show me.", {
      x: 0.5, y: ey + 1.28, w: 9, h: 0.80, rectRadius: 0.08,
      fill: { color: C.WHITE }, line: { color: C.SUCCESS, width: 1.5 },
    }, {
      fontSize: 19, fontFace: FONT_B, color: C.SUCCESS, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_WEDO);
    runSlideDiagnostics(s, pres);
  })();

  // You Do
  (() => {
    const s = P.customSlide(pres, T, "You Do", C.ASSESS, "Your reason. Your proof.");
    const y0 = CONTENT_TOP;

    addTextOnShape(s, "Your own topic, or the BTN issue from the start", {
      x: 0.5, y: y0, w: 9, h: 0.72, rectRadius: 0.08,
      fill: { color: C.ASSESS },
    }, {
      fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    const moves = [
      { label: "Reason", frame: "[Position] because [reason].", color: C.PRIMARY },
      { label: "Evidence", frame: "For example... / According to...", color: C.SECONDARY },
      { label: "Explain", frame: "This means that...", color: C.ASSESS },
      { label: "Link", frame: "Therefore, ...", color: C.SUCCESS },
    ];
    const rh = 0.52;
    const gap = 0.09;
    moves.forEach((move, i) => {
      const ry = y0 + 0.88 + i * (rh + gap);
      s.addShape("roundRect", {
        x: 0.5, y: ry, w: 1.85, h: rh, rectRadius: 0.07,
        fill: { color: move.color },
      });
      s.addText(move.label, {
        x: 0.5, y: ry, w: 1.85, h: rh,
        fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      s.addShape("roundRect", {
        x: 2.50, y: ry, w: 7.0, h: rh, rectRadius: 0.07,
        fill: { color: C.WHITE }, line: { color: move.color, width: 1.2 },
      });
      s.addText(move.frame, {
        x: 2.70, y: ry, w: 6.6, h: rh,
        fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL,
        valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
    });

    s.addText("One reason only. Fifteen minutes on the Builder.", {
      x: 0.5, y: y0 + 3.34, w: 9, h: 0.44,
      fontSize: 16, fontFace: FONT_B, color: C.MUTED, italic: true,
      align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
    runSlideDiagnostics(s, pres);
  })();

  P.exitTicketPanel(pres, T, {
    topic: "Our school should plant more trees",
    task: "Write the REASON sentence, then the EVIDENCE sentence.",
    cue: "Evidence means something we could check.",
    taskSize: 27,
    footer: FOOTER,
    notes: NOTES_EXIT,
  });

  closingSlide(pres, {
    reflectionPrompt: "Which of the four moves was hardest to write today?",
    scItems: [
      "I can sort words into low, medium and high modality.",
      "I can write a mighty middle paragraph with a reason and evidence.",
      "I can explain why high modality makes my writing stronger.",
    ],
    selfAssessment: "Fingers at your chest. One, two or three. Show me.",
    takeaways: ["Reason. Evidence. Explain. Link."],
  }, NOTES_CLOSING);

  const pptxPath = path.join(OUT_DIR, "Persuasive Writing W6 S4.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);
}

// ─── Resource: Mighty Middle Builder ────────────────────────────────────────

async function buildBuilder() {
  const doc = createPdf({ title: BUILDER_RES.name });
  let y = addPdfHeader(doc, "Mighty Middle Builder", {
    subtitle: "One reason. Four moves. Do not skip the evidence.",
    color: C.PRIMARY,
    lessonInfo: "Week 6 Session 4 | Year 5/6 Literacy",
  });

  y = addSectionHeading(doc, "The model, labelled", y, { color: C.PRIMARY });
  y += 4;

  // The annotated model, drawn as labelled bands so the printed page carries
  // the same representation as the slide, not a prose description of it.
  MODEL_MOVES.forEach((move) => {
    const boxH = 40;
    doc.save();
    doc.roundedRect(PAGE.MARGIN, y, 78, boxH, 3).fill(hex(C[move.color]));
    doc.fontSize(9).font("Sans-Bold").fillColor("#FFFFFF");
    doc.text(move.label, PAGE.MARGIN + 4, y + boxH / 2 - 5, { width: 70, align: "center" });
    doc.roundedRect(PAGE.MARGIN + 84, y, PAGE.CONTENT_W - 84, boxH, 3)
      .lineWidth(0.7).strokeColor("#9CA3AF").stroke();
    doc.fontSize(9).font("Sans").fillColor("#000000");
    doc.text(move.text, PAGE.MARGIN + 90, y + 6, { width: PAGE.CONTENT_W - 98 });
    doc.restore();
    y += boxH + 5;
  });

  y += 4;
  y = addBodyText(doc,
    "The paragraph argues ONE reason. The evidence move is the one writers skip, and it is the "
    + "one that makes a reader believe you.",
    y, { italic: true });
  y += 6;

  doc.addPage();
  y = PAGE.MARGIN;
  y = addSectionHeading(doc, "Now yours. One reason only.", y, { color: C.PRIMARY });
  y += 6;

  doc.fontSize(11).font("Sans-Bold").fillColor("#000000");
  doc.text("My topic:", PAGE.MARGIN, y);
  doc.moveTo(PAGE.MARGIN + 62, y + 14).lineTo(PAGE.MARGIN + PAGE.CONTENT_W, y + 14)
    .lineWidth(0.9).strokeColor("#000000").stroke();
  y += 30;

  const frames = [
    { label: "Reason", starter: "[My position] because ...", lines: 2, color: C.PRIMARY },
    { label: "Evidence", starter: "For example, ... / According to ...", lines: 2, color: C.SECONDARY },
    { label: "Explain", starter: "This means that ...", lines: 2, color: C.ASSESS },
    { label: "Link", starter: "Therefore, ...", lines: 2, color: C.SUCCESS },
  ];

  frames.forEach((frame) => {
    doc.save();
    doc.roundedRect(PAGE.MARGIN, y, PAGE.CONTENT_W, 22, 3).fill(hex(frame.color));
    doc.fontSize(10).font("Sans-Bold").fillColor("#FFFFFF");
    doc.text(frame.label + "   " + frame.starter, PAGE.MARGIN + 8, y + 6,
      { width: PAGE.CONTENT_W - 16 });
    doc.restore();
    y += 28;
    y = addLinedArea(doc, y, frame.lines, { lineSpacing: 30 });
    y += 14;
  });

  y = addTipBox(doc,
    "Challenge: write a second paragraph for a different reason, then decide which of the two "
    + "you would put first and why.",
    y, { color: C.ACCENT });

  addPdfFooter(doc, "Persuasive Writing | Week 6 Session 4 | Year 5/6 Literacy");
  await writePdf(doc, path.join(OUT_DIR, BUILDER_RES.fileName));
  console.log("Wrote " + BUILDER_RES.name);
}

(async () => {
  await build();
  await buildBuilder();
})().catch((err) => {
  console.error(err);
  process.exit(1);
});
