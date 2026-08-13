"use strict";

// Persuasive Writing - Term 3 Week 6, Session 3 (Year 5/6 Literacy)
// "Voice, part 2: hook them and tell them what to do"
//
// Victorian Curriculum 2.0: English, Language and Literacy, Levels 5-6 - how
// rhetorical questions, imperatives and overstatement position a reader.
//
// UNIT ANCHOR (locked): Bold Beginning. Mighty Middle. Excellent Ending.
// Voice all the way through. Today is still inside the VOICE band.
//
// The ARC sample text opens with "Imagine..." - a bossy verb straight off the
// teacher's own word bank. That coincidence is the spine of the I Do: the real
// Level 5-6 exemplar uses the exact tool on their printed sheet.

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

const SESSION = 3;
const FOOTER = "Persuasive Writing | Week 6 Session 3 | Year 5/6 Literacy";
const OUT_DIR = "output/Persuasive_W6_S3";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const LAB_RES = makeSessionResource(SESSION,
  "Hook Lab",
  "Worked hooks, then three hook slots for a new topic using each tool.");
const RESOURCE_ITEMS = [LAB_RES];
fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Teacher notes ──────────────────────────────────────────────────────────

const NOTES_TITLE =
  "Session 3. One line: strong words are wasted if the reader stops at line one, so today we hook them.";

const NOTES_RESOURCES =
  "Prep slide. Print the Hook Lab. Students need their Persuasive Word Bank from Session 2 and their Structure Card from Session 1.\n" +
  "Whiteboards and markers out.\n" +
  "CATCH-UP: missed Session 2? Give them a spare Word Bank and start them at the launch. Today only needs the bank, not Session 2 itself.";

const NOTES_OVERVIEW =
  "For you, not the class. Skip past it when projecting.\n" +
  "Three tools today: rhetorical questions, bossy verbs, and exaggeration or promises. All three live in the Voice band of the anchor.\n" +
  "Decision points: the keep-reading hinge, the boards check in the We Do, and the exit ticket. Between them keep the pace brisk.";

const NOTES_LAUNCH = composeGlanceNotes({
  answer: "open - any sensible upgrade of the flat word.",
  beats: [
    [
      "SAY: Quick warm up from last session. One word, one swap.",
      "SAY: Then a problem: strong words are wasted if nobody reads on.",
    ],
    [
      "ASK: Upgrade the word bad in this sentence.",
      "30 sec. Write it... chin it... show me.",
      "EXPECT: appalling, dreadful, outrageous.",
    ],
    [
      "SCAN boards.",
      "80%+ -> reveal, then straight into today.",
      "Less -> open the bank at BAD, read three aloud, re-ask.",
    ],
  ],
  trap: [
    "reaching for nice instead of a strong word.",
    "Fix: ask how bad, student picks a stronger word.",
  ],
  prep: "Low-coupling launch: retrieves Session 2 but a student who missed it can answer straight from the printed bank.",
  tag: "[Launch | Retention and recall | HITS 6]",
});

const NOTES_LI = composeGlanceNotes({
  beats: [
    [
      "POINT to the intention.",
      "SAY: Today we grab the reader and tell them what to do.",
    ],
    "SAY: Read the I can statements with me. Everyone, together, on three.",
    "SAY: Everyone leaves able to spot a rhetorical question.",
  ],
  prep: "SC1 reachable by all; SC2 is the core target the exit ticket assesses; SC3 stretches into explaining the effect.",
  tag: "[LI/SC | Planning made visible | HITS 1]",
});

const NOTES_VOCAB = composeGlanceNotes({
  answer: "a question you ask to make the reader think, not to get an answer.",
  beats: [
    [
      "SAY: A rhetorical question is not really a question.",
      "You already know what the reader will answer in their head.",
    ],
    [
      "ASK: Would you like more homework tonight?",
      "5 sec. Thumbs only, voices off. Show me... now.",
      "EXPECT: thumbs down, and some laughing.",
    ],
    [
      "SAY: I did not need your answer.",
      "I made you take a side. That is the trick.",
    ],
  ],
  trap: [
    "asking a real question the reader could answer either way.",
    "Fix: ask what answer they wanted, student rewrites it.",
  ],
  prep: "The demonstration IS the definition. Ask the homework question before explaining the term.",
  tag: "[Key Vocabulary | Knowledge and memory | HITS 6]",
});

const NOTES_IDO = composeGlanceNotes({
  answer: "Imagine is a bossy verb. Line two is rhetorical.",
  beats: [
    "SAY: The real Level 5-6 sample. Look at its first word.",
    [
      "CLICK 1 for the bossy verb. CLICK 2 for the question.",
      "SAY: Two tools, two lines. That is the whole opening.",
    ],
    [
      "ASK: What does it make you do before you read on?",
      "20 sec. Turn and tell. Partner A first.",
      "EXPECT: picture it, take a side.",
    ],
  ],
  trap: [
    "reading Imagine as a normal word, not an order.",
    "Fix: read it without Imagine, student hears the drop.",
  ],
  stretch: "Find the emotive words in these two lines.",
  help: "Cover line two, work only on Imagine.",
  care: "cyberbullying. Name it as the example topic, not a discussion.",
  prep: "Two exact lines only. The full text and its topic are handled properly in Session 4.",
  sources: P.ARC_SOURCE_LINE,
  tag: "[I Do | Explicit teaching | SC1 | HITS 4]",
});

const NOTES_QUESTIONS = composeGlanceNotes({
  answer: "open - any starter that makes the reader take a side.",
  beats: [
    [
      "SAY: Twelve ready-made openings. They are on your printed bank too.",
      "SAY: You only need two or three that sound like you.",
    ],
    [
      "ASK: Pick one starter and finish it about wasted paper.",
      "40 sec. Write it... chin it... show me.",
      "EXPECT: Fed up with... or Isn't it time... completed.",
    ],
    [
      "SCAN boards for a question mark.",
      "80%+ -> cold call two, read them aloud with feeling.",
      "Less -> finish one starter yourself on the board, re-ask.",
    ],
  ],
  trap: [
    "writing a statement and adding a question mark.",
    "Fix: read it aloud as a question, student rewrites the opening.",
  ],
  prep: "Twelve starters, quoted from the teacher's own bank. Do not read all twelve aloud.",
  sources: "Persuasive word banks supplied by the teacher.",
  tag: "[I Do | Knowledge and memory | SC1 | HITS 4]",
});

const NOTES_VERBS = composeGlanceNotes({
  answer: "bossy verbs give an order, promises tell the reader what they will get.",
  beats: [
    [
      "POINT to the left column. SAY: Bossy verbs give the reader an order.",
      "POINT to the right. SAY: Promises tell them what they will get.",
    ],
    [
      "ASK: Which one belongs at the very END of a piece?",
      "15 sec. Fingers at your chest. One for bossy, two for promise.",
      "EXPECT: one, the bossy verb.",
    ],
    "SAY: An excellent ending tells the reader what to DO. That is your bossy verb.",
  ],
  trap: [
    "using a promise where a call to action is needed.",
    "Fix: ask what the reader must do now, student swaps to a bossy verb.",
  ],
  prep: "This connects straight to the Excellent Ending band and is used again in Week 7 on the poster.",
  sources: "Persuasive word banks supplied by the teacher.",
  tag: "[I Do | Knowledge and memory | SC1 | HITS 4]",
});

const NOTES_CFU = composeGlanceNotes({
  answer: "B. It asks a question and tells you what should happen.",
  beats: [
    [
      "SAY: Three openings, same topic. Only one makes you read on.",
      "Do not call out. Boards for this one.",
    ],
    [
      "ASK: Which opening makes you keep reading, A, B or C?",
      "30 sec. Write it... chin it... show me.",
      "EXPECT: B.",
    ],
    [
      "SCAN boards, back row first.",
      "80%+ -> cold call one B: which tool did that? Then reveal.",
      "Less -> read A aloud flatly, then B, re-ask.",
    ],
    "REVEAL after boards are scanned. SAY: A question, then a must. Two tools.",
  ],
  trap: [
    "picking C because it sounds like proper writing.",
    "Fix: ask whether C made them feel anything, student re-tests.",
  ],
  prep: "The hinge. A announces the topic, C is a neutral fact. Both wrong answers show a student not yet hearing the pull of a hook.",
  tag: "[CFU hinge | Supported application | SC2 | HITS 7, 8]",
});

const NOTES_WEDO = composeGlanceNotes({
  answer: "open - a question that makes the reader answer in their head.",
  beats: [
    "SAY: Watch me. Fed up with watching perfectly good paper hit the bin?",
    [
      "ASK: Now finish a DIFFERENT starter, same topic.",
      "60 sec. Write it... chin it... show me.",
      "EXPECT: a completed starter ending in a question mark.",
    ],
    [
      "SCAN boards for a real question.",
      "80%+ -> cold call one strong, one shaky: what will they answer?",
      "Less -> finish one starter together, re-ask.",
    ],
  ],
  trap: [
    "asking something answerable either way.",
    "Fix: ask which answer they wanted, student narrows it.",
  ],
  stretch: "Add a bossy verb straight after your question.",
  help: "Choose the starter, leave only the ending blank.",
  prep: "Same topic as the last two sessions on purpose. Familiar topic, so all the thinking goes on the hook.",
  tag: "[We Do | Supported application | SC2 | HITS 4, 5]",
});

const NOTES_YOUDO = composeGlanceNotes({
  answer: "open - three hooks, one using each tool.",
  beats: [
    [
      "SAY: New topic. A quiet zone in the yard for anyone who wants one.",
      "SAY: Three hooks. One question, one bossy verb, one promise.",
    ],
    [
      "POINT to the Hook Lab. SAY: One tool per box. Use your bank.",
      "TIME: twelve minutes.",
    ],
    [
      "CIRCULATE. Check the tool matches the box before the wording.",
      "COLLECT two hooks to read aloud at the closing.",
    ],
  ],
  trap: [
    "writing three versions of the same question.",
    "Fix: point at the box label, student switches tool.",
  ],
  stretch: "Write a fourth hook that uses two tools in one line.",
  help: "Give the starter word for each box so only the ending is theirs.",
  prep: "New topic on purpose, so this is transfer and not a repeat of the We Do.",
  tag: "[You Do | Mastery and application | SC2 | HITS 10]",
});

const NOTES_EXIT = composeGlanceNotes({
  answer: "any question that makes the reader take a side on homework.",
  beats: [
    [
      "SAY: One rhetorical question. Topic: homework should be shorter.",
      "SAY: Make me answer it in my head.",
    ],
    [
      "TIME: three minutes. Sheet on the desk when you finish.",
      "COLLECT by walking the rows.",
    ],
    [
      "SORT as you collect: real hook, or statement with a question mark.",
      "A thick second pile -> reopen the starters at the top of Session 4.",
    ],
  ],
  prep: "Assesses the core target: write a hook using one named tool. The question mark alone is not the evidence.",
  tag: "[Exit Ticket | Mastery and application | SC2 | HITS 8]",
});

const NOTES_CLOSING = composeGlanceNotes({
  answer: "open - most students should show two.",
  beats: [
    "READ the two collected hooks aloud. SAY: Listen for what it makes you answer.",
    [
      "ASK: Rate yourself on the three statements.",
      "15 sec. Fingers at your chest. One, two or three. Show me.",
      "EXPECT: most on two.",
    ],
    "SAY: Next session we build the middle, where the reasons and the proof live.",
  ],
  prep: "Anyone showing one starts Session 4 with the starter list in front of them.",
  tag: "[Closing | Retention and recall | HITS 1, 9]",
});

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  titleSlide(pres, "Hook Them In",
    "Voice, part 2: questions, bossy verbs and promises",
    "Week 6 Session 3 | Year 5/6 Literacy", NOTES_TITLE);

  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  (() => {
    const s = P.customSlide(pres, T, "For the teacher", C.MUTED, "Session 3 at a glance", { badgeW: 2.4 });
    const cardY = CONTENT_TOP;
    const cardH = SAFE_BOTTOM - CONTENT_TOP;

    addCard(s, 0.5, cardY, 4.4, cardH, { strip: C.PRIMARY });
    s.addText("Where this sits", {
      x: 0.7, y: cardY + 0.12, w: 4.0, h: 0.28,
      fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText([
      { text: "Still inside the VOICE band of the anchor.", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: P.ANCHOR_PHRASE, options: { color: C.PRIMARY, bold: true, breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "Three tools today:", options: { bold: true, breakLine: true } },
      { text: "rhetorical questions (make them take a side)", options: { breakLine: true } },
      { text: "bossy verbs (tell them what to do)", options: { breakLine: true } },
      { text: "exaggeration and promises (what they get)", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "The bossy verb comes back in Week 7 as the poster's call to action.", options: { italic: true, color: C.MUTED } },
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
      { text: "Shape: example-first. The ARC sample text opens with Imagine, a bossy verb straight off the printed bank.", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "Decision points:", options: { bold: true, color: C.ACCENT, breakLine: true } },
      { text: "the keep-reading hinge, the boards check in the We Do, and the exit ticket.", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "Sensitive content:", options: { bold: true, color: C.ACCENT, breakLine: true } },
      { text: "two lines of the cyberbullying sample appear in the I Do. Name it as the example topic and move on.", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "Flex: cut the promises column if the launch runs long.", options: { italic: true, color: C.MUTED } },
    ], {
      x: 5.3, y: cardY + 0.46, w: 4.0, h: cardH - 0.60,
      fontSize: 12.5, fontFace: FONT_B, color: C.CHARCOAL,
      valign: "top", margin: 0, paraSpaceAfter: 3,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_OVERVIEW);
    runSlideDiagnostics(s, pres);
  })();

  // Launch - retrieval swap, then the problem
  (() => {
    const s = P.customSlide(pres, T, "Launch", C.SECONDARY, "Warm up, then a problem");
    const y0 = CONTENT_TOP;

    addCard(s, 0.5, y0, 9, 1.05, { strip: C.SECONDARY });
    s.addText("Upgrade one word", {
      x: 0.75, y: y0 + 0.09, w: 4, h: 0.28,
      fontSize: 14, fontFace: FONT_B, color: C.SECONDARY, bold: true, margin: 0,
    });
    s.addText([
      { text: "Throwing away good paper is ", options: { color: C.CHARCOAL } },
      { text: "bad", options: { color: C.ALERT, bold: true } },
      { text: ".", options: { color: C.CHARCOAL } },
    ], {
      x: 0.75, y: y0 + 0.40, w: 8.5, h: 0.56,
      fontSize: 24, fontFace: FONT_B, valign: "middle", margin: 0,
    });

    s.addText("But a strong word is wasted if nobody reads past line one.", {
      x: 0.5, y: y0 + 1.32, w: 9, h: 0.90,
      fontSize: 27, fontFace: FONT_H, color: C.PRIMARY, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    addTextOnShape(s, "Write it... chin it... show me.", {
      x: 2.95, y: y0 + 2.36, w: 4.1, h: 0.62, rectRadius: 0.08,
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
    "We are learning how a writer hooks a reader and pushes them to act.",
    [
      "I can spot a rhetorical question in a piece of writing.",
      "I can write a hook using a rhetorical question or a bossy verb.",
      "I can explain why my hook works on my reader.",
    ],
    NOTES_LI, FOOTER);

  keyWordSlide(pres, {
    word: "rhetorical question",
    meaning: "A question you ask to make the reader think, not to get an answer.",
    example: "Would you like more homework tonight?",
    routine: ["Say it", "Answer it in your head", "Use it"],
    color: C.SECONDARY,
    title: "The word for today",
  }, NOTES_VOCAB, FOOTER);

  // I Do - the ARC opening, two tools labelled
  (() => {
    const s = P.customSlide(pres, T, "I Do", C.PRIMARY, "How a real writer opens");
    const y0 = CONTENT_TOP;

    addCard(s, 0.5, y0, 9, 1.98, { strip: C.PRIMARY });
    s.addText("The opening of the Level 5-6 sample text", {
      x: 0.75, y: y0 + 0.09, w: 8.5, h: 0.28,
      fontSize: 14, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText(P.ARC.hook, {
      x: 0.75, y: y0 + 0.42, w: 8.5, h: 0.62,
      fontSize: 19, fontFace: FONT_H, color: C.CHARCOAL, italic: true,
      valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
    });
    s.addText(P.ARC.rhetorical, {
      x: 0.75, y: y0 + 1.14, w: 8.5, h: 0.68,
      fontSize: 19, fontFace: FONT_H, color: C.CHARCOAL, italic: true,
      valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
    });

    const ly = y0 + 2.20;
    clickBuild(s, [
      () => {
        addTextOnShape(s, "Imagine = a bossy verb", {
          x: 0.5, y: ly, w: 4.4, h: 0.82, rectRadius: 0.08,
          fill: { color: C.SECONDARY },
        }, {
          fontSize: 19, fontFace: FONT_B, color: C.WHITE, bold: true,
          align: "center", valign: "middle", margin: 0,
        });
      },
      () => {
        addTextOnShape(s, "Line 2 = a rhetorical question", {
          x: 5.1, y: ly, w: 4.4, h: 0.82, rectRadius: 0.08,
          fill: { color: C.ASSESS },
        }, {
          fontSize: 19, fontFace: FONT_B, color: C.WHITE, bold: true,
          align: "center", valign: "middle", margin: 0,
        });
      },
    ]);

    s.addText("It is on your printed bank. A real writer used it first.", {
      x: 0.5, y: ly + 0.94, w: 9, h: 0.40,
      fontSize: 15, fontFace: FONT_B, color: C.MUTED, italic: true,
      align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO);
    runSlideDiagnostics(s, pres);
  })();

  // Rhetorical question starters
  (() => {
    const s = P.customSlide(pres, T, "Tool Bank", C.ASSESS, "Ready-made openings");
    const y0 = CONTENT_TOP;
    const endY = P.drawWordChips(s, T, {
      x: 0.5, y: y0, w: 9, words: P.RHETORICAL_STARTERS, perRow: 3,
      chipH: 0.58, gap: 0.11, fill: C.BG_LIGHT, fontSize: 16,
    });

    addTextOnShape(s, "Pick one. Finish it about wasted paper.", {
      x: 0.5, y: endY + 0.26, w: 9, h: 0.70, rectRadius: 0.08,
      fill: { color: C.ASSESS },
    }, {
      fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_QUESTIONS);
    runSlideDiagnostics(s, pres);
  })();

  // Bossy verbs and promises
  (() => {
    const s = P.customSlide(pres, T, "Tool Bank", C.SECONDARY, "Order them, or promise them");
    const y0 = CONTENT_TOP;
    const colW = 4.4;

    addTextOnShape(s, "Bossy verbs: give an order", {
      x: 0.5, y: y0, w: colW, h: 0.56, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, {
      fontSize: 16, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
    P.drawWordChips(s, T, {
      x: 0.5, y: y0 + 0.68, w: colW, words: P.BOSSY_VERBS, perRow: 2,
      chipH: 0.62, gap: 0.12, fill: C.WHITE, fontSize: 17,
    });

    addTextOnShape(s, "Promises: say what they get", {
      x: 5.1, y: y0, w: colW, h: 0.56, rectRadius: 0.08,
      fill: { color: C.ASSESS },
    }, {
      fontSize: 16, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
    P.drawWordChips(s, T, {
      x: 5.1, y: y0 + 0.68, w: colW, words: P.EXAGGERATION_PROMISES.slice(0, 6), perRow: 2,
      chipH: 0.62, gap: 0.12, fill: C.WHITE, fontSize: 15,
    });

    addTextOnShape(s, "An excellent ending tells the reader what to DO.", {
      x: 0.5, y: y0 + 3.06, w: 9, h: 0.66, rectRadius: 0.08,
      fill: { color: C.WHITE }, line: { color: C.PRIMARY, width: 1.5 },
    }, {
      fontSize: 18, fontFace: FONT_B, color: C.PRIMARY, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_VERBS);
    runSlideDiagnostics(s, pres);
  })();

  // CFU hinge
  (() => {
    const s = P.customSlide(pres, T, "CFU", C.ALERT, "Which one makes you keep reading?");
    const y0 = CONTENT_TOP;

    addTextOnShape(s, "Topic: our school should do more sport", {
      x: 0.5, y: y0, w: 9, h: 0.48, rectRadius: 0.08,
      fill: { color: C.ALERT },
    }, {
      fontSize: 16, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    const options = [
      { key: "A", text: "This piece is going to be about sport at our school." },
      { key: "B", text: "Fed up with sitting still all day? Our school must double its sport." },
      { key: "C", text: "Sport is quite popular among many students in our school." },
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
        addRevealAnswerBar(s, ["B - a question, then a must"], {
          y: P.REVEAL_Y, h: P.REVEAL_H, fontSize: 22, color: C.SUCCESS,
        });
      },
    ]);

    addFooter(s, FOOTER);
    s.addNotes(NOTES_CFU);
    runSlideDiagnostics(s, pres);
  })();

  // We Do
  (() => {
    const s = P.customSlide(pres, T, "We Do", C.SUCCESS, "Finish a starter together");
    const y0 = CONTENT_TOP;

    addTextOnShape(s, P.MODEL_TOPIC, {
      x: 0.5, y: y0, w: 9, h: 0.72, rectRadius: 0.08,
      fill: { color: C.SUCCESS },
    }, {
      fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    const my = y0 + 0.88;
    addCard(s, 0.5, my, 9, 1.25, { strip: C.PRIMARY });
    s.addText("Mine", {
      x: 0.75, y: my + 0.09, w: 3, h: 0.28,
      fontSize: 14, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText([
      { text: "Fed up with ", options: { color: C.SECONDARY, bold: true } },
      { text: "watching perfectly good paper hit the bin?", options: { color: C.CHARCOAL } },
    ], {
      x: 0.75, y: my + 0.42, w: 8.5, h: 0.72,
      fontSize: 23, fontFace: FONT_B, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    addTextOnShape(s, "Now finish a DIFFERENT starter. Write it... chin it... show me.", {
      x: 0.5, y: my + 1.42, w: 9, h: 0.80, rectRadius: 0.08,
      fill: { color: C.WHITE }, line: { color: C.SUCCESS, width: 1.5 },
    }, {
      fontSize: 19, fontFace: FONT_B, color: C.SUCCESS, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_WEDO);
    runSlideDiagnostics(s, pres);
  })();

  // You Do - Hook Lab
  (() => {
    const s = P.customSlide(pres, T, "You Do", C.ASSESS, "Three hooks, three tools");
    const y0 = CONTENT_TOP;

    addTextOnShape(s, "New topic: our yard should have a quiet zone", {
      x: 0.5, y: y0, w: 9, h: 0.72, rectRadius: 0.08,
      fill: { color: C.ASSESS },
    }, {
      fontSize: 23, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    const boxes = [
      { label: "Hook 1", tool: "a rhetorical question", color: C.PRIMARY },
      { label: "Hook 2", tool: "a bossy verb", color: C.SECONDARY },
      { label: "Hook 3", tool: "a promise", color: C.SUCCESS },
    ];
    const bw = 2.87;
    const gap = 0.20;
    boxes.forEach((box, i) => {
      const bx = 0.5 + i * (bw + gap);
      s.addShape("roundRect", {
        x: bx, y: y0 + 0.90, w: bw, h: 1.90, rectRadius: 0.08,
        fill: { color: C.WHITE }, line: { color: box.color, width: 1.5 },
      });
      s.addShape("roundRect", {
        x: bx, y: y0 + 0.90, w: bw, h: 0.52, rectRadius: 0.08,
        fill: { color: box.color },
      });
      s.addText(box.label, {
        x: bx, y: y0 + 0.90, w: bw, h: 0.52,
        fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      s.addText(box.tool, {
        x: bx + 0.14, y: y0 + 1.52, w: bw - 0.28, h: 1.20,
        fontSize: 17, fontFace: FONT_B, color: C.CHARCOAL,
        align: "center", valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
    });

    addTextOnShape(s, "Hook Lab sheet. Twelve minutes. One tool per box.", {
      x: 0.5, y: y0 + 2.98, w: 9, h: 0.72, rectRadius: 0.08,
      fill: { color: C.WHITE }, line: { color: C.ASSESS, width: 1.5 },
    }, {
      fontSize: 18, fontFace: FONT_B, color: C.ASSESS, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
    runSlideDiagnostics(s, pres);
  })();

  P.exitTicketPanel(pres, T, {
    topic: "Homework should be shorter",
    task: "Write ONE rhetorical question to hook your reader.",
    cue: "Make me answer it in my head.",
    footer: FOOTER,
    notes: NOTES_EXIT,
  });

  closingSlide(pres, {
    reflectionPrompt: "Which of the three tools felt most like you?",
    scItems: [
      "I can spot a rhetorical question in a piece of writing.",
      "I can write a hook using a rhetorical question or a bossy verb.",
      "I can explain why my hook works on my reader.",
    ],
    selfAssessment: "Fingers at your chest. One, two or three. Show me.",
    takeaways: ["A hook makes the reader answer before they have read a word."],
  }, NOTES_CLOSING);

  const pptxPath = path.join(OUT_DIR, "Persuasive Writing W6 S3.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);
}

// ─── Resource: Hook Lab ─────────────────────────────────────────────────────

async function buildLab() {
  const doc = createPdf({ title: LAB_RES.name });
  let y = addPdfHeader(doc, "Hook Lab", {
    subtitle: "Three hooks for one topic. One tool in each box.",
    color: C.ASSESS,
    lessonInfo: "Week 6 Session 3 | Year 5/6 Literacy",
  });

  y = addSectionHeading(doc, "Watch two first", y, { color: C.ASSESS });
  y += 4;

  // Worked hooks, drawn as labelled boxes so the printed page matches the
  // structure students are asked to fill in below.
  const worked = [
    { tool: "Rhetorical question", text: "Fed up with watching perfectly good paper hit the bin?", color: C.PRIMARY },
    { tool: "Bossy verb", text: "Imagine a classroom where nothing useful gets thrown away.", color: C.SECONDARY },
  ];
  worked.forEach((w) => {
    doc.save();
    doc.roundedRect(PAGE.MARGIN, y, PAGE.CONTENT_W, 40, 3)
      .lineWidth(0.8).strokeColor(hex(w.color)).stroke();
    doc.roundedRect(PAGE.MARGIN, y, 120, 40, 3).fill(hex(w.color));
    doc.fontSize(9).font("Sans-Bold").fillColor("#FFFFFF");
    doc.text(w.tool, PAGE.MARGIN + 6, y + 15, { width: 108, align: "center" });
    doc.fontSize(11).font("Sans-Italic").fillColor("#000000");
    doc.text(w.text, PAGE.MARGIN + 130, y + 13, { width: PAGE.CONTENT_W - 140 });
    doc.restore();
    y += 46;
  });

  y = addBodyText(doc,
    "Both hooks make the reader do something before they have read the argument: "
    + "answer a question, or picture a scene.",
    y, { italic: true });
  y += 8;

  y = addSectionHeading(doc, "Your topic: our yard should have a quiet zone", y, { color: C.ASSESS });
  y += 4;

  const boxes = [
    { label: "Hook 1: a rhetorical question", hint: "Use a starter from your word bank.", color: C.PRIMARY },
    { label: "Hook 2: a bossy verb", hint: "Don't... Go on... Try a... Imagine... Enjoy the...", color: C.SECONDARY },
    { label: "Hook 3: a promise", hint: "You will be... It will... Just think how... Now you can...", color: C.SUCCESS },
  ];

  boxes.forEach((box) => {
    doc.save();
    doc.roundedRect(PAGE.MARGIN, y, PAGE.CONTENT_W, 22, 3).fill(hex(box.color));
    doc.fontSize(10).font("Sans-Bold").fillColor("#FFFFFF");
    doc.text(box.label, PAGE.MARGIN + 8, y + 6, { width: PAGE.CONTENT_W - 16 });
    doc.restore();
    y += 26;
    doc.fontSize(9).font("Sans-Italic").fillColor(hex(C.MUTED));
    doc.text(box.hint, PAGE.MARGIN + 2, y, { width: PAGE.CONTENT_W - 4 });
    y += 14;
    y = addLinedArea(doc, y, 2, { lineSpacing: 26 });
    y += 12;
  });

  y = addTipBox(doc,
    "Challenge: write a fourth hook that uses two tools in one line, then read all four aloud "
    + "and circle the one you would actually use.",
    y, { color: C.ACCENT });

  addPdfFooter(doc, "Persuasive Writing | Week 6 Session 3 | Year 5/6 Literacy");
  await writePdf(doc, path.join(OUT_DIR, LAB_RES.fileName));
  console.log("Wrote " + LAB_RES.name);
}

(async () => {
  await build();
  await buildLab();
})().catch((err) => {
  console.error(err);
  process.exit(1);
});
