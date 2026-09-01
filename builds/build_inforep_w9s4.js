"use strict";

// Information Reports - Term 3 Week 9, Session 4 (Year 5/6 Literacy)
// "The words that make it sound like a report"
//
// Victorian Curriculum 2.0: English, Language, Levels 5-6 - VC2E6LA06, how
// the choice of verb and elaborated tenses expands and sharpens ideas, plus
// technical vocabulary and expanded noun groups for precision.
//
// UNIT ANCHOR (locked): "Classify it. Describe it, one aspect at a time.
// Wrap it up. Facts all the way through." Today is the FACTS band.
//
// Lesson shape: error-analysis-led. Every teaching move starts from a sentence
// that is wrong in a named way (opinion, past tense, vague noun) and fixes it,
// because these are the four errors that will appear in next week's drafts.
//
// Sources: teacher's unit plan. Example sentences are original.

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme } = require("../themes/factory");
const { composeGlanceNotes } = require("../themes/core/composeNotes");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addWriteLine,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
  PAGE, hex,
} = require("../themes/pdf_helpers");

const P = require("./inforep_lib");

const UNIT_VARIANT = 2;
const T = createTheme("literacy", "grade56", UNIT_VARIANT);
const {
  C, FONT_H, FONT_B, SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, keyWordSlide, closingSlide,
  addFooter, addTextOnShape, addRevealAnswerBar,
  clickBuild, runSlideDiagnostics,
} = T;

const SESSION = 4;
const FOOTER = "Information Reports | Week 9 Session 4 | Year 5/6 Literacy";
const OUT_DIR = path.join(__dirname, "..", "output", "InfoReport_W9_S4");
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const SHEET_RES = makeSessionResource(SESSION,
  "Language Features Sheet",
  "Sort fact and opinion, fix the tense, grow noun groups, add an adverb group.");
const KEY_RES = makeSessionResource(SESSION,
  "Answer Key",
  "Answers for the language features sheet, with what to accept.");
const RESOURCE_ITEMS = [SHEET_RES, KEY_RES];
fs.mkdirSync(RES_DIR, { recursive: true });

/* --- Teacher notes (Glance Format, megaprompt sections 45-47) ------------- */

const NOTES_TITLE =
  "Session 4, the last of Week 9. We know the parts and we can find facts. Today we work on the words themselves.";

const NOTES_RESOURCES =
  "Prep slide. Print the Language Features Sheet, one per student. Keep the Answer Key for yourself.\n" +
  "Students need their map card. Whiteboards out for the We Do.\n" +
  "CATCH-UP: missed a session? Today needs nothing except the map card. Section 1 of the sheet is the same sort we model together.";

const NOTES_LAUNCH = composeGlanceNotes({
  answer: "B. You can check it, and it uses the present tense.",
  beats: [
    [
      "SAY: Two sentences, same frog.",
      "Only one belongs in a report.",
    ],
    [
      "ASK: Which one belongs in a report?",
      "20 sec. Fingers at your chest: one or two.",
      "EXPECT: two.",
    ],
    [
      "SCAN the fingers.",
      "80%+ -> reveal, then move on.",
      "Less -> read A again, ask who could argue, re-ask.",
    ],
    [
      "REVEAL after the fingers are up.",
      "SAY: I reckon is the giveaway. Reports never say I.",
    ],
  ],
  trap: [
    "thinking a strong sentence is a factual one.",
    "Fix: ask could we check it, student re-sorts.",
  ],
  prep: "Low-coupling launch: two sentences on screen, nothing earlier assumed. Whole block under 5 minutes.",
  tag: "[Launch | Retention and recall | HITS 6]",
});

const NOTES_LI = composeGlanceNotes({
  beats: [
    [
      "POINT to the learning intention.",
      "SAY: Today we work on the words that make it sound like a report.",
    ],
    "SAY: Read the I can statements with me. Everyone, together, on three.",
    "SAY: The first one everyone leaves with. Tell a fact from an opinion.",
  ],
  prep: "SC1 is reachable by every student; SC2 is the core target the exit ticket assesses; SC3 stretches into noun groups. Tier labels stay off the slide.",
  tag: "[LI/SC | Planning made visible | HITS 1]",
});

const NOTES_VOCAB_TECH = composeGlanceNotes({
  answer: "a technical word is the exact word experts use for a topic.",
  beats: [
    [
      "SAY: Every topic has its own exact words.",
      "Amphibian. Habitat. Nocturnal.",
    ],
    [
      "ASK: What is the technical word for active at night?",
      "10 sec. Choral response: everyone, together, on three.",
      "EXPECT: nocturnal.",
    ],
    "SAY: One technical word can replace a whole explaining sentence.",
  ],
  trap: [
    "using a hard word without knowing it.",
    "Fix: ask for the meaning, student says it plainly.",
  ],
  prep: "Technical words are precision, not decoration. Students add them to the class glossary as they research next week.",
  tag: "[Key Vocabulary | Knowledge and memory | HITS 6]",
});

const NOTES_VOCAB_NOUN = composeGlanceNotes({
  answer: "a noun group is a noun plus the words that describe it.",
  beats: [
    [
      "SAY: A noun group is the noun and its describing words.",
      "The frog becomes the large, bright green tree frog.",
    ],
    [
      "ASK: Grow this one. The pads.",
      "20 sec. Turn and tell, partner A first.",
      "EXPECT: the wide, round pads on its toes.",
    ],
    "SAY: More detail means a clearer picture for the reader.",
  ],
  trap: [
    "adding words that describe nothing new.",
    "Fix: ask what the reader now knows, student cuts a word.",
  ],
  prep: "The precision tool for the Describe band. It is the difference between a list of facts and a report a reader can picture.",
  tag: "[Key Vocabulary | Knowledge and memory | HITS 6]",
});

const NOTES_IDO = composeGlanceNotes({
  answer: "cut the opinion, then use the present tense.",
  beats: [
    [
      "SAY: Two sentences that nearly work.",
      "Watch me fix each one.",
    ],
    [
      "CLICK to the fix. SAY: Best is an opinion, so out it goes.",
      "Now anyone can check it.",
    ],
    [
      "CLICK again. SAY: Yesterday makes it a story.",
      "Frogs eat moths. True every day.",
    ],
    [
      "ASK: Which word had to go in the first fix?",
      "15 sec. Write it... chin it... show me.",
      "EXPECT: best.",
    ],
  ],
  trap: [
    "changing the whole sentence instead of one word.",
    "Fix: ring the guilty word, student fixes only that.",
  ],
  stretch: "Find another opinion word that could sneak in.",
  help: "Give the fix with one blank in it.",
  prep: "Say the anchor once here, on the FACTS band. These two fixes are the two errors that dominate next week's drafts.",
  tag: "[I Do | Explicit teaching | SC1 | HITS 3, 4]",
});

const NOTES_CFU = composeGlanceNotes({
  answer: "B. Ten centimetres can be measured.",
  beats: [
    [
      "SAY: Three sentences. All sound factual. Only one is.",
      "Do not call out. This one goes on boards.",
    ],
    [
      "ASK: Which sentence is a fact?",
      "30 sec. Boards up on cue.",
      "EXPECT: B.",
    ],
    [
      "SCAN boards, back row first.",
      "80%+ -> cold call one B: how would you check it?",
      "Less -> read A, ask who decides most useful, re-ask.",
    ],
    [
      "REVEAL after boards are scanned.",
      "SAY: Most useful and amazing are judgements, not facts.",
    ],
  ],
  trap: [
    "picking C because climbing really is true.",
    "Fix: ring the word amazing, student re-tests.",
  ],
  prep: "The hinge of the lesson. A and C both hide a judgement inside a factual-sounding sentence, which is exactly the error in student drafts.",
  tag: "[CFU hinge | Supported application | SC2 | HITS 7, 8]",
});

const NOTES_WEDO = composeGlanceNotes({
  answer: "open - listen for two describing words and a where phrase.",
  beats: [
    [
      "SAY: Watch a noun group grow. The frog. Then more detail.",
      "Each click adds something the reader can picture.",
    ],
    [
      "ASK: Now grow this one. Its tongue.",
      "60 sec. Write it... chin it... show me.",
      "EXPECT: its long, sticky tongue.",
    ],
    [
      "SCAN boards for two describing words.",
      "80%+ -> cold call one: what can we picture now?",
      "Less -> grow one together on the board, re-ask.",
    ],
  ],
  trap: [
    "stacking words that add nothing, like the very big huge frog.",
    "Fix: ask which word to cut, student cuts it.",
  ],
  stretch: "Add a where phrase: on the smooth glass.",
  help: "Give one describing word. Ask for the second.",
  prep: "Different move from the You Do: growing one phrase together before they meet four alone. Keep the growth visible on the board.",
  tag: "[We Do | Supported application | SC3 | HITS 4, 5]",
});

const NOTES_YOUDO = composeGlanceNotes({
  answer: "three facts, three opinions, three tense fixes, four noun groups.",
  beats: [
    [
      "SAY: Four short jobs on the sheet. Same four tools.",
      "Sort, fix, grow, then the stretch box.",
    ],
    [
      "POINT to the steps. TIME: fifteen minutes.",
      "Section 1 first, it is the sort we just did.",
    ],
    [
      "CIRCULATE. Look at the tense fixes first, they are hardest.",
      "COLLECT two noun groups to read at the closing.",
    ],
  ],
  trap: [
    "fixing the tense but leaving yesterday in the sentence.",
    "Fix: ask if it is true every day, student cuts the time word.",
  ],
  stretch: "Section 4: add a where, when or how phrase.",
  help: "Section 3 has the first describing word filled in.",
  prep: "Four short tasks rather than one long one, so a student who stalls on tense still finishes noun groups. Section 1 re-grounds today's sort.",
  tag: "[You Do | Mastery and application | SC2 | HITS 10]",
});

const NOTES_EXIT = composeGlanceNotes({
  answer: "any present-tense sentence about frogs with the opinion removed.",
  beats: [
    [
      "SAY: One sentence. Two things wrong with it.",
      "Take out the opinion. Put it in the present tense.",
    ],
    [
      "TIME: three minutes. Books open on the desk when you finish.",
      "COLLECT by walking the rows, not by calling names.",
    ],
    [
      "SORT into two piles as you collect: both fixed, one fixed.",
      "A thick second pile -> re-model the tense fix in Week 10.",
    ],
  ],
  prep: "Assesses the core target: factual language in the timeless present. Two errors in one sentence, so the pile sort tells you which one to re-teach.",
  tag: "[Exit Ticket | Mastery and application | SC2 | HITS 8]",
});

const NOTES_CLOSING = composeGlanceNotes({
  answer: "open - most students should show two.",
  beats: [
    "READ the two collected noun groups aloud. SAY: Picture it as I read.",
    [
      "ASK: Rate yourself on the three statements.",
      "15 sec. Fingers at your chest, one, two or three. Show me.",
      "EXPECT: most on two.",
    ],
    "SAY: Next week you choose your own topic and write the whole thing.",
  ],
  prep: "End of Week 9. Self-assessment feeds the Week 10 planning groups. Anyone showing one starts Week 10 beside you.",
  tag: "[Closing | Retention and recall | HITS 1, 9]",
});

/* --- Slide content -------------------------------------------------------- */

const LAUNCH_PAIR = [
  { key: "1", text: "I reckon green tree frogs are the coolest animals you will ever see." },
  { key: "2", text: "Green tree frogs live in the warm, wet parts of northern Australia." },
];

const IDO_ROWS = [
  { label: "Opinion", text: "Green tree frogs are the best pets in Australia.", bad: true },
  { label: "Fact", text: "Green tree frogs are often kept as pets.", bad: false },
  { label: "Past tense", text: "Yesterday the frog ate a moth.", bad: true },
  { label: "Present tense", text: "Green tree frogs eat moths.", bad: false },
];

const CFU_OPTIONS = [
  { key: "A", text: "Green tree frogs are the most useful animals in a garden." },
  { key: "B", text: "Green tree frogs can grow to about ten centimetres long." },
  { key: "C", text: "Green tree frogs are amazing climbers." },
];

const NOUN_GROWTH = [
  "the frog",
  "the green frog",
  "the large green frog",
  "the large, bright green tree frog",
];

/* --- Build ---------------------------------------------------------------- */

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // 1. Title
  titleSlide(pres, "Sounding Like a Report",
    "Facts. Present tense. Technical words. Noun groups.",
    "Week 9 Session 4 | Year 5/6 Literacy", NOTES_TITLE);

  // 2. Teacher Resources
  addResourceSlide(pres, {
    resources: RESOURCE_ITEMS,
    studentTools: [
      "Mini-whiteboards and markers",
      "Writing books and pencils",
      "The map card",
    ],
    boardSetup: [
      "Draw two columns headed Fact and Opinion",
      "Leave room to grow a noun group live",
    ],
  }, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // 3. Launch - which one belongs in a report?
  (() => {
    const s = P.customSlide(pres, T, "Launch", C.SECONDARY, "Which one belongs in a report?");
    const y0 = CONTENT_TOP;
    const cardH = 1.28;
    const gap = 0.18;

    LAUNCH_PAIR.forEach((item, i) => {
      const cy = y0 + i * (cardH + gap);
      s.addShape("roundRect", {
        x: 0.5, y: cy, w: 9, h: cardH, rectRadius: 0.08,
        fill: { color: C.WHITE }, line: { color: C.SECONDARY, width: 1.5 },
      });
      s.addShape("roundRect", {
        x: 0.72, y: cy + (cardH - 0.52) / 2, w: 0.52, h: 0.52, rectRadius: 0.26,
        fill: { color: C.SECONDARY },
      });
      s.addText(item.key, {
        x: 0.72, y: cy + (cardH - 0.52) / 2, w: 0.52, h: 0.52,
        fontSize: 19, fontFace: FONT_H, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      s.addText(item.text, {
        x: 1.44, y: cy, w: 7.8, h: cardH,
        fontSize: 21, fontFace: FONT_B, color: C.CHARCOAL,
        valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
    });

    clickBuild(s, [
      () => {
        addRevealAnswerBar(s, ["2 - you can check it, and it is true every day"], {
          y: P.REVEAL_Y, h: P.REVEAL_H, fontSize: 20, color: C.PRIMARY,
        });
      },
    ]);

    addFooter(s, FOOTER);
    s.addNotes(NOTES_LAUNCH);
    runSlideDiagnostics(s, pres);
  })();

  // 4. Learning Intention and Success Criteria
  liSlide(pres,
    "We are learning the language that makes writing sound like an information report.",
    [
      "I can tell a fact from an opinion.",
      "I can write about my topic in the present tense.",
      "I can add a technical word and grow a noun group to give more detail.",
    ],
    NOTES_LI, FOOTER);

  // 5. Key vocabulary - technical word
  keyWordSlide(pres, {
    word: "technical word",
    meaning: "The exact word experts use for a topic.",
    example: "Nocturnal means active at night. One word does the whole job.",
    routine: ["Say it", "Say the meaning", "Use it"],
    color: C.PRIMARY,
    title: "Words for precision",
  }, NOTES_VOCAB_TECH, FOOTER);

  // 6. Key vocabulary - noun group
  keyWordSlide(pres, {
    word: "noun group",
    meaning: "A noun and the words that describe it.",
    example: "The frog becomes the large, bright green tree frog.",
    routine: ["Say it", "Grow one", "Use it"],
    color: C.SECONDARY,
    title: "Words for detail",
  }, NOTES_VOCAB_NOUN, FOOTER);

  // 7. I Do - fix the opinion, fix the tense
  (() => {
    const s = P.customSlide(pres, T, "I Do", C.PRIMARY, "Two sentences that nearly work");
    const y0 = CONTENT_TOP;
    const rowH = 0.78;
    const gap = 0.14;
    const bigGap = 0.30;

    // Rows 0 and 2 are the flawed sentences and are visible from the start.
    // Rows 1 and 3 are the fixes and arrive on a click each.
    const tops = [
      y0,
      y0 + rowH + gap,
      y0 + 2 * rowH + gap + bigGap,
      y0 + 3 * rowH + 2 * gap + bigGap,
    ];

    function drawRow(i) {
      const row = IDO_ROWS[i];
      const accent = row.bad ? C.ALERT : C.SUCCESS;
      s.addShape("roundRect", {
        x: 0.5, y: tops[i], w: 9, h: rowH, rectRadius: 0.08,
        fill: { color: C.WHITE }, line: { color: accent, width: 1.5 },
      });
      addTextOnShape(s, row.label, {
        x: 0.5, y: tops[i], w: 1.85, h: rowH, rectRadius: 0.08,
        fill: { color: accent },
      }, {
        fontSize: 13.5, fontFace: FONT_B, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      s.addText(row.text, {
        x: 2.55, y: tops[i], w: 6.7, h: rowH,
        fontSize: 19, fontFace: FONT_B, color: C.CHARCOAL,
        valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
    }

    drawRow(0);
    drawRow(2);
    clickBuild(s, [
      () => drawRow(1),
      () => drawRow(3),
    ]);

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO);
    runSlideDiagnostics(s, pres);
  })();

  // 8. CFU hinge - which one is a fact?
  (() => {
    const s = P.customSlide(pres, T, "CFU", C.ALERT, "Which sentence is a fact?");
    const y0 = CONTENT_TOP;

    addTextOnShape(s, "A fact can be checked. An opinion cannot.", {
      x: 0.5, y: y0, w: 9, h: 0.48, rectRadius: 0.08,
      fill: { color: C.ALERT },
    }, {
      fontSize: 16, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    P.drawOptionStack(s, T, CFU_OPTIONS, {
      y: y0 + 0.60, optionH: 0.66, gap: 0.12, color: C.PRIMARY, fontSize: 17,
    });

    clickBuild(s, [
      () => {
        addRevealAnswerBar(s, ["B - ten centimetres can be measured"], {
          y: P.REVEAL_Y, h: P.REVEAL_H, fontSize: 21, color: C.SUCCESS,
        });
      },
    ]);

    addFooter(s, FOOTER);
    s.addNotes(NOTES_CFU);
    runSlideDiagnostics(s, pres);
  })();

  // 9. We Do - grow a noun group
  (() => {
    const s = P.customSlide(pres, T, "We Do", C.SUCCESS, "Grow the noun group");
    const y0 = CONTENT_TOP;
    const rowH = 0.60;
    const gap = 0.10;
    const leftW = 5.2;

    // Row 0 is the bare noun and is visible; each click adds one grown version.
    function drawGrowth(i) {
      const gy = y0 + i * (rowH + gap);
      const isLast = i === NOUN_GROWTH.length - 1;
      s.addShape("roundRect", {
        x: 0.5, y: gy, w: leftW, h: rowH, rectRadius: 0.07,
        fill: { color: isLast ? C.SUCCESS : C.WHITE },
        line: isLast ? null : { color: C.MUTED, width: 1.2 },
      });
      s.addText(NOUN_GROWTH[i], {
        x: 0.70, y: gy, w: leftW - 0.40, h: rowH,
        fontSize: isLast ? 18 : 16, fontFace: FONT_B, bold: isLast,
        color: isLast ? C.WHITE : C.CHARCOAL,
        valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
    }

    T.addInstructionCard(s, [
      { text: "Now your turn", role: "header" },
      { text: "Grow this noun group." },
      // The phrase students grow is the hero of this panel, so it is sized
      // explicitly: the emphasis role caps at 15.5 for Y3-6, which would make
      // it smaller than the instructions around it.
      { text: "its tongue", role: "emphasis", fontSize: 28, color: C.SECONDARY },
      { text: "Two describing words at least." },
      { text: "Write it... chin it... show me." },
    ], {
      x: 5.9, y: y0, w: 3.6, h: 4 * rowH + 3 * gap,
      strip: C.SUCCESS, fill: C.WHITE,
      headerColor: C.SUCCESS, emphasisColor: C.SECONDARY,
    });

    drawGrowth(0);
    clickBuild(s, [
      () => drawGrowth(1),
      () => drawGrowth(2),
      () => drawGrowth(3),
      () => {
        addRevealAnswerBar(s, ["its long, sticky tongue"], {
          y: P.REVEAL_Y, h: P.REVEAL_H, fontSize: 24, color: C.SUCCESS,
        });
      },
    ]);

    addFooter(s, FOOTER);
    s.addNotes(NOTES_WEDO);
    runSlideDiagnostics(s, pres);
  })();

  // 10. You Do - the language features sheet
  (() => {
    const s = P.customSlide(pres, T, "You Do", C.ASSESS, "Four short jobs on the sheet");
    const y0 = CONTENT_TOP;
    const h = SAFE_BOTTOM - y0;

    T.addInstructionCard(s, [
      { text: "Language Features Sheet", role: "header" },
      { text: "1. Sort six sentences: fact or opinion." },
      { text: "2. Fix the tense in three sentences." },
      { text: "3. Grow four noun groups." },
      { text: "Fifteen minutes. Map card open." },
    ], {
      x: 0.5, y: y0, w: 4.6, h,
      strip: C.ASSESS, fill: C.WHITE,
      headerColor: C.ASSESS, emphasisColor: C.ALERT,
    });

    const tools = [
      { label: "Facts", detail: "Could anyone check it?", color: C.PRIMARY },
      { label: "Present tense", detail: "True today, true always.", color: C.SECONDARY },
      { label: "Technical words", detail: "The exact word for the topic.", color: C.ACCENT },
      { label: "Noun groups", detail: "The noun plus its describing words.", color: C.SUCCESS },
    ];
    const toolH = (h - 0.36) / 4;
    tools.forEach((tool, i) => {
      const ty = y0 + i * (toolH + 0.12);
      s.addShape("roundRect", {
        x: 5.3, y: ty, w: 4.2, h: toolH, rectRadius: 0.07,
        fill: { color: C.WHITE }, line: { color: tool.color, width: 1.4 },
      });
      s.addShape("roundRect", {
        x: 5.3, y: ty, w: 0.10, h: toolH, rectRadius: 0.03,
        fill: { color: tool.color },
      });
      s.addText([
        { text: tool.label, options: { fontSize: 15, bold: true, color: tool.color, breakLine: true } },
        { text: tool.detail, options: { fontSize: 12.5, color: C.CHARCOAL } },
      ], {
        x: 5.56, y: ty, w: 3.80, h: toolH,
        fontFace: FONT_B, valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
    runSlideDiagnostics(s, pres);
  })();

  // 11. Exit ticket
  P.exitTicketPanel(pres, T, {
    topic: "Frogs were the best animals we saw yesterday.",
    task: "Rewrite this so it belongs in an information report.",
    cue: "Take out the opinion. Use the present tense.",
    taskSize: 27,
    footer: FOOTER,
    notes: NOTES_EXIT,
  });

  // 12. Closing
  closingSlide(pres, {
    reflectionPrompt: "Which of the four tools will you use most in your own report?",
    scItems: [
      "I can tell a fact from an opinion.",
      "I can write about my topic in the present tense.",
      "I can add a technical word and grow a noun group to give more detail.",
    ],
    selfAssessment: "Fingers at your chest. One, two or three. Show me.",
    takeaways: [P.ANCHOR_PHRASE],
  }, NOTES_CLOSING);

  const pptxPath = path.join(OUT_DIR, "Information Reports W9 S4.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);
}

/* --- Resources ------------------------------------------------------------ */

function sheetBody(doc, isKey) {
  let y = addPdfHeader(doc, isKey ? "Language Features: Answer Key" : "Language Features Sheet", {
    subtitle: isKey
      ? "Answers, with what to accept as you circulate."
      : "Sort. Fix the tense. Grow the noun group.",
    color: isKey ? C.ALERT : C.PRIMARY,
    lessonInfo: "Year 5/6 Literacy | Information Reports | Session 4",
    showNameDate: !isKey,
  });

  y = addTipBox(doc, "Facts all the way through. Present tense. Technical words. Noun groups.", y,
    { color: C.ACCENT });

  y = addSectionHeading(doc, "Section 1: fact or opinion?", y,
    { color: isKey ? C.ALERT : C.PRIMARY });
  y = addBodyText(doc, "Write F for fact or O for opinion beside each sentence.", y);

  P.FACT_OPINION.forEach((item) => {
    const textW = PAGE.CONTENT_W - 60;
    const textH = doc.fontSize(11).font("Sans").heightOfString(item.text, { width: textW });
    doc.save();
    if (isKey) {
      doc.roundedRect(PAGE.MARGIN, y - 2, 26, 20, 3).fill(hex(item.fact ? C.SUCCESS : C.ALERT));
      doc.fontSize(11).font("Sans-Bold").fillColor("#FFFFFF");
      doc.text(item.fact ? "F" : "O", PAGE.MARGIN, y + 3, { width: 26, align: "center", lineBreak: false });
    } else {
      doc.roundedRect(PAGE.MARGIN, y - 2, 26, 20, 3).lineWidth(1).strokeColor(hex(C.PRIMARY)).stroke();
    }
    doc.restore();
    doc.fontSize(11).font("Sans").fillColor("#000000");
    doc.text(item.text, PAGE.MARGIN + 38, y, { width: textW });
    y += Math.max(textH, 20) + 10;
  });
  y += 4;

  y = addSectionHeading(doc, "Section 2: fix the tense", y,
    { color: isKey ? C.ALERT : C.SECONDARY });
  y = addBodyText(doc,
    "Rewrite each sentence so it is true every day, not just once.", y);

  P.TENSE_FIXES.forEach((item) => {
    doc.fontSize(11).font("Sans-Italic").fillColor(hex(C.MUTED));
    doc.text(item.before, PAGE.MARGIN + 8, y, { width: PAGE.CONTENT_W - 16 });
    y = doc.y + 4;
    y = addWriteLine(doc, "", y, { answer: isKey ? item.after : null, color: C.SUCCESS });
    // A second ruled line: rewriting a sentence in Year 5/6 handwriting rarely
    // fits on one, and the page has the room.
    if (!isKey) y = addWriteLine(doc, "", y, {});
    y += 6;
  });

  doc.addPage();
  y = PAGE.MARGIN;

  y = addSectionHeading(doc, "Section 3: grow the noun group", y,
    { color: isKey ? C.ALERT : C.ASSESS });
  y = addBodyText(doc,
    "Add describing words so the reader can picture it. The first one is started for you.", y);

  P.NOUN_GROUPS.forEach((item, i) => {
    doc.fontSize(12).font("Sans-Bold").fillColor(hex(C.ASSESS));
    doc.text(item.plain, PAGE.MARGIN + 8, y, { width: 140, lineBreak: false });
    doc.fontSize(11).font("Sans").fillColor(hex(C.MUTED));
    doc.text("becomes", PAGE.MARGIN + 150, y + 1, { width: 60, lineBreak: false });
    doc.save();
    doc.moveTo(PAGE.MARGIN + 210, y + 15).lineTo(PAGE.MARGIN + PAGE.CONTENT_W, y + 15)
      .lineWidth(0.9).strokeColor("#000000").stroke();
    doc.restore();
    // The enabler is on the sheet: item one is worked, so nobody starts blank.
    const shown = isKey ? item.grown : (i === 0 ? item.grown : null);
    if (shown) {
      doc.fontSize(11).font(isKey ? "Sans-Italic" : "Sans").fillColor(hex(isKey ? C.SUCCESS : C.CHARCOAL));
      doc.text(shown, PAGE.MARGIN + 214, y + 1, { width: PAGE.CONTENT_W - 220, lineBreak: false });
    }
    y += 34;
  });
  y += 6;

  y = addSectionHeading(doc, "Section 4: the stretch box", y,
    { color: isKey ? C.ALERT : C.SUCCESS });
  y = addBodyText(doc,
    "Choose one of your noun groups. Add a phrase that says WHERE, WHEN or HOW. Then write the whole sentence.", y);
  if (isKey) {
    y = addBodyText(doc,
      "One way: its long, sticky tongue flicks out at night. Look for a phrase that answers where, when or how, not just another adjective.",
      y, { italic: true, color: "2D6B4A" });
  } else {
    doc.save();
    doc.roundedRect(PAGE.MARGIN, y + 4, PAGE.CONTENT_W, 150, 4)
      .lineWidth(1).strokeColor(hex(C.SUCCESS)).stroke();
    doc.restore();
    for (let i = 1; i <= 4; i += 1) {
      const ly = y + 4 + i * 30;
      doc.moveTo(PAGE.MARGIN + 12, ly).lineTo(PAGE.MARGIN + PAGE.CONTENT_W - 12, ly)
        .lineWidth(0.9).strokeColor("#000000").stroke();
    }
    y += 164;
  }

  y = addSectionHeading(doc, "Word bank: technical words for this topic", y,
    { color: isKey ? C.ALERT : C.PRIMARY });
  P.TECHNICAL_WORDS.forEach((item) => {
    doc.fontSize(11).font("Sans-Bold").fillColor(hex(C.PRIMARY));
    doc.text(item.word, PAGE.MARGIN + 8, y, { width: 100, lineBreak: false });
    doc.font("Sans").fillColor("#000000");
    doc.text(item.meaning, PAGE.MARGIN + 112, y, { width: PAGE.CONTENT_W - 120 });
    y = doc.y + 6;
  });

  addPdfFooter(doc, "Information Reports | Year 5/6 Literacy | Session 4");
}

async function buildSheet() {
  const doc = createPdf({ title: SHEET_RES.name });
  sheetBody(doc, false);
  await writePdf(doc, path.join(OUT_DIR, SHEET_RES.fileName));
  console.log("Wrote " + SHEET_RES.name);
}

async function buildKey() {
  const doc = createPdf({ title: KEY_RES.name });
  sheetBody(doc, true);
  await writePdf(doc, path.join(OUT_DIR, KEY_RES.fileName));
  console.log("Wrote " + KEY_RES.name);
}

(async () => {
  await build();
  await buildSheet();
  await buildKey();
})().catch((err) => {
  console.error(err);
  process.exit(1);
});
