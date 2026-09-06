"use strict";

// Information Reports - Term 3 Week 9, Session 1 (Year 5/6 Literacy)
// "What is an information report?"
//
// Victorian Curriculum 2.0: English, Literacy and Language, Levels 5-6 -
// understanding how texts are structured for purpose and audience, working
// towards VC2E5LY10 and VC2E6LY09 (create texts for purpose and audience).
//
// UNIT ANCHOR (locked, megaprompt sections 5b and 79 - built from the
// teacher's own planning template wording):
//   "Classify it. Describe it, one aspect at a time. Wrap it up.
//    Facts all the way through."
//
// Lesson shape: example-first. Three real openings carry the launch, then the
// page model and its four features carry the I Do. New, high-novelty content,
// so the classic shape is the right one.
//
// Sources: the teacher's unit plan and class topic. The model reports are
// original, written for this unit and fact-checked (see inforep_lib.js).

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme } = require("../themes/factory");
const { composeGlanceNotes } = require("../themes/core/composeNotes");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
  PAGE, hex,
} = require("../themes/pdf_helpers");

const P = require("./inforep_lib");

// Weeks 9 and 10 are ONE information report unit with one locked anchor, so
// both weeks ship on the same palette. Variant 2 is also Week 9's natural
// rotation, and its green/burgundy pair suits the class topic.
const UNIT_VARIANT = 2;
const T = createTheme("literacy", "grade56", UNIT_VARIANT);
const {
  C, FONT_H, FONT_B, SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, keyWordSlide, closingSlide, annotatedModelSlide,
  addCard, addFooter, addTextOnShape, addRevealAnswerBar,
  clickBuild, runSlideDiagnostics,
} = T;

const SESSION = 1;
const FOOTER = "Information Reports | Week 9 Session 1 | Year 5/6 Literacy";
const OUT_DIR = path.join(__dirname, "..", "output", "InfoReport_W9_S1");
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const MAP_RES = makeSessionResource(SESSION,
  "Information Report Map Card",
  "The anchor card for the whole fortnight. Print, glue in books, use every session.");
const RESOURCE_ITEMS = [MAP_RES];
fs.mkdirSync(RES_DIR, { recursive: true });

/* --- Teacher notes (Glance Format, megaprompt sections 45-47) ------------- */

const NOTES_TITLE =
  "Open Week 9. Say the fortnight in one line: this week we read reports like writers, next week we write and publish our own.";

const NOTES_RESOURCES =
  "Prep slide. Print the Information Report Map Card, one per student. It is glued in books and used in all eight sessions.\n" +
  "Have whiteboards, markers and books ready. No other printing today.\n" +
  "CATCH-UP: this is Session 1, so nothing is assumed. A student who joins later re-enters through the map card and the launch.";

const NOTES_OVERVIEW =
  "For you, not the class. Skip past it when projecting.\n" +
  "Week 9 reads reports like writers, Week 10 plans, drafts, edits and publishes one. The anchor phrase is said in every I Do, both weeks, word for word.\n" +
  "Decision points today: the A/B/C hinge, the boards check in the We Do, and the exit ticket. Between them keep the pace brisk.";

const NOTES_LAUNCH = composeGlanceNotes({
  answer: "open - all three say what the thing IS.",
  beats: [
    [
      "SAY: Three openings, three different topics.",
      "Read all three with me.",
    ],
    [
      "ASK: What do these three have in common?",
      "30 sec. Turn and tell, partner A first.",
      "EXPECT: they all say what the thing is.",
    ],
    [
      "SCAN the room. Listen for is or are.",
      "80%+ -> reveal, then move on.",
      "Less -> read the frog one again, ask its job, re-ask.",
    ],
    [
      "REVEAL after the partner talk.",
      "SAY: Each one classifies. It says what the thing is.",
    ],
  ],
  trap: [
    "naming the topic only, such as frogs.",
    "Fix: point at the word is, student re-reads it.",
  ],
  prep: "Low-coupling launch: answerable from the three sentences on screen, so no earlier lesson is assumed. Whole block under 6 minutes.",
  sources: P.SOURCE_LINE_MODEL,
  tag: "[Launch | Retention and recall | HITS 6]",
});

const NOTES_LI = composeGlanceNotes({
  beats: [
    [
      "POINT to the learning intention.",
      "SAY: Today we work out what makes a text a report.",
    ],
    "SAY: Read the I can statements with me. Everyone, together, on three.",
    "SAY: The first one everyone leaves with. Who is it written for?",
  ],
  prep: "SC1 is reachable by every student; SC2 is the core target the exit ticket assesses; SC3 stretches into explaining fact and opinion. Tier labels stay off the slide.",
  tag: "[LI/SC | Planning made visible | HITS 1]",
});

const NOTES_VOCAB = composeGlanceNotes({
  answer: "audience means the people who will read it.",
  beats: [
    [
      "SAY: Every report is written for somebody. That somebody is the audience.",
      "The audience changes the words you choose.",
    ],
    [
      "ASK: A report on frogs for Preps, or for adults. What changes?",
      "20 sec. Turn and tell, partner B first.",
      "EXPECT: easier words, more pictures, shorter sentences.",
    ],
    "SAY: Same facts, different audience, different words. Hold onto that.",
  ],
  trap: [
    "thinking audience means the topic.",
    "Fix: ask who is reading, student answers with people, not a subject.",
  ],
  prep: "The one word that carries the purpose and audience half of today. Students who can name an audience can judge word choice next week.",
  tag: "[Key Vocabulary | Knowledge and memory | HITS 6]",
});

const NOTES_IDO = composeGlanceNotes({
  answer: "facts, present tense, headings, a visual with a caption.",
  beats: [
    [
      "SAY: A report page. Look at the shape before you read.",
      "Title, headings, picture, caption.",
    ],
    [
      "POINT to each feature card. SAY: Facts you can check.",
      "Present tense: frogs hunt, not hunted.",
    ],
    [
      "ASK: Which feature would a story NOT have?",
      "20 sec. Write it... chin it... show me.",
      "EXPECT: headings.",
    ],
    [
      "SCAN boards.",
      "80%+ -> cold call: how do you know?",
      "Less -> show a novel, flip pages, re-ask.",
    ],
  ],
  trap: [
    "facts alone do not make a report.",
    "Fix: ask what the text is FOR, student re-checks.",
  ],
  stretch: "Name a feature not on our list.",
  help: "Give the map card. Point to a heading.",
  prep: "The unit anchor starts here. Say the anchor phrase once at the end of this slide, word for word, and again in every I Do for two weeks.",
  sources: P.SOURCE_LINE_MODEL,
  tag: "[I Do | Explicit teaching | SC2 | HITS 3, 4]",
});

const NOTES_CFU = composeGlanceNotes({
  answer: "B. It says what honey bees are, and their group.",
  beats: [
    [
      "SAY: Three openings, one topic, three kinds of text.",
      "Do not call out. This one goes on boards.",
    ],
    [
      "ASK: Which opening is from an information report?",
      "30 sec. Boards up on cue.",
      "EXPECT: B.",
    ],
    [
      "SCAN boards, back row first.",
      "80%+ -> cold call one B: how do you know?",
      "Less -> read A aloud, ask whose story it is, re-ask.",
    ],
    [
      "REVEAL after boards are scanned.",
      "SAY: It classifies. That is the tell.",
    ],
  ],
  trap: [
    "picking C because it sounds factual and important.",
    "Fix: ask what C wants you to DO, student re-tests.",
  ],
  prep: "The hinge of the lesson. A is a narrative opening and C is persuasive, so both wrong answers name a real confusion.",
  tag: "[CFU hinge | Supported application | SC2 | HITS 7, 8]",
});

const NOTES_WEDO = composeGlanceNotes({
  answer: "sentence three: It is definitely the most beautiful place.",
  beats: [
    [
      "SAY: New topic, the Great Barrier Reef.",
      "Read the four sentences aloud with me.",
    ],
    [
      "ASK: Which sentence is an opinion?",
      "45 sec. Write the first three words, chin it, show me.",
      "EXPECT: It is definitely.",
    ],
    [
      "SCAN boards for the word definitely.",
      "80%+ -> cold call one: how could we check it?",
      "Less -> read two and three again, ask which one proves, re-ask.",
    ],
  ],
  trap: [
    "calling a big number an opinion because it sounds unlikely.",
    "Fix: ask could we count it, student sorts again.",
  ],
  stretch: "Rewrite the opinion as a fact about the reef.",
  help: "Cover all but two sentences. Which one can you check?",
  prep: "Different topic from the I Do on purpose. The test students take away is: could I check it?",
  sources: P.SOURCE_LINE_MODEL,
  tag: "[We Do | Supported application | SC3 | HITS 4, 5]",
});

const NOTES_YOUDO = composeGlanceNotes({
  answer: "volcanoes, wind farms and honey bees work.",
  beats: [
    [
      "SAY: A report topic is a whole group.",
      "Not one story. Not an argument.",
      "Sort these five in your book.",
    ],
    [
      "POINT to the steps. TIME: eight minutes in books.",
      "Write two topics of your own at the end.",
    ],
    [
      "CIRCULATE. Read for whole groups, not single events.",
      "COLLECT two strong topics for the closing.",
    ],
  ],
  trap: [
    "choosing a topic that is really an argument.",
    "Fix: ask what the reader learns, student rewrites it.",
  ],
  stretch: "Split one of your own topics into three aspects.",
  help: "Give three ready-made topics. Ask only for the sort.",
  prep: "Different task and content from the We Do: judging topics, not spotting opinions. Books, not sheets.",
  tag: "[You Do | Mastery and application | SC2 | HITS 10]",
});

const NOTES_EXIT = composeGlanceNotes({
  answer: "no, it is a story about one day, and it uses I.",
  beats: [
    [
      "SAY: One sentence. Read it, decide, then give me one reason.",
      "Yes or no, and why.",
    ],
    [
      "TIME: three minutes. Books open on the desk when you finish.",
      "COLLECT by walking the rows, not by calling names.",
    ],
    [
      "SORT into two piles as you collect: reason given, no reason.",
      "A thick second pile -> open Session 2 by re-reading the four features.",
    ],
  ],
  prep: "Assesses the core target: name what makes a text an information report. Sort into two piles as you collect.",
  tag: "[Exit Ticket | Mastery and application | SC2 | HITS 8]",
});

const NOTES_CLOSING = composeGlanceNotes({
  answer: "open - most students should show two.",
  beats: [
    "READ the two collected topics aloud. SAY: Listen for a whole group.",
    [
      "ASK: Rate yourself on the three statements.",
      "15 sec. Fingers at your chest, one, two or three. Show me.",
      "EXPECT: most on two.",
    ],
    "SAY: Next session we pull a report apart and find the job each part does.",
  ],
  prep: "Self-assessment feeds Session 2 grouping. Anyone showing one gets the map card open in front of them next session.",
  tag: "[Closing | Retention and recall | HITS 1, 9]",
});

/* --- Slide content -------------------------------------------------------- */

const LAUNCH_OPENINGS = [
  { label: "Text 1", text: P.FROG.classify.split(".")[0] + ".", color: C.PRIMARY },
  { label: "Text 2", text: P.REEF.classify.split(".")[0] + ".", color: C.SECONDARY },
  { label: "Text 3", text: P.THUNDER.classify.split(".")[0] + ".", color: C.ASSESS },
];

const CFU_OPTIONS = [
  { key: "A", text: "The morning the bees came, Nan grabbed a bucket and ran outside." },
  { key: "B", text: "Honey bees are insects that live together in large colonies." },
  { key: "C", text: "We must protect honey bees before it is too late." },
];

const WEDO_SENTENCES = [
  { text: "The Great Barrier Reef is a coral reef system off the coast of Queensland.", opinion: false },
  { text: "It is the largest reef system on Earth.", opinion: false },
  { text: "It is definitely the most beautiful place in the world.", opinion: true },
  { text: "The reef is home to more than fifteen hundred species of fish.", opinion: false },
];

const YOUDO_TOPICS = [
  { text: "Volcanoes", works: true },
  { text: "The day I broke my arm", works: false },
  { text: "Why school should start later", works: false },
  { text: "Wind farms", works: true },
  { text: "Honey bees", works: true },
];

/* --- Build ---------------------------------------------------------------- */

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // 1. Title
  titleSlide(pres, "Information Reports",
    P.ANCHOR_SHORT,
    "Week 9 Session 1 | Year 5/6 Literacy", NOTES_TITLE);

  // 2. Teacher Resources
  addResourceSlide(pres, {
    resources: RESOURCE_ITEMS,
    studentTools: [
      "Mini-whiteboards, markers and erasers",
      "Writing books and pencils",
      "The map card, glued in the front of the book",
    ],
    boardSetup: [
      "Nothing to pre-write",
      "Keep a novel on the desk for the I Do pivot",
    ],
  }, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // 3. Teacher-facing overview
  (() => {
    const s = P.customSlide(pres, T, "For the teacher", C.MUTED, "The fortnight at a glance");
    const cardY = CONTENT_TOP;
    const cardH = SAFE_BOTTOM - CONTENT_TOP;

    addCard(s, 0.5, cardY, 4.4, cardH, { strip: C.PRIMARY });
    s.addText("Where this sits", {
      x: 0.7, y: cardY + 0.12, w: 4.0, h: 0.28,
      fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText([
      { text: "Week 9, read like a writer", options: { bold: true, color: C.PRIMARY, breakLine: true } },
      { text: "S1  What is an information report?", options: { breakLine: true } },
      { text: "S2  The parts and the job each one does", options: { breakLine: true } },
      { text: "S3  Research questions and notes", options: { breakLine: true } },
      { text: "S4  Factual language and noun groups", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "Week 10, write and publish", options: { bold: true, color: C.ASSESS, breakLine: true } },
      { text: "Plan, draft, add visuals, edit, publish.", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "Anchor, said in every I Do:", options: { color: C.MUTED, breakLine: true } },
      { text: P.ANCHOR_PHRASE, options: { color: C.PRIMARY, bold: true } },
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
      { text: "Shape: example-first. Three real openings, then one page model and its four features.", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "Decision points:", options: { bold: true, color: C.ACCENT, breakLine: true } },
      { text: "the A/B/C hinge, the boards check in the We Do, and the exit ticket.", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "Assumption:", options: { bold: true, color: C.ACCENT, breakLine: true } },
      { text: "students have met reports before but not this lens. Language stays mixed-readiness throughout.", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "Model texts are written for this unit, not quoted. Swap in real mentor texts if you have them.", options: { italic: true, color: C.MUTED } },
    ], {
      x: 5.3, y: cardY + 0.46, w: 4.0, h: cardH - 0.60,
      fontSize: 12.5, fontFace: FONT_B, color: C.CHARCOAL,
      valign: "top", margin: 0, paraSpaceAfter: 3,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_OVERVIEW);
    runSlideDiagnostics(s, pres);
  })();

  // 4. Launch - three openings, one thing in common
  (() => {
    const s = P.customSlide(pres, T, "Launch", C.SECONDARY, "What do these three have in common?");
    const y0 = CONTENT_TOP;
    // Three cards must stop above the click-revealed bar: 1.3 + 3 x 0.84 +
    // 2 x 0.12 = 4.06, just under the 4.13 content floor.
    const cardH = 0.84;
    const gap = 0.12;

    LAUNCH_OPENINGS.forEach((item, i) => {
      const cy = y0 + i * (cardH + gap);
      s.addShape("roundRect", {
        x: 0.5, y: cy, w: 9, h: cardH, rectRadius: 0.08,
        fill: { color: C.WHITE }, line: { color: item.color, width: 1.5 },
      });
      s.addShape("roundRect", {
        x: 0.5, y: cy, w: 1.05, h: cardH, rectRadius: 0.08,
        fill: { color: item.color },
      });
      s.addText(item.label, {
        x: 0.5, y: cy, w: 1.05, h: cardH,
        fontSize: 13, fontFace: FONT_B, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      s.addText(item.text, {
        x: 1.72, y: cy, w: 7.6, h: cardH,
        fontSize: 17, fontFace: FONT_B, color: C.CHARCOAL,
        valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
    });

    clickBuild(s, [
      () => {
        addRevealAnswerBar(s, ["Each one CLASSIFIES: it says what the thing is"], {
          y: P.REVEAL_Y, h: P.REVEAL_H, fontSize: 20, label: "Notice", color: C.PRIMARY,
        });
      },
    ]);

    addFooter(s, FOOTER);
    s.addNotes(NOTES_LAUNCH);
    runSlideDiagnostics(s, pres);
  })();

  // 5. Learning Intention and Success Criteria
  liSlide(pres,
    "We are learning what an information report is and what makes it different from other texts.",
    [
      "I can say who an information report is written for.",
      "I can name three features that make a text an information report.",
      "I can explain why a fact belongs in a report and an opinion does not.",
    ],
    NOTES_LI, FOOTER);

  // 6. Key vocabulary
  keyWordSlide(pres, {
    word: "audience",
    meaning: "The people who will read your writing.",
    example: "A report written for Preps uses simpler words than one written for adults.",
    routine: ["Say it", "Name yours", "Use it"],
    color: C.SECONDARY,
    title: "The word for today",
  }, NOTES_VOCAB, FOOTER);

  // 7. I Do - what a report page looks like, and the four features
  annotatedModelSlide(pres, "I Do", "What makes it a report?",
    [
      { text: "Look at the page", role: "header" },
      { text: "Notice the shape before you read a word." },
      { text: "Title. Headings. A picture with words underneath." },
      { text: "Then check the four features." },
      { text: "Last question: who is this written for?" },
    ],
    "An information report page",
    P.REPORT_FEATURES,
    NOTES_IDO, FOOTER,
    {
      badgeColor: C.PRIMARY,
      previewSpec: P.reportPageSpec(C, P.FROG),
      sourceType: "Model report",
    });

  // 8. CFU hinge - which opening is from an information report?
  (() => {
    const s = P.customSlide(pres, T, "CFU", C.ALERT, "Which one is the information report?");
    const y0 = CONTENT_TOP;

    addTextOnShape(s, "Topic: honey bees", {
      x: 0.5, y: y0, w: 9, h: 0.48, rectRadius: 0.08,
      fill: { color: C.ALERT },
    }, {
      fontSize: 16, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    // Options must clear the click-revealed answer bar: last card bottom is
    // 1.90 + 3 x 0.78 - 0.12 = 4.12, just under the 4.13 content floor.
    P.drawOptionStack(s, T, CFU_OPTIONS, {
      y: y0 + 0.60, optionH: 0.66, gap: 0.12, color: C.PRIMARY, fontSize: 17,
    });

    clickBuild(s, [
      () => {
        addRevealAnswerBar(s, ["B - it says what bees ARE and their group"], {
          y: P.REVEAL_Y, h: P.REVEAL_H, fontSize: 21, color: C.SUCCESS,
        });
      },
    ]);

    addFooter(s, FOOTER);
    s.addNotes(NOTES_CFU);
    runSlideDiagnostics(s, pres);
  })();

  // 9. We Do - one sentence does not belong
  (() => {
    const s = P.customSlide(pres, T, "We Do", C.SUCCESS, "Which sentence is an opinion?");
    const y0 = CONTENT_TOP;

    // Extract + cue must stop above the reveal bar: 1.3 + 2.26 + 0.12 + 0.44
    // = 4.12, just under the 4.13 content floor.
    P.drawReportExtract(s, T, {
      x: 0.5, y: y0, w: 9, h: 2.26,
      title: P.REEF.title,
      titleSize: 19,
      titleColor: C.SECONDARY,
      bodySize: 15.5,
      showLabels: false,
      sections: WEDO_SENTENCES.map((item) => ({ text: item.text })),
    });

    addTextOnShape(s, "Write the first three words of the opinion sentence.", {
      x: 0.5, y: y0 + 2.38, w: 9, h: 0.44, rectRadius: 0.08,
      fill: { color: C.SUCCESS },
    }, {
      fontSize: 17, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    clickBuild(s, [
      () => {
        addRevealAnswerBar(s, ["It is definitely... - nobody can check that"], {
          y: P.REVEAL_Y, h: P.REVEAL_H, fontSize: 21, color: C.SUCCESS,
        });
      },
    ]);

    addFooter(s, FOOTER);
    s.addNotes(NOTES_WEDO);
    runSlideDiagnostics(s, pres);
  })();

  // 10. You Do - which topics work?
  (() => {
    const s = P.customSlide(pres, T, "You Do", C.ASSESS, "Which topics would work?");
    const y0 = CONTENT_TOP;
    const rowH = 0.54;
    const gap = 0.10;

    YOUDO_TOPICS.forEach((item, i) => {
      const ry = y0 + i * (rowH + gap);
      s.addShape("roundRect", {
        x: 0.5, y: ry, w: 5.9, h: rowH, rectRadius: 0.07,
        fill: { color: C.WHITE }, line: { color: C.ASSESS, width: 1.4 },
      });
      s.addText(item.text, {
        x: 0.75, y: ry, w: 5.4, h: rowH,
        fontSize: 17, fontFace: FONT_B, color: C.CHARCOAL,
        valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
    });

    const steps = [
      "Sort the five into yes and no.",
      "Write one reason for each no.",
      "Write two topics of your own.",
    ];
    addCard(s, 6.7, y0, 2.8, 2.30, { strip: C.ASSESS });
    s.addText("Steps", {
      x: 6.92, y: y0 + 0.10, w: 2.4, h: 0.30,
      fontSize: 15, fontFace: FONT_B, color: C.ASSESS, bold: true, margin: 0,
    });
    steps.forEach((step, i) => {
      s.addText(String(i + 1) + ". " + step, {
        x: 6.92, y: y0 + 0.48 + i * 0.58, w: 2.4, h: 0.54,
        fontSize: 13.5, fontFace: FONT_B, color: C.CHARCOAL,
        valign: "middle", margin: 0,
      });
    });

    addTextOnShape(s, "A report topic is a whole group. Not one story. Not an argument.", {
      x: 6.7, y: y0 + 2.44, w: 2.8, h: 0.76, rectRadius: 0.08,
      fill: { color: C.ASSESS },
    }, {
      fontSize: 13, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0.06,
    });

    s.addText("Eight minutes in your books.", {
      x: 0.5, y: y0 + 3.30, w: 5.9, h: 0.42,
      fontSize: 15, fontFace: FONT_B, color: C.ASSESS, italic: true,
      align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
    runSlideDiagnostics(s, pres);
  })();

  // 11. Exit ticket
  P.exitTicketPanel(pres, T, {
    topic: "On Saturday I saw a huge frog on our letterbox and it made me jump.",
    task: "Is this from an information report? Write yes or no, and one reason.",
    cue: "One sentence. One reason.",
    taskSize: 26,
    footer: FOOTER,
    notes: NOTES_EXIT,
  });

  // 12. Closing
  closingSlide(pres, {
    reflectionPrompt: "Which feature of an information report will be easiest for you to use?",
    scItems: [
      "I can say who an information report is written for.",
      "I can name three features that make a text an information report.",
      "I can explain why a fact belongs in a report and an opinion does not.",
    ],
    selfAssessment: "Fingers at your chest. One, two or three. Show me.",
    takeaways: [P.ANCHOR_PHRASE],
  }, NOTES_CLOSING);

  const pptxPath = path.join(OUT_DIR, "Information Reports W9 S1.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);
}

/* --- Resource: the Information Report Map Card ---------------------------- */

async function buildMapCard() {
  const doc = createPdf({ title: MAP_RES.name });
  let y = addPdfHeader(doc, "Information Report Map", {
    subtitle: "Glue this in your book. You will use it every session for two weeks.",
    color: C.PRIMARY,
    lessonInfo: "Year 5/6 Literacy | Information Reports",
    showNameDate: false,
  });

  y = addTipBox(doc, P.ANCHOR_PHRASE, y, { color: C.ACCENT });
  y += 6;

  y = P.addReportMapPdf(doc, y, {
    colors: {
      classify: hex(C.PRIMARY),
      describe: hex(C.SECONDARY),
      wrap: hex(C.ASSESS),
      facts: hex(C.ACCENT),
    },
    aspects: P.GENERIC_ASPECTS,
  });

  y = addSectionHeading(doc, "What each part does", y, { color: C.PRIMARY });
  const detail = [
    ["Classify", "Say what the thing is and what group it belongs to. One or two sentences. Frame: The [topic] is a [group] that [something true about it]."],
    ["Describe", "One aspect per paragraph. Appearance, habitat, diet, size, uses, dangers. Start each paragraph with a topic sentence that names the aspect."],
    ["Wrap up", "Sum up in new words, or finish with one last interesting fact."],
    ["Facts", "Present tense. No opinions. Technical words. Noun groups that describe. A visual with a caption."],
  ];
  const colors = [hex(C.PRIMARY), hex(C.SECONDARY), hex(C.ASSESS), hex(C.ACCENT)];
  detail.forEach((row, i) => {
    doc.fontSize(11).font("Sans-Bold").fillColor(colors[i]);
    doc.text(row[0], PAGE.MARGIN, y, { width: 80, continued: false });
    doc.fontSize(10.5).font("Sans").fillColor("#000000");
    doc.text(row[1], PAGE.MARGIN + 80, y, { width: PAGE.CONTENT_W - 80 });
    y = doc.y + 8;
  });

  y += 4;
  y = addSectionHeading(doc, "Four features to look for", y, { color: C.ACCENT });
  y = addBodyText(doc,
    "Facts you can check. Present tense: frogs hunt, not frogs hunted. Headings, one aspect under each. A visual with a caption that adds something.",
    y);

  y = addSectionHeading(doc, "Topic sentence starters for the middle", y, { color: C.SECONDARY });
  const starters = [
    "The [topic] is a ... that ...",
    "[Topic] have ... / [Topic] has ...",
    "[Topic] live in ...",
    "[Topic] eat ...",
    "Most [topic] ...",
    "The biggest danger to [topic] is ...",
  ];
  starters.forEach((line) => {
    doc.circle(PAGE.MARGIN + 8, y + 5, 2).fill(hex(C.SECONDARY));
    doc.fontSize(10.5).font("Sans").fillColor("#000000");
    doc.text(line, PAGE.MARGIN + 18, y, { width: PAGE.CONTENT_W - 26 });
    y += 15;
  });
  y += 8;

  y = addSectionHeading(doc, "Words to leave out", y, { color: C.ALERT });
  y = addBodyText(doc,
    "best, worst, amazing, cool, favourite, beautiful, boring, should, I think. Every one is an opinion.",
    y);

  y = addBodyText(doc,
    "Ask first: who is my audience? Same facts, different words.",
    y, { italic: true });

  addPdfFooter(doc, "Information Reports | Year 5/6 Literacy");
  await writePdf(doc, path.join(OUT_DIR, MAP_RES.fileName));
  console.log("Wrote " + MAP_RES.name);
}

(async () => {
  await build();
  await buildMapCard();
})().catch((err) => {
  console.error(err);
  process.exit(1);
});
