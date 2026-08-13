"use strict";

// Persuasive Writing - Term 3 Week 7, Session 1 (Year 5/6 Literacy)
// "Pick your fight: debate, choose a topic, plan the reasons"
//
// Victorian Curriculum 2.0: English, Literacy (Interacting with others, and
// Creating texts), Levels 5-6 - presenting a point of view, and planning a
// sustained persuasive text for a stated audience.
//
// UNIT ANCHOR (locked, carried from Week 6):
//   Bold Beginning. Mighty Middle. Excellent Ending. Voice all the way through.
// Week 7 extends it rather than replacing it:
//   Same belief. Same reasons. Different voice.
//
// This week each student writes ONE topic three ways: a letter, a poster and a
// news article. Today builds the plan all three will draw on.

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
  titleSlide, liSlide, closingSlide,
  addCard, addFooter, addTextOnShape, addRevealAnswerBar,
  clickBuild, runSlideDiagnostics,
} = T;

const SESSION = 1;
const FOOTER = "Persuasive Writing | Week 7 Session 1 | Year 5/6 Literacy";
const OUT_DIR = "output/Persuasive_W7_S1";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const PLAN_RES = makeSessionResource(SESSION,
  "Persuasive Plan",
  "Your topic, your belief and three reasons with evidence. Used all week.");
const RESOURCE_ITEMS = [PLAN_RES];
fs.mkdirSync(RES_DIR, { recursive: true });

const FORMATS = [
  { label: "Letter", who: "someone who can say yes", color: "PRIMARY" },
  { label: "Poster", who: "everyone walking past", color: "SECONDARY" },
  { label: "News article", who: "the whole school community", color: "ASSESS" },
];

const DEBATE_STATEMENTS = [
  "Homework should be banned on weekends.",
  "Every classroom should have a scrap paper bin.",
  "Students should choose where they sit.",
];

// ─── Teacher notes ──────────────────────────────────────────────────────────

const NOTES_TITLE =
  "Week 7, Session 1. One line: this week you pick one thing you actually care about and publish it three ways.";

const NOTES_RESOURCES =
  "Prep slide. Print the Persuasive Plan, one per student. It is the spine of the whole week and is not reprinted.\n" +
  "Students need their Week 6 Word Bank and Structure Card. Have spares.\n" +
  "CATCH-UP: missed Week 6? Give a spare Word Bank and Structure Card, say the anchor phrase, and start them at the debate. The plan can be built from scratch today.";

const NOTES_OVERVIEW =
  "For you, not the class. Skip past it when projecting.\n" +
  "Week 7 is one topic published three ways: letter, poster, news article. Today builds the plan all three sessions draw on, so protect the You Do time.\n" +
  "Decision points: the reason-or-opinion hinge, the boards check in the We Do, and the exit ticket. Between them keep the pace brisk.";

const NOTES_LAUNCH = composeGlanceNotes({
  answer: "open - listen for a reason, not just a side.",
  beats: [
    [
      "SAY: Quick debate. I read a statement. You stand if you agree.",
      "SAY: Then you have to convince the people still sitting down.",
    ],
    [
      "ASK: Convince the other side. Give me your best reason.",
      "60 sec per statement. Stand if you agree, then cold call two standers.",
      "EXPECT: a reason with because in it.",
    ],
    [
      "SCAN who stands and who does not move.",
      "80%+ giving reasons -> next slide.",
      "Less -> model one yourself: I stand because..., re-ask.",
    ],
  ],
  trap: [
    "saying a side loudly with no reason attached.",
    "Fix: ask for the word because, student finishes the sentence.",
  ],
  prep: "The teacher's own opener. Low coupling: any student can take a side on these three, whatever they missed. Whole block under 8 minutes.",
  tag: "[Launch | Attention, focus and regulation | HITS 5, 6]",
});

const NOTES_LI = composeGlanceNotes({
  beats: [
    "POINT to the intention. SAY: Today you choose your fight and plan it.",
    "SAY: Read the I can statements with me. Everyone, together, on three.",
    "SAY: Everyone leaves today with a topic and a side written down.",
  ],
  prep: "SC1 reachable by all; SC2 is the core target the exit ticket assesses; SC3 stretches into ranking the reasons.",
  tag: "[LI/SC | Planning made visible | HITS 1]",
});

const NOTES_ANCHOR = composeGlanceNotes({
  answer: "the belief and reasons stay the same, only the voice changes.",
  beats: [
    [
      "POINT to the three formats. SAY: One topic. Three readers. Three pieces.",
      "SAY: Same belief. Same reasons. Different voice.",
    ],
    [
      "CLICK through the three readers.",
      "SAY: A letter argues, a poster shouts, an article reports and then argues.",
    ],
    [
      "ASK: Which one gets the fewest words?",
      "10 sec. Fingers at your chest. One, two or three.",
      "EXPECT: two, the poster.",
    ],
  ],
  trap: [
    "thinking three pieces means three different topics.",
    "Fix: point at one plan, student says all three come from it.",
  ],
  stretch: "Say which format would change your topic's mind fastest, and why.",
  help: "Give the Structure Card and name only the letter for now.",
  prep: "Say the Week 7 phrase word for word every session. It is what stops students starting over three times.",
  tag: "[I Do | Explicit teaching | SC1 | HITS 2, 3]",
});

const NOTES_IDO = composeGlanceNotes({
  answer: "belief, then three reasons, each with something checkable.",
  beats: [
    [
      "SAY: Watch me plan mine. My topic is the scrap paper bin.",
      "SAY: First my belief, in one sentence, with a must in it.",
    ],
    [
      "CLICK through my three reasons.",
      "SAY: Each one needs proof. Watch me reject a reason I cannot prove.",
    ],
    [
      "ASK: Which of my three reasons is strongest?",
      "20 sec. Turn and tell. Partner A first.",
      "EXPECT: the one with the countable proof.",
    ],
  ],
  trap: [
    "planning three reasons that are the same reason reworded.",
    "Fix: read two aloud together, student hears the repeat.",
  ],
  stretch: "Add a fourth reason someone who disagrees would find hardest to answer.",
  help: "Give the plan with the belief sentence already written.",
  prep: "Model rejecting a weak reason out loud. That move is what stops students padding the plan.",
  tag: "[I Do | Explicit teaching | SC2 | HITS 4]",
});

const NOTES_CFU = composeGlanceNotes({
  answer: "B. It gives a reason with proof behind it.",
  beats: [
    [
      "SAY: Three lines about the same topic. Only one is a real reason.",
      "Do not call out. Boards for this one.",
    ],
    [
      "ASK: Which one is a reason, A, B or C?",
      "30 sec. Write it... chin it... show me.",
      "EXPECT: B.",
    ],
    [
      "SCAN boards, back row first.",
      "80%+ -> cold call one B: what makes it a reason? Then reveal.",
      "Less -> read A aloud, ask what proof it has, re-ask.",
    ],
    "REVEAL after boards are scanned. SAY: A reason can be argued with. An opinion cannot.",
  ],
  trap: [
    "picking A because it sounds strong and certain.",
    "Fix: ask how you would prove A, student re-tests all three.",
  ],
  prep: "The hinge. A restates the belief, C is a fact with no argument attached. Both wrong answers are diagnostic for the plan.",
  tag: "[CFU hinge | Supported application | SC2 | HITS 7, 8]",
});

const NOTES_WEDO = composeGlanceNotes({
  answer: "open - a reason with something checkable behind it.",
  beats: [
    "SAY: I give the reason. You find the proof.",
    [
      "ASK: What proof could we actually get for that?",
      "60 sec. Write it... chin it... show me.",
      "EXPECT: a count, a price, or ask the office.",
    ],
    [
      "SCAN boards for checkable proof.",
      "80%+ -> cold call one strong, one shaky: how would you check it?",
      "Less -> name one proof yourself, re-ask for a second.",
    ],
  ],
  trap: [
    "offering an opinion as proof.",
    "Fix: ask who could confirm it, student rewrites it.",
  ],
  stretch: "Find proof someone who disagrees would still accept.",
  help: "Give two proof options, have them choose one.",
  prep: "Proof does not have to be a statistic. A count, a price or a person to ask all work at this level.",
  tag: "[We Do | Supported application | SC2 | HITS 4, 5]",
});

const NOTES_YOUDO = composeGlanceNotes({
  answer: "open - a topic, a belief and three reasons on the plan.",
  beats: [
    [
      "SAY: Your topic now. Something you actually want changed.",
      "SAY: It has to be arguable. Someone must be able to disagree.",
    ],
    [
      "POINT to the plan. SAY: Belief first. Then three reasons. Then proof.",
      "TIME: eighteen minutes. This plan runs all week.",
    ],
    [
      "CIRCULATE. Approve every topic before they write reasons.",
      "COLLECT nothing. Plans stay with students all week.",
    ],
  ],
  trap: [
    "choosing a topic nobody could disagree with.",
    "Fix: ask who would argue back, student narrows the topic.",
  ],
  stretch: "Rank your three reasons and mark which goes first.",
  help: "Offer a shortlist of three approved topics to choose from.",
  prep: "Protect this time. A weak plan today costs three sessions. Check topics early, not at the end.",
  tag: "[You Do | Mastery and application | SC2 | HITS 10]",
});

const NOTES_EXIT = composeGlanceNotes({
  answer: "a topic, a side, and one reason with because in it.",
  beats: [
    [
      "SAY: On the bottom of your plan. Three things.",
      "SAY: Your topic, your side, and your best reason.",
    ],
    [
      "TIME: four minutes. Plans stay with you.",
      "WALK the rows and read them over shoulders.",
    ],
    [
      "SORT in your head: ready to write, or topic still shaky.",
      "Shaky ones -> five minutes with you at the start of Session 2.",
    ],
  ],
  prep: "Assesses the core target. Do not collect the plans; students need them tomorrow. Read them while walking.",
  tag: "[Exit Ticket | Mastery and application | SC2 | HITS 8]",
});

const NOTES_CLOSING = composeGlanceNotes({
  answer: "open - most students should show two.",
  beats: [
    "COLD CALL three students for their topic in one sentence. No reasons.",
    [
      "ASK: Rate yourself on the three statements.",
      "15 sec. Fingers at your chest. One, two or three. Show me.",
      "EXPECT: most on two.",
    ],
    "SAY: Tomorrow the same belief becomes a letter to someone who can act.",
  ],
  prep: "Note the shaky topics now. Those students get five minutes with you before Session 2 starts.",
  tag: "[Closing | Retention and recall | HITS 1, 9]",
});

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  titleSlide(pres, "Pick Your Fight",
    "One topic. Three readers. Three published pieces.",
    "Week 7 Session 1 | Year 5/6 Literacy", NOTES_TITLE);

  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  (() => {
    const s = P.customSlide(pres, T, "For the teacher", C.MUTED, "Week 7 at a glance", { badgeW: 2.4 });
    const cardY = CONTENT_TOP;
    const cardH = SAFE_BOTTOM - CONTENT_TOP;

    addCard(s, 0.5, cardY, 4.4, cardH, { strip: C.PRIMARY });
    s.addText("The week", {
      x: 0.7, y: cardY + 0.12, w: 4.0, h: 0.28,
      fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText([
      { text: "S1  Debate, choose a topic, plan", options: { bold: true, breakLine: true } },
      { text: "S2  Write the letter", options: { bold: true, breakLine: true } },
      { text: "S3  Design the poster", options: { bold: true, breakLine: true } },
      { text: "S4  Write the news article", options: { bold: true, breakLine: true } },
      { text: "S5  Publish, share and reflect", options: { bold: true, breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "Carried from Week 6:", options: { color: C.MUTED, breakLine: true } },
      { text: P.ANCHOR_PHRASE, options: { color: C.PRIMARY, bold: true, breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "New this week:", options: { color: C.MUTED, breakLine: true } },
      { text: P.ANCHOR_PHRASE_W7, options: { color: C.ASSESS, bold: true } },
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
      { text: "Shape: problem-first. The debate exposes who already has a position before any planning is modelled.", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "Protect the You Do:", options: { bold: true, color: C.ACCENT, breakLine: true } },
      { text: "a weak plan today costs three sessions. Approve every topic before students write reasons.", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "Decision points:", options: { bold: true, color: C.ACCENT, breakLine: true } },
      { text: "the reason-or-opinion hinge, the We Do boards check, and the exit ticket.", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "Flex: cut the debate to two statements if topics are slow to settle.", options: { italic: true, color: C.MUTED } },
    ], {
      x: 5.3, y: cardY + 0.46, w: 4.0, h: cardH - 0.60,
      fontSize: 12.5, fontFace: FONT_B, color: C.CHARCOAL,
      valign: "top", margin: 0, paraSpaceAfter: 3,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_OVERVIEW);
    runSlideDiagnostics(s, pres);
  })();

  // Launch - the debate
  (() => {
    const s = P.customSlide(pres, T, "Launch", C.SECONDARY, "Stand up if you agree");
    const y0 = CONTENT_TOP;
    const ch = 0.80;
    DEBATE_STATEMENTS.forEach((statement, i) => {
      const cy = y0 + i * (ch + 0.14);
      s.addShape("roundRect", {
        x: 0.5, y: cy, w: 9, h: ch, rectRadius: 0.08,
        fill: { color: C.WHITE }, line: { color: C.SECONDARY, width: 1.5 },
      });
      s.addText(statement, {
        x: 0.85, y: cy, w: 8.3, h: ch,
        fontSize: 22, fontFace: FONT_B, color: C.CHARCOAL,
        valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
    });

    addTextOnShape(s, "Now convince the people still sitting down.", {
      x: 0.5, y: y0 + 3 * (ch + 0.14) + 0.06, w: 9, h: 0.72, rectRadius: 0.08,
      fill: { color: C.SECONDARY },
    }, {
      fontSize: 21, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_LAUNCH);
    runSlideDiagnostics(s, pres);
  })();

  liSlide(pres,
    "We are learning to choose a topic and plan the reasons that will convince a reader.",
    [
      "I can say what I believe about my topic.",
      "I can plan three reasons with proof for each one.",
      "I can explain which of my reasons is strongest and why.",
    ],
    NOTES_LI, FOOTER);

  // Anchor - same belief, different voice
  (() => {
    const s = P.customSlide(pres, T, "I Do", C.PRIMARY, "One topic. Three readers.");
    const y0 = CONTENT_TOP;

    addTextOnShape(s, P.ANCHOR_PHRASE_W7, {
      x: 0.5, y: y0, w: 9, h: 0.78, rectRadius: 0.08,
      fill: { color: C.ACCENT },
    }, {
      fontSize: 25, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    const bw = 2.87;
    const gap = 0.20;
    clickBuild(s, FORMATS.map((fmt, i) => () => {
      const bx = 0.5 + i * (bw + gap);
      const by = y0 + 0.96;
      s.addShape("roundRect", {
        x: bx, y: by, w: bw, h: 1.95, rectRadius: 0.08,
        fill: { color: C.WHITE }, line: { color: C[fmt.color], width: 1.5 },
      });
      s.addShape("roundRect", {
        x: bx, y: by, w: bw, h: 0.62, rectRadius: 0.08,
        fill: { color: C[fmt.color] },
      });
      s.addText(fmt.label, {
        x: bx, y: by, w: bw, h: 0.62,
        fontSize: 19, fontFace: FONT_H, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      s.addText("Written for", {
        x: bx + 0.12, y: by + 0.72, w: bw - 0.24, h: 0.28,
        fontSize: 12, fontFace: FONT_B, color: C.MUTED,
        align: "center", valign: "middle", margin: 0,
      });
      s.addText(fmt.who, {
        x: bx + 0.12, y: by + 1.02, w: bw - 0.24, h: 0.82,
        fontSize: 17, fontFace: FONT_B, color: C.CHARCOAL, bold: true,
        align: "center", valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
    }));

    s.addText("The belief does not change. The words do.", {
      x: 0.5, y: y0 + 3.06, w: 9, h: 0.46,
      fontSize: 17, fontFace: FONT_B, color: C.PRIMARY, italic: true, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_ANCHOR);
    runSlideDiagnostics(s, pres);
  })();

  // I Do - plan the model topic
  (() => {
    const s = P.customSlide(pres, T, "I Do", C.PRIMARY, "Watch me plan mine");
    const y0 = CONTENT_TOP;

    addCard(s, 0.5, y0, 9, 0.92, { strip: C.ALERT });
    s.addText("My belief", {
      x: 0.75, y: y0 + 0.07, w: 3, h: 0.26,
      fontSize: 13, fontFace: FONT_B, color: C.ALERT, bold: true, margin: 0,
    });
    s.addText("Every classroom must have a scrap paper bin.", {
      x: 0.75, y: y0 + 0.34, w: 8.5, h: 0.50,
      fontSize: 21, fontFace: FONT_B, color: C.CHARCOAL,
      valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
    });

    const reasons = [
      { r: "It saves the school money.", p: "Ask the office what paper costs a term." },
      { r: "It cuts what we send to landfill.", p: "Count one week of our own bin." },
      { r: "It teaches a habit we keep for life.", p: "Compare with the schools already doing it." },
    ];
    const rh = 0.66;
    const gap = 0.09;
    clickBuild(s, reasons.map((item, i) => () => {
      const ry = y0 + 1.08 + i * (rh + gap);
      s.addShape("roundRect", {
        x: 0.5, y: ry, w: 4.7, h: rh, rectRadius: 0.07,
        fill: { color: C.PRIMARY },
      });
      s.addText(item.r, {
        x: 0.68, y: ry, w: 4.34, h: rh,
        fontSize: 15.5, fontFace: FONT_B, color: C.WHITE, bold: true,
        valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
      s.addShape("roundRect", {
        x: 5.35, y: ry, w: 4.15, h: rh, rectRadius: 0.07,
        fill: { color: C.WHITE }, line: { color: C.SECONDARY, width: 1.2 },
      });
      s.addText(item.p, {
        x: 5.52, y: ry, w: 3.81, h: rh,
        fontSize: 13.5, fontFace: FONT_B, color: C.CHARCOAL,
        valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
    }));

    s.addText("Reason on the left. Proof on the right. No proof, no reason.", {
      x: 0.5, y: y0 + 3.34, w: 9, h: 0.42,
      fontSize: 15, fontFace: FONT_B, color: C.MUTED, italic: true,
      align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO);
    runSlideDiagnostics(s, pres);
  })();

  // CFU hinge - reason or opinion
  (() => {
    const s = P.customSlide(pres, T, "CFU", C.ALERT, "Which one is a reason?");
    const y0 = CONTENT_TOP;

    addTextOnShape(s, "Belief: our school should start a walking bus", {
      x: 0.5, y: y0, w: 9, h: 0.48, rectRadius: 0.08,
      fill: { color: C.ALERT },
    }, {
      fontSize: 16, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    const options = [
      { key: "A", text: "A walking bus is definitely the best idea our school has had." },
      { key: "B", text: "It cuts traffic at the gate, because forty fewer cars would arrive." },
      { key: "C", text: "Some schools in Victoria have a walking bus every Friday." },
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
        addRevealAnswerBar(s, ["B - a because, and a number behind it"], {
          y: P.REVEAL_Y, h: P.REVEAL_H, fontSize: 21, color: C.SUCCESS,
        });
      },
    ]);

    addFooter(s, FOOTER);
    s.addNotes(NOTES_CFU);
    runSlideDiagnostics(s, pres);
  })();

  // We Do
  (() => {
    const s = P.customSlide(pres, T, "We Do", C.SUCCESS, "I give the reason. You find the proof.");
    const y0 = CONTENT_TOP;

    addCard(s, 0.5, y0, 9, 1.10, { strip: C.PRIMARY });
    s.addText("Reason (mine)", {
      x: 0.75, y: y0 + 0.09, w: 4, h: 0.28,
      fontSize: 14, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("A scrap paper bin saves money, because paper costs the school real money.", {
      x: 0.75, y: y0 + 0.40, w: 8.5, h: 0.60,
      fontSize: 20, fontFace: FONT_B, color: C.CHARCOAL,
      valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
    });

    const py = y0 + 1.26;
    addCard(s, 0.5, py, 9, 1.10, { strip: C.SECONDARY });
    s.addText("Proof (yours)", {
      x: 0.75, y: py + 0.09, w: 4, h: 0.28,
      fontSize: 14, fontFace: FONT_B, color: C.SECONDARY, bold: true, margin: 0,
    });
    s.addText("How could we actually check that? A count, a price, or someone to ask.", {
      x: 0.75, y: py + 0.40, w: 8.5, h: 0.60,
      fontSize: 20, fontFace: FONT_B, color: C.CHARCOAL,
      valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
    });

    addTextOnShape(s, "Write it... chin it... show me.", {
      x: 0.5, y: py + 1.28, w: 9, h: 0.80, rectRadius: 0.08,
      fill: { color: C.WHITE }, line: { color: C.SUCCESS, width: 1.5 },
    }, {
      fontSize: 19, fontFace: FONT_B, color: C.SUCCESS, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_WEDO);
    runSlideDiagnostics(s, pres);
  })();

  // You Do - plan your own
  (() => {
    const s = P.customSlide(pres, T, "You Do", C.ASSESS, "Your topic. Your plan.");
    const y0 = CONTENT_TOP;

    addTextOnShape(s, "Pick something you actually want changed", {
      x: 0.5, y: y0, w: 9, h: 0.78, rectRadius: 0.08,
      fill: { color: C.ASSESS },
    }, {
      fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    const steps = [
      { n: "1", t: "Write your belief in one sentence." },
      { n: "2", t: "Write three reasons. One line each." },
      { n: "3", t: "Beside each reason, write how you would prove it." },
    ];
    const sh = 0.68;
    steps.forEach((step, i) => {
      const sy = y0 + 0.94 + i * (sh + 0.11);
      s.addShape("roundRect", {
        x: 0.5, y: sy, w: 9, h: sh, rectRadius: 0.08,
        fill: { color: C.WHITE }, line: { color: C.ASSESS, width: 1.5 },
      });
      s.addShape("roundRect", {
        x: 0.70, y: sy + 0.14, w: 0.44, h: 0.44, rectRadius: 0.22,
        fill: { color: C.ASSESS },
      });
      s.addText(step.n, {
        x: 0.70, y: sy + 0.14, w: 0.44, h: 0.44,
        fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      s.addText(step.t, {
        x: 1.32, y: sy, w: 7.9, h: sh,
        fontSize: 18, fontFace: FONT_B, color: C.CHARCOAL,
        valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
    });

    s.addText("Your topic must be arguable. Someone has to be able to disagree.", {
      x: 0.5, y: y0 + 3.30, w: 9, h: 0.42,
      fontSize: 15, fontFace: FONT_B, color: C.MUTED, italic: true,
      align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
    runSlideDiagnostics(s, pres);
  })();

  P.exitTicketPanel(pres, T, {
    topic: "On the bottom of your plan",
    task: "Your topic. Your side. Your best reason.",
    cue: "Keep your plan. You need it all week.",
    footer: FOOTER,
    notes: NOTES_EXIT,
  });

  closingSlide(pres, {
    reflectionPrompt: "Who is the one person you most want to convince about your topic?",
    scItems: [
      "I can say what I believe about my topic.",
      "I can plan three reasons with proof for each one.",
      "I can explain which of my reasons is strongest and why.",
    ],
    selfAssessment: "Fingers at your chest. One, two or three. Show me.",
    takeaways: [P.ANCHOR_PHRASE_W7],
  }, NOTES_CLOSING);

  const pptxPath = path.join(OUT_DIR, "Persuasive Writing W7 S1.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);
}

// ─── Resource: the Persuasive Plan ──────────────────────────────────────────

async function buildPlan() {
  const doc = createPdf({ title: PLAN_RES.name });
  let y = addPdfHeader(doc, "Persuasive Plan", {
    subtitle: "One topic, three pieces. This plan runs all week. Do not lose it.",
    color: C.ASSESS,
    lessonInfo: "Week 7 | Year 5/6 Literacy",
  });

  y = addTipBox(doc, P.ANCHOR_PHRASE_W7, y, { color: C.ACCENT });
  y += 6;

  doc.fontSize(11).font("Sans-Bold").fillColor("#000000");
  doc.text("My topic:", PAGE.MARGIN, y);
  doc.moveTo(PAGE.MARGIN + 62, y + 14).lineTo(PAGE.MARGIN + PAGE.CONTENT_W, y + 14)
    .lineWidth(0.9).strokeColor("#000000").stroke();
  y += 32;

  doc.save();
  doc.roundedRect(PAGE.MARGIN, y, PAGE.CONTENT_W, 22, 3).fill(hex(C.ALERT));
  doc.fontSize(10).font("Sans-Bold").fillColor("#FFFFFF");
  doc.text("My belief, in one sentence. Use a high modality word: must, should, never.",
    PAGE.MARGIN + 8, y + 6, { width: PAGE.CONTENT_W - 16 });
  doc.restore();
  y += 28;
  y = addLinedArea(doc, y, 2, { lineSpacing: 30 });
  y += 16;

  y = addSectionHeading(doc, "Three reasons, and how I would prove each one", y, { color: C.PRIMARY });
  y += 6;

  // A two-column reason/proof grid, drawn so the structure is on the page
  // rather than described in words.
  const colW = PAGE.CONTENT_W / 2;
  doc.save();
  doc.rect(PAGE.MARGIN, y, colW, 22).fill(hex(C.PRIMARY));
  doc.rect(PAGE.MARGIN + colW, y, colW, 22).fill(hex(C.SECONDARY));
  doc.fontSize(10).font("Sans-Bold").fillColor("#FFFFFF");
  doc.text("Reason", PAGE.MARGIN, y + 6, { width: colW, align: "center" });
  doc.text("How I would prove it", PAGE.MARGIN + colW, y + 6, { width: colW, align: "center" });
  doc.restore();
  y += 22;

  const rowH = 62;
  for (let r = 0; r < 3; r += 1) {
    doc.save();
    doc.rect(PAGE.MARGIN, y, colW, rowH).lineWidth(0.9).strokeColor("#000000").stroke();
    doc.rect(PAGE.MARGIN + colW, y, colW, rowH).lineWidth(0.9).strokeColor("#000000").stroke();
    doc.circle(PAGE.MARGIN + 14, y + 14, 9).fill(hex(C.PRIMARY));
    doc.fontSize(10).font("Sans-Bold").fillColor("#FFFFFF");
    doc.text(String(r + 1), PAGE.MARGIN + 7, y + 9, { width: 14, align: "center" });
    doc.restore();
    // Faint writing rules inside each cell so students know how much to write.
    doc.save();
    doc.strokeColor("#000000").lineWidth(0.5);
    [30, 50].forEach((off) => {
      doc.moveTo(PAGE.MARGIN + 8, y + off).lineTo(PAGE.MARGIN + colW - 8, y + off).stroke();
      doc.moveTo(PAGE.MARGIN + colW + 8, y + off).lineTo(PAGE.MARGIN + PAGE.CONTENT_W - 8, y + off).stroke();
    });
    doc.restore();
    y += rowH;
  }
  y += 16;

  doc.save();
  doc.roundedRect(PAGE.MARGIN, y, PAGE.CONTENT_W, 22, 3).fill(hex(C.SUCCESS));
  doc.fontSize(10).font("Sans-Bold").fillColor("#FFFFFF");
  doc.text("My strongest reason is number ______ , because ...",
    PAGE.MARGIN + 8, y + 6, { width: PAGE.CONTENT_W - 16 });
  doc.restore();
  y += 28;
  y = addLinedArea(doc, y, 2, { lineSpacing: 28 });

  doc.addPage();
  y = PAGE.MARGIN;
  y = addSectionHeading(doc, "Three readers, one belief", y, { color: C.ASSESS });
  y += 6;
  y = addBodyText(doc,
    "This week you will turn this one plan into three pieces. The belief and the reasons stay "
    + "the same. Only the voice changes, because the reader changes.",
    y, {});
  y += 8;

  FORMATS.forEach((fmt) => {
    doc.save();
    doc.roundedRect(PAGE.MARGIN, y, PAGE.CONTENT_W, 26, 3).fill(hex(C[fmt.color]));
    doc.fontSize(11).font("Sans-Bold").fillColor("#FFFFFF");
    doc.text(fmt.label + "   written for " + fmt.who, PAGE.MARGIN + 8, y + 8,
      { width: PAGE.CONTENT_W - 16 });
    doc.restore();
    y += 32;
    doc.fontSize(10).font("Sans-Italic").fillColor("#000000");
    doc.text("Who exactly is my reader for this one?", PAGE.MARGIN + 2, y);
    y += 14;
    y = addLinedArea(doc, y, 1, { lineSpacing: 26 });
    y += 14;
  });

  y = addTipBox(doc,
    "Before you write anything this week, read your belief out loud. If it does not sound "
    + "certain, add a high modality word from your Word Bank.",
    y, { color: C.ACCENT });

  addPdfFooter(doc, "Persuasive Writing | Week 7 | Year 5/6 Literacy");
  await writePdf(doc, path.join(OUT_DIR, PLAN_RES.fileName));
  console.log("Wrote " + PLAN_RES.name);
}

(async () => {
  await build();
  await buildPlan();
})().catch((err) => {
  console.error(err);
  process.exit(1);
});
