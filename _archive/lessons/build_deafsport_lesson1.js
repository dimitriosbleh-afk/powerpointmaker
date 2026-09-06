"use strict";

// Deaf Sports in Australia (Years 5-6 Auslan, Term 4) - Session 1 of 10.
// MEETING A SIGNER: the interview routine.
//   Establishes the voice-off routines, the attention protocol (waving) and the
//   response cue the other nine lessons run on. Deliberately the lightest lesson
//   in the unit: 3 production signs (DEAF, SPORT, FAVOURITE), no printed cards.
// Unit anchor: Set it up, ask it, watch it, check it.
// Unit variant fixed (literacy/grade56/v0, from Term 4 week 1) across all 10.
// Catch-up: nothing in this lesson assumes prior Auslan beyond fingerspelling a
// name, and the Do Now rebuilds that before it is needed.

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading, addBodyText,
  addTipBox, addPdfFooter, addResourceSlide,
  getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");
const {
  signSlide, stepCard, createSignReport, ATTRIBUTION,
} = require("./auslan_signs_lib");

const T = createTheme("literacy", "grade56", weekToVariant(1)); // variant 0, fixed for the unit
const {
  C, FONT_H, FONT_B,
  titleSlide, liSlide, contentSlide, closingSlide, exitTicketSlide,
  addTextOnShape, addCard,
  clickBuild, runSlideDiagnostics,
} = T;

const SESSION = 1;
const TOTAL = 10;
const UNIT_TITLE = "Deaf Sports in Australia";
const ANCHOR = "Set it up, ask it, watch it, check it.";
const FOOTER = `${UNIT_TITLE} | Session ${SESSION} of ${TOTAL} | Years 5-6 Auslan`;
const OUT_DIR = "output/DeafSport_Lesson1_Meeting_A_Signer";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const PROGRESS_RES = makeSessionResource(SESSION,
  "Unit Progress Card",
  "One per student for the whole unit. Print once, keep in a class tub, three ticks a week.");
const RESOURCE_ITEMS = [PROGRESS_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

const REPORT = createSignReport();

/* ---------------------------------------------------------------- notes ---- */

const NOTES_TITLE = "Session 1 of 10. The lightest lesson in the unit on purpose: it builds the routines the other nine run on.";

const NOTES_RESOURCES = T.composeGlanceNotes({
  beats: [
    ["SHOW while students settle.", "SAY: One card each, and it lives in the tub all term."],
    ["COLLECT the cards at the door every week.", "They are the only record students keep."],
  ],
  prep: [
    "Print the Unit Progress Card once, one per student, on card.",
    "Materials: cleared floor space, a timer the class can see.",
    ATTRIBUTION,
  ],
  tag: "[Setup | Planning | HITS 2]",
});

const NOTES_DONOW = T.composeGlanceNotes({
  answer: "every student fingerspells their own name to two people",
  beats: [
    ["SAY: Stand up. Find two people.", "Fingerspell your own name to each one."],
    ["SAY: Some of you may remember this.", "If it feels new, that is okay. We will build it together."],
    "TIME: 3 minutes, then stop them where they stand.",
    ["SAY: You just told two people who you are.", "Today we add asking them something back."],
  ],
  prep: [
    "Retrieval plus bridge. Nothing here is new.",
    "It exists so every student starts the lesson having succeeded once.",
  ],
  tag: "[Do Now | Attention, focus and regulation | HITS 6]",
});

const NOTES_LI = T.composeGlanceNotes({
  beats: [
    ["POINT to each criterion.", "SAY: Three things, and the first is just waiting for eyes."],
    "SAY: By the end, everyone has asked someone what sport they like best.",
  ],
  prep: [
    "SC1 is reachable by every student with support.",
    "SC2 is what the exit ticket collects. SC3 is the stretch.",
  ],
  tag: "[LI and SC | Planning | HITS 1]",
});

const NOTES_ANCHOR = T.composeGlanceNotes({
  beats: [
    ["SAY: Four moves, every week this term.",
      "You set it up, you ask it, you watch it, you check it.",
      "Nothing on that list is optional."],
    ["POINT to move 1.", "SAY: Before any signing happens, I need their eyes."],
    ["SAY: I wave, or I tap the table, and then I wait.",
      "I do not start until they are looking."],
    "MODEL the wave, then wait visibly for a student to look up.",
  ],
  prep: [
    "The anchor is restated in these exact words in every I Do.",
    "Do not reword it. A student who missed a week finds today through it.",
  ],
  tag: "[I Do | Explicit teaching | HITS 3]",
});

const NOTES_SIGNS = T.composeGlanceNotes({
  answer: "DEAF, SPORT, FAVOURITE",
  beats: [
    "MODEL each sign six times, slowly, then at normal pace.",
    ["SAY: Watch my face on this one.",
      "My hands are asking, but my face is doing half the work."],
    "SAY: I watch the whole answer, not the first bit then off thinking.",
    ["GET IT WRONG: ask a question at someone looking away.",
      "SAY: I got nothing back, because I skipped the first move."],
  ],
  trap: ["waving at a back and starting anyway.",
    "Fix: hold the wave until eyes arrive, student redoes it."],
  stretch: "add a second question of their own after the sport question.",
  help: "a three-sport picture card, so the answer is a point plus a sign.",
  prep: [
    "3 production signs only, the low end of the budget.",
    "Lesson 1 is buying routines, not vocabulary.",
    "Rehearse FAVOURITE first; Signbank also lists adoration and obsession for it.",
  ],
  tag: "[I Do | Explicit teaching | HITS 3, 4]",
});

const NOTES_GAME = T.composeGlanceNotes({
  answer: "a normie fingerspells their real name, a Bob fingerspells B-O-B",
  beats: [
    ["SET UP: clear the furniture. Eyes closed.",
      "Tap one or two shoulders. Those are the Original Bobs."],
    "SAY: No running, no voices, and do not dodge anyone who approaches you.",
    "TIME: 2 minutes on the visible timer.",
    ["CIRCULATE: watch for waves landing on backs, not faces.",
      "That is the thing to fix live."],
    "COLLECT: ask who is a Bob, then who is a normie.",
  ],
  trap: ["starting to sign before the partner looked up.",
    "Fix: freeze the pair, redo the opening."],
  prep: [
    "Game 1 Bob Virus, from the unit game bank.",
    "Voice off is the mechanic, not a rule on top of it.",
    "A fingerspelled name only counts if the partner watched it land.",
  ],
  tag: "[We Do | Collaborative learning | HITS 5]",
});

const NOTES_CHECK = T.composeGlanceNotes({
  answer: "SPORT",
  beats: [
    ["ASK: Show me the sign for sport.",
      "5 sec think. Cue: Everyone signs it to me on three... one, two, three.",
      "EXPECT: SPORT"],
    ["SCAN the room, back row first.",
      "80%+ -> go to You Do and add the sport question.",
      "Less -> two facing lines, mirror the person opposite, swap, re-ask."],
    ["FOLLOW UP one student.",
      "SAY: You waited for her eyes before you started. How did you know?"],
  ],
  trap: ["a handshape copied from a neighbour, not from you.",
    "Fix: re-model facing them, student redoes it."],
  prep: [
    "The decision point that changes Lesson 2.",
    "If it fails, Lesson 2 opens with the attention protocol, not new questions.",
  ],
  tag: "[CFU | Evaluating impact | HITS 7]",
});

const NOTES_YOUDO = T.composeGlanceNotes({
  beats: [
    ["SAY: Voices off. Three different people.",
      "Ask what sport they like best and remember one answer."],
    "TIME: 12 minutes. This stage belongs to them, so do not script it.",
    ["CIRCULATE: look for the wait, not the sign.",
      "The wait is what is new today."],
  ],
  stretch: "find a sport nobody else named, and say how you found that person.",
  help: "two people instead of three, with the three-sport picture card on the desk.",
  prep: [
    "Different content from the game: names there, an opinion here.",
    "Year 5 asks two people. Year 6 asks three and adds a follow-up.",
  ],
  tag: "[You Do | Explicit teaching | HITS 10]",
});

const NOTES_EXIT = T.composeGlanceNotes({
  answer: "one person's name, fingerspelled, plus that person's favourite sport",
  beats: [
    "COLLECT at the door, one student at a time, signed to you.",
    "Two piles in your head: got both, or got one or neither.",
    "Nothing is written and nothing is marked.",
    "Then the Unit Progress Card, Lesson 1 row, three ticks.",
  ],
  prep: [
    "Assesses SC2.",
    "If the ticket and the tick disagree, the ticket is the evidence.",
  ],
  tag: "[Exit ticket | Evaluating impact | HITS 7]",
});

const NOTES_CLOSING = T.composeGlanceNotes({
  beats: [
    "SAY: Turn and tell your partner one thing you found out today.",
    "SAY: Next week we add three more questions to the one you used.",
  ],
  care: "talk privately to any Deaf or hard of hearing student before this lesson.",
  prep: [
    "Offer them a prepared answer, and never make them the live example.",
    "Protocol practised today: waving. It is required in every remaining lesson.",
  ],
  tag: "[Closing | Planning | HITS 9]",
});

/* ---------------------------------------------------------------- deck ----- */

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // 1 Title
  titleSlide(pres, UNIT_TITLE, "Session 1: Meeting a signer",
    `Years 5-6 Auslan | Session ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // 2 Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // 3 Do Now
  contentSlide(pres, "Do Now", C.SECONDARY, "Tell two people who you are",
    [
      "Stand up and find two people",
      "Fingerspell your own name to each one",
      "Watch their whole name before you swap",
    ],
    NOTES_DONOW, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.15, { strip: C.SECONDARY });
      slide.addText("Voices off", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.2, w: lg.rightW - 0.4, h: 0.4,
        fontSize: 20, fontFace: FONT_H, color: C.ALERT, bold: true, align: "center", margin: 0,
      });
      slide.addText("Some of you may remember this.\nIf it feels new, that is okay.\nWe will build it together.", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.72, w: lg.rightW - 0.4, h: 1.2,
        fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, align: "center", margin: 0, lineSpacingMultiple: 1.2,
      });
    });

  // 4 LI and SC
  liSlide(pres,
    "I am learning to open a signed exchange and ask one question about sport.",
    [
      "I can wave or tap to get someone's attention and wait for their eyes.",
      "I can ask one person about their favourite sport and watch their whole answer.",
      "I can tell the class one thing I found out.",
    ],
    NOTES_LI, FOOTER);

  // 5 I Do - the unit anchor, built one move at a time
  const anchorSlide = pres.addSlide();
  anchorSlide.background = { color: C.BG_LIGHT };
  T.addTopBar(anchorSlide, C.PRIMARY);
  T.addTitle(anchorSlide, "Four moves, every week");
  addTextOnShape(anchorSlide, ANCHOR, {
    x: 0.5, y: 1.5, w: 9.0, h: 0.72, rectRadius: 0.1,
    fill: { color: C.PRIMARY },
  }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
  const MOVES = [
    ["Set it up", "Get their eyes first"],
    ["Ask it", "One clear question"],
    ["Watch it", "The whole answer"],
    ["Check it", "Did I understand?"],
  ];
  const mW = (9.0 - 0.24 * 3) / 4;
  clickBuild(anchorSlide, MOVES.map(([head, sub], i) => () => {
    const mx = 0.5 + i * (mW + 0.24);
    addTextOnShape(anchorSlide, String(i + 1), {
      x: mx + mW / 2 - 0.22, y: 2.52, w: 0.44, h: 0.44, rectRadius: 0.22,
      fill: { color: C.SECONDARY },
    }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });
    addTextOnShape(anchorSlide, head, {
      x: mx, y: 3.08, w: mW, h: 0.5, rectRadius: 0.08,
      fill: { color: C.BG_CARD }, line: { color: C.SECONDARY, width: 1 },
    }, { fontSize: 16, fontFace: FONT_H, color: C.PRIMARY, bold: true });
    anchorSlide.addText(sub, {
      x: mx, y: 3.64, w: mW, h: 0.6, fontSize: 12.5, fontFace: FONT_B,
      color: C.CHARCOAL, align: "center", margin: 0,
    });
  }));
  T.addFooter(anchorSlide, FOOTER);
  anchorSlide.addNotes(NOTES_ANCHOR);

  // 6 I Do - the three new signs
  REPORT.add(signSlide(T, pres, {
    title: "Three signs for today",
    lead: "Watch me first. Then copy. Six times each.",
    signs: [
      { gloss: "DEAF", meaning: "deaf" },
      { gloss: "SPORT", meaning: "sport" },
      { gloss: "FAVOURITE", meaning: "favourite" },
    ],
    notes: NOTES_SIGNS, footer: FOOTER,
  }).results);

  // 7 We Do - Game 1 Bob Virus
  const gameSlide = pres.addSlide();
  gameSlide.background = { color: C.BG_LIGHT };
  T.addTopBar(gameSlide, C.SUCCESS);
  T.addTitle(gameSlide, "Game 1 Bob Virus");
  stepCard(T, gameSlide, {
    title: "How to play",
    steps: [
      "Move around and ask someone for their name",
      "A normie fingerspells their real name",
      "A Bob fingerspells B-O-B",
      "Meet a Bob and you become a Bob",
      "Two minutes on the timer",
    ],
    x: 0.5, y: 1.5, w: 5.4, h: 3.1, strip: C.SUCCESS,
  });
  addTextOnShape(gameSlide, "Rules", {
    x: 6.1, y: 1.5, w: 3.4, h: 0.46, rectRadius: 0.08, fill: { color: C.ALERT },
  }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
  ["No running", "No voices", "Do not dodge anyone", "Wait for their eyes"].forEach((r, i) => {
    addTextOnShape(gameSlide, r, {
      x: 6.1, y: 2.12 + i * 0.62, w: 3.4, h: 0.5, rectRadius: 0.08,
      fill: { color: C.BG_CARD }, line: { color: C.ALERT, width: 1 },
    }, { fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL });
  });
  T.addFooter(gameSlide, FOOTER);
  gameSlide.addNotes(NOTES_GAME);
  runSlideDiagnostics(gameSlide, pres);

  // 8 Check it - the decision point
  const checkSlide = pres.addSlide();
  checkSlide.background = { color: C.BG_LIGHT };
  T.addTopBar(checkSlide, C.ALERT);
  T.addTitle(checkSlide, "Check it");
  addTextOnShape(checkSlide, "Everyone signs it to me on three", {
    x: 1.1, y: 1.85, w: 7.8, h: 0.85, rectRadius: 0.12, fill: { color: C.ALERT },
  }, { fontSize: 26, fontFace: FONT_H, color: C.WHITE, bold: true });
  checkSlide.addText("Sign: sport", {
    x: 1.1, y: 2.95, w: 7.8, h: 0.7,
    fontSize: 32, fontFace: FONT_H, color: C.PRIMARY, bold: true, align: "center", margin: 0,
  });
  checkSlide.addText("Think first. Hands ready. One, two, three.", {
    x: 1.1, y: 3.75, w: 7.8, h: 0.5,
    fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL, align: "center", margin: 0,
  });
  T.addFooter(checkSlide, FOOTER);
  checkSlide.addNotes(NOTES_CHECK);
  runSlideDiagnostics(checkSlide, pres);

  // 9 You Do
  const youSlide = pres.addSlide();
  youSlide.background = { color: C.BG_LIGHT };
  T.addTopBar(youSlide, C.SUCCESS);
  T.addTitle(youSlide, "You Do: ask three people");
  stepCard(T, youSlide, {
    title: "With three different people",
    steps: [
      "Wave or tap. Wait for their eyes",
      "Ask what sport they like best",
      "Watch their whole answer",
      "Remember one answer to tell the class",
    ],
    x: 0.5, y: 1.5, w: 5.4, h: 2.9, strip: C.SUCCESS,
  });
  addTextOnShape(youSlide, "Voices off", {
    x: 6.1, y: 1.5, w: 3.4, h: 0.9, rectRadius: 0.12, fill: { color: C.ALERT },
  }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
  addTextOnShape(youSlide, "3", {
    x: 7.35, y: 2.7, w: 0.9, h: 0.9, rectRadius: 0.45, fill: { color: C.SUCCESS },
  }, { fontSize: 34, fontFace: FONT_H, color: C.WHITE, bold: true });
  youSlide.addText("people", {
    x: 6.1, y: 3.72, w: 3.4, h: 0.4,
    fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL, align: "center", margin: 0,
  });
  T.addFooter(youSlide, FOOTER);
  youSlide.addNotes(NOTES_YOUDO);
  runSlideDiagnostics(youSlide, pres);

  // 10 Exit ticket
  exitTicketSlide(pres,
    [
      "Sign me one person's name and their favourite sport.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // 11 Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner the one thing you found out about somebody today.",
      scItems: [
        "I can wave or tap to get someone's attention and wait for their eyes.",
        "I can ask one person about their favourite sport and watch their whole answer.",
        "I can tell the class one thing I found out.",
      ],
      selfAssessment: "Unit Progress Card, Lesson 1 row",
    },
    NOTES_CLOSING, FOOTER);

  const pptxPath = path.join(OUT_DIR, `Deaf Sports Session ${SESSION} Meeting A Signer.pptx`);
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  /* ------------------------------------------------------------- pdf ------ */

  const doc = createPdf({ title: PROGRESS_RES.name });
  let y = addPdfHeader(doc, PROGRESS_RES.name, {
    subtitle: "Deaf Sports in Australia | Years 5-6 Auslan | Term 4",
    color: C.PRIMARY,
    lessonInfo: "One card per student for the whole unit. Keep it in the class tub.",
  });

  y = addBodyText(doc, "Name: ______________________________________________", y);
  y = addSectionHeading(doc, "Tick a box each week if you can do it", y, { color: C.SECONDARY });
  y = addBodyText(doc,
    "The three I can statements are on the board every lesson. If you cannot tick one, that is useful information, not a problem.",
    y);

  const rows = [
    "Lesson 1", "Lesson 2", "Lesson 3", "Lesson 4", "Lesson 5",
    "Lesson 6", "Lesson 7", "Lesson 8", "Lesson 9", "Lesson 10",
  ];
  const tableX = 50;
  const labelW = 130;
  const boxW = 110;
  const rowH = 30;
  const headH = 34;

  doc.save();
  doc.font("Sans-Bold").fontSize(10).fillColor("#" + C.WHITE);
  doc.rect(tableX, y, labelW + boxW * 3, headH).fill("#" + C.PRIMARY);
  doc.fillColor("#" + C.WHITE);
  doc.text("Lesson", tableX + 8, y + 11, { width: labelW - 16 });
  ["First one", "Second one", "Third one"].forEach((h, i) => {
    doc.text(h, tableX + labelW + i * boxW + 8, y + 11, { width: boxW - 16, align: "center" });
  });
  doc.restore();
  y += headH;

  rows.forEach((label, r) => {
    doc.save();
    doc.rect(tableX, y, labelW + boxW * 3, rowH)
      .lineWidth(0.8).strokeColor("#" + C.MUTED).stroke();
    doc.moveTo(tableX + labelW, y).lineTo(tableX + labelW, y + rowH).stroke();
    for (let i = 1; i < 3; i += 1) {
      doc.moveTo(tableX + labelW + i * boxW, y).lineTo(tableX + labelW + i * boxW, y + rowH).stroke();
    }
    doc.font("Sans").fontSize(11).fillColor("#" + C.CHARCOAL);
    doc.text(label, tableX + 8, y + 10, { width: labelW - 16 });
    for (let i = 0; i < 3; i += 1) {
      const bx = tableX + labelW + i * boxW + boxW / 2 - 8;
      doc.rect(bx, y + 7, 16, 16).lineWidth(1).strokeColor("#" + C.SECONDARY).stroke();
    }
    doc.restore();
    y += rowH;
  });

  y += 14;
  addTipBox(doc,
    "Teacher: collect these at the door each week and keep them in the tub. They are never marked. If a student's tick and their exit ticket disagree, the exit ticket is the evidence.",
    y, { color: C.SECONDARY });

  addPdfFooter(doc, `${UNIT_TITLE} | Unit Progress Card | Years 5-6 Auslan`);
  await writePdf(doc, path.join(OUT_DIR, PROGRESS_RES.fileName));
  console.log("PDF written: " + PROGRESS_RES.fileName);

  REPORT.print(`Session ${SESSION}`);
  console.log(`Session ${SESSION} build complete.`);
}

build().catch((err) => { console.error(err); process.exit(1); });
