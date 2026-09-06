"use strict";

// Deaf Sports in Australia (Years 5-6 Auslan, Term 4) - Session 2 of 10.
// WHO, WHAT, WHERE: the profile questions, plus the first repair sign.
//   4 production signs: WHO, WHAT, WHERE, AGAIN. Game 2 Find Your Team.
// Unit anchor: Set it up, ask it, watch it, check it.
// Unit variant fixed (literacy/grade56/v0) across all 10 sessions.
// Catch-up: the Do Now retrieves Lesson 1's three signs from images, so a
// student who missed Lesson 1 meets them here before they are needed.

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
  signSlide, signRecallSlide, stepCard, createSignReport, ATTRIBUTION,
} = require("./auslan_signs_lib");

const T = createTheme("literacy", "grade56", weekToVariant(1)); // variant 0, fixed for the unit
const {
  C, FONT_H, FONT_B,
  titleSlide, liSlide, closingSlide, exitTicketSlide,
  addTextOnShape, runSlideDiagnostics,
} = T;

const SESSION = 2;
const TOTAL = 10;
const UNIT_TITLE = "Deaf Sports in Australia";
const ANCHOR = "Set it up, ask it, watch it, check it.";
const FOOTER = `${UNIT_TITLE} | Session ${SESSION} of ${TOTAL} | Years 5-6 Auslan`;
const OUT_DIR = "output/DeafSport_Lesson2_Who_What_Where";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const CARDS_RES = makeSessionResource(SESSION,
  "Team Role Cards",
  "Six teams of four. Two teams share a sport on purpose, so one answer is never enough.");
const RESOURCE_ITEMS = [CARDS_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

const REPORT = createSignReport();

const TEAMS = [
  ["Sam", "Swimming", "Perth"], ["Ali", "Swimming", "Perth"],
  ["Jo", "Swimming", "Perth"], ["Kit", "Swimming", "Perth"],
  ["Ren", "Swimming", "Hobart"], ["Bo", "Swimming", "Hobart"],
  ["Tam", "Swimming", "Hobart"], ["Nia", "Swimming", "Hobart"],
  ["Max", "Cricket", "Adelaide"], ["Eve", "Cricket", "Adelaide"],
  ["Ari", "Cricket", "Adelaide"], ["Lou", "Cricket", "Adelaide"],
  ["Fin", "Basketball", "Sydney"], ["Zia", "Basketball", "Sydney"],
  ["Rue", "Basketball", "Sydney"], ["Dev", "Basketball", "Sydney"],
  ["Ivy", "Tennis", "Melbourne"], ["Cam", "Tennis", "Melbourne"],
  ["Rio", "Tennis", "Melbourne"], ["Gus", "Tennis", "Melbourne"],
  ["Wren", "Athletics", "Brisbane"], ["Ash", "Athletics", "Brisbane"],
  ["Nell", "Athletics", "Brisbane"], ["Ty", "Athletics", "Brisbane"],
];

/* ---------------------------------------------------------------- notes ---- */

const NOTES_TITLE = "Session 2 of 10. Three profile questions and the first repair sign. Team role cards need cutting before this lesson.";

const NOTES_RESOURCES = T.composeGlanceNotes({
  beats: [
    ["SHOW while students settle.", "SAY: One card each, and it stays flat on your desk."],
    ["SET UP: cards in holders or laid flat, never held.", "A student holding a card has no hands to sign with."],
  ],
  prep: [
    "Print and cut the Team Role Cards once; they are laminated and reused.",
    "Deal so every team of four is complete. Materials: mini whiteboards.",
    ATTRIBUTION,
  ],
  tag: "[Setup | Planning | HITS 2]",
});

const NOTES_DONOW = T.composeGlanceNotes({
  answer: "deaf, sport, favourite",
  beats: [
    ["ASK: What does each sign mean?",
      "10 sec think. Cue: Everyone signs it to me on three... one, two, three.",
      "EXPECT: DEAF, SPORT, FAVOURITE"],
    "REVEAL after the boards are scanned, one click.",
    ["SAY: You know one thing about each other. One thing is not a profile.",
      "Today we get three."],
  ],
  prep: [
    "Retrieval from Lesson 1, then the bridge into today.",
    "A student who missed Lesson 1 meets these three here, before they are needed.",
  ],
  tag: "[Do Now | Attention, focus and regulation | HITS 6]",
});

const NOTES_LI = T.composeGlanceNotes({
  beats: [
    ["POINT to each criterion.", "SAY: Three questions today, and one way to fix it when you miss."],
    "SAY: Nobody guesses in this room. We ask again instead.",
  ],
  prep: [
    "SC1 is one question, reachable with the prompt card.",
    "SC2 is what the exit ticket collects. SC3 is the repair move.",
  ],
  tag: "[LI and SC | Planning | HITS 1]",
});

const NOTES_ANCHOR = T.composeGlanceNotes({
  beats: [
    ["SAY: Same four moves. Set it up, ask it, watch it, check it.",
      "What changes today is that the asking has three different questions in it."],
    ["SAY: I set the topic first, so they know what I am asking about.",
      "If I do not, they spend the whole question working out what I mean."],
    "MODEL setting the topic, then asking, at normal pace.",
  ],
  prep: [
    "The anchor is restated in these exact words every week. Do not reword it.",
    "Setting up is the move students skip first when they are nervous.",
  ],
  tag: "[I Do | Explicit teaching | HITS 3]",
});

const NOTES_QUESTIONS = T.composeGlanceNotes({
  answer: "WHO, WHAT, WHERE",
  beats: [
    "MODEL each question form from the school reference, six times each.",
    ["SAY: Three questions, and they come in this order all term.",
      "An order you do not think about is one you can use when nervous."],
    "SAY: Watch the whole answer even when you think you know it.",
    "CHECK all three look visibly different, not one form re-faced.",
  ],
  trap: ["WHERE collapsing into WHAT under time pressure.",
    "Fix: sign three, students point to the right heading."],
  stretch: "add a fourth question of their own and say why.",
  help: "the three question words in English, face up on the desk.",
  prep: [
    "CHECK GRAMMAR: confirm the question form and where the question sign sits.",
    "WHAT has one Signbank entry and its first sense is the interactive 'What?!'.",
    "Confirm all three from the school reference before teaching them.",
  ],
  tag: "[I Do | Explicit teaching | HITS 3, 4]",
});

const NOTES_AGAIN = T.composeGlanceNotes({
  answer: "AGAIN",
  beats: [
    ["GET IT WRONG: ask a question, then look down at a card mid-answer.",
      "SAY: I missed that because I looked away."],
    ["SAY: I am not going to guess. I am going to ask for it again.",
      "It costs me four seconds and nothing else."],
    "MODEL the repair, then let the exchange carry on rather than restarting it.",
  ],
  trap: ["nodding along and guessing rather than repairing.",
    "Fix: freeze the pair, student asks again, then answers."],
  prep: [
    "AGAIN gets the same thing signed a second time.",
    "Lesson 3 adds WHAT MEAN? for when a repeat would not help.",
  ],
  tag: "[I Do | Explicit teaching | HITS 3]",
});

const NOTES_GAME = T.composeGlanceNotes({
  answer: "four people who share a name set, a sport and a city",
  beats: [
    ["SET UP: one card each, flat on the desk.",
      "Half the class moves at a time.",
      "Cards never travel in a hand."],
    "SAY: Two teams share a sport, so one answer is never enough. Check all three.",
    ["CIRCULATE: watch for students deciding on the sport alone.",
      "That is the error the deal was built to expose."],
    "COLLECT: a complete team sits down together and keeps checking new arrivals.",
  ],
  trap: ["joining the first person with a matching sport.",
    "Fix: send them back to ask the city question."],
  prep: [
    "Game 2 Find Your Team. Six teams of four, two sharing a sport.",
    "A complete team sits down so standing clusters stop blocking sight lines.",
  ],
  tag: "[We Do | Collaborative learning | HITS 5]",
});

const NOTES_CHECK = T.composeGlanceNotes({
  answer: "WHERE",
  beats: [
    ["ASK: Show me the question for where.",
      "5 sec think. Cue: Everyone signs it to me on three... one, two, three.",
      "EXPECT: WHERE"],
    ["SCAN the room, back row first.",
      "80%+ -> go to You Do with all three questions in play.",
      "Less -> three headings on the board, students point to the right one.",
      "Six times fast, then re-ask with the same routine."],
    ["FOLLOW UP one student.",
      "SAY: Do you agree with that one? Add one thing."],
  ],
  trap: ["WHERE grabbed as the nearest question sign under pressure.",
    "Fix: sort by heading, then re-ask with the same routine."],
  prep: [
    "The decision point that changes Lesson 3.",
    "If it fails, Lesson 3 opens by re-sorting the three questions.",
  ],
  tag: "[CFU | Evaluating impact | HITS 7]",
});

const NOTES_YOUDO = T.composeGlanceNotes({
  beats: [
    ["SAY: Voices off. Ask your partner all three questions.",
      "Write the three answers on your board."],
    "SAY: Pencils down while your partner signs. Watch it all, then write.",
    ["CIRCULATE: look for the pencil going down.",
      "A student writing mid-answer has missed the answer."],
  ],
  stretch: "ask a fourth question of their own and be ready to explain the choice.",
  help: "the three question words in English, face up, so recall is the sign only.",
  prep: [
    "Different content from the game: card identities there, the real partner here.",
    "Year 5 uses the prompt card. Year 6 works with it face down.",
  ],
  tag: "[You Do | Explicit teaching | HITS 10]",
});

const NOTES_EXIT = T.composeGlanceNotes({
  answer: "any one of the three answers about their partner",
  beats: [
    "COLLECT at the door, signed to you, student chooses the question.",
    "Two piles: answered the question asked, or answered a different one.",
    "Then the Unit Progress Card, Lesson 2 row, three ticks.",
  ],
  prep: [
    "Assesses SC2. Answering a different question is the thing to watch.",
    "It usually means the question signs are still blurring together.",
  ],
  tag: "[Exit ticket | Evaluating impact | HITS 7]",
});

const NOTES_CLOSING = T.composeGlanceNotes({
  beats: [
    "SAY: Turn and tell your partner one thing you learned about somebody new.",
    "SAY: Next week we add when, and we start counting years.",
  ],
  prep: [
    "Protocol practised today: the table tap, for when a wave will not reach.",
    "Running late: cut the game to one round. Never cut AGAIN.",
  ],
  tag: "[Closing | Planning | HITS 9]",
});

/* ---------------------------------------------------------------- deck ----- */

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // 1 Title
  titleSlide(pres, UNIT_TITLE, "Session 2: Who, what, where",
    `Years 5-6 Auslan | Session ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // 2 Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // 3 Do Now - retrieval reveal of Lesson 1's signs
  REPORT.add(signRecallSlide(T, pres, {
    title: "Do Now: what does each one mean?",
    lead: "Sign the meaning back to me. Answers come up on the click.",
    signs: [
      { gloss: "DEAF", meaning: "deaf" },
      { gloss: "SPORT", meaning: "sport" },
      { gloss: "FAVOURITE", meaning: "favourite" },
    ],
    notes: NOTES_DONOW, footer: FOOTER,
  }).results);

  // 4 LI and SC
  liSlide(pres,
    "I am learning to ask who, what and where, and to ask for a repeat when I miss something.",
    [
      "I can ask one of the three profile questions.",
      "I can ask all three and record the answers.",
      "I can use AGAIN when I miss an answer instead of guessing.",
    ],
    NOTES_LI, FOOTER);

  // 5 I Do - anchor restated, then set the topic first
  const anchorSlide = pres.addSlide();
  anchorSlide.background = { color: C.BG_LIGHT };
  T.addTopBar(anchorSlide, C.PRIMARY);
  T.addTitle(anchorSlide, "Set it up first");
  addTextOnShape(anchorSlide, ANCHOR, {
    x: 0.5, y: 1.45, w: 9.0, h: 0.72, rectRadius: 0.1, fill: { color: C.PRIMARY },
  }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
  addTextOnShape(anchorSlide, "Topic first, then the question", {
    x: 1.4, y: 2.5, w: 7.2, h: 0.7, rectRadius: 0.1, fill: { color: C.SECONDARY },
  }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
  anchorSlide.addText("If you do not set the topic, your partner spends the whole question working out what you mean.", {
    x: 1.4, y: 3.45, w: 7.2, h: 0.8,
    fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL, align: "center", margin: 0,
  });
  T.addFooter(anchorSlide, FOOTER);
  anchorSlide.addNotes(NOTES_ANCHOR);
  runSlideDiagnostics(anchorSlide, pres);

  // 6 I Do - the three profile questions
  REPORT.add(signSlide(T, pres, {
    title: "Three questions, always this order",
    lead: "Watch me first. Then copy. Six times each.",
    signs: [
      { gloss: "WHO", meaning: "who" },
      { gloss: "WHAT", meaning: "what" },
      { gloss: "WHERE", meaning: "where" },
    ],
    notes: NOTES_QUESTIONS, footer: FOOTER,
  }).results);

  // 7 I Do - the repair sign
  const againSlide = pres.addSlide();
  againSlide.background = { color: C.BG_LIGHT };
  T.addTopBar(againSlide, C.ALERT);
  T.addTitle(againSlide, "When you miss it, ask again");
  REPORT.add(require("./auslan_signs_lib").signCardRow(T, againSlide, [
    { gloss: "AGAIN", meaning: "again" },
  ], { x: 0.6, w: 3.6, y: 1.5, bottom: 5.0 }));
  stepCard(T, againSlide, {
    title: "Nobody guesses in this room",
    steps: [
      "You missed part of the answer",
      "Do not nod and hope",
      "Sign: again",
      "Watch the whole thing this time",
    ],
    x: 4.6, y: 1.9, w: 4.9, h: 2.4, strip: C.ALERT,
  });
  T.addFooter(againSlide, FOOTER);
  againSlide.addNotes(NOTES_AGAIN);
  runSlideDiagnostics(againSlide, pres);

  // 8 We Do - Game 2
  const gameSlide = pres.addSlide();
  gameSlide.background = { color: C.BG_LIGHT };
  T.addTopBar(gameSlide, C.SUCCESS);
  T.addTitle(gameSlide, "Game 2 Find Your Team");
  stepCard(T, gameSlide, {
    title: "Find the other three on your team",
    steps: [
      "Card flat on your desk. Never in your hand",
      "Wait for eyes, then ask who, what and where",
      "Two teams share a sport, so check all three",
      "A full team sits down together",
    ],
    x: 0.5, y: 1.5, w: 5.4, h: 2.9, strip: C.SUCCESS,
  });
  addTextOnShape(gameSlide, "Check all three", {
    x: 6.1, y: 1.6, w: 3.4, h: 0.7, rectRadius: 0.1, fill: { color: C.ALERT },
  }, { fontSize: 19, fontFace: FONT_H, color: C.WHITE, bold: true });
  ["Who are you?", "What sport?", "Where from?"].forEach((q, i) => {
    addTextOnShape(gameSlide, q, {
      x: 6.1, y: 2.5 + i * 0.62, w: 3.4, h: 0.5, rectRadius: 0.08,
      fill: { color: C.BG_CARD }, line: { color: C.SECONDARY, width: 1 },
    }, { fontSize: 15, fontFace: FONT_B, color: C.CHARCOAL });
  });
  T.addFooter(gameSlide, FOOTER);
  gameSlide.addNotes(NOTES_GAME);
  runSlideDiagnostics(gameSlide, pres);

  // 9 Check it
  const checkSlide = pres.addSlide();
  checkSlide.background = { color: C.BG_LIGHT };
  T.addTopBar(checkSlide, C.ALERT);
  T.addTitle(checkSlide, "Check it");
  addTextOnShape(checkSlide, "Everyone signs it to me on three", {
    x: 1.1, y: 1.85, w: 7.8, h: 0.85, rectRadius: 0.12, fill: { color: C.ALERT },
  }, { fontSize: 26, fontFace: FONT_H, color: C.WHITE, bold: true });
  checkSlide.addText("Show me the question for: where", {
    x: 1.1, y: 2.95, w: 7.8, h: 0.7,
    fontSize: 30, fontFace: FONT_H, color: C.PRIMARY, bold: true, align: "center", margin: 0,
  });
  checkSlide.addText("Think first. Hands ready. One, two, three.", {
    x: 1.1, y: 3.8, w: 7.8, h: 0.5,
    fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL, align: "center", margin: 0,
  });
  T.addFooter(checkSlide, FOOTER);
  checkSlide.addNotes(NOTES_CHECK);
  runSlideDiagnostics(checkSlide, pres);

  // 10 You Do
  const youSlide = pres.addSlide();
  youSlide.background = { color: C.BG_LIGHT };
  T.addTopBar(youSlide, C.SUCCESS);
  T.addTitle(youSlide, "You Do: interview your partner");
  stepCard(T, youSlide, {
    title: "All three questions, then swap",
    steps: [
      "Ask who, what and where",
      "Pencils down while your partner signs",
      "Watch it all, then write the answer",
      "Use again if you miss one",
    ],
    x: 0.5, y: 1.5, w: 5.4, h: 2.9, strip: C.SUCCESS,
  });
  addTextOnShape(youSlide, "Voices off", {
    x: 6.1, y: 1.6, w: 3.4, h: 0.9, rectRadius: 0.12, fill: { color: C.ALERT },
  }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
  addTextOnShape(youSlide, "Watch, then write", {
    x: 6.1, y: 2.75, w: 3.4, h: 0.75, rectRadius: 0.1, fill: { color: C.SECONDARY },
  }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });
  youSlide.addText("Never both at once", {
    x: 6.1, y: 3.65, w: 3.4, h: 0.4,
    fontSize: 15, fontFace: FONT_B, color: C.CHARCOAL, align: "center", margin: 0,
  });
  T.addFooter(youSlide, FOOTER);
  youSlide.addNotes(NOTES_YOUDO);
  runSlideDiagnostics(youSlide, pres);

  // 11 Exit ticket
  exitTicketSlide(pres,
    ["Sign me one answer about your partner. You choose which question."],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // 12 Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner one thing you learned about somebody new today.",
      scItems: [
        "I can ask one of the three profile questions.",
        "I can ask all three and record the answers.",
        "I can use AGAIN when I miss an answer instead of guessing.",
      ],
      selfAssessment: "Unit Progress Card, Lesson 2 row",
    },
    NOTES_CLOSING, FOOTER);

  const pptxPath = path.join(OUT_DIR, `Deaf Sports Session ${SESSION} Who What Where.pptx`);
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  /* ------------------------------------------------------------- pdf ------ */

  const doc = createPdf({ title: CARDS_RES.name });
  let y = addPdfHeader(doc, CARDS_RES.name, {
    subtitle: "Six teams of four. Cut along the lines and laminate.",
    color: C.PRIMARY,
    lessonInfo: `Session ${SESSION} | Game 2 Find Your Team | one card per student`,
  });

  y = addTipBox(doc,
    "Teams 1 and 2 both play swimming on purpose. A student who stops at the sport question joins the wrong team, which is the whole point of the game. Deal so every team of four is complete.",
    y, { color: C.ACCENT });

  y = addSectionHeading(doc, "Cut these out", y, { color: C.SECONDARY });

  const cardW = 165;
  const cardH = 92;
  const cols = 3;
  const startX = 50;
  TEAMS.forEach(([name, sport, city], i) => {
    const col = i % cols;
    const row = Math.floor(i / cols) % 6;
    if (i > 0 && i % (cols * 6) === 0) {
      doc.addPage();
      y = 60;
    }
    const cx = startX + col * cardW;
    const cy = y + row * cardH;
    doc.save();
    doc.rect(cx, cy, cardW - 8, cardH - 8).lineWidth(0.8).dash(3, { space: 2 })
      .strokeColor("#" + C.MUTED).stroke();
    doc.undash();
    doc.font("Sans-Bold").fontSize(15).fillColor("#" + C.PRIMARY)
      .text(name, cx + 12, cy + 14, { width: cardW - 32 });
    doc.font("Sans").fontSize(11).fillColor("#" + C.CHARCOAL)
      .text("Sport: " + sport, cx + 12, cy + 38, { width: cardW - 32 });
    doc.text("City: " + city, cx + 12, cy + 56, { width: cardW - 32 });
    doc.restore();
  });

  addPdfFooter(doc, `Session ${SESSION} | Team Role Cards | Years 5-6 Auslan`);
  await writePdf(doc, path.join(OUT_DIR, CARDS_RES.fileName));
  console.log("PDF written: " + CARDS_RES.fileName);

  REPORT.print(`Session ${SESSION}`);
  console.log(`Session ${SESSION} build complete.`);
}

build().catch((err) => { console.error(err); process.exit(1); });
