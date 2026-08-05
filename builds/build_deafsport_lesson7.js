"use strict";

// Deaf Sports in Australia (Years 5-6 Auslan, Term 4) - Session 7 of 10.
// THE TWO MINUTE INTERVIEW: four turns, recorded.
//   4 production signs: PRACTISE, PROUD, HISTORY, COMMUNITY.
//   Assessment piece 2 (recorded paired interview) runs in the You Do.
//   THE FORK CHECK runs the same afternoon. See the unit document 9.8.
// Stage minutes differ this session: I Do 10, You Do 20, for the recording.
// PRACTISE uses PRACTISE_3.jpg: the default is LEARNER/TRAINEE, not practise.

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

const T = createTheme("literacy", "grade56", weekToVariant(1));
const {
  C, FONT_H, FONT_B,
  titleSlide, liSlide, closingSlide, exitTicketSlide,
  addTextOnShape, clickBuild, runSlideDiagnostics,
} = T;

const SESSION = 7;
const TOTAL = 10;
const UNIT_TITLE = "Deaf Sports in Australia";
const ANCHOR = "Set it up, ask it, watch it, check it.";
const FOOTER = `${UNIT_TITLE} | Session ${SESSION} of ${TOTAL} | Years 5-6 Auslan`;
const OUT_DIR = "output/DeafSport_Lesson7_The_Two_Minute_Interview";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const PROMPT_RES = makeSessionResource(SESSION,
  "Interview Prompt Card",
  "One per pair, laid flat between them. Laminated; reused if branch B runs.");
const RESOURCE_ITEMS = [PROMPT_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

const REPORT = createSignReport();

/* ---------------------------------------------------------------- notes ---- */

const NOTES_TITLE = "Session 7 of 10. Assessment piece 2 is recorded today. Check filming consent before the lesson, not after. The fork check runs this afternoon.";

const NOTES_RESOURCES = T.composeGlanceNotes({
  beats: [
    ["SHOW while students settle.", "SAY: One card per pair, flat on the desk between you."],
    "CHECK devices charged, assignment visible, consent list read. Before you start.",
  ],
  prep: [
    "Google Classroom: one assignment per class, file upload on, due end of lesson.",
    "A student without consent does the same interview live with you, arranged in advance.",
    ATTRIBUTION,
  ],
  tag: "[Setup | Planning | HITS 2]",
});

const NOTES_DONOW = T.composeGlanceNotes({
  answer: "who, what, where, when",
  beats: [
    ["ASK: Which question is each one?",
      "10 sec think. Cue: Everyone signs it to me on three... one, two, three.",
      "EXPECT: WHO, WHAT, WHERE, WHEN"],
    "REVEAL after the room has answered, one click.",
    ["SAY: Four questions, four weeks apart, all in your hands at once.",
      "Today they run as one conversation, not four separate ones."],
  ],
  prep: [
    "Widest retrieval yet: all four question signs, plus one Lesson 5 modification.",
    "The sequence has to be automatic before a camera is turned on.",
  ],
  tag: "[Do Now | Attention, focus and regulation | HITS 6]",
});

const NOTES_LI = T.composeGlanceNotes({
  beats: [
    ["POINT to each criterion.", "SAY: Four turns today, and one question that is yours."],
    "SAY: By the end everyone has asked two questions and watched both answers.",
  ],
  prep: [
    "SC1 is two turns, reachable with the card. SC2 is the recorded piece.",
    "SC3 is the follow-up, which is what the exit ticket collects.",
  ],
  tag: "[LI and SC | Planning | HITS 1]",
});

const NOTES_SIGNS = T.composeGlanceNotes({
  answer: "PRACTISE, PROUD, HISTORY, COMMUNITY",
  beats: [
    ["SAY: Same four moves. Set it up, ask it, watch it, check it.",
      "Two of today's signs turn an interview into a person."],
    ["SAY: Ask somebody how much they practise and what they are proud of.",
      "You get the answer nobody else got."],
    "MODEL each sign six times, then inside a question, from the school reference.",
  ],
  trap: ["asking the proud question as an afterthought at the end.",
    "Fix: run the interview backwards, proud question first, twice."],
  prep: [
    "PRACTISE uses PRACTISE_3: the default Signbank entry is LEARNER, not practise.",
    "HISTORY and COMMUNITY are tagged CHECK SIGNBANK in the vocabulary bank.",
    "PROUD also carries boast and brag. Confirm the pride sense before teaching.",
  ],
  tag: "[I Do | Explicit teaching | HITS 3, 4]",
});

const NOTES_MODEL = T.composeGlanceNotes({
  beats: [
    ["MODEL a full four-turn interview with a student who agreed in advance.",
      "Four questions, four answers, no stopping."],
    ["SAY: Watch how I do not look away between questions.",
      "My next question is already ready, so my eyes never leave the answer."],
    ["GET IT WRONG: miss an answer, repair with again, miss another, use what mean.",
      "SAY: Two breakdowns and neither one ended the interview."],
    "SAY: A repair is a pass, not a fail. That is what I want on the recording.",
  ],
  trap: ["restarting the whole interview after a breakdown.",
    "Fix: repair the one turn, then carry on from where it stopped."],
  prep: [
    "Arrange the student partner before the lesson so the model runs clean.",
    "The recording shows where eyes went, which is the easiest thing to fix.",
  ],
  tag: "[I Do | Explicit teaching | HITS 4]",
});

const NOTES_GAME = T.composeGlanceNotes({
  answer: "four questions and four answers inside two minutes",
  beats: [
    ["SET UP: pairs facing, prompt card flat between them, timer visible.",
      "Two minutes one way, then swap."],
    "SAY: Nobody writes anything. The whole two minutes is signing and watching.",
    ["CIRCULATE: watch for the fourth question dropping off.",
      "It is the one attention runs out on."],
    "COLLECT: run it twice with different partners. The second is the rehearsal.",
  ],
  trap: ["the fourth question disappearing under time pressure.",
    "Fix: run it backwards once, fourth question first, then forwards."],
  prep: [
    "Game 7 Two Minute Interview. The clock keeps running through a repair.",
    "That is deliberate: it teaches that a breakdown does not end the exchange.",
  ],
  tag: "[We Do | Collaborative learning | HITS 5]",
});

const NOTES_CHECK = T.composeGlanceNotes({
  answer: "PROUD",
  beats: [
    ["ASK: Show me the question about what a person is proud of.",
      "5 sec think. Cue: Everyone signs it to me on three... one, two, three.",
      "EXPECT: PROUD"],
    ["SCAN the room, back row first.",
      "80%+ -> record straight away.",
      "Less -> run the interview backwards twice, proud question first, then re-ask."],
    ["FOLLOW UP one student.",
      "SAY: That is a strong fourth question. Ask me a harder one on the same idea."],
  ],
  trap: ["the fourth question dropping because the first three used the attention.",
    "Fix: reverse the order so the last question stops being last."],
  prep: [
    "This is the check that feeds the fork decision this afternoon.",
    "Question: could most of them run four turns without the card? Unsure means no.",
  ],
  tag: "[CFU | Evaluating impact | HITS 7]",
});

const NOTES_RECORD = T.composeGlanceNotes({
  beats: [
    ["SET UP: pairs side on to the camera, both faces lit.",
      "Nobody with a window behind them. The camera never sits between two signers."],
    ["SAY: One continuous take. Four questions, four answers, one follow-up, then swap.",
      "No retakes for tidiness."],
    "COLLECT: the camera student uploads before the end of the lesson.",
  ],
  care: "check filming consent against your list before the lesson starts.",
  prep: [
    "Assessment piece 2. One retake only if the take missed part of the exchange.",
    "A student without consent does the same interview live with you, same criteria.",
    "Arranged in advance so it is never visible as an exception on the day.",
  ],
  tag: "[You Do | Evaluating impact | HITS 8]",
});

const NOTES_EXIT = T.composeGlanceNotes({
  answer: "any question that was not one of the four on the card",
  beats: [
    "COLLECT at the door, signed to you, the follow-up they asked.",
    "Two piles: asked one of their own, or used a card question.",
    "Then the Unit Progress Card, Lesson 7 row, three ticks.",
  ],
  prep: [
    "Assesses SC3. This pile is also useful evidence for the fork.",
    "A room that mostly reused card questions is a room still on the card.",
  ],
  tag: "[Exit ticket | Evaluating impact | HITS 7]",
});

const NOTES_CLOSING = T.composeGlanceNotes({
  beats: [
    "SAY: Turn and tell your partner the best answer you got today.",
    "SAY: Next week the stories start, and you will be the one telling them.",
  ],
  prep: [
    "FORK CHECK this afternoon, at your desk, five minutes, nothing to mark.",
    "Could most of them run four turns without the card? Unsure is a no.",
    "Print nothing for Lesson 8 until you have decided.",
  ],
  tag: "[Closing | Planning | HITS 9]",
});

/* ---------------------------------------------------------------- deck ----- */

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  titleSlide(pres, UNIT_TITLE, "Session 7: The two minute interview",
    `Years 5-6 Auslan | Session ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  REPORT.add(signRecallSlide(T, pres, {
    title: "Do Now: which question is each one?",
    lead: "Sign the meaning back to me. Answers come up on the click.",
    signs: [
      { gloss: "WHO", meaning: "who" },
      { gloss: "WHAT", meaning: "what" },
      { gloss: "WHERE", meaning: "where" },
      { gloss: "WHEN", meaning: "when" },
    ],
    notes: NOTES_DONOW, footer: FOOTER,
  }).results);

  liSlide(pres,
    "I am learning to run a four-turn interview and record it.",
    [
      "I can ask two questions and watch both answers.",
      "I can run a four-turn interview with a partner, voice off.",
      "I can add one follow-up question that was not on the card.",
    ],
    NOTES_LI, FOOTER);

  // I Do - the four signs
  REPORT.add(signSlide(T, pres, {
    title: "Signs that turn an interview into a person",
    lead: "Set it up, ask it, watch it, check it. Watch me first, then copy.",
    signs: [
      { gloss: "PRACTISE", meaning: "practise" },
      { gloss: "PROUD", meaning: "proud" },
      { gloss: "HISTORY", meaning: "history" },
      { gloss: "COMMUNITY", meaning: "community" },
    ],
    notes: NOTES_SIGNS, footer: FOOTER,
  }).results);

  // I Do - the four turns, built on clicks
  const turnSlide = pres.addSlide();
  turnSlide.background = { color: C.BG_LIGHT };
  T.addTopBar(turnSlide, C.PRIMARY);
  T.addTitle(turnSlide, "Four turns, then one of your own");
  const TURNS = [
    "Who are you?",
    "What sport do you play?",
    "When did you start?",
    "What are you proud of?",
  ];
  clickBuild(turnSlide, [
    ...TURNS.map((q, i) => () => {
      addTextOnShape(turnSlide, `${i + 1}.  ${q}`, {
        x: 1.5, y: 1.62 + i * 0.72, w: 7.0, h: 0.6, rectRadius: 0.1,
        fill: { color: C.PRIMARY },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
    }),
    () => {
      addTextOnShape(turnSlide, "5.  Something you actually want to know", {
        x: 1.5, y: 4.5, w: 7.0, h: 0.6, rectRadius: 0.1,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
    },
  ]);
  T.addFooter(turnSlide, FOOTER);
  turnSlide.addNotes(NOTES_MODEL);

  // We Do - Game 7
  const gameSlide = pres.addSlide();
  gameSlide.background = { color: C.BG_LIGHT };
  T.addTopBar(gameSlide, C.SUCCESS);
  T.addTitle(gameSlide, "Game 7 Two Minute Interview");
  stepCard(T, gameSlide, {
    title: "Two minutes each way",
    steps: [
      "Card flat on the desk between you",
      "All four questions in two minutes",
      "Timer goes, swap roles, go again",
      "Nobody writes anything at all",
    ],
    x: 0.5, y: 1.5, w: 5.4, h: 2.9, strip: C.SUCCESS,
  });
  addTextOnShape(gameSlide, "2:00", {
    x: 6.1, y: 1.7, w: 3.4, h: 1.0, rectRadius: 0.12, fill: { color: C.PRIMARY },
  }, { fontSize: 40, fontFace: FONT_H, color: C.WHITE, bold: true });
  addTextOnShape(gameSlide, "A repair is a pass", {
    x: 6.1, y: 2.92, w: 3.4, h: 0.7, rectRadius: 0.1, fill: { color: C.ALERT },
  }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });
  gameSlide.addText("The clock keeps running", {
    x: 6.1, y: 3.75, w: 3.4, h: 0.4,
    fontSize: 15, fontFace: FONT_B, color: C.CHARCOAL, align: "center", margin: 0,
  });
  T.addFooter(gameSlide, FOOTER);
  gameSlide.addNotes(NOTES_GAME);
  runSlideDiagnostics(gameSlide, pres);

  // Check it
  const checkSlide = pres.addSlide();
  checkSlide.background = { color: C.BG_LIGHT };
  T.addTopBar(checkSlide, C.ALERT);
  T.addTitle(checkSlide, "Check it");
  addTextOnShape(checkSlide, "Everyone signs it to me on three", {
    x: 1.1, y: 1.9, w: 7.8, h: 0.85, rectRadius: 0.12, fill: { color: C.ALERT },
  }, { fontSize: 26, fontFace: FONT_H, color: C.WHITE, bold: true });
  checkSlide.addText("Show me the question about what you are proud of", {
    x: 1.1, y: 2.95, w: 7.8, h: 0.85,
    fontSize: 24, fontFace: FONT_H, color: C.PRIMARY, bold: true, align: "center", margin: 0,
  });
  checkSlide.addText("Think first. Hands ready. One, two, three.", {
    x: 1.1, y: 3.95, w: 7.8, h: 0.5,
    fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL, align: "center", margin: 0,
  });
  T.addFooter(checkSlide, FOOTER);
  checkSlide.addNotes(NOTES_CHECK);
  runSlideDiagnostics(checkSlide, pres);

  // You Do - the recording
  const recSlide = pres.addSlide();
  recSlide.background = { color: C.BG_LIGHT };
  T.addTopBar(recSlide, C.SUCCESS);
  T.addTitle(recSlide, "You Do: record your interview");
  stepCard(T, recSlide, {
    title: "One continuous take",
    steps: [
      "Sit side on to the camera, faces lit",
      "Four questions, four answers, then your own",
      "Swap roles and go again",
      "Upload before the end of the lesson",
    ],
    x: 0.5, y: 1.5, w: 5.4, h: 2.9, strip: C.SUCCESS,
  });
  addTextOnShape(recSlide, "Voices off", {
    x: 6.1, y: 1.6, w: 3.4, h: 0.85, rectRadius: 0.12, fill: { color: C.ALERT },
  }, { fontSize: 23, fontFace: FONT_H, color: C.WHITE, bold: true });
  addTextOnShape(recSlide, "No retakes for tidiness", {
    x: 6.1, y: 2.65, w: 3.4, h: 0.8, rectRadius: 0.1, fill: { color: C.SECONDARY },
  }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
  recSlide.addText("A repair belongs on the recording", {
    x: 6.1, y: 3.6, w: 3.4, h: 0.6,
    fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, align: "center", margin: 0,
  });
  T.addFooter(recSlide, FOOTER);
  recSlide.addNotes(NOTES_RECORD);
  runSlideDiagnostics(recSlide, pres);

  exitTicketSlide(pres,
    ["Sign me the one follow-up question you asked that was not on the card."],
    NOTES_EXIT, FOOTER, { assessesSc: 3 });

  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner the best answer you got today.",
      scItems: [
        "I can ask two questions and watch both answers.",
        "I can run a four-turn interview with a partner, voice off.",
        "I can add one follow-up question that was not on the card.",
      ],
      selfAssessment: "Unit Progress Card, Lesson 7 row",
    },
    NOTES_CLOSING, FOOTER);

  const pptxPath = path.join(OUT_DIR, `Deaf Sports Session ${SESSION} The Two Minute Interview.pptx`);
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  /* ------------------------------------------------------------- pdf ------ */

  const doc = createPdf({ title: PROMPT_RES.name });
  let y = addPdfHeader(doc, PROMPT_RES.name, {
    subtitle: "Lay this flat on the desk between you. Do not hold it.",
    color: C.PRIMARY,
    lessonInfo: `Session ${SESSION} | Game 7 Two Minute Interview | one per pair`,
  });

  y = addSectionHeading(doc, "Ask these four, in this order", y, { color: C.PRIMARY });

  const QUESTIONS = [
    "Who are you?",
    "What sport do you play?",
    "When did you start?",
    "What are you proud of?",
  ];
  QUESTIONS.forEach((q, i) => {
    doc.save();
    doc.rect(50, y, 495, 46).lineWidth(1).strokeColor("#" + C.SECONDARY).stroke();
    doc.font("Sans-Bold").fontSize(20).fillColor("#" + C.PRIMARY)
      .text(String(i + 1), 64, y + 13, { width: 30 });
    doc.font("Sans").fontSize(17).fillColor("#" + C.CHARCOAL)
      .text(q, 104, y + 15, { width: 420 });
    doc.restore();
    y += 54;
  });

  y += 8;
  y = addSectionHeading(doc, "Then ask one of your own", y, { color: C.SECONDARY });
  y = addBodyText(doc,
    "Something you actually want to know, that is not on this list.", y);

  y += 6;
  y = addSectionHeading(doc, "If you miss an answer", y, { color: C.SECONDARY });
  y = addBodyText(doc,
    "Ask for it again. If a repeat would not help, ask what they mean instead.", y);

  addTipBox(doc,
    "The teacher will model the Auslan for each of these. Read the English here; sign what you were shown.",
    y, { color: C.ACCENT });

  addPdfFooter(doc, `Session ${SESSION} | Interview Prompt Card | Years 5-6 Auslan`);
  await writePdf(doc, path.join(OUT_DIR, PROMPT_RES.fileName));
  console.log("PDF written: " + PROMPT_RES.fileName);

  REPORT.print(`Session ${SESSION}`);
  console.log(`Session ${SESSION} build complete.`);
}

build().catch((err) => { console.error(err); process.exit(1); });
