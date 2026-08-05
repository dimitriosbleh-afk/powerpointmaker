"use strict";

// Deaf Sports in Australia (Years 5-6 Auslan, Term 4) - Session 9 of 10.
// TRANSLATING, NOT SWAPPING: a written story into Auslan.
//   4 production signs: SOON, LATER, NEXT, FINISH. Game 9 Same Story Two Ways.
//   NEXT uses NEXT_2.jpg: the default NEXT.jpg is DEMOTION, not the next event.
// Which reordering to show is confirmed from the school reference before the
// lesson. This deck never states an Auslan word order of its own.
// Branch B note: this lesson is dropped if the fork went to consolidation.

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading, addBodyText,
  addTipBox, addPdfFooter, addWriteLine, addResourceSlide,
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

const SESSION = 9;
const TOTAL = 10;
const UNIT_TITLE = "Deaf Sports in Australia";
const ANCHOR = "Set it up, ask it, watch it, check it.";
const FOOTER = `${UNIT_TITLE} | Session ${SESSION} of ${TOTAL} | Years 5-6 Auslan`;
const OUT_DIR = "output/DeafSport_Lesson9_Translating_Not_Swapping";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const STORY_RES = makeSessionResource(SESSION,
  "Story Strip Two The Race That Started With A Light",
  "One per pair. Cut in half for Game 9 so each pair holds two events.");
const RESOURCE_ITEMS = [STORY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

const REPORT = createSignReport();

const STORY_TWO = [
  "A swimmer is on the blocks. Beside her, in the next lane, is the swimmer she has raced all season.",
  "The starter raises the gun. She is not listening for it. She is looking down at the side of the pool.",
  "The gun fires and the light comes on. Both swimmers are gone before the sound reaches the back row.",
  "Later, at the wall, she turns and looks up at the board. Her name is at the top.",
];

/* ---------------------------------------------------------------- notes ---- */

const NOTES_TITLE = "Session 9 of 10. Translation, not word-for-word signing. Decide which single reordering you will show before the lesson.";

const NOTES_RESOURCES = T.composeGlanceNotes({
  beats: [
    ["SHOW while students settle.", "SAY: One story per pair, and I will cut it in half for the game."],
    "SET UP: pairs at an angle so both signing spaces are clear.",
  ],
  prep: [
    "Print the story strip per pair. Cut half the copies in two for Game 9.",
    "Six support copies have the four events already underlined.",
    ATTRIBUTION,
  ],
  tag: "[Setup | Planning | HITS 2]",
});

const NOTES_DONOW = T.composeGlanceNotes({
  answer: "story, tell, happen",
  beats: [
    ["ASK: What does each sign mean?",
      "10 sec think. Cue: Everyone signs it to me on three... one, two, three.",
      "EXPECT: STORY, TELL, HAPPEN"],
    "REVEAL after the room has answered, one click.",
    ["SAY: Last week the story came from me. Today it comes off a page.",
      "Paper is written in English, which is not the language you will tell it in."],
  ],
  prep: [
    "Retrieval from Lesson 8, then pairs set last week's scene to each other.",
    "The bridge names the whole problem the lesson solves.",
  ],
  tag: "[Do Now | Attention, focus and regulation | HITS 6]",
});

const NOTES_LI = T.composeGlanceNotes({
  beats: [
    ["POINT to each criterion.", "SAY: Two steps today. Find the events, then translate them."],
    "SAY: By the end everyone can find the four events in a paragraph.",
  ],
  prep: [
    "SC1 is finding the events, reachable by every student.",
    "SC2 is what the exit ticket collects. SC3 explains one reordering.",
  ],
  tag: "[LI and SC | Planning | HITS 1]",
});

const NOTES_SIGNS = T.composeGlanceNotes({
  answer: "SOON, LATER, NEXT, FINISH",
  beats: [
    ["SAY: Same four moves. Set it up, ask it, watch it, check it.",
      "Four more time signs, and these move a story along."],
    ["SAY: Before, after and back then put you in a year.",
      "Soon, later and next move you through an afternoon."],
    "MODEL each sign six times, then joining two events together.",
  ],
  trap: ["using before and after inside a single scene.",
    "Fix: sort six events into a year or an afternoon, then redo."],
  prep: [
    "NEXT uses NEXT_2: the default Signbank entry for next is DEMOTION.",
    "All four time words named in the term overview are now covered.",
  ],
  tag: "[I Do | Explicit teaching | HITS 3, 4]",
});

const NOTES_STEPS = T.composeGlanceNotes({
  answer: "four events; the rest of the paragraph is description",
  beats: [
    ["SAY: Step one, I am not signing anything yet.",
      "I am finding the four things that actually happen.",
      "Everything else in this paragraph is decoration."],
    ["SAY: Step two, now I translate. Not word by word.",
      "The English says one thing first and the Auslan may not."],
    ["SAY: That is not me getting it wrong. That is two languages being different.",
      "Watch for the moment my hands do something the paper did not say."],
  ],
  trap: ["counting sentences instead of events.",
    "Fix: cut the paragraph into strips, sort into happens and describes."],
  prep: [
    "CHECK GRAMMAR: confirm which reordering you will show, and show only that one.",
    "Do not state an Auslan word order this document has not verified.",
  ],
  tag: "[I Do | Explicit teaching | HITS 4]",
});

const NOTES_WRONG = T.composeGlanceNotes({
  beats: [
    ["GET IT WRONG: sign the first sentence word for word off the page.",
      "Include the small English words."],
    ["SAY: That was English on my hands.",
      "Every word from the paper, in the paper's order, and it is not Auslan."],
    "MODEL it properly, then name the one place the order changed.",
  ],
  trap: ["signing every English word including the small ones.",
    "Fix: cover the paper, retell from memory, then compare."],
  stretch: "translate it, then write two sentences on what you changed and why.",
  help: "the paragraph with the four events already underlined.",
  prep: [
    "The deliberate error is the clearest way to show what translation is not.",
    "Year 5 translates two events underlined. Year 6 translates four unmarked.",
  ],
  tag: "[I Do | Explicit teaching | HITS 3]",
});

const NOTES_GAME = T.composeGlanceNotes({
  answer: "the four events joined into one story across two pairs",
  beats: [
    ["SET UP: groups of four, two pairs. Same story, different halves.",
      "Pair A translates and signs their two events."],
    "SAY: Pair B watches all of it, pencils down, then signs the events that come next.",
    ["CIRCULATE: a pair starting before the other pair is watching has told nobody.",
      "Every turn opens by waiting for two sets of eyes."],
    "COLLECT: the group decides which events could have come in a different order.",
  ],
  trap: ["joining the halves without a time sign, so the seam is invisible.",
    "Fix: pair B signs the join again using soon, later or next."],
  prep: [
    "Game 9 Same Story Two Ways. Swap which pair goes first and run it again.",
    "For a harder round, do not tell them which half is first.",
  ],
  tag: "[We Do | Collaborative learning | HITS 5]",
});

const NOTES_CHECK = T.composeGlanceNotes({
  answer: "four",
  beats: [
    ["ASK: How many events are in the paragraph?",
      "5 sec think. Cue: Write it. Boards up on three... one, two, three.",
      "EXPECT: 4"],
    ["SCAN the boards, back row first.",
      "80%+ -> go to You Do with the full translation.",
      "Less -> cut the paragraph into strips, sort happens and describes, re-ask."],
    ["FOLLOW UP one student.",
      "SAY: Do you agree with four? Add which sentence you left out and why."],
  ],
  trap: ["counting sentences, which is a reading problem in an Auslan costume.",
    "Fix: sort the strips, then re-ask with the same routine."],
  prep: [
    "The decision point that changes Lesson 10.",
    "If finding events fails, the presentation next week has nothing to present.",
  ],
  tag: "[CFU | Evaluating impact | HITS 7]",
});

const NOTES_YOUDO = T.composeGlanceNotes({
  beats: [
    ["SAY: Voices off. Translate the whole paragraph to your partner.",
      "Scene first, then four events joined with the new time signs."],
    "SAY: Your partner watches it all, then signs back one event.",
    ["CIRCULATE: listen for the join, not the events.",
      "The time signs are what makes it one story."],
  ],
  stretch: "write two English sentences saying what you changed and why.",
  help: "the paragraph with the four events underlined, so step one is done.",
  prep: [
    "Different content from the game: half the story there, all of it here.",
    "SC3 only shows up here, so protect this stage if time is short.",
  ],
  tag: "[You Do | Explicit teaching | HITS 10]",
});

const NOTES_EXIT = T.composeGlanceNotes({
  answer: "any real place the Auslan order differed from the English",
  beats: [
    "COLLECT at the door, signed to you, one place the order changed.",
    "Two piles: named a real one, or could not find one.",
    "Then the Unit Progress Card, Lesson 9 row, three ticks.",
  ],
  prep: [
    "Assesses SC3. A student who found none has probably signed word for word.",
    "That is the thing to pick up in the Lesson 10 presentations.",
  ],
  tag: "[Exit ticket | Evaluating impact | HITS 7]",
});

const NOTES_CLOSING = T.composeGlanceNotes({
  beats: [
    "SAY: Turn and tell your partner one thing the paper said that your hands did not.",
    "SAY: Next week you present, and somebody tells you one useful thing.",
  ],
  prep: [
    "Protocol practised today: waiting for eyes before you start.",
    "Running late: cut the game to one round. Never cut the You Do.",
  ],
  tag: "[Closing | Planning | HITS 9]",
});

/* ---------------------------------------------------------------- deck ----- */

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  titleSlide(pres, UNIT_TITLE, "Session 9: Translating, not swapping",
    `Years 5-6 Auslan | Session ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  REPORT.add(signRecallSlide(T, pres, {
    title: "Do Now: what does each one mean?",
    lead: "Sign the meaning back to me. Answers come up on the click.",
    signs: [
      { gloss: "STORY", meaning: "story" },
      { gloss: "TELL", meaning: "tell" },
      { gloss: "HAPPEN", meaning: "happen" },
    ],
    notes: NOTES_DONOW, footer: FOOTER,
  }).results);

  liSlide(pres,
    "I am learning to translate a written story into Auslan rather than sign it word by word.",
    [
      "I can find the four events in a written story.",
      "I can translate the story into Auslan using soon, later and next.",
      "I can say one place where the Auslan does not match the English word order and why.",
    ],
    NOTES_LI, FOOTER);

  // I Do - the four signs
  REPORT.add(signSlide(T, pres, {
    title: "Signs that move a story along",
    lead: "Set it up, ask it, watch it, check it. Watch me first, then copy.",
    signs: [
      { gloss: "SOON", meaning: "soon" },
      { gloss: "LATER", meaning: "later" },
      { gloss: "NEXT", meaning: "next" },
      { gloss: "FINISH", meaning: "finish" },
    ],
    notes: NOTES_SIGNS, footer: FOOTER,
  }).results);

  // I Do - the two steps, on clicks
  const stepSlide = pres.addSlide();
  stepSlide.background = { color: C.BG_LIGHT };
  T.addTopBar(stepSlide, C.PRIMARY);
  T.addTitle(stepSlide, "Two steps, and step one is not signing");
  clickBuild(stepSlide, [
    () => {
      addTextOnShape(stepSlide, "Step 1: find the four things that happen", {
        x: 0.7, y: 1.7, w: 8.6, h: 0.78, rectRadius: 0.1, fill: { color: C.PRIMARY },
      }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });
      stepSlide.addText("Everything else in the paragraph is decoration.", {
        x: 0.7, y: 2.58, w: 8.6, h: 0.45,
        fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL, align: "center", margin: 0,
      });
    },
    () => {
      addTextOnShape(stepSlide, "Step 2: translate. Not word by word", {
        x: 0.7, y: 3.2, w: 8.6, h: 0.78, rectRadius: 0.1, fill: { color: C.SUCCESS },
      }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });
      stepSlide.addText("Two languages being different is not you getting it wrong.", {
        x: 0.7, y: 4.08, w: 8.6, h: 0.45,
        fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL, align: "center", margin: 0,
      });
    },
  ]);
  T.addFooter(stepSlide, FOOTER);
  stepSlide.addNotes(NOTES_STEPS);

  // I Do - the deliberate error
  const wrongSlide = pres.addSlide();
  wrongSlide.background = { color: C.BG_LIGHT };
  T.addTopBar(wrongSlide, C.ALERT);
  T.addTitle(wrongSlide, "Watch me get it wrong");
  addTextOnShape(wrongSlide, "That was English on my hands", {
    x: 1.3, y: 1.9, w: 7.4, h: 0.9, rectRadius: 0.12, fill: { color: C.ALERT },
  }, { fontSize: 26, fontFace: FONT_H, color: C.WHITE, bold: true });
  wrongSlide.addText("Every word from the paper, in the paper's order.", {
    x: 1.3, y: 2.95, w: 7.4, h: 0.6,
    fontSize: 20, fontFace: FONT_H, color: C.PRIMARY, bold: true, align: "center", margin: 0,
  });
  addTextOnShape(wrongSlide, "Watch for the moment my hands do something the paper did not say", {
    x: 0.9, y: 3.75, w: 8.2, h: 0.72, rectRadius: 0.1, fill: { color: C.SECONDARY },
  }, { fontSize: 15, fontFace: FONT_B, color: C.WHITE, bold: true });
  T.addFooter(wrongSlide, FOOTER);
  wrongSlide.addNotes(NOTES_WRONG);
  runSlideDiagnostics(wrongSlide, pres);

  // We Do - Game 9
  const gameSlide = pres.addSlide();
  gameSlide.background = { color: C.BG_LIGHT };
  T.addTopBar(gameSlide, C.SUCCESS);
  T.addTitle(gameSlide, "Game 9 Same Story Two Ways");
  stepCard(T, gameSlide, {
    title: "Two pairs, one story",
    steps: [
      "Your pair has half the events",
      "Translate them together, then sign them",
      "The other pair signs what comes next",
      "Join the halves with soon, later or next",
    ],
    x: 0.5, y: 1.5, w: 5.4, h: 2.9, strip: C.SUCCESS,
  });
  addTextOnShape(gameSlide, "Wait for eyes", {
    x: 6.1, y: 1.7, w: 3.4, h: 0.85, rectRadius: 0.12, fill: { color: C.ALERT },
  }, { fontSize: 21, fontFace: FONT_H, color: C.WHITE, bold: true });
  gameSlide.addText("A pair who starts before the other pair is watching has told nobody anything.", {
    x: 6.1, y: 2.75, w: 3.4, h: 1.25,
    fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, align: "center", margin: 0,
  });
  T.addFooter(gameSlide, FOOTER);
  gameSlide.addNotes(NOTES_GAME);
  runSlideDiagnostics(gameSlide, pres);

  // Check it
  const checkSlide = pres.addSlide();
  checkSlide.background = { color: C.BG_LIGHT };
  T.addTopBar(checkSlide, C.ALERT);
  T.addTitle(checkSlide, "Check it");
  addTextOnShape(checkSlide, "Write it. Boards up on three", {
    x: 1.1, y: 1.9, w: 7.8, h: 0.85, rectRadius: 0.12, fill: { color: C.ALERT },
  }, { fontSize: 26, fontFace: FONT_H, color: C.WHITE, bold: true });
  checkSlide.addText("How many events are in the paragraph?", {
    x: 1.1, y: 3.0, w: 7.8, h: 0.7,
    fontSize: 26, fontFace: FONT_H, color: C.PRIMARY, bold: true, align: "center", margin: 0,
  });
  checkSlide.addText("Events, not sentences.", {
    x: 1.1, y: 3.85, w: 7.8, h: 0.5,
    fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL, align: "center", margin: 0,
  });
  T.addFooter(checkSlide, FOOTER);
  checkSlide.addNotes(NOTES_CHECK);
  runSlideDiagnostics(checkSlide, pres);

  // You Do
  const youSlide = pres.addSlide();
  youSlide.background = { color: C.BG_LIGHT };
  T.addTopBar(youSlide, C.SUCCESS);
  T.addTitle(youSlide, "You Do: translate the whole story");
  stepCard(T, youSlide, {
    title: "All four events, joined up",
    steps: [
      "Set the scene first",
      "Tell the four events in order",
      "Join them with soon, later or next",
      "Partner signs back one event",
    ],
    x: 0.5, y: 1.5, w: 5.4, h: 2.9, strip: C.SUCCESS,
  });
  addTextOnShape(youSlide, "Voices off", {
    x: 6.1, y: 1.6, w: 3.4, h: 0.9, rectRadius: 0.12, fill: { color: C.ALERT },
  }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
  ["soon", "later", "next", "finish"].forEach((wd, i) => {
    addTextOnShape(youSlide, wd, {
      x: 6.1, y: 2.68 + i * 0.53, w: 3.4, h: 0.44, rectRadius: 0.07,
      fill: { color: C.BG_CARD }, line: { color: C.SECONDARY, width: 1 },
    }, { fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL });
  });
  T.addFooter(youSlide, FOOTER);
  youSlide.addNotes(NOTES_YOUDO);
  runSlideDiagnostics(youSlide, pres);

  exitTicketSlide(pres,
    ["Sign me one place where your Auslan did not match the English word order."],
    NOTES_EXIT, FOOTER, { assessesSc: 3 });

  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner one thing the paper said that your hands did not.",
      scItems: [
        "I can find the four events in a written story.",
        "I can translate the story into Auslan using soon, later and next.",
        "I can say one place where the Auslan does not match the English word order and why.",
      ],
      selfAssessment: "Unit Progress Card, Lesson 9 row",
    },
    NOTES_CLOSING, FOOTER);

  const pptxPath = path.join(OUT_DIR, `Deaf Sports Session ${SESSION} Translating Not Swapping.pptx`);
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  /* ------------------------------------------------------------- pdf ------ */

  const doc = createPdf({ title: STORY_RES.name });
  let y = addPdfHeader(doc, STORY_RES.name, {
    subtitle: "The Race That Started With A Light",
    color: C.PRIMARY,
    lessonInfo: `Session ${SESSION} | one per pair | cut in half for Game 9`,
  });

  y = addTipBox(doc,
    "This story is written for this lesson. The real detail underneath it is that in Deaf swimming a light on the side of the pool is wired to the starter's gun, so the race starts when the light comes on. The source is in the unit document's Sources appendix.",
    y, { color: C.ACCENT });

  y = addSectionHeading(doc, "The Race That Started With A Light", y, { color: C.PRIMARY });
  STORY_TWO.forEach((para, i) => {
    if (i === 2) {
      doc.save();
      doc.moveTo(50, y - 6).lineTo(545, y - 6).lineWidth(0.8)
        .dash(4, { space: 3 }).strokeColor("#" + C.MUTED).stroke();
      doc.undash();
      doc.restore();
      y += 6;
    }
    doc.font("Sans").fontSize(14).fillColor("#" + C.CHARCOAL);
    doc.text(para, 50, y, { width: 495, lineGap: 5 });
    y += doc.heightOfString(para, { width: 495, lineGap: 5 }) + 14;
  });

  y += 4;
  y = addBodyText(doc,
    "Cut along the dashed line for Game 9. The top pair has the first two events, the bottom pair has the last two.",
    y);

  y += 4;
  y = addSectionHeading(doc, "Before you sign it", y, { color: C.SECONDARY });
  y = addWriteLine(doc, "1.  Where does it happen? Show that first.", y);
  y = addWriteLine(doc, "2.  Who is in it? Two swimmers, two different places.", y);
  y = addWriteLine(doc, "3.  Find the four things that actually happen.", y);
  y = addWriteLine(doc, "4.  Join them with soon, later or next.", y);

  addPdfFooter(doc, `Session ${SESSION} | Story Strip Two | Years 5-6 Auslan`);
  await writePdf(doc, path.join(OUT_DIR, STORY_RES.fileName));
  console.log("PDF written: " + STORY_RES.fileName);

  REPORT.print(`Session ${SESSION}`);
  console.log(`Session ${SESSION} build complete.`);
}

build().catch((err) => { console.error(err); process.exit(1); });
