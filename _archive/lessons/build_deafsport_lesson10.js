"use strict";

// Deaf Sports in Australia (Years 5-6 Auslan, Term 4) - Session 10 of 10.
// SHOWCASE AND REFLECTION: present a retell, give one useful piece of feedback.
//   No new signs. All 35 retrieved. Game 4 reused with the full ten-card
//   timeline, and both story strips reused from Sessions 8 and 9.
// This is the compressible lesson named in the unit header. If a lesson is lost
// to an assembly or a public holiday, this is the one to drop.
// Branch B note: this lesson does not happen if the fork went to consolidation.

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant } = require("../themes/factory");
const {
  signRecallSlide, stepCard, noPrintResourceSlide, createSignReport, ATTRIBUTION,
} = require("./auslan_signs_lib");

const T = createTheme("literacy", "grade56", weekToVariant(1));
const {
  C, FONT_H, FONT_B,
  titleSlide, liSlide, closingSlide, exitTicketSlide,
  addTextOnShape, clickBuild, runSlideDiagnostics,
} = T;

const SESSION = 10;
const TOTAL = 10;
const UNIT_TITLE = "Deaf Sports in Australia";
const ANCHOR = "Set it up, ask it, watch it, check it.";
const FOOTER = `${UNIT_TITLE} | Session ${SESSION} of ${TOTAL} | Years 5-6 Auslan`;
const OUT_DIR = "output/DeafSport_Lesson10_Showcase_And_Reflection";

fs.mkdirSync(OUT_DIR, { recursive: true });

const REPORT = createSignReport();

/* ---------------------------------------------------------------- notes ---- */

const NOTES_TITLE = "Session 10 of 10. Nothing new to teach and nothing new to print. This is the compressible lesson if a week is lost.";

const NOTES_RESOURCES = T.composeGlanceNotes({
  beats: [
    ["SHOW while students settle.", "SAY: Everything today has been out of this tub before."],
    "SET UP: timeline cards from Session 4, both story strips, progress cards.",
  ],
  prep: [
    "Nothing new to print. Session 4 timeline cards, now all ten in play.",
    "Collect the Unit Progress Cards at the end of this lesson for good.",
    ATTRIBUTION,
  ],
  tag: "[Setup | Planning | HITS 2]",
});

const NOTES_DONOW = T.composeGlanceNotes({
  answer: "soon, later, next",
  beats: [
    ["ASK: What does each sign mean?",
      "10 sec think. Cue: Everyone signs it to me on three... one, two, three.",
      "EXPECT: SOON, LATER, NEXT"],
    "REVEAL after the room has answered, then run the four question signs too.",
    ["SAY: Everything from ten weeks is in your hands at once.",
      "Today somebody watches you use it and tells you one thing."],
  ],
  prep: [
    "The widest retrieval in the unit: four question signs and four time signs.",
    "Ten fast rounds. It is meant to feel like everything at once.",
  ],
  tag: "[Do Now | Attention, focus and regulation | HITS 6]",
});

const NOTES_LI = T.composeGlanceNotes({
  beats: [
    ["POINT to each criterion.", "SAY: You present, you get one useful thing, you run it again."],
    "SAY: By the end everyone has presented to one other pair.",
  ],
  prep: [
    "SC1 is presenting, reachable by every student who has a story.",
    "SC2 is what the exit ticket collects. SC3 is acting on the feedback.",
  ],
  tag: "[LI and SC | Planning | HITS 1]",
});

const NOTES_FEEDBACK = T.composeGlanceNotes({
  answer: "feedback that names what the signer did",
  beats: [
    ["SAY: Same four moves. Set it up, ask it, watch it, check it.",
      "A presentation sets up twice: for the story, and the audience."],
    ["MODEL a retell, then model feedback on yourself.",
      "SAY: You kept the two people apart, so I knew who was who."],
    ["SAY: That names what I did. Good job names nothing.",
      "GET IT WRONG: give yourself empty praise.",
      "SAY: Well done, that was lovely. That helps nobody."],
  ],
  trap: ["feedback that praises rather than names.",
    "Fix: student gives it again starting with you kept."],
  stretch: "present to a different table, then teach one move.",
  help: "the shorter story, plus the starter you kept... so I knew...",
  prep: [
    "Modelling feedback on yourself is safer than modelling it on a student.",
    "Year 5 uses the sentence starter. Year 6 gives feedback unscaffolded.",
  ],
  tag: "[I Do | Explicit teaching | HITS 3, 8]",
});

const NOTES_GAME = T.composeGlanceNotes({
  answer: "all ten events laid out oldest to newest",
  beats: [
    ["SET UP: groups of four, all ten cards face down in the middle.",
      "Two or three cards each, flat on the desk."],
    "SAY: Ten cards this time, so you have to hold a running order, not compare two.",
    ["CIRCULATE: this is the last retrieval of the history content.",
      "Note who still cannot place the 1880s card."],
    "COLLECT: read the whole line back around the circle.",
  ],
  trap: ["placing late cards by guessing once the line gets long.",
    "Fix: cover the line, ask for the year, student re-places it."],
  prep: [
    "Game 4 reused, now with all ten cards including 2011 and 2017.",
    "Nothing new to print. The Session 4 set already has them.",
  ],
  tag: "[We Do | Collaborative learning | HITS 6]",
});

const NOTES_CHECK = T.composeGlanceNotes({
  answer: "any real event, with a reason attached to it",
  beats: [
    ["ASK: Which event mattered most to Deaf people in Australia?",
      "10 sec think. Cue: Everyone signs it to me on three... one, two, three.",
      "EXPECT: a real event, not a year on its own"],
    ["SCAN the room, back row first.",
      "80%+ give an event with a reason ready -> go to You Do.",
      "Less -> pairs take one event and finish, this changed things because... re-ask."],
    ["FOLLOW UP one student.",
      "SAY: Do you agree with that one? Add what it changed."],
  ],
  trap: ["naming an event with no impact attached to it.",
    "Fix: finish the sentence frame in pairs, then re-ask."],
  prep: [
    "Ten seconds of think time here, not five. It is a judgement, not a recall.",
    "This is the last evidence for the recount statement in the coverage check.",
  ],
  tag: "[CFU | Evaluating impact | HITS 7]",
});

const NOTES_YOUDO = T.composeGlanceNotes({
  beats: [
    ["SAY: Voices off. Pair A presents the whole retell to pair B.",
      "Pair B gives one piece of feedback that names something."],
    "SAY: Pair A runs it again with one change. Then swap over.",
    ["CIRCULATE: the second run is the assessment, not the first.",
      "Watch for the change actually landing."],
  ],
  stretch: "present to a different table, then teach one move to a younger buddy.",
  help: "the shorter story and the feedback sentence starter on a strip.",
  prep: [
    "Different content from the game: ordering events there, a whole retell here.",
    "SC3 only shows up on the second run, so protect the time for it.",
  ],
  tag: "[You Do | Explicit teaching | HITS 5, 8]",
});

const NOTES_EXIT = T.composeGlanceNotes({
  answer: "one specific change made between the first and second run",
  beats: [
    "COLLECT at the door, signed to you, the change they made.",
    "Two piles: named a specific change, or ran it the same way twice.",
    "Then the Unit Progress Card, final row, and collect the cards for good.",
  ],
  prep: [
    "Assesses SC3. A student who ran it identically did not use the feedback.",
    "That is worth a line in your notes for next term, not a reteach today.",
  ],
  tag: "[Exit ticket | Evaluating impact | HITS 7]",
});

const NOTES_CLOSING = T.composeGlanceNotes({
  beats: [
    "SAY: Ten weeks ago you could not ask anybody anything without your voice.",
    "SAY: Turn and tell your partner the one thing you can do now that you could not.",
  ],
  prep: [
    "Protocol today: all of them, student-led. Nominate two per group and stay out.",
    "Collect the progress cards. They are the only record students kept.",
  ],
  tag: "[Closing | Planning | HITS 9]",
});

/* ---------------------------------------------------------------- deck ----- */

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  titleSlide(pres, UNIT_TITLE, "Session 10: Showcase and reflection",
    `Years 5-6 Auslan | Session ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  noPrintResourceSlide(T, pres, {
    items: [
      "Timeline cards from Session 4, all ten this time",
      "Story Strip One and Story Strip Two, from Sessions 8 and 9",
      "Unit Progress Cards, collected for good at the end",
    ],
    notes: NOTES_RESOURCES, footer: FOOTER,
  });

  REPORT.add(signRecallSlide(T, pres, {
    title: "Do Now: what does each one mean?",
    lead: "Sign the meaning back to me. Answers come up on the click.",
    signs: [
      { gloss: "SOON", meaning: "soon" },
      { gloss: "LATER", meaning: "later" },
      { gloss: "NEXT", meaning: "next" },
    ],
    notes: NOTES_DONOW, footer: FOOTER,
  }).results);

  liSlide(pres,
    "I am learning to present a retell to an audience and give one piece of useful feedback.",
    [
      "I can present my retell to one other pair.",
      "I can give feedback that names one thing the signer did.",
      "I can take feedback and run my retell again with one change.",
    ],
    NOTES_LI, FOOTER);

  // I Do - what useful feedback sounds like
  const fbSlide = pres.addSlide();
  fbSlide.background = { color: C.BG_LIGHT };
  T.addTopBar(fbSlide, C.PRIMARY);
  T.addTitle(fbSlide, "What useful feedback sounds like");
  clickBuild(fbSlide, [
    () => {
      addTextOnShape(fbSlide, "You kept the two people apart, so I always knew who was who.", {
        x: 0.7, y: 1.75, w: 8.6, h: 0.85, rectRadius: 0.1, fill: { color: C.SUCCESS },
      }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });
      fbSlide.addText("That names what you did.", {
        x: 0.7, y: 2.7, w: 8.6, h: 0.45,
        fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL, align: "center", margin: 0,
      });
    },
    () => {
      addTextOnShape(fbSlide, "Well done. That was lovely.", {
        x: 0.7, y: 3.32, w: 8.6, h: 0.8, rectRadius: 0.1, fill: { color: C.ALERT },
      }, { fontSize: 19, fontFace: FONT_H, color: C.WHITE, bold: true });
      fbSlide.addText("Names nothing. Helps nobody.", {
        x: 0.7, y: 4.22, w: 8.6, h: 0.45,
        fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL, align: "center", margin: 0,
      });
    },
  ]);
  T.addFooter(fbSlide, FOOTER);
  fbSlide.addNotes(NOTES_FEEDBACK);

  // We Do - Game 4 reused with all ten cards
  const gameSlide = pres.addSlide();
  gameSlide.background = { color: C.BG_LIGHT };
  T.addTopBar(gameSlide, C.SUCCESS);
  T.addTitle(gameSlide, "Game 4 again, all ten cards");
  stepCard(T, gameSlide, {
    title: "The whole timeline this time",
    steps: [
      "Two or three cards each, flat on the desk",
      "Sign your event and its year",
      "The group decides where it goes",
      "Read the finished line back around the circle",
    ],
    x: 0.5, y: 1.5, w: 5.4, h: 2.9, strip: C.SUCCESS,
  });
  addTextOnShape(gameSlide, "10 cards", {
    x: 6.1, y: 1.7, w: 3.4, h: 0.9, rectRadius: 0.12, fill: { color: C.PRIMARY },
  }, { fontSize: 30, fontFace: FONT_H, color: C.WHITE, bold: true });
  gameSlide.addText("Long enough that you have to hold a running order, not just compare two.", {
    x: 6.1, y: 2.8, w: 3.4, h: 1.2,
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
  addTextOnShape(checkSlide, "Everyone signs it to me on three", {
    x: 1.1, y: 1.85, w: 7.8, h: 0.85, rectRadius: 0.12, fill: { color: C.ALERT },
  }, { fontSize: 26, fontFace: FONT_H, color: C.WHITE, bold: true });
  checkSlide.addText("Which event mattered most to Deaf people in Australia?", {
    x: 1.1, y: 2.9, w: 7.8, h: 0.9,
    fontSize: 24, fontFace: FONT_H, color: C.PRIMARY, bold: true, align: "center", margin: 0,
  });
  checkSlide.addText("Ten seconds. This one is a judgement, not a memory.", {
    x: 1.1, y: 3.95, w: 7.8, h: 0.5,
    fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL, align: "center", margin: 0,
  });
  T.addFooter(checkSlide, FOOTER);
  checkSlide.addNotes(NOTES_CHECK);
  runSlideDiagnostics(checkSlide, pres);

  // You Do - present, feedback, run again
  const youSlide = pres.addSlide();
  youSlide.background = { color: C.BG_LIGHT };
  T.addTopBar(youSlide, C.SUCCESS);
  T.addTitle(youSlide, "You Do: present, then run it again");
  stepCard(T, youSlide, {
    title: "Pair to pair",
    steps: [
      "Pair A presents the whole retell",
      "Pair B gives one piece of feedback",
      "Pair A runs it again with one change",
      "Then swap over",
    ],
    x: 0.5, y: 1.5, w: 5.4, h: 2.9, strip: C.SUCCESS,
  });
  addTextOnShape(youSlide, "Voices off", {
    x: 6.1, y: 1.6, w: 3.4, h: 0.85, rectRadius: 0.12, fill: { color: C.ALERT },
  }, { fontSize: 23, fontFace: FONT_H, color: C.WHITE, bold: true });
  addTextOnShape(youSlide, "You kept... so I knew...", {
    x: 6.1, y: 2.65, w: 3.4, h: 0.8, rectRadius: 0.1, fill: { color: C.SECONDARY },
  }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
  youSlide.addText("Name what they did", {
    x: 6.1, y: 3.6, w: 3.4, h: 0.4,
    fontSize: 15, fontFace: FONT_B, color: C.CHARCOAL, align: "center", margin: 0,
  });
  T.addFooter(youSlide, FOOTER);
  youSlide.addNotes(NOTES_YOUDO);
  runSlideDiagnostics(youSlide, pres);

  exitTicketSlide(pres,
    ["Sign me the one change you made after your feedback."],
    NOTES_EXIT, FOOTER, { assessesSc: 3 });

  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner the one thing you can do now that you could not ten weeks ago.",
      scItems: [
        "I can present my retell to one other pair.",
        "I can give feedback that names one thing the signer did.",
        "I can take feedback and run my retell again with one change.",
      ],
      selfAssessment: "Unit Progress Card, final row, then hand it in",
    },
    NOTES_CLOSING, FOOTER);

  const pptxPath = path.join(OUT_DIR, `Deaf Sports Session ${SESSION} Showcase And Reflection.pptx`);
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  REPORT.print(`Session ${SESSION}`);
  console.log(`Session ${SESSION} build complete.`);
}

build().catch((err) => { console.error(err); process.exit(1); });
