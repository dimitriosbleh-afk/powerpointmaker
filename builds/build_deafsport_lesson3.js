"use strict";

// Deaf Sports in Australia (Years 5-6 Auslan, Term 4) - Session 3 of 10.
// WHEN AND HOW MANY: years and numbers, plus the second repair move.
//   4 production signs: WHEN, HOW-MANY, YEAR, WHAT MEAN?. Game 3 More or Less.
//   MEAN gets a lookup card on purpose: its first Signbank sense is SIGN
//   LANGUAGE, not meaning, so this deck does not assert it.
// Unit anchor: Set it up, ask it, watch it, check it.
// No printed resource this session; boards only.

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant } = require("../themes/factory");
const {
  signSlide, signRecallSlide, signCardRow, stepCard,
  noPrintResourceSlide, createSignReport, ATTRIBUTION,
} = require("./auslan_signs_lib");

const T = createTheme("literacy", "grade56", weekToVariant(1));
const {
  C, FONT_H, FONT_B,
  titleSlide, liSlide, closingSlide, exitTicketSlide,
  addTextOnShape, clickBuild, runSlideDiagnostics,
} = T;

const SESSION = 3;
const TOTAL = 10;
const UNIT_TITLE = "Deaf Sports in Australia";
const ANCHOR = "Set it up, ask it, watch it, check it.";
const FOOTER = `${UNIT_TITLE} | Session ${SESSION} of ${TOTAL} | Years 5-6 Auslan`;
const OUT_DIR = "output/DeafSport_Lesson3_When_And_How_Many";

fs.mkdirSync(OUT_DIR, { recursive: true });

const REPORT = createSignReport();

/* ---------------------------------------------------------------- notes ---- */

const NOTES_TITLE = "Session 3 of 10. Years and numbers. Confirm how a four-digit year is produced before this lesson; do not improvise it.";

const NOTES_RESOURCES = T.composeGlanceNotes({
  beats: [
    ["SHOW while students settle.", "SAY: Boards out. Nothing to hand round today."],
    "CHECK the number range 1 to 100 is displayed before the game starts.",
  ],
  prep: [
    "Materials: one mini whiteboard and marker per student, number range on display.",
    "Unit Progress Cards come out of the tub at the end.",
    ATTRIBUTION,
  ],
  tag: "[Setup | Planning | HITS 2]",
});

const NOTES_DONOW = T.composeGlanceNotes({
  answer: "who, what, where",
  beats: [
    ["ASK: Which question is each one?",
      "10 sec think. Cue: Everyone signs it to me on three... one, two, three.",
      "EXPECT: WHO, WHAT, WHERE"],
    "REVEAL after the room has answered, one click.",
    ["SAY: All three get you a person. None of them gets you a date.",
      "This term is full of dates."],
  ],
  prep: [
    "Retrieval from Lesson 2, mixed with Lesson 1 sport signs if time allows.",
    "The bridge names the gap today fills.",
  ],
  tag: "[Do Now | Attention, focus and regulation | HITS 6]",
});

const NOTES_LI = T.composeGlanceNotes({
  beats: [
    ["POINT to each criterion.", "SAY: Numbers today, and a second way to fix a breakdown."],
    "SAY: By the end everyone can write a year I sign to them.",
  ],
  prep: [
    "SC1 is a number to 100, reachable by every student.",
    "SC2 is what the exit ticket collects. SC3 is the clarification move.",
  ],
  tag: "[LI and SC | Planning | HITS 1]",
});

const NOTES_ANCHOR = T.composeGlanceNotes({
  beats: [
    ["SAY: Same four moves. Set it up, ask it, watch it, check it.",
      "With a number question, setting up means saying what you are counting."],
    ["SAY: Otherwise they get the how many and no idea what of.",
      "Medals? Years? Brothers?"],
    "MODEL setting the topic, then the number question, at pace.",
  ],
  prep: [
    "The anchor is restated in these exact words every week. Do not reword it.",
    "Naming what is counted is the transfer from Lesson 2's topic-first move.",
  ],
  tag: "[I Do | Explicit teaching | HITS 3]",
});

const NOTES_SIGNS = T.composeGlanceNotes({
  answer: "WHEN, HOW-MANY, YEAR",
  beats: [
    "MODEL each sign six times, then in a question, from the school reference.",
    ["SAY: When asks for a time. How many asks for a number.",
      "They are not interchangeable and students will swap them."],
    "SAY: Numbers move fast and they do not repeat themselves. Watch it all.",
  ],
  trap: ["asking how many when they meant when.",
    "Fix: two headings, sign six questions, students sort, then redo."],
  stretch: "ask both about the same fact, then say which answer was more useful.",
  help: "the two question words in English on the desk, so recall is the sign only.",
  prep: [
    "HOW-MANY carries quantity and number in Signbank; the sense is right.",
    "Confirm all three from the school reference before teaching.",
  ],
  tag: "[I Do | Explicit teaching | HITS 3, 4]",
});

const NOTES_YEARS = T.composeGlanceNotes({
  answer: "1955, 1965, 2005",
  beats: [
    ["SAY: Three years to hold onto this term.",
      "Nineteen fifty five. Nineteen sixty five. Two thousand and five."],
    "SAY: You will meet all three again next week, in order.",
    "MODEL each year twice, with a clear pause between the two halves.",
  ],
  trap: ["catching nineteen and losing sixty five.",
    "Fix: two boxes on the board, sign each half, then join them."],
  prep: [
    "CHECK GRAMMAR: how a four-digit year is produced is a reference question.",
    "Confirm it before the lesson. Do not improvise a year in front of them.",
  ],
  tag: "[I Do | Explicit teaching | HITS 4]",
});

const NOTES_REPAIR = T.composeGlanceNotes({
  answer: "WHAT MEAN?",
  beats: [
    ["SAY: I understood the number but not what it was the year of.",
      "Again would get me the same year, and I already had the year."],
    ["SAY: What mean? gets me an explanation, which is what I needed.",
      "Two repairs now, and they do different jobs."],
    "MODEL both back to back so the difference is visible.",
  ],
  trap: ["using again when a repeat will not help.",
    "Fix: student names which repair they need before they use it."],
  prep: [
    "WHAT MEAN? is the school's supplied fixed form. Reproduce it exactly.",
    "MEAN shows a lookup card: its first Signbank sense is SIGN LANGUAGE.",
    "Confirm the form you use from the school reference before teaching it.",
  ],
  tag: "[I Do | Explicit teaching | HITS 3]",
});

const NOTES_GAME = T.composeGlanceNotes({
  answer: "more, less or correct, until the partner's number is found",
  beats: [
    ["SET UP: range 1 to 100 on display. Each student hides a number.",
      "Boards face down between turns."],
    "SAY: A guess is signed, never written, or we stop practising numbers.",
    ["CIRCULATE: watch for guesses written on boards.",
      "That is the failure mode and it kills the game."],
    "COLLECT: first to find their partner's number wins, then swap who starts.",
  ],
  trap: ["writing the guess instead of signing it.",
    "Fix: turn the board over, student signs the guess again."],
  prep: [
    "Game 3 More or Less, from the unit game bank.",
    "Extension range is 1880 to 2026, which is next week's timeline practice.",
  ],
  tag: "[We Do | Collaborative learning | HITS 5]",
});

const NOTES_CHECK = T.composeGlanceNotes({
  answer: "1965",
  beats: [
    ["ASK: Write the year I sign.",
      "5 sec think. Cue: Write it. Boards up on three... one, two, three.",
      "EXPECT: 1965"],
    ["SCAN the boards, back row first.",
      "80%+ -> go to You Do. Lesson 4's timeline is safe.",
      "Less -> write 19 and 65 in two boxes, sign each, join, re-ask with 1955."],
    ["FOLLOW UP one student.",
      "SAY: How did you know it was sixty five and not fifty six?"],
  ],
  trap: ["the two halves of the year fusing into one number.",
    "Fix: separate boxes, clear pause, student rewrites it."],
  prep: [
    "The decision point that changes Lesson 4.",
    "Lesson 4 cannot run if years are not landing, so pivot rather than push on.",
  ],
  tag: "[CFU | Evaluating impact | HITS 7]",
});

const NOTES_YOUDO = T.composeGlanceNotes({
  beats: [
    ["SAY: Voices off. Ask your partner when they started their favourite sport.",
      "Then ask how many years they have played it."],
    "SAY: Digits only on the board. Use again, or what mean, if you miss it.",
    ["CIRCULATE: look for the right repair being chosen.",
      "That is today's real target, not the numbers."],
  ],
  stretch: "play More or Less with years only, from 1880 to 2026.",
  help: "a 1 to 100 number line on the desk, so a guess is point plus sign.",
  prep: [
    "Different content from the game: a hidden number there, a real one here.",
    "Year 5 reads years receptively. Year 6 produces them.",
  ],
  tag: "[You Do | Explicit teaching | HITS 10]",
});

const NOTES_EXIT = T.composeGlanceNotes({
  answer: "the four digits of the year signed at the door",
  beats: [
    "COLLECT at the door on boards, held up as they leave.",
    "Two piles: four digits right, or not.",
    "Then the Unit Progress Card, Lesson 3 row, three ticks.",
  ],
  prep: [
    "Assesses SC2. This is the receptive check, not the expressive one.",
    "A student who cannot read a year cannot order the timeline next week.",
  ],
  tag: "[Exit ticket | Evaluating impact | HITS 7]",
});

const NOTES_CLOSING = T.composeGlanceNotes({
  beats: [
    "SAY: Turn and tell your partner how long they have played their sport.",
    "SAY: Next week those three years turn into a story about Deaf sport.",
  ],
  prep: [
    "Protocol practised today: eye gaze held through a whole answer.",
    "Running late: cut You Do to one question. Never cut the year practice.",
  ],
  tag: "[Closing | Planning | HITS 9]",
});

/* ---------------------------------------------------------------- deck ----- */

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  titleSlide(pres, UNIT_TITLE, "Session 3: When and how many",
    `Years 5-6 Auslan | Session ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  noPrintResourceSlide(T, pres, {
    items: [
      "One mini whiteboard and marker per student",
      "The number range 1 to 100 displayed where everyone can see it",
      "Unit Progress Cards in the tub, out at the end",
    ],
    notes: NOTES_RESOURCES, footer: FOOTER,
  });

  REPORT.add(signRecallSlide(T, pres, {
    title: "Do Now: which question is each one?",
    lead: "Sign the meaning back to me. Answers come up on the click.",
    signs: [
      { gloss: "WHO", meaning: "who" },
      { gloss: "WHAT", meaning: "what" },
      { gloss: "WHERE", meaning: "where" },
    ],
    notes: NOTES_DONOW, footer: FOOTER,
  }).results);

  liSlide(pres,
    "I am learning to ask when something happened and how many, using years and numbers.",
    [
      "I can sign a number to 100.",
      "I can ask when and how many, and read a year in an answer.",
      "I can use WHAT MEAN? to ask for an explanation, not just a repeat.",
    ],
    NOTES_LI, FOOTER);

  // I Do - anchor
  const anchorSlide = pres.addSlide();
  anchorSlide.background = { color: C.BG_LIGHT };
  T.addTopBar(anchorSlide, C.PRIMARY);
  T.addTitle(anchorSlide, "Say what you are counting");
  addTextOnShape(anchorSlide, ANCHOR, {
    x: 0.5, y: 1.45, w: 9.0, h: 0.72, rectRadius: 0.1, fill: { color: C.PRIMARY },
  }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
  addTextOnShape(anchorSlide, "How many... what?", {
    x: 1.4, y: 2.5, w: 7.2, h: 0.7, rectRadius: 0.1, fill: { color: C.SECONDARY },
  }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
  anchorSlide.addText("Medals? Years? Brothers? Set up what you are counting before you ask the number.", {
    x: 1.4, y: 3.45, w: 7.2, h: 0.8,
    fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL, align: "center", margin: 0,
  });
  T.addFooter(anchorSlide, FOOTER);
  anchorSlide.addNotes(NOTES_ANCHOR);
  runSlideDiagnostics(anchorSlide, pres);

  // I Do - the three signs
  REPORT.add(signSlide(T, pres, {
    title: "Asking about time and number",
    lead: "Watch me first. Then copy. Six times each.",
    signs: [
      { gloss: "WHEN", meaning: "when" },
      { gloss: "HOW-MANY", meaning: "how many" },
      { gloss: "YEAR", meaning: "year" },
    ],
    notes: NOTES_SIGNS, footer: FOOTER,
  }).results);

  // I Do - three years, built on clicks
  const yearSlide = pres.addSlide();
  yearSlide.background = { color: C.BG_LIGHT };
  T.addTopBar(yearSlide, C.PRIMARY);
  T.addTitle(yearSlide, "Three years to hold onto");
  const YEARS = [
    ["1955", "Deaf Sports Australia joins the world body"],
    ["1965", "Australia's first medals at the world games"],
    ["2005", "Melbourne hosts the Deaflympics"],
  ];
  const yw = (9.0 - 0.3 * 2) / 3;
  clickBuild(yearSlide, YEARS.map(([yr, what], i) => () => {
    const yx = 0.5 + i * (yw + 0.3);
    addTextOnShape(yearSlide, yr, {
      x: yx, y: 1.85, w: yw, h: 1.0, rectRadius: 0.1, fill: { color: C.PRIMARY },
    }, { fontSize: 36, fontFace: FONT_H, color: C.WHITE, bold: true });
    yearSlide.addText(what, {
      x: yx, y: 3.0, w: yw, h: 1.0, fontSize: 14, fontFace: FONT_B,
      color: C.CHARCOAL, align: "center", margin: 0,
    });
  }));
  T.addFooter(yearSlide, FOOTER);
  yearSlide.addNotes(NOTES_YEARS);

  // I Do - the second repair move
  const repairSlide = pres.addSlide();
  repairSlide.background = { color: C.BG_LIGHT };
  T.addTopBar(repairSlide, C.ALERT);
  T.addTitle(repairSlide, "Two repairs, two different jobs");
  REPORT.add(signCardRow(T, repairSlide, [
    { gloss: "AGAIN", meaning: "again" },
    { gloss: "WHAT", meaning: "what" },
    { gloss: "MEAN", meaning: "meaning", lookupWord: "meaning" },
  ], { y: 1.45, bottom: 4.15, center: false }));
  addTextOnShape(repairSlide, "Again = show me that again", {
    x: 0.7, y: 4.3, w: 4.15, h: 0.6, rectRadius: 0.1, fill: { color: C.SECONDARY },
  }, { fontSize: 15, fontFace: FONT_B, color: C.WHITE, bold: true });
  addTextOnShape(repairSlide, "What mean? = explain it to me", {
    x: 5.15, y: 4.3, w: 4.15, h: 0.6, rectRadius: 0.1, fill: { color: C.ALERT },
  }, { fontSize: 15, fontFace: FONT_B, color: C.WHITE, bold: true });
  T.addFooter(repairSlide, FOOTER);
  repairSlide.addNotes(NOTES_REPAIR);
  runSlideDiagnostics(repairSlide, pres);

  // We Do - Game 3
  const gameSlide = pres.addSlide();
  gameSlide.background = { color: C.BG_LIGHT };
  T.addTopBar(gameSlide, C.SUCCESS);
  T.addTitle(gameSlide, "Game 3 More or Less");
  stepCard(T, gameSlide, {
    title: "Guess your partner's number",
    steps: [
      "Write a hidden number from 1 to 100",
      "Board face down between turns",
      "Sign your guess. Never write it",
      "Partner answers more, less or correct",
    ],
    x: 0.5, y: 1.5, w: 5.4, h: 2.9, strip: C.SUCCESS,
  });
  addTextOnShape(gameSlide, "1 to 100", {
    x: 6.1, y: 1.7, w: 3.4, h: 0.95, rectRadius: 0.12, fill: { color: C.PRIMARY },
  }, { fontSize: 30, fontFace: FONT_H, color: C.WHITE, bold: true });
  ["more", "less", "correct"].forEach((wd, i) => {
    addTextOnShape(gameSlide, wd, {
      x: 6.1, y: 2.85 + i * 0.55, w: 3.4, h: 0.45, rectRadius: 0.08,
      fill: { color: C.BG_CARD }, line: { color: C.SUCCESS, width: 1 },
    }, { fontSize: 15, fontFace: FONT_B, color: C.CHARCOAL });
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
  checkSlide.addText("Write the year I sign to you", {
    x: 1.1, y: 3.0, w: 7.8, h: 0.7,
    fontSize: 28, fontFace: FONT_H, color: C.PRIMARY, bold: true, align: "center", margin: 0,
  });
  checkSlide.addText("Four digits. Watch it all, then write.", {
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
  T.addTitle(youSlide, "You Do: ask about their sport");
  stepCard(T, youSlide, {
    title: "Two questions, then swap",
    steps: [
      "When did you start your favourite sport?",
      "How many years have you played it?",
      "Write the answers as digits only",
      "Miss it? Use again, or what mean",
    ],
    x: 0.5, y: 1.5, w: 5.4, h: 2.9, strip: C.SUCCESS,
  });
  addTextOnShape(youSlide, "Voices off", {
    x: 6.1, y: 1.6, w: 3.4, h: 0.9, rectRadius: 0.12, fill: { color: C.ALERT },
  }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
  addTextOnShape(youSlide, "Digits only", {
    x: 6.1, y: 2.75, w: 3.4, h: 0.75, rectRadius: 0.1, fill: { color: C.SECONDARY },
  }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
  youSlide.addText("No words on the board", {
    x: 6.1, y: 3.65, w: 3.4, h: 0.4,
    fontSize: 15, fontFace: FONT_B, color: C.CHARCOAL, align: "center", margin: 0,
  });
  T.addFooter(youSlide, FOOTER);
  youSlide.addNotes(NOTES_YOUDO);
  runSlideDiagnostics(youSlide, pres);

  exitTicketSlide(pres,
    ["Write the year I sign to you. Four digits."],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner how long they have played their sport.",
      scItems: [
        "I can sign a number to 100.",
        "I can ask when and how many, and read a year in an answer.",
        "I can use WHAT MEAN? to ask for an explanation, not just a repeat.",
      ],
      selfAssessment: "Unit Progress Card, Lesson 3 row",
    },
    NOTES_CLOSING, FOOTER);

  const pptxPath = path.join(OUT_DIR, `Deaf Sports Session ${SESSION} When And How Many.pptx`);
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  REPORT.print(`Session ${SESSION}`);
  console.log(`Session ${SESSION} build complete.`);
}

build().catch((err) => { console.error(err); process.exit(1); });
