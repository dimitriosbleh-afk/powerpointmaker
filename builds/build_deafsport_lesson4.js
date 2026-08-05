"use strict";

// Deaf Sports in Australia (Years 5-6 Auslan, Term 4) - Session 4 of 10.
// BACK THEN AND AFTER: ordering real events in Deaf sport history.
//   4 production signs: BEFORE, AFTER, PAST, THEN. Game 4 Order The Years.
// Every date on the timeline cards comes from Deaf Sports Australia's own
// history page and fact sheet. Sources are in the unit document appendix.
// Unit anchor: Set it up, ask it, watch it, check it.

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
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

const SESSION = 4;
const TOTAL = 10;
const UNIT_TITLE = "Deaf Sports in Australia";
const ANCHOR = "Set it up, ask it, watch it, check it.";
const FOOTER = `${UNIT_TITLE} | Session ${SESSION} of ${TOTAL} | Years 5-6 Auslan`;
const OUT_DIR = "output/DeafSport_Lesson4_Back_Then_And_After";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const TIMELINE_RES = makeSessionResource(SESSION,
  "Deaf Sport Timeline Cards",
  "Ten dated events. Cut and laminate once; reused in Session 10.");
const RESOURCE_ITEMS = [TIMELINE_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

const REPORT = createSignReport();

// Cards 1 to 8 are the Session 4 set; 9 and 10 join them in Session 10.
const TIMELINE = [
  ["1880s", "A Deaf cricket club starts in Melbourne. It is still going today as the Melbourne Deaf Cricket Club, and it is one of the oldest Deaf sport clubs in the world."],
  ["1924", "The first Deaflympic Games are held in Paris."],
  ["1954", "Deaf Sports Australia is set up as the national body for Deaf sport in Australia."],
  ["1955", "Deaf Sports Australia joins the international Deaf sports committee."],
  ["1964", "The first Australian Deaf Games are held in Sydney, over the summer of 1964 and 1965."],
  ["1965", "Australia sends two athletes to the world games for the Deaf. Both of them win a medal."],
  ["1985", "The national Deaf sports body is formally recognised and funded."],
  ["2005", "Melbourne hosts the Deaflympics. It is the only time Australia has ever hosted them."],
  ["2011", "The Active Deaf Kids school program starts."],
  ["2017", "All sixteen national and six state Deaf sport organisations join together as one body."],
];

/* ---------------------------------------------------------------- notes ---- */

const NOTES_TITLE = "Session 4 of 10. Real dated events from Deaf Sports Australia's own history. Timeline cards need cutting before this lesson.";

const NOTES_RESOURCES = T.composeGlanceNotes({
  beats: [
    ["SHOW while students settle.", "SAY: One set of cards per table of four, face down."],
    "SET UP: circle seating so every face is visible before the game starts.",
  ],
  prep: [
    "Print, cut and laminate the timeline cards once. Session 10 reuses them.",
    "Every date comes from Deaf Sports Australia. Check them before printing.",
    ATTRIBUTION,
  ],
  tag: "[Setup | Planning | HITS 2]",
});

const NOTES_DONOW = T.composeGlanceNotes({
  answer: "when, how many, year",
  beats: [
    ["ASK: What does each sign mean?",
      "10 sec think. Cue: Everyone signs it to me on three... one, two, three.",
      "EXPECT: WHEN, HOW-MANY, YEAR"],
    "REVEAL after the room has answered, one click.",
    ["SAY: You can write those three years. Now you need to say which came first.",
      "You cannot do that with numbers alone."],
  ],
  prep: [
    "Retrieval from Lesson 3, then the bridge into ordering.",
    "If years were shaky last week, reteach the two-box split before you go on.",
  ],
  tag: "[Do Now | Attention, focus and regulation | HITS 6]",
});

const NOTES_LI = T.composeGlanceNotes({
  beats: [
    ["POINT to each criterion.", "SAY: Four time signs today, and a real story to put in order."],
    "SAY: By the end everyone can say which of two events came first.",
  ],
  prep: [
    "SC1 is two events, reachable by every student.",
    "SC2 is what the exit ticket collects. SC3 adds the reason.",
  ],
  tag: "[LI and SC | Planning | HITS 1]",
});

const NOTES_ANCHOR = T.composeGlanceNotes({
  answer: "PAST",
  beats: [
    ["SAY: Same four moves. Set it up, ask it, watch it, check it.",
      "The setting up today is the time. Tell them when we are first."],
    ["SAY: If I do not, everything I sign after that lands in the wrong year.",
      "Back then. Eighteen eighties."],
    "MODEL setting the time, then the event, before adding any detail.",
  ],
  prep: [
    "The anchor is restated in these exact words every week. Do not reword it.",
    "PAST is introduced here as the setting-up move, not as vocabulary.",
  ],
  tag: "[I Do | Explicit teaching | HITS 3]",
});

const NOTES_SIGNS = T.composeGlanceNotes({
  answer: "BEFORE, AFTER, THEN",
  beats: [
    "MODEL each sign six times against a real pair of events, not an abstract one.",
    ["SAY: Back in the eighteen eighties, Deaf men in Melbourne started a cricket club.",
      "It is still going today."],
    ["SAY: That was before Deaf Sports Australia existed, which started in 1954.",
      "Then, in 1964, the first Australian Deaf Games were held in Sydney."],
  ],
  trap: ["reading the second pair of digits and calling 1880s the later date.",
    "Fix: eight students hold cards and walk into order, class directs them."],
  stretch: "recount three events and add what changed because of the middle one.",
  help: "three cards face up in the right order, so the task is reading not recall.",
  prep: [
    "BEFORE has two widely used forms, one for time and one for place.",
    "Teach the time form to production; students only recognise the other.",
    "Confirm both from the school reference before teaching.",
  ],
  tag: "[I Do | Explicit teaching | HITS 3, 4]",
});

const NOTES_WRONG = T.composeGlanceNotes({
  beats: [
    ["GET IT WRONG: recount the club and the Games with the time signs swapped.",
      "SAY: I just told you the Games came first and the club came after."],
    ["SAY: That is a hundred years out.",
      "Watch it again with the time signs the right way round."],
    "MODEL it correctly, at pace, without commentary this time.",
  ],
  trap: ["missing the time sign because it moves fast.",
    "Fix: sign the time sign alone first, then the whole event."],
  prep: [
    "The deliberate error is the hardest thing to improvise, so it is scripted.",
    "It targets the exact confusion the check will test.",
  ],
  tag: "[I Do | Explicit teaching | HITS 3]",
});

const NOTES_GAME = T.composeGlanceNotes({
  answer: "the eight cards laid out oldest to newest",
  beats: [
    ["SET UP: groups of four in a circle, cards face down in the middle.",
      "Two cards each, flat on the desk, never in a hand."],
    "SAY: Sign your event and its year. The group decides where it goes.",
    ["CIRCULATE: watch for hands reaching for cards mid-turn.",
      "A group reaching has stopped watching."],
    "COLLECT: read the finished line back around the circle.",
  ],
  trap: ["placing a card by guessing rather than by asking for the year again.",
    "Fix: cover the line, ask for the year, student re-places it."],
  prep: [
    "Game 4 Order The Years. Eight cards for four players.",
    "The turn passes left in a fixed direction so nobody is called on.",
  ],
  tag: "[We Do | Collaborative learning | HITS 5]",
});

const NOTES_CHECK = T.composeGlanceNotes({
  answer: "the Melbourne cricket club",
  beats: [
    ["ASK: Which came first, the cricket club or the first Australian Deaf Games?",
      "5 sec think. Cue: Everyone signs it to me on three... one, two, three.",
      "EXPECT: the cricket club"],
    ["SCAN the room, back row first.",
      "80%+ -> go to You Do with three events.",
      "Less -> eight students hold cards and walk into order, then re-ask."],
    ["FOLLOW UP one student.",
      "SAY: You said the club came first. How do you know?"],
  ],
  trap: ["treating the biggest number as the earliest date.",
    "Fix: read the century aloud, not the last two digits, then redo."],
  prep: [
    "The decision point that changes Lesson 5.",
    "Ordering has to hold before the impact question in Lesson 5 makes sense.",
  ],
  tag: "[CFU | Evaluating impact | HITS 7]",
});

const NOTES_YOUDO = T.composeGlanceNotes({
  beats: [
    ["SAY: Voices off. Pick three events and recount them in order to your partner.",
      "Use at least two of the four time signs."],
    "SAY: Your partner writes the three years in the order you give them.",
    ["CIRCULATE: check the written order against the cards.",
      "A mismatch is the evidence, not a mistake to correct mid-turn."],
  ],
  stretch: "add one sentence about what changed because of the middle event.",
  help: "three cards face up in the correct order, so the task becomes reading.",
  prep: [
    "Different content from the game: ordering for a group there, recounting here.",
    "Year 5 recounts two events with one time sign. Year 6 recounts three.",
  ],
  tag: "[You Do | Explicit teaching | HITS 10]",
});

const NOTES_EXIT = T.composeGlanceNotes({
  answer: "whichever of the two events is the earlier one",
  beats: [
    "COLLECT at the door: sign two events, student signs which came first.",
    "Two piles: right order, or not.",
    "Then the Unit Progress Card, Lesson 4 row, three ticks.",
  ],
  care: "the 1880s club's original name used words we do not use now.",
  prep: [
    "The cards say Melbourne Deaf Cricket Club. If a student finds the old name,",
    "name it as historical, from a time Deaf people did not choose the words.",
  ],
  tag: "[Exit ticket | Evaluating impact | HITS 7]",
});

const NOTES_CLOSING = T.composeGlanceNotes({
  beats: [
    "SAY: Turn and tell your partner which event you think mattered most.",
    "SAY: Next week you find out how Deaf sport actually runs.",
  ],
  prep: [
    "Protocol practised today: turn taking, one signer at a time.",
    "Running late: cut the game to four cards. Never cut PAST.",
  ],
  tag: "[Closing | Planning | HITS 9]",
});

/* ---------------------------------------------------------------- deck ----- */

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  titleSlide(pres, UNIT_TITLE, "Session 4: Back then, and after",
    `Years 5-6 Auslan | Session ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  REPORT.add(signRecallSlide(T, pres, {
    title: "Do Now: what does each one mean?",
    lead: "Sign the meaning back to me. Answers come up on the click.",
    signs: [
      { gloss: "WHEN", meaning: "when" },
      { gloss: "HOW-MANY", meaning: "how many" },
      { gloss: "YEAR", meaning: "year" },
    ],
    notes: NOTES_DONOW, footer: FOOTER,
  }).results);

  liSlide(pres,
    "I am learning to put events in order using before, after, back then and then.",
    [
      "I can put two events in the right order.",
      "I can recount three events from Deaf sport history in order using the time signs.",
      "I can say which event came first and explain how I know.",
    ],
    NOTES_LI, FOOTER);

  // I Do - anchor, set the time first
  const anchorSlide = pres.addSlide();
  anchorSlide.background = { color: C.BG_LIGHT };
  T.addTopBar(anchorSlide, C.PRIMARY);
  T.addTitle(anchorSlide, "Set the time first");
  addTextOnShape(anchorSlide, ANCHOR, {
    x: 0.5, y: 1.45, w: 9.0, h: 0.72, rectRadius: 0.1, fill: { color: C.PRIMARY },
  }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
  addTextOnShape(anchorSlide, "Back then... 1880s", {
    x: 1.4, y: 2.5, w: 7.2, h: 0.75, rectRadius: 0.1, fill: { color: C.SECONDARY },
  }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
  anchorSlide.addText("Miss the time sign and you have the right event in the wrong century.", {
    x: 1.4, y: 3.5, w: 7.2, h: 0.7,
    fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL, align: "center", margin: 0,
  });
  T.addFooter(anchorSlide, FOOTER);
  anchorSlide.addNotes(NOTES_ANCHOR);
  runSlideDiagnostics(anchorSlide, pres);

  // I Do - the time signs
  REPORT.add(signSlide(T, pres, {
    title: "Putting events in order",
    lead: "Watch me first. Then copy. Six times each.",
    signs: [
      { gloss: "PAST", meaning: "back then" },
      { gloss: "BEFORE", meaning: "before" },
      { gloss: "AFTER", meaning: "after" },
      { gloss: "THEN", meaning: "then" },
    ],
    notes: NOTES_SIGNS, footer: FOOTER,
  }).results);

  // I Do - three real events, built on clicks
  const storySlide = pres.addSlide();
  storySlide.background = { color: C.BG_LIGHT };
  T.addTopBar(storySlide, C.PRIMARY);
  T.addTitle(storySlide, "Watch me get it wrong first");
  const THREE = [
    ["1880s", "A Deaf cricket club starts in Melbourne"],
    ["1954", "Deaf Sports Australia is set up"],
    ["1964", "The first Australian Deaf Games, in Sydney"],
  ];
  clickBuild(storySlide, THREE.map(([yr, what], i) => () => {
    const ty = 1.75 + i * 1.02;
    addTextOnShape(storySlide, yr, {
      x: 0.8, y: ty, w: 1.85, h: 0.78, rectRadius: 0.1, fill: { color: C.PRIMARY },
    }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
    addTextOnShape(storySlide, what, {
      x: 2.85, y: ty, w: 6.35, h: 0.78, rectRadius: 0.1,
      fill: { color: C.BG_CARD }, line: { color: C.SECONDARY, width: 1 },
    }, { fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL });
  }));
  T.addFooter(storySlide, FOOTER);
  storySlide.addNotes(NOTES_WRONG);

  // We Do - Game 4
  const gameSlide = pres.addSlide();
  gameSlide.background = { color: C.BG_LIGHT };
  T.addTopBar(gameSlide, C.SUCCESS);
  T.addTitle(gameSlide, "Game 4 Order The Years");
  stepCard(T, gameSlide, {
    title: "Lay the whole timeline out",
    steps: [
      "Two cards each, flat on your desk",
      "Sign your event and its year",
      "The group decides where it goes",
      "Hands off the table while somebody signs",
    ],
    x: 0.5, y: 1.5, w: 5.4, h: 2.9, strip: C.SUCCESS,
  });
  addTextOnShape(gameSlide, "Oldest", {
    x: 6.1, y: 1.7, w: 3.4, h: 0.55, rectRadius: 0.08, fill: { color: C.PRIMARY },
  }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
  gameSlide.addText("|", {
    x: 6.1, y: 2.32, w: 3.4, h: 1.35, fontSize: 30, fontFace: FONT_B,
    color: C.SECONDARY, align: "center", margin: 0,
  });
  addTextOnShape(gameSlide, "Newest", {
    x: 6.1, y: 3.72, w: 3.4, h: 0.55, rectRadius: 0.08, fill: { color: C.SECONDARY },
  }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
  T.addFooter(gameSlide, FOOTER);
  gameSlide.addNotes(NOTES_GAME);
  runSlideDiagnostics(gameSlide, pres);

  // Check it
  const checkSlide = pres.addSlide();
  checkSlide.background = { color: C.BG_LIGHT };
  T.addTopBar(checkSlide, C.ALERT);
  T.addTitle(checkSlide, "Check it");
  addTextOnShape(checkSlide, "Everyone signs it to me on three", {
    x: 1.1, y: 1.75, w: 7.8, h: 0.8, rectRadius: 0.12, fill: { color: C.ALERT },
  }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
  checkSlide.addText("Which came first?", {
    x: 1.1, y: 2.7, w: 7.8, h: 0.55,
    fontSize: 26, fontFace: FONT_H, color: C.PRIMARY, bold: true, align: "center", margin: 0,
  });
  addTextOnShape(checkSlide, "The Melbourne cricket club", {
    x: 0.9, y: 3.35, w: 4.05, h: 0.7, rectRadius: 0.1,
    fill: { color: C.BG_CARD }, line: { color: C.SECONDARY, width: 1 },
  }, { fontSize: 15, fontFace: FONT_B, color: C.CHARCOAL });
  addTextOnShape(checkSlide, "The first Australian Deaf Games", {
    x: 5.15, y: 3.35, w: 4.05, h: 0.7, rectRadius: 0.1,
    fill: { color: C.BG_CARD }, line: { color: C.SECONDARY, width: 1 },
  }, { fontSize: 15, fontFace: FONT_B, color: C.CHARCOAL });
  T.addFooter(checkSlide, FOOTER);
  checkSlide.addNotes(NOTES_CHECK);
  runSlideDiagnostics(checkSlide, pres);

  // You Do
  const youSlide = pres.addSlide();
  youSlide.background = { color: C.BG_LIGHT };
  T.addTopBar(youSlide, C.SUCCESS);
  T.addTitle(youSlide, "You Do: recount three events");
  stepCard(T, youSlide, {
    title: "Tell it, then swap",
    steps: [
      "Pick three events from the cards",
      "Recount them in order to your partner",
      "Use at least two time signs",
      "Your partner writes the three years down",
    ],
    x: 0.5, y: 1.5, w: 5.4, h: 2.9, strip: C.SUCCESS,
  });
  addTextOnShape(youSlide, "Voices off", {
    x: 6.1, y: 1.6, w: 3.4, h: 0.9, rectRadius: 0.12, fill: { color: C.ALERT },
  }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
  ["back then", "before", "after", "then"].forEach((wd, i) => {
    addTextOnShape(youSlide, wd, {
      x: 6.1, y: 2.68 + i * 0.53, w: 3.4, h: 0.44, rectRadius: 0.07,
      fill: { color: C.BG_CARD }, line: { color: C.SECONDARY, width: 1 },
    }, { fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL });
  });
  T.addFooter(youSlide, FOOTER);
  youSlide.addNotes(NOTES_YOUDO);
  runSlideDiagnostics(youSlide, pres);

  exitTicketSlide(pres,
    ["I will sign two events. Sign me which one came first."],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner which event you think mattered most, and why.",
      scItems: [
        "I can put two events in the right order.",
        "I can recount three events from Deaf sport history in order using the time signs.",
        "I can say which event came first and explain how I know.",
      ],
      selfAssessment: "Unit Progress Card, Lesson 4 row",
    },
    NOTES_CLOSING, FOOTER);

  const pptxPath = path.join(OUT_DIR, `Deaf Sports Session ${SESSION} Back Then And After.pptx`);
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  /* ------------------------------------------------------------- pdf ------ */

  const doc = createPdf({ title: TIMELINE_RES.name });
  let y = addPdfHeader(doc, TIMELINE_RES.name, {
    subtitle: "Ten dated events in Deaf sport in Australia. Cut and laminate.",
    color: C.PRIMARY,
    lessonInfo: `Session ${SESSION} | Game 4 Order The Years | one set per group of four`,
  });

  y = addTipBox(doc,
    "Cards 1 to 8 are the Session 4 set. Cards 9 and 10 join them in Session 10. Every date here comes from Deaf Sports Australia's own history page and sport modifications fact sheet; check them against the unit document's Sources appendix before printing.",
    y, { color: C.ACCENT });

  y = addSectionHeading(doc, "Cut along the dashed lines", y, { color: C.SECONDARY });

  const cardW = 495;
  const cardH = 64;
  TIMELINE.forEach(([year, text], i) => {
    if (y + cardH > 780) { doc.addPage(); y = 60; }
    doc.save();
    doc.rect(50, y, cardW, cardH - 8).lineWidth(0.8).dash(3, { space: 2 })
      .strokeColor("#" + C.MUTED).stroke();
    doc.undash();
    doc.font("Sans-Bold").fontSize(17).fillColor("#" + C.PRIMARY)
      .text(year, 62, y + 18, { width: 70 });
    doc.font("Sans").fontSize(10.5).fillColor("#" + C.CHARCOAL)
      .text(text, 140, y + 12, { width: cardW - 105 });
    doc.restore();
    y += cardH;
  });

  addPdfFooter(doc, `Session ${SESSION} | Deaf Sport Timeline Cards | Years 5-6 Auslan`);
  await writePdf(doc, path.join(OUT_DIR, TIMELINE_RES.fileName));
  console.log("PDF written: " + TIMELINE_RES.fileName);

  REPORT.print(`Session ${SESSION}`);
  console.log(`Session ${SESSION} build complete.`);
}

build().catch((err) => { console.error(err); process.exit(1); });
