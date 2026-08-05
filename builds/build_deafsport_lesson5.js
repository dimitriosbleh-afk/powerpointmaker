"use strict";

// Deaf Sports in Australia (Years 5-6 Auslan, Term 4) - Session 5 of 10.
// STARTING LIGHTS AND FLAGS: how Deaf sport is set up to be seen.
//   4 production signs: START, LIGHT, FLAG, CHANGE. Game 5 How Many Medals.
// Every modification on these slides comes from Deaf Sports Australia's own
// Sport Modifications fact sheet. Sources are in the unit document appendix.
// Care: barriers are design problems somebody already solved. Never framed as
// something an athlete had to overcome.

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

const SESSION = 5;
const TOTAL = 10;
const UNIT_TITLE = "Deaf Sports in Australia";
const ANCHOR = "Set it up, ask it, watch it, check it.";
const FOOTER = `${UNIT_TITLE} | Session ${SESSION} of ${TOTAL} | Years 5-6 Auslan`;
const OUT_DIR = "output/DeafSport_Lesson5_Starting_Lights_And_Flags";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const TALLY_RES = makeSessionResource(SESSION,
  "Medal Tally Sheets A and B",
  "Half the class gets A, half gets B. An information gap for asking how many.");
const RESOURCE_ITEMS = [TALLY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

const REPORT = createSignReport();

const SPORTS = ["Swimming", "Athletics", "Cricket", "Basketball", "Tennis", "Sock wrestling"];
const TALLY_A = [7, 12, 3, 9, 5, 41];
const TALLY_B = [11, 4, 8, 6, 14, 38];

/* ---------------------------------------------------------------- notes ---- */

const NOTES_TITLE = "Session 5 of 10. How Deaf sport actually runs. Every modification here is from Deaf Sports Australia's own fact sheet.";

const NOTES_RESOURCES = T.composeGlanceNotes({
  beats: [
    ["SHOW while students settle.", "SAY: Half of you get sheet A, half get sheet B."],
    "SET UP: one propped folder per pair, so the sheet is hidden and the face is not.",
  ],
  prep: [
    "Print 50 of sheet A and 50 of sheet B across four classes. Not laminated.",
    "Folders from classroom stock. Pairs face each other, sheets flat.",
    ATTRIBUTION,
  ],
  tag: "[Setup | Planning | HITS 2]",
});

const NOTES_DONOW = T.composeGlanceNotes({
  answer: "back then, before, after",
  beats: [
    ["ASK: What does each sign mean?",
      "10 sec think. Cue: Everyone signs it to me on three... one, two, three.",
      "EXPECT: PAST, BEFORE, AFTER"],
    "REVEAL after the room has answered, one click.",
    ["SAY: You know when Deaf sport started.",
      "Today you find out how it runs, and it is cleverer than you think."],
  ],
  prep: [
    "Retrieval reaches back to Lesson 4, and one question sign from Lesson 2.",
    "The bridge sets up the whole lesson in one line.",
  ],
  tag: "[Do Now | Attention, focus and regulation | HITS 6]",
});

const NOTES_LI = T.composeGlanceNotes({
  beats: [
    ["POINT to each criterion.", "SAY: One idea today. Everything you need to know can be seen."],
    "SAY: By the end everyone can name one thing shown with a light or a flag.",
  ],
  prep: [
    "SC1 is naming one modification, reachable by every student.",
    "SC2 is what the exit ticket collects. SC3 generalises beyond sport.",
  ],
  tag: "[LI and SC | Planning | HITS 1]",
});

const NOTES_ANCHOR = T.composeGlanceNotes({
  answer: "START, LIGHT, FLAG",
  beats: [
    ["SAY: Same four moves. Set it up, ask it, watch it, check it.",
      "Deaf Sports Australia says Deaf athletes play every sport."],
    ["SAY: The modifications there are all do one job.",
      "They take something you were supposed to hear and make it something you can see."],
    "MODEL each of the three signs six times before the examples.",
  ],
  prep: [
    "The anchor is restated in these exact words every week. Do not reword it.",
    "The one-job framing is the whole lesson. Say it before any example.",
  ],
  tag: "[I Do | Explicit teaching | HITS 3]",
});

const NOTES_MODS = T.composeGlanceNotes({
  beats: [
    ["SAY: In swimming a light on the side of the pool is wired to the starter's gun.",
      "The gun fires, the light comes on, the race starts."],
    ["SAY: A referee gets an athlete's attention with a flag, or by waving.",
      "A captain who wants the referee taps them on the shoulder."],
    ["SAY: That is a normal thing to do, not a rude one.",
      "Scores go on a board every time, so nobody waits to be told."],
  ],
  trap: ["naming a modification without the problem it solved.",
    "Fix: three problems and three solutions on the board, students pair them."],
  prep: [
    "All four modifications come from the Deaf Sports Australia fact sheet.",
    "The shoulder tap is also this week's Deaf-friendly protocol.",
  ],
  tag: "[I Do | Explicit teaching | HITS 3, 4]",
});

const NOTES_CHANGE = T.composeGlanceNotes({
  answer: "CHANGE",
  beats: [
    ["SAY: Notice what changed in every one of those. Not the athlete.",
      "The starting gun changed. The referee changed. The scoreboard changed."],
    ["GET IT WRONG: SAY: I nearly told you Deaf athletes need special sports.",
      "That is wrong and it is worth knowing why."],
    ["SAY: Same sports, same rules. What got redesigned was the equipment.",
      "A barrier is a design problem somebody already fixed."],
  ],
  trap: ["hearing this as Deaf athletes overcoming something.",
    "Fix: redirect to the result. They are athletes with results."],
  stretch: "find a barrier in this classroom that could be designed away the same way.",
  help: "three picture prompts on the desk: a pool light, a flag, a scoreboard.",
  prep: [
    "This is the care point of the whole unit, taught rather than warned about.",
    "Never frame a Deaf athlete as inspirational for competing.",
  ],
  tag: "[I Do | Explicit teaching | HITS 3]",
});

const NOTES_GAME = T.composeGlanceNotes({
  answer: "the six numbers on the partner's sheet, then the totals compared",
  beats: [
    ["SET UP: pairs facing, sheet flat behind a propped folder.",
      "The folder hides the sheet, never the partner."],
    "SAY: Pencils down while your partner signs. Watch the whole answer, then write.",
    ["CIRCULATE: a pair writing mid-answer will be a number out within three turns.",
      "Fix it on the first turn, not the fourth."],
    "COLLECT: both add their totals. Ask who has more sock wrestling medals.",
  ],
  trap: ["writing while the partner is still signing.",
    "Fix: pencils on the desk, partner signs again, then write."],
  prep: [
    "Game 5 How Many Medals. The last row is meant to be ridiculous.",
    "It is the ending the class remembers, so do not cut it.",
  ],
  tag: "[We Do | Collaborative learning | HITS 5]",
});

const NOTES_CHECK = T.composeGlanceNotes({
  answer: "the light",
  beats: [
    ["ASK: What starts a race in Deaf swimming?",
      "5 sec think. Cue: Everyone signs it to me on three... one, two, three.",
      "EXPECT: LIGHT"],
    ["SCAN the room, back row first.",
      "80%+ -> go to You Do and add the explanation.",
      "Less -> three problems and three solutions on the board, pair them, re-ask."],
    ["FOLLOW UP one student.",
      "SAY: You said light. Add what the light replaced."],
  ],
  trap: ["naming the light with no problem attached to it.",
    "Fix: student finishes the sentence, the light replaced the..."],
  prep: [
    "The decision point that changes Lesson 6.",
    "SC2 needs the problem as well as the modification, so probe for it here.",
  ],
  tag: "[CFU | Evaluating impact | HITS 7]",
});

const NOTES_YOUDO = T.composeGlanceNotes({
  beats: [
    ["SAY: Voices off. Pick one modification. Sign what it is and what it fixed.",
      "Your partner watches all of it."],
    "SAY: Then they sign back one more barrier somebody designed away.",
    ["CIRCULATE: listen for the problem, not the modification.",
      "The modification is the easy half."],
  ],
  stretch: "name a barrier outside sport and say what would change and what would not.",
  help: "three picture prompts, so the task is choose and explain, not recall.",
  prep: [
    "Different content from the game: numbers there, an explanation here.",
    "Year 5 names one modification. Year 6 explains two and generalises.",
  ],
  tag: "[You Do | Explicit teaching | HITS 10]",
});

const NOTES_EXIT = T.composeGlanceNotes({
  answer: "a starting light, a referee flag, or a visual scoreboard",
  beats: [
    "COLLECT at the door, signed to you, one thing shown with a light or a flag.",
    "Two piles: named one, or could not.",
    "Then the Unit Progress Card, Lesson 5 row, three ticks.",
  ],
  care: "teach every modification as a design problem, never as something overcome.",
  prep: [
    "Assesses SC2. If a student says a Deaf athlete is amazing for competing,",
    "redirect to the result: amazing because of what they did in the pool.",
  ],
  tag: "[Exit ticket | Evaluating impact | HITS 7]",
});

const NOTES_CLOSING = T.composeGlanceNotes({
  beats: [
    "SAY: Turn and tell your partner one barrier somebody designed away.",
    "SAY: Next week you meet two real Australian Deaf athletes.",
  ],
  prep: [
    "Protocol practised today: tapping the shoulder, straight out of the content.",
    "Running late: cut the game to four sports. Never cut the You Do explanation.",
  ],
  tag: "[Closing | Planning | HITS 9]",
});

/* ---------------------------------------------------------------- deck ----- */

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  titleSlide(pres, UNIT_TITLE, "Session 5: Starting lights and flags",
    `Years 5-6 Auslan | Session ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  REPORT.add(signRecallSlide(T, pres, {
    title: "Do Now: what does each one mean?",
    lead: "Sign the meaning back to me. Answers come up on the click.",
    signs: [
      { gloss: "PAST", meaning: "back then" },
      { gloss: "BEFORE", meaning: "before" },
      { gloss: "AFTER", meaning: "after" },
    ],
    notes: NOTES_DONOW, footer: FOOTER,
  }).results);

  liSlide(pres,
    "I am learning to explain how Deaf sport is set up so everything you need to know can be seen.",
    [
      "I can name one thing in sport that is shown with a light or a flag.",
      "I can explain what problem the light or flag solved.",
      "I can give one more example of a barrier somebody designed away.",
    ],
    NOTES_LI, FOOTER);

  // I Do - anchor + the three signs
  REPORT.add(signSlide(T, pres, {
    title: "Made to be seen",
    lead: "Set it up, ask it, watch it, check it. Watch me first, then copy.",
    signs: [
      { gloss: "START", meaning: "start" },
      { gloss: "LIGHT", meaning: "light" },
      { gloss: "FLAG", meaning: "flag" },
    ],
    notes: NOTES_ANCHOR, footer: FOOTER,
  }).results);

  // I Do - the real modifications, built on clicks
  const modSlide = pres.addSlide();
  modSlide.background = { color: C.BG_LIGHT };
  T.addTopBar(modSlide, C.PRIMARY);
  T.addTitle(modSlide, "Four things you can see");
  const MODS = [
    ["A light on the pool", "wired to the starter's gun"],
    ["A flag, or a wave", "how a referee gets your eyes"],
    ["A tap on the shoulder", "how a captain gets the referee"],
    ["A scoreboard", "so nobody waits to be told"],
  ];
  const mw = (9.0 - 0.24) / 2;
  clickBuild(modSlide, MODS.map(([head, sub], i) => () => {
    const mx = 0.5 + (i % 2) * (mw + 0.24);
    const my = 1.62 + Math.floor(i / 2) * 1.62;
    addTextOnShape(modSlide, head, {
      x: mx, y: my, w: mw, h: 0.66, rectRadius: 0.1, fill: { color: C.PRIMARY },
    }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
    addTextOnShape(modSlide, sub, {
      x: mx, y: my + 0.72, w: mw, h: 0.58, rectRadius: 0.1,
      fill: { color: C.BG_CARD }, line: { color: C.SECONDARY, width: 1 },
    }, { fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL });
  }));
  T.addFooter(modSlide, FOOTER);
  modSlide.addNotes(NOTES_MODS);

  // I Do - what actually changed
  const changeSlide = pres.addSlide();
  changeSlide.background = { color: C.BG_LIGHT };
  T.addTopBar(changeSlide, C.SECONDARY);
  T.addTitle(changeSlide, "What changed?");
  REPORT.add(require("./auslan_signs_lib").signCardRow(T, changeSlide, [
    { gloss: "CHANGE", meaning: "change" },
  ], { x: 0.6, w: 3.4, y: 1.5, bottom: 4.9 }));
  addTextOnShape(changeSlide, "Not the athlete", {
    x: 4.5, y: 1.75, w: 5.0, h: 0.75, rectRadius: 0.1, fill: { color: C.ALERT },
  }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
  ["The starting gun changed", "The referee changed", "The scoreboard changed"].forEach((t, i) => {
    addTextOnShape(changeSlide, t, {
      x: 4.5, y: 2.72 + i * 0.66, w: 5.0, h: 0.55, rectRadius: 0.08,
      fill: { color: C.BG_CARD }, line: { color: C.SECONDARY, width: 1 },
    }, { fontSize: 15, fontFace: FONT_B, color: C.CHARCOAL });
  });
  T.addFooter(changeSlide, FOOTER);
  changeSlide.addNotes(NOTES_CHANGE);
  runSlideDiagnostics(changeSlide, pres);

  // We Do - Game 5
  const gameSlide = pres.addSlide();
  gameSlide.background = { color: C.BG_LIGHT };
  T.addTopBar(gameSlide, C.SUCCESS);
  T.addTitle(gameSlide, "Game 5 How Many Medals");
  stepCard(T, gameSlide, {
    title: "Fill in your partner's column",
    steps: [
      "Sheet flat behind the folder",
      "Ask how many medals in each sport",
      "Pencils down while your partner signs",
      "Watch it all, then write the number",
    ],
    x: 0.5, y: 1.5, w: 5.4, h: 2.9, strip: C.SUCCESS,
  });
  addTextOnShape(gameSlide, "Watch, then write", {
    x: 6.1, y: 1.7, w: 3.4, h: 0.8, rectRadius: 0.12, fill: { color: C.ALERT },
  }, { fontSize: 19, fontFace: FONT_H, color: C.WHITE, bold: true });
  gameSlide.addText("A pair who writes while watching will be a number out within three turns.", {
    x: 6.1, y: 2.7, w: 3.4, h: 1.1,
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
    x: 1.1, y: 1.9, w: 7.8, h: 0.85, rectRadius: 0.12, fill: { color: C.ALERT },
  }, { fontSize: 26, fontFace: FONT_H, color: C.WHITE, bold: true });
  checkSlide.addText("What starts a race in Deaf swimming?", {
    x: 1.1, y: 3.0, w: 7.8, h: 0.7,
    fontSize: 26, fontFace: FONT_H, color: C.PRIMARY, bold: true, align: "center", margin: 0,
  });
  checkSlide.addText("Think first. Hands ready. One, two, three.", {
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
  T.addTitle(youSlide, "You Do: what problem did it solve?");
  stepCard(T, youSlide, {
    title: "Explain one, then swap",
    steps: [
      "Pick one modification",
      "Sign what it is and what it fixed",
      "Your partner watches all of it",
      "They sign back one more barrier",
    ],
    x: 0.5, y: 1.5, w: 5.4, h: 2.9, strip: C.SUCCESS,
  });
  addTextOnShape(youSlide, "Voices off", {
    x: 6.1, y: 1.6, w: 3.4, h: 0.9, rectRadius: 0.12, fill: { color: C.ALERT },
  }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
  addTextOnShape(youSlide, "The light replaced the...", {
    x: 6.1, y: 2.75, w: 3.4, h: 0.8, rectRadius: 0.1, fill: { color: C.SECONDARY },
  }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
  youSlide.addText("Finish the sentence", {
    x: 6.1, y: 3.7, w: 3.4, h: 0.4,
    fontSize: 15, fontFace: FONT_B, color: C.CHARCOAL, align: "center", margin: 0,
  });
  T.addFooter(youSlide, FOOTER);
  youSlide.addNotes(NOTES_YOUDO);
  runSlideDiagnostics(youSlide, pres);

  exitTicketSlide(pres,
    ["Sign me one thing in sport that is shown with a light or a flag."],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner one barrier that somebody designed away.",
      scItems: [
        "I can name one thing in sport that is shown with a light or a flag.",
        "I can explain what problem the light or flag solved.",
        "I can give one more example of a barrier somebody designed away.",
      ],
      selfAssessment: "Unit Progress Card, Lesson 5 row",
    },
    NOTES_CLOSING, FOOTER);

  const pptxPath = path.join(OUT_DIR, `Deaf Sports Session ${SESSION} Starting Lights And Flags.pptx`);
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  /* ------------------------------------------------------------- pdf ------ */

  const doc = createPdf({ title: TALLY_RES.name });
  let y = addPdfHeader(doc, TALLY_RES.name, {
    subtitle: "Half the class gets A, half gets B. Do not let them see each other's sheet.",
    color: C.PRIMARY,
    lessonInfo: `Session ${SESSION} | Game 5 How Many Medals | one per student`,
  });

  y = addTipBox(doc,
    "The last row is meant to be ridiculous. Ask who has more sock wrestling medals at the end; it is the bit the class remembers. Pencils go down while the partner signs.",
    y, { color: C.ACCENT });

  [["Sheet A", TALLY_A], ["Sheet B", TALLY_B]].forEach(([label, values], si) => {
    if (si === 1) { doc.addPage(); y = 60; }
    y = addSectionHeading(doc, label, y, { color: C.SECONDARY });
    y = addBodyText(doc,
      "Your medals are already filled in. Write your partner's numbers in the empty column as they tell you.",
      y);

    const colX = [50, 260, 380];
    const rowH = 34;
    doc.save();
    doc.rect(50, y, 495, 30).fill("#" + C.PRIMARY);
    doc.font("Sans-Bold").fontSize(11).fillColor("#" + C.WHITE);
    ["Sport", "My medals", "My partner's medals"].forEach((h, i) => {
      doc.text(h, colX[i] + 10, y + 10, { width: 160 });
    });
    doc.restore();
    y += 30;

    SPORTS.forEach((sport, i) => {
      doc.save();
      doc.rect(50, y, 495, rowH).lineWidth(0.8).strokeColor("#" + C.MUTED).stroke();
      doc.moveTo(260, y).lineTo(260, y + rowH).stroke();
      doc.moveTo(380, y).lineTo(380, y + rowH).stroke();
      doc.font("Sans").fontSize(12).fillColor("#" + C.CHARCOAL);
      doc.text(sport, colX[0] + 10, y + 11, { width: 190 });
      doc.font("Sans-Bold").fontSize(13).fillColor("#" + C.PRIMARY);
      doc.text(String(values[i]), colX[1] + 10, y + 10, { width: 100 });
      doc.restore();
      y += rowH;
    });

    doc.save();
    doc.rect(50, y, 495, rowH).lineWidth(1.2).strokeColor("#" + C.SECONDARY).stroke();
    doc.moveTo(260, y).lineTo(260, y + rowH).stroke();
    doc.moveTo(380, y).lineTo(380, y + rowH).stroke();
    doc.font("Sans-Bold").fontSize(12).fillColor("#" + C.CHARCOAL)
      .text("Total", colX[0] + 10, y + 11, { width: 190 });
    doc.restore();
    y += rowH + 16;

    y = addBodyText(doc, "Who has more sock wrestling medals? ____________________", y);
  });

  addPdfFooter(doc, `Session ${SESSION} | Medal Tally Sheets | Years 5-6 Auslan`);
  await writePdf(doc, path.join(OUT_DIR, TALLY_RES.fileName));
  console.log("PDF written: " + TALLY_RES.fileName);

  REPORT.print(`Session ${SESSION}`);
  console.log(`Session ${SESSION} build complete.`);
}

build().catch((err) => { console.error(err); process.exit(1); });
