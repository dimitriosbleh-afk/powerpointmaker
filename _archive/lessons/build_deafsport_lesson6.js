"use strict";

// Deaf Sports in Australia (Years 5-6 Auslan, Term 4) - Session 6 of 10.
// BUILDING A DEAF PROFILE: four facts, same order, every time.
//   4 production signs: WIN, MEDAL, TEAM, COMPETE. Game 6 Which Card Is Mine.
//   Assessment piece 1 (observational checklist) runs live during the game.
// Two profile cards are real people, one of them living. Everything on them
// comes from official ICSD athlete records; keep every profile to what they
// did in their sport. Sources are in the unit document appendix.
// WIN uses WIN_3.jpg: the default WIN.jpg is CELEBRATION, not winning.

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

const SESSION = 6;
const TOTAL = 10;
const UNIT_TITLE = "Deaf Sports in Australia";
const ANCHOR = "Set it up, ask it, watch it, check it.";
const FOOTER = `${UNIT_TITLE} | Session ${SESSION} of ${TOTAL} | Years 5-6 Auslan`;
const OUT_DIR = "output/DeafSport_Lesson6_Building_A_Profile";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const PROFILE_RES = makeSessionResource(SESSION,
  "Deaf Athlete Profile Cards",
  "Six cards, four facts each. Cut and laminate; reused in Sessions 7 and 10.");
const CHECKLIST_RES = makeSessionResource(SESSION,
  "Observational Checklist",
  "Assessment piece 1. One page per class, ticked one-handed during the game.");
const RESOURCE_ITEMS = [PROFILE_RES, CHECKLIST_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

const REPORT = createSignReport();

// Four slots, same order on every card. Sources: ICSD athlete records and
// Deaf Sports Australia. Nothing here touches health, hearing or family.
const PROFILES = [
  ["Barry Knapman, Australia", "Diving", "1965 and 1969",
    "Gold on the three metre springboard in Washington in 1965. Silver in the same event in Belgrade in 1969."],
  ["Cindy-Lu Bailey, Australia", "Swimming", "1977 to 1997",
    "Competed at six world games for the Deaf, from Bucharest to Copenhagen. Won 29 medals, 19 of them gold."],
  ["The Melbourne Deaf Cricket Club", "Cricket", "Started in the 1880s",
    "Still playing today. One of the oldest Deaf sport clubs in the world."],
  ["The Australian Deaf Games", "Many sports", "First held in Sydney, summer 1964 and 1965",
    "Held every three or four years since. One of the oldest ongoing Deaf sporting events in the world."],
  ["The Melbourne Deaflympics", "Many sports", "2005",
    "Over 3,500 Deaf athletes and officials came to Melbourne. Australia sent its largest ever team."],
  ["Deaf Sports Australia", "All sports", "Set up in 1954",
    "The national body for Deaf sport in Australia. Helps Deaf Australians play sport at every level."],
];

const CRITERIA = [
  "Waits for eyes before starting",
  "Gives four facts in order",
  "Keeps going after losing a fact",
  "Watches a whole profile",
  "Repairs with again or what mean",
  "Answers with the card face down",
];

/* ---------------------------------------------------------------- notes ---- */

const NOTES_TITLE = "Session 6 of 10. Assessment piece 1 runs live during the game. Clipboard and checklist ready before the lesson starts.";

const NOTES_RESOURCES = T.composeGlanceNotes({
  beats: [
    ["SHOW while students settle.", "SAY: Six cards per table, face up in the middle."],
    "SET UP: clipboard and checklist in your hand before the game starts.",
  ],
  prep: [
    "Print and cut the profile cards once; Sessions 7 and 10 reuse them.",
    "One checklist per class. Tick during the game, never stop it to assess.",
    ATTRIBUTION,
  ],
  tag: "[Setup | Planning | HITS 2]",
});

const NOTES_DONOW = T.composeGlanceNotes({
  answer: "start, light, flag",
  beats: [
    ["ASK: What does each sign mean?",
      "10 sec think. Cue: Everyone signs it to me on three... one, two, three.",
      "EXPECT: START, LIGHT, FLAG"],
    "REVEAL after the room has answered, one click.",
    ["SAY: You have facts about Deaf sport and you have four questions.",
      "Today you put them together into thirty seconds you can give somebody."],
  ],
  prep: [
    "Retrieval from Lesson 5, plus a number between 10 and 40 on boards.",
    "The bridge names what a profile is before the first one is modelled.",
  ],
  tag: "[Do Now | Attention, focus and regulation | HITS 6]",
});

const NOTES_LI = T.composeGlanceNotes({
  beats: [
    ["POINT to each criterion.", "SAY: Four facts, one order, and it never changes."],
    "SAY: By the end everyone can give two facts about somebody on a card.",
  ],
  prep: [
    "SC1 is two facts with the card visible, reachable by every student.",
    "SC2 is what the exit ticket collects. SC3 is the card face down.",
  ],
  tag: "[LI and SC | Planning | HITS 1]",
});

const NOTES_SIGNS = T.composeGlanceNotes({
  answer: "WIN, MEDAL, TEAM, COMPETE",
  beats: [
    ["SAY: Same four moves. Set it up, ask it, watch it, check it.",
      "A profile is those four moves with the asking already done."],
    "MODEL each sign six times, then inside a sentence about a real athlete.",
    ["SAY: Every profile has four facts, in this order.",
      "Who. What sport. When. What they won."],
  ],
  trap: ["reciting the four facts as a list rather than using them as slots.",
    "Fix: put the four headings up, fill two, pairs supply the rest."],
  prep: [
    "WIN uses WIN_3: the default Signbank entry for win is CELEBRATION.",
    "MEDAL and COMPETE are tagged CHECK SIGNBANK in the vocabulary bank.",
    "Confirm all four from the school reference before teaching.",
  ],
  tag: "[I Do | Explicit teaching | HITS 3, 4]",
});

const NOTES_MODEL = T.composeGlanceNotes({
  answer: "who, what sport, when, what they won",
  beats: [
    ["SAY: Barry Knapman. Australia. Diving.",
      "He won gold on the three metre springboard in Washington in 1965.",
      "Four years later in Belgrade he won silver in the same event."],
    ["SAY: Cindy-Lu Bailey. Australia. Swimming.",
      "She competed at six of these games, from 1977 through to 1997.",
      "She won twenty nine medals. Nineteen of them were gold."],
    ["SAY: If you lose a fact, keep going and come back.",
      "Do not stop and restart; your partner is holding what you already gave."],
  ],
  prep: [
    "Deliver both at pace, without stopping. The pace is part of the model.",
    "Both are real people, one living. Keep to what they did in their sport.",
  ],
  tag: "[I Do | Explicit teaching | HITS 4]",
});

const NOTES_WRONG = T.composeGlanceNotes({
  beats: [
    ["GET IT WRONG: deliver a profile starting with the medals.",
      "SAY: I gave you twenty nine medals and you had no idea whose."],
    ["SAY: The setting up is not decoration.",
      "It is the thing that makes everything after it land."],
    "MODEL it correctly, name first, without commentary this time.",
  ],
  trap: ["opening with the most exciting fact instead of the name.",
    "Fix: student redoes the profile, name first, same card."],
  stretch: "build two profiles and say which person you would rather interview.",
  help: "the four slot headings on a strip, card face up, so recall is not the load.",
  prep: [
    "The deliberate error targets exactly what the check will test.",
    "Year 5 gives three facts card visible. Year 6 gives four, card face down.",
  ],
  tag: "[I Do | Explicit teaching | HITS 3]",
});

const NOTES_GAME = T.composeGlanceNotes({
  answer: "the card the describer secretly chose",
  beats: [
    ["SET UP: groups of four, all six cards face up in the middle.",
      "Each student secretly picks one and remembers it."],
    "SAY: Give your four facts without naming the card. The group points at the end.",
    ["CIRCULATE with the checklist. Tick one-handed as you go.",
      "Do not stop the game to assess."],
    "COLLECT: the describer says whether the group got it right, then the next player.",
  ],
  trap: ["the group scanning cards mid-description and missing the language.",
    "Fix: cards face down until the description finishes, then turn them up."],
  prep: [
    "Game 6 Which Card Is Mine. Assessment piece 1 runs here, live.",
    "Six criteria, plain language, all visible from two metres away.",
  ],
  tag: "[We Do | Collaborative learning | HITS 5, 8]",
});

const NOTES_CHECK = T.composeGlanceNotes({
  answer: "who it is",
  beats: [
    ["ASK: What is the first fact in a profile?",
      "5 sec think. Cue: Everyone signs it to me on three... one, two, three.",
      "EXPECT: WHO"],
    ["SCAN the room, back row first.",
      "80%+ -> go to You Do and let them build their own.",
      "Less -> four headings up, fill the first and last, pairs supply the middle, re-ask."],
    ["FOLLOW UP one student.",
      "SAY: You started with the name. Say why that has to come first."],
  ],
  trap: ["the four slots blurring into a recited list.",
    "Fix: partial model on the board, students fill the gaps, then redo."],
  prep: [
    "The decision point that changes Lesson 7.",
    "If it fails, Lesson 7 spends its first fifteen minutes rebuilding the profile.",
  ],
  tag: "[CFU | Evaluating impact | HITS 7]",
});

const NOTES_YOUDO = T.composeGlanceNotes({
  beats: [
    ["SAY: Voices off. Pick a card your group did not use.",
      "Build the four facts and sign the profile to your partner."],
    "SAY: Card face down, then your partner asks you one question about it.",
    ["CIRCULATE: the question is the point.",
      "A profile that survives a question is a profile they own."],
  ],
  stretch: "build a second profile and say which of the two you would interview.",
  help: "the four slot headings on a strip, card left face up.",
  prep: [
    "Different content from the game: guessing a card there, surviving a question here.",
    "SC3 only shows up here, so protect this stage if time is short.",
  ],
  tag: "[You Do | Explicit teaching | HITS 10]",
});

const NOTES_EXIT = T.composeGlanceNotes({
  answer: "any two of the four facts, in their right slots",
  beats: [
    "COLLECT at the door, signed to you, two facts about their card.",
    "Two piles: two facts in the right slots, or fewer.",
    "Then the Unit Progress Card, Lesson 6 row, three ticks.",
  ],
  care: "two cards are real people, one living. Keep it to what they did in sport.",
  prep: [
    "Assesses SC2. Nothing about health, hearing or family, even if a student",
    "finds it somewhere. Redirect in one line: we are asking what they did.",
  ],
  tag: "[Exit ticket | Evaluating impact | HITS 7]",
});

const NOTES_CLOSING = T.composeGlanceNotes({
  beats: [
    "SAY: Turn and tell your partner which person you would most like to meet.",
    "SAY: Next week you interview each other, and I record it.",
  ],
  prep: [
    "Protocol practised today: choosing again or what mean, whichever fits.",
    "File the checklist today. Assessment piece 2 is next week.",
  ],
  tag: "[Closing | Planning | HITS 9]",
});

/* ---------------------------------------------------------------- deck ----- */

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  titleSlide(pres, UNIT_TITLE, "Session 6: Building a Deaf profile",
    `Years 5-6 Auslan | Session ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  REPORT.add(signRecallSlide(T, pres, {
    title: "Do Now: what does each one mean?",
    lead: "Sign the meaning back to me. Answers come up on the click.",
    signs: [
      { gloss: "START", meaning: "start" },
      { gloss: "LIGHT", meaning: "light" },
      { gloss: "FLAG", meaning: "flag" },
    ],
    notes: NOTES_DONOW, footer: FOOTER,
  }).results);

  liSlide(pres,
    "I am learning to build a short profile of a Deaf Australian athlete from four facts.",
    [
      "I can give two facts about a person on a card.",
      "I can build and sign a four-fact profile.",
      "I can answer a question about my profile without reading the card.",
    ],
    NOTES_LI, FOOTER);

  // I Do - the four signs
  REPORT.add(signSlide(T, pres, {
    title: "Four signs for a profile",
    lead: "Set it up, ask it, watch it, check it. Watch me first, then copy.",
    signs: [
      { gloss: "TEAM", meaning: "team" },
      { gloss: "COMPETE", meaning: "compete" },
      { gloss: "WIN", meaning: "win" },
      { gloss: "MEDAL", meaning: "medal" },
    ],
    notes: NOTES_SIGNS, footer: FOOTER,
  }).results);

  // I Do - the four slots, built on clicks
  const slotSlide = pres.addSlide();
  slotSlide.background = { color: C.BG_LIGHT };
  T.addTopBar(slotSlide, C.PRIMARY);
  T.addTitle(slotSlide, "Four facts, always this order");
  const SLOTS = [
    ["1. Who", "Barry Knapman, Australia"],
    ["2. What sport", "Diving"],
    ["3. When", "1965 and 1969"],
    ["4. What they won", "Gold in Washington, silver in Belgrade"],
  ];
  clickBuild(slotSlide, SLOTS.map(([head, eg], i) => () => {
    const sy = 1.66 + i * 0.85;
    addTextOnShape(slotSlide, head, {
      x: 0.7, y: sy, w: 2.7, h: 0.68, rectRadius: 0.1, fill: { color: C.PRIMARY },
    }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });
    addTextOnShape(slotSlide, eg, {
      x: 3.6, y: sy, w: 5.7, h: 0.68, rectRadius: 0.1,
      fill: { color: C.BG_CARD }, line: { color: C.SECONDARY, width: 1 },
    }, { fontSize: 15, fontFace: FONT_B, color: C.CHARCOAL });
  }));
  T.addFooter(slotSlide, FOOTER);
  slotSlide.addNotes(NOTES_MODEL);

  // I Do - the deliberate error
  const wrongSlide = pres.addSlide();
  wrongSlide.background = { color: C.BG_LIGHT };
  T.addTopBar(wrongSlide, C.ALERT);
  T.addTitle(wrongSlide, "Watch me get it wrong");
  addTextOnShape(wrongSlide, "Twenty nine medals!", {
    x: 1.3, y: 1.75, w: 7.4, h: 0.85, rectRadius: 0.12, fill: { color: C.ALERT },
  }, { fontSize: 26, fontFace: FONT_H, color: C.WHITE, bold: true });
  wrongSlide.addText("...and you had no idea whose.", {
    x: 1.3, y: 2.8, w: 7.4, h: 0.6,
    fontSize: 22, fontFace: FONT_H, color: C.PRIMARY, bold: true, align: "center", margin: 0,
  });
  addTextOnShape(wrongSlide, "Name first. Every time.", {
    x: 2.6, y: 3.62, w: 4.8, h: 0.7, rectRadius: 0.1, fill: { color: C.SECONDARY },
  }, { fontSize: 19, fontFace: FONT_H, color: C.WHITE, bold: true });
  T.addFooter(wrongSlide, FOOTER);
  wrongSlide.addNotes(NOTES_WRONG);
  runSlideDiagnostics(wrongSlide, pres);

  // We Do - Game 6
  const gameSlide = pres.addSlide();
  gameSlide.background = { color: C.BG_LIGHT };
  T.addTopBar(gameSlide, C.SUCCESS);
  T.addTitle(gameSlide, "Game 6 Which Card Is Mine");
  stepCard(T, gameSlide, {
    title: "Give the facts, not the name",
    steps: [
      "All six cards face up in the middle",
      "Secretly pick one and remember it",
      "Sign your four facts, never the name",
      "The group points at the end, not during",
    ],
    x: 0.5, y: 1.5, w: 5.4, h: 2.9, strip: C.SUCCESS,
  });
  addTextOnShape(gameSlide, "Watch it all first", {
    x: 6.1, y: 1.7, w: 3.4, h: 0.8, rectRadius: 0.12, fill: { color: C.ALERT },
  }, { fontSize: 19, fontFace: FONT_H, color: C.WHITE, bold: true });
  gameSlide.addText("A student scanning cards mid-description has stopped watching the language.", {
    x: 6.1, y: 2.7, w: 3.4, h: 1.2,
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
  checkSlide.addText("What is the first fact in a profile?", {
    x: 1.1, y: 3.0, w: 7.8, h: 0.7,
    fontSize: 27, fontFace: FONT_H, color: C.PRIMARY, bold: true, align: "center", margin: 0,
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
  T.addTitle(youSlide, "You Do: build your own profile");
  stepCard(T, youSlide, {
    title: "Then survive a question",
    steps: [
      "Pick a card your group did not use",
      "Build the four facts in order",
      "Sign the profile to your partner",
      "Card face down. They ask you one question",
    ],
    x: 0.5, y: 1.5, w: 5.4, h: 2.9, strip: C.SUCCESS,
  });
  addTextOnShape(youSlide, "Voices off", {
    x: 6.1, y: 1.6, w: 3.4, h: 0.9, rectRadius: 0.12, fill: { color: C.ALERT },
  }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
  addTextOnShape(youSlide, "Card face down", {
    x: 6.1, y: 2.75, w: 3.4, h: 0.8, rectRadius: 0.1, fill: { color: C.SECONDARY },
  }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
  youSlide.addText("for the question", {
    x: 6.1, y: 3.7, w: 3.4, h: 0.4,
    fontSize: 15, fontFace: FONT_B, color: C.CHARCOAL, align: "center", margin: 0,
  });
  T.addFooter(youSlide, FOOTER);
  youSlide.addNotes(NOTES_YOUDO);
  runSlideDiagnostics(youSlide, pres);

  exitTicketSlide(pres,
    ["Sign me two facts about the person or event on your card."],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner which person you would most like to meet, and why.",
      scItems: [
        "I can give two facts about a person on a card.",
        "I can build and sign a four-fact profile.",
        "I can answer a question about my profile without reading the card.",
      ],
      selfAssessment: "Unit Progress Card, Lesson 6 row",
    },
    NOTES_CLOSING, FOOTER);

  const pptxPath = path.join(OUT_DIR, `Deaf Sports Session ${SESSION} Building A Profile.pptx`);
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  /* ------------------------------------------------------- pdf: cards ----- */

  const doc = createPdf({ title: PROFILE_RES.name });
  let y = addPdfHeader(doc, PROFILE_RES.name, {
    subtitle: "Six cards, four facts each, always in the same order.",
    color: C.PRIMARY,
    lessonInfo: `Session ${SESSION} | Game 6 Which Card Is Mine | one set per group of four`,
  });

  y = addTipBox(doc,
    "Cards 1 and 2 are real people, one of them living. Everything on them comes from official ICSD athlete records and Deaf Sports Australia; check them against the unit document's Sources appendix before printing. Keep every profile to what they did in their sport.",
    y, { color: C.ACCENT });

  PROFILES.forEach(([who, sport, when, what]) => {
    if (y + 132 > 780) { doc.addPage(); y = 60; }
    doc.save();
    doc.rect(50, y, 495, 124).lineWidth(0.8).dash(3, { space: 2 })
      .strokeColor("#" + C.MUTED).stroke();
    doc.undash();
    doc.font("Sans-Bold").fontSize(15).fillColor("#" + C.PRIMARY)
      .text(who, 62, y + 12, { width: 470 });
    const rows = [["Sport", sport], ["When", when], ["What they won or did", what]];
    let ry = y + 38;
    rows.forEach(([label, value]) => {
      doc.font("Sans-Bold").fontSize(9.5).fillColor("#" + C.SECONDARY)
        .text(label, 62, ry, { width: 130 });
      doc.font("Sans").fontSize(10.5).fillColor("#" + C.CHARCOAL)
        .text(value, 196, ry - 1, { width: 336 });
      ry += Math.max(18, doc.heightOfString(value, { width: 336 }) + 6);
    });
    doc.restore();
    y += 132;
  });

  addPdfFooter(doc, `Session ${SESSION} | Deaf Athlete Profile Cards | Years 5-6 Auslan`);
  await writePdf(doc, path.join(OUT_DIR, PROFILE_RES.fileName));
  console.log("PDF written: " + PROFILE_RES.fileName);

  /* --------------------------------------------------- pdf: checklist ----- */

  const cdoc = createPdf({ title: CHECKLIST_RES.name, layout: "landscape" });
  let cy = addPdfHeader(cdoc, CHECKLIST_RES.name, {
    subtitle: "Assessment piece 1. Tick one-handed during Game 6. Do not stop the game.",
    color: C.PRIMARY,
    lessonInfo: `Session ${SESSION} | one page per class | S = supported`,
  });

  cy = addSectionHeading(cdoc, "Class: ______________________  Date: ____________", cy,
    { color: C.SECONDARY });

  const nameW = 120;
  const critW = 62;
  const rowH = 21;
  cdoc.save();
  cdoc.rect(50, cy, nameW + critW * 6, 46).fill("#" + C.PRIMARY);
  cdoc.font("Sans-Bold").fontSize(8).fillColor("#" + C.WHITE);
  cdoc.text("Student", 54, cy + 18, { width: nameW - 8 });
  CRITERIA.forEach((crit, i) => {
    cdoc.text(crit, 50 + nameW + i * critW + 4, cy + 8, { width: critW - 8, align: "center" });
  });
  cdoc.restore();
  cy += 46;

  for (let r = 0; r < 26; r += 1) {
    if (cy + rowH > 800) { cdoc.addPage(); cy = 60; }
    cdoc.save();
    cdoc.rect(50, cy, nameW + critW * 6, rowH).lineWidth(0.6)
      .strokeColor("#" + C.MUTED).stroke();
    for (let i = 0; i <= 6; i += 1) {
      const lx = 50 + nameW + i * critW;
      if (i < 6) cdoc.moveTo(lx, cy).lineTo(lx, cy + rowH).stroke();
    }
    cdoc.moveTo(50 + nameW, cy).lineTo(50 + nameW, cy + rowH).stroke();
    cdoc.restore();
    cy += rowH;
  }

  cy += 12;
  addTipBox(cdoc,
    "Not yet: fewer than two facts, or starts before the group is watching. Approaching: two or three facts, card visible, stops when one is lost. At standard: all four in order, keeps going after a breakdown. Beyond: all four card face down, repairs unprompted, answers a question on top.",
    cy, { color: C.SECONDARY });

  addPdfFooter(cdoc, `Session ${SESSION} | Observational Checklist | Years 5-6 Auslan`);
  await writePdf(cdoc, path.join(OUT_DIR, CHECKLIST_RES.fileName));
  console.log("PDF written: " + CHECKLIST_RES.fileName);

  REPORT.print(`Session ${SESSION}`);
  console.log(`Session ${SESSION} build complete.`);
}

build().catch((err) => { console.error(err); process.exit(1); });
