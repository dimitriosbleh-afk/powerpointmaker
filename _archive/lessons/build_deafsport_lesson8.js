"use strict";

// Deaf Sports in Australia (Years 5-6 Auslan, Term 4) - Session 8 of 10.
// SHOWING THE STORY: set the scene before anything happens in it.
//   4 production signs: STORY, TELL, FIRST, HAPPEN. Game 8 Watch And Draw It.
//   Assessment piece 3 (hard copy diagnostic) runs here, and runs here on
//   branch B too, in the final slot of the term. It is the rejoin point.
// Depicting signs, constructed action and the use of signing space are modelled
// from the school reference. This deck never describes how a sign is produced.
// Stage minutes differ: We Do 10, exit ticket and diagnostic 10.

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

const SESSION = 8;
const TOTAL = 10;
const UNIT_TITLE = "Deaf Sports in Australia";
const ANCHOR = "Set it up, ask it, watch it, check it.";
const FOOTER = `${UNIT_TITLE} | Session ${SESSION} of ${TOTAL} | Years 5-6 Auslan`;
const OUT_DIR = "output/DeafSport_Lesson8_Showing_The_Story";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const STORY_RES = makeSessionResource(SESSION,
  "Story Strip One The Club That Kept Going",
  "One per pair. Four events, two people. Reused in Session 10 presentations.");
const DIAG_RES = makeSessionResource(SESSION,
  "Hard Copy Diagnostic",
  "Assessment piece 3. Eight items, about eight minutes, markable at a glance.");
const RESOURCE_ITEMS = [STORY_RES, DIAG_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

const REPORT = createSignReport();

const STORY_ONE = [
  "It is Melbourne, a long time ago. Two friends stand at the edge of a cricket ground. One of them is holding a bat.",
  "Nobody will let them join a team, because the umpire calls out and they cannot hear him.",
  "So they start their own club. They agree on hand signals for out, for over, and for wait.",
  "More than a hundred years later, the club is still playing.",
];

/* ---------------------------------------------------------------- notes ---- */

const NOTES_TITLE = "Session 8 of 10. Assessment piece 3 runs today. On branch B this lesson is the rejoin point and runs in the final slot of the term.";

const NOTES_RESOURCES = T.composeGlanceNotes({
  beats: [
    ["SHOW while students settle.", "SAY: One story per pair, and a sheet each at the end."],
    "SET UP: plain paper and pencils out for the game, pencils down to start.",
  ],
  prep: [
    "Print the story strip once and laminate; Session 10 reuses it.",
    "Print the diagnostic per student. It is collected and kept as evidence.",
    ATTRIBUTION,
  ],
  tag: "[Setup | Planning | HITS 2]",
});

const NOTES_DONOW = T.composeGlanceNotes({
  answer: "back then, before, then",
  beats: [
    ["ASK: What does each sign mean?",
      "10 sec think. Cue: Everyone signs it to me on three... one, two, three.",
      "EXPECT: PAST, BEFORE, THEN"],
    "REVEAL after the room has answered, one click.",
    ["SAY: You can list events in order. A story is not a list.",
      "A story has a place and people in it, and today you put them there first."],
  ],
  prep: [
    "Retrieval reaches back to Lesson 4, the last time the time signs were taught.",
    "The bridge names the difference between a list and a story.",
  ],
  tag: "[Do Now | Attention, focus and regulation | HITS 6]",
});

const NOTES_LI = T.composeGlanceNotes({
  beats: [
    ["POINT to each criterion.", "SAY: One idea today. Show me where it happens before it happens."],
    "SAY: By the end everyone can show where a story happens.",
  ],
  prep: [
    "SC1 is showing the place, reachable by every student.",
    "SC2 is what the exit ticket collects. SC3 is keeping two people apart.",
  ],
  tag: "[LI and SC | Planning | HITS 1]",
});

const NOTES_SIGNS = T.composeGlanceNotes({
  answer: "STORY, TELL, FIRST, HAPPEN",
  beats: [
    ["SAY: Same four moves. Set it up, ask it, watch it, check it.",
      "In an interview you set the topic. In a story you set the place and the people."],
    ["SAY: You do it before anything happens to them.",
      "Watching and checking do not change at all."],
    "MODEL each sign six times, then inside the opening of the story.",
  ],
  trap: ["starting the events before the scene is set.",
    "Fix: stop them, set the place, then let the story start again."],
  prep: [
    "CHECK GRAMMAR: depicting signs, constructed action and signing space are",
    "modelled from the school reference. Do not improvise them in front of them.",
    "TELL carries telling and rendition in Signbank, which is the narrative sense.",
  ],
  tag: "[I Do | Explicit teaching | HITS 3, 4]",
});

const NOTES_MODEL = T.composeGlanceNotes({
  beats: [
    ["MODEL the scene first: Melbourne, a cricket ground, a long time ago, two people.",
      "Nothing happens yet."],
    ["SAY: Watch where I put each person. I keep them there for the whole story.",
      "The second I move them, you will not know who is doing what."],
    ["SAY: When you watch a story, the place tells you as much as the hands do.",
      "Where somebody is standing is information."],
    "MODEL the four events in order, keeping both people where you put them.",
  ],
  trap: ["both people ending up in the same space.",
    "Fix: two chairs at the front, one person each, retell pointing at chairs."],
  prep: [
    "Set the scene fully before the first event. That is the whole model.",
    "Confirm the depicting and constructed action forms before the lesson.",
  ],
  tag: "[I Do | Explicit teaching | HITS 4]",
});

const NOTES_WRONG = T.composeGlanceNotes({
  beats: [
    ["GET IT WRONG: retell two events with both people in the same place.",
      "SAY: I just told you a story where you could not tell who did anything."],
    ["SAY: Nothing wrong with my signs.",
      "Everything wrong with where I put them."],
    "MODEL it correctly, both people apart, without commentary this time.",
  ],
  care: "constructed action is taking on what somebody did, not doing an impression.",
  prep: [
    "If a retell tips into caricature, stop it once and name what changed.",
    "You moved from showing what he did to making him funny. Then run it straight.",
  ],
  tag: "[I Do | Explicit teaching | HITS 3]",
});

const NOTES_GAME = T.composeGlanceNotes({
  answer: "a sketch with the two people in the places the signer put them",
  beats: [
    ["SET UP: pairs facing, paper flat, pencils down.",
      "Student A signs a scene only: a place and two people."],
    "SAY: Watch the whole thing with pencils down, then draw it.",
    ["CIRCULATE: a student head down drawing has missed the placement.",
      "That is the whole content of this game."],
    "COLLECT: compare. If the people are in the wrong places, A signs it again.",
  ],
  trap: ["drawing while the partner is still signing.",
    "Fix: pencils on the desk, partner signs the scene again, then draw."],
  prep: [
    "Game 8 Watch And Draw It. Ten minutes only; the diagnostic needs its slot.",
    "The scene, not the events. Events come in the You Do.",
  ],
  tag: "[We Do | Collaborative learning | HITS 5]",
});

const NOTES_CHECK = T.composeGlanceNotes({
  answer: "two people, kept in two different places",
  beats: [
    ["ASK: Show me where the two people in the story were.",
      "5 sec think. Cue: Everyone signs it to me on three... one, two, three.",
      "EXPECT: two distinct places"],
    ["SCAN the room, back row first.",
      "80%+ -> go to You Do with the full four-event retell.",
      "Less -> two chairs at the front, one person each, retell pointing, then re-ask."],
    ["FOLLOW UP one student.",
      "SAY: Do you agree with where he put her? Add one thing."],
  ],
  trap: ["both people landing in the same space because handshapes are being copied.",
    "Fix: physical markers first, then the same retell without them."],
  prep: [
    "The decision point that changes Lesson 9.",
    "Translation needs placement to hold, so pivot rather than push on.",
  ],
  tag: "[CFU | Evaluating impact | HITS 7]",
});

const NOTES_YOUDO = T.composeGlanceNotes({
  beats: [
    ["SAY: Voices off. Retell the whole story to your partner.",
      "Scene first, then the four events, both people kept apart."],
    "SAY: Your partner watches it all, then signs back which event came first.",
    ["CIRCULATE: check the scene arrives before the first event.",
      "That is the criterion, not the number of events."],
  ],
  stretch: "retell the story from the second person's point of view.",
  help: "the story as four pictures, with the two people drawn in different places.",
  prep: [
    "Different content from the game: the scene there, scene plus events here.",
    "Year 5 sets the scene and retells two events. Year 6 retells all four.",
  ],
  tag: "[You Do | Explicit teaching | HITS 10]",
});

const NOTES_DIAG = T.composeGlanceNotes({
  answer: "eight items covering the whole unit, not just today",
  beats: [
    "SET UP: one sheet each, silent, independent. About eight minutes.",
    "SAY: This covers everything since week one, not just today.",
    ["COLLECT the sheets. Then the exit ticket at the door:",
      "sign me one event from the story."],
  ],
  prep: [
    "Assessment piece 3. Items 1 to 4 are right or wrong on sight.",
    "Items 5 to 8 are one line each. Markable at a glance, about a minute a sheet.",
    "On branch B this runs in the final slot of the term instead.",
  ],
  tag: "[Exit ticket | Evaluating impact | HITS 7]",
});

const NOTES_CLOSING = T.composeGlanceNotes({
  beats: [
    "SAY: Turn and tell your partner where your story happened.",
    "SAY: Next week the story comes off a piece of paper instead of from me.",
  ],
  prep: [
    "Protocol practised today: signing space and sight lines, pairs at an angle.",
    "Running late: cut the game to one round. Never cut the diagnostic.",
  ],
  tag: "[Closing | Planning | HITS 9]",
});

/* ---------------------------------------------------------------- deck ----- */

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  titleSlide(pres, UNIT_TITLE, "Session 8: Showing the story",
    `Years 5-6 Auslan | Session ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  REPORT.add(signRecallSlide(T, pres, {
    title: "Do Now: what does each one mean?",
    lead: "Sign the meaning back to me. Answers come up on the click.",
    signs: [
      { gloss: "PAST", meaning: "back then" },
      { gloss: "BEFORE", meaning: "before" },
      { gloss: "THEN", meaning: "then" },
    ],
    notes: NOTES_DONOW, footer: FOOTER,
  }).results);

  liSlide(pres,
    "I am learning to set a scene before I tell what happened in it.",
    [
      "I can show where a story happens before I start it.",
      "I can retell a four-event story with the scene set first.",
      "I can keep two people apart so my partner knows who is who.",
    ],
    NOTES_LI, FOOTER);

  // I Do - the four signs
  REPORT.add(signSlide(T, pres, {
    title: "Signs for telling a story",
    lead: "Set it up, ask it, watch it, check it. Watch me first, then copy.",
    signs: [
      { gloss: "STORY", meaning: "story" },
      { gloss: "TELL", meaning: "tell" },
      { gloss: "FIRST", meaning: "first" },
      { gloss: "HAPPEN", meaning: "happen" },
    ],
    notes: NOTES_SIGNS, footer: FOOTER,
  }).results);

  // I Do - the scene, then the events, on clicks
  const sceneSlide = pres.addSlide();
  sceneSlide.background = { color: C.BG_LIGHT };
  T.addTopBar(sceneSlide, C.PRIMARY);
  T.addTitle(sceneSlide, "Set the scene before anything happens");
  addTextOnShape(sceneSlide, "Where? Melbourne, a cricket ground, a long time ago.", {
    x: 0.6, y: 1.55, w: 8.8, h: 0.68, rectRadius: 0.1, fill: { color: C.PRIMARY },
  }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });
  addTextOnShape(sceneSlide, "Who? Two friends. Put them in two different places.", {
    x: 0.6, y: 2.35, w: 8.8, h: 0.68, rectRadius: 0.1, fill: { color: C.SECONDARY },
  }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });
  clickBuild(sceneSlide, [
    () => {
      addTextOnShape(sceneSlide, "Now, and only now, the events", {
        x: 2.1, y: 3.3, w: 5.8, h: 0.62, rectRadius: 0.1, fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
      sceneSlide.addText("Where somebody is standing is information.", {
        x: 1.0, y: 4.05, w: 8.0, h: 0.5,
        fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL, align: "center", margin: 0,
      });
    },
  ]);
  T.addFooter(sceneSlide, FOOTER);
  sceneSlide.addNotes(NOTES_MODEL);

  // I Do - the deliberate error
  const wrongSlide = pres.addSlide();
  wrongSlide.background = { color: C.BG_LIGHT };
  T.addTopBar(wrongSlide, C.ALERT);
  T.addTitle(wrongSlide, "Watch me get it wrong");
  addTextOnShape(wrongSlide, "Both people in the same place", {
    x: 1.3, y: 1.85, w: 7.4, h: 0.85, rectRadius: 0.12, fill: { color: C.ALERT },
  }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
  wrongSlide.addText("You could not tell who did anything.", {
    x: 1.3, y: 2.9, w: 7.4, h: 0.6,
    fontSize: 21, fontFace: FONT_H, color: C.PRIMARY, bold: true, align: "center", margin: 0,
  });
  addTextOnShape(wrongSlide, "Nothing wrong with the signs. Everything wrong with the places.", {
    x: 1.0, y: 3.72, w: 8.0, h: 0.7, rectRadius: 0.1, fill: { color: C.SECONDARY },
  }, { fontSize: 16, fontFace: FONT_B, color: C.WHITE, bold: true });
  T.addFooter(wrongSlide, FOOTER);
  wrongSlide.addNotes(NOTES_WRONG);
  runSlideDiagnostics(wrongSlide, pres);

  // We Do - Game 8
  const gameSlide = pres.addSlide();
  gameSlide.background = { color: C.BG_LIGHT };
  T.addTopBar(gameSlide, C.SUCCESS);
  T.addTitle(gameSlide, "Game 8 Watch And Draw It");
  stepCard(T, gameSlide, {
    title: "Sign the scene, not the story",
    steps: [
      "A place and two people. Nothing happens yet",
      "Partner watches with pencils down",
      "Then they draw the scene",
      "Wrong places? Sign it again, do not explain",
    ],
    x: 0.5, y: 1.5, w: 5.4, h: 2.9, strip: C.SUCCESS,
  });
  addTextOnShape(gameSlide, "Pencils down", {
    x: 6.1, y: 1.7, w: 3.4, h: 0.85, rectRadius: 0.12, fill: { color: C.ALERT },
  }, { fontSize: 21, fontFace: FONT_H, color: C.WHITE, bold: true });
  gameSlide.addText("A student head down drawing has missed the placement, which is the whole content.", {
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
  addTextOnShape(checkSlide, "Everyone signs it to me on three", {
    x: 1.1, y: 1.9, w: 7.8, h: 0.85, rectRadius: 0.12, fill: { color: C.ALERT },
  }, { fontSize: 26, fontFace: FONT_H, color: C.WHITE, bold: true });
  checkSlide.addText("Show me where the two people were", {
    x: 1.1, y: 3.0, w: 7.8, h: 0.7,
    fontSize: 27, fontFace: FONT_H, color: C.PRIMARY, bold: true, align: "center", margin: 0,
  });
  checkSlide.addText("Two people. Two different places.", {
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
  T.addTitle(youSlide, "You Do: retell the whole story");
  stepCard(T, youSlide, {
    title: "Scene first, then four events",
    steps: [
      "Show where it happens",
      "Put the two people in two places",
      "Tell the four events in order",
      "Partner signs back which came first",
    ],
    x: 0.5, y: 1.5, w: 5.4, h: 2.9, strip: C.SUCCESS,
  });
  addTextOnShape(youSlide, "Voices off", {
    x: 6.1, y: 1.6, w: 3.4, h: 0.9, rectRadius: 0.12, fill: { color: C.ALERT },
  }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
  addTextOnShape(youSlide, "Keep them where you put them", {
    x: 6.1, y: 2.75, w: 3.4, h: 0.9, rectRadius: 0.1, fill: { color: C.SECONDARY },
  }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });
  youSlide.addText("All the way through", {
    x: 6.1, y: 3.8, w: 3.4, h: 0.4,
    fontSize: 15, fontFace: FONT_B, color: C.CHARCOAL, align: "center", margin: 0,
  });
  T.addFooter(youSlide, FOOTER);
  youSlide.addNotes(NOTES_YOUDO);
  runSlideDiagnostics(youSlide, pres);

  exitTicketSlide(pres,
    [
      "Hard copy diagnostic: eight items, on your own, about eight minutes.",
      "Then at the door, sign me one event from the story.",
    ],
    NOTES_DIAG, FOOTER, { assessesSc: 2 });

  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner where your story happened, before you say what happened.",
      scItems: [
        "I can show where a story happens before I start it.",
        "I can retell a four-event story with the scene set first.",
        "I can keep two people apart so my partner knows who is who.",
      ],
      selfAssessment: "Unit Progress Card, Lesson 8 row",
    },
    NOTES_CLOSING, FOOTER);

  const pptxPath = path.join(OUT_DIR, `Deaf Sports Session ${SESSION} Showing The Story.pptx`);
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  /* ------------------------------------------------------- pdf: story ----- */

  const doc = createPdf({ title: STORY_RES.name });
  let y = addPdfHeader(doc, STORY_RES.name, {
    subtitle: "The Club That Kept Going",
    color: C.PRIMARY,
    lessonInfo: `Session ${SESSION} | one per pair | four events, two people`,
  });

  y = addTipBox(doc,
    "This story is written for this lesson. The real detail underneath it is that a Deaf cricket club started in Melbourne in the 1880s and is still playing today as the Melbourne Deaf Cricket Club. The source is in the unit document's Sources appendix. The two people are invented so you have somebody to place in the scene.",
    y, { color: C.ACCENT });

  y = addSectionHeading(doc, "The Club That Kept Going", y, { color: C.PRIMARY });
  STORY_ONE.forEach((para) => {
    doc.font("Sans").fontSize(14).fillColor("#" + C.CHARCOAL);
    doc.text(para, 50, y, { width: 495, lineGap: 5 });
    y += doc.heightOfString(para, { width: 495, lineGap: 5 }) + 14;
  });

  y += 6;
  y = addSectionHeading(doc, "Before you sign it", y, { color: C.SECONDARY });
  y = addWriteLine(doc, "1.  Where does it happen? Show that first.", y);
  y = addWriteLine(doc, "2.  Who is in it? Put them in two different places and keep them there.", y);
  y = addWriteLine(doc, "3.  What are the four things that happen?", y);

  addPdfFooter(doc, `Session ${SESSION} | Story Strip One | Years 5-6 Auslan`);
  await writePdf(doc, path.join(OUT_DIR, STORY_RES.fileName));
  console.log("PDF written: " + STORY_RES.fileName);

  /* -------------------------------------------------- pdf: diagnostic ----- */

  const ddoc = createPdf({ title: DIAG_RES.name });
  let dy = addPdfHeader(ddoc, DIAG_RES.name, {
    subtitle: "Eight items. On your own. About eight minutes.",
    color: C.PRIMARY,
    lessonInfo: `Session ${SESSION} | assessment piece 3 | collected and kept`,
  });

  dy = addBodyText(ddoc, "Name: ______________________________________________", dy);

  // 1 match
  dy = addSectionHeading(ddoc, "1. Match the question to what it asks about. Draw a line.", dy,
    { color: C.SECONDARY });
  const LEFT = ["Who", "What", "Where", "When"];
  const RIGHT = ["A place", "A person", "A year or a date", "A thing or an activity"];
  LEFT.forEach((l, i) => {
    ddoc.save();
    ddoc.font("Sans").fontSize(13).fillColor("#" + C.CHARCOAL);
    ddoc.rect(60, dy, 150, 26).lineWidth(0.8).strokeColor("#" + C.MUTED).stroke();
    ddoc.text(l, 70, dy + 7, { width: 130 });
    ddoc.rect(330, dy, 200, 26).lineWidth(0.8).strokeColor("#" + C.MUTED).stroke();
    ddoc.text(RIGHT[i], 340, dy + 7, { width: 180 });
    ddoc.restore();
    dy += 32;
  });

  dy += 6;
  dy = addSectionHeading(ddoc, "2. Put these three events in order. Write 1, 2 and 3.", dy,
    { color: C.SECONDARY });
  ["Melbourne hosts the Deaflympics.",
    "A Deaf cricket club starts in Melbourne.",
    "The first Australian Deaf Games are held in Sydney."].forEach((ev) => {
    ddoc.save();
    ddoc.rect(60, dy, 28, 26).lineWidth(1).strokeColor("#" + C.SECONDARY).stroke();
    ddoc.font("Sans").fontSize(13).fillColor("#" + C.CHARCOAL)
      .text(ev, 100, dy + 7, { width: 430 });
    ddoc.restore();
    dy += 32;
  });

  dy += 6;
  dy = addWriteLine(ddoc, "3. Write the year. Melbourne hosted the Deaflympics in ______________.", dy);

  dy = addSectionHeading(ddoc, "4. Circle the word that means the same as back then.", dy,
    { color: C.SECONDARY });
  ddoc.save();
  ["soon", "before", "in the past", "next"].forEach((w, i) => {
    ddoc.font("Sans").fontSize(14).fillColor("#" + C.CHARCOAL)
      .text(w, 70 + i * 120, dy, { width: 110 });
  });
  ddoc.restore();
  dy += 34;

  dy = addWriteLine(ddoc, "5. Name one thing in Deaf sport that is shown with a light or a flag.", dy);
  dy = addWriteLine(ddoc, "______________________________________________________________", dy);
  dy = addWriteLine(ddoc, "6. What problem did that light or flag solve?", dy);
  dy = addWriteLine(ddoc, "______________________________________________________________", dy);

  if (dy > 640) { ddoc.addPage(); dy = 60; }
  dy = addSectionHeading(ddoc, "7. Write four facts about one person or thing from the profile cards.", dy,
    { color: C.SECONDARY });
  ["Who or what: _________________________________________________",
    "What sport: __________________________________________________",
    "When: ________________________________________________________",
    "What they won or did: ________________________________________"].forEach((l) => {
    dy = addWriteLine(ddoc, l, dy);
  });

  dy = addWriteLine(ddoc, "8. When you are about to tell a story in Auslan, what do you show first?", dy);
  dy = addWriteLine(ddoc, "______________________________________________________________", dy);

  addPdfFooter(ddoc, `Session ${SESSION} | Hard Copy Diagnostic | Years 5-6 Auslan`);
  await writePdf(ddoc, path.join(OUT_DIR, DIAG_RES.fileName));
  console.log("PDF written: " + DIAG_RES.fileName);

  REPORT.print(`Session ${SESSION}`);
  console.log(`Session ${SESSION} build complete.`);
}

build().catch((err) => { console.error(err); process.exit(1); });
