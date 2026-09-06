"use strict";

// Persuasive Writing - Term 3 Week 6, Session 1 (Year 5/6 Literacy)
// "What makes writing persuasive - and can it be funny?"
//
// Victorian Curriculum 2.0: English, Literacy (Creating texts) and Literature,
// Levels 5-6 - understanding how persuasive texts are structured and how
// language choices position a reader.
//
// UNIT ANCHOR (locked from the school's own Persuasive Structure chart,
// megaprompt section 5b - never swapped for "introduction / body / conclusion"):
//   "Bold Beginning. Mighty Middle. Excellent Ending. Voice all the way through."
//
// Lesson shape: example-first. The school's own chart and its own 2019 Grade 6
// exemplar carry the modelling, because students recognise both and the
// exemplar is funny, which is exactly the teacher's Tuesday hook.
//
// Sources, all teacher-supplied: the DCE Persuasive Structure chart, the 2019
// DCE Grade 6 exemplar "Why Donkey Should Leave Shrek's Swamp", and the
// teacher's three funny topics.

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme } = require("../themes/factory");
const { composeGlanceNotes } = require("../themes/core/composeNotes");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
  PAGE, hex,
} = require("../themes/pdf_helpers");

const P = require("./persuasive_lib");

// Weeks 6 and 7 are ONE persuasive unit with one locked anchor, so both weeks
// ship on the same palette. Variant 5 is also the natural Week 6 rotation.
const UNIT_VARIANT = 5;
const T = createTheme("literacy", "grade56", UNIT_VARIANT);
const {
  C, FONT_H, FONT_B, SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, keyWordSlide, closingSlide,
  addTopBar, addBadge, addTitle, addCard, addFooter, addTextOnShape,
  addInstructionalImageCard, addRevealAnswerBar,
  clickBuild, runSlideDiagnostics,
} = T;

const SESSION = 1;
const FOOTER = "Persuasive Writing | Week 6 Session 1 | Year 5/6 Literacy";
const OUT_DIR = "output/Persuasive_W6_S1";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const CARD_RES = makeSessionResource(SESSION,
  "Persuasive Structure Card",
  "The anchor card for the whole fortnight. Print, glue in books, use every session.");
const RESOURCE_ITEMS = [CARD_RES];
fs.mkdirSync(RES_DIR, { recursive: true });

// Reveal band: content above any click-revealed bar must stop by 4.13.
const REVEAL_Y = 4.28;
const REVEAL_H = 0.78;
const CONTENT_FLOOR = 4.13;

// ─── Teacher notes (Glance Format, megaprompt sections 45-47) ────────────────

const NOTES_TITLE =
  "Open Week 6. Say the fortnight in one line: this week we learn it, next week we write and publish it.";

const NOTES_RESOURCES =
  "Prep slide. Print the Persuasive Structure Card, one per student - it is used every session for two weeks.\n" +
  "Have whiteboards, markers and books ready. The school's Persuasive Structure chart appears on screen this session; put the wall copy up too if you have it.\n" +
  "CATCH-UP: this is Session 1, so nothing is assumed. A student who joins later re-enters through the structure card and the launch.";

const NOTES_OVERVIEW =
  "For you, not the class. Skip past it when projecting.\n" +
  "Week 6 learns the craft, Week 7 writes and publishes three pieces on one topic. The anchor is the school's own chart, so the language on screen matches the wall.\n" +
  "Decision points today: the hinge after the exemplar, the boards check after the We Do, and the exit ticket. Between them keep the pace brisk.";

const NOTES_LAUNCH = composeGlanceNotes({
  answer: "open - listen for making the reader agree.",
  beats: [
    [
      "SAY: Some of you may remember persuasive writing from this year.",
      "If it feels new, that is okay.",
    ],
    [
      "ASK: What is a persuasive piece trying to do to its reader?",
      "30 sec. Turn and tell. Partner A first.",
      "EXPECT: change your mind, make you agree.",
    ],
    [
      "SCAN the room. Listen for the word reader.",
      "80%+ -> reveal the purpose, next slide.",
      "Less -> read an advert line aloud, ask what it wants, re-ask.",
    ],
    [
      "REVEAL after partner talk.",
      "SAY: To persuade someone to think and feel a certain way.",
    ],
  ],
  trap: [
    "thinking persuasive means being rude.",
    "Fix: point at the word feel, student re-says it.",
  ],
  prep: "Low-coupling launch: answerable from any advert a student has seen, so no earlier session is assumed. Whole block under 6 minutes.",
  tag: "[Launch | Retention and recall | HITS 6]",
});

const NOTES_HOOK = composeGlanceNotes({
  answer: "yes - all three could be persuasive, because each one takes a side.",
  beats: [
    [
      "SAY: Persuasive writing does not have to be serious.",
      "It has to have passion. That is voice.",
    ],
    [
      "ASK: Could we write a funny persuasive? Which topic would you pick?",
      "20 sec. Fingers at your chest, show me one, two or three.",
      "EXPECT: a spread across all three.",
    ],
    [
      "SCAN the fingers. Every hand at chest height.",
      "80%+ showing a number -> cold call one: Why that one?",
      "Less -> read the broccoli topic aloud in a serious voice, re-ask.",
    ],
  ],
  trap: [
    "picking a topic with no side to take, like My dog.",
    "Fix: ask what the reader must DO, student rewrites it as a should.",
  ],
  prep: "The teacher's own three topics, quoted from the planner. They give the lesson its energy and set up the You Do.",
  sources: P.SCHOOL_SOURCE_LINE,
  tag: "[Hook | Attention, focus and regulation | HITS 6]",
});

const NOTES_LI = composeGlanceNotes({
  beats: [
    "POINT to the learning intention. SAY: Today we learn how a persuasive piece is built.",
    "SAY: Read the I can statements with me. Everyone, together, on three.",
    "SAY: The first one everyone leaves with today - name the three parts.",
  ],
  prep: "SC1 is reachable by every student; SC2 is the core target the exit ticket assesses; SC3 stretches into explaining voice. Tier labels stay off the slide.",
  tag: "[LI/SC | Planning made visible | HITS 1]",
});

const NOTES_VOCAB = composeGlanceNotes({
  answer: "perspective means the side you are on, and how you see it.",
  beats: [
    [
      "SAY: Perspective is the word on our school chart.",
      "It means your side. Where you stand on the topic.",
    ],
    [
      "ASK: Homework on weekends. What is YOUR perspective?",
      "10 sec. Thumbs only, voices off. Up for yes, down for no.",
      "EXPECT: a split room.",
    ],
    "SAY: Same topic, different perspectives. That is why we have to persuade.",
  ],
  trap: [
    "confusing perspective with topic.",
    "Fix: name the topic, then ask for the side, student says both.",
  ],
  prep: "The one word carried from the school's chart. Students who can say their perspective can write a bold beginning.",
  sources: P.SCHOOL_SOURCE_LINE,
  tag: "[Key Vocabulary | Knowledge and memory | HITS 6]",
});

const NOTES_ANCHOR = composeGlanceNotes({
  answer: "Bold Beginning, Mighty Middle, Excellent Ending, voice throughout.",
  beats: [
    [
      "POINT to the school chart.",
      "SAY: You know this one. It is on our wall.",
    ],
    [
      "CLICK through the four parts.",
      "SAY: Topic and side. Then reasons with proof. Then one last push.",
      "And voice runs through all of it.",
    ],
    [
      "ASK: Say the three parts back to me.",
      "10 sec. Choral response: Everyone, together, on three.",
      "EXPECT: bold beginning, mighty middle, excellent ending.",
    ],
  ],
  trap: [
    "reciting the names without knowing each job.",
    "Fix: point at one band, student says its job.",
  ],
  stretch: "Say which part is hardest to write, and why.",
  help: "Give the structure card now, point to each band.",
  prep: "The unit anchor. Say this phrase in every I Do for the next two weeks, word for word, so the language never drifts.",
  sources: P.SCHOOL_SOURCE_LINE,
  tag: "[I Do | Explicit teaching | SC1 | HITS 3, 4]",
});

const NOTES_EXEMPLAR = composeGlanceNotes({
  answer: "paragraph one is the bold beginning, the last is the ending.",
  beats: [
    [
      "SAY: A Grade 6 piece from this school, 2019.",
      "Read the whole thing aloud. It is meant to be funny.",
    ],
    [
      "ASK: How many reasons are in the middle?",
      "20 sec. Write it... chin it... show me.",
      "EXPECT: three.",
    ],
    [
      "SCAN boards.",
      "80%+ -> cold call one: which reason is strongest? Convince us.",
      "Less -> re-read the middle, count on fingers, re-ask.",
    ],
  ],
  trap: [
    "counting the ending as a reason too.",
    "Fix: ask what the last paragraph adds, student recounts.",
  ],
  stretch: "Find the line where his voice is strongest. Say why.",
  help: "Cover all but one paragraph, ask only which part it is.",
  prep: "The exemplar is a low-resolution screenshot, so read it aloud rather than expecting students to read it from the board.",
  sources: P.SCHOOL_SOURCE_LINE,
  tag: "[I Do | Explicit teaching | SC1 | HITS 4, 6]",
});

const NOTES_CFU = composeGlanceNotes({
  answer: "B. It names the topic and the side.",
  beats: [
    [
      "SAY: Three openings, one topic. Only one is bold.",
      "Do not call out. This one is on boards.",
    ],
    [
      "ASK: Which is the bold beginning, A, B or C?",
      "30 sec. Write it... chin it... show me.",
      "EXPECT: B.",
    ],
    [
      "SCAN boards, back row first.",
      "80%+ -> cold call one B: how do you know? Then reveal.",
      "Less -> read A aloud, ask what side it takes, re-ask.",
    ],
    "REVEAL after boards are scanned. SAY: Topic plus side, or not bold.",
  ],
  trap: [
    "picking A because it sounds exciting but takes no side.",
    "Fix: ask what the writer wants, student re-tests all three.",
  ],
  prep: "The hinge of the lesson. A tests interest without a position; C states a topic with no perspective. Both wrong answers are diagnostic.",
  tag: "[CFU hinge | Supported application | SC2 | HITS 7, 8]",
});

const NOTES_WEDO = composeGlanceNotes({
  answer: "open - topic and a clear side in one sentence.",
  beats: [
    [
      "SAY: Our topic is the scrap paper bin. Watch how I pick my side.",
      "SAY: I believe every classroom must have a scrap paper bin.",
    ],
    [
      "ASK: Now write your own bold beginning.",
      "60 sec. Write it... chin it... show me.",
      "EXPECT: a must or should sentence about paper.",
    ],
    [
      "SCAN boards for topic AND side.",
      "80%+ -> cold call one strong, one shaky: name the side.",
      "Less -> rebuild one board on the frame, re-ask.",
    ],
  ],
  trap: [
    "writing about paper without taking a side.",
    "Fix: ring the gap, student adds must.",
  ],
  stretch: "Add a rhetorical question in front.",
  help: "Give the frame with the topic filled in.",
  prep: "The scrap paper bin is the teacher's own topic and becomes the modelled topic for all of Week 7. Keep it.",
  tag: "[We Do | Supported application | SC2 | HITS 4, 5]",
});

const NOTES_YOUDO = composeGlanceNotes({
  answer: "open - a bold beginning naming the topic and the side.",
  beats: [
    [
      "SAY: Your turn, your topic. Pick the broccoli or the snake.",
      "Be as funny as you like. Just take a real side.",
    ],
    [
      "POINT to the steps. SAY: Pick, decide your side, write two sentences.",
      "TIME: eight minutes in books.",
    ],
    [
      "CIRCULATE. Read for a side, not for neat handwriting.",
      "COLLECT two you will read aloud at the closing.",
    ],
  ],
  trap: [
    "describing the topic instead of arguing about it.",
    "Fix: ask what the reader must DO, student rewrites.",
  ],
  stretch: "Write a second bold beginning for the opposite side.",
  help: "Give the frame with the topic filled in, side blank.",
  prep: "Different topics from the We Do on purpose, so this is application and not a copy. Books, not sheets.",
  tag: "[You Do | Mastery and application | SC2 | HITS 10]",
});

const NOTES_EXIT = composeGlanceNotes({
  answer: "any sentence naming dogs at school AND the writer's side.",
  beats: [
    [
      "SAY: One sentence. New topic. Dogs at school.",
      "Name the topic and your side. That is all.",
    ],
    [
      "TIME: three minutes. Books open on the desk when you finish.",
      "COLLECT by walking the rows, not by calling names.",
    ],
    [
      "SORT into two piles as you collect: side named, side missing.",
      "A thick second pile -> open Session 2 by rebuilding one together.",
    ],
  ],
  prep: "Assesses the core target: write a bold beginning naming topic and side. Sort into two piles as you collect.",
  tag: "[Exit Ticket | Mastery and application | SC2 | HITS 8]",
});

const NOTES_CLOSING = composeGlanceNotes({
  answer: "open - most students should show two.",
  beats: [
    "READ the two collected bold beginnings aloud. SAY: Listen for the side.",
    [
      "ASK: Rate yourself on the three statements.",
      "15 sec. Fingers at your chest. One, two or three. Show me.",
      "EXPECT: most on two.",
    ],
    "SAY: Next session we make the words themselves do the persuading.",
  ],
  prep: "Self-assessment feeds Session 2 grouping. Anyone showing one gets the structure card in front of them next session.",
  tag: "[Closing | Retention and recall | HITS 1, 9]",
});

// ─── Slide chrome helper for custom layouts ─────────────────────────────────

function customSlide(pres, badgeText, badgeColor, title) {
  const s = pres.addSlide();
  addTopBar(s, badgeColor);
  addBadge(s, badgeText, { color: badgeColor, w: 1.9 });
  addTitle(s, title);
  return s;
}

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // 1. Title
  titleSlide(pres, "Persuasive Writing",
    "Bold Beginning. Mighty Middle. Excellent Ending.",
    "Week 6 Session 1 | Year 5/6 Literacy", NOTES_TITLE);

  // 2. Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // 3. Teacher-facing overview
  (() => {
    const s = customSlide(pres, "For the teacher", C.MUTED, "The fortnight at a glance");
    const cardY = CONTENT_TOP;
    const cardH = SAFE_BOTTOM - CONTENT_TOP;

    addCard(s, 0.5, cardY, 4.4, cardH, { strip: C.PRIMARY });
    s.addText("Where this sits", {
      x: 0.7, y: cardY + 0.12, w: 4.0, h: 0.28,
      fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText([
      { text: "Week 6, learn the craft", options: { bold: true, color: C.PRIMARY, breakLine: true } },
      { text: "S1  Structure and voice", options: { breakLine: true } },
      { text: "S2  Emotive and precise words", options: { breakLine: true } },
      { text: "S3  Questions, bossy verbs, promises", options: { breakLine: true } },
      { text: "S4  Modality and the mighty middle", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "Week 7, write and publish", options: { bold: true, color: C.ASSESS, breakLine: true } },
      { text: "One topic, three published pieces:", options: { breakLine: true } },
      { text: "a letter, a poster, a news article.", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "Anchor, said in every I Do:", options: { color: C.MUTED, breakLine: true } },
      { text: P.ANCHOR_PHRASE, options: { color: C.PRIMARY, bold: true } },
    ], {
      x: 0.7, y: cardY + 0.46, w: 4.0, h: cardH - 0.60,
      fontSize: 12.5, fontFace: FONT_B, color: C.CHARCOAL,
      valign: "top", margin: 0, paraSpaceAfter: 3,
    });

    addCard(s, 5.1, cardY, 4.4, cardH, { strip: C.ACCENT });
    s.addText("How today works", {
      x: 5.3, y: cardY + 0.12, w: 4.0, h: 0.28,
      fontSize: 14, fontFace: FONT_H, color: C.ACCENT, bold: true, margin: 0,
    });
    s.addText([
      { text: "Shape: example-first. The school's own chart and its own 2019 Grade 6 exemplar carry the modelling.", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "Decision points:", options: { bold: true, color: C.ACCENT, breakLine: true } },
      { text: "the A/B/C hinge, the boards check in the We Do, and the exit ticket.", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "Assumption:", options: { bold: true, color: C.ACCENT, breakLine: true } },
      { text: "students have met persuasive writing this year but not this lens. Language stays mixed-readiness throughout.", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "Flex: cut the You Do to five minutes if the exemplar runs long.", options: { italic: true, color: C.MUTED } },
    ], {
      x: 5.3, y: cardY + 0.46, w: 4.0, h: cardH - 0.60,
      fontSize: 12.5, fontFace: FONT_B, color: C.CHARCOAL,
      valign: "top", margin: 0, paraSpaceAfter: 3,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_OVERVIEW);
    runSlideDiagnostics(s, pres);
  })();

  // 4. Launch - what do we already know?
  (() => {
    const s = customSlide(pres, "Launch", C.SECONDARY, "What do we already know?");
    const y0 = CONTENT_TOP;

    addTextOnShape(s, "A persuasive piece is written to do something to its reader.", {
      x: 0.5, y: y0, w: 9, h: 0.70, rectRadius: 0.08,
      fill: { color: C.SECONDARY },
    }, {
      fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0.1,
    });

    s.addText("What is it trying to do?", {
      x: 0.5, y: y0 + 0.88, w: 9, h: 1.20,
      fontSize: 48, fontFace: FONT_H, color: C.PRIMARY, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    addTextOnShape(s, "Turn and tell. Partner A first.", {
      x: 2.85, y: y0 + 2.24, w: 4.3, h: 0.62, rectRadius: 0.08,
      fill: { color: C.WHITE }, line: { color: C.SECONDARY, width: 1.5 },
    }, {
      fontSize: 17, fontFace: FONT_B, color: C.SECONDARY, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    clickBuild(s, [
      () => {
        addRevealAnswerBar(s, ["To persuade someone to think and FEEL a certain way"], {
          y: REVEAL_Y, h: REVEAL_H, fontSize: 21, label: "Purpose", color: C.PRIMARY,
        });
      },
    ]);

    addFooter(s, FOOTER);
    s.addNotes(NOTES_LAUNCH);
    runSlideDiagnostics(s, pres);
  })();

  // 5. Hook - could we write a funny persuasive?
  (() => {
    const s = customSlide(pres, "Hook", C.ACCENT, "Could we write a funny persuasive?");
    const y0 = CONTENT_TOP;
    const chipH = 0.88;
    const gap = 0.15;

    P.FUNNY_TOPICS.forEach((topic, i) => {
      const cy = y0 + i * (chipH + gap);
      s.addShape("roundRect", {
        x: 0.5, y: cy, w: 9, h: chipH, rectRadius: 0.08,
        fill: { color: C.WHITE }, line: { color: C.ACCENT, width: 1.5 },
      });
      s.addShape("roundRect", {
        x: 0.70, y: cy + 0.20, w: 0.55, h: 0.55, rectRadius: 0.275,
        fill: { color: C.ACCENT },
      });
      s.addText(String(i + 1), {
        x: 0.70, y: cy + 0.20, w: 0.55, h: 0.55,
        fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      s.addText(topic, {
        x: 1.44, y: cy, w: 7.8, h: chipH,
        fontSize: 21, fontFace: FONT_B, color: C.CHARCOAL,
        valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
    });

    const noteY = y0 + 3 * (chipH + gap) + 0.06;
    s.addText("Persuasion is not about being serious. It is about showing passion.", {
      x: 0.5, y: noteY, w: 9, h: 0.50,
      fontSize: 17, fontFace: FONT_H, color: C.PRIMARY, italic: true, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_HOOK);
    runSlideDiagnostics(s, pres);
  })();

  // 6. Learning Intention and Success Criteria
  liSlide(pres,
    "We are learning how a persuasive piece is built and how it convinces a reader.",
    [
      "I can name the three parts of a persuasive piece.",
      "I can write a bold beginning that names my topic and my side.",
      "I can explain how a writer's voice makes me want to agree.",
    ],
    NOTES_LI, FOOTER);

  // 7. Key vocabulary
  keyWordSlide(pres, {
    word: "perspective",
    meaning: "The side you are on, and how you see the topic.",
    example: "Two people can look at the same topic and have a different perspective.",
    routine: ["Say it", "Take a side", "Say why"],
    color: C.SECONDARY,
    title: "The word on our chart",
  }, NOTES_VOCAB, FOOTER);

  // 8. I Do A - the anchor, built one band per click
  (() => {
    const s = customSlide(pres, "I Do", C.PRIMARY, "Our chart, one part at a time");
    const y0 = CONTENT_TOP;
    const h = SAFE_BOTTOM - y0;

    addInstructionalImageCard(s, P.IMAGES.structureChart, {
      x: 0.5, y: y0, w: 3.55, h, fit: "contain", strip: C.PRIMARY,
      caption: "Our school Persuasive Structure chart",
    });

    const base = { x: 4.35, y: y0 + 0.05, w: 5.15, h: h - 0.10 };
    clickBuild(s, [
      () => P.drawStructureStrip(s, T, { ...base, parts: ["BB"] }),
      () => P.drawStructureStrip(s, T, { ...base, parts: ["MM"] }),
      () => P.drawStructureStrip(s, T, { ...base, parts: ["EE"] }),
      () => P.drawStructureStrip(s, T, { ...base, parts: ["voice"] }),
    ]);

    addFooter(s, FOOTER);
    s.addNotes(NOTES_ANCHOR);
    runSlideDiagnostics(s, pres);
  })();

  // 9. I Do B - the school's own Grade 6 exemplar
  (() => {
    const s = customSlide(pres, "I Do", C.PRIMARY, "A Grade 6 piece from this school");
    const y0 = CONTENT_TOP;
    const h = SAFE_BOTTOM - y0;

    addInstructionalImageCard(s, P.IMAGES.shrekExemplar, {
      x: 0.5, y: y0, w: 3.55, h, fit: "contain", strip: C.ASSESS,
      caption: "Why Donkey Should Leave Shrek's Swamp, 2019",
    });

    const rx = 4.35;
    const rw = 5.15;
    addCard(s, rx, y0, rw, 1.05, { strip: C.PRIMARY });
    s.addText("Find the three parts", {
      x: rx + 0.2, y: y0 + 0.08, w: rw - 0.4, h: 0.30,
      fontSize: 16, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("Which paragraph does each job?", {
      x: rx + 0.2, y: y0 + 0.42, w: rw - 0.4, h: 0.52,
      fontSize: 15, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
      valign: "top",
    });

    P.drawStructureStrip(s, T, {
      x: rx, y: y0 + 1.20, w: rw, h: h - 1.20, showVoice: true,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_EXEMPLAR);
    runSlideDiagnostics(s, pres);
  })();

  // 10. CFU hinge - which one is the bold beginning?
  (() => {
    const s = customSlide(pres, "CFU", C.ALERT, "Which one is the bold beginning?");
    const y0 = CONTENT_TOP;

    addTextOnShape(s, "Topic: school should start at 10 am", {
      x: 0.5, y: y0, w: 9, h: 0.48, rectRadius: 0.08,
      fill: { color: C.ALERT },
    }, {
      fontSize: 16, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    const options = [
      { key: "A", text: "Have you ever been so tired you forgot your own name?" },
      { key: "B", text: "School must start at 10 am, because tired students cannot learn." },
      { key: "C", text: "This piece is going to be about what time school starts." },
    ];
    // Options must clear the click-revealed answer bar: last card bottom
    // 1.90 + 3 x 0.78 - 0.12 = 4.12, just under the 4.13 content floor.
    const oh = 0.66;
    const og = 0.12;
    options.forEach((opt, i) => {
      const oy = y0 + 0.60 + i * (oh + og);
      s.addShape("roundRect", {
        x: 0.5, y: oy, w: 9, h: oh, rectRadius: 0.08,
        fill: { color: C.WHITE }, line: { color: C.PRIMARY, width: 1.5 },
      });
      s.addShape("roundRect", {
        x: 0.68, y: oy + 0.11, w: 0.44, h: 0.44, rectRadius: 0.22,
        fill: { color: C.PRIMARY },
      });
      s.addText(opt.key, {
        x: 0.68, y: oy + 0.11, w: 0.44, h: 0.44,
        fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      s.addText(opt.text, {
        x: 1.28, y: oy, w: 7.97, h: oh,
        fontSize: 17, fontFace: FONT_B, color: C.CHARCOAL,
        valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
    });

    clickBuild(s, [
      () => {
        addRevealAnswerBar(s, ["B - it names the topic AND the side"], {
          y: REVEAL_Y, h: REVEAL_H, fontSize: 22, color: C.SUCCESS,
        });
      },
    ]);

    addFooter(s, FOOTER);
    s.addNotes(NOTES_CFU);
    runSlideDiagnostics(s, pres);
  })();

  // 11. We Do - build a bold beginning together
  (() => {
    const s = customSlide(pres, "We Do", C.SUCCESS, "Build a bold beginning together");
    const y0 = CONTENT_TOP;

    addTextOnShape(s, P.MODEL_TOPIC, {
      x: 0.5, y: y0, w: 9, h: 0.82, rectRadius: 0.08,
      fill: { color: C.SUCCESS },
    }, {
      fontSize: 26, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    const fy = y0 + 0.96;
    addCard(s, 0.5, fy, 9, 1.75, { strip: C.PRIMARY });
    s.addText("The frame", {
      x: 0.75, y: fy + 0.10, w: 4, h: 0.28,
      fontSize: 15, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText([
      { text: "I believe every classroom ", options: { color: C.CHARCOAL } },
      { text: "must", options: { color: C.ALERT, bold: true } },
      { text: " have a scrap paper bin, because ", options: { color: C.CHARCOAL } },
      { text: "[your reason]", options: { color: C.MUTED, italic: true } },
      { text: ".", options: { color: C.CHARCOAL } },
    ], {
      x: 0.75, y: fy + 0.44, w: 8.5, h: 1.20,
      fontSize: 24, fontFace: FONT_B, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    addTextOnShape(s, "Your turn on boards. Write it... chin it... show me.", {
      x: 0.5, y: fy + 1.93, w: 9, h: 0.72, rectRadius: 0.08,
      fill: { color: C.WHITE }, line: { color: C.SUCCESS, width: 1.5 },
    }, {
      fontSize: 19, fontFace: FONT_B, color: C.SUCCESS, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_WEDO);
    runSlideDiagnostics(s, pres);
  })();

  // 12. You Do - your topic, your side
  (() => {
    const s = customSlide(pres, "You Do", C.ASSESS, "Your topic. Your side.");
    const y0 = CONTENT_TOP;

    const picks = [P.FUNNY_TOPICS[1], P.FUNNY_TOPICS[2]];
    const ph = 1.00;
    picks.forEach((topic, i) => {
      const py = y0 + i * (ph + 0.16);
      s.addShape("roundRect", {
        x: 0.5, y: py, w: 5.6, h: ph, rectRadius: 0.08,
        fill: { color: C.WHITE }, line: { color: C.ASSESS, width: 1.5 },
      });
      s.addText(topic, {
        x: 0.74, y: py, w: 5.12, h: ph,
        fontSize: 18, fontFace: FONT_B, color: C.CHARCOAL,
        valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
    });

    const steps = ["Pick one topic.", "Decide your side.", "Write two sentences."];
    const sy = y0;
    addCard(s, 6.3, sy, 3.2, 2.16, { strip: C.ASSESS });
    s.addText("Steps", {
      x: 6.52, y: sy + 0.10, w: 2.8, h: 0.30,
      fontSize: 15, fontFace: FONT_B, color: C.ASSESS, bold: true, margin: 0,
    });
    steps.forEach((step, i) => {
      s.addText(String(i + 1) + ". " + step, {
        x: 6.52, y: sy + 0.48 + i * 0.52, w: 2.8, h: 0.46,
        fontSize: 15, fontFace: FONT_B, color: C.CHARCOAL,
        valign: "middle", margin: 0,
      });
    });

    const by = y0 + 2.36;
    addTextOnShape(s, "Name the topic. Name your side. Be as funny as you like.", {
      x: 0.5, y: by, w: 9, h: 0.78, rectRadius: 0.08,
      fill: { color: C.ASSESS },
    }, {
      fontSize: 21, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    s.addText("Eight minutes in your books.", {
      x: 0.5, y: by + 0.86, w: 9, h: 0.46,
      fontSize: 16, fontFace: FONT_B, color: C.MUTED, italic: true,
      align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
    runSlideDiagnostics(s, pres);
  })();

  // 13. Exit ticket
  P.exitTicketPanel(pres, T, {
    topic: "Dogs should be allowed at school",
    task: "Write ONE bold beginning sentence.",
    cue: "Name the topic. Name your side.",
    footer: FOOTER,
    notes: NOTES_EXIT,
  });

  // 14. Closing
  closingSlide(pres, {
    reflectionPrompt: "Which part of a persuasive piece do you already feel ready to write?",
    scItems: [
      "I can name the three parts of a persuasive piece.",
      "I can write a bold beginning that names my topic and my side.",
      "I can explain how a writer's voice makes me want to agree.",
    ],
    selfAssessment: "Fingers at your chest. One, two or three. Show me.",
    takeaways: [P.ANCHOR_PHRASE],
  }, NOTES_CLOSING);

  const pptxPath = path.join(OUT_DIR, "Persuasive Writing W6 S1.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);
}

// ─── Resource: the Persuasive Structure Card ────────────────────────────────

async function buildCard() {
  const doc = createPdf({ title: CARD_RES.name });
  let y = addPdfHeader(doc, "Persuasive Structure Card", {
    subtitle: "Keep this in your book. You will use it every session for two weeks.",
    color: C.PRIMARY,
    lessonInfo: "Year 5/6 Literacy | Persuasive Writing",
    showNameDate: false,
  });

  y = addTipBox(doc, P.ANCHOR_PHRASE, y, { color: C.ACCENT });
  y += 6;

  const bandColors = [C.PRIMARY, C.SECONDARY, C.ASSESS];
  const bandDetail = [
    [
      "Name the topic.",
      "Say which side you are on.",
      "Give the reader a reason to keep reading.",
      "Frame: I believe [topic] should / must [what], because [reason].",
    ],
    [
      "One reason per paragraph. Three is plenty.",
      "Reason: [Position] because [reason].",
      "Evidence: According to... / For example...",
      "Explain: This means that...",
      "Link: Therefore, ...",
    ],
    [
      "Say your side one more time, in new words.",
      "Tell the reader what to DO.",
      "Finish on a short, strong line.",
    ],
  ];

  P.STRUCTURE.forEach((band, i) => {
    doc.save();
    doc.roundedRect(PAGE.MARGIN, y, PAGE.CONTENT_W, 22, 3).fill(hex(bandColors[i]));
    doc.fontSize(12).font("Sans-Bold").fillColor("#FFFFFF");
    doc.text(band.label + "   (" + band.sub + ")", PAGE.MARGIN + 8, y + 6,
      { width: PAGE.CONTENT_W - 16 });
    doc.restore();
    y += 27;

    doc.fontSize(10.5).font("Sans").fillColor("#000000");
    bandDetail[i].forEach((line) => {
      doc.circle(PAGE.MARGIN + 8, y + 5, 2).fill(hex(bandColors[i]));
      doc.fillColor("#000000").text(line, PAGE.MARGIN + 18, y, { width: PAGE.CONTENT_W - 26 });
      y += 15;
    });
    y += 8;
  });

  y = addSectionHeading(doc, "Voice: all the way through", y, { color: C.ACCENT });
  y += 4;
  y = P.addWordChipsPdf(doc, y, "The four voice tools",
    P.VOICE_TOOLS.map((t) => t.label + ": " + t.cue),
    { color: C.ACCENT, perRow: 2, chipH: 22, fontSize: 9 });

  y = addBodyText(doc,
    "Voice is passion. It is what makes a reader believe you, whether your topic is serious or silly.",
    y, { italic: true });

  addPdfFooter(doc, "Persuasive Writing | Year 5/6 Literacy | Diamond Creek East PS");
  await writePdf(doc, path.join(OUT_DIR, CARD_RES.fileName));
  console.log("Wrote " + CARD_RES.name);
}

(async () => {
  await build();
  await buildCard();
})().catch((err) => {
  console.error(err);
  process.exit(1);
});
