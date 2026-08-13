"use strict";

// Persuasive Writing - Term 3 Week 7, Session 3 (Year 5/6 Literacy)
// "Piece two: the poster, for a reader who gives you three seconds"
//
// Victorian Curriculum 2.0: English, Literacy (Creating texts) and Language
// (Text structure and organisation), Levels 5-6 - selecting and compressing
// language and layout for a glancing audience.
//
// UNIT ANCHOR (locked): Bold Beginning. Mighty Middle. Excellent Ending.
// Voice all the way through.  Week 7: Same belief. Same reasons. Different voice.
//
// The poster is the SAME belief compressed. Every element on the poster maps to
// a band of the anchor: the slogan is the bold beginning, the one line of proof
// is the mighty middle, the call to action is the excellent ending.
//
// Designed-visual lesson: the I Do and the We Do both use structured poster
// specs through the shared drawMockupPreview path, never flat text blocks.

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme } = require("../themes/factory");
const { composeGlanceNotes } = require("../themes/core/composeNotes");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addLinedArea,
  addPosterPairPdf, addPosterMockupPdf,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
  PAGE, hex,
} = require("../themes/pdf_helpers");

const P = require("./persuasive_lib");

const UNIT_VARIANT = 5;
const T = createTheme("literacy", "grade56", UNIT_VARIANT);
const {
  C, FONT_H, FONT_B, SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, closingSlide,
  annotatedModelSlide, compareVisualSlide,
  addCard, addFooter, addTextOnShape, addRevealAnswerBar,
  clickBuild, runSlideDiagnostics,
} = T;

const SESSION = 3;
const FOOTER = "Persuasive Writing | Week 7 Session 3 | Year 5/6 Literacy";
const OUT_DIR = "output/Persuasive_W7_S3";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const POSTER_RES = makeSessionResource(SESSION,
  "Poster Planner",
  "Two posters to compare, then a poster frame to plan and sketch your own.");
const RESOURCE_ITEMS = [POSTER_RES];
fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Poster specs (structured, so students infer from layout) ───────────────

// The I Do poster: the scrap paper bin, done well.
const POSTER_MODEL = {
  accent: C.PRIMARY,
  components: [
    { kind: "masthead", text: "DON'T BIN IT. REUSE IT.", scale: 1.15 },
    { kind: "hero", mode: "photo", scale: 1.9 },
    { kind: "stat", text: "Half our paper is used on one side only", scale: 0.85 },
    { kind: "cta", text: "Put your scraps in the blue tray", scale: 0.9 },
  ],
};

// We Do pair: same topic, same belief, one poster designed for a glance and
// one written like a letter. Students compare and say WHY.
const POSTER_STRONG = {
  accent: C.SUCCESS,
  components: [
    { kind: "masthead", text: "THREE MINUTES. ONE PLANET.", scale: 1.15 },
    { kind: "hero", mode: "photo", scale: 1.9 },
    { kind: "stat", text: "Litter picked up in 3 minutes a day", scale: 0.85 },
    { kind: "cta", text: "Grab a bag at the office", scale: 0.9 },
  ],
};

const POSTER_WEAK = {
  accent: C.MUTED,
  pageFill: "F2F3F5",
  components: [
    { kind: "heading", text: "information about the yard clean up idea", scale: 0.8 },
    { kind: "textBlock", count: 6, scale: 2.8 },
    { kind: "caption", text: "please see your teacher for further details", scale: 0.7 },
  ],
};

// ─── Teacher notes ──────────────────────────────────────────────────────────

const NOTES_TITLE =
  "Week 7, Session 3. One line: same belief again, but this reader walks past and gives you three seconds.";

const NOTES_RESOURCES =
  "Prep slide. Print the Poster Planner. Students need their Persuasive Plan and their Letter Builder draft.\n" +
  "Coloured pencils or markers for the sketch. A3 paper if you want the finished posters to go up.\n" +
  "CATCH-UP: no plan or letter? Give a spare Persuasive Plan. A belief and one reason is enough to design a poster today.";

const NOTES_OVERVIEW =
  "For you, not the class. Skip past it when projecting.\n" +
  "The poster is the same belief compressed. Slogan is the bold beginning, one line of proof is the mighty middle, the call to action is the excellent ending. Say that mapping out loud.\n" +
  "Decision points: the three-second hinge, the compare-and-say-why check, and the exit ticket. Between them keep the pace brisk.";

const NOTES_LAUNCH = composeGlanceNotes({
  answer: "the slogan and the picture. Almost nobody reads the small print.",
  beats: [
    [
      "SAY: I am going to show you a poster for three seconds only.",
      "SHOW it, count three, cover it.",
    ],
    [
      "ASK: What did you actually take in?",
      "20 sec. Turn and tell. Partner B first.",
      "EXPECT: the big words and the picture.",
    ],
    [
      "SCAN the room as pairs talk.",
      "80%+ naming the big words -> next slide.",
      "Less -> show it again for three seconds, re-ask.",
    ],
  ],
  trap: [
    "claiming to have read the whole poster.",
    "Fix: cover it and ask for the small print, student cannot recall it.",
  ],
  prep: "Low-coupling launch: needs no plan and no earlier session. The three-second rule is the whole lesson in one move.",
  tag: "[Launch | Attention, focus and regulation | HITS 6]",
});

const NOTES_LI = composeGlanceNotes({
  beats: [
    "POINT to the intention. SAY: Today, the same belief in about twelve words.",
    "SAY: Read the I can statements with me. Everyone, together, on three.",
    "SAY: Everyone leaves with a slogan they could put on a wall.",
  ],
  prep: "SC1 reachable by all; SC2 is the core target the exit ticket assesses; SC3 stretches into justifying design choices.",
  tag: "[LI/SC | Planning made visible | HITS 1]",
});

const NOTES_MODEL = composeGlanceNotes({
  answer: "slogan, picture, one line of proof, then the call to action.",
  beats: [
    "POINT to the poster. SAY: Read it the way you would in a corridor.",
    [
      "SAY: The slogan IS my bold beginning.",
      "SAY: The proof IS my mighty middle. The bottom line IS my ending.",
    ],
    [
      "ASK: How many words are on the whole poster?",
      "15 sec. Write it... chin it... show me.",
      "EXPECT: under twenty.",
    ],
  ],
  trap: [
    "thinking a poster is a shrunk down letter.",
    "Fix: count the words aloud, student compares their letter.",
  ],
  stretch: "Say what the picture adds that words cannot.",
  help: "Cover all but the slogan, name only that band.",
  prep: "Say the mapping out loud, in the anchor's own words. It is what stops students starting the topic again from scratch.",
  tag: "[I Do | Explicit teaching | SC1 | HITS 3, 4]",
});

const NOTES_COMPARE = composeGlanceNotes({
  answer: "A. It reads at a glance. B has to be stood in front of.",
  beats: [
    "SAY: Two posters, one belief. Not prettier. Which one WORKS.",
    [
      "ASK: Which works at a glance, and what makes it work?",
      "60 sec. Turn and tell, then write it on boards.",
      "EXPECT: A, because the big words carry it.",
    ],
    [
      "SCAN boards for a REASON, not just a letter.",
      "80%+ -> cold call two for their reason.",
      "Less -> stand at the back, ask what you can read, re-ask.",
    ],
  ],
  trap: [
    "choosing A with no reason attached.",
    "Fix: ask what B needs to change, student names it.",
  ],
  stretch: "Rewrite B's heading so it works.",
  help: "Give the stems: A works because... B fails because...",
  prep: "Both posters argue the same thing. The difference is compression, not content. Keep students on that.",
  tag: "[We Do | Supported application | SC2 | HITS 5, 7]",
});

const NOTES_CFU = composeGlanceNotes({
  answer: "B. Short, certain, and it tells you what to do.",
  beats: [
    [
      "SAY: Three slogans for the same poster. Only one belongs on a wall.",
      "Do not call out. Boards for this one.",
    ],
    [
      "ASK: Which is the slogan, A, B or C?",
      "30 sec. Write it... chin it... show me.",
      "EXPECT: B.",
    ],
    [
      "SCAN boards, back row first.",
      "80%+ -> cold call one B: how many words is that? Then reveal.",
      "Less -> read A aloud at walking pace, re-ask.",
    ],
    "REVEAL after boards are scanned. SAY: Four words. A bossy verb. Done.",
  ],
  trap: [
    "picking A because it explains the most.",
    "Fix: time three seconds, read A, student sees it does not fit.",
  ],
  prep: "The hinge. A is a full sentence, C is vague. Both wrong answers show compression not yet landing.",
  tag: "[CFU hinge | Supported application | SC2 | HITS 7, 8]",
});

const NOTES_YOUDO = composeGlanceNotes({
  answer: "open - a slogan, one line of proof and a call to action.",
  beats: [
    [
      "SAY: Same belief as your letter. Do not change topic.",
      "SAY: Twelve words on the whole poster is plenty.",
    ],
    [
      "POINT to the Planner. SAY: Slogan, proof, call to action.",
      "TIME: eight minutes on words, then twelve sketching.",
    ],
    [
      "CIRCULATE. Read the slogan first. If it is a sentence, send it back.",
      "COLLECT nothing. Posters are published in Session 5.",
    ],
  ],
  trap: [
    "writing paragraphs on the poster.",
    "Fix: cover all but twelve words, student picks which twelve.",
  ],
  stretch: "Design a second slogan aimed at teachers.",
  help: "Give the bossy verb list, build the slogan from one.",
  prep: "Words before pictures. A student who sketches first spends the session colouring, not persuading.",
  tag: "[You Do | Mastery and application | SC2 | HITS 10]",
});

const NOTES_EXIT = composeGlanceNotes({
  answer: "a slogan of about six words, with a bossy verb in it.",
  beats: [
    [
      "SAY: One slogan. Your topic. Six words or fewer.",
      "SAY: A bossy verb somewhere in it.",
    ],
    [
      "TIME: three minutes. Planners stay with you.",
      "WALK the rows and read over shoulders.",
    ],
    [
      "SORT in your head: short with a verb, or still a sentence.",
      "Still sentences -> two minutes each before Session 4 starts.",
    ],
  ],
  prep: "Assesses the core target: compress the belief into a slogan. Word count is the visible evidence.",
  tag: "[Exit Ticket | Mastery and application | SC2 | HITS 8]",
});

const NOTES_CLOSING = composeGlanceNotes({
  answer: "open - most students should show two.",
  beats: [
    "COLD CALL four students for their slogan only. No explaining.",
    [
      "ASK: Rate yourself on the three statements.",
      "15 sec. Fingers at your chest. One, two or three. Show me.",
      "EXPECT: most on two.",
    ],
    "SAY: Tomorrow the same belief again, but this reader wants the facts first.",
  ],
  prep: "Reading four slogans aloud back to back makes compression audible. Do not let students explain them.",
  tag: "[Closing | Retention and recall | HITS 1, 9]",
});

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  titleSlide(pres, "The Poster",
    "Piece two: three seconds to convince",
    "Week 7 Session 3 | Year 5/6 Literacy", NOTES_TITLE);

  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  (() => {
    const s = P.customSlide(pres, T, "For the teacher", C.MUTED, "Session 3 at a glance", { badgeW: 2.4 });
    const cardY = CONTENT_TOP;
    const cardH = SAFE_BOTTOM - CONTENT_TOP;

    addCard(s, 0.5, cardY, 4.4, cardH, { strip: C.PRIMARY });
    s.addText("Where this sits", {
      x: 0.7, y: cardY + 0.12, w: 4.0, h: 0.28,
      fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText([
      { text: "Piece two of three, from the same plan.", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: P.ANCHOR_PHRASE_W7, options: { color: C.ASSESS, bold: true, breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "The mapping, say it out loud:", options: { bold: true, breakLine: true } },
      { text: "the slogan IS the bold beginning", options: { breakLine: true } },
      { text: "one line of proof IS the mighty middle", options: { breakLine: true } },
      { text: "the call to action IS the excellent ending", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "Students must not change topic. This is compression, not a new piece.", options: { italic: true, color: C.MUTED } },
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
      { text: "Shape: problem-first. The three second reveal in the launch proves the point before any modelling.", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "Words before pictures:", options: { bold: true, color: C.ACCENT, breakLine: true } },
      { text: "a student who sketches first spends the session colouring, not persuading. Hold the pencils back.", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "Decision points:", options: { bold: true, color: C.ACCENT, breakLine: true } },
      { text: "the slogan hinge, the compare-and-say-why check, and the exit ticket.", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "Flex: sketching can finish in Session 5 publishing time.", options: { italic: true, color: C.MUTED } },
    ], {
      x: 5.3, y: cardY + 0.46, w: 4.0, h: cardH - 0.60,
      fontSize: 12.5, fontFace: FONT_B, color: C.CHARCOAL,
      valign: "top", margin: 0, paraSpaceAfter: 3,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_OVERVIEW);
    runSlideDiagnostics(s, pres);
  })();

  // Launch - the three second rule
  (() => {
    const s = P.customSlide(pres, T, "Launch", C.SECONDARY, "Three seconds. Go.");
    const y0 = CONTENT_TOP;

    s.addText("3", {
      x: 0.5, y: y0, w: 9, h: 1.55,
      fontSize: 108, fontFace: FONT_H, color: C.SECONDARY, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
    s.addText("seconds is all a poster ever gets", {
      x: 0.5, y: y0 + 1.60, w: 9, h: 0.66,
      fontSize: 28, fontFace: FONT_H, color: C.PRIMARY, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    addTextOnShape(s, "What did you actually take in? Turn and tell. Partner B first.", {
      x: 0.5, y: y0 + 2.44, w: 9, h: 0.72, rectRadius: 0.08,
      fill: { color: C.WHITE }, line: { color: C.SECONDARY, width: 1.5 },
    }, {
      fontSize: 18, fontFace: FONT_B, color: C.SECONDARY, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    s.addText("Show the next slide for three seconds, then click straight past it.", {
      x: 0.5, y: y0 + 3.26, w: 9, h: 0.42,
      fontSize: 14, fontFace: FONT_B, color: C.MUTED, italic: true,
      align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_LAUNCH);
    runSlideDiagnostics(s, pres);
  })();

  liSlide(pres,
    "We are learning to compress a whole argument into a poster a reader takes in at a glance.",
    [
      "I can say what a reader notices first on a poster.",
      "I can write a slogan and a call to action for my topic.",
      "I can explain why my poster works at a glance.",
    ],
    NOTES_LI, FOOTER);

  // I Do - the annotated model poster
  annotatedModelSlide(pres, "I Do", "The same belief, in twelve words",
    [
      { text: "Read it like a corridor", role: "header" },
      "What hits you first?",
      "Where does the proof sit?",
      "What does it tell you to DO?",
      { text: "Same belief as my letter.", role: "body", bold: true },
    ],
    "Scrap paper bin poster",
    [
      { label: "Slogan", detail: "This IS the bold beginning.", color: C.PRIMARY },
      { label: "One proof", detail: "This IS the mighty middle.", color: C.SECONDARY },
      { label: "Call to action", detail: "This IS the excellent ending.", color: C.ASSESS },
      { label: "Picture", detail: "Does the work of a paragraph.", color: C.ACCENT },
    ],
    NOTES_MODEL, FOOTER,
    { previewSpec: POSTER_MODEL, previewAccent: C.PRIMARY, sourceType: "Model poster" });

  // We Do - compare two posters, same belief
  compareVisualSlide(pres, "We Do", "Same belief. Which one works?",
    "Not which is prettier. Which one works when you walk past it?",
    {
      panelTitle: "Poster A",
      title: "Three minutes a day",
      previewSpec: POSTER_STRONG,
      previewAccent: C.SUCCESS,
      strip: C.SUCCESS,
      previewH: 2.05,
    },
    {
      panelTitle: "Poster B",
      title: "Yard clean up information",
      previewSpec: POSTER_WEAK,
      previewAccent: C.MUTED,
      strip: C.MUTED,
      previewH: 2.05,
    },
    // The posters ARE the hero here, so the compare cards grow to fill the
    // slide rather than leaving the default dead bottom half (section 15h).
    NOTES_COMPARE, FOOTER, { cardH: 3.05 });

  // CFU hinge - which is the slogan?
  (() => {
    const s = P.customSlide(pres, T, "CFU", C.ALERT, "Which one is a slogan?");
    const y0 = CONTENT_TOP;

    addTextOnShape(s, "Poster topic: bring a reusable water bottle", {
      x: 0.5, y: y0, w: 9, h: 0.48, rectRadius: 0.08,
      fill: { color: C.ALERT },
    }, {
      fontSize: 16, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    const options = [
      { key: "A", text: "Students should consider bringing a reusable bottle to school each day." },
      { key: "B", text: "REFILL. DON'T LANDFILL." },
      { key: "C", text: "Water bottles and the environment." },
    ];
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
        addRevealAnswerBar(s, ["B - four words, and a bossy verb"], {
          y: P.REVEAL_Y, h: P.REVEAL_H, fontSize: 22, color: C.SUCCESS,
        });
      },
    ]);

    addFooter(s, FOOTER);
    s.addNotes(NOTES_CFU);
    runSlideDiagnostics(s, pres);
  })();

  // You Do - design your poster
  (() => {
    const s = P.customSlide(pres, T, "You Do", C.ASSESS, "Your poster. Twelve words.");
    const y0 = CONTENT_TOP;

    addTextOnShape(s, "Same belief as your letter. Do not change topic.", {
      x: 0.5, y: y0, w: 9, h: 0.72, rectRadius: 0.08,
      fill: { color: C.ASSESS },
    }, {
      fontSize: 23, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    const parts = [
      { label: "Slogan", detail: "Six words or fewer. Use a bossy verb.", color: C.PRIMARY },
      { label: "One proof", detail: "Your strongest reason, in one line.", color: C.SECONDARY },
      { label: "Call to action", detail: "Exactly what you want them to do.", color: C.SUCCESS },
    ];
    const rh = 0.72;
    const gap = 0.10;
    parts.forEach((part, i) => {
      const ry = y0 + 0.86 + i * (rh + gap);
      s.addShape("roundRect", {
        x: 0.5, y: ry, w: 2.85, h: rh, rectRadius: 0.07,
        fill: { color: part.color },
      });
      s.addText(part.label, {
        x: 0.5, y: ry, w: 2.85, h: rh,
        fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      s.addShape("roundRect", {
        x: 3.50, y: ry, w: 6.0, h: rh, rectRadius: 0.07,
        fill: { color: C.WHITE }, line: { color: part.color, width: 1.3 },
      });
      s.addText(part.detail, {
        x: 3.70, y: ry, w: 5.6, h: rh,
        fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL,
        valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
    });

    s.addText("Words first. Pencils come out when your three parts are written.", {
      x: 0.5, y: y0 + 3.30, w: 9, h: 0.42,
      fontSize: 15, fontFace: FONT_B, color: C.MUTED, italic: true,
      align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
    runSlideDiagnostics(s, pres);
  })();

  P.exitTicketPanel(pres, T, {
    topic: "At the top of your Poster Planner",
    task: "Write your slogan. Six words or fewer.",
    cue: "A bossy verb somewhere in it.",
    footer: FOOTER,
    notes: NOTES_EXIT,
  });

  closingSlide(pres, {
    reflectionPrompt: "Which word in your slogan is doing the most work?",
    scItems: [
      "I can say what a reader notices first on a poster.",
      "I can write a slogan and a call to action for my topic.",
      "I can explain why my poster works at a glance.",
    ],
    selfAssessment: "Fingers at your chest. One, two or three. Show me.",
    takeaways: ["Same belief. Fewer words. Same fight."],
  }, NOTES_CLOSING);

  const pptxPath = path.join(OUT_DIR, "Persuasive Writing W7 S3.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);
}

// ─── Resource: Poster Planner ───────────────────────────────────────────────

async function buildPlanner() {
  const doc = createPdf({ title: POSTER_RES.name });
  let y = addPdfHeader(doc, "Poster Planner", {
    subtitle: "Same belief. About twelve words. Three seconds to land it.",
    color: C.SECONDARY,
    lessonInfo: "Week 7 Session 3 | Year 5/6 Literacy",
  });

  y = addSectionHeading(doc, "Two posters, one belief. Which one works?", y, { color: C.SECONDARY });
  y += 4;

  // Real drawn poster mockups on the printed page, so the paper carries the
  // same visual object students compared on screen (megaprompt scaffold rule).
  y = addPosterPairPdf(doc, y, POSTER_STRONG, POSTER_WEAK, {
    leftTitle: "Poster A", rightTitle: "Poster B",
    posterH: 200, color: C.SECONDARY,
  });
  y += 6;

  y = addBodyText(doc,
    "Poster A works because __________________________________________________",
    y, {});
  y += 4;
  y = addBodyText(doc,
    "Poster B fails because ___________________________________________________",
    y, {});
  y += 10;

  y = addTipBox(doc,
    "The slogan is your bold beginning. One line of proof is your mighty middle. "
    + "The call to action is your excellent ending.",
    y, { color: C.ACCENT });

  doc.addPage();
  y = PAGE.MARGIN;
  y = addSectionHeading(doc, "Your poster. Words first.", y, { color: C.SECONDARY });
  y += 6;

  const parts = [
    { label: "Slogan", hint: "Six words or fewer. Use a bossy verb from your bank.", lines: 1, color: C.PRIMARY },
    { label: "One line of proof", hint: "Your strongest reason, in one short line.", lines: 1, color: C.SECONDARY },
    { label: "Call to action", hint: "Exactly what you want them to do.", lines: 1, color: C.SUCCESS },
  ];

  parts.forEach((part) => {
    doc.save();
    doc.roundedRect(PAGE.MARGIN, y, PAGE.CONTENT_W, 22, 3).fill(hex(part.color));
    doc.fontSize(10).font("Sans-Bold").fillColor("#FFFFFF");
    doc.text(part.label + "   " + part.hint, PAGE.MARGIN + 8, y + 6,
      { width: PAGE.CONTENT_W - 16 });
    doc.restore();
    y += 26;
    y = addLinedArea(doc, y, part.lines, { lineSpacing: 30 });
    y += 12;
  });

  y = addSectionHeading(doc, "Now sketch it. The frame is drawn for you.", y, { color: C.SECONDARY });
  y += 6;

  // A blank poster frame with the three zones marked, so the layout decision
  // is made for the student and only the content is theirs.
  const frameW = 240;
  const frameX = PAGE.MARGIN + (PAGE.CONTENT_W - frameW) / 2;
  const frameH = 300;
  doc.save();
  doc.roundedRect(frameX, y, frameW, frameH, 4).lineWidth(1.2).strokeColor("#000000").stroke();
  const zones = [
    { h: 58, label: "SLOGAN", color: C.PRIMARY },
    { h: 130, label: "PICTURE", color: C.MUTED },
    { h: 54, label: "ONE LINE OF PROOF", color: C.SECONDARY },
    { h: 58, label: "CALL TO ACTION", color: C.SUCCESS },
  ];
  let zy = y;
  zones.forEach((zone) => {
    doc.rect(frameX + 6, zy + 6, frameW - 12, zone.h - 12)
      .lineWidth(0.8).dash(3, { space: 3 }).strokeColor(hex(zone.color)).stroke();
    doc.undash();
    doc.fontSize(8).font("Sans-Bold").fillColor(hex(zone.color));
    doc.text(zone.label, frameX + 10, zy + 10, { width: frameW - 20 });
    zy += zone.h;
  });
  doc.restore();
  y += frameH + 12;

  addPdfFooter(doc, "Persuasive Writing | Week 7 Session 3 | Year 5/6 Literacy");
  await writePdf(doc, path.join(OUT_DIR, POSTER_RES.fileName));
  console.log("Wrote " + POSTER_RES.name);
}

(async () => {
  await build();
  await buildPlanner();
})().catch((err) => {
  console.error(err);
  process.exit(1);
});
