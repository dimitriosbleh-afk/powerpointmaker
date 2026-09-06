"use strict";

// Persuasive Writing - Term 3 Week 7, Session 5 (Year 5/6 Literacy)
// "Publish: three pieces, one belief, out into the world"
//
// Victorian Curriculum 2.0: English, Literacy (Creating texts), Levels 5-6 -
// editing and publishing texts for a real audience, and reflecting on how
// language choices suit purpose and audience.
//
// UNIT ANCHOR (locked): Bold Beginning. Mighty Middle. Excellent Ending.
// Voice all the way through.  Week 7: Same belief. Same reasons. Different voice.
//
// The whole fortnight lands here. Students edit against the anchor, give each
// other one specific piece of feedback, publish all three pieces, and read the
// three side by side to see what changed and what did not.

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme } = require("../themes/factory");
const { composeGlanceNotes } = require("../themes/core/composeNotes");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addLinedArea,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
  PAGE, hex,
} = require("../themes/pdf_helpers");

const P = require("./persuasive_lib");

const UNIT_VARIANT = 5;
const T = createTheme("literacy", "grade56", UNIT_VARIANT);
const {
  C, FONT_H, FONT_B, SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, closingSlide,
  addCard, addFooter, addTextOnShape, addRevealAnswerBar,
  clickBuild, runSlideDiagnostics,
} = T;

const SESSION = 5;
const FOOTER = "Persuasive Writing | Week 7 Session 5 | Year 5/6 Literacy";
const OUT_DIR = "output/Persuasive_W7_S5";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const PUBLISH_RES = makeSessionResource(SESSION,
  "Publishing Checklist",
  "Check each piece against the anchor, then give and record one piece of partner feedback.");
const RESOURCE_ITEMS = [PUBLISH_RES];
fs.mkdirSync(RES_DIR, { recursive: true });

// The four checks, held identical on screen and on the printed checklist.
const CHECKS = [
  { label: "Bold Beginning", ask: "Can a stranger tell my side from the opening?", color: "PRIMARY" },
  { label: "Mighty Middle", ask: "Does every reason have proof behind it?", color: "SECONDARY" },
  { label: "Excellent Ending", ask: "Does it say exactly what I want them to do?", color: "ASSESS" },
  { label: "Voice", ask: "Do I sound certain? Emotive words, high modality?", color: "ACCENT" },
];

// ─── Teacher notes ──────────────────────────────────────────────────────────

const NOTES_TITLE =
  "Week 7, Session 5. One line: today the three pieces stop being drafts and go somewhere people will read them.";

const NOTES_RESOURCES =
  "Prep slide. Print the Publishing Checklist. Students need all three drafts, their plan and their word bank.\n" +
  "Decide before the lesson where each piece will actually go: letters delivered or emailed, posters on a named wall, articles in the newsletter or a class display.\n" +
  "CATCH-UP: only one or two drafts finished? Publish what exists. One finished piece properly published beats three unfinished ones.";

const NOTES_OVERVIEW =
  "For you, not the class. Skip past it when projecting.\n" +
  "The fortnight lands here. Edit against the anchor, one specific piece of partner feedback each, then publish. Real destinations matter more than neat handwriting.\n" +
  "Decision points: the feedback quality check after the We Do, and the three-way comparison at the end. Between them, students work and you conference.";

const NOTES_LAUNCH = composeGlanceNotes({
  answer: "the belief and the reasons stayed. The words and the order changed.",
  beats: [
    [
      "SAY: Put all three pieces on your desk. Side by side. Look at them.",
      "SAY: One belief. Three readers. Three very different pieces of writing.",
    ],
    [
      "ASK: What stayed the same across all three?",
      "40 sec. Turn and tell. Partner B first.",
      "EXPECT: the belief and the reasons.",
    ],
    [
      "SCAN the room as pairs talk.",
      "80%+ naming the belief -> next slide.",
      "Less -> hold up one student's letter and poster, ask what is shared, re-ask.",
    ],
  ],
  trap: [
    "thinking they wrote three unrelated things.",
    "Fix: read one student's belief off all three, student hears it repeat.",
  ],
  prep: "This is the payoff moment of the fortnight. Do not rush it. Whole block under 8 minutes.",
  tag: "[Launch | Retention and recall | HITS 6, 9]",
});

const NOTES_LI = composeGlanceNotes({
  beats: [
    "POINT to the intention. SAY: Today we make them good enough to send.",
    "SAY: Read the I can statements with me. Everyone, together, on three.",
    "SAY: Everyone leaves with at least one piece properly published.",
  ],
  prep: "SC1 reachable by all; SC2 is the core target the exit ticket assesses; SC3 stretches into explaining audience choices.",
  tag: "[LI/SC | Planning made visible | HITS 1]",
});

const NOTES_CHECKS = composeGlanceNotes({
  answer: "four checks, one per band of the anchor.",
  beats: [
    [
      "SAY: We do not edit for spelling first. We edit for whether it works.",
      "CLICK through the four checks, in anchor order.",
    ],
    [
      "SAY: Watch me run check one on my letter.",
      "SAY: Cover all but my first line. Can you tell my side?",
    ],
    [
      "ASK: Run check one on YOUR letter. Does it pass?",
      "45 sec. Thumbs only, voices off. Show me... now.",
      "EXPECT: a mix. That is honest.",
    ],
  ],
  trap: [
    "editing spelling and handwriting instead of the argument.",
    "Fix: cover the whole page but one line, student judges that line only.",
  ],
  stretch: "Run all four checks on your article.",
  help: "Do check one together, out loud, on their letter.",
  prep: "Same four checks as the printed checklist and the same wording as the anchor. No new editing language today.",
  tag: "[I Do | Explicit teaching | SC2 | HITS 4, 8]",
});

const NOTES_FEEDBACK = composeGlanceNotes({
  answer: "one specific thing, named against one of the four checks.",
  beats: [
    [
      "SAY: Feedback that helps names ONE check and ONE line.",
      "SAY: Not good job. Not I liked it. One check, one line.",
    ],
    [
      "ASK: Which of these two comments would actually help you?",
      "20 sec. Fingers at your chest. One or two.",
      "EXPECT: two.",
    ],
    [
      "SCAN the fingers.",
      "80%+ on two -> cold call one: what made it useful?",
      "Less -> read comment one aloud, ask what to change, re-ask.",
    ],
  ],
  trap: [
    "giving praise instead of a change.",
    "Fix: ask what the partner should DO next, student rewrites the comment.",
  ],
  prep: "Model the sentence stem before pairs start. Vague feedback wastes the whole partner block.",
  tag: "[We Do | Supported application | SC2 | HITS 5, 8]",
});

const NOTES_YOUDO = composeGlanceNotes({
  answer: "open - each piece checked, one partner comment acted on, then published.",
  beats: [
    "SAY: Four checks on each piece. Then swap. One comment each.",
    [
      "POINT to the destinations. SAY: These are real. Someone will read them.",
      "TIME: fifteen minutes editing, ten swapping, then publish.",
    ],
    [
      "CIRCULATE. Conference with the students you flagged this week.",
      "COLLECT the published pieces at the end, not the drafts.",
    ],
  ],
  trap: [
    "starting a piece again from scratch instead of editing it.",
    "Fix: ring one line to change, student edits only that.",
  ],
  stretch: "Write a short note to your reader saying why you wrote three versions.",
  help: "Publish one piece properly rather than three in a hurry.",
  prep: "Real destinations are what make the fortnight matter. Agree them before the lesson, not during it.",
  tag: "[You Do | Mastery and application | SC2 | HITS 10]",
});

const NOTES_SHARE = composeGlanceNotes({
  answer: "open - listen for the same belief in three different voices.",
  beats: [
    [
      "SAY: Four students. Each reads their slogan, then their letter's first line.",
      "SAY: Everyone else listens for the belief underneath both.",
    ],
    [
      "ASK: What was the belief? Say it back in your own words.",
      "15 sec. Choral response: Everyone, together, on three.",
      "EXPECT: the belief, roughly right.",
    ],
    "SAY: Same belief. Same reasons. Different voice. You did that four times over.",
  ],
  trap: [
    "reading whole pieces, so the block runs long.",
    "Fix: two lines each only, timed. Move on.",
  ],
  prep: "Two lines each keeps it sharp and makes the contrast between formats audible. Do not let it become a full reading.",
  tag: "[Share | Mastery and application | HITS 5, 9]",
});

const NOTES_EXIT = composeGlanceNotes({
  answer: "any honest answer naming the audience as the reason.",
  beats: [
    [
      "SAY: One sentence on the back of your checklist.",
      "SAY: Which piece would change the most minds, and why?",
    ],
    [
      "TIME: four minutes.",
      "COLLECT the checklists. These ones you keep.",
    ],
    [
      "SORT as you collect: reason names the audience, or reason is about effort.",
      "Effort answers -> the audience idea has not landed. Note it for next unit.",
    ],
  ],
  prep: "Assesses whether audience, not effort, is what students think decides persuasive power. Keep these for reporting.",
  tag: "[Exit Ticket | Mastery and application | SC3 | HITS 8, 9]",
});

const NOTES_CLOSING = composeGlanceNotes({
  answer: "open - most students should show two or three.",
  beats: [
    "SAY: Two weeks ago some of us were not sure what persuasive writing even was.",
    [
      "ASK: Rate yourself on the three statements.",
      "15 sec. Fingers at your chest. One, two or three. Show me.",
      "EXPECT: most on two or three.",
    ],
    "SAY: Say the anchor with me. Bold Beginning. Mighty Middle. Excellent Ending.",
  ],
  prep: "Last session of the fortnight. Finishing on the anchor phrase is what makes it stick for the next unit.",
  tag: "[Closing | Retention and recall | HITS 1, 9]",
});

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  titleSlide(pres, "Publish It",
    "Three pieces. One belief. Out into the world.",
    "Week 7 Session 5 | Year 5/6 Literacy", NOTES_TITLE);

  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  (() => {
    const s = P.customSlide(pres, T, "For the teacher", C.MUTED, "Session 5 at a glance", { badgeW: 2.4 });
    const cardY = CONTENT_TOP;
    const cardH = SAFE_BOTTOM - CONTENT_TOP;

    addCard(s, 0.5, cardY, 4.4, cardH, { strip: C.PRIMARY });
    s.addText("Where this sits", {
      x: 0.7, y: cardY + 0.12, w: 4.0, h: 0.28,
      fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText([
      { text: "The last session of the fortnight.", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: P.ANCHOR_PHRASE, options: { color: C.PRIMARY, bold: true, breakLine: true } },
      { text: P.ANCHOR_PHRASE_W7, options: { color: C.ASSESS, bold: true, breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "Editing is against the anchor, not for spelling. Four checks, one per band, in the same words students have heard for two weeks.", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "One piece properly published beats three in a hurry.", options: { italic: true, color: C.MUTED } },
    ], {
      x: 0.7, y: cardY + 0.46, w: 4.0, h: cardH - 0.60,
      fontSize: 12.5, fontFace: FONT_B, color: C.CHARCOAL,
      valign: "top", margin: 0, paraSpaceAfter: 3,
    });

    addCard(s, 5.1, cardY, 4.4, cardH, { strip: C.ACCENT });
    s.addText("Before the lesson", {
      x: 5.3, y: cardY + 0.12, w: 4.0, h: 0.28,
      fontSize: 14, fontFace: FONT_H, color: C.ACCENT, bold: true, margin: 0,
    });
    s.addText([
      { text: "Agree the real destinations:", options: { bold: true, color: C.ACCENT, breakLine: true } },
      { text: "letters delivered or emailed to the named reader", options: { breakLine: true } },
      { text: "posters up on a named wall or corridor", options: { breakLine: true } },
      { text: "articles into the newsletter or a class display", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "Real destinations are what make the fortnight matter. Agree them now, not during the lesson.", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "Use the You Do block to conference with the students you flagged this week.", options: { italic: true, color: C.MUTED } },
    ], {
      x: 5.3, y: cardY + 0.46, w: 4.0, h: cardH - 0.60,
      fontSize: 12.5, fontFace: FONT_B, color: C.CHARCOAL,
      valign: "top", margin: 0, paraSpaceAfter: 3,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_OVERVIEW);
    runSlideDiagnostics(s, pres);
  })();

  // Launch - three pieces side by side
  (() => {
    const s = P.customSlide(pres, T, "Launch", C.SECONDARY, "All three, side by side");
    const y0 = CONTENT_TOP;

    const pieces = [
      { label: "Letter", color: C.PRIMARY },
      { label: "Poster", color: C.SECONDARY },
      { label: "News article", color: C.ASSESS },
    ];
    const bw = 2.87;
    const gap = 0.20;
    pieces.forEach((piece, i) => {
      const bx = 0.5 + i * (bw + gap);
      s.addShape("roundRect", {
        x: bx, y: y0, w: bw, h: 1.30, rectRadius: 0.08,
        fill: { color: piece.color },
      });
      s.addText(piece.label, {
        x: bx, y: y0, w: bw, h: 1.30,
        fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
    });

    s.addText("What stayed the same in all three?", {
      x: 0.5, y: y0 + 1.52, w: 9, h: 1.10,
      fontSize: 38, fontFace: FONT_H, color: C.PRIMARY, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    addTextOnShape(s, "Turn and tell. Partner B first.", {
      x: 2.95, y: y0 + 2.78, w: 4.1, h: 0.64, rectRadius: 0.08,
      fill: { color: C.WHITE }, line: { color: C.SECONDARY, width: 1.5 },
    }, {
      fontSize: 17, fontFace: FONT_B, color: C.SECONDARY, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_LAUNCH);
    runSlideDiagnostics(s, pres);
  })();

  liSlide(pres,
    "We are learning to edit our writing against what it is meant to do, then publish it.",
    [
      "I can check my writing against the three parts and voice.",
      "I can give my partner one specific piece of feedback.",
      "I can explain why my three pieces sound different.",
    ],
    NOTES_LI, FOOTER);

  // I Do - the four checks
  (() => {
    const s = P.customSlide(pres, T, "I Do", C.PRIMARY, "Edit for whether it WORKS");
    const y0 = CONTENT_TOP;
    const rh = 0.70;
    const gap = 0.10;

    clickBuild(s, CHECKS.map((check, i) => () => {
      const ry = y0 + i * (rh + gap);
      s.addShape("roundRect", {
        x: 0.5, y: ry, w: 2.85, h: rh, rectRadius: 0.07,
        fill: { color: C[check.color] },
      });
      s.addText(check.label, {
        x: 0.5, y: ry, w: 2.85, h: rh,
        fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
      s.addShape("roundRect", {
        x: 3.50, y: ry, w: 6.0, h: rh, rectRadius: 0.07,
        fill: { color: C.WHITE }, line: { color: C[check.color], width: 1.3 },
      });
      s.addText(check.ask, {
        x: 3.70, y: ry, w: 5.6, h: rh,
        fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL,
        valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
    }));

    s.addText("Spelling comes last. This comes first.", {
      x: 0.5, y: y0 + 4 * (rh + gap) + 0.04, w: 9, h: 0.44,
      fontSize: 17, fontFace: FONT_B, color: C.MUTED, italic: true,
      align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_CHECKS);
    runSlideDiagnostics(s, pres);
  })();

  // We Do - what feedback actually helps
  (() => {
    const s = P.customSlide(pres, T, "We Do", C.SUCCESS, "Which comment would help you?");
    const y0 = CONTENT_TOP;

    const comments = [
      { key: "1", text: "This is really good, I liked reading it.", color: C.MUTED },
      { key: "2", text: "Excellent Ending check: your last line does not say what you want them to DO.", color: C.SUCCESS },
    ];
    const ch = 1.05;
    comments.forEach((comment, i) => {
      const cy = y0 + i * (ch + 0.18);
      s.addShape("roundRect", {
        x: 0.5, y: cy, w: 9, h: ch, rectRadius: 0.08,
        fill: { color: C.WHITE }, line: { color: comment.color, width: 1.5 },
      });
      s.addShape("roundRect", {
        x: 0.72, y: cy + 0.28, w: 0.5, h: 0.5, rectRadius: 0.25,
        fill: { color: comment.color },
      });
      s.addText(comment.key, {
        x: 0.72, y: cy + 0.28, w: 0.5, h: 0.5,
        fontSize: 19, fontFace: FONT_H, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      s.addText(comment.text, {
        x: 1.42, y: cy, w: 7.85, h: ch,
        fontSize: 19, fontFace: FONT_B, color: C.CHARCOAL,
        valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
    });

    addTextOnShape(s, "Name ONE check. Name ONE line. That is all.", {
      x: 0.5, y: y0 + 2.52, w: 9, h: 0.76, rectRadius: 0.08,
      fill: { color: C.SUCCESS },
    }, {
      fontSize: 21, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    s.addText("Fingers at your chest. One or two. Show me.", {
      x: 0.5, y: y0 + 3.38, w: 9, h: 0.42,
      fontSize: 15, fontFace: FONT_B, color: C.MUTED, italic: true,
      align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_FEEDBACK);
    runSlideDiagnostics(s, pres);
  })();

  // You Do - edit, swap, publish
  (() => {
    const s = P.customSlide(pres, T, "You Do", C.ASSESS, "Edit. Swap. Publish.");
    const y0 = CONTENT_TOP;

    const steps = [
      { n: "1", t: "Four checks on each piece.", time: "15 min", color: C.PRIMARY },
      { n: "2", t: "Swap. One comment each, naming a check.", time: "10 min", color: C.SECONDARY },
      { n: "3", t: "Fix that one thing, then publish.", time: "rest", color: C.SUCCESS },
    ];
    const rh = 0.82;
    const gap = 0.12;
    steps.forEach((step, i) => {
      const ry = y0 + i * (rh + gap);
      s.addShape("roundRect", {
        x: 0.5, y: ry, w: 0.68, h: rh, rectRadius: 0.07,
        fill: { color: step.color },
      });
      s.addText(step.n, {
        x: 0.5, y: ry, w: 0.68, h: rh,
        fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      s.addShape("roundRect", {
        x: 1.30, y: ry, w: 6.85, h: rh, rectRadius: 0.07,
        fill: { color: C.WHITE }, line: { color: step.color, width: 1.3 },
      });
      s.addText(step.t, {
        x: 1.52, y: ry, w: 6.41, h: rh,
        fontSize: 18, fontFace: FONT_B, color: C.CHARCOAL,
        valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
      s.addShape("roundRect", {
        x: 8.27, y: ry, w: 1.23, h: rh, rectRadius: 0.07,
        fill: { color: C.BG_LIGHT }, line: { color: step.color, width: 1 },
      });
      s.addText(step.time, {
        x: 8.27, y: ry, w: 1.23, h: rh,
        fontSize: 15, fontFace: FONT_B, color: step.color, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
    });

    addTextOnShape(s, "These go somewhere real. Someone will read them.", {
      x: 0.5, y: y0 + 3 * (rh + gap) + 0.08, w: 9, h: 0.76, rectRadius: 0.08,
      fill: { color: C.ASSESS },
    }, {
      fontSize: 21, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
    runSlideDiagnostics(s, pres);
  })();

  // Share
  (() => {
    const s = P.customSlide(pres, T, "Share", C.SECONDARY, "Two lines each");
    const y0 = CONTENT_TOP;

    addCard(s, 0.5, y0, 9, 1.55, { strip: C.SECONDARY });
    s.addText("Read out", {
      x: 0.75, y: y0 + 0.10, w: 4, h: 0.28,
      fontSize: 14, fontFace: FONT_B, color: C.SECONDARY, bold: true, margin: 0,
    });
    s.addText([
      { text: "Your poster slogan.", options: { bold: true, color: C.CHARCOAL, breakLine: true } },
      { text: "Then the first line of your letter.", options: { bold: true, color: C.CHARCOAL } },
    ], {
      x: 0.75, y: y0 + 0.42, w: 8.5, h: 1.02,
      fontSize: 24, fontFace: FONT_B, valign: "middle", margin: 0,
      paraSpaceAfter: 4, fit: "shrink", shrinkText: true,
    });

    s.addText("Everyone else: listen for the belief underneath both.", {
      x: 0.5, y: y0 + 1.74, w: 9, h: 0.90,
      fontSize: 26, fontFace: FONT_H, color: C.PRIMARY, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    addTextOnShape(s, P.ANCHOR_PHRASE_W7, {
      x: 0.5, y: y0 + 2.78, w: 9, h: 0.74, rectRadius: 0.08,
      fill: { color: C.ACCENT },
    }, {
      fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_SHARE);
    runSlideDiagnostics(s, pres);
  })();

  P.exitTicketPanel(pres, T, {
    topic: "On the back of your Publishing Checklist",
    task: "Which of your three pieces would change the most minds, and why?",
    cue: "One sentence. Think about who reads it.",
    taskSize: 26,
    footer: FOOTER,
    notes: NOTES_EXIT,
  });

  closingSlide(pres, {
    reflectionPrompt: "What is the one thing you will take into the next piece you write?",
    scItems: [
      "I can check my writing against the three parts and voice.",
      "I can give my partner one specific piece of feedback.",
      "I can explain why my three pieces sound different.",
    ],
    selfAssessment: "Fingers at your chest. One, two or three. Show me.",
    takeaways: [P.ANCHOR_PHRASE],
  }, NOTES_CLOSING);

  const pptxPath = path.join(OUT_DIR, "Persuasive Writing W7 S5.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);
}

// ─── Resource: Publishing Checklist ─────────────────────────────────────────

async function buildChecklist() {
  const doc = createPdf({ title: PUBLISH_RES.name });
  let y = addPdfHeader(doc, "Publishing Checklist", {
    subtitle: "Check what it DOES before you check the spelling.",
    color: C.SUCCESS,
    lessonInfo: "Week 7 Session 5 | Year 5/6 Literacy",
  });

  y = addTipBox(doc, P.ANCHOR_PHRASE, y, { color: C.ACCENT });
  y += 6;

  y = addSectionHeading(doc, "Four checks, on each piece", y, { color: C.SUCCESS });
  y += 6;

  // A drawn check grid: four checks down, three pieces across, so the same
  // four questions are asked of every piece.
  const labelW = 190;
  const colW = (PAGE.CONTENT_W - labelW) / 3;
  const hdrH = 26;

  doc.save();
  doc.rect(PAGE.MARGIN, y, labelW, hdrH).fill(hex(C.PRIMARY));
  doc.fontSize(10).font("Sans-Bold").fillColor("#FFFFFF");
  doc.text("Ask this of every piece", PAGE.MARGIN + 6, y + 8, { width: labelW - 12 });
  ["Letter", "Poster", "Article"].forEach((piece, i) => {
    const cx = PAGE.MARGIN + labelW + i * colW;
    doc.rect(cx, y, colW, hdrH).fill(hex(C.SECONDARY));
    doc.fontSize(10).font("Sans-Bold").fillColor("#FFFFFF");
    doc.text(piece, cx, y + 8, { width: colW, align: "center" });
  });
  doc.restore();
  y += hdrH;

  const rowH = 52;
  CHECKS.forEach((check) => {
    doc.save();
    doc.rect(PAGE.MARGIN, y, labelW, rowH)
      .lineWidth(0.9).strokeColor("#000000").stroke();
    doc.fontSize(9.5).font("Sans-Bold").fillColor(hex(C[check.color]));
    doc.text(check.label, PAGE.MARGIN + 6, y + 6, { width: labelW - 12 });
    doc.fontSize(9).font("Sans").fillColor("#000000");
    doc.text(check.ask, PAGE.MARGIN + 6, y + 20, { width: labelW - 12 });
    for (let i = 0; i < 3; i += 1) {
      const cx = PAGE.MARGIN + labelW + i * colW;
      doc.rect(cx, y, colW, rowH).lineWidth(0.9).strokeColor("#000000").stroke();
      // A tick box big enough for a Year 6 hand, plus a fix line beside it.
      doc.rect(cx + 8, y + 8, 18, 18).lineWidth(1).strokeColor("#000000").stroke();
      doc.moveTo(cx + 8, y + 40).lineTo(cx + colW - 8, y + 40)
        .lineWidth(0.6).strokeColor("#000000").stroke();
      doc.fontSize(7).font("Sans-Italic").fillColor("#6B7280");
      doc.text("what I fixed", cx + 8, y + 42, { width: colW - 16 });
    }
    doc.restore();
    y += rowH;
  });

  y += 16;
  y = addSectionHeading(doc, "Partner feedback: one check, one line", y, { color: C.SECONDARY });
  y += 4;

  y = addTipBox(doc,
    "Use this stem: \"[Check name] check: your [which line] does not yet [what it should do].\" "
    + "Never good job. Never I liked it.",
    y, { color: C.SECONDARY });
  y += 4;

  doc.fontSize(11).font("Sans-Bold").fillColor("#000000");
  doc.text("My partner:", PAGE.MARGIN, y);
  doc.moveTo(PAGE.MARGIN + 76, y + 14).lineTo(PAGE.MARGIN + PAGE.CONTENT_W, y + 14)
    .lineWidth(0.9).strokeColor("#000000").stroke();
  y += 28;

  doc.fontSize(10).font("Sans-Bold").fillColor(hex(C.SECONDARY));
  doc.text("The comment I gave them:", PAGE.MARGIN, y);
  y += 16;
  y = addLinedArea(doc, y, 2, { lineSpacing: 28 });
  y += 12;

  doc.fontSize(10).font("Sans-Bold").fillColor(hex(C.SUCCESS));
  doc.text("The comment they gave me, and what I changed because of it:", PAGE.MARGIN, y);
  y += 16;
  y = addLinedArea(doc, y, 2, { lineSpacing: 28 });

  doc.addPage();
  y = PAGE.MARGIN;
  y = addSectionHeading(doc, "Where each piece is going", y, { color: C.ASSESS });
  y += 6;
  y = addBodyText(doc,
    "A persuasive piece nobody reads persuades nobody. Write down where each of yours is "
    + "actually going, and who will read it.",
    y, {});
  y += 8;

  [
    { piece: "My letter", color: C.PRIMARY },
    { piece: "My poster", color: C.SECONDARY },
    { piece: "My news article", color: C.ASSESS },
  ].forEach((item) => {
    doc.save();
    doc.roundedRect(PAGE.MARGIN, y, PAGE.CONTENT_W, 22, 3).fill(hex(item.color));
    doc.fontSize(10).font("Sans-Bold").fillColor("#FFFFFF");
    doc.text(item.piece + "   goes to ... and will be read by ...", PAGE.MARGIN + 8, y + 6,
      { width: PAGE.CONTENT_W - 16 });
    doc.restore();
    y += 26;
    y = addLinedArea(doc, y, 2, { lineSpacing: 28 });
    y += 14;
  });

  y = addSectionHeading(doc, "Last question", y, { color: C.SUCCESS });
  y += 4;
  doc.fontSize(11).font("Sans-Bold").fillColor("#000000");
  doc.text("Which of your three pieces would change the most minds, and why?",
    PAGE.MARGIN, y, { width: PAGE.CONTENT_W });
  y += 20;
  y = addLinedArea(doc, y, 3, { lineSpacing: 28 });

  addPdfFooter(doc, "Persuasive Writing | Week 7 Session 5 | Year 5/6 Literacy");
  await writePdf(doc, path.join(OUT_DIR, PUBLISH_RES.fileName));
  console.log("Wrote " + PUBLISH_RES.name);
}

(async () => {
  await build();
  await buildChecklist();
})().catch((err) => {
  console.error(err);
  process.exit(1);
});
