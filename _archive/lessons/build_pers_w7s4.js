"use strict";

// Persuasive Writing - Term 3 Week 7, Session 4 (Year 5/6 Literacy)
// "Piece three: the news article, facts first, then the argument"
//
// Victorian Curriculum 2.0: English, Literacy (Creating texts) and Literature,
// Levels 5-6 - shaping a text for a public audience, and distinguishing fact
// from opinion within one piece.
//
// UNIT ANCHOR (locked): Bold Beginning. Mighty Middle. Excellent Ending.
// Voice all the way through.  Week 7: Same belief. Same reasons. Different voice.
//
// The article is the one format where the belief does NOT come first. Facts
// earn the reader's trust, and only then does the writer argue. That single
// difference is the lesson.
//
// Designed-visual lesson: the article layout is a structured previewSpec via
// the shared drawMockupPreview path, never a flat text description.

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
  titleSlide, liSlide, closingSlide, annotatedModelSlide,
  addCard, addFooter, addTextOnShape, addRevealAnswerBar,
  clickBuild, runSlideDiagnostics,
} = T;

const SESSION = 4;
const FOOTER = "Persuasive Writing | Week 7 Session 4 | Year 5/6 Literacy";
const OUT_DIR = "output/Persuasive_W7_S4";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const ARTICLE_RES = makeSessionResource(SESSION,
  "News Article Builder",
  "The article layout labelled, then a frame for headline, lead, quote and argument.");
const RESOURCE_ITEMS = [ARTICLE_RES];
fs.mkdirSync(RES_DIR, { recursive: true });

// The article layout as a structured spec, so students infer the hierarchy by
// looking rather than reading a description of it.
const ARTICLE_SPEC = {
  accent: C.ASSESS,
  components: [
    { kind: "masthead", text: "DIAMOND CREEK EAST NEWS", scale: 0.85 },
    { kind: "heading", text: "Classrooms throw out paper by the boxful", scale: 1.1 },
    { kind: "caption", text: "By Room 12  |  Term 3", scale: 0.7 },
    { kind: "photo", scale: 1.5 },
    { kind: "textBlock", count: 3, scale: 1.4 },
    { kind: "quote", text: "We fill a recycling box every week, says our cleaner.", scale: 1.0 },
    { kind: "textBlock", count: 3, scale: 1.4 },
  ],
};

// ─── Teacher notes ──────────────────────────────────────────────────────────

const NOTES_TITLE =
  "Week 7, Session 4. One line: the last format, and the only one where you earn the reader's trust before you argue.";

const NOTES_RESOURCES =
  "Prep slide. Print the News Article Builder. Students need their Persuasive Plan, their letter draft and their poster plan.\n" +
  "Have one BTN story or a real school newsletter open. Any short news piece works as the shape reference.\n" +
  "CATCH-UP: missed the week so far? Give a spare Persuasive Plan. A belief and one reason is enough to write the lead and one paragraph today.";

const NOTES_OVERVIEW =
  "For you, not the class. Skip past it when projecting.\n" +
  "The article is the one format where the belief does NOT come first. Facts earn trust, then the argument lands. That single difference is the whole lesson.\n" +
  "Decision points: the fact-or-opinion hinge, the boards check in the We Do, and the exit ticket. Between them keep the pace brisk.";

const NOTES_LAUNCH = composeGlanceNotes({
  answer: "the article starts with facts. The letter and poster start with the side.",
  beats: [
    [
      "SAY: You have written this belief twice now. A letter, then a poster.",
      "SHOW the article layout beside them. SAY: This one is different.",
    ],
    [
      "ASK: What comes first in a news article that did not come first before?",
      "30 sec. Turn and tell. Partner A first.",
      "EXPECT: facts, or what happened.",
    ],
    [
      "SCAN the room as pairs talk.",
      "80%+ saying facts -> next slide.",
      "Less -> read the first line of a real news story aloud, re-ask.",
    ],
  ],
  trap: [
    "thinking a news article cannot persuade at all.",
    "Fix: read the last paragraph of the model, student hears the argument.",
  ],
  prep: "Low-coupling launch: works from the layout on screen alone, so a student who missed the week can still join.",
  tag: "[Launch | Retention and recall | HITS 6]",
});

const NOTES_LI = composeGlanceNotes({
  beats: [
    "POINT to the intention. SAY: Facts first, then the argument.",
    "SAY: Read the I can statements with me. Everyone, together, on three.",
    "SAY: Everyone leaves with a headline and a first line.",
  ],
  prep: "SC1 reachable by all; SC2 is the core target the exit ticket assesses; SC3 stretches into justifying the fact-then-argue order.",
  tag: "[LI/SC | Planning made visible | HITS 1]",
});

const NOTES_MODEL = composeGlanceNotes({
  answer: "headline, byline, photo, facts, a quote, then the argument.",
  beats: [
    [
      "POINT to the layout. SAY: The headline is the biggest thing.",
      "SAY: It says what happened. Not my side.",
    ],
    [
      "SAY: The first paragraph is facts. That is how I earn your trust.",
      "SAY: Only after the quote do I start arguing.",
    ],
    [
      "ASK: Where does my opinion first appear?",
      "15 sec. Fingers at your chest. Point to the screen.",
      "EXPECT: the block after the quote.",
    ],
  ],
  trap: [
    "putting the opinion in the headline.",
    "Fix: read the headline aloud, ask which side it takes, student hears none.",
  ],
  stretch: "Write another headline that is still neutral.",
  help: "Cover all but the headline and first paragraph.",
  prep: "Same belief, same reasons as the letter. Only the ORDER changes. Say that in the anchor's own words.",
  tag: "[I Do | Explicit teaching | SC1 | HITS 3, 4]",
});

const NOTES_QUOTE = composeGlanceNotes({
  answer: "a quote is evidence someone else gave you, in their words.",
  beats: [
    [
      "SAY: A quote is somebody else saying it, not you.",
      "SAY: It is the strongest evidence a young writer can get.",
    ],
    [
      "ASK: Who could you actually quote about your topic?",
      "40 sec. Write it... chin it... show me.",
      "EXPECT: a named real person you could ask this week.",
    ],
    [
      "SCAN boards for a real, reachable person.",
      "80%+ -> cold call two: what would you ask them?",
      "Less -> name three people in this building yourself, re-ask.",
    ],
  ],
  trap: [
    "inventing a quote from a made up expert.",
    "Fix: ask who said it and when, student replaces it with a real person.",
  ],
  prep: "Insist on a real, reachable person. An invented quote is the one thing that would sink a published piece.",
  tag: "[I Do | Explicit teaching | SC2 | HITS 4]",
});

const NOTES_CFU = composeGlanceNotes({
  answer: "B. It says what happened without taking a side.",
  beats: [
    [
      "SAY: Three headlines, same story. Only one belongs on a news article.",
      "Do not call out. Boards for this one.",
    ],
    [
      "ASK: Which is the news headline, A, B or C?",
      "30 sec. Write it... chin it... show me.",
      "EXPECT: B.",
    ],
    [
      "SCAN boards, back row first.",
      "80%+ -> cold call one B: which side does it take? Then reveal.",
      "Less -> read A aloud, ask whose opinion that is, re-ask.",
    ],
    "REVEAL after boards are scanned. SAY: A headline reports. The article argues.",
  ],
  trap: [
    "picking A because it is the most persuasive.",
    "Fix: ask if a reporter could write A, student re-tests all three.",
  ],
  prep: "The hinge. A is an opinion headline, C is vague. Both wrong answers show the fact-then-argue order not yet landing.",
  tag: "[CFU hinge | Supported application | SC2 | HITS 7, 8]",
});

const NOTES_WEDO = composeGlanceNotes({
  answer: "open - a neutral headline, then a first line of facts.",
  beats: [
    "SAY: My headline. Classrooms throw out paper by the boxful. No side.",
    [
      "ASK: Write the first line. Facts only, no opinion.",
      "90 sec. Write it... chin it... show me.",
      "EXPECT: what, where, how much. No should or must.",
    ],
    [
      "SCAN boards for a smuggled opinion.",
      "80%+ -> cold call one strong, one shaky: any opinion there?",
      "Less -> rewrite one board, cut the opinion, re-ask.",
    ],
  ],
  trap: [
    "smuggling should or must into the facts line.",
    "Fix: ring the word, student cuts it and rereads.",
  ],
  stretch: "Add a second facts line with a number.",
  help: "Give the first four words of the facts line.",
  prep: "Hunting for smuggled modality is the move here. It is what makes the fact-then-argue order real rather than decorative.",
  tag: "[We Do | Supported application | SC2 | HITS 4, 5]",
});

const NOTES_YOUDO = composeGlanceNotes({
  answer: "open - headline, facts, a quote, then the argument.",
  beats: [
    [
      "SAY: Your article. Same belief, same reasons, new order.",
      "SAY: Headline, facts, quote, then argue.",
    ],
    [
      "POINT to the Builder. SAY: The quote can be blank if you have not asked yet.",
      "TIME: twenty minutes drafting.",
    ],
    [
      "CIRCULATE. Read the headline first. If it takes a side, send it back.",
      "COLLECT nothing. Drafts are published in Session 5.",
    ],
  ],
  trap: [
    "rewriting the letter with a headline on top.",
    "Fix: ask where the facts are, student moves the argument down.",
  ],
  stretch: "Write the article as if you disagreed with yourself, then compare.",
  help: "Give the Builder with the headline and first facts line written.",
  prep: "Leave the quote blank rather than inventing one. Session 5 has time to go and ask a real person.",
  tag: "[You Do | Mastery and application | SC2 | HITS 10]",
});

const NOTES_EXIT = composeGlanceNotes({
  answer: "a headline that reports, and one line of facts with no opinion in it.",
  beats: [
    [
      "SAY: Two lines at the top of your Builder.",
      "SAY: Your headline, and your first line of facts.",
    ],
    [
      "TIME: four minutes. Builders stay with you.",
      "WALK the rows and read over shoulders.",
    ],
    [
      "SORT in your head: neutral headline, or opinion headline.",
      "Opinion headlines -> two minutes each before Session 5 starts.",
    ],
  ],
  prep: "Assesses the core target. The visible evidence is the absence of a side in the headline.",
  tag: "[Exit Ticket | Mastery and application | SC2 | HITS 8]",
});

const NOTES_CLOSING = composeGlanceNotes({
  answer: "open - most students should show two.",
  beats: [
    "COLD CALL four students for their headline only. Listen for a smuggled side.",
    [
      "ASK: Rate yourself on the three statements.",
      "15 sec. Fingers at your chest. One, two or three. Show me.",
      "EXPECT: most on two.",
    ],
    "SAY: Tomorrow you finish all three and put them where people will read them.",
  ],
  prep: "Note anyone still leading with their opinion. They get two minutes with you before Session 5.",
  tag: "[Closing | Retention and recall | HITS 1, 9]",
});

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  titleSlide(pres, "The News Article",
    "Piece three: facts first, then the argument",
    "Week 7 Session 4 | Year 5/6 Literacy", NOTES_TITLE);

  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  (() => {
    const s = P.customSlide(pres, T, "For the teacher", C.MUTED, "Session 4 at a glance", { badgeW: 2.4 });
    const cardY = CONTENT_TOP;
    const cardH = SAFE_BOTTOM - CONTENT_TOP;

    addCard(s, 0.5, cardY, 4.4, cardH, { strip: C.PRIMARY });
    s.addText("Where this sits", {
      x: 0.7, y: cardY + 0.12, w: 4.0, h: 0.28,
      fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText([
      { text: "Piece three of three, from the same plan.", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: P.ANCHOR_PHRASE_W7, options: { color: C.ASSESS, bold: true, breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "The one real difference:", options: { bold: true, breakLine: true } },
      { text: "in the letter and the poster, the belief comes FIRST.", options: { breakLine: true } },
      { text: "In the article, facts come first and earn the reader's trust. Only then does the writer argue.", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "Same belief. Same reasons. New order.", options: { italic: true, color: C.MUTED } },
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
      { text: "Shape: example-first. The layout is analysed before any writing.", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "The quote rule:", options: { bold: true, color: C.ACCENT, breakLine: true } },
      { text: "a real, reachable person only. Leave the quote blank rather than invent one. There is time in Session 5 to go and ask.", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "Watch for:", options: { bold: true, color: C.ACCENT, breakLine: true } },
      { text: "should and must smuggled into the facts line. Hunting those is what makes the order real.", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "Flex: BTN or the school newsletter both work as the shape reference.", options: { italic: true, color: C.MUTED } },
    ], {
      x: 5.3, y: cardY + 0.46, w: 4.0, h: cardH - 0.60,
      fontSize: 12.5, fontFace: FONT_B, color: C.CHARCOAL,
      valign: "top", margin: 0, paraSpaceAfter: 3,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_OVERVIEW);
    runSlideDiagnostics(s, pres);
  })();

  // Launch - what comes first?
  (() => {
    const s = P.customSlide(pres, T, "Launch", C.SECONDARY, "What comes first this time?");
    const y0 = CONTENT_TOP;

    const rows = [
      { fmt: "Letter", first: "My side, in line one", color: C.PRIMARY },
      { fmt: "Poster", first: "My side, in six words", color: C.SECONDARY },
      { fmt: "News article", first: "Not my side at all", color: C.ASSESS },
    ];
    const rh = 0.86;
    const gap = 0.14;
    rows.forEach((row, i) => {
      const ry = y0 + i * (rh + gap);
      s.addShape("roundRect", {
        x: 0.5, y: ry, w: 3.2, h: rh, rectRadius: 0.07,
        fill: { color: row.color },
      });
      s.addText(row.fmt, {
        x: 0.5, y: ry, w: 3.2, h: rh,
        fontSize: 19, fontFace: FONT_H, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      s.addShape("roundRect", {
        x: 3.85, y: ry, w: 5.65, h: rh, rectRadius: 0.07,
        fill: { color: C.WHITE }, line: { color: row.color, width: 1.3 },
      });
      s.addText(row.first, {
        x: 4.05, y: ry, w: 5.25, h: rh,
        fontSize: 19, fontFace: FONT_B, color: C.CHARCOAL,
        valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
    });

    addTextOnShape(s, "So what DOES come first? Turn and tell. Partner A first.", {
      x: 0.5, y: y0 + 3 * (rh + gap) + 0.06, w: 9, h: 0.72, rectRadius: 0.08,
      fill: { color: C.WHITE }, line: { color: C.SECONDARY, width: 1.5 },
    }, {
      fontSize: 18, fontFace: FONT_B, color: C.SECONDARY, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_LAUNCH);
    runSlideDiagnostics(s, pres);
  })();

  liSlide(pres,
    "We are learning to argue in a news article by reporting the facts first.",
    [
      "I can tell a fact from an opinion in a news article.",
      "I can write a neutral headline and a first line of facts.",
      "I can explain why facts before argument convinces a reader.",
    ],
    NOTES_LI, FOOTER);

  // I Do - the annotated article layout
  annotatedModelSlide(pres, "I Do", "Where does the opinion hide?",
    [
      { text: "Read the layout first", role: "header" },
      "What is the biggest thing?",
      "Which block is facts?",
      "Where does the arguing start?",
      { text: "Same belief. New order.", role: "body", bold: true },
    ],
    "School newsletter article",
    [
      { label: "Headline", detail: "Reports. Takes no side.", color: C.PRIMARY },
      { label: "Facts first", detail: "This earns the reader's trust.", color: C.SECONDARY },
      { label: "Quote", detail: "Someone else says it, not you.", color: C.SUCCESS },
      { label: "Then argue", detail: "The mighty middle, moved down.", color: C.ASSESS },
    ],
    NOTES_MODEL, FOOTER,
    { previewSpec: ARTICLE_SPEC, previewAccent: C.ASSESS, sourceType: "Model article" });

  // I Do - the quote
  (() => {
    const s = P.customSlide(pres, T, "I Do", C.PRIMARY, "Someone else says it");
    const y0 = CONTENT_TOP;

    addCard(s, 0.5, y0, 9, 1.30, { strip: C.SUCCESS });
    s.addText("A quote in my article", {
      x: 0.75, y: y0 + 0.10, w: 5, h: 0.28,
      fontSize: 14, fontFace: FONT_B, color: C.SUCCESS, bold: true, margin: 0,
    });
    s.addText("\"We fill a recycling box every week,\" says our cleaner, Mr Tran.", {
      x: 0.75, y: y0 + 0.44, w: 8.5, h: 0.74,
      fontSize: 23, fontFace: FONT_H, color: C.CHARCOAL, italic: true,
      valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
    });

    s.addText("A quote is the strongest evidence you can get, because it is not you saying it.", {
      x: 0.5, y: y0 + 1.48, w: 9, h: 0.80,
      fontSize: 22, fontFace: FONT_H, color: C.PRIMARY, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    addTextOnShape(s, "Who could YOU actually ask this week?", {
      x: 0.5, y: y0 + 2.42, w: 9, h: 0.68, rectRadius: 0.08,
      fill: { color: C.ASSESS },
    }, {
      fontSize: 21, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    addTextOnShape(s, "A real person. Write it... chin it... show me.", {
      x: 2.4, y: y0 + 3.14, w: 5.2, h: 0.62, rectRadius: 0.08,
      fill: { color: C.WHITE }, line: { color: C.ASSESS, width: 1.5 },
    }, {
      fontSize: 17, fontFace: FONT_B, color: C.ASSESS, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_QUOTE);
    runSlideDiagnostics(s, pres);
  })();

  // CFU hinge - which headline?
  (() => {
    const s = P.customSlide(pres, T, "CFU", C.ALERT, "Which one is a news headline?");
    const y0 = CONTENT_TOP;

    addTextOnShape(s, "Story: the school is running out of bike racks", {
      x: 0.5, y: y0, w: 9, h: 0.48, rectRadius: 0.08,
      fill: { color: C.ALERT },
    }, {
      fontSize: 16, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    const options = [
      { key: "A", text: "We desperately need more bike racks and the school must act now." },
      { key: "B", text: "Bike numbers double as racks fill by 8.40 am" },
      { key: "C", text: "Some thoughts about bikes at our school" },
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
        fontSize: 16.5, fontFace: FONT_B, color: C.CHARCOAL,
        valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
    });

    clickBuild(s, [
      () => {
        addRevealAnswerBar(s, ["B - it reports, and takes no side"], {
          y: P.REVEAL_Y, h: P.REVEAL_H, fontSize: 21, color: C.SUCCESS,
        });
      },
    ]);

    addFooter(s, FOOTER);
    s.addNotes(NOTES_CFU);
    runSlideDiagnostics(s, pres);
  })();

  // We Do
  (() => {
    const s = P.customSlide(pres, T, "We Do", C.SUCCESS, "I write the headline. You write the facts.");
    const y0 = CONTENT_TOP;

    addCard(s, 0.5, y0, 9, 1.05, { strip: C.PRIMARY });
    s.addText("Headline (mine)", {
      x: 0.75, y: y0 + 0.08, w: 4, h: 0.28,
      fontSize: 14, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("Classrooms throw out paper by the boxful", {
      x: 0.75, y: y0 + 0.38, w: 8.5, h: 0.58,
      fontSize: 24, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
      valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
    });

    const fy = y0 + 1.20;
    addCard(s, 0.5, fy, 9, 1.05, { strip: C.SUCCESS });
    s.addText("First line, facts only (yours)", {
      x: 0.75, y: fy + 0.08, w: 5, h: 0.28,
      fontSize: 14, fontFace: FONT_B, color: C.SUCCESS, bold: true, margin: 0,
    });
    s.addText("What, where, how much. No should. No must.", {
      x: 0.75, y: fy + 0.38, w: 8.5, h: 0.58,
      fontSize: 21, fontFace: FONT_B, color: C.CHARCOAL,
      valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
    });

    addTextOnShape(s, "If a should or a must sneaks in, cut it. Write it... chin it... show me.", {
      x: 0.5, y: fy + 1.20, w: 9, h: 0.76, rectRadius: 0.08,
      fill: { color: C.WHITE }, line: { color: C.SUCCESS, width: 1.5 },
    }, {
      fontSize: 18, fontFace: FONT_B, color: C.SUCCESS, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_WEDO);
    runSlideDiagnostics(s, pres);
  })();

  // You Do
  (() => {
    const s = P.customSlide(pres, T, "You Do", C.ASSESS, "Draft your article");
    const y0 = CONTENT_TOP;

    const order = [
      { n: "1", label: "Headline", detail: "Reports the story. Takes no side.", color: C.PRIMARY },
      { n: "2", label: "Facts", detail: "What, where, how much. No opinion yet.", color: C.SECONDARY },
      { n: "3", label: "Quote", detail: "A real person. Leave blank if not asked yet.", color: C.SUCCESS },
      { n: "4", label: "Then argue", detail: "Your strongest reason, from your plan.", color: C.ASSESS },
    ];
    const rh = 0.62;
    const gap = 0.10;
    order.forEach((row, i) => {
      const ry = y0 + i * (rh + gap);
      s.addShape("roundRect", {
        x: 0.5, y: ry, w: 0.62, h: rh, rectRadius: 0.07,
        fill: { color: row.color },
      });
      s.addText(row.n, {
        x: 0.5, y: ry, w: 0.62, h: rh,
        fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      s.addShape("roundRect", {
        x: 1.24, y: ry, w: 2.5, h: rh, rectRadius: 0.07,
        fill: { color: C.WHITE }, line: { color: row.color, width: 1.3 },
      });
      s.addText(row.label, {
        x: 1.24, y: ry, w: 2.5, h: rh,
        fontSize: 16, fontFace: FONT_H, color: row.color, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      s.addText(row.detail, {
        x: 3.95, y: ry, w: 5.55, h: rh,
        fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL,
        valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
    });

    addTextOnShape(s, "News Article Builder. Twenty minutes. Same belief, new order.", {
      x: 0.5, y: y0 + 4 * (rh + gap) + 0.06, w: 9, h: 0.72, rectRadius: 0.08,
      fill: { color: C.ASSESS },
    }, {
      fontSize: 19, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
    runSlideDiagnostics(s, pres);
  })();

  P.exitTicketPanel(pres, T, {
    topic: "At the top of your News Article Builder",
    task: "Your headline, and your first line of facts.",
    cue: "No side in the headline. No should. No must.",
    footer: FOOTER,
    notes: NOTES_EXIT,
  });

  closingSlide(pres, {
    reflectionPrompt: "Which of your three pieces do you think would change the most minds?",
    scItems: [
      "I can tell a fact from an opinion in a news article.",
      "I can write a neutral headline and a first line of facts.",
      "I can explain why facts before argument convinces a reader.",
    ],
    selfAssessment: "Fingers at your chest. One, two or three. Show me.",
    takeaways: ["Facts earn the trust. Then you argue."],
  }, NOTES_CLOSING);

  const pptxPath = path.join(OUT_DIR, "Persuasive Writing W7 S4.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);
}

// ─── Resource: News Article Builder ─────────────────────────────────────────

async function buildArticle() {
  const doc = createPdf({ title: ARTICLE_RES.name });
  let y = addPdfHeader(doc, "News Article Builder", {
    subtitle: "Same belief. Same reasons. Facts first, then the argument.",
    color: C.ASSESS,
    lessonInfo: "Week 7 Session 4 | Year 5/6 Literacy",
  });

  y = addSectionHeading(doc, "The layout, labelled", y, { color: C.ASSESS });
  y += 6;

  // A drawn article wireframe, so the printed scaffold carries the same
  // visual hierarchy students analysed on screen.
  const boxX = PAGE.MARGIN;
  const boxW = PAGE.CONTENT_W;
  const zones = [
    { h: 20, label: "MASTHEAD", note: "The paper's name", color: C.MUTED },
    { h: 34, label: "HEADLINE", note: "Reports the story. Takes NO side.", color: C.PRIMARY },
    { h: 16, label: "BYLINE", note: "Who wrote it, and when", color: C.MUTED },
    { h: 60, label: "PHOTO", note: "Shows what is happening", color: C.MUTED },
    { h: 44, label: "FACTS", note: "What, where, how much. This earns trust.", color: C.SECONDARY },
    { h: 30, label: "QUOTE", note: "Someone else says it, in their words", color: C.SUCCESS },
    { h: 52, label: "THE ARGUMENT", note: "Now your mighty middle, moved down here", color: C.ASSESS },
  ];
  zones.forEach((zone) => {
    doc.save();
    doc.rect(boxX, y, 96, zone.h).fill(hex(zone.color));
    doc.fontSize(8).font("Sans-Bold").fillColor("#FFFFFF");
    doc.text(zone.label, boxX + 4, y + zone.h / 2 - 4, { width: 88, align: "center" });
    doc.rect(boxX + 96, y, boxW - 96, zone.h)
      .lineWidth(0.7).strokeColor("#9CA3AF").stroke();
    doc.fontSize(9).font("Sans-Italic").fillColor("#000000");
    doc.text(zone.note, boxX + 104, y + zone.h / 2 - 5, { width: boxW - 112 });
    doc.restore();
    y += zone.h;
  });

  y += 12;
  y = addTipBox(doc,
    "The headline is the only part that must take NO side. Everything below it is still "
    + "persuasive writing, just in a new order.",
    y, { color: C.ACCENT });

  doc.addPage();
  y = PAGE.MARGIN;
  y = addSectionHeading(doc, "Your article. Draft it here.", y, { color: C.ASSESS });
  y += 6;

  const frames = [
    { label: "Headline", starter: "Reports the story. No side.", lines: 1, color: C.PRIMARY },
    { label: "First paragraph: facts", starter: "What, where, how much. No should, no must.", lines: 3, color: C.SECONDARY },
    { label: "Quote", starter: "\"...,\" says [a real person you asked].", lines: 2, color: C.SUCCESS },
    { label: "The argument", starter: "Reason. Evidence. Explain. Link.", lines: 5, color: C.ASSESS },
  ];

  frames.forEach((frame) => {
    doc.save();
    doc.roundedRect(PAGE.MARGIN, y, PAGE.CONTENT_W, 22, 3).fill(hex(frame.color));
    doc.fontSize(10).font("Sans-Bold").fillColor("#FFFFFF");
    doc.text(frame.label + "   " + frame.starter, PAGE.MARGIN + 8, y + 6,
      { width: PAGE.CONTENT_W - 16 });
    doc.restore();
    y += 26;
    y = addLinedArea(doc, y, frame.lines, { lineSpacing: 28 });
    y += 12;
  });

  y = addTipBox(doc,
    "Who I will ask for my quote: ______________________________   "
    + "What I will ask them: ______________________________",
    y, { color: C.SUCCESS });

  addPdfFooter(doc, "Persuasive Writing | Week 7 Session 4 | Year 5/6 Literacy");
  await writePdf(doc, path.join(OUT_DIR, ARTICLE_RES.fileName));
  console.log("Wrote " + ARTICLE_RES.name);
}

(async () => {
  await build();
  await buildArticle();
})().catch((err) => {
  console.error(err);
  process.exit(1);
});
