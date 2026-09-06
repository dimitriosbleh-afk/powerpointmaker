"use strict";

// Persuasive Writing - Term 3 Week 7, Session 2 (Year 5/6 Literacy)
// "Piece one: the letter to someone who can say yes"
//
// Victorian Curriculum 2.0: English, Literacy (Creating texts), Levels 5-6 -
// creating a sustained persuasive text for a specific audience and purpose.
//
// UNIT ANCHOR (locked): Bold Beginning. Mighty Middle. Excellent Ending.
// Voice all the way through.  Week 7: Same belief. Same reasons. Different voice.
//
// The letter is the full-structure piece: every band of the anchor appears in
// it. The poster and the article that follow are compressions of this.

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

const SESSION = 2;
const FOOTER = "Persuasive Writing | Week 7 Session 2 | Year 5/6 Literacy";
const OUT_DIR = "output/Persuasive_W7_S2";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const LETTER_RES = makeSessionResource(SESSION,
  "Letter Builder",
  "The model letter labelled by band, then a drafting frame for your own.");
const RESOURCE_ITEMS = [LETTER_RES];
fs.mkdirSync(RES_DIR, { recursive: true });

// Wireframe of a persuasive letter. Structured previewSpec, not flat text, so
// students infer the shape by looking (megaprompt previewSpec consistency rule).
const LETTER_SPEC = {
  accent: C.PRIMARY,
  components: [
    { kind: "caption", text: "Your address, then the date", scale: 0.75 },
    { kind: "subheading", text: "Dear Mrs Okafor,", scale: 0.85 },
    { kind: "heading", text: "I am writing because every classroom needs a scrap paper bin.", scale: 1.0 },
    { kind: "textBlock", count: 3, scale: 1.5 },
    { kind: "textBlock", count: 3, scale: 1.5 },
    { kind: "cta", text: "Please put one bin in every room this term.", scale: 0.95 },
    { kind: "caption", text: "Yours sincerely, and your name", scale: 0.75 },
  ],
};

// ─── Teacher notes ──────────────────────────────────────────────────────────

const NOTES_TITLE =
  "Week 7, Session 2. One line: same belief as yesterday, now aimed at the one person who can actually say yes.";

const NOTES_RESOURCES =
  "Prep slide. Print the Letter Builder. Students need yesterday's Persuasive Plan, plus their Week 6 Word Bank.\n" +
  "Decide before the lesson who the realistic audience is for common topics: principal, business manager, school council, local council.\n" +
  "CATCH-UP: no plan from yesterday? Give a spare Persuasive Plan and let them fill the belief and one reason in the launch. One reason is enough to write today.";

const NOTES_OVERVIEW =
  "For you, not the class. Skip past it when projecting.\n" +
  "The letter is the full-structure piece: all three anchor bands appear. The poster and article later in the week are compressions of it, so this session carries the most modelling.\n" +
  "Decision points: the audience hinge, the boards check in the We Do, and the exit ticket. Between them keep the pace brisk.";

const NOTES_LAUNCH = composeGlanceNotes({
  answer: "open - the person with the power to change it.",
  beats: [
    [
      "SAY: Read your belief off your plan, silently, once.",
      "SAY: Now the real question. Who could actually say yes to it?",
    ],
    [
      "ASK: Who is the one person who could make your change happen?",
      "40 sec. Turn and tell. Partner A first.",
      "EXPECT: a named role, not everyone.",
    ],
    [
      "SCAN the room as pairs talk.",
      "80%+ naming a role -> next slide.",
      "Less -> name three real roles yourself, re-ask.",
    ],
  ],
  trap: [
    "naming everyone, or naming a friend with no power.",
    "Fix: ask who signs it off, student names the role.",
  ],
  prep: "Low-coupling launch: a student without yesterday's plan can pick any belief here and still join in.",
  tag: "[Launch | Retention and recall | HITS 6]",
});

const NOTES_LI = composeGlanceNotes({
  beats: [
    "POINT to the intention. SAY: Today the letter. One reader, all three bands.",
    "SAY: Read the I can statements with me. Everyone, together, on three.",
    "SAY: Everyone leaves with a bold beginning written to a real person.",
  ],
  prep: "SC1 reachable by all; SC2 is the core target the exit ticket assesses; SC3 stretches into matching tone to audience.",
  tag: "[LI/SC | Planning made visible | HITS 1]",
});

const NOTES_MODEL = composeGlanceNotes({
  answer: "greeting, bold beginning, two mighty middles, then the ask.",
  beats: [
    [
      "POINT to the shape before the words.",
      "SAY: Greeting, my position, my reasons, then what I want.",
    ],
    "SAY: The bold beginning is the first line after Dear. Topic and side only.",
    [
      "ASK: Which part tells the reader what to DO?",
      "10 sec. Fingers at your chest. Point to the screen.",
      "EXPECT: the line near the bottom.",
    ],
  ],
  trap: [
    "burying the position in the middle.",
    "Fix: cover the first line, ask if you know the side.",
  ],
  stretch: "Say what changes if you write to the school council.",
  help: "Give the Builder, point only at the greeting.",
  prep: "Same three anchor bands as the Structure Card, laid out as a letter. Do not introduce new structure language.",
  tag: "[I Do | Explicit teaching | SC1 | HITS 3, 4]",
});

const NOTES_TONE = composeGlanceNotes({
  answer: "the right hand column - polite, certain, and specific.",
  beats: [
    [
      "SAY: Same belief, two tones. One gets read. One gets binned.",
      "READ the left column aloud in a whining voice. Then the right.",
    ],
    [
      "ASK: Which one would you act on if you were the principal?",
      "15 sec. Fingers at your chest. One or two.",
      "EXPECT: two.",
    ],
    [
      "SCAN the fingers.",
      "80%+ on two -> cold call one: what made it polite AND certain?",
      "Less -> read the left column again, ask how it feels, re-ask.",
    ],
  ],
  trap: [
    "thinking polite means unsure, so the modality drops.",
    "Fix: point at must in the right column, student reads it aloud.",
  ],
  prep: "Politeness and high modality are not opposites. That is the whole point of this slide.",
  tag: "[I Do | Explicit teaching | SC2 | HITS 4]",
});

const NOTES_CFU = composeGlanceNotes({
  answer: "B. It is polite, certain and says exactly what it wants.",
  beats: [
    [
      "SAY: Three opening lines to a principal. Only one works.",
      "Do not call out. Boards for this one.",
    ],
    [
      "ASK: Which opening would a principal act on, A, B or C?",
      "30 sec. Write it... chin it... show me.",
      "EXPECT: B.",
    ],
    [
      "SCAN boards, back row first.",
      "80%+ -> cold call one B: which word made it certain? Then reveal.",
      "Less -> read A aloud, ask what it actually wants, re-ask.",
    ],
    "REVEAL after boards are scanned. SAY: Polite and certain. Both, every time.",
  ],
  trap: [
    "picking C because it is the most polite.",
    "Fix: ask what C is asking for, student finds nothing, re-tests.",
  ],
  prep: "The hinge. A is rude and certain, C is polite and vague. Both wrong answers are diagnostic for tone.",
  tag: "[CFU hinge | Supported application | SC2 | HITS 7, 8]",
});

const NOTES_WEDO = composeGlanceNotes({
  answer: "open - a greeting, then a line naming the topic and the side.",
  beats: [
    "SAY: Dear Mrs Okafor. Then my topic and my side.",
    [
      "ASK: Write the greeting and bold beginning for YOUR letter.",
      "90 sec. Write it... chin it... show me.",
      "EXPECT: topic and side, with a should or must.",
    ],
    [
      "SCAN boards for a named reader and a clear side.",
      "80%+ -> cold call one strong, one shaky: who is it to?",
      "Less -> rebuild one board on the frame, re-ask.",
    ],
  ],
  trap: [
    "writing to nobody, so the tone drifts.",
    "Fix: student says the reader's name aloud, then rereads.",
  ],
  stretch: "Add a rhetorical question first.",
  help: "Give the greeting already written.",
  prep: "The greeting matters here. Writing to a named person is what fixes the tone for the rest of the draft.",
  tag: "[We Do | Supported application | SC2 | HITS 4, 5]",
});

const NOTES_YOUDO = composeGlanceNotes({
  answer: "open - a draft letter with a beginning, two reasons and an ask.",
  beats: [
    [
      "SAY: Your letter now. Your reader, your belief, your reasons.",
      "SAY: Two mighty middles is enough. Quality beats quantity.",
    ],
    [
      "POINT to the Builder. SAY: Beginning, two reasons, then what you want.",
      "TIME: twenty minutes drafting.",
    ],
    [
      "CIRCULATE. Check the reader is named before you read the argument.",
      "COLLECT nothing. Drafts stay for Session 5 publishing.",
    ],
  ],
  trap: [
    "writing all three reasons thinly instead of two properly.",
    "Fix: ring the weakest reason, student cuts it and grows the others.",
  ],
  stretch: "Add a paragraph answering what someone who disagrees would say.",
  help: "Give the Builder with the greeting and first line already written.",
  prep: "Protect the drafting time. Two well-proved reasons beat three thin ones, and the poster needs them sharp.",
  tag: "[You Do | Mastery and application | SC2 | HITS 10]",
});

const NOTES_EXIT = composeGlanceNotes({
  answer: "a named reader, and one sentence saying exactly what you want.",
  beats: [
    [
      "SAY: Two things at the bottom of your Builder.",
      "SAY: Who it is to, and the one thing you are asking for.",
    ],
    [
      "TIME: four minutes. Builders stay with you.",
      "WALK the rows and read over shoulders.",
    ],
    [
      "SORT in your head: clear ask, or vague ask.",
      "Vague ones -> two minutes each at the start of Session 3.",
    ],
  ],
  prep: "Assesses the core target. Do not collect the Builders; students draft on them again in Session 5.",
  tag: "[Exit Ticket | Mastery and application | SC2 | HITS 8]",
});

const NOTES_CLOSING = composeGlanceNotes({
  answer: "open - most students should show two.",
  beats: [
    "COLD CALL three students: who is your letter to, in one word?",
    [
      "ASK: Rate yourself on the three statements.",
      "15 sec. Fingers at your chest. One, two or three. Show me.",
      "EXPECT: most on two.",
    ],
    "SAY: Tomorrow the same belief, but the reader only gives you three seconds.",
  ],
  prep: "The three second line sets up the poster. Say it exactly as written so tomorrow's launch lands.",
  tag: "[Closing | Retention and recall | HITS 1, 9]",
});

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  titleSlide(pres, "The Letter",
    "Piece one: written to someone who can say yes",
    "Week 7 Session 2 | Year 5/6 Literacy", NOTES_TITLE);

  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  (() => {
    const s = P.customSlide(pres, T, "For the teacher", C.MUTED, "Session 2 at a glance", { badgeW: 2.4 });
    const cardY = CONTENT_TOP;
    const cardH = SAFE_BOTTOM - CONTENT_TOP;

    addCard(s, 0.5, cardY, 4.4, cardH, { strip: C.PRIMARY });
    s.addText("Where this sits", {
      x: 0.7, y: cardY + 0.12, w: 4.0, h: 0.28,
      fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText([
      { text: "Piece one of three, from one plan.", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: P.ANCHOR_PHRASE_W7, options: { color: C.ASSESS, bold: true, breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "The letter is the FULL structure piece:", options: { bold: true, breakLine: true } },
      { text: "all three anchor bands appear in it.", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "The poster and article later this week are compressions of this letter, so this session carries the most modelling.", options: { italic: true, color: C.MUTED } },
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
      { text: "Shape: example-first. The model letter is analysed by shape before wording.", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "Before the lesson:", options: { bold: true, color: C.ACCENT, breakLine: true } },
      { text: "decide the realistic audience for the common topics. Principal, business manager, school council, local council.", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "The idea students miss:", options: { bold: true, color: C.ACCENT, breakLine: true } },
      { text: "polite and certain are not opposites. A letter can be both, and it has to be.", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "Flex: two reasons is enough. Do not push for three.", options: { italic: true, color: C.MUTED } },
    ], {
      x: 5.3, y: cardY + 0.46, w: 4.0, h: cardH - 0.60,
      fontSize: 12.5, fontFace: FONT_B, color: C.CHARCOAL,
      valign: "top", margin: 0, paraSpaceAfter: 3,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_OVERVIEW);
    runSlideDiagnostics(s, pres);
  })();

  // Launch - who can say yes?
  (() => {
    const s = P.customSlide(pres, T, "Launch", C.SECONDARY, "Who can actually say yes?");
    const y0 = CONTENT_TOP;

    addTextOnShape(s, "Read your belief off your plan. Silently. Once.", {
      x: 0.5, y: y0, w: 9, h: 0.72, rectRadius: 0.08,
      fill: { color: C.SECONDARY },
    }, {
      fontSize: 21, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    s.addText("Now: who is the ONE person who could make it happen?", {
      x: 0.5, y: y0 + 0.94, w: 9, h: 1.25,
      fontSize: 36, fontFace: FONT_H, color: C.PRIMARY, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    const roles = ["The principal", "The business manager", "School council", "The local council"];
    P.drawWordChips(s, T, {
      x: 0.5, y: y0 + 2.32, w: 9, words: roles, perRow: 4,
      chipH: 0.62, gap: 0.14, fill: C.WHITE, fontSize: 15,
    });

    addTextOnShape(s, "Turn and tell. Partner A first.", {
      x: 2.95, y: y0 + 3.10, w: 4.1, h: 0.62, rectRadius: 0.08,
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
    "We are learning to write a persuasive letter to a reader who can act on it.",
    [
      "I can name the person my letter is written to.",
      "I can write a letter with a bold beginning, reasons and a clear ask.",
      "I can explain how my tone changes for that reader.",
    ],
    NOTES_LI, FOOTER);

  // I Do - the annotated model letter
  annotatedModelSlide(pres, "I Do", "The shape of a persuasive letter",
    [
      { text: "Look at the SHAPE first", role: "header" },
      "Where does the position sit?",
      "How many reason blocks?",
      "What is the last thing before the sign off?",
      { text: "The three bands are all here.", role: "body", bold: true },
    ],
    "Letter to Mrs Okafor",
    [
      { label: "Bold Beginning", detail: "One line: the topic and my side.", color: C.PRIMARY },
      { label: "Mighty Middle", detail: "One block per reason, with proof.", color: C.SECONDARY },
      { label: "Excellent Ending", detail: "Exactly what I want, in one line.", color: C.ASSESS },
      { label: "Voice", detail: "Polite AND certain. Both, always.", color: C.ACCENT },
    ],
    NOTES_MODEL, FOOTER,
    { previewSpec: LETTER_SPEC, previewAccent: C.PRIMARY, sourceType: "Model letter" });

  // I Do - tone comparison
  (() => {
    const s = P.customSlide(pres, T, "I Do", C.PRIMARY, "Same belief. One gets binned.");
    const y0 = CONTENT_TOP;
    const colW = 4.4;

    const cols = [
      {
        label: "Gets binned",
        color: C.ALERT,
        lines: [
          "You never listen to us.",
          "It is so unfair that we do not have a paper bin.",
          "Someone should probably do something about it.",
        ],
      },
      {
        label: "Gets read",
        color: C.SUCCESS,
        lines: [
          "Thank you for the time you give our ideas.",
          "Every classroom must have a scrap paper bin.",
          "Please add one bin per room this term.",
        ],
      },
    ];

    cols.forEach((col, i) => {
      const cx = 0.5 + i * (colW + 0.2);
      addTextOnShape(s, col.label, {
        x: cx, y: y0, w: colW, h: 0.60, rectRadius: 0.08,
        fill: { color: col.color },
      }, {
        fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      col.lines.forEach((line, j) => {
        const ly = y0 + 0.74 + j * 0.80;
        s.addShape("roundRect", {
          x: cx, y: ly, w: colW, h: 0.70, rectRadius: 0.07,
          fill: { color: C.WHITE }, line: { color: col.color, width: 1.2 },
        });
        s.addText(line, {
          x: cx + 0.16, y: ly, w: colW - 0.32, h: 0.70,
          fontSize: 15, fontFace: FONT_B, color: C.CHARCOAL,
          valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
        });
      });
    });

    addTextOnShape(s, "Polite and certain are not opposites.", {
      x: 0.5, y: y0 + 3.16, w: 9, h: 0.60, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, {
      fontSize: 19, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_TONE);
    runSlideDiagnostics(s, pres);
  })();

  // CFU hinge
  (() => {
    const s = P.customSlide(pres, T, "CFU", C.ALERT, "Which opening gets acted on?");
    const y0 = CONTENT_TOP;

    addTextOnShape(s, "A letter to the principal about lunchtime clubs", {
      x: 0.5, y: y0, w: 9, h: 0.48, rectRadius: 0.08,
      fill: { color: C.ALERT },
    }, {
      fontSize: 16, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    const options = [
      { key: "A", text: "Our school never lets us do anything fun and it has to stop." },
      { key: "B", text: "Thank you for reading. Our school must run lunchtime clubs this term." },
      { key: "C", text: "I was wondering if maybe some sort of club might possibly be nice." },
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
        addRevealAnswerBar(s, ["B - polite, and must"], {
          y: P.REVEAL_Y, h: P.REVEAL_H, fontSize: 24, color: C.SUCCESS,
        });
      },
    ]);

    addFooter(s, FOOTER);
    s.addNotes(NOTES_CFU);
    runSlideDiagnostics(s, pres);
  })();

  // We Do
  (() => {
    const s = P.customSlide(pres, T, "We Do", C.SUCCESS, "Open the letter together");
    const y0 = CONTENT_TOP;

    addCard(s, 0.5, y0, 9, 1.55, { strip: C.PRIMARY });
    s.addText("Mine", {
      x: 0.75, y: y0 + 0.09, w: 3, h: 0.28,
      fontSize: 14, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText([
      { text: "Dear Mrs Okafor,", options: { color: C.SECONDARY, bold: true, breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "I am writing because every classroom ", options: { color: C.CHARCOAL } },
      { text: "must", options: { color: C.ALERT, bold: true } },
      { text: " have a scrap paper bin.", options: { color: C.CHARCOAL } },
    ], {
      x: 0.75, y: y0 + 0.40, w: 8.5, h: 1.05,
      fontSize: 20, fontFace: FONT_B, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true, paraSpaceAfter: 2,
    });

    const yy = y0 + 1.72;
    addCard(s, 0.5, yy, 9, 1.10, { strip: C.SUCCESS });
    s.addText("Yours", {
      x: 0.75, y: yy + 0.09, w: 3, h: 0.28,
      fontSize: 14, fontFace: FONT_B, color: C.SUCCESS, bold: true, margin: 0,
    });
    s.addText("Your greeting, then ONE line with your topic and your side.", {
      x: 0.75, y: yy + 0.40, w: 8.5, h: 0.60,
      fontSize: 20, fontFace: FONT_B, color: C.CHARCOAL,
      valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
    });

    addTextOnShape(s, "Write it... chin it... show me.", {
      x: 0.5, y: yy + 1.26, w: 9, h: 0.72, rectRadius: 0.08,
      fill: { color: C.WHITE }, line: { color: C.SUCCESS, width: 1.5 },
    }, {
      fontSize: 19, fontFace: FONT_B, color: C.SUCCESS, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_WEDO);
    runSlideDiagnostics(s, pres);
  })();

  // You Do
  (() => {
    const s = P.customSlide(pres, T, "You Do", C.ASSESS, "Draft your letter");
    const y0 = CONTENT_TOP;

    const bands = [
      { label: "Bold Beginning", detail: "Greeting, then your topic and side.", color: C.PRIMARY },
      { label: "Mighty Middle", detail: "TWO reasons. One block each, with proof.", color: C.SECONDARY },
      { label: "Excellent Ending", detail: "Exactly what you want, in one line.", color: C.ASSESS },
    ];
    const rh = 0.86;
    const gap = 0.12;
    bands.forEach((band, i) => {
      const ry = y0 + i * (rh + gap);
      s.addShape("roundRect", {
        x: 0.5, y: ry, w: 3.1, h: rh, rectRadius: 0.07,
        fill: { color: band.color },
      });
      s.addText(band.label, {
        x: 0.5, y: ry, w: 3.1, h: rh,
        fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      s.addShape("roundRect", {
        x: 3.75, y: ry, w: 5.75, h: rh, rectRadius: 0.07,
        fill: { color: C.WHITE }, line: { color: band.color, width: 1.3 },
      });
      s.addText(band.detail, {
        x: 3.95, y: ry, w: 5.35, h: rh,
        fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL,
        valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
    });

    addTextOnShape(s, "Letter Builder. Twenty minutes. Two reasons, done properly.", {
      x: 0.5, y: y0 + 3 * (rh + gap) + 0.06, w: 9, h: 0.76, rectRadius: 0.08,
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
    topic: "At the bottom of your Letter Builder",
    task: "Who is your letter to? What exactly are you asking for?",
    cue: "One name. One sentence.",
    taskSize: 27,
    footer: FOOTER,
    notes: NOTES_EXIT,
  });

  closingSlide(pres, {
    reflectionPrompt: "What would you change about your tone if you wrote to a friend instead?",
    scItems: [
      "I can name the person my letter is written to.",
      "I can write a letter with a bold beginning, reasons and a clear ask.",
      "I can explain how my tone changes for that reader.",
    ],
    selfAssessment: "Fingers at your chest. One, two or three. Show me.",
    takeaways: ["Polite and certain. Both, every time."],
  }, NOTES_CLOSING);

  const pptxPath = path.join(OUT_DIR, "Persuasive Writing W7 S2.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);
}

// ─── Resource: Letter Builder ───────────────────────────────────────────────

async function buildLetter() {
  const doc = createPdf({ title: LETTER_RES.name });
  let y = addPdfHeader(doc, "Letter Builder", {
    subtitle: "One reader. Two reasons. One clear ask.",
    color: C.PRIMARY,
    lessonInfo: "Week 7 Session 2 | Year 5/6 Literacy",
  });

  y = addSectionHeading(doc, "The model, labelled by band", y, { color: C.PRIMARY });
  y += 4;

  // A schematic of the model letter, drawn on the page, so the printed
  // scaffold shows the same shape students analysed on screen.
  const modelBands = [
    { band: "Greeting", text: "Dear Mrs Okafor,", color: C.MUTED, lines: 1 },
    { band: "Bold Beginning", text: "I am writing because every classroom must have a scrap paper bin.", color: C.PRIMARY, lines: 1 },
    { band: "Mighty Middle 1", text: "A bin saves money because paper costs the school real money. For example, the office could tell us the termly bill. This means every unused sheet is money gone. Therefore, a bin pays for itself.", color: C.SECONDARY, lines: 3 },
    { band: "Mighty Middle 2", text: "A bin cuts what we send to landfill because most of our paper is only used on one side. For example, one week of our own bin would show it. Therefore, we should reuse before we recycle.", color: C.SECONDARY, lines: 3 },
    { band: "Excellent Ending", text: "Please put one scrap paper bin in every room this term. Our class would set them up.", color: C.ASSESS, lines: 2 },
    { band: "Sign off", text: "Yours sincerely, Room 12", color: C.MUTED, lines: 1 },
  ];

  modelBands.forEach((band) => {
    const boxH = 16 + band.lines * 12;
    doc.save();
    doc.roundedRect(PAGE.MARGIN, y, 96, boxH, 3).fill(hex(band.color));
    doc.fontSize(8.5).font("Sans-Bold").fillColor("#FFFFFF");
    doc.text(band.band, PAGE.MARGIN + 4, y + boxH / 2 - 5, { width: 88, align: "center" });
    doc.roundedRect(PAGE.MARGIN + 102, y, PAGE.CONTENT_W - 102, boxH, 3)
      .lineWidth(0.7).strokeColor("#9CA3AF").stroke();
    doc.fontSize(9).font("Sans").fillColor("#000000");
    doc.text(band.text, PAGE.MARGIN + 108, y + 6, { width: PAGE.CONTENT_W - 116 });
    doc.restore();
    y += boxH + 5;
  });

  y += 6;
  y = addTipBox(doc,
    "Polite and certain are not opposites. Thank the reader, then use must, should or never.",
    y, { color: C.ACCENT });

  doc.addPage();
  y = PAGE.MARGIN;
  y = addSectionHeading(doc, "Your letter. Draft it here.", y, { color: C.PRIMARY });
  y += 6;

  doc.fontSize(11).font("Sans-Bold").fillColor("#000000");
  doc.text("My letter is to:", PAGE.MARGIN, y);
  doc.moveTo(PAGE.MARGIN + 92, y + 14).lineTo(PAGE.MARGIN + PAGE.CONTENT_W, y + 14)
    .lineWidth(0.9).strokeColor("#000000").stroke();
  y += 30;

  const frames = [
    { label: "Greeting and Bold Beginning", starter: "Dear ..., I am writing because ...", lines: 2, color: C.PRIMARY },
    { label: "Mighty Middle 1", starter: "Reason. Evidence. Explain. Link.", lines: 4, color: C.SECONDARY },
    { label: "Mighty Middle 2", starter: "Reason. Evidence. Explain. Link.", lines: 4, color: C.SECONDARY },
    { label: "Excellent Ending and sign off", starter: "Please ... Yours sincerely, ...", lines: 3, color: C.ASSESS },
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
    "Challenge: add a paragraph that answers what someone who disagrees would say, "
    + "then turn it around with However.",
    y, { color: C.ACCENT });

  addPdfFooter(doc, "Persuasive Writing | Week 7 Session 2 | Year 5/6 Literacy");
  await writePdf(doc, path.join(OUT_DIR, LETTER_RES.fileName));
  console.log("Wrote " + LETTER_RES.name);
}

(async () => {
  await build();
  await buildLetter();
})().catch((err) => {
  console.error(err);
  process.exit(1);
});
