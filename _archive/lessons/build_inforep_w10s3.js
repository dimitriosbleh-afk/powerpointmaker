"use strict";

// Information Reports - Term 3 Week 10, Session 3 (Year 5/6 Literacy)
// "A picture that teaches, then a careful edit"
//
// Victorian Curriculum 2.0: English, Levels 5-6 - VC2E6LA07, the way still
// images (figures, diagrams, maps, tables) are used in texts, and VC2E6LY10,
// re-reading and editing your own texts and the texts of others using agreed
// criteria. The Editor's Challenge on the sheet reaches into VC2E7LY09
// (removing repetition, reordering, substituting words for impact).
//
// UNIT ANCHOR (locked): "Classify it. Describe it, one aspect at a time.
// Wrap it up. Facts all the way through."
//
// Lesson shape: compare-two-models for the launch, then error-analysis-led for
// the editing half, because students edit better when they have just fixed
// someone else's paragraph.
//
// Sources: teacher's unit plan - the editing checklist and the Two Stars and a
// Wish routine are quoted verbatim.

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

const P = require("./inforep_lib");

const UNIT_VARIANT = 2;
const T = createTheme("literacy", "grade56", UNIT_VARIANT);
const {
  C, FONT_H, FONT_B, SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, keyWordSlide, closingSlide,
  annotatedModelSlide, compareVisualSlide,
  addFooter, addTextOnShape, addRevealAnswerBar,
  clickBuild, runSlideDiagnostics,
} = T;

const SESSION = 3;
const FOOTER = "Information Reports | Week 10 Session 3 | Year 5/6 Literacy";
const OUT_DIR = path.join(__dirname, "..", "output", "InfoReport_W10_S3");
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const SHEET_RES = makeSessionResource(SESSION,
  "Visuals and Editing Sheet",
  "Plan a visual and its caption, then self-edit and peer-edit with the checklist.");
const RESOURCE_ITEMS = [SHEET_RES];
fs.mkdirSync(RES_DIR, { recursive: true });

/* --- Teacher notes (Glance Format, megaprompt sections 45-47) ------------- */

const NOTES_TITLE =
  "Session 3. Drafts are written. Today we add a picture that teaches, then we edit properly.";

const NOTES_RESOURCES =
  "Prep slide. Print the Visuals and Editing Sheet, one per student, and hand back their drafts.\n" +
  "Have pencils for sketching. Devices only if students are searching for copyright-free images.\n" +
  "CATCH-UP: draft not finished? Give five minutes now while others start their visual, then join at the caption model. The editing half needs only one paragraph.";

const NOTES_LAUNCH = composeGlanceNotes({
  answer: "the right one. The diagram and caption tell you something extra.",
  beats: [
    [
      "SAY: Same report, two pages. Look at both.",
      "One has a picture with words underneath.",
    ],
    [
      "ASK: Which page helps a reader more?",
      "20 sec. Fingers at your chest: left one, right two.",
      "EXPECT: two.",
    ],
    [
      "SCAN the fingers.",
      "80%+ -> cold call one: what does the picture add?",
      "Less -> cover the right page, ask what is missing, re-ask.",
    ],
  ],
  trap: [
    "saying the picture just looks nicer.",
    "Fix: ask what the reader now knows, student names a fact.",
  ],
  prep: "Low-coupling launch: both pages are on screen, so nothing earlier is assumed. Whole block under 5 minutes.",
  sources: P.SOURCE_LINE_MODEL,
  tag: "[Launch | Attention, focus and regulation | HITS 6]",
});

const NOTES_LI = composeGlanceNotes({
  beats: [
    [
      "POINT to the learning intention.",
      "SAY: Today your report gets a picture that teaches, then a real edit.",
    ],
    "SAY: Read the I can statements with me. Everyone, together, on three.",
    "SAY: The first one everyone leaves with. Pick a visual that helps.",
  ],
  prep: "SC1 is reachable by every student; SC2 is the core target the exit ticket assesses; SC3 stretches into peer editing. Tier labels stay off the slide.",
  tag: "[LI/SC | Planning made visible | HITS 1]",
});

const NOTES_VOCAB = composeGlanceNotes({
  answer: "a caption is the line under a picture that explains it.",
  beats: [
    [
      "SAY: A caption is the little line under a picture.",
      "It is not decoration. It teaches.",
    ],
    [
      "ASK: What should a caption tell you?",
      "20 sec. Turn and tell, partner A first.",
      "EXPECT: what it is, and something extra.",
    ],
    "SAY: If the caption only says what you can see, it has wasted its space.",
  ],
  trap: [
    "writing a title instead of a caption.",
    "Fix: ask what it adds, student adds a fact.",
  ],
  prep: "The one word for today. Captions are where most students give up, so name and model it explicitly.",
  tag: "[Key Vocabulary | Knowledge and memory | HITS 6]",
});

const NOTES_VISUALS = composeGlanceNotes({
  answer: "match the visual to the aspect it explains.",
  beats: [
    [
      "SAY: Four kinds of visual, each good at one job.",
      "Read them with me.",
    ],
    [
      "POINT to each row. A map for habitat, a diagram for parts.",
      "A table for comparing numbers.",
    ],
    [
      "ASK: Which visual suits your habitat paragraph?",
      "20 sec. Write it... chin it... show me.",
      "EXPECT: a map.",
    ],
    [
      "SCAN boards.",
      "80%+ -> cold call: why not a photograph there?",
      "Less -> read the habitat paragraph, re-ask.",
    ],
  ],
  trap: [
    "choosing a photograph for everything.",
    "Fix: ask what it explains, student re-matches.",
  ],
  stretch: "Name an aspect a table suits better than a photo.",
  help: "Give the aspect. Student picks from two.",
  prep: "Say the anchor once here. Choosing the visual is a thinking task, not a decorating task.",
  tag: "[I Do | Explicit teaching | SC1 | HITS 3, 4]",
});

const NOTES_CAPTION = composeGlanceNotes({
  answer: "number it, name it, add a new fact.",
  beats: [
    [
      "SAY: Read my caption, then the four rules beside it.",
      "Figure 1. The toe pads. Then the extra fact.",
    ],
    [
      "POINT to each rule card as you name it.",
      "SAY: The last rule is the forgotten one. Add something.",
    ],
    [
      "ASK: What does my caption add that the report does not?",
      "20 sec. Turn and tell, partner B first.",
      "EXPECT: the pads work on smooth surfaces.",
    ],
  ],
  trap: [
    "repeating the paragraph inside the caption.",
    "Fix: cover the caption, ask what is new, student rewrites.",
  ],
  stretch: "Write a caption for a map of your topic.",
  help: "Give the naming half. Student adds the fact.",
  prep: "The caption is where VC2E6LA07 lives: the image and its caption carry information the prose does not.",
  sources: P.SOURCE_LINE_MODEL,
  tag: "[I Do | Explicit teaching | SC2 | HITS 4]",
});

const NOTES_CFU = composeGlanceNotes({
  answer: "B. It numbers it, names it, and adds the toe pad fact.",
  beats: [
    [
      "SAY: Three captions for the same picture. Only one works.",
      "Do not call out. This one goes on boards.",
    ],
    [
      "ASK: Which caption does its job?",
      "30 sec. Boards up on cue.",
      "EXPECT: B.",
    ],
    [
      "SCAN boards, back row first.",
      "80%+ -> cold call one B: what does A leave out?",
      "Less -> read A, ask what it adds, re-ask.",
    ],
    [
      "REVEAL after boards are scanned.",
      "SAY: A only names it. C is about you, not the picture.",
    ],
  ],
  trap: [
    "picking C because it sounds honest.",
    "Fix: ask what the reader learns from C, student re-tests.",
  ],
  prep: "The hinge of the lesson. A is the naming-only caption and C is the where-I-got-it caption, which are the two failures in student work.",
  tag: "[CFU hinge | Supported application | SC2 | HITS 7, 8]",
});

const NOTES_WEDO = composeGlanceNotes({
  answer: "three fixes: the best, lived, and the small letter t.",
  beats: [
    [
      "SAY: Someone else's paragraph. Three things break our checklist.",
      "Read it aloud with me.",
    ],
    [
      "ASK: Find all three. Write the words that have to change.",
      "90 sec. Write it... chin it... show me.",
      "EXPECT: best, lived, they.",
    ],
    [
      "SCAN boards for all three.",
      "80%+ -> cold call one: which checklist line catches each?",
      "Less -> read one sentence at a time, check it, re-ask.",
    ],
  ],
  trap: [
    "finding the opinion and stopping there.",
    "Fix: ask how many are left, student keeps hunting.",
  ],
  stretch: "Rewrite the paragraph with all three fixed.",
  help: "Say how many errors are in each sentence.",
  prep: "Editing someone else's work first is what makes self-editing possible. Different paragraph from the sheet.",
  tag: "[We Do | Supported application | SC3 | HITS 4, 5]",
});

const NOTES_YOUDO = composeGlanceNotes({
  answer: "a visual with a caption, a self-edit, then a partner's feedback.",
  beats: [
    [
      "SAY: Three jobs. Visual and caption, self-edit, then swap.",
      "Two stars and a wish, and both stars must be specific.",
    ],
    [
      "POINT to the steps. TIME: twenty-five minutes.",
      "Sketch first. Do not go looking for a picture yet.",
    ],
    [
      "CIRCULATE. Read the captions, not the sketches.",
      "COLLECT two captions to read at the closing.",
    ],
  ],
  trap: [
    "giving a wish that is really an insult, or a star that says good.",
    "Fix: ask which checklist line it points at, student rewrites it.",
  ],
  stretch: "The Editor's Challenge box on the sheet.",
  help: "Edit for two checklist lines only, not all seven.",
  prep: "Peer editing uses the teacher's own Two Stars and a Wish wording. Model one star aloud before they swap.",
  sources: P.SOURCE_LINE_UNIT,
  tag: "[You Do | Mastery and application | SC3 | HITS 5, 8]",
});

const NOTES_EXIT = composeGlanceNotes({
  answer: "a caption that numbers it, names it, and adds a fact.",
  beats: [
    [
      "SAY: One caption. The one for your own visual.",
      "Number it, name it, add something new.",
    ],
    [
      "TIME: three minutes. Sheets open on the desk when you finish.",
      "COLLECT by walking the rows, not by calling names.",
    ],
    [
      "SORT into two piles as you collect: adds a fact, names only.",
      "A thick second pile -> re-model one caption before publishing.",
    ],
  ],
  prep: "Assesses the core target. The pile sort tells you whether to re-model captions at the start of Session 4.",
  tag: "[Exit Ticket | Mastery and application | SC2 | HITS 8]",
});

const NOTES_CLOSING = composeGlanceNotes({
  answer: "open - most students should show two.",
  beats: [
    "READ the two collected captions aloud. SAY: Listen for the extra fact.",
    [
      "ASK: Rate yourself on the three statements.",
      "15 sec. Fingers at your chest, one, two or three. Show me.",
      "EXPECT: most on two.",
    ],
    "SAY: Next session we publish, and we read each other's work.",
  ],
  prep: "Anyone showing one gets their caption re-modelled with you at the start of Session 4.",
  tag: "[Closing | Retention and recall | HITS 1, 9]",
});

/* --- Slide content -------------------------------------------------------- */

const CAPTION_RULES = [
  { label: "Number it", detail: "Figure 1, Figure 2, in order." },
  { label: "Name it", detail: "Say what the reader is looking at." },
  { label: "Add something", detail: "A fact the paragraph does not have." },
  { label: "Keep it short", detail: "One or two sentences, no more." },
];

const VISUAL_MATCH = [
  { label: "Labelled diagram", use: "Shows the parts of something", aspect: "Appearance" },
  { label: "Map", use: "Shows where something is", aspect: "Habitat" },
  { label: "Photograph", use: "Shows what it really looks like", aspect: "Appearance" },
  { label: "Table", use: "Compares numbers or groups", aspect: "Types or size" },
];

// A paragraph with three planted errors, one per checklist line.
const EDIT_PARAGRAPH = [
  "Green tree frogs are the best climbers in Australia.",
  "They lived in trees and reeds near still water.",
  "they also hide in downpipes and water tanks.",
];

/* --- Build ---------------------------------------------------------------- */

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // 1. Title
  titleSlide(pres, "Visuals, Captions and Editing",
    "A picture that teaches. Then a careful edit.",
    "Week 10 Session 3 | Year 5/6 Literacy", NOTES_TITLE);

  // 2. Teacher Resources
  addResourceSlide(pres, {
    resources: RESOURCE_ITEMS,
    studentTools: [
      "Their draft, handed back at the door",
      "The Visuals and Editing Sheet, one each",
      "Pencils for sketching",
      "Mini-whiteboards for the We Do",
    ],
    boardSetup: [
      "Devices only for students searching copyright-free images",
      "Sketch first is the rule; say it before they start",
    ],
  }, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // 3. Launch - same report, two pages
  compareVisualSlide(pres, "Launch", "Which page helps a reader more?",
    "Fingers at your chest: one for the left page, two for the right.",
    {
      panelTitle: "Page 1",
      title: "Words only",
      strip: C.SECONDARY,
      previewSpec: P.noVisualPageSpec(C, P.FROG),
      previewH: 2.10,
    },
    {
      panelTitle: "Page 2",
      title: "With a diagram and a caption",
      strip: C.PRIMARY,
      previewSpec: P.reportPageSpec(C, P.FROG),
      previewH: 2.10,
    },
    NOTES_LAUNCH, FOOTER,
    // The default 2.15" card leaves the bottom third of the slide empty and
    // shrinks both pages to thumbnails; students have to READ these.
    { badgeColor: C.SECONDARY, badgeFill: C.SECONDARY, badgeW: 1.8,
      promptFill: C.PRIMARY, cardH: 3.00 });

  // 4. Learning Intention and Success Criteria
  liSlide(pres,
    "We are learning to add a visual with a caption and to edit against a checklist.",
    [
      "I can choose a visual that helps the reader understand one part of my report.",
      "I can write a caption that adds information.",
      "I can edit my writing and a partner's using the checklist.",
    ],
    NOTES_LI, FOOTER);

  // 5. Key vocabulary
  keyWordSlide(pres, {
    word: "caption",
    meaning: "The line under a picture that explains it.",
    example: "Figure 1: The toe pads of a green tree frog. Wide pads let it climb glass.",
    routine: ["Say it", "Find one", "Write one"],
    color: C.ASSESS,
    title: "The word for today",
  }, NOTES_VOCAB, FOOTER);

  // 6. I Do A - match the visual to the aspect
  (() => {
    const s = P.customSlide(pres, T, "I Do", C.PRIMARY, "Which visual for which aspect?");
    const y0 = CONTENT_TOP;
    // Four rows plus a closing line must fit above 5.10:
    // 1.30 + 4 x 0.84 = 4.66, closing line to 5.06.
    const rowH = 0.72;
    const gap = 0.12;
    const colors = [C.PRIMARY, C.SECONDARY, C.ACCENT, C.SUCCESS];

    VISUAL_MATCH.forEach((item, i) => {
      const ry = y0 + i * (rowH + gap);
      const accent = colors[i];
      s.addShape("roundRect", {
        x: 0.5, y: ry, w: 9, h: rowH, rectRadius: 0.08,
        fill: { color: C.WHITE }, line: { color: accent, width: 1.4 },
      });
      addTextOnShape(s, item.label, {
        x: 0.5, y: ry, w: 2.5, h: rowH, rectRadius: 0.08,
        fill: { color: accent },
      }, {
        fontSize: 15, fontFace: FONT_B, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      s.addText(item.use, {
        x: 3.20, y: ry, w: 3.5, h: rowH,
        fontSize: 15.5, fontFace: FONT_B, color: C.CHARCOAL,
        valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
      addTextOnShape(s, item.aspect, {
        x: 6.95, y: ry + 0.16, w: 2.3, h: rowH - 0.32, rectRadius: 0.07,
        fill: { color: C.WHITE }, line: { color: accent, width: 1.4 },
      }, {
        fontSize: 13, fontFace: FONT_B, color: accent, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
    });

    const noteY = y0 + 4 * (rowH + gap);
    s.addText("Pick the visual that explains the aspect. Not the prettiest one.", {
      x: 0.5, y: noteY, w: 9, h: 0.40,
      fontSize: 17, fontFace: FONT_H, color: C.PRIMARY, bold: true, italic: true,
      align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_VISUALS);
    runSlideDiagnostics(s, pres);
  })();

  // 7. I Do B - what a caption has to do
  annotatedModelSlide(pres, "I Do", "A caption that does its job",
    // The caption itself sits in the prompt card, at a size students can read.
    // In the page mockup it renders at caption size, which shows WHERE it sits
    // but not what it says.
    [
      { text: "Read the caption", role: "header" },
      { text: "Figure 1: The toe pads of a green tree frog. Wide, sticky pads let it climb smooth surfaces.", role: "emphasis" },
      { text: "Number it. Name it." },
      { text: "Then a fact the paragraph does not have." },
      { text: "Two sentences at most." },
    ],
    "Where the caption sits",
    CAPTION_RULES,
    NOTES_CAPTION, FOOTER,
    {
      badgeColor: C.PRIMARY,
      previewSpec: P.reportPageSpec(C, P.FROG),
      sourceType: "Model report",
    });

  // 8. CFU hinge - which caption does its job?
  (() => {
    const s = P.customSlide(pres, T, "CFU", C.ALERT, "Which caption does its job?");
    const y0 = CONTENT_TOP;

    addTextOnShape(s, "The picture: a green tree frog gripping a window", {
      x: 0.5, y: y0, w: 9, h: 0.48, rectRadius: 0.08,
      fill: { color: C.ALERT },
    }, {
      fontSize: 16, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    P.drawOptionStack(s, T, P.CAPTION_OPTIONS, {
      y: y0 + 0.60, optionH: 0.66, gap: 0.12, color: C.PRIMARY, fontSize: 16,
    });

    clickBuild(s, [
      () => {
        addRevealAnswerBar(s, ["B - it names it AND adds the toe pad fact"], {
          y: P.REVEAL_Y, h: P.REVEAL_H, fontSize: 20, color: C.SUCCESS,
        });
      },
    ]);

    addFooter(s, FOOTER);
    s.addNotes(NOTES_CFU);
    runSlideDiagnostics(s, pres);
  })();

  // 9. We Do - edit someone else's paragraph
  (() => {
    const s = P.customSlide(pres, T, "We Do", C.SUCCESS, "Three things break our checklist");
    const y0 = CONTENT_TOP;

    P.drawReportExtract(s, T, {
      x: 0.5, y: y0, w: 9, h: 1.85,
      bodySize: 17,
      showLabels: false,
      sections: EDIT_PARAGRAPH.map((line) => ({ text: line })),
    });

    const chipY = y0 + 1.97;
    const chips = ["No opinions?", "Present tense?", "Capitals?"];
    const chipGap = 0.16;
    const chipW = (9 - chipGap * 2) / 3;
    chips.forEach((label, i) => {
      addTextOnShape(s, label, {
        x: 0.5 + i * (chipW + chipGap), y: chipY, w: chipW, h: 0.50, rectRadius: 0.08,
        fill: { color: C.SECONDARY },
      }, {
        fontSize: 15, fontFace: FONT_B, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
    });

    addTextOnShape(s, "Write the three words that have to change.", {
      x: 0.5, y: chipY + 0.62, w: 9, h: 0.54, rectRadius: 0.08,
      fill: { color: C.SUCCESS },
    }, {
      fontSize: 19, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    clickBuild(s, [
      () => {
        addRevealAnswerBar(s, ["best (opinion)", "lived becomes live", "they needs a capital"], {
          y: P.REVEAL_Y, h: P.REVEAL_H, fontSize: 16, color: C.SUCCESS,
        });
      },
    ]);

    addFooter(s, FOOTER);
    s.addNotes(NOTES_WEDO);
    runSlideDiagnostics(s, pres);
  })();

  // 10. You Do - visual, caption, self-edit, peer edit
  (() => {
    const s = P.customSlide(pres, T, "You Do", C.ASSESS, "Your visual, then your edit");
    const y0 = CONTENT_TOP;
    const h = SAFE_BOTTOM - y0;

    T.addInstructionCard(s, [
      { text: "Visuals and Editing Sheet", role: "header" },
      { text: "1. Sketch one visual. Write its caption." },
      { text: "2. Self-edit with the seven-line checklist." },
      { text: "3. Swap. Two stars and a wish." },
      { text: "Twenty-five minutes. Sketch before searching." },
    ], {
      x: 0.5, y: y0, w: 4.7, h,
      strip: C.ASSESS, fill: C.WHITE,
      headerColor: C.ASSESS, emphasisColor: C.ALERT,
    });

    const rx = 5.5;
    const rw = 4.0;
    const fb = [
      { label: "Star", detail: "One thing that works, and why.", color: C.SUCCESS },
      { label: "Star", detail: "A second thing, from the checklist.", color: C.SUCCESS },
      { label: "Wish", detail: "One change, pointing at one checklist line.", color: C.ACCENT },
    ];
    // Header 0.50 at 1.30, cards start 1.92, three at 0.96 with 0.12 gaps
    // finish at 5.04, clear of the 5.10 floor.
    const fbH = 0.96;
    addTextOnShape(s, "Two stars and a wish", {
      x: rx, y: y0, w: rw, h: 0.50, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, {
      fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
    fb.forEach((item, i) => {
      const fy = y0 + 0.62 + i * (fbH + 0.12);
      s.addShape("roundRect", {
        x: rx, y: fy, w: rw, h: fbH, rectRadius: 0.07,
        fill: { color: C.WHITE }, line: { color: item.color, width: 1.4 },
      });
      s.addText([
        { text: item.label, options: { fontSize: 15, bold: true, color: item.color, breakLine: true } },
        { text: item.detail, options: { fontSize: 13, color: C.CHARCOAL } },
      ], {
        x: rx + 0.20, y: fy, w: rw - 0.40, h: fbH,
        fontFace: FONT_B, valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
    runSlideDiagnostics(s, pres);
  })();

  // 11. Exit ticket
  P.exitTicketPanel(pres, T, {
    topic: "The visual you chose for your own report",
    task: "Write its caption.",
    cue: "Number it. Name it. Add something new.",
    taskSize: 34,
    footer: FOOTER,
    notes: NOTES_EXIT,
  });

  // 12. Closing
  closingSlide(pres, {
    reflectionPrompt: "What did your partner's wish make you want to change?",
    scItems: [
      "I can choose a visual that helps the reader understand one part of my report.",
      "I can write a caption that adds information.",
      "I can edit my writing and a partner's using the checklist.",
    ],
    selfAssessment: "Fingers at your chest. One, two or three. Show me.",
    takeaways: [P.ANCHOR_PHRASE],
  }, NOTES_CLOSING);

  const pptxPath = path.join(OUT_DIR, "Information Reports W10 S3.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);
}

/* --- Resource: Visuals and Editing Sheet ---------------------------------- */

async function buildSheet() {
  const doc = createPdf({ title: SHEET_RES.name });
  let y = addPdfHeader(doc, "Visuals and Editing Sheet", {
    subtitle: "Plan your visual and its caption. Then edit, yours and a partner's.",
    color: C.PRIMARY,
    lessonInfo: "Year 5/6 Literacy | Information Reports | Week 10",
  });

  y = addSectionHeading(doc, "Part 1: which visual, and where does it go?", y, { color: C.PRIMARY });

  // The four visual types, on paper, so the choice is made against the same
  // four options students saw on screen.
  P.VISUAL_TYPES.forEach((item, i) => {
    const accent = hex([C.PRIMARY, C.SECONDARY, C.ACCENT, C.SUCCESS][i]);
    doc.save();
    doc.roundedRect(PAGE.MARGIN, y, 130, 20, 3).fill(accent);
    doc.fontSize(9.5).font("Sans-Bold").fillColor("#FFFFFF");
    doc.text(item.label, PAGE.MARGIN, y + 6, { width: 130, align: "center", lineBreak: false });
    doc.restore();
    doc.fontSize(10.5).font("Sans").fillColor("#000000");
    doc.text(item.use, PAGE.MARGIN + 140, y + 4, { width: PAGE.CONTENT_W - 148, lineBreak: false });
    y += 26;
  });
  y += 6;

  y = P.addNoteBoxPdf(doc, y, "Which aspect will your visual belong to, and which kind will it be?",
    { color: hex(C.PRIMARY), lines: 2, spacing: 24, dots: false });

  y = addSectionHeading(doc, "Part 2: sketch it here", y, { color: C.SECONDARY });
  y = addBodyText(doc,
    "Sketch is enough. Label the parts you want the reader to notice.", y);
  doc.save();
  doc.roundedRect(PAGE.MARGIN, y, PAGE.CONTENT_W, 210, 4)
    .lineWidth(1.2).strokeColor(hex(C.SECONDARY)).stroke();
  doc.restore();
  y += 222;

  y = P.addNoteBoxPdf(doc, y, "Your caption: number it, name it, add something new",
    { color: hex(C.ASSESS), lines: 2, spacing: 26, dots: false });

  doc.addPage();
  y = PAGE.MARGIN;

  // The teacher's own checklist, wording locked (megaprompt section 5b).
  y = addSectionHeading(doc, "Part 3: self-edit. Tick each line only when it is true.", y,
    { color: C.PRIMARY });
  y = P.addChecklistPdf(doc, y, P.EDIT_CHECKLIST, { color: hex(C.PRIMARY) });
  y += 4;

  y = addSectionHeading(doc, "Part 4: peer feedback", y, { color: C.SUCCESS });
  y = addBodyText(doc,
    "Swap with a partner. " + P.STARS_AND_WISH + ". Point every one of them at a line from the checklist above.", y);

  y = P.addNoteBoxPdf(doc, y, "Star 1: what works, and why",
    { color: hex(C.SUCCESS), lines: 2, spacing: 24, dots: false });
  y = P.addNoteBoxPdf(doc, y, "Star 2: a second thing that works",
    { color: hex(C.SUCCESS), lines: 2, spacing: 24, dots: false });
  y = P.addNoteBoxPdf(doc, y, "Wish: one change, and which checklist line it comes from",
    { color: hex(C.ACCENT), lines: 2, spacing: 24, dots: false });

  doc.addPage();
  y = PAGE.MARGIN;

  // Extension. Same concept, deeper: this is the Year 7 editing standard the
  // teacher listed (VC2E7LY09), startable without help.
  y = addSectionHeading(doc, "Part 5: the Editor's Challenge", y, { color: C.ASSESS });
  y = addBodyText(doc,
    "For anyone whose checklist is fully ticked. Real editors do more than fix mistakes. Do all three, then answer the last question.", y);

  y = addTipBox(doc,
    "1. Find a word or idea you repeated. Cut one of them. " +
    "2. Find two sentences that would read better swapped. Swap them. " +
    "3. Find one flat word and replace it with a sharper one.", y, { color: C.ASSESS });

  y = P.addNoteBoxPdf(doc, y, "The repetition I cut", { color: hex(C.ASSESS), lines: 2, spacing: 25, dots: false });
  y = P.addNoteBoxPdf(doc, y, "The two sentences I reordered", { color: hex(C.ASSESS), lines: 2, spacing: 25, dots: false });
  y = P.addNoteBoxPdf(doc, y, "The word I swapped, and what I swapped it for", { color: hex(C.ASSESS), lines: 2, spacing: 25, dots: false });
  y = P.addNoteBoxPdf(doc, y, "Which change made the biggest difference to a reader, and why?",
    { color: hex(C.PRIMARY), lines: 4, spacing: 26, dots: false });

  addPdfFooter(doc, "Information Reports | Year 5/6 Literacy | Week 10");
  await writePdf(doc, path.join(OUT_DIR, SHEET_RES.fileName));
  console.log("Wrote " + SHEET_RES.name);
}

(async () => {
  await build();
  await buildSheet();
})().catch((err) => {
  console.error(err);
  process.exit(1);
});
