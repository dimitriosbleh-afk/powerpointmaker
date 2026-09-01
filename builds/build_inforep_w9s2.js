"use strict";

// Information Reports - Term 3 Week 9, Session 2 (Year 5/6 Literacy)
// "The parts, and the job each part does"
//
// Victorian Curriculum 2.0: English, Literacy, Levels 5-6 - text structure
// appropriate for topic, purpose and audience (VC2E5LY10, VC2E6LY09), and
// VC2E6LA07 for the way a diagram and its caption carry information.
//
// UNIT ANCHOR (locked): "Classify it. Describe it, one aspect at a time.
// Wrap it up. Facts all the way through."
//
// Lesson shape: example-first, then compare. The teacher's plan bundles
// introduction, body paragraphs, conclusion and visuals into one session, so
// the scope here is ONE idea - each part does one job - with the depth on
// classification and topic sentences. Visuals are named on the map today and
// taught properly in Week 10 Session 3, which is where students make one.
//
// Sources: teacher's unit plan, including the bicycle model sentence quoted
// verbatim. Model reports are original (see inforep_lib.js).

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

const P = require("./inforep_lib");

const UNIT_VARIANT = 2;
const T = createTheme("literacy", "grade56", UNIT_VARIANT);
const {
  C, FONT_H, FONT_B, SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, keyWordSlide, closingSlide,
  addCard, addFooter, addTextOnShape, addRevealAnswerBar,
  clickBuild, runSlideDiagnostics,
} = T;

const SESSION = 2;
const FOOTER = "Information Reports | Week 9 Session 2 | Year 5/6 Literacy";
const OUT_DIR = path.join(__dirname, "..", "output", "InfoReport_W9_S2");
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const SHEET_RES = makeSessionResource(SESSION,
  "Deconstruction Sheet",
  "Label the parts of a report, find the topic sentences, read the caption.");
const KEY_RES = makeSessionResource(SESSION,
  "Answer Key",
  "Labelled answers for the deconstruction sheet, with what to look for.");
const RESOURCE_ITEMS = [SHEET_RES, KEY_RES];
fs.mkdirSync(RES_DIR, { recursive: true });

/* --- Teacher notes (Glance Format, megaprompt sections 45-47) ------------- */

const NOTES_TITLE =
  "Session 2. Yesterday we worked out what a report is. Today we pull one apart and name the job of every part.";

const NOTES_RESOURCES =
  "Prep slide. Print the Deconstruction Sheet, one per student. Keep the Answer Key for yourself.\n" +
  "Students need their Information Report Map Card from Session 1, glued in books. Spares are in the Session 1 folder.\n" +
  "CATCH-UP: missed Session 1? Hand over the map card, read the launch text aloud with them, and start at Task 1 on the sheet. Nothing today needs yesterday's exit ticket.";

const NOTES_LAUNCH = composeGlanceNotes({
  answer: "open - listen for parts, headings or sections.",
  beats: [
    [
      "READ the whole report aloud. Students follow on screen.",
      "SAY: Listen for where one idea stops and the next starts.",
    ],
    [
      "ASK: How many separate ideas did you hear?",
      "20 sec. Fingers at your chest, show me.",
      "EXPECT: four or five.",
    ],
    [
      "SCAN the fingers.",
      "80%+ -> SAY: Good. Every one of those is a part with a job.",
      "Less -> re-read, pause hard between paragraphs, re-ask.",
    ],
  ],
  trap: [
    "counting sentences instead of ideas.",
    "Fix: read one paragraph, ask what it is about, student recounts.",
  ],
  prep: "Low-coupling launch: the whole text is on screen, so a student who missed Session 1 hears everything they need. Whole block under 6 minutes.",
  sources: P.SOURCE_LINE_MODEL,
  tag: "[Text Launch | Attention, focus and regulation | HITS 6]",
});

const NOTES_LI = composeGlanceNotes({
  beats: [
    [
      "POINT to the learning intention.",
      "SAY: Today we name the job that every part of a report does.",
    ],
    "SAY: Read the I can statements with me. Everyone, together, on three.",
    "SAY: The first one everyone leaves with. Point to the three parts.",
  ],
  prep: "SC1 is reachable by every student; SC2 is the core target the exit ticket assesses; SC3 stretches into writing a classification. Tier labels stay off the slide.",
  tag: "[LI/SC | Planning made visible | HITS 1]",
});

const NOTES_VOCAB_CLASSIFY = composeGlanceNotes({
  answer: "classify means to say what something is and what group it is in.",
  beats: [
    [
      "SAY: To classify is to put something in its group.",
      "A frog is an amphibian. A bicycle is a vehicle.",
    ],
    [
      "ASK: A dolphin. What group?",
      "10 sec. Choral response: everyone, together, on three.",
      "EXPECT: a mammal, or a sea animal.",
    ],
    "SAY: That is the whole job of the first paragraph. Say what it is.",
  ],
  trap: [
    "describing instead of classifying.",
    "Fix: ask what group, not what it looks like, student answers again.",
  ],
  prep: "The teacher's own template word. Every classification sentence students write for two weeks starts from this.",
  sources: P.SOURCE_LINE_UNIT,
  tag: "[Key Vocabulary | Knowledge and memory | HITS 6]",
});

const NOTES_VOCAB_ASPECT = composeGlanceNotes({
  answer: "an aspect is one part of the topic, such as habitat or diet.",
  beats: [
    [
      "SAY: An aspect is one slice of the topic.",
      "Appearance is one. Habitat is another. Diet is another.",
    ],
    [
      "ASK: Name one more aspect a frog report could have.",
      "20 sec. Turn and tell, partner B first.",
      "EXPECT: life cycle, predators, sounds, size.",
    ],
    "SAY: One aspect, one paragraph. That is the rule for the middle.",
  ],
  trap: [
    "listing facts instead of naming an aspect.",
    "Fix: ask what heading would sit above it, student names the aspect.",
  ],
  prep: "The teacher's own template word for the body paragraphs. Aspect and paragraph stay one to one for the whole unit.",
  sources: P.SOURCE_LINE_UNIT,
  tag: "[Key Vocabulary | Knowledge and memory | HITS 6]",
});

const NOTES_MAP = composeGlanceNotes({
  answer: "classify, describe one aspect at a time, wrap up, facts throughout.",
  beats: [
    [
      "SAY: Here is the map for every report you will read or write.",
      "Say the anchor with me.",
    ],
    [
      "CLICK through the four parts. Match each to the report we just read.",
      "SAY: This paragraph does this job. Nothing else.",
    ],
    [
      "ASK: Say the three parts back to me.",
      "10 sec. Choral response: everyone, together, on three.",
      "EXPECT: classify, describe, wrap up.",
    ],
  ],
  trap: [
    "reciting the names without knowing each job.",
    "Fix: point at one band, student says its job.",
  ],
  stretch: "Which part would be hardest to write? Say why.",
  help: "Open the map card. Point to each band as it appears.",
  prep: "The unit anchor. Say this phrase word for word in every I Do, both weeks, so the language never drifts.",
  tag: "[I Do | Explicit teaching | SC1 | HITS 3, 4]",
});

const NOTES_CLASSIFY_MODEL = composeGlanceNotes({
  answer: "name it, name its group, add one true thing.",
  beats: [
    [
      "SAY: Watch the first sentence do three jobs at once.",
      "Name it. Name its group. Add one true thing.",
    ],
    [
      "POINT to each coloured part.",
      "SAY: Green tree frog. Amphibian. Lives in northern Australia.",
    ],
    [
      "CLICK to the bicycle sentence. Read it aloud.",
      "SAY: Different topic, same three jobs.",
    ],
    [
      "ASK: What is missing from Frogs are cool?",
      "20 sec. Turn and tell, partner A first.",
      "EXPECT: the group, and a checkable fact.",
    ],
  ],
  trap: [
    "writing a title instead of a sentence.",
    "Fix: ask for a verb, student rewrites it.",
  ],
  stretch: "Write a classification sentence for your topic.",
  help: "Give the frame with the topic filled.",
  prep: "The bicycle sentence is the teacher's own model, quoted from the unit plan. Keep it word for word.",
  sources: P.SOURCE_LINE_UNIT,
  tag: "[I Do | Explicit teaching | SC3 | HITS 4]",
});

const NOTES_CFU = composeGlanceNotes({
  answer: "B. It names the aspect, habitat, in the first sentence.",
  beats: [
    [
      "SAY: A topic sentence tells the reader which aspect is coming.",
      "Do not call out. This one goes on boards.",
    ],
    [
      "ASK: Which one opens a paragraph about habitat?",
      "30 sec. Write it... chin it... show me.",
      "EXPECT: B.",
    ],
    [
      "SCAN boards, back row first.",
      "80%+ -> cold call one B: which word told you? Then reveal.",
      "Less -> read C, ask which aspect it names, re-ask.",
    ],
    [
      "REVEAL after boards are scanned.",
      "SAY: Name the aspect first. Then the details follow.",
    ],
  ],
  trap: [
    "picking A because it sounds like a strong opening.",
    "Fix: ask which aspect A names, student re-tests all three.",
  ],
  prep: "The hinge of the lesson. A names no aspect and is an opinion; C names the wrong aspect. Both wrong answers are diagnostic.",
  tag: "[CFU hinge | Supported application | SC2 | HITS 7, 8]",
});

const NOTES_WEDO = composeGlanceNotes({
  answer: "A why it is later, B distance, C how it forms.",
  beats: [
    [
      "SAY: New topic. Three paragraphs about thunder.",
      "The headings have been taken off.",
    ],
    [
      "ASK: Which heading belongs to each paragraph?",
      "60 sec. Write all three, chin it, show me.",
      "EXPECT: A later, B distance, C forms.",
    ],
    [
      "SCAN boards for all three matched.",
      "80%+ -> cold call one: which word gave B away?",
      "Less -> read A again, ask what it explains, re-ask.",
    ],
  ],
  trap: [
    "matching on one word, not the whole paragraph.",
    "Fix: read it again, student checks the whole idea.",
  ],
  stretch: "Write a fourth heading this report could use.",
  help: "Give one match free. Ask only for the other two.",
  prep: "Different topic from the I Do on purpose. Matching heading to paragraph is the reverse of writing a topic sentence.",
  sources: P.SOURCE_LINE_MODEL,
  tag: "[We Do | Supported application | SC2 | HITS 4, 5]",
});

const NOTES_YOUDO = composeGlanceNotes({
  answer: "classify, three aspects, wrap up, in that order down the page.",
  beats: [
    [
      "SAY: Your turn, on the sheet. A new report, the Great Barrier Reef.",
      "Label every part, then underline each topic sentence.",
    ],
    [
      "POINT to the steps. TIME: twelve minutes.",
      "Map card open beside the sheet.",
    ],
    [
      "CIRCULATE. Look for labels in the right order first.",
      "COLLECT two sheets to show at the closing.",
    ],
  ],
  trap: [
    "underlining the most interesting sentence, not the topic sentence.",
    "Fix: ask which sentence names the aspect, student re-underlines.",
  ],
  stretch: "Task 5: name an aspect the reef report is missing.",
  help: "Labels are listed at the top of the sheet. Match, do not recall.",
  prep: "Different text and different task from the We Do. Section 1 of the sheet is doable from today's launch alone, so a returning student can start.",
  sources: P.SOURCE_LINE_MODEL,
  tag: "[You Do | Mastery and application | SC2 | HITS 10]",
});

const NOTES_EXIT = composeGlanceNotes({
  answer: "any sentence naming volcanoes, their group, and one true thing.",
  beats: [
    [
      "SAY: One sentence. A new topic, volcanoes.",
      "Name it, name its group, add one true thing.",
    ],
    [
      "TIME: three minutes. Books open on the desk when you finish.",
      "COLLECT by walking the rows, not by calling names.",
    ],
    [
      "SORT into two piles as you collect: group named, group missing.",
      "A thick second pile -> open Session 3 by rebuilding one together.",
    ],
  ],
  prep: "Assesses the core target: the job of the opening part. Sort into two piles as you collect.",
  tag: "[Exit Ticket | Mastery and application | SC2 | HITS 8]",
});

const NOTES_CLOSING = composeGlanceNotes({
  answer: "open - most students should show two.",
  beats: [
    "SHOW the two collected sheets. SAY: Look at the order of the labels.",
    [
      "ASK: Rate yourself on the three statements.",
      "15 sec. Fingers at your chest, one, two or three. Show me.",
      "EXPECT: most on two.",
    ],
    "SAY: Next session we go and find the facts that fill those parts.",
  ],
  prep: "Self-assessment feeds Session 3 grouping. Anyone showing one works beside you during the Session 3 You Do.",
  tag: "[Closing | Retention and recall | HITS 1, 9]",
});

/* --- Slide content -------------------------------------------------------- */

// The teacher's own model sentence, quoted from the unit plan (section 5b).
const BICYCLE_SENTENCE = "The bicycle is a human-powered vehicle with two wheels.";

const CFU_OPTIONS = [
  { key: "A", text: "Frogs are amazing animals that everyone should know about." },
  { key: "B", text: "Green tree frogs live in warm, damp places near still water." },
  { key: "C", text: "A green tree frog can swallow a moth in one gulp." },
];

// Paragraphs shown out of order so the match cannot be made by position.
const WEDO_PARAGRAPHS = [
  { key: "A", text: P.THUNDER.aspects[1].text, heading: P.THUNDER.aspects[1].heading },
  { key: "B", text: P.THUNDER.aspects[2].text, heading: P.THUNDER.aspects[2].heading },
  { key: "C", text: P.THUNDER.aspects[0].text, heading: P.THUNDER.aspects[0].heading },
];

/* --- Build ---------------------------------------------------------------- */

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // 1. Title
  titleSlide(pres, "The Parts of a Report",
    P.ANCHOR_SHORT,
    "Week 9 Session 2 | Year 5/6 Literacy", NOTES_TITLE);

  // 2. Teacher Resources
  addResourceSlide(pres, {
    resources: RESOURCE_ITEMS,
    studentTools: [
      "Mini-whiteboards and markers",
      "The map card from Session 1",
      "A highlighter or coloured pencil for underlining",
    ],
    boardSetup: [
      "Write the three part names down one side of the board",
      "Leave room to build a classification sentence live",
    ],
  }, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // 3. Text launch - the whole model report, unlabelled
  (() => {
    const s = P.customSlide(pres, T, "Text Launch", C.SECONDARY, P.FROG.title);
    P.drawReportExtract(s, T, {
      x: 0.5, y: CONTENT_TOP, w: 9, h: SAFE_BOTTOM - CONTENT_TOP,
      bodySize: 13.5,
      showLabels: false,
      sections: [
        { text: P.FROG.classify },
        { text: P.FROG.aspects[0].text },
        { text: P.FROG.aspects[1].text },
        { text: P.FROG.aspects[2].text },
        { text: P.FROG.wrap },
      ],
    });
    addFooter(s, FOOTER);
    s.addNotes(NOTES_LAUNCH);
    runSlideDiagnostics(s, pres);
  })();

  // 4. Learning Intention and Success Criteria
  liSlide(pres,
    "We are learning how an information report is built, part by part.",
    [
      "I can point to the three parts of an information report.",
      "I can find the topic sentence that starts a body paragraph.",
      "I can write a classification sentence that opens a report.",
    ],
    NOTES_LI, FOOTER);

  // 5. Key vocabulary - classify
  keyWordSlide(pres, {
    word: "classify",
    meaning: "To say what something is and what group it belongs to.",
    example: "A green tree frog is an amphibian. A bicycle is a vehicle.",
    routine: ["Say it", "Name a group", "Use it"],
    color: C.PRIMARY,
    title: "The word for the opening",
  }, NOTES_VOCAB_CLASSIFY, FOOTER);

  // 6. Key vocabulary - aspect
  keyWordSlide(pres, {
    word: "aspect",
    meaning: "One part of the topic, such as appearance, habitat or diet.",
    example: "One aspect gets one paragraph, with a heading above it.",
    routine: ["Say it", "Name one", "Use it"],
    color: C.SECONDARY,
    title: "The word for the middle",
  }, NOTES_VOCAB_ASPECT, FOOTER);

  // 7. I Do A - the map, built one part per click, against the report we read
  (() => {
    const s = P.customSlide(pres, T, "I Do", C.PRIMARY, "The map for every report");
    const y0 = CONTENT_TOP;
    const h = SAFE_BOTTOM - y0;
    const mapX = 0.5;
    const mapW = 4.15;
    const rx = 4.95;
    const rw = 4.55;

    // Row tops mirror drawReportMap's band arithmetic for y=1.3, h=3.8:
    // classify 1.30, describe 2.12, wrap 3.82, facts 4.64.
    const rows = [
      { y: 1.30, h: 0.72, color: C.PRIMARY, lines: ["The green tree frog is an amphibian that lives in northern Australia."] },
      { y: 2.12, h: 1.59, color: C.SECONDARY, lines: [
        "Appearance: smooth, bright green skin.",
        "Habitat: trees and reeds near still water.",
        "Diet: hunts at night, eats insects.",
      ] },
      { y: 3.82, h: 0.72, color: C.ASSESS, lines: ["A common but remarkable Australian animal."] },
      { y: 4.64, h: 0.46, color: C.ACCENT, lines: ["Present tense. No opinions. Technical words."] },
    ];

    function drawRow(i) {
      const row = rows[i];
      s.addShape("roundRect", {
        x: rx, y: row.y, w: rw, h: row.h, rectRadius: 0.08,
        fill: { color: C.WHITE }, line: { color: row.color, width: 1.4 },
      });
      s.addText(row.lines.map((line, k) => ({
        text: line,
        options: { breakLine: k < row.lines.length - 1 },
      })), {
        x: rx + 0.16, y: row.y, w: rw - 0.32, h: row.h,
        fontSize: i === 3 ? 12 : 13, fontFace: FONT_B, color: C.CHARCOAL,
        valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
        paraSpaceAfter: 2,
      });
    }

    const base = { x: mapX, y: y0, w: mapW, h, aspects: P.FROG_ASPECTS };
    clickBuild(s, [
      () => { P.drawReportMap(s, T, { ...base, parts: ["classify"] }); drawRow(0); },
      () => { P.drawReportMap(s, T, { ...base, parts: ["describe"] }); drawRow(1); },
      () => { P.drawReportMap(s, T, { ...base, parts: ["wrap"] }); drawRow(2); },
      () => { P.drawReportMap(s, T, { ...base, parts: ["facts"] }); drawRow(3); },
    ]);

    addFooter(s, FOOTER);
    s.addNotes(NOTES_MAP);
    runSlideDiagnostics(s, pres);
  })();

  // 8. I Do B - the classification sentence does three jobs
  (() => {
    const s = P.customSlide(pres, T, "I Do", C.PRIMARY, "The opening sentence does three jobs");
    const y0 = CONTENT_TOP;

    addCard(s, 0.5, y0, 9, 0.86, { strip: C.PRIMARY });
    s.addText([
      { text: "The ", options: { color: C.CHARCOAL } },
      { text: "[topic]", options: { color: C.PRIMARY, bold: true } },
      { text: " is a ", options: { color: C.CHARCOAL } },
      { text: "[group]", options: { color: C.SECONDARY, bold: true } },
      { text: " that ", options: { color: C.CHARCOAL } },
      { text: "[one true thing]", options: { color: C.ACCENT, bold: true } },
      { text: ".", options: { color: C.CHARCOAL } },
    ], {
      x: 0.8, y: y0 + 0.08, w: 8.4, h: 0.70,
      fontSize: 23, fontFace: FONT_B, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    const exY = y0 + 1.02;
    s.addShape("roundRect", {
      x: 0.5, y: exY, w: 9, h: 1.02, rectRadius: 0.08,
      fill: { color: C.WHITE }, line: { color: C.PRIMARY, width: 1.4 },
    });
    s.addText([
      { text: "The green tree frog", options: { color: C.PRIMARY, bold: true } },
      { text: " is ", options: { color: C.CHARCOAL } },
      { text: "an amphibian", options: { color: C.SECONDARY, bold: true } },
      { text: " that ", options: { color: C.CHARCOAL } },
      { text: "lives in the warm, wet parts of northern Australia", options: { color: C.ACCENT, bold: true } },
      { text: ".", options: { color: C.CHARCOAL } },
    ], {
      x: 0.8, y: exY + 0.06, w: 8.4, h: 0.90,
      fontSize: 19, fontFace: FONT_B, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    const bicY = exY + 1.16;
    clickBuild(s, [
      () => {
        s.addShape("roundRect", {
          x: 0.5, y: bicY, w: 9, h: 0.86, rectRadius: 0.08,
          fill: { color: C.WHITE }, line: { color: C.PRIMARY, width: 1.4 },
        });
        s.addText([
          { text: "The bicycle", options: { color: C.PRIMARY, bold: true } },
          { text: " is ", options: { color: C.CHARCOAL } },
          { text: "a human-powered vehicle", options: { color: C.SECONDARY, bold: true } },
          { text: " ", options: { color: C.CHARCOAL } },
          { text: "with two wheels", options: { color: C.ACCENT, bold: true } },
          { text: ".", options: { color: C.CHARCOAL } },
        ], {
          x: 0.8, y: bicY + 0.06, w: 8.4, h: 0.74,
          fontSize: 19, fontFace: FONT_B, valign: "middle", margin: 0,
          fit: "shrink", shrinkText: true,
        });
      },
      () => {
        addTextOnShape(s, "Name it. Name its group. Add one true thing.", {
          x: 0.5, y: bicY + 1.00, w: 9, h: 0.62, rectRadius: 0.08,
          fill: { color: C.PRIMARY },
        }, {
          fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true,
          align: "center", valign: "middle", margin: 0,
        });
      },
    ]);

    addFooter(s, FOOTER);
    s.addNotes(NOTES_CLASSIFY_MODEL);
    runSlideDiagnostics(s, pres);
  })();

  // 9. CFU hinge - which one opens a paragraph about habitat?
  (() => {
    const s = P.customSlide(pres, T, "CFU", C.ALERT, "Which one opens a habitat paragraph?");
    const y0 = CONTENT_TOP;

    addTextOnShape(s, "A topic sentence names the aspect that is coming.", {
      x: 0.5, y: y0, w: 9, h: 0.48, rectRadius: 0.08,
      fill: { color: C.ALERT },
    }, {
      fontSize: 16, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    P.drawOptionStack(s, T, CFU_OPTIONS, {
      y: y0 + 0.60, optionH: 0.66, gap: 0.12, color: C.PRIMARY, fontSize: 17,
    });

    clickBuild(s, [
      () => {
        addRevealAnswerBar(s, ["B - it names the aspect: where they live"], {
          y: P.REVEAL_Y, h: P.REVEAL_H, fontSize: 21, color: C.SUCCESS,
        });
      },
    ]);

    addFooter(s, FOOTER);
    s.addNotes(NOTES_CFU);
    runSlideDiagnostics(s, pres);
  })();

  // 10. We Do - match the heading to the paragraph
  (() => {
    const s = P.customSlide(pres, T, "We Do", C.SUCCESS, "Which heading belongs to each paragraph?");
    const y0 = CONTENT_TOP;
    const heights = [0.62, 0.78, 0.62];
    const gap = 0.10;

    let cy = y0;
    WEDO_PARAGRAPHS.forEach((para, i) => {
      const ph = heights[i];
      s.addShape("roundRect", {
        x: 0.5, y: cy, w: 9, h: ph, rectRadius: 0.08,
        fill: { color: C.WHITE }, line: { color: C.SUCCESS, width: 1.3 },
      });
      s.addShape("roundRect", {
        x: 0.68, y: cy + (ph - 0.38) / 2, w: 0.38, h: 0.38, rectRadius: 0.19,
        fill: { color: C.SUCCESS },
      });
      s.addText(para.key, {
        x: 0.68, y: cy + (ph - 0.38) / 2, w: 0.38, h: 0.38,
        fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      s.addText(para.text, {
        x: 1.20, y: cy + 0.04, w: 8.1, h: ph - 0.08,
        fontSize: 12.5, fontFace: FONT_B, color: C.CHARCOAL,
        valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
      cy += ph + gap;
    });

    const chipY = cy + 0.06;
    const chips = P.THUNDER.aspects.map((a) => a.heading);
    const chipGap = 0.14;
    const chipW = (9 - chipGap * 2) / 3;
    chips.forEach((label, i) => {
      addTextOnShape(s, label, {
        x: 0.5 + i * (chipW + chipGap), y: chipY, w: chipW, h: 0.52, rectRadius: 0.08,
        fill: { color: C.SECONDARY },
      }, {
        fontSize: 15, fontFace: FONT_B, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
    });

    clickBuild(s, [
      () => {
        addRevealAnswerBar(s, ["A = why it is later", "B = counting distance", "C = how it forms"], {
          y: P.REVEAL_Y, h: P.REVEAL_H, fontSize: 16, color: C.SUCCESS,
        });
      },
    ]);

    addFooter(s, FOOTER);
    s.addNotes(NOTES_WEDO);
    runSlideDiagnostics(s, pres);
  })();

  // 11. You Do - label a new report on the sheet
  (() => {
    const s = P.customSlide(pres, T, "You Do", C.ASSESS, "Label a report you have not seen");
    const y0 = CONTENT_TOP;
    const h = SAFE_BOTTOM - y0;

    T.addInstructionCard(s, [
      { text: "Deconstruction Sheet", role: "header" },
      { text: "The Great Barrier Reef." },
      { text: "1. Label every part in the boxes." },
      { text: "2. Underline each topic sentence." },
      { text: "3. Turn over: three writing tasks." },
      { text: "Twelve minutes. Map card open." },
    ], {
      x: 0.5, y: y0, w: 4.6, h,
      strip: C.ASSESS, fill: C.WHITE,
      headerColor: C.ASSESS, emphasisColor: C.ALERT,
    });

    P.drawReportMap(s, T, {
      x: 5.3, y: y0, w: 4.2, h,
      aspects: ["Aspect 1", "Aspect 2", "Aspect 3"],
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
    runSlideDiagnostics(s, pres);
  })();

  // 12. Exit ticket
  P.exitTicketPanel(pres, T, {
    topic: "New topic: volcanoes",
    task: "Write the first sentence of a report about volcanoes.",
    cue: "Name it. Name its group. Add one true thing.",
    taskSize: 27,
    footer: FOOTER,
    notes: NOTES_EXIT,
  });

  // 13. Closing
  closingSlide(pres, {
    reflectionPrompt: "Which part of a report will be easiest for you to write?",
    scItems: [
      "I can point to the three parts of an information report.",
      "I can find the topic sentence that starts a body paragraph.",
      "I can write a classification sentence that opens a report.",
    ],
    selfAssessment: "Fingers at your chest. One, two or three. Show me.",
    takeaways: [P.ANCHOR_PHRASE],
  }, NOTES_CLOSING);

  const pptxPath = path.join(OUT_DIR, "Information Reports W9 S2.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);
}

/* --- Resources ------------------------------------------------------------ */

function sheetBody(doc, isKey) {
  let y = addPdfHeader(doc, isKey ? "Deconstruction Sheet: Answer Key" : "Deconstruction Sheet", {
    subtitle: isKey
      ? "Labelled answers, with what to look for as you circulate."
      : "Label the parts. Find the topic sentences. Read the caption.",
    color: isKey ? C.ALERT : C.PRIMARY,
    lessonInfo: "Year 5/6 Literacy | Information Reports | Session 2",
    showNameDate: !isKey,
  });

  y = addTipBox(doc, P.ANCHOR_PHRASE, y, { color: C.ACCENT });

  y = addSectionHeading(doc, "Task 1: label every part in the box beside it", y,
    { color: isKey ? C.ALERT : C.PRIMARY });
  y = addBodyText(doc,
    "Use these labels: Classify, Aspect 1, Aspect 2, Aspect 3, Wrap up.", y);

  y = P.addReportTextPdf(doc, y, P.REEF, {
    color: hex(C.PRIMARY),
    labels: isKey ? ["Classify", "Aspect 1", "Aspect 2", "Aspect 3", "Wrap up"] : null,
  });

  y = addSectionHeading(doc, "Task 2: underline the topic sentence in each aspect", y,
    { color: isKey ? C.ALERT : C.SECONDARY });
  if (isKey) {
    P.REEF.aspects.forEach((aspect, i) => {
      doc.fontSize(10.5).font("Sans-Bold").fillColor(hex(C.SECONDARY));
      doc.text("Aspect " + (i + 1) + ": ", PAGE.MARGIN, y, { continued: true });
      doc.font("Sans").fillColor("#000000");
      doc.text(aspect.topicSentence);
      y = doc.y + 6;
    });
  } else {
    y = addBodyText(doc,
      "Underline it on the text above. The topic sentence names the aspect.", y, { italic: true });
  }

  // Page 1 is annotation ON the text. Page 2 is the writing, with real space
  // for it, rather than three cramped lines squeezed under the report.
  doc.addPage();
  y = PAGE.MARGIN;

  y = addSectionHeading(doc, "Task 3: the caption", y, { color: isKey ? C.ALERT : C.ASSESS });
  y = addBodyText(doc,
    "What does the caption tell you that the words of the report do not?", y);
  if (isKey) {
    y = addBodyText(doc,
      "Look for: it tells you WHERE the reef is, which none of the paragraphs state. Accept any answer that names something only the map can show.",
      y, { italic: true, color: "2D6B4A" });
    y += 8;
  } else {
    y = addLinedArea(doc, y + 6, 4, { lineSpacing: 30 });
    y += 16;
  }

  y = addSectionHeading(doc, "Task 4: rewrite the opening", y, { color: isKey ? C.ALERT : C.SUCCESS });
  y = addBodyText(doc,
    "Write the opening sentence again in your own words. Name it, name its group, add one true thing.", y);
  if (isKey) {
    y = addBodyText(doc,
      "Look for: a group word such as coral reef, reef system or natural wonder, plus one checkable fact. A sentence with no group word is not yet a classification.",
      y, { italic: true, color: "2D6B4A" });
    y += 8;
  } else {
    y = addLinedArea(doc, y + 6, 5, { lineSpacing: 30 });
    y += 16;
  }

  y = addSectionHeading(doc, "Task 5: the missing aspect", y, { color: isKey ? C.ALERT : C.PRIMARY });
  y = addBodyText(doc,
    "This report has three aspects. Name one more aspect it could have had, then write the topic sentence that would start it.", y);
  if (isKey) {
    y = addBodyText(doc,
      "Look for: a named aspect (how it formed, tourism, protection, coral itself) AND a topic sentence that states it. A fact on its own is not a topic sentence.",
      y, { italic: true, color: "2D6B4A" });
  } else {
    y = addLinedArea(doc, y + 6, 5, { lineSpacing: 30 });
  }

  addPdfFooter(doc, "Information Reports | Year 5/6 Literacy | Session 2");
}

async function buildSheet() {
  const doc = createPdf({ title: SHEET_RES.name });
  sheetBody(doc, false);
  await writePdf(doc, path.join(OUT_DIR, SHEET_RES.fileName));
  console.log("Wrote " + SHEET_RES.name);
}

async function buildKey() {
  const doc = createPdf({ title: KEY_RES.name });
  sheetBody(doc, true);
  await writePdf(doc, path.join(OUT_DIR, KEY_RES.fileName));
  console.log("Wrote " + KEY_RES.name);
}

(async () => {
  await build();
  await buildSheet();
  await buildKey();
})().catch((err) => {
  console.error(err);
  process.exit(1);
});
