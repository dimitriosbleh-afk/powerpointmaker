"use strict";

// Information Reports - Term 3 Week 10, Session 1 (Year 5/6 Literacy)
// "Pick a topic. Make a plan. Start writing."
//
// Victorian Curriculum 2.0: English, Literacy, Levels 5-6 - VC2E5LY10 and
// VC2E6LY09, creating texts with organised ideas and a text structure suited
// to topic, purpose and audience.
//
// UNIT ANCHOR (locked, week 2 of 2): "Classify it. Describe it, one aspect at
// a time. Wrap it up. Facts all the way through." The planning template IS
// that map, which is why the plan and the anchor use the same words.
//
// Lesson shape: example-first, with the modelling split into three short
// cycles (choose, plan, draft) because the teacher's plan puts planning AND
// the start of drafting in one session.
//
// Sources: teacher's unit plan - the broad class themes, the planning template
// wording, and the class topic used for all modelling.

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
  titleSlide, liSlide, closingSlide,
  addCard, addFooter, addTextOnShape, addRevealAnswerBar,
  clickBuild, runSlideDiagnostics,
} = T;

const SESSION = 1;
const FOOTER = "Information Reports | Week 10 Session 1 | Year 5/6 Literacy";
const OUT_DIR = path.join(__dirname, "..", "output", "InfoReport_W10_S1");
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const PLAN_RES = makeSessionResource(SESSION,
  "Report Plan",
  "The planning template, with the report map on it. Used for the rest of the week.");
const RESCUE_RES = makeSessionResource(SESSION,
  "Research Rescue Card",
  "For a student whose research stalls: worked notes, question stems and a fact bank.");
const RESOURCE_ITEMS = [PLAN_RES, RESCUE_RES];
fs.mkdirSync(RES_DIR, { recursive: true });

/* --- Teacher notes (Glance Format, megaprompt sections 45-47) ------------- */

const NOTES_TITLE =
  "Week 10. Last week we read reports like writers. This week each student writes and publishes one.";

const NOTES_RESOURCES =
  "Prep slide. Print the Report Plan, one per student. Print a few Research Rescue Cards and keep them at your desk.\n" +
  "Book the library or a device trolley for the second half. Have last week's map cards out.\n" +
  "CATCH-UP: missed Week 9? The launch rebuilds the whole map, and the plan sheet carries it too. Hand over a Rescue Card so research does not become the blocker.";

const NOTES_OVERVIEW =
  "For you, not the class. Skip past it when projecting.\n" +
  "Four sessions: plan and start drafting, finish drafting, visuals and editing, publish and share. The teacher's plan listed five; publishing and sharing are combined on Friday.\n" +
  "Decision points today: the topic hinge, the boards check after the plan, and the exit ticket. Between them keep the pace brisk.";

const NOTES_LAUNCH = composeGlanceNotes({
  answer: "classify, describe one aspect at a time, wrap up, facts throughout.",
  beats: [
    [
      "SAY: Before we write, say the map back to me.",
      "Everyone, together, on three.",
    ],
    [
      "ASK: What is the job of the middle?",
      "20 sec. Turn and tell, partner A first.",
      "EXPECT: one aspect per paragraph.",
    ],
    [
      "SCAN the room. Listen for the word aspect.",
      "80%+ -> reveal the map, then move on.",
      "Less -> point to the map card, read each band, re-ask.",
    ],
    [
      "REVEAL after the partner talk.",
      "SAY: That is your plan. The plan IS the map.",
    ],
  ],
  trap: [
    "naming the parts but not their jobs.",
    "Fix: point at one band, student says its job.",
  ],
  prep: "Low-coupling launch: the map is rebuilt from scratch on screen, so a student who missed all of Week 9 can join here. Whole block under 6 minutes.",
  tag: "[Launch | Retention and recall | HITS 6]",
});

const NOTES_LI = composeGlanceNotes({
  beats: [
    [
      "POINT to the learning intention.",
      "SAY: Today you choose your topic and build your plan.",
    ],
    "SAY: Read the I can statements with me. Everyone, together, on three.",
    "SAY: The first one everyone leaves with. A topic you can find facts about.",
  ],
  prep: "SC1 is reachable by every student; SC2 is the core target the exit ticket assesses; SC3 stretches into drafting. Tier labels stay off the slide.",
  tag: "[LI/SC | Planning made visible | HITS 1]",
});

const NOTES_TOPIC = composeGlanceNotes({
  answer: "just right is a whole group you can split into three aspects.",
  beats: [
    [
      "SAY: A topic can be too big, too small, or just right.",
      "Watch me test three.",
    ],
    [
      "POINT to each row. SAY: Animals is a library. Rex is a story.",
      "Emperor penguins is a group. That one works.",
    ],
    [
      "ASK: Two tests. What are they?",
      "20 sec. Turn and tell, partner B first.",
      "EXPECT: can I find facts, can I split it into aspects.",
    ],
  ],
  trap: [
    "choosing a pet or a favourite thing.",
    "Fix: ask if it is a whole group, student widens it.",
  ],
  stretch: "Take a too-big topic and narrow it twice.",
  help: "Offer three ready-made topics from the theme list.",
  prep: "Say the anchor once here. The four broad themes are the teacher's own, from the unit plan.",
  sources: P.SOURCE_LINE_UNIT,
  tag: "[I Do | Explicit teaching | SC1 | HITS 3, 4]",
});

const NOTES_CFU = composeGlanceNotes({
  answer: "B. Tropical cyclones is a group you can split into aspects.",
  beats: [
    [
      "SAY: Three topics from the weather theme. One is just right.",
      "Do not call out. This one goes on boards.",
    ],
    [
      "ASK: Which topic would work for a report?",
      "30 sec. Boards up on cue.",
      "EXPECT: B.",
    ],
    [
      "SCAN boards, back row first.",
      "80%+ -> cold call one B: name an aspect for it.",
      "Less -> read A, ask how many books it would take, re-ask.",
    ],
    [
      "REVEAL after boards are scanned.",
      "SAY: A is too big. C is one event, so it is a story.",
    ],
  ],
  trap: [
    "picking C because it sounds interesting.",
    "Fix: ask if it happened once, student re-tests.",
  ],
  prep: "The hinge of the lesson. A is too broad and C is a single event, which are the two ways a topic choice fails.",
  tag: "[CFU hinge | Supported application | SC1 | HITS 7, 8]",
});

const NOTES_PLAN = composeGlanceNotes({
  answer: "notes, not sentences, in every aspect row.",
  beats: [
    [
      "SAY: Watch me fill the plan for our class topic.",
      "Row by row, and only in notes.",
    ],
    [
      "CLICK through the rows. SAY: Topic. Then the group it is in.",
      "Then three aspects, each with dot points.",
    ],
    [
      "ASK: Why notes and not sentences in the plan?",
      "20 sec. Turn and tell, partner A first.",
      "EXPECT: so we write it in our own words later.",
    ],
  ],
  trap: [
    "writing full sentences into the plan.",
    "Fix: ring one sentence, student cuts it back to notes.",
  ],
  stretch: "Add a fourth aspect row to your own plan.",
  help: "The plan sheet already names three aspects. Fill those.",
  prep: "The plan template is the teacher's own, and it is the same map as the anchor. Keep the wording identical.",
  sources: P.SOURCE_LINE_UNIT,
  tag: "[I Do | Explicit teaching | SC2 | HITS 4]",
});

const NOTES_DRAFT = composeGlanceNotes({
  answer: "the dot points become a topic sentence and two detail sentences.",
  beats: [
    [
      "SAY: Now watch the notes turn into a paragraph.",
      "I cover the plan first, so I use my own words.",
    ],
    [
      "CLICK to the topic sentence. SAY: It names the aspect.",
      "CLICK again. Two details, and a noun group in each.",
    ],
    [
      "ASK: Which sentence names the aspect?",
      "15 sec. Write it... chin it... show me.",
      "EXPECT: the first one.",
    ],
  ],
  trap: [
    "copying the dot points in as a list.",
    "Fix: cover the plan, student says the sentence aloud first.",
  ],
  stretch: "Add a technical word to the paragraph.",
  help: "Give the topic sentence. Student adds one detail.",
  prep: "The drafting think-aloud the unit plan asks for. Say the words you are choosing and why, out loud, as you write.",
  tag: "[I Do | Explicit teaching | SC3 | HITS 4, 9]",
});

const NOTES_WEDO = composeGlanceNotes({
  answer: "open - three aspects such as how they form, where, and damage.",
  beats: [
    [
      "SAY: Same topic for everyone. Tropical cyclones.",
      "You choose the three aspects.",
    ],
    [
      "ASK: What three aspects would this report have?",
      "60 sec. Write three... chin it... show me.",
      "EXPECT: three separate aspects, not three facts.",
    ],
    [
      "SCAN boards for three DIFFERENT aspects.",
      "80%+ -> cold call one: which would be hardest to research?",
      "Less -> name one aspect yourself, ask for two more, re-ask.",
    ],
  ],
  trap: [
    "writing three facts instead of three aspects.",
    "Fix: ask what heading sits above it, student names the aspect.",
  ],
  stretch: "Write the topic sentence for one of your aspects.",
  help: "Give two aspects. Ask for the third.",
  prep: "One shared topic here so the thinking is about aspects, not research. Their own topic comes in the You Do.",
  sources: P.SOURCE_LINE_UNIT,
  tag: "[We Do | Supported application | SC2 | HITS 4, 5]",
});

const NOTES_YOUDO = composeGlanceNotes({
  answer: "a topic, three named aspects, and notes in each row.",
  beats: [
    [
      "SAY: Your topic, from one of the four themes.",
      "Plan first. Notes only. Then start your opening.",
    ],
    [
      "POINT to the steps. TIME: twenty-five minutes.",
      "Library books and screens open once the plan has three aspects.",
    ],
    [
      "CIRCULATE. Check the topic passes both tests before they research.",
      "COLLECT two plans to show at the closing.",
    ],
  ],
  trap: [
    "starting to research before the aspects are named.",
    "Fix: cover the book, ask for three aspects, student names them.",
  ],
  stretch: "Draft the whole first aspect paragraph.",
  help: "Hand over a Research Rescue Card with the fact bank.",
  prep: "Their own topic, unlike the We Do. A student with no facts gets the Rescue Card, so writing is never blocked by research.",
  tag: "[You Do | Mastery and application | SC2 | HITS 10]",
});

const NOTES_EXIT = composeGlanceNotes({
  answer: "a sentence naming their topic, its group, and one true thing.",
  beats: [
    [
      "SAY: One sentence. Your own topic this time.",
      "Name it, name its group, add one true thing.",
    ],
    [
      "TIME: three minutes. Write it at the top of your plan.",
      "COLLECT the plans as you walk the rows.",
    ],
    [
      "SORT into two piles as you collect: group named, group missing.",
      "A thick second pile -> re-model the opening at the start of Session 2.",
    ],
  ],
  prep: "Assesses the core target on their own topic. Collecting the plan means you can see the aspects and the opening together.",
  tag: "[Exit Ticket | Mastery and application | SC2 | HITS 8]",
});

const NOTES_CLOSING = composeGlanceNotes({
  answer: "open - most students should show two.",
  beats: [
    "SHOW the two collected plans. SAY: Look at the three aspect rows.",
    [
      "ASK: Rate yourself on the three statements.",
      "15 sec. Fingers at your chest, one, two or three. Show me.",
      "EXPECT: most on two.",
    ],
    "SAY: Next session we turn every one of those rows into a paragraph.",
  ],
  prep: "Anyone showing one starts Session 2 beside you with a Rescue Card. Keep the collected plans to hand back at the door.",
  tag: "[Closing | Retention and recall | HITS 1, 9]",
});

/* --- Slide content -------------------------------------------------------- */

const CFU_OPTIONS = [
  { key: "A", text: "Weather" },
  { key: "B", text: "Tropical cyclones" },
  { key: "C", text: "The storm that hit our school last winter" },
];

// The class topic plan, filled the way students should fill theirs: notes only.
const FROG_PLAN = [
  { label: "Topic", text: "Green tree frog", color: C.PRIMARY },
  { label: "Introduction", text: "amphibian / warm, wet places / northern Australia / one of our largest tree frogs", color: C.PRIMARY },
  { label: "Aspect 1", text: "Appearance: smooth green skin / white belly / about 10 cm / wide toe pads", color: C.SECONDARY },
  { label: "Aspect 2", text: "Habitat: trees, reeds, rocks near still water / downpipes and water tanks", color: C.SECONDARY },
  { label: "Aspect 3", text: "Diet: hunts at night / moths, crickets, cockroaches / sticky tongue", color: C.SECONDARY },
  { label: "Conclusion", text: "common but remarkable / skin chemicals kill germs", color: C.ASSESS },
];

const DRAFT_NOTES = ["smooth green skin", "white belly", "about 10 cm", "wide toe pads"];

const DRAFT_SENTENCES = [
  { label: "Topic sentence", text: P.FROG.aspects[0].topicSentence, color: C.PRIMARY },
  { label: "Details", text: "An adult can grow to about ten centimetres long. Wide, round pads on its toes help it grip and climb.", color: C.SECONDARY },
];

/* --- Build ---------------------------------------------------------------- */

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // 1. Title
  titleSlide(pres, "Plan Your Report",
    "Pick a topic. Make a plan. Start writing.",
    "Week 10 Session 1 | Year 5/6 Literacy", NOTES_TITLE);

  // 2. Teacher Resources
  addResourceSlide(pres, {
    resources: RESOURCE_ITEMS,
    studentTools: [
      "The Report Plan, one each",
      "Pencils, not pens, for planning",
      "The map card from Week 9",
      "Mini-whiteboards for the two checks",
    ],
    boardSetup: [
      "Book the library or a device trolley for the second half",
      "A few Rescue Cards at your desk, not handed out yet",
    ],
  }, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // 3. Teacher-facing overview
  (() => {
    const s = P.customSlide(pres, T, "For the teacher", C.MUTED, "Week 10 at a glance");
    const cardY = CONTENT_TOP;
    const cardH = SAFE_BOTTOM - CONTENT_TOP;

    addCard(s, 0.5, cardY, 4.4, cardH, { strip: C.PRIMARY });
    s.addText("This week", {
      x: 0.7, y: cardY + 0.12, w: 4.0, h: 0.28,
      fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText([
      { text: "S1  Choose a topic, plan, start the opening", options: { breakLine: true } },
      { text: "S2  Draft the body paragraphs and conclusion", options: { breakLine: true } },
      { text: "S3  Visuals, captions, then editing", options: { breakLine: true } },
      { text: "S4  Publish and share", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "The plan listed five sessions. Publishing and sharing are combined on Friday.", options: { italic: true, color: C.MUTED, breakLine: true } },
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
      { text: "Two objectives in one session: plan, then start drafting. The modelling is split into three short cycles so neither is rushed.", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "Decision points:", options: { bold: true, color: C.ACCENT, breakLine: true } },
      { text: "the topic hinge, the boards check after the plan, and the exit ticket.", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "Research:", options: { bold: true, color: C.ACCENT, breakLine: true } },
      { text: "library books and screens open only once a plan has three named aspects. A Rescue Card unblocks anyone with no facts.", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "You model on the green tree frog all week. Students write their own topic.", options: { italic: true, color: C.MUTED } },
    ], {
      x: 5.3, y: cardY + 0.46, w: 4.0, h: cardH - 0.60,
      fontSize: 12.5, fontFace: FONT_B, color: C.CHARCOAL,
      valign: "top", margin: 0, paraSpaceAfter: 3,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_OVERVIEW);
    runSlideDiagnostics(s, pres);
  })();

  // 4. Launch - rebuild the map from memory
  (() => {
    const s = P.customSlide(pres, T, "Launch", C.SECONDARY, "Say the map back to me");
    const y0 = CONTENT_TOP;
    const h = SAFE_BOTTOM - y0;
    const base = { x: 2.2, y: y0, w: 5.6, h, aspects: P.GENERIC_ASPECTS };

    clickBuild(s, [
      () => P.drawReportMap(s, T, { ...base, parts: ["classify"] }),
      () => P.drawReportMap(s, T, { ...base, parts: ["describe"] }),
      () => P.drawReportMap(s, T, { ...base, parts: ["wrap"] }),
      () => P.drawReportMap(s, T, { ...base, parts: ["facts"] }),
    ]);

    addFooter(s, FOOTER);
    s.addNotes(NOTES_LAUNCH);
    runSlideDiagnostics(s, pres);
  })();

  // 5. Learning Intention and Success Criteria
  liSlide(pres,
    "We are learning to choose a topic and plan an information report.",
    [
      "I can choose a topic I can find facts about.",
      "I can sort my facts into three aspects on a plan.",
      "I can draft the sentence that opens my report.",
    ],
    NOTES_LI, FOOTER);

  // 6. I Do A - too big, too small, just right
  (() => {
    const s = P.customSlide(pres, T, "I Do", C.PRIMARY, "Too big, too small, just right");
    const y0 = CONTENT_TOP;
    // Three rows, a test bar and a theme chip row all have to fit above 5.10:
    // 1.30 + 3 x 0.80 = 3.70, bar to 4.22, chips 4.34 to 4.80.
    const rowH = 0.68;
    const gap = 0.12;
    const verdictColor = { "too big": C.ALERT, "too small": C.ALERT, "just right": C.SUCCESS };

    P.TOPIC_TESTS.forEach((item, i) => {
      const ry = y0 + i * (rowH + gap);
      const accent = verdictColor[item.verdict];
      s.addShape("roundRect", {
        x: 0.5, y: ry, w: 9, h: rowH, rectRadius: 0.08,
        fill: { color: C.WHITE }, line: { color: accent, width: 1.5 },
      });
      s.addText(item.topic, {
        x: 0.76, y: ry, w: 2.9, h: rowH,
        fontSize: 19, fontFace: FONT_B, color: C.CHARCOAL, bold: true,
        valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
      addTextOnShape(s, item.verdict, {
        x: 3.75, y: ry + 0.14, w: 1.55, h: rowH - 0.28, rectRadius: 0.07,
        fill: { color: accent },
      }, {
        fontSize: 13, fontFace: FONT_B, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      s.addText(item.why, {
        x: 5.50, y: ry, w: 3.75, h: rowH,
        fontSize: 15, fontFace: FONT_B, color: C.CHARCOAL,
        valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
    });

    const testY = y0 + 3 * (rowH + gap);
    const themeY = testY + 0.64;
    addTextOnShape(s, "Two tests: can I find facts? Can I split it into three aspects?", {
      x: 0.5, y: testY, w: 9, h: 0.52, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, {
      fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    const themeGap = 0.14;
    const themeW = (9 - themeGap * 3) / 4;
    P.CLASS_THEMES.forEach((theme, i) => {
      addTextOnShape(s, theme, {
        x: 0.5 + i * (themeW + themeGap), y: themeY, w: themeW, h: 0.46, rectRadius: 0.07,
        fill: { color: C.WHITE }, line: { color: C.ACCENT, width: 1.4 },
      }, {
        fontSize: 13, fontFace: FONT_B, color: C.ACCENT, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_TOPIC);
    runSlideDiagnostics(s, pres);
  })();

  // 7. CFU hinge - which topic would work?
  (() => {
    const s = P.customSlide(pres, T, "CFU", C.ALERT, "Which topic would work for a report?");
    const y0 = CONTENT_TOP;

    addTextOnShape(s, "Theme: Extreme Weather", {
      x: 0.5, y: y0, w: 9, h: 0.48, rectRadius: 0.08,
      fill: { color: C.ALERT },
    }, {
      fontSize: 16, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    P.drawOptionStack(s, T, CFU_OPTIONS, {
      y: y0 + 0.60, optionH: 0.66, gap: 0.12, color: C.PRIMARY, fontSize: 18,
    });

    clickBuild(s, [
      () => {
        addRevealAnswerBar(s, ["B - a whole group, and it splits into aspects"], {
          y: P.REVEAL_Y, h: P.REVEAL_H, fontSize: 20, color: C.SUCCESS,
        });
      },
    ]);

    addFooter(s, FOOTER);
    s.addNotes(NOTES_CFU);
    runSlideDiagnostics(s, pres);
  })();

  // 8. I Do B - fill the plan, row by row
  (() => {
    const s = P.customSlide(pres, T, "I Do", C.PRIMARY, "Filling the plan: notes, not sentences");
    const y0 = CONTENT_TOP;
    const base = {
      x: 0.5, y: y0, w: 9, h: SAFE_BOTTOM - y0,
      rows: FROG_PLAN, labelW: 1.60, fontSize: 13,
    };

    clickBuild(s, FROG_PLAN.map((row, i) => () => {
      P.drawPlanRows(s, T, { ...base, parts: [i] });
    }));

    addFooter(s, FOOTER);
    s.addNotes(NOTES_PLAN);
    runSlideDiagnostics(s, pres);
  })();

  // 9. I Do C - notes become a paragraph
  (() => {
    const s = P.customSlide(pres, T, "I Do", C.PRIMARY, "From notes to a paragraph");
    const y0 = CONTENT_TOP;
    const h = SAFE_BOTTOM - y0;

    P.drawNoteCard(s, T, {
      x: 0.5, y: y0, w: 3.9, h,
      heading: "Aspect 1: Appearance",
      points: DRAFT_NOTES,
      color: C.SECONDARY,
      fontSize: 15,
    });

    const rx = 4.7;
    const rw = 4.8;
    const cardH = (h - 0.16) / 2;

    clickBuild(s, DRAFT_SENTENCES.map((item, i) => () => {
      const cy = y0 + i * (cardH + 0.16);
      s.addShape("roundRect", {
        x: rx, y: cy, w: rw, h: cardH, rectRadius: 0.08,
        fill: { color: C.WHITE }, line: { color: item.color, width: 1.5 },
      });
      addTextOnShape(s, item.label, {
        x: rx, y: cy, w: rw, h: 0.34, rectRadius: 0.08,
        fill: { color: item.color },
      }, {
        fontSize: 12.5, fontFace: FONT_B, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      s.addText(item.text, {
        x: rx + 0.18, y: cy + 0.40, w: rw - 0.36, h: cardH - 0.52,
        fontSize: 15, fontFace: FONT_B, color: C.CHARCOAL,
        valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
    }));

    addFooter(s, FOOTER);
    s.addNotes(NOTES_DRAFT);
    runSlideDiagnostics(s, pres);
  })();

  // 10. We Do - choose three aspects for a shared topic
  (() => {
    const s = P.customSlide(pres, T, "We Do", C.SUCCESS, "Choose the three aspects");
    const y0 = CONTENT_TOP;

    addTextOnShape(s, "Topic: tropical cyclones", {
      x: 0.5, y: y0, w: 9, h: 0.92, rectRadius: 0.08,
      fill: { color: C.SUCCESS },
    }, {
      fontSize: 30, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    const slotY = y0 + 1.06;
    const slotGap = 0.18;
    const slotW = (9 - slotGap * 2) / 3;
    ["Aspect 1", "Aspect 2", "Aspect 3"].forEach((label, i) => {
      const sx = 0.5 + i * (slotW + slotGap);
      s.addShape("roundRect", {
        x: sx, y: slotY, w: slotW, h: 1.30, rectRadius: 0.08,
        fill: { color: C.WHITE }, line: { color: C.SUCCESS, width: 1.5 },
      });
      s.addText(label, {
        x: sx, y: slotY + 0.12, w: slotW, h: 0.36,
        fontSize: 15, fontFace: FONT_B, color: C.SUCCESS, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      s.addText("?", {
        x: sx, y: slotY + 0.52, w: slotW, h: 0.66,
        fontSize: 34, fontFace: FONT_H, color: C.MUTED, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
    });

    addTextOnShape(s, "Write three aspects. Not three facts.", {
      x: 0.5, y: slotY + 1.44, w: 9, h: 0.56, rectRadius: 0.08,
      fill: { color: C.WHITE }, line: { color: C.SUCCESS, width: 1.5 },
    }, {
      fontSize: 19, fontFace: FONT_B, color: C.SUCCESS, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    clickBuild(s, [
      () => {
        addRevealAnswerBar(s, ["How they form", "Where they happen", "The damage they cause"], {
          y: P.REVEAL_Y, h: P.REVEAL_H, fontSize: 17, label: "One way", color: C.SUCCESS,
        });
      },
    ]);

    addFooter(s, FOOTER);
    s.addNotes(NOTES_WEDO);
    runSlideDiagnostics(s, pres);
  })();

  // 11. You Do - your topic, your plan
  (() => {
    const s = P.customSlide(pres, T, "You Do", C.ASSESS, "Your topic. Your plan.");
    const y0 = CONTENT_TOP;
    const h = SAFE_BOTTOM - y0;

    T.addInstructionCard(s, [
      { text: "Report Plan", role: "header" },
      { text: "1. Choose a topic from one of the four themes." },
      { text: "2. Name three aspects on your plan." },
      { text: "3. Research, then fill the rows with notes." },
      { text: "Twenty-five minutes. Notes only, no sentences." },
    ], {
      x: 0.5, y: y0, w: 4.7, h,
      strip: C.ASSESS, fill: C.WHITE,
      headerColor: C.ASSESS, emphasisColor: C.ALERT,
    });

    const rx = 5.5;
    const rw = 4.0;
    addCard(s, rx, y0, rw, 1.72, { strip: C.ACCENT });
    s.addText("Where to look", {
      x: rx + 0.22, y: y0 + 0.10, w: rw - 0.44, h: 0.30,
      fontSize: 14, fontFace: FONT_H, color: C.ACCENT, bold: true, margin: 0,
    });
    s.addText([
      { text: "Library books on the topic shelf.", options: { breakLine: true } },
      { text: "Sites ending in .gov.au or .edu.au.", options: { breakLine: true } },
      { text: "Check the date. Old facts go stale.", options: { breakLine: true } },
      { text: "Two sources, not one.", options: {} },
    ], {
      x: rx + 0.22, y: y0 + 0.44, w: rw - 0.44, h: 1.18,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL,
      valign: "top", margin: 0, paraSpaceAfter: 2,
    });

    const themeY = y0 + 1.88;
    addTextOnShape(s, "Choose from these themes", {
      x: rx, y: themeY, w: rw, h: 0.44, rectRadius: 0.07,
      fill: { color: C.PRIMARY },
    }, {
      fontSize: 14, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
    const tGap = 0.10;
    const tW = (rw - tGap) / 2;
    P.CLASS_THEMES.forEach((theme, i) => {
      const col = i % 2;
      const row = Math.floor(i / 2);
      addTextOnShape(s, theme, {
        x: rx + col * (tW + tGap), y: themeY + 0.56 + row * 0.62, w: tW, h: 0.54,
        rectRadius: 0.07,
        fill: { color: C.WHITE }, line: { color: C.ACCENT, width: 1.4 },
      }, {
        fontSize: 13, fontFace: FONT_B, color: C.ACCENT, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
    runSlideDiagnostics(s, pres);
  })();

  // 12. Exit ticket
  P.exitTicketPanel(pres, T, {
    topic: "Your own topic",
    task: "Write the opening sentence of YOUR report at the top of your plan.",
    cue: "Name it. Name its group. Add one true thing.",
    taskSize: 26,
    footer: FOOTER,
    notes: NOTES_EXIT,
  });

  // 13. Closing
  closingSlide(pres, {
    reflectionPrompt: "Which of your three aspects will you have the most to say about?",
    scItems: [
      "I can choose a topic I can find facts about.",
      "I can sort my facts into three aspects on a plan.",
      "I can draft the sentence that opens my report.",
    ],
    selfAssessment: "Fingers at your chest. One, two or three. Show me.",
    takeaways: [P.ANCHOR_PHRASE],
  }, NOTES_CLOSING);

  const pptxPath = path.join(OUT_DIR, "Information Reports W10 S1.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);
}

/* --- Resource: the Report Plan -------------------------------------------- */

async function buildPlan() {
  const doc = createPdf({ title: PLAN_RES.name });
  let y = addPdfHeader(doc, "Report Plan", {
    subtitle: "Notes only. No full sentences. You will use this all week.",
    color: C.PRIMARY,
    lessonInfo: "Year 5/6 Literacy | Information Reports | Week 10",
  });

  y = addTipBox(doc, P.ANCHOR_PHRASE, y, { color: C.ACCENT });

  // The plan IS the map, so the map is printed on the plan (section 79).
  y = P.addReportMapPdf(doc, y, {
    colors: {
      classify: hex(C.PRIMARY),
      describe: hex(C.SECONDARY),
      wrap: hex(C.ASSESS),
      facts: hex(C.ACCENT),
    },
    aspects: P.GENERIC_ASPECTS,
  });

  y = P.addNoteBoxPdf(doc, y, "Topic", { color: hex(C.PRIMARY), lines: 1, spacing: 26, dots: false });

  y = addSectionHeading(doc, "Does your topic pass both tests?", y, { color: C.ACCENT });
  y = P.addChecklistPdf(doc, y, [
    "I can find facts about it. Not just one story about one thing.",
    "I can split it into three aspects, each with something to say.",
  ], { color: hex(C.ACCENT) });

  y = P.addNoteBoxPdf(doc, y, "Introduction: what is it, and what group is it in?",
    { color: hex(C.PRIMARY), lines: 3, spacing: 25 });

  doc.addPage();
  y = PAGE.MARGIN;
  y = addSectionHeading(doc, "The middle: one aspect per box", y, { color: C.SECONDARY });
  y = addBodyText(doc,
    "Name the aspect on the line, then dot point the facts you found. Aim for three or four facts in each box.", y);

  [1, 2, 3].forEach((n) => {
    y = P.addNoteBoxPdf(doc, y, "Aspect " + n + ", which is:",
      { color: hex(C.SECONDARY), lines: 4, spacing: 25, headingLine: true });
  });

  y = P.addNoteBoxPdf(doc, y, "Conclusion: sum up, or one last interesting fact",
    { color: hex(C.ASSESS), lines: 2, spacing: 25 });

  y = P.addNoteBoxPdf(doc, y, "Where my facts came from (book titles or websites)",
    { color: hex(C.MUTED), lines: 2, spacing: 25, dots: false });

  addPdfFooter(doc, "Information Reports | Year 5/6 Literacy | Week 10");
  await writePdf(doc, path.join(OUT_DIR, PLAN_RES.fileName));
  console.log("Wrote " + PLAN_RES.name);
}

/* --- Resource: the Research Rescue Card ----------------------------------- */

// Enabling scaffold. It changes the FORM of the task rather than shortening it:
// the facts are supplied, so a student who cannot yet research still practises
// the thing the learning intention names (megaprompt section 73).
async function buildRescue() {
  const doc = createPdf({ title: RESCUE_RES.name });
  let y = addPdfHeader(doc, "Research Rescue Card", {
    subtitle: "For when the facts will not come. The facts are here. The writing is still yours.",
    color: C.SUCCESS,
    lessonInfo: "Year 5/6 Literacy | Information Reports | Week 10",
    showNameDate: false,
  });

  y = addSectionHeading(doc, "Step 1: the four questions that always work", y, { color: C.PRIMARY });
  P.RESEARCH_QUESTIONS.forEach((item) => {
    doc.fontSize(11).font("Sans").fillColor("#000000");
    doc.text(item.wondering, PAGE.MARGIN + 8, y, { width: 280, lineBreak: false });
    doc.font("Sans-Bold").fillColor(hex(C.SECONDARY));
    doc.text("becomes the " + item.aspect + " paragraph", PAGE.MARGIN + 290, y,
      { width: 200, lineBreak: false });
    y += 22;
  });
  y += 8;

  y = addSectionHeading(doc, "Step 2: what finished notes look like", y, { color: C.SECONDARY });
  y = addBodyText(doc,
    "This is the class topic, planned. Yours should look like this: dot points, never sentences.", y);
  y = P.addNoteBoxPdf(doc, y, "Aspect 1: Appearance", {
    color: hex(C.SECONDARY), lines: 4, spacing: 24,
    answers: ["smooth, bright green skin", "creamy white belly", "about ten centimetres long", "wide, round pads on its toes"],
  });

  doc.addPage();
  y = PAGE.MARGIN;

  y = addSectionHeading(doc, "Step 3: a topic with the facts already found", y, { color: C.SUCCESS });
  y = addBodyText(doc,
    "If your own topic is not working, take this one. The facts are here, so you can spend the time on the writing.", y);

  y = P.addSourceBoxPdf(doc, y, "Thunder: what it is",
    P.THUNDER.classify, { color: hex(C.PRIMARY) });
  P.THUNDER.aspects.forEach((aspect) => {
    y = P.addSourceBoxPdf(doc, y, "Aspect: " + aspect.heading, aspect.text,
      { color: hex(C.SECONDARY) });
  });
  y = P.addSourceBoxPdf(doc, y, "One last interesting fact", P.THUNDER.wrap,
    { color: hex(C.ASSESS) });

  y = addBodyText(doc,
    "Use your own words. Copying these boxes into your report is not writing a report.",
    y, { italic: true });

  addPdfFooter(doc, "Information Reports | Year 5/6 Literacy | Week 10");
  await writePdf(doc, path.join(OUT_DIR, RESCUE_RES.fileName));
  console.log("Wrote " + RESCUE_RES.name);
}

(async () => {
  await build();
  await buildPlan();
  await buildRescue();
})().catch((err) => {
  console.error(err);
  process.exit(1);
});
