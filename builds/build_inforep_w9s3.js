"use strict";

// Information Reports - Term 3 Week 9, Session 3 (Year 5/6 Literacy)
// "Ask the question. Take the notes. Use your own words."
//
// Victorian Curriculum 2.0: English, Literacy, Levels 5-6 - locating,
// selecting and synthesising information from more than one source, working
// towards VC2E5LY10 and VC2E6LY09.
//
// UNIT ANCHOR (locked): "Classify it. Describe it, one aspect at a time.
// Wrap it up. Facts all the way through."
//
// Lesson shape: short-cycle loops. Note-taking is a multi-step skill, so the
// modelling is three small model-then-try cycles (question, notes, own words)
// rather than one long demonstration.
//
// Sources: the teacher's own research questions, quoted from the unit plan.
// The two short sources are original, written for this unit so that two texts
// say overlapping things in different words (see inforep_lib.js).

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

const SESSION = 3;
const FOOTER = "Information Reports | Week 9 Session 3 | Year 5/6 Literacy";
const OUT_DIR = path.join(__dirname, "..", "output", "InfoReport_W9_S3");
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const SHEET_RES = makeSessionResource(SESSION,
  "Research and Notes Sheet",
  "Sort research questions, take dot point notes from two sources, write one sentence.");
const KEY_RES = makeSessionResource(SESSION,
  "Answer Key",
  "Model notes and a model sentence, with what to look for as you circulate.");
const RESOURCE_ITEMS = [SHEET_RES, KEY_RES];
fs.mkdirSync(RES_DIR, { recursive: true });

/* --- Teacher notes (Glance Format, megaprompt sections 45-47) ------------- */

const NOTES_TITLE =
  "Session 3. We know how a report is built. Today we go and find the facts that fill it.";

const NOTES_RESOURCES =
  "Prep slide. Print the Research and Notes Sheet, one per student. Keep the Answer Key for yourself.\n" +
  "Students need their map card and a pencil. If you have two real books on the class topic, put them on the desk and swap them in for the two printed sources.\n" +
  "CATCH-UP: missed a session? Everything today starts from a question anyone can ask. Hand over the map card and start at Section 1 of the sheet.";

const NOTES_LAUNCH = composeGlanceNotes({
  answer: "open - what they look like, where they live, what they eat.",
  beats: [
    [
      "SAY: Our topic is Australian frogs.",
      "You know nothing yet. What would you want to find out?",
    ],
    [
      "ASK: What would a reader want to know?",
      "45 sec. Turn and tell, partner A first.",
      "EXPECT: look like, live, eat, survive.",
    ],
    [
      "SCAN the room for questions, not facts.",
      "80%+ -> reveal the four, then move on.",
      "Less -> model one out loud, re-ask.",
    ],
    [
      "REVEAL after the partner talk.",
      "SAY: Every one of those is an aspect of our map.",
    ],
  ],
  trap: [
    "offering facts instead of questions.",
    "Fix: ask for a question mark, student rephrases it.",
  ],
  prep: "Low-coupling launch: any student can wonder about frogs, so nothing earlier is assumed. The teacher's own four questions. Whole block under 6 minutes.",
  sources: P.SOURCE_LINE_UNIT,
  tag: "[Launch | Retention and recall | HITS 6]",
});

const NOTES_LI = composeGlanceNotes({
  beats: [
    [
      "POINT to the learning intention.",
      "SAY: Today we ask good questions and take notes in our own words.",
    ],
    "SAY: Read the I can statements with me. Everyone, together, on three.",
    "SAY: The first one everyone leaves with. Turn a wondering into a question.",
  ],
  prep: "SC1 is reachable by every student; SC2 is the core target the exit ticket assesses; SC3 stretches into joining two sources. Tier labels stay off the slide.",
  tag: "[LI/SC | Planning made visible | HITS 1]",
});

const NOTES_VOCAB = composeGlanceNotes({
  answer: "paraphrase means to say the same idea in your own words.",
  beats: [
    [
      "SAY: To paraphrase is to say the same fact a different way.",
      "Same meaning. Your words.",
    ],
    [
      "ASK: Say this in your own words. Frogs are nocturnal.",
      "20 sec. Turn and tell, partner B first.",
      "EXPECT: they come out at night.",
    ],
    "SAY: Copying is easy and useless. Paraphrasing proves you understood.",
  ],
  trap: [
    "swapping one word and calling it a paraphrase.",
    "Fix: cover the sentence, student says it from memory.",
  ],
  prep: "The word that separates note-taking from copying. It is the reason Section 3 of the sheet exists.",
  tag: "[Key Vocabulary | Knowledge and memory | HITS 6]",
});

const NOTES_QUESTIONS = composeGlanceNotes({
  answer: "each question becomes one aspect of the report.",
  beats: [
    [
      "SAY: A wondering is a thought. A question is a job.",
      "Watch me turn one into the other.",
    ],
    [
      "POINT to each question and its aspect.",
      "SAY: What does it eat becomes the diet paragraph.",
    ],
    [
      "ASK: Which aspect does How does it hide belong to?",
      "20 sec. Boards up on cue.",
      "EXPECT: appearance.",
    ],
    [
      "SCAN boards.",
      "80%+ -> cold call: how do you know?",
      "Less -> read it slowly again, re-ask.",
    ],
  ],
  trap: [
    "writing a yes or no question.",
    "Fix: ask what a paragraph would say, student rewrites.",
  ],
  stretch: "Write a fifth question and name its aspect.",
  help: "Give the four aspects. Match, do not recall.",
  prep: "Say the anchor once here. The four questions are the teacher's own, from the unit plan.",
  sources: P.SOURCE_LINE_UNIT,
  tag: "[I Do | Explicit teaching | SC1 | HITS 3, 4]",
});

const NOTES_NOTETAKING = composeGlanceNotes({
  answer: "amphibian, warm and wet places, northern Australia.",
  beats: [
    [
      "SAY: Watch me take notes. Three steps, every time.",
      "Read it, pick key words, write dots.",
    ],
    [
      "CLICK to the key words. SAY: Not the whole sentence.",
      "CLICK to the notes. Three dots, no sentences.",
    ],
    [
      "ASK: Why not write the whole sentence?",
      "20 sec. Turn and tell, partner A first.",
      "EXPECT: to use our own words later.",
    ],
  ],
  trap: [
    "copying the sentence and calling it notes.",
    "Fix: cover the source, student writes from dots.",
  ],
  stretch: "Take dot points from the next sentence alone.",
  help: "Give two of the three dots. Ask for the third.",
  prep: "The three steps here are the same three steps on the sheet, so the paper and the screen match exactly.",
  sources: P.SOURCE_LINE_MODEL,
  tag: "[I Do | Explicit teaching | SC2 | HITS 4]",
});

const NOTES_CFU = composeGlanceNotes({
  answer: "B. Key words only, and it is not copied.",
  beats: [
    [
      "SAY: Same sentence. Three sets of notes.",
      "Do not call out. This one goes on boards.",
    ],
    [
      "ASK: Which notes are the most useful?",
      "30 sec. Boards up on cue.",
      "EXPECT: B.",
    ],
    [
      "SCAN boards, back row first.",
      "80%+ -> cold call one B: what is wrong with A?",
      "Less -> read A aloud, ask whose words those are, re-ask.",
    ],
    [
      "REVEAL after boards are scanned.",
      "SAY: A is copying. C is not enough to write from.",
    ],
  ],
  trap: [
    "choosing A because it holds the most information.",
    "Fix: ask what happens when they copy it in, student re-tests.",
  ],
  prep: "The hinge of the lesson. A is the plagiarism risk the unit plan names; C is the too-thin note that cannot be written from.",
  sources: P.SOURCE_LINE_UNIT,
  tag: "[CFU hinge | Supported application | SC2 | HITS 7, 8]",
});

const NOTES_WEDO = composeGlanceNotes({
  answer: "open - one sentence covering night feeding and insects.",
  beats: [
    [
      "SAY: Two sources, same aspect. Read both with me.",
      "Some facts appear twice. Some appear once.",
    ],
    [
      "ASK: Write ONE sentence that uses both sources.",
      "90 sec. Write it... chin it... show me.",
      "EXPECT: night plus insects, in their own words.",
    ],
    [
      "SCAN boards for facts from BOTH sources.",
      "80%+ -> cold call one: which source gave you that?",
      "Less -> ring one fact in each source, re-ask.",
    ],
  ],
  trap: [
    "copying one source and ignoring the other.",
    "Fix: ask what source two added, student adds it in.",
  ],
  stretch: "Add a third fact only one source mentions.",
  help: "Give the first four words of the sentence.",
  prep: "The synthesis move the unit plan asks for. Different aspect from the You Do, so the sheet is application, not repetition.",
  sources: P.SOURCE_LINE_MODEL,
  tag: "[We Do | Supported application | SC3 | HITS 4, 5]",
});

const NOTES_YOUDO = composeGlanceNotes({
  answer: "notes from both sources, then one sentence about the frog call.",
  beats: [
    [
      "SAY: Your turn, on the sheet. A new aspect, the frog's call.",
      "Same three steps as the model.",
    ],
    [
      "POINT to the steps. TIME: fifteen minutes.",
      "Section 1 first, it warms up the questions.",
    ],
    [
      "CIRCULATE. Look for dot points, not sentences.",
      "COLLECT two sentences to read at the closing.",
    ],
  ],
  trap: [
    "writing the sentence before taking the notes.",
    "Fix: cover the sources, student writes from the dots only.",
  ],
  stretch: "Section 4: write a question these sources do not answer.",
  help: "One dot point is filled in on each source box.",
  prep: "Different aspect and different sources from the We Do. Section 1 is doable from today's launch alone, so a returning student can start.",
  sources: P.SOURCE_LINE_MODEL,
  tag: "[You Do | Mastery and application | SC2 | HITS 10]",
});

const NOTES_EXIT = composeGlanceNotes({
  answer: "any sentence meaning green tree frogs come out at night.",
  beats: [
    [
      "SAY: One sentence. Say this fact in your own words.",
      "Do not use the word nocturnal.",
    ],
    [
      "TIME: three minutes. Books open on the desk when you finish.",
      "COLLECT by walking the rows, not by calling names.",
    ],
    [
      "SORT into two piles as you collect: own words, copied.",
      "A thick second pile -> re-model paraphrasing in Session 4.",
    ],
  ],
  prep: "Assesses the core target: put a fact into your own words. Banning the word nocturnal is what makes it a real paraphrase.",
  tag: "[Exit Ticket | Mastery and application | SC2 | HITS 8]",
});

const NOTES_CLOSING = composeGlanceNotes({
  answer: "open - most students should show two.",
  beats: [
    "READ the two collected sentences aloud. SAY: Listen for two sources in one.",
    [
      "ASK: Rate yourself on the three statements.",
      "15 sec. Fingers at your chest, one, two or three. Show me.",
      "EXPECT: most on two.",
    ],
    "SAY: Next session we work on the words that make it sound like a report.",
  ],
  prep: "Self-assessment feeds Session 4 grouping. Anyone showing one gets the model notes in front of them next session.",
  tag: "[Closing | Retention and recall | HITS 1, 9]",
});

/* --- Build ---------------------------------------------------------------- */

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // 1. Title
  titleSlide(pres, "Finding the Facts",
    "Ask the question. Take the notes. Use your own words.",
    "Week 9 Session 3 | Year 5/6 Literacy", NOTES_TITLE);

  // 2. Teacher Resources
  addResourceSlide(pres, {
    resources: RESOURCE_ITEMS,
    studentTools: [
      "Mini-whiteboards and markers",
      "Writing books and pencils",
      "The map card",
    ],
    boardSetup: [
      "Two real books on the class topic, if you have them",
      "Leave room to model dot points live",
    ],
  }, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // 3. Launch - what would you want to know?
  (() => {
    const s = P.customSlide(pres, T, "Launch", C.SECONDARY, "What would you want to know?");
    const y0 = CONTENT_TOP;

    addTextOnShape(s, "Topic: Australian frogs", {
      x: 0.5, y: y0, w: 9, h: 0.92, rectRadius: 0.08,
      fill: { color: C.SECONDARY },
    }, {
      fontSize: 30, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    const gridY = y0 + 1.06;
    const chipW = 4.4;
    const chipH = 0.80;
    const gapX = 0.20;
    const gapY = 0.12;

    clickBuild(s, P.RESEARCH_QUESTIONS.map((item, i) => () => {
      const col = i % 2;
      const row = Math.floor(i / 2);
      const cx = 0.5 + col * (chipW + gapX);
      const cy = gridY + row * (chipH + gapY);
      s.addShape("roundRect", {
        x: cx, y: cy, w: chipW, h: chipH, rectRadius: 0.08,
        fill: { color: C.WHITE }, line: { color: C.PRIMARY, width: 1.5 },
      });
      s.addText(item.wondering, {
        x: cx + 0.16, y: cy, w: chipW - 1.42, h: chipH,
        fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL,
        valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
      s.addShape("roundRect", {
        x: cx + chipW - 1.20, y: cy + 0.16, w: 1.04, h: 0.48, rectRadius: 0.07,
        fill: { color: C.PRIMARY },
      });
      s.addText(item.aspect, {
        x: cx + chipW - 1.20, y: cy + 0.16, w: 1.04, h: 0.48,
        fontSize: 11, fontFace: FONT_B, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
    }));

    addFooter(s, FOOTER);
    s.addNotes(NOTES_LAUNCH);
    runSlideDiagnostics(s, pres);
  })();

  // 4. Learning Intention and Success Criteria
  liSlide(pres,
    "We are learning to ask research questions and take notes in our own words.",
    [
      "I can turn a wondering into a research question.",
      "I can take dot point notes that answer my question.",
      "I can join facts from two sources into one new sentence of my own.",
    ],
    NOTES_LI, FOOTER);

  // 5. Key vocabulary
  keyWordSlide(pres, {
    word: "paraphrase",
    meaning: "To say the same idea in your own words.",
    example: "Frogs are nocturnal becomes Frogs come out at night.",
    routine: ["Say it", "Cover it", "Say it again"],
    color: C.ASSESS,
    title: "The word for today",
  }, NOTES_VOCAB, FOOTER);

  // 6. I Do A - a wondering becomes a research question, and an aspect
  (() => {
    const s = P.customSlide(pres, T, "I Do", C.PRIMARY, "A question becomes an aspect");
    const y0 = CONTENT_TOP;
    const rowH = 0.68;
    const gap = 0.14;

    P.RESEARCH_QUESTIONS.forEach((item, i) => {
      const ry = y0 + i * (rowH + gap);
      s.addShape("roundRect", {
        x: 0.5, y: ry, w: 5.6, h: rowH, rectRadius: 0.08,
        fill: { color: C.WHITE }, line: { color: C.SECONDARY, width: 1.4 },
      });
      s.addText(item.wondering, {
        x: 0.76, y: ry, w: 5.1, h: rowH,
        fontSize: 17, fontFace: FONT_B, color: C.CHARCOAL,
        valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
      s.addText("becomes", {
        x: 6.20, y: ry, w: 0.90, h: rowH,
        fontSize: 12, fontFace: FONT_B, color: C.MUTED, italic: true,
        align: "center", valign: "middle", margin: 0,
      });
      addTextOnShape(s, item.aspect, {
        x: 7.20, y: ry, w: 2.3, h: rowH, rectRadius: 0.08,
        fill: { color: C.PRIMARY },
      }, {
        fontSize: 15, fontFace: FONT_B, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
    });

    const noteY = y0 + 4 * (rowH + gap);
    s.addText("One question. One aspect. One paragraph.", {
      x: 0.5, y: noteY, w: 9, h: 0.44,
      fontSize: 17, fontFace: FONT_H, color: C.PRIMARY, bold: true, italic: true,
      align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_QUESTIONS);
    runSlideDiagnostics(s, pres);
  })();

  // 7. I Do B - three steps to a dot point note
  (() => {
    const s = P.customSlide(pres, T, "I Do", C.PRIMARY, "Read it. Pick the key words. Write dots.");
    const y0 = CONTENT_TOP;

    P.drawReportExtract(s, T, {
      x: 0.5, y: y0, w: 9, h: 1.00,
      bodySize: 16,
      showLabels: true,
      labelW: 1.35,
      sections: [{ label: "Source", text: P.NOTE_SOURCE_SENTENCE, color: C.SECONDARY }],
    });

    const chipY = y0 + 1.14;
    const chipGap = 0.16;
    const chipW = (9 - chipGap * 2) / 3;

    clickBuild(s, [
      () => {
        P.NOTE_KEY_WORDS.forEach((word, i) => {
          addTextOnShape(s, word, {
            x: 0.5 + i * (chipW + chipGap), y: chipY, w: chipW, h: 0.62, rectRadius: 0.08,
            fill: { color: C.ACCENT },
          }, {
            fontSize: 17, fontFace: FONT_B, color: C.WHITE, bold: true,
            align: "center", valign: "middle", margin: 0,
          });
        });
      },
      () => {
        P.drawNoteCard(s, T, {
          x: 0.5, y: chipY + 0.76, w: 9, h: 1.86,
          heading: "My notes",
          points: P.NOTE_MODEL_POINTS,
          color: C.SUCCESS,
          fontSize: 17,
        });
      },
    ]);

    addFooter(s, FOOTER);
    s.addNotes(NOTES_NOTETAKING);
    runSlideDiagnostics(s, pres);
  })();

  // 8. CFU hinge - which notes are the most useful?
  (() => {
    const s = P.customSlide(pres, T, "CFU", C.ALERT, "Which notes are the most useful?");
    const y0 = CONTENT_TOP;

    addTextOnShape(s, P.NOTE_SOURCE_SENTENCE, {
      x: 0.5, y: y0, w: 9, h: 0.62, rectRadius: 0.08,
      fill: { color: C.ALERT },
    }, {
      fontSize: 15, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0.06,
    });

    // Options must stop above the reveal bar: 2.02 + 3 x 0.62 + 2 x 0.10 = 4.08.
    P.drawOptionStack(s, T, P.NOTE_OPTIONS, {
      y: y0 + 0.72, optionH: 0.62, gap: 0.10, color: C.PRIMARY, fontSize: 15,
    });

    clickBuild(s, [
      () => {
        addRevealAnswerBar(s, ["B - key words only, and not copied"], {
          y: P.REVEAL_Y, h: P.REVEAL_H, fontSize: 21, color: C.SUCCESS,
        });
      },
    ]);

    addFooter(s, FOOTER);
    s.addNotes(NOTES_CFU);
    runSlideDiagnostics(s, pres);
  })();

  // 9. We Do - two sources, one new sentence
  (() => {
    const s = P.customSlide(pres, T, "We Do", C.SUCCESS, "Two sources. One new sentence.");
    const y0 = CONTENT_TOP;
    const colW = 4.35;

    P.DIET_SOURCES.forEach((src, i) => {
      P.drawReportExtract(s, T, {
        x: 0.5 + i * (colW + 0.30), y: y0, w: colW, h: 2.10,
        bodySize: 13,
        showLabels: true,
        labelW: 1.15,
        sections: [{ label: src.label, text: src.text, color: i === 0 ? C.PRIMARY : C.SECONDARY }],
      });
    });

    addTextOnShape(s, "Write ONE sentence that uses facts from both.", {
      x: 0.5, y: y0 + 2.22, w: 9, h: 0.58, rectRadius: 0.08,
      fill: { color: C.SUCCESS },
    }, {
      fontSize: 19, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    clickBuild(s, [
      () => {
        addRevealAnswerBar(s, [P.DIET_PARAPHRASE], {
          y: P.REVEAL_Y, h: P.REVEAL_H, fontSize: 15, label: "One way", color: C.SUCCESS,
        });
      },
    ]);

    addFooter(s, FOOTER);
    s.addNotes(NOTES_WEDO);
    runSlideDiagnostics(s, pres);
  })();

  // 10. You Do - the sheet, on a new aspect
  (() => {
    const s = P.customSlide(pres, T, "You Do", C.ASSESS, "Your turn: the frog's call");
    const y0 = CONTENT_TOP;
    const h = SAFE_BOTTOM - y0;

    T.addInstructionCard(s, [
      { text: "Research and Notes Sheet", role: "header" },
      { text: "Two new sources about the sound a frog makes." },
      { text: "1. Sort the questions into aspects." },
      { text: "2. Dot point notes from each source." },
      { text: "3. One sentence, your own words." },
      { text: "Fifteen minutes. Notes before sentence." },
    ], {
      x: 0.5, y: y0, w: 4.6, h,
      strip: C.ASSESS, fill: C.WHITE,
      headerColor: C.ASSESS, emphasisColor: C.ALERT,
    });

    const steps = [
      { n: "1", label: "Read both sources", color: C.SECONDARY },
      { n: "2", label: "Dot points from each", color: C.SUCCESS },
      { n: "3", label: "One sentence, your words", color: C.ASSESS },
    ];
    const stepH = (h - 0.24) / 3;
    steps.forEach((step, i) => {
      const sy = y0 + i * (stepH + 0.12);
      s.addShape("roundRect", {
        x: 5.3, y: sy, w: 4.2, h: stepH, rectRadius: 0.08,
        fill: { color: step.color },
      });
      s.addShape("roundRect", {
        x: 5.50, y: sy + (stepH - 0.50) / 2, w: 0.50, h: 0.50, rectRadius: 0.25,
        fill: { color: C.WHITE },
      });
      s.addText(step.n, {
        x: 5.50, y: sy + (stepH - 0.50) / 2, w: 0.50, h: 0.50,
        fontSize: 19, fontFace: FONT_H, color: step.color, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      s.addText(step.label, {
        x: 6.14, y: sy, w: 3.20, h: stepH,
        fontSize: 17, fontFace: FONT_B, color: C.WHITE, bold: true,
        valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
    runSlideDiagnostics(s, pres);
  })();

  // 11. Exit ticket
  P.exitTicketPanel(pres, T, {
    topic: "Green tree frogs are nocturnal, which means they are active at night.",
    task: "Write that fact in your own words.",
    cue: "You may not use the word nocturnal.",
    taskSize: 30,
    footer: FOOTER,
    notes: NOTES_EXIT,
  });

  // 12. Closing
  closingSlide(pres, {
    reflectionPrompt: "Which is harder for you: finding the facts, or saying them your own way?",
    scItems: [
      "I can turn a wondering into a research question.",
      "I can take dot point notes that answer my question.",
      "I can join facts from two sources into one new sentence of my own.",
    ],
    selfAssessment: "Fingers at your chest. One, two or three. Show me.",
    takeaways: [P.ANCHOR_PHRASE],
  }, NOTES_CLOSING);

  const pptxPath = path.join(OUT_DIR, "Information Reports W9 S3.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);
}

/* --- Resources ------------------------------------------------------------ */

function sheetBody(doc, isKey) {
  let y = addPdfHeader(doc, isKey ? "Research and Notes: Answer Key" : "Research and Notes Sheet", {
    subtitle: isKey
      ? "Model notes and a model sentence, with what to look for."
      : "Sort the questions. Take the notes. Write one sentence of your own.",
    color: isKey ? C.ALERT : C.PRIMARY,
    lessonInfo: "Year 5/6 Literacy | Information Reports | Session 3",
    showNameDate: !isKey,
  });

  y = addTipBox(doc, "Read it. Pick the key words. Write dot points. Then use your own words.", y,
    { color: C.ACCENT });

  y = addSectionHeading(doc, "Section 1: which aspect does each question belong to?", y,
    { color: isKey ? C.ALERT : C.PRIMARY });
  y = addBodyText(doc,
    "Write the aspect beside each question. Choose from: Appearance, Habitat, Diet, Adaptations.", y);

  P.RESEARCH_QUESTIONS.forEach((item) => {
    doc.fontSize(11).font("Sans").fillColor("#000000");
    doc.text(item.wondering, PAGE.MARGIN + 8, y, { width: 280, lineBreak: false });
    if (isKey) {
      doc.font("Sans-Bold").fillColor(hex(C.SUCCESS));
      doc.text(item.aspect, PAGE.MARGIN + 300, y, { width: 160, lineBreak: false });
    } else {
      doc.moveTo(PAGE.MARGIN + 300, y + 13).lineTo(PAGE.MARGIN + 460, y + 13)
        .lineWidth(0.9).strokeColor("#000000").stroke();
    }
    y += 26;
  });
  y += 8;

  y = addSectionHeading(doc, "Section 2: notes from two sources", y,
    { color: isKey ? C.ALERT : C.SECONDARY });
  y = addBodyText(doc,
    "Both sources are about the same aspect: the sound a green tree frog makes.", y);

  P.CALL_SOURCES.forEach((src, i) => {
    // A source and its notes box are one unit: break the page before the pair
    // rather than letting the notes box land on the next page on its own.
    if (y + 210 > PAGE.H - PAGE.MARGIN - 30) {
      doc.addPage();
      y = PAGE.MARGIN;
    }
    y = P.addSourceBoxPdf(doc, y, src.label, src.text, {
      color: hex(i === 0 ? C.PRIMARY : C.SECONDARY),
    });
    y = P.addNoteBoxPdf(doc, y, "My notes from " + src.label, {
      color: hex(C.SUCCESS),
      lines: 3,
      spacing: 26,
      // The enabler is on the sheet itself: the first dot point is filled in,
      // so a student who cannot start still has a model in front of them.
      answers: isKey ? src.notes : [src.notes[0]],
    });
  });

  // No forced page break here: Sections 3 and 4 fit under the second source
  // pair, which keeps the sheet to two pages.
  y = addSectionHeading(doc, "Section 3: one sentence, in your own words", y,
    { color: isKey ? C.ALERT : C.ASSESS });
  y = addBodyText(doc,
    "Cover both sources. Using only your dot points, write ONE sentence that uses facts from both.", y);
  if (isKey) {
    y = addBodyText(doc, "One way: " + P.CALL_PARAPHRASE, y, { italic: true, color: "2D6B4A" });
    y = addBodyText(doc,
      "Look for: a fact from EACH source, and wording that is not lifted from either. A sentence that repeats one source word for word has not been paraphrased.",
      y, { italic: true, color: "2D6B4A" });
    y += 8;
  } else {
    y = addLinedArea(doc, y + 6, 5, { lineSpacing: 30 });
    y += 16;
  }

  y = addSectionHeading(doc, "Section 4: the question nobody answered", y,
    { color: isKey ? C.ALERT : C.SUCCESS });
  y = addBodyText(doc,
    "Write one research question about green tree frogs that neither source answers. Then name the aspect it belongs to.", y);
  if (isKey) {
    y = addBodyText(doc,
      "Look for: a real question with a question mark, not a yes or no question, plus a named aspect. Anything about predators, life cycle or size is a strong answer.",
      y, { italic: true, color: "2D6B4A" });
  } else {
    y = addLinedArea(doc, y + 6, 5, { lineSpacing: 30 });
  }

  addPdfFooter(doc, "Information Reports | Year 5/6 Literacy | Session 3");
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
