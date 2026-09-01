"use strict";

// Information Reports - Term 3 Week 10, Session 2 (Year 5/6 Literacy)
// "Notes into paragraphs, then finish strong"
//
// Victorian Curriculum 2.0: English, Literacy, Levels 5-6 - VC2E5LY10 and
// VC2E6LY09, creating a text with elaborated, sequenced ideas; VC2E6LA06 for
// verb choice and noun groups inside the drafted paragraph.
//
// UNIT ANCHOR (locked): "Classify it. Describe it, one aspect at a time.
// Wrap it up. Facts all the way through." Today is Describe and Wrap up.
//
// Lesson shape: example-first, then sustained writing. This is a drafting
// session, so the deck is deliberately short: two teaching moves and a long
// You Do, with the checklist left on screen while students write.
//
// Sources: teacher's unit plan. Model text is original (see inforep_lib.js).

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

const SESSION = 2;
const FOOTER = "Information Reports | Week 10 Session 2 | Year 5/6 Literacy";
const OUT_DIR = path.join(__dirname, "..", "output", "InfoReport_W10_S2");
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const CHECK_RES = makeSessionResource(SESSION,
  "Drafting Checklist",
  "One page to keep beside the draft: paragraph shape, the four tools, the two endings.");
const RESOURCE_ITEMS = [CHECK_RES];
fs.mkdirSync(RES_DIR, { recursive: true });

/* --- Teacher notes (Glance Format, megaprompt sections 45-47) ------------- */

const NOTES_TITLE =
  "Session 2. Plans are done. Today every aspect row becomes a paragraph, and the report gets an ending.";

const NOTES_RESOURCES =
  "Prep slide. Print the Drafting Checklist, one per student, and hand back their plans at the door.\n" +
  "Leave the You Do slide up while they write. Whiteboards out for the We Do only.\n" +
  "CATCH-UP: no plan yet? Hand over a Research Rescue Card from Session 1 and let them plan one aspect, then draft it. They join the class at the same paragraph.";

const NOTES_LAUNCH = composeGlanceNotes({
  answer: "the middle. One aspect per paragraph, with a topic sentence.",
  beats: [
    [
      "SAY: Plans out. Say the map with me.",
      "Everyone, together, on three.",
    ],
    [
      "ASK: Which part are you working on today?",
      "15 sec. Fingers at your chest: one, two or three.",
      "EXPECT: mostly two.",
    ],
    [
      "SCAN the fingers. Note anyone showing one.",
      "80%+ showing two or three -> move on.",
      "Less -> re-read the map bands, point at the plan rows, re-ask.",
    ],
  ],
  trap: [
    "starting the middle without a topic sentence.",
    "Fix: point at the Describe band, student says its job.",
  ],
  prep: "Low-coupling launch: the map is on screen, so a student with no plan can still locate themselves. Note the ones on one and sit near them. Under 5 minutes.",
  tag: "[Launch | Retention and recall | HITS 6]",
});

const NOTES_LI = composeGlanceNotes({
  beats: [
    [
      "POINT to the learning intention.",
      "SAY: Today your notes become paragraphs, and your report gets an ending.",
    ],
    "SAY: Read the I can statements with me. Everyone, together, on three.",
    "SAY: The first one everyone leaves with. Start with a topic sentence.",
  ],
  prep: "SC1 is reachable by every student; SC2 is the core target the exit ticket assesses; SC3 stretches into the ending. Tier labels stay off the slide.",
  tag: "[LI/SC | Planning made visible | HITS 1]",
});

const NOTES_IDO = composeGlanceNotes({
  answer: "topic sentence first, then details with noun groups.",
  beats: [
    [
      "SAY: Four dot points. Watch them become a paragraph.",
      "I cover the plan, so the words are mine.",
    ],
    [
      "CLICK to the topic sentence. It names the aspect.",
      "CLICK again. Two details, in my own words.",
      "CLICK again. Noun groups: still water, cool damp places.",
    ],
    [
      "ASK: Which sentence tells the reader the aspect?",
      "15 sec. Write it... chin it... show me.",
      "EXPECT: the first one.",
    ],
  ],
  trap: [
    "listing the dot points as short separate sentences.",
    "Fix: join two dots aloud, student writes it.",
  ],
  stretch: "Add a technical word to the paragraph.",
  help: "Give the topic sentence. Student writes one detail.",
  prep: "Say the anchor once here, on the Describe band. Think aloud about the word choices, not just the order.",
  sources: P.SOURCE_LINE_MODEL,
  tag: "[I Do | Explicit teaching | SC1 | HITS 4, 9]",
});

const NOTES_CFU = composeGlanceNotes({
  answer: "B. Every sentence is about where they live.",
  beats: [
    [
      "SAY: Three paragraphs. Only one stays on a single aspect.",
      "Do not call out. This one goes on boards.",
    ],
    [
      "ASK: Which paragraph stays on one aspect?",
      "30 sec. Boards up on cue.",
      "EXPECT: B.",
    ],
    [
      "SCAN boards, back row first.",
      "80%+ -> cold call one B: how many aspects are in A?",
      "Less -> read A sentence by sentence, name each aspect, re-ask.",
    ],
    [
      "REVEAL after boards are scanned.",
      "SAY: A has three aspects in one paragraph. C has opinions.",
    ],
  ],
  trap: [
    "choosing A because it holds the most facts.",
    "Fix: ask for one heading that fits all of A, student re-tests.",
  ],
  prep: "The hinge of the lesson, and the most common fault in drafts: a paragraph that drifts across aspects. C is the opinion drift.",
  tag: "[CFU hinge | Supported application | SC2 | HITS 7, 8]",
});

const NOTES_WEDO = composeGlanceNotes({
  answer: "open - a summary in new words, or one last fact.",
  beats: [
    [
      "SAY: An ending does one of two jobs.",
      "Read my ending for the frog.",
    ],
    [
      "ASK: Write an ending for tropical cyclones.",
      "90 sec. Boards up on cue.",
      "EXPECT: a summary, or one striking fact.",
    ],
    [
      "SCAN boards for one of the two jobs.",
      "80%+ -> cold call: which job did you pick?",
      "Less -> write one together, name its job, re-ask.",
    ],
  ],
  trap: [
    "ending with an opinion, such as cyclones are scary.",
    "Fix: ask could we check it, student swaps in a fact.",
  ],
  stretch: "Write the other kind of ending too.",
  help: "Give the first four words of a summary ending.",
  prep: "The shared topic carries over from Session 1, so the thinking is about the ending, not the research. Their own ending comes in the You Do.",
  tag: "[We Do | Supported application | SC3 | HITS 4, 5]",
});

const NOTES_YOUDO = composeGlanceNotes({
  answer: "three aspect paragraphs and an ending, from their own plan.",
  beats: [
    [
      "SAY: Your plan, your paragraphs. One aspect per paragraph.",
      "Checklist stays beside you the whole time.",
    ],
    [
      "POINT to the steps. TIME: thirty minutes of writing.",
      "This slide stays up. Do not put your hand up to start.",
    ],
    [
      "CIRCULATE. Read the first sentence of each paragraph only.",
      "COLLECT two endings to read at the closing.",
    ],
  ],
  trap: [
    "writing one long paragraph for all three aspects.",
    "Fix: ring where the aspect changes, student starts a new paragraph.",
  ],
  stretch: "Add a technical word to every paragraph.",
  help: "Draft one aspect only, from a plan row you have already filled.",
  prep: "Their own topic, unlike the We Do. Protect the writing time: sit and confer, do not stop the class. Time budget: thirty minutes.",
  tag: "[You Do | Mastery and application | SC2 | HITS 10]",
});

const NOTES_EXIT = composeGlanceNotes({
  answer: "one word, summary or fact, plus their own last sentence.",
  beats: [
    [
      "SAY: Look at your last sentence. Which job does it do?",
      "Write the word, then copy the sentence.",
    ],
    [
      "TIME: three minutes. Drafts open on the desk when you finish.",
      "COLLECT by walking the rows, not by calling names.",
    ],
    [
      "SORT into two piles as you collect: job named, no ending yet.",
      "A thick second pile -> give five minutes to finish at the start of Session 3.",
    ],
  ],
  prep: "Assesses the core target and tells you who has not finished drafting, which decides how Session 3 opens.",
  tag: "[Exit Ticket | Mastery and application | SC2 | HITS 8]",
});

const NOTES_CLOSING = composeGlanceNotes({
  answer: "open - most students should show two.",
  beats: [
    "READ the two collected endings aloud. SAY: Name the job as I read.",
    [
      "ASK: Rate yourself on the three statements.",
      "15 sec. Fingers at your chest, one, two or three. Show me.",
      "EXPECT: most on two.",
    ],
    "SAY: Next session we add a picture that teaches, then we edit.",
  ],
  prep: "Anyone showing one gets five minutes of writing time at the start of Session 3 while others begin the visual.",
  tag: "[Closing | Retention and recall | HITS 1, 9]",
});

/* --- Slide content -------------------------------------------------------- */

const LAUNCH_STATUS = [
  { n: "1", label: "Still planning", color: C.ALERT },
  { n: "2", label: "Opening written", color: C.SECONDARY },
  { n: "3", label: "First aspect done", color: C.SUCCESS },
];

const HABITAT_NOTES = [
  "trees, reeds, rocks",
  "near still water",
  "downpipes, water tanks",
  "cool and damp",
];

const HABITAT_DRAFT = [
  { label: "Topic sentence: it names the aspect", text: P.FROG.aspects[1].topicSentence, color: C.PRIMARY },
  { label: "Details, in my own words", text: "They also live in cool, damp places around houses, such as downpipes, water tanks and letterboxes.", color: C.SECONDARY },
  { label: "Noun groups doing the work", text: "still water / cool, damp places around houses", color: C.ACCENT },
];

const CFU_OPTIONS = [
  { key: "A", text: "Green tree frogs shelter near still water. They also eat moths and crickets. Their skin is bright green." },
  { key: "B", text: "Green tree frogs shelter in trees and reeds near still water. They also hide in downpipes and water tanks. Both places stay cool and damp." },
  { key: "C", text: "Green tree frogs are amazing. They live in lots of places. Everyone should see one." },
];

const ENDING_JOBS = [
  { label: "Sum it up", detail: "Say your main idea again, in new words." },
  { label: "One last fact", detail: "Finish on the most interesting thing you found." },
];

/* --- Build ---------------------------------------------------------------- */

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // 1. Title
  titleSlide(pres, "Drafting the Middle and the End",
    "One aspect per paragraph. Then finish strong.",
    "Week 10 Session 2 | Year 5/6 Literacy", NOTES_TITLE);

  // 2. Teacher Resources
  addResourceSlide(pres, {
    resources: RESOURCE_ITEMS,
    studentTools: [
      "Their plan from Session 1, handed back at the door",
      "The Drafting Checklist, one each",
      "Writing books or devices",
      "Mini-whiteboards for the We Do only",
    ],
    boardSetup: [
      "Leave the You Do slide up through the writing block",
      "A chair beside you for conferring, not a queue",
    ],
  }, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // 3. Launch - say the map, then locate yourself
  (() => {
    const s = P.customSlide(pres, T, "Launch", C.SECONDARY, "Where are you up to?");
    const y0 = CONTENT_TOP;
    const h = SAFE_BOTTOM - y0;

    P.drawReportMap(s, T, {
      x: 0.5, y: y0, w: 4.3, h,
      focus: "describe",
      aspects: P.GENERIC_ASPECTS,
    });

    const rx = 5.2;
    const rw = 4.3;
    const cardH = 0.92;
    const gap = 0.16;
    LAUNCH_STATUS.forEach((item, i) => {
      const cy = y0 + i * (cardH + gap);
      s.addShape("roundRect", {
        x: rx, y: cy, w: rw, h: cardH, rectRadius: 0.08,
        fill: { color: C.WHITE }, line: { color: item.color, width: 1.5 },
      });
      s.addShape("roundRect", {
        x: rx + 0.20, y: cy + (cardH - 0.52) / 2, w: 0.52, h: 0.52, rectRadius: 0.26,
        fill: { color: item.color },
      });
      s.addText(item.n, {
        x: rx + 0.20, y: cy + (cardH - 0.52) / 2, w: 0.52, h: 0.52,
        fontSize: 19, fontFace: FONT_H, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      s.addText(item.label, {
        x: rx + 0.88, y: cy, w: rw - 1.10, h: cardH,
        fontSize: 18, fontFace: FONT_B, color: C.CHARCOAL,
        valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
    });

    const cueY = y0 + 3 * (cardH + gap);
    addTextOnShape(s, "Fingers at your chest. Show me.", {
      x: rx, y: cueY, w: rw, h: 0.56, rectRadius: 0.08,
      fill: { color: C.SECONDARY },
    }, {
      fontSize: 17, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_LAUNCH);
    runSlideDiagnostics(s, pres);
  })();

  // 4. Learning Intention and Success Criteria
  liSlide(pres,
    "We are learning to draft body paragraphs and a conclusion from our plan.",
    [
      "I can start each paragraph with a topic sentence.",
      "I can keep each paragraph on one aspect.",
      "I can finish with an ending that sums up or adds one last fact.",
    ],
    NOTES_LI, FOOTER);

  // 5. I Do - notes become a paragraph
  (() => {
    const s = P.customSlide(pres, T, "I Do", C.PRIMARY, "Four dot points become a paragraph");
    const y0 = CONTENT_TOP;
    const h = SAFE_BOTTOM - y0;

    P.drawNoteCard(s, T, {
      x: 0.5, y: y0, w: 3.5, h,
      heading: "Aspect 2: Habitat",
      points: HABITAT_NOTES,
      color: C.SECONDARY,
      fontSize: 15,
    });

    const rx = 4.3;
    const rw = 5.2;
    const cardH = (h - 0.28) / 3;

    clickBuild(s, HABITAT_DRAFT.map((item, i) => () => {
      const cy = y0 + i * (cardH + 0.14);
      s.addShape("roundRect", {
        x: rx, y: cy, w: rw, h: cardH, rectRadius: 0.08,
        fill: { color: C.WHITE }, line: { color: item.color, width: 1.5 },
      });
      addTextOnShape(s, item.label, {
        x: rx, y: cy, w: rw, h: 0.32, rectRadius: 0.08,
        fill: { color: item.color },
      }, {
        fontSize: 12, fontFace: FONT_B, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      s.addText(item.text, {
        x: rx + 0.18, y: cy + 0.38, w: rw - 0.36, h: cardH - 0.50,
        fontSize: 15, fontFace: FONT_B, color: C.CHARCOAL,
        valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
    }));

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO);
    runSlideDiagnostics(s, pres);
  })();

  // 6. CFU hinge - which paragraph stays on one aspect?
  (() => {
    const s = P.customSlide(pres, T, "CFU", C.ALERT, "Which paragraph stays on one aspect?");
    const y0 = CONTENT_TOP;

    // Three tall options must clear the reveal bar:
    // 1.30 + 3 x 0.86 + 2 x 0.12 = 4.12.
    P.drawOptionStack(s, T, CFU_OPTIONS, {
      y: y0, optionH: 0.86, gap: 0.12, color: C.PRIMARY, fontSize: 14.5,
    });

    clickBuild(s, [
      () => {
        addRevealAnswerBar(s, ["B - every sentence is about where they live"], {
          y: P.REVEAL_Y, h: P.REVEAL_H, fontSize: 20, color: C.SUCCESS,
        });
      },
    ]);

    addFooter(s, FOOTER);
    s.addNotes(NOTES_CFU);
    runSlideDiagnostics(s, pres);
  })();

  // 7. We Do - the ending does one of two jobs
  (() => {
    const s = P.customSlide(pres, T, "We Do", C.SUCCESS, "An ending does one of two jobs");
    const y0 = CONTENT_TOP;
    const jobGap = 0.20;
    const jobW = (9 - jobGap) / 2;

    ENDING_JOBS.forEach((job, i) => {
      const jx = 0.5 + i * (jobW + jobGap);
      s.addShape("roundRect", {
        x: jx, y: y0, w: jobW, h: 0.78, rectRadius: 0.08,
        fill: { color: i === 0 ? C.PRIMARY : C.ASSESS },
      });
      s.addText([
        { text: job.label, options: { fontSize: 17, bold: true, color: C.WHITE, breakLine: true } },
        { text: job.detail, options: { fontSize: 12.5, color: C.WHITE } },
      ], {
        x: jx + 0.18, y: y0 + 0.04, w: jobW - 0.36, h: 0.70,
        fontFace: FONT_B, valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
    });

    const modelY = y0 + 0.94;
    P.drawReportExtract(s, T, {
      x: 0.5, y: modelY, w: 9, h: 1.30,
      bodySize: 15.5,
      showLabels: true,
      labelW: 2.3,
      sections: [{ label: "My ending, the frog", text: P.FROG.wrap, color: C.SECONDARY }],
    });

    addTextOnShape(s, "Now write an ending for tropical cyclones.", {
      x: 0.5, y: modelY + 1.42, w: 9, h: 0.60, rectRadius: 0.08,
      fill: { color: C.SUCCESS },
    }, {
      fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    clickBuild(s, [
      () => {
        addRevealAnswerBar(s, ["Tropical cyclones are among the most powerful storms on Earth. Australia's cyclone season runs from November to April."], {
          y: P.REVEAL_Y, h: P.REVEAL_H, fontSize: 14, label: "One way", color: C.SUCCESS,
        });
      },
    ]);

    addFooter(s, FOOTER);
    s.addNotes(NOTES_WEDO);
    runSlideDiagnostics(s, pres);
  })();

  // 8. You Do - draft the rest. This slide stays up while they write.
  (() => {
    const s = P.customSlide(pres, T, "You Do", C.ASSESS, "Draft the middle and the end");
    const y0 = CONTENT_TOP;
    const h = SAFE_BOTTOM - y0;

    T.addInstructionCard(s, [
      { text: "Your plan, your paragraphs", role: "header" },
      { text: "1. One paragraph for each aspect row." },
      { text: "2. Topic sentence first, every time." },
      { text: "3. Finish with an ending." },
      { text: "Thirty minutes. Checklist beside you." },
    ], {
      x: 0.5, y: y0, w: 4.5, h,
      strip: C.ASSESS, fill: C.WHITE,
      headerColor: C.ASSESS, emphasisColor: C.ALERT,
    });

    const tools = [
      { label: "Facts only", detail: "Could anyone check it?", color: C.PRIMARY },
      { label: "Present tense", detail: "True today, true always.", color: C.SECONDARY },
      { label: "Technical words", detail: "The exact word for your topic.", color: C.ACCENT },
      { label: "Noun groups", detail: "Describing words the reader can picture.", color: C.SUCCESS },
    ];
    const toolH = (h - 0.36) / 4;
    tools.forEach((tool, i) => {
      const ty = y0 + i * (toolH + 0.12);
      s.addShape("roundRect", {
        x: 5.2, y: ty, w: 4.3, h: toolH, rectRadius: 0.07,
        fill: { color: C.WHITE }, line: { color: tool.color, width: 1.4 },
      });
      s.addShape("roundRect", {
        x: 5.2, y: ty, w: 0.10, h: toolH, rectRadius: 0.03,
        fill: { color: tool.color },
      });
      s.addText([
        { text: tool.label, options: { fontSize: 15, bold: true, color: tool.color, breakLine: true } },
        { text: tool.detail, options: { fontSize: 12.5, color: C.CHARCOAL } },
      ], {
        x: 5.46, y: ty, w: 3.90, h: toolH,
        fontFace: FONT_B, valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
    runSlideDiagnostics(s, pres);
  })();

  // 9. Exit ticket
  P.exitTicketPanel(pres, T, {
    topic: "Look at the last sentence of your report",
    task: "Write SUM UP or LAST FACT, then copy your last sentence.",
    cue: "One word, then one sentence.",
    taskSize: 26,
    footer: FOOTER,
    notes: NOTES_EXIT,
  });

  // 10. Closing
  closingSlide(pres, {
    reflectionPrompt: "Which paragraph are you proudest of, and what made it work?",
    scItems: [
      "I can start each paragraph with a topic sentence.",
      "I can keep each paragraph on one aspect.",
      "I can finish with an ending that sums up or adds one last fact.",
    ],
    selfAssessment: "Fingers at your chest. One, two or three. Show me.",
    takeaways: [P.ANCHOR_PHRASE],
  }, NOTES_CLOSING);

  const pptxPath = path.join(OUT_DIR, "Information Reports W10 S2.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);
}

/* --- Resource: the Drafting Checklist ------------------------------------- */

async function buildChecklist() {
  const doc = createPdf({ title: CHECK_RES.name });
  let y = addPdfHeader(doc, "Drafting Checklist", {
    subtitle: "Keep this beside your draft. Check as you go, not at the end.",
    color: C.PRIMARY,
    lessonInfo: "Year 5/6 Literacy | Information Reports | Week 10",
    showNameDate: false,
  });

  y = addTipBox(doc, P.ANCHOR_PHRASE, y, { color: C.ACCENT });

  y = addSectionHeading(doc, "The shape of one body paragraph", y, { color: C.SECONDARY });
  const shape = [
    ["Topic sentence", "Names the aspect. Green tree frogs shelter in trees and reeds near still water."],
    ["Detail", "A fact from your notes, in your own words."],
    ["Detail", "Another fact, with a noun group so the reader can picture it."],
    ["Optional", "One more detail, or a technical word with its meaning."],
  ];
  shape.forEach((row, i) => {
    const accent = i === 0 ? hex(C.PRIMARY) : hex(C.SECONDARY);
    doc.save();
    doc.roundedRect(PAGE.MARGIN, y, 108, 20, 3).fill(accent);
    doc.fontSize(9.5).font("Sans-Bold").fillColor("#FFFFFF");
    doc.text(row[0], PAGE.MARGIN, y + 6, { width: 108, align: "center", lineBreak: false });
    doc.restore();
    doc.fontSize(10.5).font("Sans").fillColor("#000000");
    doc.text(row[1], PAGE.MARGIN + 118, y + 3, { width: PAGE.CONTENT_W - 126 });
    y = Math.max(doc.y, y + 20) + 8;
  });
  y += 4;

  y = addSectionHeading(doc, "The four tools, every paragraph", y, { color: C.ACCENT });
  y = P.addChecklistPdf(doc, y, [
    "Facts only. Could anyone check every sentence?",
    "Present tense. Is it true every day, not just once?",
    "At least one technical word, used correctly.",
    "At least one noun group the reader can picture.",
  ], { color: hex(C.ACCENT) });

  y = addSectionHeading(doc, "Your ending does ONE of these", y, { color: C.ASSESS });
  y = P.addChecklistPdf(doc, y, [
    "Sum it up: say your main idea again, in new words.",
    "One last fact: finish on the most interesting thing you found.",
  ], { color: hex(C.ASSESS) });

  y = addSectionHeading(doc, "Before you say you have finished", y, { color: C.SUCCESS });
  y = P.addChecklistPdf(doc, y, [
    "One paragraph for each aspect on my plan.",
    "Every paragraph starts with a topic sentence.",
    "No paragraph drifts into a different aspect.",
    "No opinions anywhere.",
    "Capitals and full stops in every sentence.",
  ], { color: hex(C.SUCCESS) });

  y = addBodyText(doc,
    "Stuck on a paragraph? Cover your plan, say the sentence out loud, then write exactly what you said.",
    y, { italic: true });

  y = addSectionHeading(doc, "Technical words you might use", y, { color: C.PRIMARY });
  y = addBodyText(doc,
    "Every topic has its own. Write the ones you meet in your research here, with their meanings.", y);
  y = P.addNoteBoxPdf(doc, y, "My technical words", { color: hex(C.PRIMARY), lines: 4, spacing: 24 });

  addPdfFooter(doc, "Information Reports | Year 5/6 Literacy | Week 10");
  await writePdf(doc, path.join(OUT_DIR, CHECK_RES.fileName));
  console.log("Wrote " + CHECK_RES.name);
}

(async () => {
  await build();
  await buildChecklist();
})().catch((err) => {
  console.error(err);
  process.exit(1);
});
