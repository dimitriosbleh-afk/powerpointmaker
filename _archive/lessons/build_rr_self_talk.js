"use strict";

// Respectful Relationships (Resilience, Rights and Respectful Relationships)
// Topic 3: Positive coping - Activity 1: Introducing the concept of self-talk
// Grade 5/6 Wellbeing - first RR session for the term (single session)
//
// Source: RRRR Level 5-6, Topic 3 "Positive coping", Activity 1. Lan's Day
// scenario and the two-thought-bubble map are taken from the source activity.

const path = require("path");
const fs = require("fs");
const PptxGenJS = require("pptxgenjs");
const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf,
  addPdfHeader, addSectionHeading, addBodyText, addTipBox,
  addPdfFooter, addLinedArea,
  addResourceSlide, makeSessionResource,
  getSessionResourceFolder,
  hex,
} = require("../themes/pdf_helpers");

// ── Theme (first RR session for the term -> week 1 -> variant 0) ──
const T = createTheme("wellbeing", "grade56", weekToVariant(1));
const {
  C, FONT_H, FONT_B,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  keyWordSlide, exitTicketSlide, scenarioSlide,
  addCard, addTextOnShape, addRevealAnswerBar, withReveal,
  composeGlanceNotes, composeRevealNotes,
  CONTENT_TOP, SAFE_BOTTOM,
} = T;

// ── Output paths ──
const UNIT = "RR_Positive_Coping_Self_Talk";
const LESSON_FOLDER = path.join(__dirname, "..", "output", UNIT);
const PPTX_NAME = "Positive Coping Self-Talk.pptx";
const FOOTER = "Respectful Relationships | Grade 5/6 Wellbeing";
const SESSION = 1;
const RES_DIR = path.join(LESSON_FOLDER, getSessionResourceFolder(SESSION));

// ── Resource ──
const HELPER_RESOURCE = makeSessionResource(
  SESSION,
  "Positive Self-Talk Helper",
  "Reference card of kind, true self-talk students can borrow, with a worked flip and a space to try their own."
);
const RESOURCE_ITEMS = [HELPER_RESOURCE];

fs.mkdirSync(RES_DIR, { recursive: true });

// ============================================================
// Shared visual: the two-bubble "flip" anchor (unit anchor)
//   Unhelpful self-talk (top) -> Flip it -> Helpful self-talk (bottom)
// Used on Launch, I Do, We Do and You Do so the representation is
// identical every time (megaprompt section 79).
// ============================================================

function drawFlipBubbles(s, opts) {
  const o = opts || {};
  const x = o.x;
  const y = o.y;
  const w = o.w;
  const h = o.h;
  const gap = 0.12;
  const flipH = 0.40;
  const bubbleH = (h - flipH - gap * 2) / 2;

  // Top bubble: unhelpful self-talk (ALERT)
  addCard(s, x, y, w, bubbleH, { strip: C.ALERT, fill: C.WHITE });
  s.addText("Unhelpful self-talk", {
    x: x + 0.2, y: y + 0.09, w: w - 0.4, h: 0.26,
    fontSize: 13, fontFace: FONT_B, color: C.ALERT, bold: true, margin: 0,
  });
  if (o.neg) {
    s.addText(String(o.neg), {
      x: x + 0.2, y: y + 0.38, w: w - 0.4, h: bubbleH - 0.46,
      fontSize: 15, fontFace: FONT_B, color: C.CHARCOAL,
      valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
    });
  } else if (o.negHint) {
    s.addText(String(o.negHint), {
      x: x + 0.2, y: y + 0.38, w: w - 0.4, h: bubbleH - 0.46,
      fontSize: 13, fontFace: FONT_B, color: C.MUTED,
      italic: true, valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
    });
  }

  // Flip connector chip
  const flipY = y + bubbleH + gap;
  const flipW = 1.5;
  addTextOnShape(s, "Flip it", {
    x: x + (w - flipW) / 2, y: flipY, w: flipW, h: flipH, rectRadius: 0.08,
    fill: { color: C.PRIMARY },
  }, {
    fontSize: 14, fontFace: FONT_B, color: C.WHITE, bold: true,
    align: "center", valign: "middle", margin: 0,
  });

  // Bottom bubble: helpful self-talk (SUCCESS)
  const bY = flipY + flipH + gap;
  addCard(s, x, bY, w, bubbleH, { strip: C.SUCCESS, fill: C.WHITE });
  s.addText("Helpful self-talk", {
    x: x + 0.2, y: bY + 0.09, w: w - 0.4, h: 0.26,
    fontSize: 13, fontFace: FONT_B, color: C.SUCCESS, bold: true, margin: 0,
  });
  if (o.pos) {
    s.addText(String(o.pos), {
      x: x + 0.2, y: bY + 0.38, w: w - 0.4, h: bubbleH - 0.46,
      fontSize: 15, fontFace: FONT_B, color: C.CHARCOAL,
      valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
    });
  } else if (o.posHint) {
    s.addText(String(o.posHint), {
      x: x + 0.2, y: bY + 0.38, w: w - 0.4, h: bubbleH - 0.46,
      fontSize: 13, fontFace: FONT_B, color: C.MUTED,
      italic: true, valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
    });
  }
}

// ============================================================
// Teacher notes (Glance Format v12.3)
// ============================================================

const NOTES_TITLE =
  "Display as students arrive. First Respectful Relationships session this term. " +
  "Settle the room and re-establish the class agreement before you begin.";

const NOTES_RESOURCES = [
  "Materials: mini-whiteboards and markers, the Positive Self-Talk Helper (one per student), board space for a two-bubble map.",
  "Prep: this is the first RR session this term. Re-establish the class agreement (listen, respect, you may pass) before slide 3.",
  "CARE: self-talk touches on distress and anxiety. Keep examples about the fictional characters, not students' own hard moments.",
  "If a student shows their inner voice is often unkind or they feel low a lot, follow up quietly and refer to a trusted adult or wellbeing staff.",
  "Print the Positive Self-Talk Helper before the lesson. Students read it in the We Do and use it again in the You Do. No answer key needed (open responses).",
].join("\n");

const NOTES_LAUNCH = composeGlanceNotes({
  answer: "the unhelpful thought - it makes the upset feeling bigger",
  beats: [
    ["SAY: Some of you may remember how feelings go up and down.",
     "Today we look at how our thinking changes them."],
    ["POINT to both thoughts. SAY: Same moment, a dropped catch.",
     "Two very different things to say to yourself."],
    ["ASK: Which thought makes the upset feeling bigger?",
     "10 sec. Turn and tell your partner.",
     "EXPECT: the unhelpful one, I always drop it, I'm useless."],
  ],
  trap: ["thinking the event causes the feeling, not the thought.",
         "Fix: point to both, same event. Student re-says which is bigger."],
  prep: "Launch: bridges prior work on emotions to today's self-talk. Keep it brisk, under 4 minutes.",
  tag: "[Launch | VTLM 2.0: Retention and recall | HITS 6]",
});

const NOTES_LI = composeGlanceNotes({
  beats: [
    ["SAY: Read the learning intention with me."],
    ["SAY: Three success criteria. We check these at the end.",
     "Read each one aloud."],
    ["POINT to each I can statement as you read it."],
  ],
  prep: "SC1 is the floor everyone reaches. SC2 is the core, checked by the exit ticket.",
  tag: "[LI/SC | VTLM 2.0: Planning made visible | HITS 1]",
});

const NOTES_VOCAB = composeGlanceNotes({
  beats: [
    ["SAY: Self-talk is the words you say to yourself inside your head."],
    ["SAY: Miss the ball? Your self-talk might say nice try,",
     "or it might say you're hopeless."],
    ["SAY: Say it with me on three. One, two, three... self-talk."],
  ],
  prep: "Vocab comes after LI/SC. The one word students need for the whole lesson.",
  tag: "[Vocabulary | VTLM 2.0: Knowledge and memory | HITS 3]",
});

const NOTES_IDO = composeGlanceNotes({
  beats: [
    ["MODEL. SAY: Watch how I handle a low mark on a test."],
    ["SAY: My first thought is, I'm no good, give up.",
     "Notice it. That is unhelpful self-talk."],
    ["SAY: I'll flip it. That's disappointing, but I'll do better.",
     "Same mark, but now I can keep going."],
  ],
  trap: ["students think positive self-talk means pretending it went well.",
         "Fix: model honest and kind, not fake. Re-say the flip."],
  stretch: "explain how the helpful thought changes what you do next.",
  help: "sentence stem That was hard, but I can... for students stuck on words.",
  prep: "Model the Notice, Name, Flip process. Anchor phrase: same moment, different thought.",
  tag: "[I Do | VTLM 2.0: Explicit teaching | HITS 4]",
});

const NOTES_CFU = composeGlanceNotes({
  answer: "B - it is kind and keeps you going",
  beats: [
    ["SAY: Two thoughts after losing a game. One helps, one does not."],
    ["ASK: Which is positive self-talk, A or B?",
     "10 sec. Write A or B... chin it... show me.",
     "EXPECT: B"],
    ["SCAN boards.",
     "80%+ -> cold call one B: how do you know? Then reveal.",
     "Less -> re-read A, ask does it help or hurt? Re-ask A or B."],
  ],
  trap: ["picking A because it sounds realistic.",
         "Fix: kind AND true is the test. Student re-checks B."],
  prep: "The hinge decision point before We Do. Wrong A confuses realistic with helpless.",
  tag: "[CFU hinge | VTLM 2.0: Supported application | SC2 | HITS 7, 8]",
});

const NOTES_CFU_REVEAL = composeRevealNotes({
  answer: "B - That was tough, I'll get the next one. Kind and true.",
  beats: [
    ["SAY: B is kind AND honest. It keeps her trying."],
    ["Tick or fix your board. Cold call one A board: what would help more?"],
  ],
  prep: "Reveal after boards scanned.",
});

const NOTES_MEETLAN = composeGlanceNotes({
  answer: "open - listen for the own goal as the hardest moment",
  beats: [
    ["SAY: This is Lan. A lot went wrong in one day.",
     "Read her day with me."],
    ["ASK: Which moment would be hardest for her self-talk?",
     "10 sec. Turn and tell.",
     "EXPECT: the own goal that lost the game."],
    ["SAY: We will help Lan flip her self-talk about the own goal."],
  ],
  care: "keep this about Lan. If a student takes it personally, acknowledge and move on.",
  prep: "We Do orientation. Pick one moment so the flip stays focused.",
  tag: "[We Do | VTLM 2.0: Supported application | HITS 5]",
});

const NOTES_WEDO = composeGlanceNotes({
  answer: "open - listen for kind and true flips that keep Lan going",
  beats: [
    ["SAY: Lan's unhelpful thought is up top. Let's flip it."],
    ["ASK: What could Lan tell herself instead?",
     "20 sec. Boards up on cue. Use your Helper.",
     "EXPECT: kind, true self-talk, one mistake is not the whole game."],
    ["SCAN boards.",
     "80%+ -> cold call two flips, then reveal.",
     "Less -> read one Helper line together, re-ask for a flip."],
  ],
  stretch: "explain how the flip changes what Lan does at the next game.",
  help: "point to two Helper lines and choose the one that fits, for stuck students.",
  prep: "Guided flip. The Helper is the scaffold and the catch-up entry point.",
  tag: "[We Do | VTLM 2.0: Supported application | SC3 | HITS 4, 5]",
});

const NOTES_WEDO_REVEAL = composeRevealNotes({
  answer: "helpful self-talk: one mistake is not the whole game; I'll practise and go again",
  beats: [
    ["SAY: These are kind AND true. They help Lan keep going."],
    ["Tick or fix your board. Cold call one student: why does this help?"],
  ],
  prep: "Reveal after boards scanned.",
});

const NOTES_YOUDO = composeGlanceNotes({
  answer: "open - kind, true self-talk, e.g. everyone forgets; I can move on",
  beats: [
    ["SAY: Your turn. New moment, on the slide."],
    ["CIRCULATE. Read boards. Prompt stuck students with the Helper."],
    ["TIME: about 5 minutes. Then we share two flips."],
  ],
  stretch: "write a second, different helpful thought for the same moment.",
  help: "give the stem It was embarrassing, but I can... and the Helper.",
  prep: ["You Do uses a new moment, not Lan, so We Do and You Do differ.",
         "Catch-up: the Helper lets a student who missed earlier work start."],
  tag: "[You Do | VTLM 2.0: Mastery and application | HITS 10]",
});

const NOTES_EXIT = composeGlanceNotes({
  answer: "open - a kind, true thought that keeps them going",
  beats: [
    ["SAY: On your own. Flip the unhelpful thought to helpful self-talk.",
     "e.g. this is hard, I can keep trying."],
    ["Write it. Kind AND true. Hand it in."],
  ],
  prep: "Assesses SC2 and reaches into SC3. Sort into got-it and not-yet.",
  tag: "[Exit Ticket | VTLM 2.0: Mastery and application | SC2]",
});

const NOTES_CLOSING = composeGlanceNotes({
  answer: "open - listen for a real place they could flip their self-talk",
  beats: [
    ["SAY: Read each I can statement. Thumbs up, sideways, or down."],
    ["ASK: Where could positive self-talk help you this week?",
     "15 sec. Turn and tell one place.",
     "EXPECT: a real situation, a test, a game, a fall-out."],
    ["SAY: If your inner voice is often unkind, talk to a trusted adult."],
  ],
  care: "some students carry heavy self-talk. Note quiet thumbs-down and follow up privately.",
  prep: "Closing ties back to the SC and to help-seeking.",
  tag: "[Closing | VTLM 2.0: Retention and recall | HITS 9]",
});

// Lan's Day scenario (faithful to the source, chunked for students)
const LAN_SCENARIO =
  "Lan finished only 9 of the 20 questions on her maths test. At interschool " +
  "sport she found she had left her lunch at home, so her friends shared theirs. " +
  "In the last minute of a 3-all soccer game, the ball flicked off her hand and " +
  "into the goal, and her team lost.";

// ============================================================
// Build
// ============================================================

async function build() {
  const pres = new PptxGenJS();
  pres.layout = "LAYOUT_16x9";

  // -- 1. Title --
  titleSlide(
    pres,
    "Positive Coping:\nSelf-Talk",
    "Respectful Relationships",
    "Grade 5/6 Wellbeing  |  Session 1",
    NOTES_TITLE
  );

  // -- 2. Teacher Resources --
  addResourceSlide(pres, RESOURCE_ITEMS, T, FOOTER, NOTES_RESOURCES);

  // -- 3. Launch: same moment, two thoughts --
  contentSlide(
    pres,
    "Launch",
    C.SECONDARY,
    "Same Moment, Two Thoughts",
    [
      "Same moment. Two different thoughts.",
      "Read both thoughts in the bubbles.",
      "Which one makes the upset feeling bigger?",
    ],
    NOTES_LAUNCH,
    FOOTER,
    (s, lg) => {
      drawFlipBubbles(s, {
        x: lg.rightX, y: lg.panelTopPadded, w: lg.rightW,
        h: 5.0 - lg.panelTopPadded,
        neg: "I always drop it. I'm useless.",
        pos: "Nice try. I'll get the next one.",
      });
    }
  );

  // -- 4. Learning Intention & Success Criteria --
  liSlide(
    pres,
    ["We are learning to notice our self-talk and use positive self-talk to help us cope when things go wrong"],
    [
      "I can explain what self-talk is",
      "I can tell positive self-talk from negative self-talk",
      "I can change negative self-talk into positive self-talk",
    ],
    NOTES_LI,
    FOOTER
  );

  // -- 5. Key Vocabulary: self-talk --
  keyWordSlide(
    pres,
    {
      word: "self-talk",
      meaning: "The words you say to yourself inside your head.",
      example: "Miss the ball? Your self-talk might say 'nice try' or 'you're hopeless'.",
      routine: ["Say it", "Notice it", "Use it"],
      color: C.PRIMARY,
      title: "Word for today",
    },
    NOTES_VOCAB,
    FOOTER
  );

  // -- 6. I Do: model the flip --
  contentSlide(
    pres,
    "I Do",
    C.PRIMARY,
    "Watch Me Flip a Thought",
    [
      "Notice it: catch the thought.",
      "Name it: helpful or unhelpful?",
      "Flip it: kinder and still true.",
    ],
    NOTES_IDO,
    FOOTER,
    (s, lg) => {
      drawFlipBubbles(s, {
        x: lg.rightX, y: lg.panelTopPadded, w: lg.rightW,
        h: 5.0 - lg.panelTopPadded,
        neg: "I'm no good at this. I may as well give up.",
        pos: "That's disappointing, but I'll work at doing better.",
      });
    }
  );

  // -- 7. CFU hinge (reveal pair): which is positive self-talk? --
  withReveal(
    () => cfuSlide(
      pres,
      "Hinge",
      "Which Is Positive Self-Talk?",
      "Show Me Boards",
      "A: I always mess up. No point trying.\nB: That was tough. I'll get the next one.\n\nShow A or B on your board.",
      NOTES_CFU,
      FOOTER
    ),
    (s) => {
      addRevealAnswerBar(s, ["B - kind and keeps you going"], { color: C.SUCCESS });
    },
    { revealNotes: NOTES_CFU_REVEAL }
  );

  // -- 8. We Do: meet Lan (scenario) --
  scenarioSlide(
    pres,
    "We Do",
    "Meet Lan",
    LAN_SCENARIO,
    [
      "Which moment would be hardest for Lan's self-talk?",
      "What might Lan start telling herself?",
    ],
    NOTES_MEETLAN,
    FOOTER
  );

  // -- 9. We Do: flip Lan's thought (reveal pair) --
  withReveal(
    () => contentSlide(
      pres,
      "We Do",
      C.SUCCESS,
      "Flip Lan's Thought",
      [
        "Lan's unhelpful thought is up top.",
        "Turn and tell: what could she say instead?",
        "Use your Positive Self-Talk Helper.",
        "Write your flip on your board.",
      ],
      NOTES_WEDO,
      FOOTER,
      (s, lg) => {
        drawFlipBubbles(s, {
          x: lg.rightX, y: lg.panelTopPadded, w: lg.rightW,
          h: 4.1 - lg.panelTopPadded,
          neg: "I lost the game for everyone. I'm hopeless at sport.",
          posHint: "Flip it on your board...",
        });
      }
    ),
    (s) => {
      addRevealAnswerBar(
        s,
        ["One mistake is not the whole game", "I'll practise and try again"],
        { color: C.SUCCESS }
      );
    },
    { revealNotes: NOTES_WEDO_REVEAL }
  );

  // -- 10. You Do: flip your own thought (new moment) --
  contentSlide(
    pres,
    "You Do",
    C.ACCENT,
    "Flip Your Own Thought",
    [
      "Moment: you forgot your words at assembly and some kids laughed.",
      "First: write the unhelpful thought up top.",
      "Then: flip it to helpful self-talk below.",
      "Use your Helper if you get stuck.",
    ],
    NOTES_YOUDO,
    FOOTER,
    (s, lg) => {
      drawFlipBubbles(s, {
        x: lg.rightX, y: lg.panelTopPadded, w: lg.rightW,
        h: 5.0 - lg.panelTopPadded,
        negHint: "Your unhelpful thought...",
        posHint: "Your helpful self-talk...",
      });
    }
  );

  // -- 11. Exit Ticket --
  exitTicketSlide(
    pres,
    "Flip this unhelpful thought into helpful self-talk:\n'I'm terrible at this. I might as well give up.'",
    NOTES_EXIT,
    FOOTER,
    { assessesSc: 2 }
  );

  // -- 12. Closing --
  closingSlide(
    pres,
    {
      reflectionPrompt: "Where could positive self-talk help you this week?",
      scItems: [
        "I can explain what self-talk is",
        "I can tell positive self-talk from negative self-talk",
        "I can change negative self-talk into positive self-talk",
      ],
      selfAssessment: {
        prompt: "Thumbs up, sideways, or down for each I can statement.",
        options: ["Got it", "Getting there", "Need practice"],
      },
    },
    NOTES_CLOSING
  );

  // Write PPTX
  fs.mkdirSync(LESSON_FOLDER, { recursive: true });
  await pres.writeFile({ fileName: path.join(LESSON_FOLDER, PPTX_NAME) });
  console.log("PPTX written to", path.join(LESSON_FOLDER, PPTX_NAME));

  // ============================================================
  // Companion PDF: Positive Self-Talk Helper
  // ============================================================
  {
    const doc = createPdf({ title: "Positive Self-Talk Helper" });
    let y = addPdfHeader(doc, "Session 1 Positive Self-Talk Helper", {
      subtitle: "Kind, true self-talk you can borrow",
      color: C.PRIMARY,
      lessonInfo: "Grade 5/6 Wellbeing - Respectful Relationships",
    });

    y = addTipBox(
      doc,
      "Self-talk is what you say to yourself inside your head. When it turns unhelpful, you can flip it to something kinder that is still true.",
      y,
      { color: C.ACCENT }
    );

    y = addSectionHeading(doc, "How to flip a thought", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Unhelpful: I'll never understand this.", y, { fontSize: 11 });
    y = addBodyText(doc, "Helpful: This is tricky now, but I can get better with practice.", y, { fontSize: 11 });

    y = addSectionHeading(doc, "Helpful self-talk to borrow", y, { color: C.SUCCESS });

    const GROUPS = [
      { when: "When I make a mistake", lines: [
        "Everyone makes mistakes. This is how I learn.",
        "One mistake is not the whole story.",
      ] },
      { when: "When something feels hard", lines: [
        "This is hard now, but I can keep going.",
        "I'll try a different way.",
      ] },
      { when: "When I lose or things go wrong", lines: [
        "One mistake is not the whole game.",
        "I'll practise and try again.",
      ] },
      { when: "When I feel upset or left out", lines: [
        "This feeling will pass. I can ask for help.",
        "I did my best today.",
      ] },
    ];

    GROUPS.forEach((g) => {
      y = addSectionHeading(doc, g.when, y, { color: C.SECONDARY, fontSize: 11 });
      y = addBodyText(doc, g.lines.map((l) => "- " + l).join("\n"), y, { fontSize: 11 });
    });

    y = addSectionHeading(doc, "Your flip", y, { color: C.ACCENT });
    y = addBodyText(doc, "An unhelpful thought I sometimes have:", y, { fontSize: 11 });
    y = addLinedArea(doc, y, 1, { lineSpacing: 24 });
    y = addBodyText(doc, "Flip it to helpful self-talk (kind and true):", y, { fontSize: 11 });
    y = addLinedArea(doc, y, 2, { lineSpacing: 24 });

    addPdfFooter(doc, "Respectful Relationships | Grade 5/6 | Session 1");
    // Write to the exact path the resource slide links to (cleanResourceLabel
    // strips the hyphen from "Self-Talk", so derive from the resource fileName).
    const helperPath = path.join(LESSON_FOLDER, HELPER_RESOURCE.fileName);
    await writePdf(doc, helperPath);
    console.log("PDF written:", helperPath);
  }

  console.log("\nBuild complete!");
  console.log("Output folder:", LESSON_FOLDER);
}

build().catch((err) => { console.error(err); process.exit(1); });
