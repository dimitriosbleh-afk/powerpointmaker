"use strict";

// Respectful Relationships (Resilience, Rights and Respectful Relationships)
// Topic 3 "Positive coping", Level 5-6 - Session 1 of 4
// Source activity 1: Introducing the concept of self-talk
//
// Source-locked from the supplied RRRR Topic 3 PDF:
//   - the definition of self-talk
//   - the missed-catch and low-test-result examples
//   - the Lan's Day scenario
//   - the two-thought-bubble mapping method
//   - every line of the Positive self-talk handout (9 statements)

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
} = require("../themes/pdf_helpers");

// Theme: one variant for the whole unit (theme cohesion rule).
const T = createTheme("wellbeing", "grade56", weekToVariant(1));
const {
  C, FONT_H, FONT_B,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  keyWordSlide, exitTicketSlide, scenarioSlide,
  addCard, addTextOnShape, addRevealAnswerBar, addTopBar, addBadge, addTitle,
  clickBuild, runSlideDiagnostics,
  composeGlanceNotes,
  CONTENT_TOP, SAFE_BOTTOM,
} = T;

const UNIT_FOLDER = "RR3_Session1_Notice_Your_Self_Talk";
const LESSON_FOLDER = path.join(__dirname, "..", "output", UNIT_FOLDER);
const PPTX_NAME = "Session 1 Notice Your Self-Talk.pptx";
const SESSION = 1;
const FOOTER = "Respectful Relationships | Grade 5/6 | Session 1";
const RES_DIR = path.join(LESSON_FOLDER, getSessionResourceFolder(SESSION));

const HELPER = makeSessionResource(
  SESSION,
  "Positive Self-Talk Helper",
  "The positive self-talk statements from the RRRR handout, plus a worked flip and space to try your own."
);
const RESOURCE_ITEMS = [HELPER];

fs.mkdirSync(RES_DIR, { recursive: true });

/* ══════════════════════════════════════════════════════════════════════
   Unit anchor (megaprompt section 79)

   Representation : the two-bubble flip map
   Anchor phrase  : "Same moment, different choice."
   Method         : Notice it. Name it. Flip it.

   Sessions 1 and 2 flip a thought. Sessions 3 and 4 flip what you do.
   The drawing, the phrase and the three steps stay identical all unit.
   ══════════════════════════════════════════════════════════════════════ */

function drawFlipMap(s, opts) {
  const o = opts || {};
  const x = o.x;
  const y = o.y;
  const w = o.w;
  const h = o.h;
  const gap = 0.12;
  const flipH = 0.40;
  const bubbleH = (h - flipH - gap * 2) / 2;

  const topLabel = o.topLabel || "Unhelpful self-talk";
  const bottomLabel = o.bottomLabel || "Helpful self-talk";

  // Top bubble: the thought that turns up on its own.
  addCard(s, x, y, w, bubbleH, { strip: C.ALERT, fill: C.WHITE });
  s.addText(topLabel, {
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

  // "Flip it" connector.
  const flipY = y + bubbleH + gap;
  const flipW = 1.5;
  addTextOnShape(s, "Flip it", {
    x: x + (w - flipW) / 2, y: flipY, w: flipW, h: flipH, rectRadius: 0.08,
    fill: { color: C.PRIMARY },
  }, {
    fontSize: 14, fontFace: FONT_B, color: C.WHITE, bold: true,
    align: "center", valign: "middle", margin: 0,
  });

  // Bottom bubble: the thought you choose.
  const bY = flipY + flipH + gap;
  addCard(s, x, bY, w, bubbleH, { strip: C.SUCCESS, fill: C.WHITE });
  s.addText(bottomLabel, {
    x: x + 0.2, y: bY + 0.09, w: w - 0.4, h: 0.26,
    fontSize: 13, fontFace: FONT_B, color: C.SUCCESS, bold: true, margin: 0,
  });

  const posBox = {
    x: x + 0.2, y: bY + 0.38, w: w - 0.4, h: bubbleH - 0.46,
  };
  const writePos = (text, style) => {
    s.addText(String(text), Object.assign({}, posBox, {
      fontSize: style === "hint" ? 13 : 15,
      fontFace: FONT_B,
      color: style === "hint" ? C.MUTED : C.CHARCOAL,
      italic: style === "hint",
      valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
    }));
  };

  if (o.pos) writePos(o.pos);
  else if (o.posHint) writePos(o.posHint, "hint");

  // Staged reveal: the empty bubble is on screen from the start so students
  // can see there is a slot to fill; the click drops the answer into it.
  return {
    revealPos: () => { if (o.stagePos) writePos(o.stagePos); },
  };
}

/* ══════════════════════════════════════════════════════════════════════
   Two-option hinge slide (unit layout)

   cfuSlide grows its question card to within 0.14" of a reveal bar, so a
   click-revealed answer overlaps it. This layout puts the two options in
   side-by-side cards - the classic A/B hinge - and leaves the bottom strip
   free for the reveal. Same layout in all four sessions.
   ══════════════════════════════════════════════════════════════════════ */

const HINGE_CARD_Y = CONTENT_TOP + 0.58;
const HINGE_CARD_H = 2.15;
const HINGE_BAR_Y = HINGE_CARD_Y + HINGE_CARD_H + 0.19;

function hingeSlide(pres, o) {
  const s = pres.addSlide();
  addTopBar(s, C.ALERT);
  addBadge(s, "Hinge", { color: C.ALERT });
  addTitle(s, o.title, { color: C.ALERT });

  // Label as well as colour (megaprompt section 18a).
  addTextOnShape(s, "CHECK", {
    x: 8.2, y: 0.20, w: 1.3, h: 0.32, rectRadius: 0.08,
    fill: { color: C.WHITE }, line: { color: C.ALERT, width: 1.5 },
  }, {
    fontSize: 11, fontFace: FONT_B, color: C.ALERT, bold: true,
    align: "center", valign: "middle", margin: 0,
  });

  addTextOnShape(s, o.routine || "Show Me Boards", {
    x: 0.5, y: CONTENT_TOP, w: 2.8, h: 0.42, rectRadius: 0.08,
    fill: { color: C.ALERT },
  }, {
    fontSize: 14, fontFace: FONT_B, color: C.WHITE, bold: true,
    align: "center", valign: "middle", margin: 0,
  });

  [["A", o.optionA, 0.5], ["B", o.optionB, 5.1]].forEach((entry) => {
    const letter = entry[0];
    const text = entry[1];
    const x = entry[2];
    addCard(s, x, HINGE_CARD_Y, 4.4, HINGE_CARD_H, { strip: C.PRIMARY, fill: C.WHITE });
    s.addShape("roundRect", {
      x: x + 0.2, y: HINGE_CARD_Y + 0.16, w: 0.5, h: 0.5, rectRadius: 0.25,
      fill: { color: C.PRIMARY },
    });
    s.addText(letter, {
      x: x + 0.2, y: HINGE_CARD_Y + 0.16, w: 0.5, h: 0.5,
      fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
    s.addText(String(text), {
      x: x + 0.2, y: HINGE_CARD_Y + 0.80, w: 4.0, h: HINGE_CARD_H - 0.96,
      fontSize: 19, fontFace: FONT_B, color: C.CHARCOAL,
      valign: "top", margin: 0, fit: "shrink", shrinkText: true,
    });
  });

  clickBuild(s, [
    () => {
      addRevealAnswerBar(s, [o.answer], {
        y: HINGE_BAR_Y, h: 0.72, fontSize: 22, color: C.SUCCESS,
      });
    },
  ]);

  T.addFooter(s, FOOTER);
  s.addNotes(o.notes);
  runSlideDiagnostics(s, pres);
  return s;
}

/* ══════════════════════════════════════════════════════════════════════
   Source text (locked, megaprompt section 5a)
   ══════════════════════════════════════════════════════════════════════ */

// RRRR Topic 3, Activity 1. Chunked for reading aloud; wording follows the
// source. The source's "into the other team's goal ... her team was defeated"
// is self-contradictory, so the goal is left unattributed rather than guessed.
const LAN_SCENARIO =
  "Lan struggled through a maths test and could only do nine of the 20 questions. " +
  "She was looking forward to interschool sport because she was playing soccer, her " +
  "favourite. She went to get her lunch out of her bag but she had left it on the " +
  "kitchen bench, so her friends shared their food. In the final minute the score was " +
  "three all. Lan stopped an attack, but the ball flicked off her hand and into the " +
  "goal. The siren sounded and her team was defeated.";

// The nine statements on the RRRR "Positive self-talk" handout, verbatim.
const HANDOUT_LINES = [
  "Even if I don't get a good mark, at least I know I tried hard.",
  "It might be lonely at first but I will eventually get to know people and settle in.",
  "Even if it is scary to talk in front of the school, it won't last that long and I can keep control of my nerves.",
  "I am going to stay calm and focused and give this my best effort.",
  "It hurts when people say mean things, but that meanness is more about them than me, and I don't have to believe what they say.",
  "I have stuck at things before, so I am not going to give up this time.",
  "Even though I missed a shot, I tried my best.",
  "The team will be disappointed, but there's another game next week.",
  "I can stick at this.",
];

/* ══════════════════════════════════════════════════════════════════════
   Teacher notes (Glance Format v12.3)
   ══════════════════════════════════════════════════════════════════════ */

const NOTES_TITLE =
  "Display as students arrive. Session 1 of 4 in Positive Coping. " +
  "Settle the room and re-establish the class agreement before you begin.";

const NOTES_RESOURCES = [
  "Materials: mini-whiteboards and markers, Positive Self-Talk Helper (one each), board space for a two-bubble map.",
  "Prep: first session of this unit. Re-establish the class agreement (listen, respect, you may pass) before the launch.",
  "Decision points: the hinge after I Do, the boards check in We Do, the exit ticket. Between those, keep it brisk.",
  "CARE: self-talk touches distress and anxiety. Keep every example about the fictional character, never a student's own hard moment.",
  "If a student's inner voice sounds often unkind, or they seem low, follow up quietly and involve wellbeing staff.",
  "SOURCES: RRRR Level 5-6, Topic 3 Positive coping, Activity 1 and the Positive self-talk handout.",
].join("\n");

const NOTES_OVERVIEW = [
  "Teacher-facing. Read before teaching, skip past it in front of the class.",
  "Unit anchor: the flip map, the phrase 'same moment, different choice', and Notice it, Name it, Flip it. Hold all three identical across the four sessions.",
  "Catch-up: every session opens on a flip map a student can do with no prior session. Hand a returner the Session 1 Helper first.",
  "Reflect on your own practice: how does your language model positive talk, and what self-talk do you use when a lesson goes wrong?",
].join("\n");

const NOTES_LAUNCH = composeGlanceNotes({
  answer: "the unhelpful one - it makes the upset feeling bigger",
  beats: [
    ["SAY: Some of you may remember that feelings go up and down.",
     "Today we look at how our thinking changes them."],
    ["POINT to both bubbles. SAY: One dropped catch.",
     "Two very different things to say to yourself."],
    ["ASK: Which thought makes the upset feeling bigger?",
     "10 sec. Turn and tell. Eyes back in 3, 2, 1.",
     "EXPECT: the top one, I always drop it, I'm useless."],
  ],
  trap: ["thinking the event causes the feeling, not the thought.",
         "Fix: cover the bubbles, same catch either way. Student re-says which is bigger."],
  prep: "Launch: bridges earlier work on emotions into self-talk. Whole block under 4 minutes.",
  tag: "[Launch | VTLM 2.0: Retention and recall | HITS 6]",
});

const NOTES_LI = composeGlanceNotes({
  beats: [
    ["SAY: Read the learning intention with me."],
    ["POINT to each I can statement as you read it aloud."],
    ["SAY: We check these three at the end of the session."],
  ],
  prep: "The first criterion is the floor everyone reaches. The second is the core the exit ticket checks.",
  tag: "[LI/SC | VTLM 2.0: Planning made visible | HITS 1]",
});

const NOTES_VOCAB = composeGlanceNotes({
  beats: [
    ["SAY: Self-talk is the talking you do inside your own head."],
    ["SAY: When you were little you thought out loud.",
     "Now most of that talk happens inside."],
    ["SAY: Say it with me. Everyone, together, on three.",
     "One, two, three... self-talk."],
  ],
  prep: "Vocabulary sits after the LI so students know why the word matters. The one word the whole session runs on.",
  sources: "definition follows RRRR Topic 3 Activity 1",
  tag: "[Vocabulary | VTLM 2.0: Knowledge and memory | HITS 3]",
});

const NOTES_IDO = composeGlanceNotes({
  beats: [
    ["MODEL. SAY: Watch me. I have just got a low mark on a test."],
    ["SAY: My first thought is, I am no good at this.",
     "I might as well give up. Notice it. Name it. Unhelpful."],
    ["CLICK to flip. SAY: That's disappointing, but I am going",
     "to work at doing better next time."],
    ["SAY: Same moment, different choice. Same mark, and I keep going."],
  ],
  trap: ["thinking helpful self-talk means pretending it went well.",
         "Fix: say the honest version aloud, student re-says the flip."],
  stretch: "explain what the helpful thought changes about what you do next.",
  help: "give the stem That was hard, but I can... when words won't come.",
  prep: "Models Notice it, Name it, Flip it on the anchor. Restate the anchor phrase word for word.",
  sources: "both thoughts are the RRRR Activity 1 test example",
  tag: "[I Do | VTLM 2.0: Explicit teaching | HITS 4]",
});

const NOTES_CFU = composeGlanceNotes({
  answer: "B - it is kind, it is true, and it keeps her going",
  beats: [
    ["SAY: Two thoughts after a tough game. One helps. One does not."],
    ["ASK: Which one is positive self-talk, A or B?",
     "15 sec. Cue: Write it... chin it... show me.",
     "EXPECT: B"],
    ["SCAN the boards, back row first.",
     "80%+ -> cold call one B: how do you know? Then click to reveal.",
     "Less -> re-read A. Ask does it help or hurt? Re-ask A or B."],
  ],
  trap: ["picking A because it sounds honest.",
         "Fix: the test is kind AND true. Student re-checks B against both."],
  prep: "Decision point before guided practice. A means realistic has been confused with hopeless.",
  tag: "[CFU hinge | VTLM 2.0: Supported application | SC2 | HITS 7, 8]",
});

const NOTES_LAN = composeGlanceNotes({
  answer: "open - listen for the goal that lost the game",
  beats: [
    ["SAY: This is Lan. A lot went wrong in one day. Read it with me."],
    ["ASK: Which moment would be hardest on Lan's self-talk?",
     "10 sec. Turn and tell.",
     "EXPECT: the goal in the last minute."],
    ["SAY: We will help Lan flip her self-talk about that moment."],
  ],
  care: "keep this about Lan. If a student takes it personally, acknowledge and move on.",
  prep: "Orients the guided practice. Pick one moment so the flip stays sharp.",
  sources: "RRRR Topic 3 Activity 1, Lan's Day",
  tag: "[We Do | VTLM 2.0: Supported application | HITS 5]",
});

const NOTES_WEDO = composeGlanceNotes({
  answer: "The team will be disappointed, but there's another game next week.",
  beats: [
    ["SAY: Lan's unhelpful thought is in the top bubble. Now we flip it."],
    ["ASK: What could Lan tell herself instead?",
     "30 sec with your Helper. Cue: Write it... chin it... show me.",
     "EXPECT: kind and true, one game is not the whole story."],
    ["SCAN the boards.",
     "80%+ -> cold call two flips, reveal, then tick it or fix it.",
     "Less -> read one Helper line together, re-ask for a flip."],
  ],
  stretch: "which Helper line would work for the maths test as well? Say why.",
  help: "point to two Helper lines and have the student choose the one that fits.",
  prep: "Guided flip. The Helper is both the scaffold and the catch-up entry point.",
  sources: "answer is a line from the RRRR Positive self-talk handout",
  tag: "[We Do | VTLM 2.0: Supported application | SC3 | HITS 4, 5]",
});

const NOTES_YOUDO = composeGlanceNotes({
  answer: "open - kind, true self-talk about reading aloud tomorrow",
  beats: [
    ["SAY: Your turn, and it is a new moment. Read it from the slide."],
    ["SAY: Write the unhelpful thought on top, then flip it underneath."],
    ["CIRCULATE. Read boards. Prompt anyone stuck with the Helper."],
    ["TIME: about 5 minutes, then we share two flips."],
  ],
  stretch: "write a second helpful thought for the same moment that is different, not just longer.",
  help: "point them to the Helper line about talking in front of the school.",
  prep: ["A new moment, not Lan's, so the You Do is not the We Do again.",
         "Catch-up: the Helper alone is enough to start this task."],
  tag: "[You Do | VTLM 2.0: Mastery and application | HITS 10]",
});

const NOTES_EXIT = composeGlanceNotes({
  answer: "open - a kind, true thought that lets them keep going",
  beats: [
    ["SAY: On your own now. No talking, no Helper."],
    ["SAY: Write helpful self-talk to argue back to that thought."],
    ["COLLECT. Sort into got it and not yet as you take them."],
  ],
  prep: "Checks the second criterion and reaches the third. A new moment again, so it is not recall.",
  tag: "[Exit Ticket | VTLM 2.0: Mastery and application | SC2]",
});

const NOTES_CLOSING = composeGlanceNotes({
  answer: "open - listen for a real place they could flip their self-talk",
  beats: [
    ["SAY: Read each I can statement. Thumbs up, sideways or down.",
     "Thumbs only, voices off."],
    ["ASK: Where could positive self-talk help you this week?",
     "15 sec. Turn and tell one place.",
     "EXPECT: a real moment, a test, a game, a fall-out."],
    ["SAY: If the voice in your head is unkind a lot, tell a trusted adult."],
  ],
  care: "note any quiet thumbs-down and follow up privately rather than in front of the class.",
  prep: ["Ties back to the criteria and to help-seeking. Next session we argue back out loud.",
         "Extension: students take the Helper home and try one statement this week."],
  tag: "[Closing | VTLM 2.0: Retention and recall | HITS 9]",
});

/* ══════════════════════════════════════════════════════════════════════
   Build
   ══════════════════════════════════════════════════════════════════════ */

function build() {
  const pres = new PptxGenJS();
  pres.layout = "LAYOUT_16x9";

  /* 1. Title */
  titleSlide(
    pres,
    "Positive Coping:\nNotice Your Self-Talk",
    "Respectful Relationships",
    "Grade 5/6 Wellbeing  |  Session 1 of 4",
    NOTES_TITLE
  );

  /* 2. Teacher Resources */
  addResourceSlide(pres, RESOURCE_ITEMS, T, FOOTER, NOTES_RESOURCES);

  /* 3. Teacher-facing unit overview (section 68b) */
  {
    const s = pres.addSlide();
    addTopBar(s, C.SECONDARY);
    addBadge(s, "For the teacher", { color: C.SECONDARY, w: 2.3 });
    addTitle(s, "Positive Coping: the unit at a glance");

    const rows = [
      ["Session 1", "Notice your self-talk", "Flip an unhelpful thought"],
      ["Session 2", "Argue back", "Make the flip specific and practical"],
      ["Session 3", "Your coping profile", "Five styles of positive coping"],
      ["Session 4", "Sharing strategies", "Listen, report back, appreciate friends"],
    ];
    const rowH = 0.62;
    const startY = CONTENT_TOP + 0.02;
    rows.forEach((r, i) => {
      const y = startY + i * (rowH + 0.1);
      addCard(s, 0.5, y, 9, rowH, { strip: i < 2 ? C.PRIMARY : C.ACCENT, fill: C.WHITE });
      s.addText(r[0], {
        x: 0.75, y: y + 0.06, w: 1.35, h: rowH - 0.12,
        fontSize: 14, fontFace: FONT_B, color: C.PRIMARY, bold: true,
        valign: "middle", margin: 0,
      });
      s.addText(r[1], {
        x: 2.15, y: y + 0.06, w: 3.1, h: rowH - 0.12,
        fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, bold: true,
        valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
      s.addText(r[2], {
        x: 5.35, y: y + 0.06, w: 4.0, h: rowH - 0.12,
        fontSize: 13, fontFace: FONT_B, color: C.MUTED,
        valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
    });

    const anchorY = startY + rows.length * (rowH + 0.1) + 0.08;
    addTextOnShape(s, "Unit anchor: the flip map. Same moment, different choice. Notice it. Name it. Flip it.", {
      x: 0.5, y: anchorY, w: 9, h: 0.56, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, {
      fontSize: 14, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    T.addFooter(s, FOOTER);
    s.addNotes(NOTES_OVERVIEW);
    runSlideDiagnostics(s, pres);
  }

  /* 4. Launch */
  contentSlide(
    pres,
    "Launch",
    C.SECONDARY,
    "Same Moment, Two Thoughts",
    [
      "One dropped catch.",
      "Read both thoughts.",
      "Which one hurts more?",
    ],
    NOTES_LAUNCH,
    FOOTER,
    (s, lg) => {
      drawFlipMap(s, {
        x: lg.rightX, y: lg.panelTopPadded, w: lg.rightW,
        h: SAFE_BOTTOM - lg.panelTopPadded,
        neg: "I always drop it. I'm useless.",
        pos: "Nice try. I'll get the next one.",
      });
    }
  );

  /* 5. Learning Intention and Success Criteria */
  liSlide(
    pres,
    ["We are learning to notice our self-talk and choose positive self-talk when things go wrong"],
    [
      "I can explain what self-talk is",
      "I can tell positive self-talk from negative self-talk",
      "I can flip negative self-talk into positive self-talk",
    ],
    NOTES_LI,
    FOOTER
  );

  /* 6. Key word */
  keyWordSlide(
    pres,
    {
      word: "self-talk",
      meaning: "The talking you do inside your own head.",
      example: "Miss a catch? Your self-talk might say 'nice try' or 'you let the team down'.",
      routine: ["Say it", "Notice it", "Name it"],
      color: C.PRIMARY,
      title: "Word for today",
    },
    NOTES_VOCAB,
    FOOTER
  );

  /* 7. I Do - model the flip (click reveals the helpful thought) */
  contentSlide(
    pres,
    "I Do",
    C.PRIMARY,
    "Watch Me Flip a Thought",
    [
      "Notice it: catch it.",
      "Name it: helpful or not?",
      "Flip it: kind and true.",
    ],
    NOTES_IDO,
    FOOTER,
    (s, lg) => {
      const map = drawFlipMap(s, {
        x: lg.rightX, y: lg.panelTopPadded, w: lg.rightW,
        h: SAFE_BOTTOM - lg.panelTopPadded,
        neg: "I am no good at this. I might as well give up.",
        stagePos: "That's disappointing, but I am going to work at doing better next time.",
      });
      clickBuild(s, [() => map.revealPos()]);
    }
  );

  /* 8. CFU hinge */
  hingeSlide(pres, {
    title: "Which Is Positive Self-Talk?",
    routine: "Show Me Boards",
    optionA: "I always mess up. There is no point trying.",
    optionB: "That was tough. I will get the next one.",
    answer: "B - kind, true, and it keeps her going",
    notes: NOTES_CFU,
  });

  /* 9. We Do - meet Lan */
  scenarioSlide(
    pres,
    "We Do",
    "Meet Lan",
    LAN_SCENARIO,
    [
      "Which moment would be hardest on Lan's self-talk?",
      "What might Lan start telling herself?",
    ],
    NOTES_LAN,
    FOOTER
  );

  /* 10. We Do - flip Lan's thought */
  contentSlide(
    pres,
    "We Do",
    C.SUCCESS,
    "Flip Lan's Thought",
    [
      "Read Lan's thought.",
      "Use your Helper.",
      "Write her flip on your board.",
    ],
    NOTES_WEDO,
    FOOTER,
    (s, lg) => {
      const map = drawFlipMap(s, {
        x: lg.rightX, y: lg.panelTopPadded, w: lg.rightW,
        h: SAFE_BOTTOM - lg.panelTopPadded,
        neg: "I lost the game for everyone. I am hopeless at sport.",
        stagePos: HANDOUT_LINES[7],
      });
      clickBuild(s, [() => map.revealPos()]);
    }
  );

  /* 11. You Do - a new moment */
  contentSlide(
    pres,
    "You Do",
    C.ACCENT,
    "Flip Your Own Thought",
    [
      "Tomorrow you read your work out to the whole class.",
      "First: write the unhelpful thought on top.",
      "Then: flip it underneath.",
      "Stuck? Open your Helper.",
    ],
    NOTES_YOUDO,
    FOOTER,
    (s, lg) => {
      drawFlipMap(s, {
        x: lg.rightX, y: lg.panelTopPadded, w: lg.rightW,
        h: SAFE_BOTTOM - lg.panelTopPadded,
        negHint: "Your unhelpful thought...",
        posHint: "Your helpful self-talk...",
      });
    }
  );

  /* 12. Exit ticket */
  exitTicketSlide(
    pres,
    [
      "You forgot your lines in the class play. The thought in your head is: " +
      "'Everyone saw me mess up. I will never live it down.'",
      "Write helpful self-talk to argue back.",
    ],
    NOTES_EXIT,
    FOOTER,
    { assessesSc: 2 }
  );

  /* 13. Closing */
  closingSlide(
    pres,
    {
      reflectionPrompt: "Where could positive self-talk help you this week?",
      scItems: [
        "I can explain what self-talk is",
        "I can tell positive self-talk from negative self-talk",
        "I can flip negative self-talk into positive self-talk",
      ],
      selfAssessment: {
        prompt: "Thumbs up, sideways or down for each I can statement.",
        options: ["Got it", "Getting there", "Need practice"],
      },
      takeaways: ["Same moment, different choice. Notice it. Name it. Flip it."],
    },
    NOTES_CLOSING
  );

  return pres;
}

/* ══════════════════════════════════════════════════════════════════════
   Companion PDF: Positive Self-Talk Helper
   ══════════════════════════════════════════════════════════════════════ */

function buildHelper() {
  const doc = createPdf({ title: "Positive Self-Talk Helper" });
  let y = addPdfHeader(doc, "Session 1 Positive Self-Talk Helper", {
    subtitle: "Kind, true self-talk you can borrow",
    color: C.PRIMARY,
    lessonInfo: "Grade 5/6 Wellbeing - Respectful Relationships",
  });

  y = addTipBox(
    doc,
    "Self-talk is the talking you do inside your own head. When it turns unhelpful, you can flip it to something kinder that is still true.",
    y,
    { color: C.ACCENT }
  );

  y = addSectionHeading(doc, "How to flip a thought", y, { color: C.PRIMARY });
  y = addBodyText(doc, "Notice it. Name it. Flip it.", y, { fontSize: 11 });
  y = addBodyText(doc, "Unhelpful: I am no good at this. I might as well give up.", y, { fontSize: 11 });
  y = addBodyText(doc, "Helpful: That's disappointing, but I am going to work at doing better next time.", y, { fontSize: 11 });

  y = addSectionHeading(doc, "Positive self-talk you can borrow", y, { color: C.SUCCESS });
  y = addBodyText(doc, HANDOUT_LINES.map((l) => "- " + l).join("\n"), y, { fontSize: 11 });

  y = addSectionHeading(doc, "Your flip", y, { color: C.ACCENT });
  y = addBodyText(doc, "An unhelpful thought I sometimes have:", y, { fontSize: 11 });
  y = addLinedArea(doc, y, 1, { lineSpacing: 24 });
  y = addBodyText(doc, "Flip it to helpful self-talk (kind and true):", y, { fontSize: 11 });
  y = addLinedArea(doc, y, 2, { lineSpacing: 24 });

  addPdfFooter(doc, "Respectful Relationships | Grade 5/6 | Session 1");
  return writePdf(doc, path.join(LESSON_FOLDER, HELPER.fileName));
}

/* ══════════════════════════════════════════════════════════════════════ */

(async function main() {
  const pres = build();
  const out = path.join(LESSON_FOLDER, PPTX_NAME);
  await pres.writeFile({ fileName: out });
  console.log("PPTX written to " + out);

  await buildHelper();
  console.log("PDF written: " + HELPER.fileName);

  console.log("Session 1 build complete.");
})().catch((err) => { console.error(err); process.exit(1); });
