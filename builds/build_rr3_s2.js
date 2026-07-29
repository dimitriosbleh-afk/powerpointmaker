"use strict";

// Respectful Relationships (Resilience, Rights and Respectful Relationships)
// Topic 3 "Positive coping", Level 5-6 - Session 2 of 4
// Source activity 2: Enacting positive and negative self-talk
//
// Source-locked from the supplied RRRR Topic 3 PDF:
//   - the trio role play (body, negative self-talk, positive self-talk)
//   - the one minute limit and "positive self-talk must win the argument"
//   - the four things negative self-talk does
//   - "specific and practical", with the failed-assignment example
//   - the coach inside your own head
//   - all ten self-talk scenario cards

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
  PAGE, hex,
} = require("../themes/pdf_helpers");

const T = createTheme("wellbeing", "grade56", weekToVariant(1));
const {
  C, FONT_H, FONT_B,
  titleSlide, liSlide, contentSlide, closingSlide,
  keyWordSlide, exitTicketSlide, scenarioSlide,
  addCard, addTextOnShape, addChipRow, addRevealAnswerBar,
  addTopBar, addBadge, addTitle,
  clickBuild, runSlideDiagnostics,
  composeGlanceNotes,
  CONTENT_TOP, SAFE_BOTTOM,
} = T;

const UNIT_FOLDER = "RR3_Session2_Argue_Back";
const LESSON_FOLDER = path.join(__dirname, "..", "output", UNIT_FOLDER);
const PPTX_NAME = "Session 2 Argue Back.pptx";
const SESSION = 2;
const FOOTER = "Respectful Relationships | Grade 5/6 | Session 2";
const RES_DIR = path.join(LESSON_FOLDER, getSessionResourceFolder(SESSION));

const CARDS = makeSessionResource(
  SESSION,
  "Self-Talk Scenario Cards",
  "The ten RRRR scenario cards. Print, copy and cut up so each trio has one."
);
const PLANNER = makeSessionResource(
  SESSION,
  "Role Play Planner",
  "One page per trio: the three roles, what negative self-talk says, and comebacks that are specific and practical."
);
const RESOURCE_ITEMS = [CARDS, PLANNER];

fs.mkdirSync(RES_DIR, { recursive: true });

/* ══════════════════════════════════════════════════════════════════════
   Unit anchor (megaprompt section 79) - identical drawing in all sessions
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

  addCard(s, x, y, w, bubbleH, { strip: C.ALERT, fill: C.WHITE });
  s.addText(o.topLabel || "Unhelpful self-talk", {
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

  const flipY = y + bubbleH + gap;
  const flipW = 1.5;
  addTextOnShape(s, "Flip it", {
    x: x + (w - flipW) / 2, y: flipY, w: flipW, h: flipH, rectRadius: 0.08,
    fill: { color: C.PRIMARY },
  }, {
    fontSize: 14, fontFace: FONT_B, color: C.WHITE, bold: true,
    align: "center", valign: "middle", margin: 0,
  });

  const bY = flipY + flipH + gap;
  addCard(s, x, bY, w, bubbleH, { strip: C.SUCCESS, fill: C.WHITE });
  s.addText(o.bottomLabel || "Helpful self-talk", {
    x: x + 0.2, y: bY + 0.09, w: w - 0.4, h: 0.26,
    fontSize: 13, fontFace: FONT_B, color: C.SUCCESS, bold: true, margin: 0,
  });

  const posBox = { x: x + 0.2, y: bY + 0.38, w: w - 0.4, h: bubbleH - 0.46 };
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

  return { revealPos: () => { if (o.stagePos) writePos(o.stagePos); } };
}

/* ══════════════════════════════════════════════════════════════════════
   Two-option hinge slide (unit layout)
   ══════════════════════════════════════════════════════════════════════ */

const HINGE_CARD_Y = CONTENT_TOP + 0.58;
const HINGE_CARD_H = 2.15;
const HINGE_BAR_Y = HINGE_CARD_Y + HINGE_CARD_H + 0.19;

function hingeSlide(pres, o) {
  const s = pres.addSlide();
  addTopBar(s, C.ALERT);
  addBadge(s, "Hinge", { color: C.ALERT });
  addTitle(s, o.title, { color: C.ALERT });

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
    s.addText(String(entry[1]), {
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

// The ten RRRR self-talk scenario cards, verbatim.
const SCENARIOS = [
  "You hurt your ankle running and will not be able to be in the end of year ballet concert.",
  "Your friend asks to meet for lunch but she does not come and you see her playing with another girl.",
  "You have trained for four weeks for cross country hoping to win. You come 11th.",
  "You overhear your best friend being invited away for the weekend by another boy in the class.",
  "You are going to have a test for your new secondary school. You are worried because you hate reading.",
  "You answer a question in class and everyone laughs at your response.",
  "You are not allowed to go out with your friends because you did not finish your homework.",
  "You fall over your feet in front of the whole assembly as you come down off the stage.",
  "Your dad says you can't go to see a film with your friends.",
  "You mis-hear the teacher's question and everyone laughs at your incorrect answer.",
];

// What negative self-talk does, from the RRRR debrief notes.
const TRICKS = [
  "Makes it bigger and longer than it is",
  "Blames all of it on you",
  "Forgets you are good at things",
  "Says it will only get worse",
];

/* ══════════════════════════════════════════════════════════════════════
   Teacher notes (Glance Format v12.3)
   ══════════════════════════════════════════════════════════════════════ */

const NOTES_TITLE =
  "Display as students arrive. Session 2 of 4 in Positive Coping. " +
  "Clear floor space for three students to stand and act.";

const NOTES_RESOURCES = [
  "Materials: mini-whiteboards, Session 1 Positive Self-Talk Helper again, scenario cards cut up (one per trio), planner (one per trio), floor space.",
  "Prep: cut the scenario cards before the lesson. Ten cards covers a class in trios.",
  "Decision points: the hinge after I Do, the boards check after the class role play, the exit ticket.",
  "CATCH-UP NOTE: missed Session 1? The launch rebuilds the flip map from scratch. Hand them the Session 1 Helper and pair them with a confident partner in their trio.",
  "CARE: role play makes self-talk audible. Keep it about the character on the card, never about a real student in the room.",
  "SOURCES: RRRR Level 5-6, Topic 3 Positive coping, Activity 2 and the Self-talk scenarios sheet.",
].join("\n");

const NOTES_LAUNCH = composeGlanceNotes({
  answer: "open - anything kind and true, e.g. I can ask for help with this bit",
  beats: [
    ["SAY: Same moment, different choice. That is our move.",
     "Read the thought in the top bubble."],
    ["ASK: What could this person say to themselves instead?",
     "20 sec. Cue: Write it... chin it... show me.",
     "EXPECT: kind, true self-talk about the maths, not about them."],
    ["COLLECT two boards. Read them out. Do not teach yet."],
  ],
  prep: ["Low-coupling launch: a student who missed Session 1 can do this from the slide alone.",
         "Whole block under 4 minutes."],
  tag: "[Launch | VTLM 2.0: Retention and recall | HITS 6]",
});

const NOTES_LI = composeGlanceNotes({
  beats: [
    ["SAY: Read the learning intention with me."],
    ["POINT to each I can statement as you read it aloud."],
    ["SAY: Today the flip has to be convincing, not just kind."],
  ],
  prep: "Builds straight on Session 1. New today: making the comeback specific and practical.",
  tag: "[LI/SC | VTLM 2.0: Planning made visible | HITS 1]",
});

const NOTES_VOCAB = composeGlanceNotes({
  beats: [
    ["SAY: Positive self-talk is like a coach inside your own head."],
    ["SAY: A good coach does not say you are amazing and leave it.",
     "A good coach tells you the next thing to do."],
    ["SAY: Say it with me. Everyone, together, on three.",
     "One, two, three... inner coach."],
  ],
  prep: "The coach image is the RRRR metaphor and it carries the whole session.",
  sources: "RRRR Topic 3 Activity 2, the coach inside our own heads",
  tag: "[Vocabulary | VTLM 2.0: Knowledge and memory | HITS 3]",
});

const NOTES_TRICKS = composeGlanceNotes({
  answer: "all four - it does every one of them in two sentences",
  beats: [
    ["POINT to the thought. SAY: Listen to what this thought is doing."],
    ["SAY: It makes one test enormous. It blames only me.",
     "It forgets I am good at other things. It says it gets worse."],
    ["ASK: Which of the four can you hear in that thought?",
     "20 sec. Turn and tell, then two hands of fingers up.",
     "EXPECT: all four, or any of them named with a reason."],
  ],
  trap: ["hearing it as just being honest about a bad result.",
         "Fix: cover the first sentence, ask what is left that is true."],
  prep: "Naming the four moves is what makes arguing back possible later.",
  sources: "the four moves are the RRRR Activity 2 debrief notes",
  tag: "[I Do | VTLM 2.0: Explicit teaching | HITS 3, 7]",
});

const NOTES_IDO = composeGlanceNotes({
  beats: [
    ["MODEL. SAY: Watch me argue back to that thought out loud."],
    ["SAY: Vague first. It'll be fine. Hear how weak that is?",
     "It does not argue with anything."],
    ["CLICK to flip. SAY: I failed this assignment, but I may pass",
     "the next one. Specific about what happened. Practical about next."],
    ["SAY: Same moment, different choice. Make it specific and practical."],
  ],
  trap: ["cheerleading instead of arguing, such as you are amazing.",
         "Fix: ask what does that change? Student re-says it with a next step."],
  stretch: "add a second sentence naming exactly what you would do first.",
  help: "give the frame It is true that... but I can... and one worked example.",
  prep: "Restates the anchor phrase, then adds this session's test: specific and practical.",
  sources: "the comeback follows the RRRR Activity 2 hint",
  tag: "[I Do | VTLM 2.0: Explicit teaching | HITS 4]",
});

const NOTES_CFU = composeGlanceNotes({
  answer: "B - it names the real problem and the next thing to do",
  beats: [
    ["SAY: Two comebacks to the same worry. One is much stronger."],
    ["ASK: Which comeback is more convincing, A or B?",
     "15 sec. Cue: Write it... chin it... show me.",
     "EXPECT: B"],
    ["SCAN the boards, back row first.",
     "80%+ -> cold call one B: what makes it convincing? Then reveal.",
     "Less -> re-read A. Ask what would you actually do? Re-ask."],
  ],
  trap: ["choosing A because it sounds nicer and calmer.",
         "Fix: nice is not the test. Specific and practical is. Student re-checks."],
  prep: "Decision point before the role play. A means the specific and practical test has not landed.",
  tag: "[CFU hinge | VTLM 2.0: Supported application | SC3 | HITS 7, 8]",
});

const NOTES_ROLES = composeGlanceNotes({
  answer: "open - listen for positive self-talk winning with a next step",
  beats: [
    ["SAY: In your trio, one is the person and two are the voices."],
    ["SAY: The voices argue with each other, out loud, about the card."],
    ["SAY: Positive self-talk has to win. One minute, that is all."],
    ["CIRCULATE while trios read their card. Check every trio has a card."],
  ],
  care: "no student aims negative self-talk at a real classmate. It talks to the card.",
  prep: "Sets up the role play structure before any acting. Keep this slide under 2 minutes.",
  sources: "the trio structure and one minute limit are from RRRR Activity 2",
  tag: "[We Do | VTLM 2.0: Supported application | HITS 5]",
});

const NOTES_TOGETHER = composeGlanceNotes({
  answer: "open - a comeback naming what happened and what to do next",
  beats: [
    ["SAY: One together first. Read the card with me."],
    ["ASK: What is negative self-talk saying here?",
     "20 sec. Turn and tell.",
     "EXPECT: everyone thinks I am stupid, I will never speak again."],
    ["ASK: What does positive self-talk argue back?",
     "30 sec. Cue: Write it... chin it... show me.",
     "EXPECT: specific and practical, e.g. I will answer again tomorrow."],
    ["SCAN the boards.",
     "80%+ -> release to trios.",
     "Less -> model one comeback aloud, re-ask for a new one."],
  ],
  stretch: "argue back twice: once about the laughing, once about answering again.",
  help: "hand them the Session 1 Helper to pick a starting line.",
  prep: "One worked round with the class before trios go alone.",
  sources: "scenario is one of the ten RRRR self-talk cards",
  tag: "[We Do | VTLM 2.0: Supported application | SC2 | HITS 4, 5]",
});

const NOTES_YOUDO = composeGlanceNotes({
  answer: "open - the trio's own comeback, specific and practical",
  beats: [
    ["SAY: Your trio, your own card. Not the one we just did."],
    ["SAY: Plan it on the planner. Two minutes. Then you perform it."],
    ["CIRCULATE. Ask one trio: what is your next step line?"],
    ["TIME: one minute per performance. Two claps and move on."],
  ],
  stretch: "swap roles and run it again with a different winning argument.",
  help: "give them a card with an obvious next step, such as the cross country one.",
  prep: ["Different card from the class round, so the You Do is not the We Do again.",
         "Catch-up: the planner carries the whole structure on paper."],
  tag: "[You Do | VTLM 2.0: Mastery and application | HITS 5, 10]",
});

const NOTES_EXIT = composeGlanceNotes({
  answer: "open - names the real problem and one thing to do next",
  beats: [
    ["SAY: On your own. No planner, no talking."],
    ["SAY: One comeback. It has to be specific and practical."],
    ["COLLECT. Sort into got it and not yet as you take them."],
  ],
  prep: "Checks the core criterion. Vague comfort answers are the not-yet pile.",
  tag: "[Exit Ticket | VTLM 2.0: Mastery and application | SC2]",
});

const NOTES_CLOSING = composeGlanceNotes({
  answer: "open - a real high-pressure moment they can name",
  beats: [
    ["SAY: Read each I can statement. Thumbs up, sideways or down.",
     "Thumbs only, voices off."],
    ["ASK: When would someone your age most need an inner coach?",
     "15 sec. Turn and tell.",
     "EXPECT: a grand final, an exam, the first day of high school."],
    ["SAY: If your inner voice is unkind a lot, tell a trusted adult."],
  ],
  care: "watch for anyone whose card hit close to home. Check in quietly afterwards.",
  prep: ["Next session we move from thoughts to what we do. Extension: try one Helper line this week.",
         "Model your own positive self-talk aloud when you set the class a challenge."],
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
    "Positive Coping:\nArgue Back",
    "Respectful Relationships",
    "Grade 5/6 Wellbeing  |  Session 2 of 4",
    NOTES_TITLE
  );

  /* 2. Teacher Resources */
  addResourceSlide(pres, RESOURCE_ITEMS, T, FOOTER, NOTES_RESOURCES);

  /* 3. Launch - re-grounds the anchor from scratch (catch-up entry) */
  contentSlide(
    pres,
    "Launch",
    C.SECONDARY,
    "Flip It On Your Board",
    [
      "Read the thought on the right.",
      "Flip it. Kind and true.",
      "Boards up on cue.",
    ],
    NOTES_LAUNCH,
    FOOTER,
    (s, lg) => {
      drawFlipMap(s, {
        x: lg.rightX, y: lg.panelTopPadded, w: lg.rightW,
        h: SAFE_BOTTOM - lg.panelTopPadded,
        neg: "I got half of them wrong. I am just bad at maths.",
        posHint: "Your flip...",
      });
    }
  );

  /* 4. Learning Intention and Success Criteria */
  liSlide(
    pres,
    ["We are learning to argue back to negative self-talk with a comeback that is specific and practical"],
    [
      "I can name what negative self-talk is doing",
      "I can argue back with positive self-talk",
      "I can make my comeback specific and practical",
    ],
    NOTES_LI,
    FOOTER
  );

  /* 5. Key word */
  keyWordSlide(
    pres,
    {
      word: "inner coach",
      meaning: "The voice in your head that encourages you and tells you what to try next.",
      example: "A coach does not just say 'you're great'. A coach says 'now work on your turns'.",
      routine: ["Say it", "Hear it", "Use it"],
      color: C.PRIMARY,
      title: "Words for today",
    },
    NOTES_VOCAB,
    FOOTER
  );

  /* 6. I Do - what negative self-talk does (annotated model) */
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "I Do", { color: C.PRIMARY });
    addTitle(s, "Listen To What It Is Doing");

    addCard(s, 0.5, CONTENT_TOP, 9, 1.9, { strip: C.ALERT, fill: C.WHITE });
    s.addText("Unhelpful self-talk", {
      x: 0.75, y: CONTENT_TOP + 0.12, w: 8.5, h: 0.28,
      fontSize: 13, fontFace: FONT_B, color: C.ALERT, bold: true, margin: 0,
    });
    s.addText("I bombed that test. I am hopeless at maths, and I always will be.", {
      x: 0.75, y: CONTENT_TOP + 0.48, w: 8.5, h: 1.24,
      fontSize: 26, fontFace: FONT_B, color: C.CHARCOAL,
      valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
    });

    s.addText("What it just did", {
      x: 0.5, y: CONTENT_TOP + 2.0, w: 9, h: 0.28,
      fontSize: 14, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    addChipRow(s, 0.5, CONTENT_TOP + 2.32, 9, TRICKS, {
      chipH: 0.95, fontSize: 14, borderColor: C.ALERT, textColor: C.CHARCOAL,
    });

    T.addFooter(s, FOOTER);
    s.addNotes(NOTES_TRICKS);
    runSlideDiagnostics(s, pres);
  }

  /* 7. I Do - argue back, specific and practical */
  contentSlide(
    pres,
    "I Do",
    C.PRIMARY,
    "Watch Me Argue Back",
    [
      "Say it back, out loud.",
      "Be specific.",
      "Be practical.",
    ],
    NOTES_IDO,
    FOOTER,
    (s, lg) => {
      const map = drawFlipMap(s, {
        x: lg.rightX, y: lg.panelTopPadded, w: lg.rightW,
        h: SAFE_BOTTOM - lg.panelTopPadded,
        neg: "I failed that assignment. I always fail.",
        stagePos: "I failed this assignment, but I may pass the next one.",
      });
      clickBuild(s, [() => map.revealPos()]);
    }
  );

  /* 8. CFU hinge */
  hingeSlide(pres, {
    title: "Which Comeback Is More Convincing?",
    routine: "Show Me Boards",
    optionA: "It'll be fine. Don't worry about it.",
    optionB: "I got the flip turns wrong. I'll practise them at training.",
    answer: "B - it names the problem and the next step",
    notes: NOTES_CFU,
  });

  /* 9. We Do - the three roles */
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "We Do", { color: C.SUCCESS });
    addTitle(s, "Three Roles, One Minute");

    const roles = [
      ["The person", "Stands in the middle. Shows what happened."],
      ["Unhelpful self-talk", "Stands on one side. Argues you cannot."],
      ["Helpful self-talk", "Stands on the other. Argues back."],
    ];
    const cardW = 2.93;
    const gap = 0.105;
    roles.forEach((r, i) => {
      const x = 0.5 + i * (cardW + gap);
      const strip = i === 1 ? C.ALERT : (i === 2 ? C.SUCCESS : C.PRIMARY);
      addCard(s, x, CONTENT_TOP, cardW, 2.4, { strip, fill: C.WHITE });
      s.addShape("roundRect", {
        x: x + (cardW - 0.56) / 2, y: CONTENT_TOP + 0.22, w: 0.56, h: 0.56,
        rectRadius: 0.28, fill: { color: strip },
      });
      s.addText(String(i + 1), {
        x: x + (cardW - 0.56) / 2, y: CONTENT_TOP + 0.22, w: 0.56, h: 0.56,
        fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      s.addText(r[0], {
        x: x + 0.16, y: CONTENT_TOP + 0.92, w: cardW - 0.32, h: 0.44,
        fontSize: 17, fontFace: FONT_B, color: strip, bold: true,
        align: "center", valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
      s.addText(r[1], {
        x: x + 0.16, y: CONTENT_TOP + 1.40, w: cardW - 0.32, h: 0.86,
        fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL,
        align: "center", valign: "top", margin: 0, fit: "shrink", shrinkText: true,
      });
    });

    addTextOnShape(s, "Helpful self-talk has to win the argument.", {
      x: 0.5, y: CONTENT_TOP + 2.62, w: 9, h: 0.62, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, {
      fontSize: 19, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    T.addFooter(s, FOOTER);
    s.addNotes(NOTES_ROLES);
    runSlideDiagnostics(s, pres);
  }

  /* 10. We Do - one round together */
  scenarioSlide(
    pres,
    "We Do",
    "Try One Together",
    SCENARIOS[5],
    [
      "What is unhelpful self-talk saying?",
      "What does helpful self-talk argue back?",
    ],
    NOTES_TOGETHER,
    FOOTER
  );

  /* 11. You Do - trios, their own card */
  contentSlide(
    pres,
    "You Do",
    C.ACCENT,
    "Your Trio, Your Card",
    [
      "First: read your card together.",
      "Then: plan on the planner.",
      "Last: perform it. One minute.",
    ],
    NOTES_YOUDO,
    FOOTER,
    (s, lg) => {
      drawFlipMap(s, {
        x: lg.rightX, y: lg.panelTopPadded, w: lg.rightW,
        h: SAFE_BOTTOM - lg.panelTopPadded,
        negHint: "What your voice one says...",
        posHint: "The comeback that wins: specific and practical...",
      });
    }
  );

  /* 12. Exit ticket */
  exitTicketSlide(
    pres,
    [
      "A friend says: 'I came last in the cross country. I am hopeless at everything.'",
      "Write one comeback that is specific and practical.",
    ],
    NOTES_EXIT,
    FOOTER,
    { assessesSc: 2 }
  );

  /* 13. Closing */
  closingSlide(
    pres,
    {
      reflectionPrompt: "When would someone your age most need an inner coach?",
      scItems: [
        "I can name what negative self-talk is doing",
        "I can argue back with positive self-talk",
        "I can make my comeback specific and practical",
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
   Companion PDF 1: Self-Talk Scenario Cards (cut up, one per trio)
   ══════════════════════════════════════════════════════════════════════ */

function buildCards() {
  const doc = createPdf({ title: "Self-Talk Scenario Cards" });
  // No name/date line: these get cut up, so there is nothing to name.
  let y = addPdfHeader(doc, "Session 2 Self-Talk Scenario Cards", {
    subtitle: "Cut along the dashed lines. One card per trio.",
    color: C.PRIMARY,
    lessonInfo: "Grade 5/6 Wellbeing - Respectful Relationships",
    showNameDate: false,
  });

  y += 10;
  const cols = 2;
  const gap = 14;
  const cardW = (PAGE.CONTENT_W - gap) / cols;
  const cardH = 104;

  SCENARIOS.forEach((text, i) => {
    const col = i % cols;
    const row = Math.floor(i / cols);
    const x = PAGE.MARGIN + col * (cardW + gap);
    const cy = y + row * (cardH + gap);

    doc.save();
    doc.dash(3, { space: 3 });
    doc.lineWidth(0.8).strokeColor(hex(C.MUTED));
    doc.rect(x, cy, cardW, cardH).stroke();
    doc.undash();
    doc.restore();

    doc.save();
    doc.rect(x, cy, cardW, 5).fill(hex(C.PRIMARY));
    doc.restore();

    doc.fontSize(12).font("Sans").fillColor(hex(C.CHARCOAL));
    doc.text(text, x + 14, cy + 22, {
      width: cardW - 28,
      height: cardH - 34,
      lineGap: 3,
    });
  });

  addPdfFooter(doc, "Respectful Relationships | Grade 5/6 | Session 2");
  return writePdf(doc, path.join(LESSON_FOLDER, CARDS.fileName));
}

/* ══════════════════════════════════════════════════════════════════════
   Companion PDF 2: Role Play Planner (one per trio)
   ══════════════════════════════════════════════════════════════════════ */

function buildPlanner() {
  const doc = createPdf({ title: "Role Play Planner" });
  let y = addPdfHeader(doc, "Session 2 Role Play Planner", {
    subtitle: "One page per trio. Plan it, then perform it in one minute.",
    color: C.PRIMARY,
    lessonInfo: "Grade 5/6 Wellbeing - Respectful Relationships",
  });

  y = addTipBox(
    doc,
    "Helpful self-talk has to win the argument. It wins by being specific (it names what actually happened) and practical (it says what to do next).",
    y,
    { color: C.ACCENT }
  );

  y = addSectionHeading(doc, "Who is who", y, { color: C.PRIMARY });
  y = addBodyText(doc, "The person: ______________________", y, { fontSize: 11 });
  y = addBodyText(doc, "Unhelpful self-talk: ______________________", y, { fontSize: 11 });
  y = addBodyText(doc, "Helpful self-talk: ______________________", y, { fontSize: 11 });

  y = addSectionHeading(doc, "Our card says", y, { color: C.SECONDARY });
  y = addLinedArea(doc, y, 2, { lineSpacing: 24 });

  y = addSectionHeading(doc, "What unhelpful self-talk says", y, { color: C.ALERT });
  y = addBodyText(doc, "Use its four moves: make it huge, blame you, forget your strengths, say it gets worse.", y, { fontSize: 10 });
  y = addLinedArea(doc, y, 3, { lineSpacing: 24 });

  y = addSectionHeading(doc, "How helpful self-talk argues back", y, { color: C.SUCCESS });
  y = addBodyText(doc, "Worked example. Unhelpful: I failed that assignment, I always fail.", y, { fontSize: 10 });
  y = addBodyText(doc, "Helpful: I failed this assignment, but I may pass the next one. I will start it on Monday.", y, { fontSize: 10 });
  y = addLinedArea(doc, y, 3, { lineSpacing: 24 });

  y = addSectionHeading(doc, "Check before you perform", y, { color: C.PRIMARY });
  y = addBodyText(doc, "Does our comeback name what actually happened?\nDoes it say what to do next?\nDoes helpful self-talk win?", y, { fontSize: 11 });

  addPdfFooter(doc, "Respectful Relationships | Grade 5/6 | Session 2");
  return writePdf(doc, path.join(LESSON_FOLDER, PLANNER.fileName));
}

/* ══════════════════════════════════════════════════════════════════════ */

(async function main() {
  const pres = build();
  const out = path.join(LESSON_FOLDER, PPTX_NAME);
  await pres.writeFile({ fileName: out });
  console.log("PPTX written to " + out);

  await buildCards();
  console.log("PDF written: " + CARDS.fileName);
  await buildPlanner();
  console.log("PDF written: " + PLANNER.fileName);

  console.log("Session 2 build complete.");
})().catch((err) => { console.error(err); process.exit(1); });
