"use strict";

// Respectful Relationships (Resilience, Rights and Respectful Relationships)
// Topic 3 "Positive coping", Level 5-6 - Session 4 of 4
// Source activity 4: Sharing positive coping strategies
//
// Source-locked from the supplied RRRR Topic 3 PDF:
//   - the listening test framing and the pair share
//   - presenting a PARTNER's favourite strategies to the class
//   - identifying one strategy similar to your own and one different
//   - the circle round and the "I find it helpful when friends..." stem,
//     with both source examples

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

const T = createTheme("wellbeing", "grade56", weekToVariant(1));
const {
  C, FONT_H, FONT_B,
  titleSlide, liSlide, contentSlide, closingSlide,
  keyWordSlide, exitTicketSlide,
  addCard, addTextOnShape, addRevealAnswerBar,
  addTopBar, addBadge, addTitle,
  clickBuild, runSlideDiagnostics,
  composeGlanceNotes,
  CONTENT_TOP, SAFE_BOTTOM,
} = T;

const UNIT_FOLDER = "RR3_Session4_Sharing_Coping_Strategies";
const LESSON_FOLDER = path.join(__dirname, "..", "output", UNIT_FOLDER);
const PPTX_NAME = "Session 4 Sharing Coping Strategies.pptx";
const SESSION = 4;
const FOOTER = "Respectful Relationships | Grade 5/6 | Session 4";
const RES_DIR = path.join(LESSON_FOLDER, getSessionResourceFolder(SESSION));

const LISTENING_CARD = makeSessionResource(
  SESSION,
  "Partner Listening Card",
  "Where students record their partner's favourite strategies, one similar, one different, and their circle sentence."
);
const RESOURCE_ITEMS = [LISTENING_CARD];

fs.mkdirSync(RES_DIR, { recursive: true });

/* ══════════════════════════════════════════════════════════════════════
   Unit anchor (megaprompt section 79) - identical drawing all four sessions
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

// Both examples given with the circle stem in RRRR Activity 4.
const CIRCLE_EXAMPLES = [
  "send me encouraging text messages before my competition",
  "invite me to play soccer with them after a hard class",
];

/* ══════════════════════════════════════════════════════════════════════
   Teacher notes (Glance Format v12.3)
   ══════════════════════════════════════════════════════════════════════ */

const NOTES_TITLE =
  "Display as students arrive. Session 4 of 4 in Positive Coping. " +
  "Plan floor space for a whole-class circle at the end.";

const NOTES_RESOURCES = [
  "Materials: each student's Session 3 coping profile, Partner Listening Card (one each), pens, floor space for a circle.",
  "Prep: pair students yourself. Pair a student who has no profile with one who has a full one.",
  "Decision points: the hinge after I Do, listening in on report-backs, the exit ticket.",
  "CATCH-UP NOTE: no profile? Give them a Session 3 Coping Styles sheet and three minutes to jot four strategies. They can then be the listener first.",
  "CARE: a student whose profile is thin, or who names no supportive friends, is telling you something. Follow up quietly afterwards.",
  "SOURCES: RRRR Level 5-6, Topic 3 Positive coping, Activity 4.",
].join("\n");

const NOTES_LAUNCH = composeGlanceNotes({
  answer: "open - anything that helps, and often it is asking someone",
  beats: [
    ["SAY: Same moment, different choice. One last time."],
    ["ASK: What could this person do instead?",
     "20 sec. Cue: Write it... chin it... show me.",
     "EXPECT: something that helps, often talking to someone."],
    ["COLLECT two boards. SAY: Notice how many of you wrote a person."],
  ],
  prep: ["Low-coupling launch and it sets up today's idea: other people are a coping strategy.",
         "Whole block under 4 minutes."],
  tag: "[Launch | VTLM 2.0: Retention and recall | HITS 6]",
});

const NOTES_LI = composeGlanceNotes({
  beats: [
    ["SAY: Read the learning intention with me."],
    ["POINT to each I can statement as you read it aloud."],
    ["SAY: Today your job is to listen well enough to speak for someone else."],
  ],
  prep: "Last session of the unit. The listening is the skill being assessed, not the profile.",
  tag: "[LI/SC | VTLM 2.0: Planning made visible | HITS 1]",
});

const NOTES_VOCAB = composeGlanceNotes({
  beats: [
    ["SAY: Active listening means listening so well you could say it back."],
    ["SAY: Not nodding. Not waiting for your turn. Actually holding it."],
    ["SAY: Say it with me. Everyone, together, on three.",
     "One, two, three... active listening."],
  ],
  prep: "The source frames this activity as a test of listening skills. Name that up front.",
  sources: "RRRR Topic 3 Activity 4",
  tag: "[Vocabulary | VTLM 2.0: Knowledge and memory | HITS 3]",
});

const NOTES_IDO = composeGlanceNotes({
  beats: [
    ["MODEL. SAY: Watch how I report back on a partner. Three parts."],
    ["CLICK. SAY: Sam's favourites are shooting hoops and drawing."],
    ["CLICK. SAY: One is similar to mine. I draw too when I need to settle."],
    ["CLICK. SAY: One is different. I never ring anyone. Sam rings his nan."],
  ],
  trap: ["reporting on themselves instead of their partner.",
         "Fix: ask whose strategy is that? Student re-says it with the partner's name."],
  stretch: "add why you think that strategy works for your partner.",
  help: "give them the three sentence starters on the listening card to fill in.",
  prep: "Models the exact three-part report students will give. Keep your example ordinary.",
  sources: "the similar and different structure is from RRRR Activity 4",
  tag: "[I Do | VTLM 2.0: Explicit teaching | HITS 4]",
});

const NOTES_CFU = composeGlanceNotes({
  answer: "B - it is about the partner, and it names both strategies",
  beats: [
    ["SAY: Two students report back. Only one actually listened."],
    ["ASK: Which report shows good listening, A or B?",
     "15 sec. Cue: Write it... chin it... show me.",
     "EXPECT: B"],
    ["SCAN the boards, back row first.",
     "80%+ -> cold call one B: what makes it a good report? Then reveal.",
     "Less -> re-read A. Ask whose strategies are those? Re-ask."],
  ],
  trap: ["choosing A because it sounds friendly and confident.",
         "Fix: the test is whose strategies. Student re-checks both cards."],
  prep: "Decision point before the pair share. A means the task has been heard as talk about yourself.",
  tag: "[CFU hinge | VTLM 2.0: Supported application | SC2 | HITS 7, 8]",
});

const NOTES_PAIRS = composeGlanceNotes({
  answer: "open - listen for listeners writing the partner's words, not their own",
  beats: [
    ["SAY: Partner A reads their profile. Partner B only listens and writes."],
    ["TIME: two minutes. Then I call swap and B reads."],
    ["CIRCULATE. If a listener is talking about themselves, redirect once."],
    ["SAY: Eyes back in 3, 2, 1. You are about to speak for your partner."],
  ],
  care: "a thin profile may feel exposing. Sit them with someone kind, keep it moving.",
  prep: "The pair share. Everyone listens and everyone is listened to, which is why pairs are teacher-chosen.",
  sources: "pair sharing structure from RRRR Activity 4",
  tag: "[We Do | VTLM 2.0: Supported application | HITS 5]",
});

const NOTES_REPORT = composeGlanceNotes({
  answer: "open - three parts, all about the partner",
  beats: [
    ["SAY: Now you speak for your partner. Use the three frames."],
    ["SAY: Their favourites, one the same as yours, one different."],
    ["COLLECT around the room. Every student reports, no volunteers."],
    ["SAY: Thank you. You listened well enough to speak for someone else."],
  ],
  stretch: "say which of your partner's strategies you might borrow, and when.",
  help: "let them read straight from their listening card. Reading it counts.",
  prep: "Every student speaks. Keep each report to about 15 seconds and the block moves.",
  tag: "[We Do | VTLM 2.0: Supported application | SC2 | HITS 5, 7]",
});

const NOTES_CIRCLE = composeGlanceNotes({
  answer: "open - one real, specific thing a friend does that helps",
  beats: [
    ["SAY: Into the circle. New question, and this one is about you."],
    ["SAY: Finish the sentence. I find it helpful when friends..."],
    ["SAY: Around the circle, everyone once. You may pass."],
    ["SAY: Friends sharing coping strategies is itself a coping strategy."],
  ],
  stretch: "name a friend you could share one of your own strategies with this week.",
  help: "point to the two examples on the slide and let them borrow the shape.",
  prep: ["New content: how friends help ME, not the partner's profile.",
         "Honour the pass. A student who passes is still in the circle."],
  sources: "the stem and both examples are from RRRR Activity 4",
  tag: "[You Do | VTLM 2.0: Mastery and application | HITS 5, 9]",
});

const NOTES_EXIT = composeGlanceNotes({
  answer: "open - one real way a friend helps, one real thing they will do",
  beats: [
    ["SAY: On your own. No talking."],
    ["SAY: One way a friend helps you.",
     "Then one thing you will do for a friend."],
    ["COLLECT. Read them tonight. They tell you who has support."],
  ],
  care: "an exit ticket naming no one is worth a quiet follow-up this week.",
  prep: "Checks the third criterion and closes the unit with a commitment, not just a reflection.",
  tag: "[Exit Ticket | VTLM 2.0: Mastery and application | SC3]",
});

const NOTES_CLOSING = composeGlanceNotes({
  answer: "open - a named friend and a realistic strategy to share",
  beats: [
    ["SAY: Read each I can statement. Thumbs up, sideways or down.",
     "Thumbs only, voices off."],
    ["SAY: Four sessions. Same move every time. Same moment, different choice."],
    ["ASK: Who could you share a coping strategy with this week?",
     "15 sec. Turn and tell.",
     "EXPECT: a named person and something they would actually do."],
    ["SAY: If coping feels hard a lot, tell a trusted adult.",
     "That is a strategy too."],
  ],
  care: "say the help-seeking line plainly. Do not rush past it.",
  prep: ["Closes the whole unit. Keep the profiles somewhere students can find them again.",
         "Reflect: does your own language in class model positive coping talk?"],
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
    "Positive Coping:\nSharing Strategies",
    "Respectful Relationships",
    "Grade 5/6 Wellbeing  |  Session 4 of 4",
    NOTES_TITLE
  );

  /* 2. Teacher Resources */
  addResourceSlide(pres, RESOURCE_ITEMS, T, FOOTER, NOTES_RESOURCES);

  /* 3. Launch */
  contentSlide(
    pres,
    "Launch",
    C.SECONDARY,
    "Flip It One More Time",
    [
      "Same moment. Different choice.",
      "Read what this person does.",
      "What could they do instead?",
    ],
    NOTES_LAUNCH,
    FOOTER,
    (s, lg) => {
      drawFlipMap(s, {
        x: lg.rightX, y: lg.panelTopPadded, w: lg.rightW,
        h: SAFE_BOTTOM - lg.panelTopPadded,
        topLabel: "What makes it worse",
        bottomLabel: "What helps",
        neg: "I sit on my own at lunch and tell nobody anything.",
        posHint: "What could they do instead...",
      });
    }
  );

  /* 4. Learning Intention and Success Criteria */
  liSlide(
    pres,
    ["We are learning to listen carefully to a partner's coping profile and share how friends help us cope"],
    [
      "I can listen so well I could say my partner's strategies back",
      "I can present my partner's strategies to the class",
      "I can name one way friends help me cope",
    ],
    NOTES_LI,
    FOOTER
  );

  /* 5. Key word */
  keyWordSlide(
    pres,
    {
      word: "active listening",
      meaning: "Listening so carefully that you could say it back.",
      example: "Not nodding along. Not waiting for your turn. Actually holding it.",
      routine: ["Say it", "Do it", "Prove it"],
      color: C.PRIMARY,
      title: "Words for today",
    },
    NOTES_VOCAB,
    FOOTER
  );

  /* 6. I Do - the three-part report, one part per click */
  contentSlide(
    pres,
    "I Do",
    C.PRIMARY,
    "Watch Me Report Back",
    [
      "Their strategies, not mine.",
      "One the same as mine.",
      "One different.",
    ],
    NOTES_IDO,
    FOOTER,
    (s, lg) => {
      const parts = [
        ["Their favourites", "Sam's are shooting hoops and drawing."],
        ["One the same", "I draw too when I need to settle."],
        ["One different", "I never ring anyone. Sam rings his nan."],
      ];
      const gap = 0.14;
      const rowH = (SAFE_BOTTOM - lg.panelTopPadded - gap * (parts.length - 1)) / parts.length;
      clickBuild(s, parts.map((p, i) => () => {
        const y = lg.panelTopPadded + i * (rowH + gap);
        addCard(s, lg.rightX, y, lg.rightW, rowH, { strip: C.SUCCESS, fill: C.WHITE });
        s.addText(p[0], {
          x: lg.rightX + 0.18, y: y + 0.08, w: lg.rightW - 0.36, h: 0.26,
          fontSize: 12, fontFace: FONT_B, color: C.SUCCESS, bold: true, margin: 0,
        });
        s.addText(p[1], {
          x: lg.rightX + 0.18, y: y + 0.36, w: lg.rightW - 0.36, h: rowH - 0.46,
          fontSize: 15, fontFace: FONT_B, color: C.CHARCOAL,
          valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
        });
      }));
    }
  );

  /* 7. CFU hinge */
  hingeSlide(pres, {
    title: "Which Report Shows Good Listening?",
    routine: "Show Me Boards",
    optionA: "I like playing footy and listening to music.",
    optionB: "Ava's are walking her dog and making lists.",
    answer: "B - it is about your partner, not about you",
    notes: NOTES_CFU,
  });

  /* 8. We Do - the pair share */
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "We Do", { color: C.SUCCESS });
    addTitle(s, "Read, Then Listen");

    const roles = [
      ["The reader", "Reads their profile out. Names their favourites."],
      ["The listener", "Says nothing. Writes their partner's strategies down."],
    ];
    const cardW = 4.4;
    roles.forEach((r, i) => {
      const x = 0.5 + i * (cardW + 0.2);
      const strip = i === 0 ? C.PRIMARY : C.SUCCESS;
      addCard(s, x, CONTENT_TOP, cardW, 2.3, { strip, fill: C.WHITE });
      s.addShape("roundRect", {
        x: x + (cardW - 0.56) / 2, y: CONTENT_TOP + 0.24, w: 0.56, h: 0.56,
        rectRadius: 0.28, fill: { color: strip },
      });
      s.addText(String(i + 1), {
        x: x + (cardW - 0.56) / 2, y: CONTENT_TOP + 0.24, w: 0.56, h: 0.56,
        fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      s.addText(r[0], {
        x: x + 0.2, y: CONTENT_TOP + 0.96, w: cardW - 0.4, h: 0.44,
        fontSize: 19, fontFace: FONT_B, color: strip, bold: true,
        align: "center", valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
      s.addText(r[1], {
        x: x + 0.2, y: CONTENT_TOP + 1.44, w: cardW - 0.4, h: 0.72,
        fontSize: 15, fontFace: FONT_B, color: C.CHARCOAL,
        align: "center", valign: "top", margin: 0, fit: "shrink", shrinkText: true,
      });
    });

    addTextOnShape(s, "Two minutes each. Then you swap.", {
      x: 0.5, y: CONTENT_TOP + 2.52, w: 9, h: 0.62, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, {
      fontSize: 20, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    s.addText("Listen hard. You will be speaking for your partner.", {
      x: 0.5, y: CONTENT_TOP + 3.28, w: 9, h: 0.36,
      fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL,
      align: "center", margin: 0, fit: "shrink", shrinkText: true,
    });

    T.addFooter(s, FOOTER);
    s.addNotes(NOTES_PAIRS);
    runSlideDiagnostics(s, pres);
  }

  /* 9. We Do - report back frames */
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "We Do", { color: C.SUCCESS });
    addTitle(s, "Speak For Your Partner");

    const frames = [
      "Their favourite strategies are ... and ...",
      "One is similar to mine. I also ...",
      "One is different. They ... and I never do that.",
    ];
    const gap = 0.16;
    const rowH = (SAFE_BOTTOM - CONTENT_TOP - 0.3 - gap * (frames.length - 1)) / frames.length;
    frames.forEach((f, i) => {
      const y = CONTENT_TOP + 0.3 + i * (rowH + gap);
      addCard(s, 0.5, y, 9, rowH, { strip: C.SUCCESS, fill: C.WHITE });
      s.addShape("roundRect", {
        x: 0.72, y: y + (rowH - 0.44) / 2, w: 0.44, h: 0.44,
        rectRadius: 0.22, fill: { color: C.SUCCESS },
      });
      s.addText(String(i + 1), {
        x: 0.72, y: y + (rowH - 0.44) / 2, w: 0.44, h: 0.44,
        fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      s.addText(f, {
        x: 1.34, y: y + 0.08, w: 7.9, h: rowH - 0.16,
        fontSize: 20, fontFace: FONT_B, color: C.CHARCOAL,
        valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
    });

    T.addFooter(s, FOOTER);
    s.addNotes(NOTES_REPORT);
    runSlideDiagnostics(s, pres);
  }

  /* 10. You Do - the circle round */
  {
    const s = pres.addSlide();
    addTopBar(s, C.ACCENT);
    addBadge(s, "You Do", { color: C.ACCENT });
    addTitle(s, "Around The Circle");

    addTextOnShape(s, "I find it helpful when friends...", {
      x: 0.5, y: CONTENT_TOP, w: 9, h: 1.2, rectRadius: 0.1,
      fill: { color: C.PRIMARY },
    }, {
      fontSize: 34, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    s.addText("For example", {
      x: 0.5, y: CONTENT_TOP + 1.36, w: 9, h: 0.28,
      fontSize: 14, fontFace: FONT_B, color: C.PRIMARY, bold: true,
      align: "center", margin: 0,
    });

    const cardW = 4.4;
    CIRCLE_EXAMPLES.forEach((eg, i) => {
      const x = 0.5 + i * (cardW + 0.2);
      addCard(s, x, CONTENT_TOP + 1.70, cardW, 1.5, { strip: C.SUCCESS, fill: C.WHITE });
      s.addText("... " + eg, {
        x: x + 0.2, y: CONTENT_TOP + 1.82, w: cardW - 0.4, h: 1.26,
        fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL,
        valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
    });

    s.addText("Everyone once around the circle. You may pass.", {
      x: 0.5, y: CONTENT_TOP + 3.34, w: 9, h: 0.36,
      fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL,
      align: "center", margin: 0, fit: "shrink", shrinkText: true,
    });

    T.addFooter(s, FOOTER);
    s.addNotes(NOTES_CIRCLE);
    runSlideDiagnostics(s, pres);
  }

  /* 11. Exit ticket */
  exitTicketSlide(
    pres,
    [
      "Write one way a friend helps you cope.",
      "Write one thing you will do this week to help a friend cope.",
    ],
    NOTES_EXIT,
    FOOTER,
    { assessesSc: 3 }
  );

  /* 12. Closing */
  closingSlide(
    pres,
    {
      reflectionPrompt: "Who could you share a coping strategy with this week?",
      scItems: [
        "I can listen so well I could say my partner's strategies back",
        "I can present my partner's strategies to the class",
        "I can name one way friends help me cope",
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
   Companion PDF: Partner Listening Card
   ══════════════════════════════════════════════════════════════════════ */

function buildListeningCard() {
  const doc = createPdf({ title: "Partner Listening Card" });
  let y = addPdfHeader(doc, "Session 4 Partner Listening Card", {
    subtitle: "Write your partner's strategies here, then use it to speak for them.",
    color: C.PRIMARY,
    lessonInfo: "Grade 5/6 Wellbeing - Respectful Relationships",
  });

  y = addTipBox(
    doc,
    "Everything on this card is about your PARTNER, not about you. The only part about you is the last box.",
    y,
    { color: C.ACCENT }
  );

  y = addSectionHeading(doc, "My partner is", y, { color: C.PRIMARY, fontSize: 12 });
  y = addLinedArea(doc, y, 1, { lineSpacing: 26 });

  y = addSectionHeading(doc, "1. Their favourite coping strategies", y, { color: C.SUCCESS, fontSize: 12 });
  y = addBodyText(doc, "Write down what they say as they say it. Do not wait until the end.", y, { fontSize: 10 });
  y = addLinedArea(doc, y, 4, { lineSpacing: 24 });

  y = addSectionHeading(doc, "2. One that is similar to mine", y, { color: C.SECONDARY, fontSize: 12 });
  y = addBodyText(doc, "Start with: One is similar to mine. I also ...", y, { fontSize: 10 });
  y = addLinedArea(doc, y, 2, { lineSpacing: 24 });

  y = addSectionHeading(doc, "3. One that is different", y, { color: C.SECONDARY, fontSize: 12 });
  y = addBodyText(doc, "Start with: One is different. They ... and I never do that.", y, { fontSize: 10 });
  y = addLinedArea(doc, y, 2, { lineSpacing: 24 });

  y = addSectionHeading(doc, "4. My circle sentence", y, { color: C.ACCENT, fontSize: 12 });
  y = addBodyText(
    doc,
    "This one is about you. Finish the sentence: I find it helpful when friends ...\n" +
    "For example: send me encouraging text messages before my competition.",
    y,
    { fontSize: 10 }
  );
  y = addLinedArea(doc, y, 2, { lineSpacing: 24 });

  addPdfFooter(doc, "Respectful Relationships | Grade 5/6 | Session 4");
  return writePdf(doc, path.join(LESSON_FOLDER, LISTENING_CARD.fileName));
}

/* ══════════════════════════════════════════════════════════════════════ */

(async function main() {
  const pres = build();
  const out = path.join(LESSON_FOLDER, PPTX_NAME);
  await pres.writeFile({ fileName: out });
  console.log("PPTX written to " + out);

  await buildListeningCard();
  console.log("PDF written: " + LISTENING_CARD.fileName);

  console.log("Session 4 build complete.");
})().catch((err) => { console.error(err); process.exit(1); });
