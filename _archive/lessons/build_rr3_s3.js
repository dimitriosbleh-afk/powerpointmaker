"use strict";

// Respectful Relationships (Resilience, Rights and Respectful Relationships)
// Topic 3 "Positive coping", Level 5-6 - Session 3 of 4
// Source activity 3: Personal coping profiles
//
// Source-locked from the supplied RRRR Topic 3 PDF:
//   - the definition of a coping strategy
//   - the five coping styles and their example lists
//   - "as long as they are not used excessively and match the situation",
//     with the shower and basketball examples
//   - the four negative coping strategies
//   - the profile task: 20 strategies, at least one from every style

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
  addCard, addTextOnShape, addChipRow, addRevealAnswerBar,
  addTopBar, addBadge, addTitle,
  clickBuild, runSlideDiagnostics,
  composeGlanceNotes,
  CONTENT_TOP, SAFE_BOTTOM,
} = T;

const UNIT_FOLDER = "RR3_Session3_Your_Coping_Profile";
const LESSON_FOLDER = path.join(__dirname, "..", "output", UNIT_FOLDER);
const PPTX_NAME = "Session 3 Your Coping Profile.pptx";
const SESSION = 3;
const FOOTER = "Respectful Relationships | Grade 5/6 | Session 3";
const RES_DIR = path.join(LESSON_FOLDER, getSessionResourceFolder(SESSION));

const STYLES_SHEET = makeSessionResource(
  SESSION,
  "Coping Styles and Strategies",
  "The five coping styles with the RRRR examples. Students keep this beside them while they build their profile."
);
const PROFILE = makeSessionResource(
  SESSION,
  "My Positive Coping Profile",
  "Twenty strategies across the five styles. Students keep this and bring it back for Session 4."
);
const RESOURCE_ITEMS = [STYLES_SHEET, PROFILE];

fs.mkdirSync(RES_DIR, { recursive: true });

/* ══════════════════════════════════════════════════════════════════════
   Unit anchor (megaprompt section 79)

   Same drawing, same phrase, same three steps as Sessions 1 and 2. From
   this session the bubbles hold what you DO rather than what you think,
   and the I Do says that out loud so the anchor carries across.
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
   Source content (locked, megaprompt section 5a)
   ══════════════════════════════════════════════════════════════════════ */

// The five coping styles and the RRRR example lists, trimmed for the slide
// face; the full source wording goes on the printed handout.
const STYLES = [
  { name: "Energetic activity", eg: "exercise, sport, dance, active play", color: "SECONDARY" },
  { name: "Self-calming activity", eg: "drawing, a quiet space, soothing music, a shower", color: "PRIMARY" },
  { name: "Social activity", eg: "talking it over, asking for help, playing with others", color: "SUCCESS" },
  { name: "Shifting attention", eg: "reading, watching TV, games, a favourite hobby", color: "ACCENT" },
  { name: "Getting organised", eg: "making lists, tidying up, making a game plan", color: "SECONDARY" },
];

// Full source example lists, for the printed handout.
const STYLES_FULL = [
  ["Energetic Activity",
   "Includes things like exercise, sport, dance, active play which can help lift your mood."],
  ["Self-Calming Activity",
   "Includes things that calm you down, like drawing, meditation, praying, being in a quiet space, listening to soothing music, taking a shower, snuggling in bed, walking the dog, stroking the cat, cuddling a teddy."],
  ["Social Activity",
   "Includes things like talking things over, help-seeking or connecting with or spending time with others."],
  ["Shifting Attention",
   "Includes things that take your mind to a different place like reading, watching TV, playing games, doing a favourite hobby."],
  ["Getting Organised",
   "Includes activities that help you plan, get organised, like making lists, tidying up, making a plan, organising an activity, making a game plan."],
];

// The RRRR negative coping list.
const NEGATIVE_COPING = [
  "Taking it out on others",
  "Using aggression or violence",
  "Alcohol or other drugs",
  "Blaming yourself",
];

/* ══════════════════════════════════════════════════════════════════════
   Teacher notes (Glance Format v12.3)
   ══════════════════════════════════════════════════════════════════════ */

const NOTES_TITLE =
  "Display as students arrive. Session 3 of 4 in Positive Coping. " +
  "Students need a pen and somewhere to keep their profile for next session.";

const NOTES_RESOURCES = [
  "Materials: Coping Styles and Strategies sheet and My Positive Coping Profile (one each), pens, board space for class examples.",
  "Prep: have your own five strategies ready to model, one from each style, and make them real ones.",
  "Decision points: the hinge after I Do, the sorting check in We Do, the exit ticket.",
  "CATCH-UP NOTE: missed a session? Nothing here depends on Sessions 1 or 2. The launch introduces the flip on actions from scratch.",
  "IMPORTANT: students keep their profile. They present it to a partner in Session 4, so collect them in or remind them to bring them back.",
  "CARE: the negative coping slide names alcohol and other drugs, from the source. Keep it brisk and factual, and follow up privately if anything is disclosed.",
  "SOURCES: RRRR Level 5-6, Topic 3 Positive coping, Activity 3 and the Coping styles and strategies handout.",
].join("\n");

const NOTES_LAUNCH = composeGlanceNotes({
  answer: "open - anything that helps and does not hurt anyone, e.g. shoot hoops",
  beats: [
    ["SAY: Same moment, different choice. Today we flip what we DO."],
    ["POINT to the top bubble.",
     "SAY: That is a real reaction. It just makes it worse."],
    ["ASK: What could this person do instead?",
     "20 sec. Cue: Write it... chin it... show me.",
     "EXPECT: something that helps, a walk, music, telling someone."],
    ["COLLECT three boards. Read them out. Do not sort them yet."],
  ],
  prep: ["Low-coupling launch: the flip map is rebuilt from scratch, so a returning student is not behind.",
         "Whole block under 4 minutes."],
  tag: "[Launch | VTLM 2.0: Retention and recall | HITS 6]",
});

const NOTES_LI = composeGlanceNotes({
  beats: [
    ["SAY: Read the learning intention with me."],
    ["POINT to each I can statement as you read it aloud."],
    ["SAY: By the end you each have your own profile to keep."],
  ],
  prep: "The profile is the product of the session and the input for Session 4.",
  tag: "[LI/SC | VTLM 2.0: Planning made visible | HITS 1]",
});

const NOTES_VOCAB = composeGlanceNotes({
  beats: [
    ["SAY: A coping strategy is something you do to help you feel better,",
     "or to keep going when things get tough."],
    ["SAY: You already use them. Today we name them and add more."],
    ["SAY: Say it with me. Everyone, together, on three.",
     "One, two, three... coping strategy."],
  ],
  prep: "Definition follows the source wording closely. Everyone can name one, which is the floor criterion.",
  sources: "RRRR Topic 3 Activity 3",
  tag: "[Vocabulary | VTLM 2.0: Knowledge and memory | HITS 3]",
});

const NOTES_STYLES = composeGlanceNotes({
  answer: "open - listen for a real strategy landing in the right style",
  beats: [
    ["SAY: Coping strategies come in five styles. Read them with me."],
    ["POINT to each row and give one example of your own as you go."],
    ["ASK: Which style do you use most?",
     "15 sec. Turn and tell, then hold up fingers one to five.",
     "EXPECT: a real strategy named with its style."],
    ["SAY: A wide range beats one favourite. That is the whole point."],
  ],
  trap: ["thinking one style is the best one.",
         "Fix: name a moment where that style fails, student picks another."],
  prep: "The five styles are the session's core representation. Leave this slide up during the You Do.",
  sources: "the five styles and examples are from the RRRR handout",
  tag: "[I Do | VTLM 2.0: Explicit teaching | HITS 3]",
});

const NOTES_MODEL = composeGlanceNotes({
  beats: [
    ["MODEL. SAY: Here is my own profile. Watch how I fill every style."],
    ["CLICK through. SAY: Energetic, I walk the dog. Self-calming,",
     "ten minutes of music. Social, I ring my sister."],
    ["CLICK. SAY: Shifting attention, I read a chapter.",
     "Getting organised, I write tomorrow's list."],
    ["SAY: All real. None of them impressive. That is what makes them work."],
  ],
  trap: ["writing what sounds good rather than what they actually do.",
         "Fix: ask when did you last do that? Student swaps it for a real one."],
  stretch: "add a second strategy to the style you use least.",
  help: "start them with the style they already named in the launch.",
  prep: "Models the profile before students build one. Keep your examples ordinary and honest.",
  tag: "[I Do | VTLM 2.0: Explicit teaching | HITS 4]",
});

const NOTES_CFU = composeGlanceNotes({
  answer: "A - getting organised. It makes a plan, it does not escape",
  beats: [
    ["SAY: One strategy, two possible styles. Only one fits."],
    ["ASK: Which style is writing a list of everything due, A or B?",
     "15 sec. Cue: Write it... chin it... show me.",
     "EXPECT: A"],
    ["SCAN the boards, back row first.",
     "80%+ -> cold call one A: how do you know? Then reveal.",
     "Less -> re-read the two styles aloud, re-ask with a new example."],
  ],
  trap: ["reading any calm activity as shifting attention.",
         "Fix: ask does it deal with the problem or step away? Student re-sorts."],
  prep: "Decision point before the class sort. B means the styles are being read as moods, not actions.",
  tag: "[CFU hinge | VTLM 2.0: Supported application | SC2 | HITS 7, 8]",
});

const NOTES_SORT = composeGlanceNotes({
  answer: "energetic, getting organised, social, self-calming - in that order",
  beats: [
    ["SAY: Four strategies. Sort each one into its style."],
    ["ASK: Which style does each one belong to?",
     "45 sec. Cue: Write all four... chin it... show me.",
     "EXPECT: all four correct, in order."],
    ["SCAN the boards.",
     "80%+ -> cold call one student per pair, reveal, tick it or fix it.",
     "Less -> re-read the styles slide, re-ask the two most got wrong."],
  ],
  stretch: "add a fifth strategy of your own and name its style.",
  help: "give them two styles to choose between for each one, not five.",
  prep: "Guided sorting before independent profile building. Leave the styles slide visible if needed.",
  tag: "[We Do | VTLM 2.0: Supported application | SC2 | HITS 5, 7]",
});

const NOTES_FIT = composeGlanceNotes({
  answer: "open - listen for a strategy that helps in one moment but not this one",
  beats: [
    ["SAY: Every strategy here is a good one. Both of these still fail."],
    ["SAY: A shower calms you down. Ten showers a day is not coping.",
     "Basketball lifts your mood. It will not get the assignment in."],
    ["ASK: When would going for a run be the wrong choice?",
     "20 sec. Turn and tell.",
     "EXPECT: when the thing that is worrying you has a deadline."],
  ],
  trap: ["hearing this as some strategies are bad.",
         "Fix: re-say it as too much, or wrong moment. Student gives an example."],
  prep: "The judgement half of the source teaching: not excessive, and matched to the situation.",
  sources: "shower and basketball examples are from RRRR Activity 3",
  tag: "[We Do | VTLM 2.0: Supported application | HITS 3]",
});

const NOTES_NEGATIVE = composeGlanceNotes({
  answer: "none of them - every one leaves the problem there and adds a new one",
  beats: [
    ["SAY: Some things people do when it gets tough make it worse."],
    ["POINT along the four. SAY: These do not belong in anyone's profile."],
    ["ASK: What do all four have in common?",
     "15 sec. Turn and tell.",
     "EXPECT: they hurt someone, including you, and fix nothing."],
  ],
  care: "keep it brisk and factual. If a student discloses anything, follow up privately.",
  prep: "Names what to cut, straight from the source list, before students write their profile.",
  sources: "the four are the RRRR Activity 3 negative coping list",
  tag: "[We Do | VTLM 2.0: Knowledge and memory | HITS 3]",
});

const NOTES_YOUDO = composeGlanceNotes({
  answer: "open - 20 strategies, at least one from every style",
  beats: [
    ["SAY: Your own profile now. Twenty strategies, four in every style."],
    ["SAY: Start with the ones you already do. Then stretch."],
    ["CIRCULATE. Ask two students to read you one from their emptiest style."],
    ["TIME: about 12 minutes. Then circle three to try this week."],
  ],
  stretch: "write when each of your three would work best, and when it would not.",
  help: "hand them the styles sheet and set a target of two per style, not four.",
  prep: ["Independent build. Different task from the class sort, which used my examples.",
         "Catch-up: the styles sheet carries everything needed to start."],
  tag: "[You Do | VTLM 2.0: Mastery and application | HITS 10]",
});

const NOTES_EXIT = composeGlanceNotes({
  answer: "open - a real strategy, correctly styled, with what it does",
  beats: [
    ["SAY: On your own. Profile face down."],
    ["SAY: One strategy, its style, and what it does for you."],
    ["COLLECT. Sort into got it and not yet as you take them."],
  ],
  prep: "Checks the sorting criterion with a strategy they chose, not one I gave them.",
  tag: "[Exit Ticket | VTLM 2.0: Mastery and application | SC2]",
});

const NOTES_CLOSING = composeGlanceNotes({
  answer: "open - listen for an honest gap and a realistic thing to try",
  beats: [
    ["SAY: Read each I can statement. Thumbs up, sideways or down.",
     "Thumbs only, voices off."],
    ["ASK: Which style is emptiest on your profile?",
     "15 sec. Turn and tell one thing you could add.",
     "EXPECT: a named style and a realistic strategy."],
    ["SAY: Bring your profile back next session. You will share it."],
  ],
  care: "a student with an almost empty profile may have little support. Follow up quietly.",
  prep: "Sets up Session 4, where partners present each other's profiles.",
  tag: "[Closing | VTLM 2.0: Retention and recall | HITS 9]",
});

/* ══════════════════════════════════════════════════════════════════════
   Build
   ══════════════════════════════════════════════════════════════════════ */

function styleColor(key) {
  return C[key] || C.PRIMARY;
}

function build() {
  const pres = new PptxGenJS();
  pres.layout = "LAYOUT_16x9";

  /* 1. Title */
  titleSlide(
    pres,
    "Positive Coping:\nYour Coping Profile",
    "Respectful Relationships",
    "Grade 5/6 Wellbeing  |  Session 3 of 4",
    NOTES_TITLE
  );

  /* 2. Teacher Resources */
  addResourceSlide(pres, RESOURCE_ITEMS, T, FOOTER, NOTES_RESOURCES);

  /* 3. Launch - the same flip map, now on what we do */
  contentSlide(
    pres,
    "Launch",
    C.SECONDARY,
    "Flip What You Do",
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
        neg: "I slam my book shut and snap at whoever talks to me.",
        posHint: "What could they do instead...",
      });
    }
  );

  /* 4. Learning Intention and Success Criteria */
  liSlide(
    pres,
    ["We are learning to sort coping strategies into styles and build our own positive coping profile"],
    [
      "I can name a positive coping strategy I already use",
      "I can sort coping strategies into the five styles",
      "I can build a profile with a strategy from every style",
    ],
    NOTES_LI,
    FOOTER
  );

  /* 5. Key word */
  keyWordSlide(
    pres,
    {
      word: "coping strategy",
      meaning: "Something you do to feel better, or to keep going when things get tough.",
      example: "Walking the dog, ringing a friend, writing a list, shooting hoops.",
      routine: ["Say it", "Name yours", "Add one"],
      color: C.PRIMARY,
      title: "Words for today",
    },
    NOTES_VOCAB,
    FOOTER
  );

  /* 6. I Do - the five coping styles */
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "I Do", { color: C.PRIMARY });
    addTitle(s, "Five Styles of Coping");

    const gap = 0.08;
    const rowH = (SAFE_BOTTOM - CONTENT_TOP - gap * (STYLES.length - 1)) / STYLES.length;
    STYLES.forEach((st, i) => {
      const y = CONTENT_TOP + i * (rowH + gap);
      const col = styleColor(st.color);
      addCard(s, 0.5, y, 9, rowH, { strip: col, fill: C.WHITE });
      s.addText(st.name, {
        x: 0.78, y: y + 0.06, w: 2.85, h: rowH - 0.12,
        fontSize: 15, fontFace: FONT_B, color: col, bold: true,
        valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
      s.addText(st.eg, {
        x: 3.75, y: y + 0.06, w: 5.55, h: rowH - 0.12,
        fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL,
        valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
    });

    T.addFooter(s, FOOTER);
    s.addNotes(NOTES_STYLES);
    runSlideDiagnostics(s, pres);
  }

  /* 7. I Do - model my own profile, one style per click */
  contentSlide(
    pres,
    "I Do",
    C.PRIMARY,
    "Watch Me Build Mine",
    [
      "One from every style.",
      "Things I really do.",
      "Nothing impressive.",
    ],
    NOTES_MODEL,
    FOOTER,
    (s, lg) => {
      const mine = [
        ["Energetic", "Walk the dog after school"],
        ["Self-calming", "Ten minutes of music"],
        ["Social", "Ring my sister"],
        ["Shifting attention", "Read a chapter"],
        ["Getting organised", "Write tomorrow's list"],
      ];
      const gap = 0.09;
      const rowH = (SAFE_BOTTOM - lg.panelTopPadded - gap * (mine.length - 1)) / mine.length;
      clickBuild(s, mine.map((m, i) => () => {
        const y = lg.panelTopPadded + i * (rowH + gap);
        addCard(s, lg.rightX, y, lg.rightW, rowH, { strip: C.SUCCESS, fill: C.WHITE });
        s.addText(m[0], {
          x: lg.rightX + 0.18, y: y + 0.05, w: lg.rightW - 0.36, h: 0.24,
          fontSize: 11, fontFace: FONT_B, color: C.SUCCESS, bold: true, margin: 0,
        });
        s.addText(m[1], {
          x: lg.rightX + 0.18, y: y + 0.29, w: lg.rightW - 0.36, h: rowH - 0.36,
          fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL,
          valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
        });
      }));
    }
  );

  /* 8. CFU hinge */
  hingeSlide(pres, {
    title: "Writing A List Of Everything Due",
    routine: "Show Me Boards",
    optionA: "Getting organised",
    optionB: "Shifting attention",
    answer: "A - it makes a plan, it does not step away",
    notes: NOTES_CFU,
  });

  /* 9. We Do - sort four strategies */
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "We Do", { color: C.SUCCESS });
    addTitle(s, "Sort Them Into Styles");

    const items = [
      "Kick the footy at lunch",
      "Write a list of what is due",
      "Ring my nan",
      "Draw in my sketchbook",
    ];
    const gap = 0.12;
    const cardW = (9 - gap * 3) / 4;
    items.forEach((item, i) => {
      const x = 0.5 + i * (cardW + gap);
      addCard(s, x, CONTENT_TOP + 0.30, cardW, 1.85, { strip: C.PRIMARY, fill: C.WHITE });
      s.addText(item, {
        x: x + 0.16, y: CONTENT_TOP + 0.42, w: cardW - 0.32, h: 1.6,
        fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL,
        align: "center", valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
    });

    s.addText("Write the style for each one, in order.", {
      x: 0.5, y: CONTENT_TOP + 2.28, w: 9, h: 0.34,
      fontSize: 16, fontFace: FONT_B, color: C.PRIMARY, bold: true,
      align: "center", margin: 0,
    });

    clickBuild(s, [
      () => {
        addRevealAnswerBar(
          s,
          ["Energetic", "Getting organised", "Social", "Self-calming"],
          { y: 4.30, h: 0.66, fontSize: 14, color: C.SUCCESS }
        );
      },
    ]);

    T.addFooter(s, FOOTER);
    s.addNotes(NOTES_SORT);
    runSlideDiagnostics(s, pres);
  }

  /* 10. We Do - a good strategy still has to fit */
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "We Do", { color: C.SUCCESS });
    addTitle(s, "Good Strategy, Wrong Moment");

    const pairs = [
      ["A shower calms you down.", "Ten showers a day is not coping."],
      ["Basketball lifts your mood.", "It will not get your assignment in."],
    ];
    const cardW = 4.4;
    pairs.forEach((p, i) => {
      const x = 0.5 + i * (cardW + 0.2);
      addCard(s, x, CONTENT_TOP, cardW, 2.1, { strip: C.SUCCESS, fill: C.WHITE });
      s.addText(p[0], {
        x: x + 0.2, y: CONTENT_TOP + 0.18, w: cardW - 0.4, h: 0.7,
        fontSize: 17, fontFace: FONT_B, color: C.SUCCESS, bold: true,
        valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
      s.addShape("rect", {
        x: x + 0.2, y: CONTENT_TOP + 0.94, w: cardW - 0.4, h: 0.02,
        fill: { color: C.MUTED },
      });
      s.addText(p[1], {
        x: x + 0.2, y: CONTENT_TOP + 1.06, w: cardW - 0.4, h: 0.86,
        fontSize: 17, fontFace: FONT_B, color: C.ALERT,
        valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
    });

    addTextOnShape(s, "Not too much of it, and it has to match the moment.", {
      x: 0.5, y: CONTENT_TOP + 2.32, w: 9, h: 0.6, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, {
      fontSize: 19, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    s.addText("Turn and tell: when would going for a run be the wrong choice?", {
      x: 0.5, y: CONTENT_TOP + 3.06, w: 9, h: 0.36,
      fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL,
      align: "center", margin: 0, fit: "shrink", shrinkText: true,
    });

    T.addFooter(s, FOOTER);
    s.addNotes(NOTES_FIT);
    runSlideDiagnostics(s, pres);
  }

  /* 11. What makes things worse */
  {
    const s = pres.addSlide();
    addTopBar(s, C.ALERT);
    addBadge(s, "Careful", { color: C.ALERT });
    addTitle(s, "These Make It Worse", { color: C.ALERT });

    addChipRow(s, 0.5, CONTENT_TOP + 0.42, 9, NEGATIVE_COPING, {
      chipH: 1.5, fontSize: 14, borderColor: C.ALERT, textColor: C.CHARCOAL,
    });

    addTextOnShape(s, "These leave the problem there and add a new one.", {
      x: 0.5, y: CONTENT_TOP + 2.24, w: 9, h: 0.66, rectRadius: 0.08,
      fill: { color: C.ALERT },
    }, {
      fontSize: 20, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    s.addText("None of these belong in your coping profile.", {
      x: 0.5, y: CONTENT_TOP + 3.04, w: 9, h: 0.4,
      fontSize: 17, fontFace: FONT_B, color: C.CHARCOAL,
      align: "center", margin: 0, fit: "shrink", shrinkText: true,
    });

    T.addFooter(s, FOOTER);
    s.addNotes(NOTES_NEGATIVE);
    runSlideDiagnostics(s, pres);
  }

  /* 12. You Do - build the profile */
  contentSlide(
    pres,
    "You Do",
    C.ACCENT,
    "Build Your Own Profile",
    [
      "First: write the ones you already do.",
      "Then: four in every style, twenty in all.",
      "Last: circle three to try this week.",
    ],
    NOTES_YOUDO,
    FOOTER,
    (s, lg) => {
      const gap = 0.09;
      const rowH = (SAFE_BOTTOM - lg.panelTopPadded - gap * (STYLES.length - 1)) / STYLES.length;
      STYLES.forEach((st, i) => {
        const y = lg.panelTopPadded + i * (rowH + gap);
        const col = styleColor(st.color);
        addCard(s, lg.rightX, y, lg.rightW, rowH, { strip: col, fill: C.WHITE });
        s.addShape("roundRect", {
          x: lg.rightX + 0.18, y: y + (rowH - 0.28) / 2, w: 0.28, h: 0.28,
          rectRadius: 0.05, fill: { color: C.WHITE }, line: { color: col, width: 1.2 },
        });
        s.addText(st.name, {
          x: lg.rightX + 0.58, y: y + 0.05, w: lg.rightW - 0.76, h: rowH - 0.1,
          fontSize: 15, fontFace: FONT_B, color: C.CHARCOAL,
          valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
        });
      });
    }
  );

  /* 13. Exit ticket */
  exitTicketSlide(
    pres,
    [
      "Write one coping strategy you will really use this week.",
      "Name its style, and say what it does for you.",
    ],
    NOTES_EXIT,
    FOOTER,
    { assessesSc: 2 }
  );

  /* 14. Closing */
  closingSlide(
    pres,
    {
      reflectionPrompt: "Which style is emptiest on your profile? What could you add?",
      scItems: [
        "I can name a positive coping strategy I already use",
        "I can sort coping strategies into the five styles",
        "I can build a profile with a strategy from every style",
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
   Companion PDF 1: Coping Styles and Strategies
   ══════════════════════════════════════════════════════════════════════ */

function buildStylesSheet() {
  const doc = createPdf({ title: "Coping Styles and Strategies" });
  let y = addPdfHeader(doc, "Session 3 Coping Styles and Strategies", {
    subtitle: "The five styles of positive coping. Keep this beside your profile.",
    color: C.PRIMARY,
    lessonInfo: "Grade 5/6 Wellbeing - Respectful Relationships",
    showNameDate: false,
  });

  y = addTipBox(
    doc,
    "A coping strategy is something you do to help you feel better, or to keep going when things get tough. Aim for a wide range across all five styles, not one favourite.",
    y,
    { color: C.ACCENT }
  );

  STYLES_FULL.forEach((st) => {
    y = addSectionHeading(doc, st[0], y, { color: C.PRIMARY, fontSize: 12 });
    y = addBodyText(doc, st[1], y, { fontSize: 11 });
    y += 2;
  });

  y = addSectionHeading(doc, "Strategies that make things worse", y, { color: C.ALERT, fontSize: 12 });
  y = addBodyText(
    doc,
    NEGATIVE_COPING.map((n) => "- " + n).join("\n") +
    "\nThese leave the problem there and add a new one. Keep them out of your profile.",
    y,
    { fontSize: 11 }
  );

  addPdfFooter(doc, "Respectful Relationships | Grade 5/6 | Session 3");
  return writePdf(doc, path.join(LESSON_FOLDER, STYLES_SHEET.fileName));
}

/* ══════════════════════════════════════════════════════════════════════
   Companion PDF 2: My Positive Coping Profile
   ══════════════════════════════════════════════════════════════════════ */

function buildProfile() {
  const doc = createPdf({ title: "My Positive Coping Profile" });
  let y = addPdfHeader(doc, "Session 3 My Positive Coping Profile", {
    subtitle: "Twenty strategies. At least one from every style. Bring this back next session.",
    color: C.PRIMARY,
    lessonInfo: "Grade 5/6 Wellbeing - Respectful Relationships",
  });

  y = addTipBox(
    doc,
    "Start with Section 1: strategies you already use. Everyone has some. Then fill the rest.",
    y,
    { color: C.ACCENT }
  );

  y = addSectionHeading(doc, "1. Strategies I already use", y, { color: C.SUCCESS, fontSize: 12 });
  y = addLinedArea(doc, y, 3, { lineSpacing: 22 });

  const boxColors = ["SECONDARY", "PRIMARY", "SUCCESS", "ACCENT", "SECONDARY"];
  STYLES.forEach((st, i) => {
    // Explicit break before style 3: two styles fit under the intro, three
    // fit on page 2. Left to flow, the heading strands at the page foot.
    if (i === 2) {
      doc.addPage();
      y = 50;
    }
    y = addSectionHeading(doc, `${i + 2}. ${st.name}`, y, {
      color: styleColor(boxColors[i]), fontSize: 12,
    });
    y = addBodyText(doc, st.eg, y, { fontSize: 9 });
    y = addLinedArea(doc, y, 4, { lineSpacing: 22 });
  });

  addPdfFooter(doc, "Respectful Relationships | Grade 5/6 | Session 3");
  return writePdf(doc, path.join(LESSON_FOLDER, PROFILE.fileName));
}

/* ══════════════════════════════════════════════════════════════════════ */

(async function main() {
  const pres = build();
  const out = path.join(LESSON_FOLDER, PPTX_NAME);
  await pres.writeFile({ fileName: out });
  console.log("PPTX written to " + out);

  await buildStylesSheet();
  console.log("PDF written: " + STYLES_SHEET.fileName);
  await buildProfile();
  console.log("PDF written: " + PROFILE.fileName);

  console.log("Session 3 build complete.");
})().catch((err) => { console.error(err); process.exit(1); });
