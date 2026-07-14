"use strict";

/**
 * Big Feelings, Calm Choices - Year 2 Wellbeing (single lesson, first on the topic)
 *
 * Content: recognising big feelings and choosing a calming strategy taught as
 * three simple, child-friendly steps - Stop, Breathe, Ask for help. 45 minutes,
 * mixed readiness. Mini-whiteboards available. Class has a calm corner.
 *
 * General-subject explicit-teaching order (megaprompt s32 / s0a item 23):
 *   Title -> Teacher Resources -> Launch (hook) -> LI/SC -> Key Vocabulary
 *   -> Notice (recognising big feelings) -> I Do (three calm steps)
 *   -> I Do (breathe together) -> CFU (reveal) -> We Do guided scenario
 *   -> We Do partner scenario -> CFU hinge (reveal) -> You Do (My Calm Plan)
 *   -> Exit Ticket -> Closing.
 *
 * Visual anchors are BUILT (no external images): recognisable emotion-face
 * icons (react-icons/md sentiment set), a size-of-feeling scale, a numbered
 * three-step calm-steps poster, and a breathing-ball visual. All custom slides
 * run runSlideDiagnostics so overlaps / out-of-bounds / underfill fail the gate.
 *
 * Notes use the v11.0 Glance Format (composeGlanceNotes).
 *
 * Companion PDFs (resources-session1/):
 *   - Session 1 Calm Steps Poster  (display for the calm corner and desks)
 *   - Session 1 My Calm Plan        (the You Do recording sheet)
 * Both are personal/open, so no answer key; teacher look-fors live in the notes.
 */

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const {
  MdSentimentVerySatisfied,
  MdSentimentSatisfiedAlt,
  MdSentimentDissatisfied,
  MdSentimentVeryDissatisfied,
  MdMoodBad,
  MdPanTool,
  MdSelfImprovement,
  MdRecordVoiceOver,
} = require("react-icons/md");

const { createTheme, weekToVariant } = require("../themes/factory");
const {
  makeSessionResource,
  addResourceSlide,
  createPdf,
  writePdf,
  addPdfHeader,
  addPdfFooter,
  addSectionHeading,
  addBodyText,
  addTipBox,
  hex,
  PAGE,
} = require("../themes/pdf_helpers");

// First lesson on the topic -> week 1 -> variant 0. Any future lessons in this
// unit MUST reuse this variant for theme cohesion.
const T = createTheme("wellbeing", "grade2", weekToVariant(1));
const {
  C, FONT_H, FONT_B,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  keyWordSlide, scenarioSlide, exitTicketSlide,
  withReveal, addRevealAnswerBar,
  addBadge, addTitle, addTopBar, addFooter, addTextOnShape, addIconCircle,
  iconToBase64Png, runSlideDiagnostics,
  composeGlanceNotes,
} = T;

const FOOTER = "Wellbeing | Year 2 | Big Feelings, Calm Choices";
const OUT_DIR = "output/Big_Feelings_Calm_Choices";
fs.mkdirSync(OUT_DIR, { recursive: true });

/* Success criteria - plain unlabelled "I can..." list (no tier labels on any
 * student-facing surface). Internal tiers: SC1 name a big feeling (floor
 * everyone reaches), SC2 choose a calm step (exit ticket target), SC3 say when
 * they would use it (depth). */
const LI = "We are learning to notice big feelings and choose a calm step to help.";
const SC = [
  "I can name a big feeling.",
  "I can choose a calm step to help with a big feeling.",
  "I can say when I would use my calm step.",
];

/* The three calm steps - the single model reused across the lesson and both
 * PDFs so students see the same representation every time. */
const STEPS = [
  { num: "1", name: "STOP",         action: "Stop and freeze your body.", color: C.ALERT,     icon: MdPanTool },
  { num: "2", name: "BREATHE",      action: "Take three slow breaths.",    color: C.PRIMARY,   icon: MdSelfImprovement },
  { num: "3", name: "ASK FOR HELP", action: "Tell a grown-up how you feel.", color: C.ACCENT,  icon: MdRecordVoiceOver },
];

/* Resource objects (PDFs generated at the end of this script). */
const poster = makeSessionResource(
  1,
  "Calm Steps Poster",
  "Stop, Breathe, Ask for help. Print big for the calm corner and small for desks."
);
const calmPlan = makeSessionResource(
  1,
  "My Calm Plan",
  "The You Do sheet: draw a big feeling, circle a calm step, draw yourself calm."
);

/* ─────────────────────────────────────────────────────────────────────────
 *  Async main - emotion-face and step icons are rendered before building.
 * ───────────────────────────────────────────────────────────────────────── */

async function main() {
  // Pre-render all icons as white PNGs (drawn inside coloured circles).
  const ICON = {};
  const iconJobs = [
    ["happy", MdSentimentVerySatisfied],
    ["worried", MdSentimentDissatisfied],
    ["angry", MdMoodBad],
    ["sad", MdSentimentVeryDissatisfied],
    ["calm", MdSentimentSatisfiedAlt],
    ["bubbling", MdSentimentDissatisfied],
    ["big", MdMoodBad],
    ["stop", MdPanTool],
    ["breathe", MdSelfImprovement],
    ["ask", MdRecordVoiceOver],
  ];
  for (const [key, Comp] of iconJobs) {
    ICON[key] = await iconToBase64Png(Comp, "FFFFFF", 256);
  }

  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.title = "Big Feelings, Calm Choices - Year 2 Wellbeing";
  pres.author = "Year 2 Wellbeing";

  /* Local visual helper: a coloured circle with a white emotion icon and a
   * caption below. Used by the launch and the size-of-feeling scale. */
  function faceUnit(s, cx, cy, r, color, iconData, label) {
    addIconCircle(s, iconData, cx, cy, r, color);
    s.addText(String(label), {
      x: cx - 1.05, y: cy + r + 0.06, w: 2.1, h: 0.42,
      fontSize: 22, fontFace: FONT_B, color: C.CHARCOAL, bold: true,
      align: "center", valign: "top", margin: 0, fit: "shrink", shrinkText: true,
    });
  }

  /* ───────────────────────────────────────────────────────────────────────
   *  1. Title
   * ─────────────────────────────────────────────────────────────────────── */

  titleSlide(
    pres,
    "Big Feelings, Calm Choices",
    "Notice a big feeling, then choose a calm step",
    "Year 2 Wellbeing  |  Mixed readiness  |  First lesson on this topic",
    "Warm and calm start. This is the first lesson on big feelings. Keep the three calm steps slide handy to flip back to during practice."
  );

  /* ───────────────────────────────────────────────────────────────────────
   *  2. Teacher Resources (immediately after title)
   * ─────────────────────────────────────────────────────────────────────── */

  addResourceSlide(
    pres,
    {
      resources: [poster, calmPlan],
      studentTools: [
        "Mini-whiteboards and markers (launch, CFU, We Do)",
        "Pencils and crayons (My Calm Plan)",
      ],
      manipulatives: [
        "Calm corner ready to use during and after the lesson",
        "Optional: a flower and a candle picture to model the breathing cue",
      ],
      routineIcons: [
        "Show me on your whiteboard",
        "Turn and tell your partner",
        "Thumbs up, sideways, down",
      ],
      boardSetup: [
        "Print the Calm Steps Poster big for the calm corner and small for desks",
        "Print one My Calm Plan per student",
        "Have the three calm steps slide ready to flip back to during practice",
      ],
    },
    T,
    FOOTER,
    "Set up before students arrive: whiteboards out, calm corner tidy, Calm Steps Poster on display, one My Calm Plan per student. Everything students need is taught on the slides first."
  );

  /* ───────────────────────────────────────────────────────────────────────
   *  3. Launch (hook) - we all have big feelings
   * ─────────────────────────────────────────────────────────────────────── */

  (function launchSlide() {
    const s = pres.addSlide();
    addTopBar(s, C.SECONDARY);
    addBadge(s, "Launch", { color: C.SECONDARY });
    addTitle(s, "We all have big feelings");

    const faces = [
      { color: C.SUCCESS,   icon: ICON.happy,   label: "happy" },
      { color: C.ACCENT,    icon: ICON.worried, label: "worried" },
      { color: C.ALERT,     icon: ICON.angry,   label: "angry" },
      { color: C.PRIMARY,   icon: ICON.sad,     label: "sad" },
    ];
    const centers = [1.625, 3.875, 6.125, 8.375];
    faces.forEach((f, i) => faceUnit(s, centers[i], 2.45, 0.72, f.color, f.icon, f.label));

    addTextOnShape(s, "Point to how you might feel. Then turn and tell your partner.", {
      x: 0.5, y: 3.95, w: 9, h: 0.8, rectRadius: 0.1,
      fill: { color: C.SECONDARY },
    }, {
      fontSize: 22, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0.1,
    });

    addFooter(s, FOOTER);
    s.addNotes(composeGlanceNotes({
      answer: "open - every feeling is okay; listen for children naming a real big feeling",
      beats: [
        "POINT to each face. SAY: Everyone has feelings. Some feel small, and some feel really big.",
        "ASK: Point to how you might feel when something goes wrong. 5 sec, everyone points. EXPECT: a point to any face.",
        "ASK: Turn and tell your partner about a time you had a big feeling. 30 sec, turn and tell. EXPECT: a short real example. ACCEPT: a one-word feeling.",
      ],
      trap: "a child saying feelings are bad or naughty. Fix: name that all feelings are okay, it is what we do next that matters, child re-says one feeling is okay.",
      care: "some children may name an upsetting event. Keep it light, thank them, and follow up privately later if needed.",
      prep: "Bridges everyday feelings to today's target: noticing a big feeling and choosing a calm step. First lesson on the topic.",
      tag: "[Launch | Retention and recall | HITS 2, 6]",
    }));
    runSlideDiagnostics(s, pres);
  })();

  /* ───────────────────────────────────────────────────────────────────────
   *  4. Learning Intention & Success Criteria
   * ─────────────────────────────────────────────────────────────────────── */

  liSlide(
    pres,
    LI,
    SC,
    composeGlanceNotes({
      beats: [
        "SAY: Here is what we are learning today, and how you will know you have got it.",
        "POINT to each success criterion. SAY: The first one is for everyone. We will all name a big feeling together.",
      ],
      prep: "Read the LI, then each I can statement. Internal only (not shown): SC1 name a big feeling, SC2 choose a calm step (exit target), SC3 say when they would use it.",
      tag: "[LI and SC | Planning made visible | HITS 1]",
    }),
    FOOTER
  );

  /* ───────────────────────────────────────────────────────────────────────
   *  5. Key Vocabulary (after LI/SC) - one word: calm
   * ─────────────────────────────────────────────────────────────────────── */

  keyWordSlide(
    pres,
    {
      word: "calm",
      meaning: "calm means your body feels quiet, still and safe",
      example: "When I am calm, my breathing is slow and my body feels relaxed.",
      routine: ["Say it", "Show it", "Feel it"],
      color: C.PRIMARY,
      badgeText: "Key Word",
      title: "Our word for today",
    },
    composeGlanceNotes({
      answer: "calm - body feels quiet, still and safe",
      beats: [
        "SHOW the word. SAY: Our word today is calm. Say it with me, calm.",
        "SAY: Calm means your body feels quiet and still and safe, like when your breathing is slow.",
        "ASK: Show me a calm body at your spot. 5 sec, everyone shows. EXPECT: still body, slow breathing.",
      ],
      trap: "thinking calm means happy. Fix: you can feel calm even when something is hard - it is a quiet, steady body, child shows a calm body again.",
      prep: "One word only for Year 2. Calm is the goal of every calm step that follows.",
      tag: "[Key Vocabulary | Knowledge and memory | HITS 3]",
    }),
    FOOTER
  );

  /* ───────────────────────────────────────────────────────────────────────
   *  6. Notice - how big is my feeling? (recognising, the visual anchor)
   * ─────────────────────────────────────────────────────────────────────── */

  (function scaleSlide() {
    const s = pres.addSlide();
    addTopBar(s, C.SECONDARY);
    addBadge(s, "Notice", { color: C.SECONDARY });
    addTitle(s, "How big is my feeling?");

    const zones = [
      { color: C.SUCCESS, icon: ICON.calm,     r: 0.50, name: "Calm",        clue: "Body feels quiet and still" },
      { color: C.ACCENT,  icon: ICON.bubbling, r: 0.63, name: "Bubbling up",  clue: "Tummy flips, breathing speeds up" },
      { color: C.ALERT,   icon: ICON.big,      r: 0.78, name: "Big feeling",  clue: "Hot face, tight fists, want to shout" },
    ];
    const cardW = 2.9;
    const cardXs = [0.5, 3.55, 6.6];
    zones.forEach((z, i) => {
      const cx = cardXs[i] + cardW / 2;
      T.addCard(s, cardXs[i], 1.45, cardW, 3.5, { strip: z.color, fill: C.WHITE });
      addTextOnShape(s, z.name, {
        x: cardXs[i] + 0.2, y: 1.62, w: cardW - 0.4, h: 0.5, rectRadius: 0.08,
        fill: { color: z.color },
      }, {
        fontSize: 18, fontFace: FONT_B, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      addIconCircle(s, z.icon, cx, 2.95, z.r, z.color);
      s.addText(z.clue, {
        x: cardXs[i] + 0.15, y: 3.85, w: cardW - 0.3, h: 0.95,
        fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL,
        align: "center", valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
    });

    addFooter(s, FOOTER);
    s.addNotes(composeGlanceNotes({
      answer: "open - a feeling can be small, getting bigger, or big; we notice the clues in our body",
      beats: [
        "POINT along the three cards. SAY: Feelings can be small and calm, or start to bubble up, or grow really big.",
        "MODEL the body clues. SAY: When mine gets big I notice a hot face and tight fists. That is my body telling me.",
        "ASK: Put a hand on your tummy. Where do you feel a big feeling in your body? 10 sec, turn and tell. EXPECT: a body part like tummy, chest, face, hands.",
      ],
      trap: "waiting until the feeling is already huge. Fix: point to the middle card - we can notice it bubbling up early, child names one early clue.",
      stretch: "name a feeling that fits each of the three sizes.",
      help: "give the child the three faces to point to as you name each size.",
      prep: "This scale is the recognising half of the LI. Noticing early makes the calm steps easier. Keep it visual - narrate, do not read the slide to them.",
      tag: "[Notice | Attention and regulation | HITS 3, 9]",
    }));
    runSlideDiagnostics(s, pres);
  })();

  /* ───────────────────────────────────────────────────────────────────────
   *  7. I Do - the three calm steps (the model)
   * ─────────────────────────────────────────────────────────────────────── */

  (function calmStepsSlide() {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "I Do", { color: C.PRIMARY });
    addTitle(s, "Three calm steps");

    const cardW = 2.9;
    const cardXs = [0.5, 3.55, 6.6];
    STEPS.forEach((st, i) => {
      const cx = cardXs[i] + cardW / 2;
      T.addCard(s, cardXs[i], 1.5, cardW, 3.3, { strip: st.color, fill: C.WHITE });
      addTextOnShape(s, st.num, {
        x: cx - 0.3, y: 1.62, w: 0.6, h: 0.5, rectRadius: 0.25,
        fill: { color: st.color },
      }, {
        fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      addIconCircle(s, ICON[["stop", "breathe", "ask"][i]], cx, 2.85, 0.55, st.color);
      s.addText(st.name, {
        x: cardXs[i] + 0.1, y: 3.45, w: cardW - 0.2, h: 0.45,
        fontSize: 22, fontFace: FONT_H, color: st.color, bold: true,
        align: "center", valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
      s.addText(st.action, {
        x: cardXs[i] + 0.1, y: 3.95, w: cardW - 0.2, h: 0.75,
        fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL,
        align: "center", valign: "top", margin: 0, fit: "shrink", shrinkText: true,
      });
    });

    addFooter(s, FOOTER);
    s.addNotes(composeGlanceNotes({
      answer: "the three calm steps in order: 1 Stop, 2 Breathe, 3 Ask for help",
      beats: [
        "MODEL with a think-aloud. SAY: Watch me. I feel a big feeling coming, so first I STOP and freeze my body.",
        "MODEL the breath. SAY: Next I BREATHE. I take three slow breaths to help my body settle down.",
        "SAY: Then, if it is still hard, I ASK FOR HELP. I tell a grown-up how I feel. Stop, breathe, ask for help.",
        "ASK: Say the three steps with me. 5 sec, choral on my signal. EXPECT: Stop, breathe, ask for help.",
      ],
      trap: "jumping straight to shouting or asking without stopping first. Fix: point to step 1 - we always Stop first, child says step 1.",
      stretch: "tell a partner one place they could use the calm steps today.",
      help: "give the Calm Steps Poster to hold and point to as you name each step.",
      prep: "This is the model - the poster IS the lesson. Narrate the steps, do not ask students to read the slide. Flip back to it during practice.",
      tag: "[I Do | Explicit teaching | HITS 3, 4]",
    }));
    runSlideDiagnostics(s, pres);
  })();

  /* ───────────────────────────────────────────────────────────────────────
   *  8. I Do - breathe with me (whole-class practice of step 2)
   * ─────────────────────────────────────────────────────────────────────── */

  (function breatheSlide() {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "I Do", { color: C.PRIMARY });
    addTitle(s, "Breathe with me");

    const cx = 5.0, cy = 2.85;
    s.addShape("ellipse", {
      x: cx - 1.25, y: cy - 1.25, w: 2.5, h: 2.5,
      fill: { color: C.SECONDARY, transparency: 80 },
      line: { color: C.SECONDARY, width: 1 },
    });
    s.addShape("ellipse", {
      x: cx - 0.92, y: cy - 0.92, w: 1.84, h: 1.84,
      fill: { color: C.SECONDARY, transparency: 55 },
      line: { color: C.SECONDARY, width: 1 },
    });
    s.addShape("ellipse", {
      x: cx - 0.6, y: cy - 0.6, w: 1.2, h: 1.2,
      fill: { color: C.PRIMARY },
      line: { color: C.PRIMARY, width: 1 },
    });
    s.addText("Breathe", {
      x: cx - 0.9, y: cy - 0.3, w: 1.8, h: 0.6,
      fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    addTextOnShape(s, "Smell the flower - breathe in slowly", {
      x: 0.5, y: 4.2, w: 4.35, h: 0.8, rectRadius: 0.1,
      fill: { color: C.SECONDARY },
    }, {
      fontSize: 16, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0.08,
    });
    addTextOnShape(s, "Blow the candle - breathe out slowly", {
      x: 5.15, y: 4.2, w: 4.35, h: 0.8, rectRadius: 0.1,
      fill: { color: C.ACCENT },
    }, {
      fontSize: 16, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0.08,
    });

    addFooter(s, FOOTER);
    s.addNotes(composeGlanceNotes({
      answer: "open - everyone joins in three slow breaths together",
      beats: [
        "SAY: Let us practise the breathe step together. Sit tall and put a hand on your tummy.",
        "MODEL slowly. SAY: Smell the flower, breathe in through your nose. Blow the candle, breathe out through your mouth.",
        "TIME three slow breaths together, watching the circle. SAY: Feel your body settle. That is your calm step working.",
      ],
      trap: "fast, shallow breaths. Fix: model a slow flower-and-candle breath again, count to three in and three out, child copies.",
      stretch: "ask the child to lead one slow breath for the class.",
      help: "let the child watch you and copy one breath at a time.",
      prep: "Do this for real, not just talk about it. Three slow breaths is plenty for Year 2. Come back to this any time the class needs to reset.",
      tag: "[I Do | Attention and regulation | HITS 3, 9]",
    }));
    runSlideDiagnostics(s, pres);
  })();

  /* ───────────────────────────────────────────────────────────────────────
   *  9. CFU 1 - the first step (reveal pair)
   * ─────────────────────────────────────────────────────────────────────── */

  withReveal(
    () => cfuSlide(
      pres,
      "CFU",
      "Quick check",
      "Show me on your whiteboards",
      "When a big feeling comes, what do we do FIRST?",
      composeGlanceNotes({
        answer: "Stop and freeze your body",
        beats: [
          "ASK: When a big feeling comes, what do we do first? 10 sec, write or draw it, boards up. EXPECT: stop.",
          "SCAN boards, back row first. 80%+ -> reveal and move on. Less -> point to step 1 on the poster, say it together, re-ask.",
          "REVEAL after boards scanned. SAY: Tick yours if you had stop. Fix it if you need to.",
        ],
        trap: "writing ask for help or shout. Fix: we always Stop first so our body can settle, child rewrites step 1.",
        stretch: "write what the second step is too.",
        help: "let the child point to step 1 on the poster instead of writing.",
        prep: "Checks the order of the steps before guided practice. Show me on whiteboards.",
        tag: "[CFU | Supported application | HITS 7, 8]",
      }),
      FOOTER
    ),
    (slide) => {
      addRevealAnswerBar(slide, "Stop and freeze your body", {
        label: "Answer",
        color: C.SUCCESS,
      });
    }
  );

  /* ───────────────────────────────────────────────────────────────────────
   *  10. We Do - guided scenario (class chooses)
   * ─────────────────────────────────────────────────────────────────────── */

  scenarioSlide(
    pres,
    "We Do",
    "Let's help Sam",
    "Sam built a tall tower. Someone bumped it and it fell down. Sam feels really angry and wants to shout.",
    ["Which calm step could Sam try first? Show me on your whiteboard."],
    composeGlanceNotes({
      answer: "Stop first, then breathe. Ask for help if it is still hard.",
      beats: [
        "READ the scenario warmly. SAY: Poor Sam. That is a big, angry feeling. Let us help Sam use a calm step.",
        "ASK: Which calm step could Sam try first? 10 sec, boards up. EXPECT: stop, then breathe.",
        "SCAN boards. 80%+ -> confirm and move on. Less -> point to the poster steps in order, re-ask for Sam.",
      ],
      trap: "choosing ask for help before stopping. Fix: Sam stops and breathes first, then asks if needed - child re-orders the steps.",
      stretch: "say what Sam could ask the grown-up for.",
      help: "give the poster and let the child point to the step for Sam.",
      prep: "Guided practice - the class helps Sam together. Confirm the answer out loud; there is no reveal on this slide.",
      tag: "[We Do | Supported application | HITS 5, 7]",
    }),
    FOOTER
  );

  /* ───────────────────────────────────────────────────────────────────────
   *  11. We Do - partner scenario (fading support)
   * ─────────────────────────────────────────────────────────────────────── */

  scenarioSlide(
    pres,
    "We Do",
    "Your turn with a partner",
    "You are on the playground. You cannot find your friend. Your tummy feels tight and you feel worried.",
    ["Talk with your partner. Which calm step would you try?"],
    composeGlanceNotes({
      answer: "open - any calm step with a reason; listen for stop and breathe, then ask a grown-up",
      beats: [
        "READ the scenario. SAY: A tight tummy is a clue this feeling is bubbling up. Now it is your turn with a partner.",
        "ASK: Which calm step would you try, and why? 40 sec, turn and tell. EXPECT: a step plus a reason. ACCEPT: a step named with help.",
        "CIRCULATE and listen. COLLECT two pairs to share. SAY: Tell us your calm step using because.",
      ],
      trap: "picking a step but not saying why. Fix: give the stem I would ... because ..., child finishes it with their partner.",
      stretch: "plan two steps in order for this worry.",
      help: "offer a choice of two steps for the pair to pick from.",
      prep: "Support fades from whole-class to partners. Use the sentence stem I would ... because ... so every pair gives a reason.",
      tag: "[We Do | Supported application | HITS 5, 9]",
    }),
    FOOTER
  );

  /* ───────────────────────────────────────────────────────────────────────
   *  12. CFU 2 - hinge (reveal pair)
   * ─────────────────────────────────────────────────────────────────────── */

  withReveal(
    () => cfuSlide(
      pres,
      "CFU",
      "Hinge question",
      "Show me 1, 2 or 3 fingers",
      "Kai stopped and breathed but still feels upset. What next?",
      composeGlanceNotes({
        answer: "3 - Ask for help. Tell a grown-up.",
        beats: [
          "POINT to the poster steps 1, 2, 3. ASK: Kai stopped and breathed but still feels upset. What next? 10 sec, hold up 1, 2 or 3 fingers. EXPECT: 3, ask for help.",
          "SCAN fingers. 80%+ -> reveal and move to You Do. Less -> point to step 3 on the poster, say when we ask for help, re-ask.",
          "REVEAL after fingers are up. SAY: When a calm step is not enough, it is brave to ask for help.",
        ],
        trap: "holding up 1 or 2 because they think you only get one step. Fix: show that steps can be used together, and asking is the next one, child re-votes.",
        stretch: "name one grown-up at school they could ask.",
        help: "remind the child of the three steps on the poster before they vote.",
        prep: "Hinge checks that students know asking for help is a real calm step, not a last resort. Finger voting, one to three.",
        tag: "[CFU | Supported application | HITS 7, 9]",
      }),
      FOOTER
    ),
    (slide) => {
      addRevealAnswerBar(slide, "3 - Ask for help", {
        label: "Answer",
        color: C.SUCCESS,
      });
    }
  );

  /* ───────────────────────────────────────────────────────────────────────
   *  13. You Do - My Calm Plan (recording sheet, calm-steps reference)
   * ─────────────────────────────────────────────────────────────────────── */

  contentSlide(
    pres,
    "You Do",
    C.SUCCESS,
    "Make your calm plan",
    [
      "First: draw a big feeling you sometimes have.",
      "Next: circle the calm step you will try.",
      "Then: draw yourself feeling calm.",
    ],
    composeGlanceNotes({
      answer: "open - a named big feeling, one circled calm step, and a picture of feeling calm",
      beats: [
        "SAY: Now it is just you. On your My Calm Plan, draw a big feeling you sometimes have.",
        "SAY: Circle the calm step you will try, then draw yourself feeling calm afterwards.",
        "CIRCULATE. If a child is stuck, point to the poster, not the answer. Note who can say when they would use their step.",
      ],
      trap: "circling every step or none. Fix: prompt them to pick the one they would try first, child circles just one.",
      stretch: "write or say when they would use their calm step.",
      help: "give the poster to copy a step from, and scribe the feeling word for them.",
      prep: "Independent recording on the My Calm Plan sheet, a different form from the We Do talk. The three steps on the right are the choices. Note strong plans to share at the close.",
      tag: "[You Do | Mastery and application | HITS 4, 10]",
    }),
    FOOTER,
    (s, g) => {
      const rx = g.rightX;
      const rw = g.rightW;
      s.addText("Your calm steps", {
        x: rx, y: 1.4, w: rw, h: 0.35,
        fontSize: 16, fontFace: FONT_B, color: C.PRIMARY, bold: true,
        align: "left", valign: "middle", margin: 0,
      });
      const chH = 0.86, gap = 0.2, top = 1.87;
      STEPS.forEach((st, i) => {
        const y = top + i * (chH + gap);
        addTextOnShape(s, st.num + "  " + st.name, {
          x: rx, y, w: rw, h: chH, rectRadius: 0.1,
          fill: { color: st.color },
        }, {
          fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true,
          align: "center", valign: "middle", margin: 0,
        });
      });
    }
  );

  /* ───────────────────────────────────────────────────────────────────────
   *  14. Exit Ticket - assesses SC2 (choose a calm step), no SC label on face
   * ─────────────────────────────────────────────────────────────────────── */

  exitTicketSlide(
    pres,
    ["Name one calm step you can use when you have a big feeling."],
    composeGlanceNotes({
      answer: "any one of: stop, breathe, ask for help",
      beats: [
        "ASK: On your whiteboard, name one calm step you can use when you have a big feeling. 30 sec, boards up. EXPECT: stop, breathe or ask for help.",
        "SCAN every board. 80%+ -> note secure for next session. Less -> flag the calm-step choice for review and re-check it first next session.",
      ],
      trap: "leaving it blank. Fix: these children need the poster in reach - note them and revisit the steps with them.",
      prep: "Checks SC2 (choose a calm step). The SC target stays in these notes only, not on the slide.",
      tag: "[Exit Ticket | Mastery and application | HITS 1, 8]",
    }),
    FOOTER,
    { assessesSc: 2 }
  );

  /* ───────────────────────────────────────────────────────────────────────
   *  15. Closing reflection
   * ─────────────────────────────────────────────────────────────────────── */

  closingSlide(
    pres,
    {
      reflectionPrompt: "Tell your partner one calm step you will use, and when you would use it.",
      scItems: SC,
      selfAssessment: {
        prompt: "Thumbs up, sideways, or down: how ready do you feel to use a calm step?",
        options: ["Ready", "Getting there", "Need more"],
      },
    },
    composeGlanceNotes({
      answer: "open - one calm step and a fitting moment to use it.",
      beats: [
        "SAY: Let us look back at what we set out to do today.",
        "ASK: Tell your partner one calm step you will use, and when. 20 sec, turn and tell. EXPECT: a step plus a moment, like when I feel angry.",
        "SAY: Now show me a thumb - up, sideways or down - for how ready you feel.",
      ],
      care: "keep sharing optional; a child can tell you privately or use the calm corner instead.",
      prep: "Read the three I can statements and have students self-check. Use the thumbs data to plan who to revisit. Acknowledge progress - the class can now name a big feeling and choose a calm step.",
      tag: "[Closing | Retention and recall | HITS 1, 9]",
    })
  );

  /* ───────────────────────────────────────────────────────────────────────
   *  Companion PDFs (resources-session1/)
   * ─────────────────────────────────────────────────────────────────────── */

  await buildPoster();
  await buildCalmPlan();

  const outFile = path.join(OUT_DIR, "Big Feelings Calm Choices Year 2 Wellbeing.pptx");
  await pres.writeFile({ fileName: outFile });
  console.log("PPTX written to " + outFile);
  console.log("PDF written to " + path.join(OUT_DIR, poster.fileName));
  console.log("PDF written to " + path.join(OUT_DIR, calmPlan.fileName));
}

/* ─────────────────────────────────────────────────────────────────────────
 *  PDF: Calm Steps Poster (display for the calm corner and desks)
 * ───────────────────────────────────────────────────────────────────────── */

function drawStopEmblem(doc, cx, cy, r) {
  const pts = [];
  for (let k = 0; k < 8; k += 1) {
    const a = (Math.PI / 8) + k * (Math.PI / 4);
    pts.push([cx + r * Math.cos(a), cy + r * Math.sin(a)]);
  }
  doc.save().lineWidth(3).strokeColor("#FFFFFF").polygon(...pts).stroke().restore();
}

function drawBreatheEmblem(doc, cx, cy, r) {
  doc.save().lineWidth(3).strokeColor("#FFFFFF");
  doc.circle(cx, cy, r).stroke();
  doc.circle(cx, cy, r * 0.62).stroke();
  doc.circle(cx, cy, r * 0.28).stroke();
  doc.restore();
}

function drawAskEmblem(doc, cx, cy, r) {
  doc.save().lineWidth(3).strokeColor("#FFFFFF");
  doc.roundedRect(cx - r, cy - r * 0.8, r * 2, r * 1.3, 8).stroke();
  doc.polygon([cx - r * 0.2, cy + r * 0.5], [cx - r * 0.55, cy + r], [cx + r * 0.1, cy + r * 0.5]).stroke();
  doc.restore();
}

function buildPoster() {
  const doc = createPdf({ title: "Calm Steps Poster" });

  let y = addPdfHeader(doc, "My Calm Steps", {
    subtitle: "Stop. Breathe. Ask for help.",
    color: C.PRIMARY,
    showNameDate: false,
  });

  y += 6;
  const x = PAGE.MARGIN;
  const w = PAGE.CONTENT_W;
  const panelH = 190;
  const gap = 18;
  const emblems = [drawStopEmblem, drawBreatheEmblem, drawAskEmblem];

  STEPS.forEach((st, i) => {
    const py = y + i * (panelH + gap);
    doc.save();
    doc.roundedRect(x, py, w, panelH, 12).fill(hex(st.color));
    doc.restore();

    // Number circle
    doc.save();
    doc.circle(x + 62, py + panelH / 2, 40).fill("#FFFFFF");
    doc.fillColor(hex(st.color)).font("Sans-Bold").fontSize(40)
      .text(st.num, x + 62 - 20, py + panelH / 2 - 24, { width: 40, align: "center", lineBreak: false });
    doc.restore();

    // Step name + action (name forced to one line so "ASK FOR HELP" cannot
    // wrap into the action text below it)
    doc.fillColor("#FFFFFF").font("Sans-Bold").fontSize(28)
      .text(st.name, x + 120, py + panelH / 2 - 40, { width: w - 250, lineBreak: false });
    doc.fillColor("#FFFFFF").font("Sans").fontSize(15)
      .text(st.action, x + 120, py + panelH / 2 + 6, { width: w - 250 });

    // Emblem on the right
    emblems[i](doc, x + w - 68, py + panelH / 2, 42);
  });

  addPdfFooter(doc, FOOTER);
  return writePdf(doc, path.join(OUT_DIR, poster.fileName));
}

/* ─────────────────────────────────────────────────────────────────────────
 *  PDF: My Calm Plan (the You Do recording sheet)
 * ───────────────────────────────────────────────────────────────────────── */

function drawDrawingBox(doc, y, h, caption) {
  const x = PAGE.MARGIN;
  const w = PAGE.CONTENT_W;
  doc.save();
  doc.roundedRect(x, y, w, h, 10).lineWidth(1.3).strokeColor(hex(C.MUTED)).stroke();
  doc.restore();
  if (caption) {
    doc.fontSize(9).font("Sans").fillColor(hex("9CA3AF"))
      .text(caption, x + 10, y + 8, { width: w - 20, lineBreak: false });
  }
  return y + h + 14;
}

function drawStepOptions(doc, y) {
  const x = PAGE.MARGIN;
  const w = PAGE.CONTENT_W;
  const gap = 16;
  const boxW = (w - gap * 2) / 3;
  const boxH = 92;
  STEPS.forEach((st, i) => {
    const bx = x + i * (boxW + gap);
    doc.save();
    doc.roundedRect(bx, y, boxW, boxH, 10).lineWidth(1.6).strokeColor(hex(st.color)).stroke();
    doc.restore();
    doc.save();
    doc.circle(bx + boxW / 2, y + 30, 17).fill(hex(st.color));
    doc.fillColor("#FFFFFF").font("Sans-Bold").fontSize(16)
      .text(st.num, bx + boxW / 2 - 9, y + 22, { width: 18, align: "center", lineBreak: false });
    doc.restore();
    doc.fillColor(hex(st.color)).font("Sans-Bold").fontSize(15)
      .text(st.name, bx + 4, y + 56, { width: boxW - 8, align: "center" });
  });
  return y + boxH + 14;
}

function buildCalmPlan() {
  const doc = createPdf({ title: "My Calm Plan" });

  let y = addPdfHeader(doc, "My Calm Plan", {
    subtitle: "Notice a big feeling and pick a calm step.",
    color: C.PRIMARY,
    showNameDate: true,
  });

  y = addTipBox(doc, "Calm steps: 1 Stop. 2 Breathe. 3 Ask for help.", y, { color: C.SECONDARY });

  y = addSectionHeading(doc, "1. Draw a big feeling you sometimes have.", y, { color: C.ALERT });
  y = drawDrawingBox(doc, y, 150, "Draw here");

  y = addSectionHeading(doc, "2. Circle the calm step you will try.", y, { color: C.PRIMARY });
  y = drawStepOptions(doc, y);

  y = addSectionHeading(doc, "3. Draw yourself feeling calm.", y, { color: C.SUCCESS });
  y = drawDrawingBox(doc, y, 150, "Draw here");

  addPdfFooter(doc, FOOTER);
  return writePdf(doc, path.join(OUT_DIR, calmPlan.fileName));
}

main().catch((err) => {
  console.error("Build failed:", err);
  process.exitCode = 1;
});
