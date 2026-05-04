"use strict";

// Harry the Helpful Hamster - Session 1
// Year 5/6 Wellbeing - 15-20 minute session
// Exploring Emotions with Harry the Helpful Hamster

const path = require("path");
const fs = require("fs");
const PptxGenJS = require("pptxgenjs");
const { createTheme, weekToVariant } = require("../themes/factory");
const { addResourceSlide } = require("../themes/pdf_helpers");

const T = createTheme("wellbeing", "grade56", weekToVariant(1));
const {
  C, FONT_H, FONT_B,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  pairShareSlide,
  addTopBar, addBadge, addTitle, addCard, addFooter,
  CONTENT_TOP, SAFE_BOTTOM,
} = T;

const UNIT = "Harry_the_Helpful_Hamster_Session1";
const LESSON_FOLDER = path.join(__dirname, "..", "output", UNIT);
const PPTX_NAME = "Exploring Emotions with Harry.pptx";
const FOOTER = "Harry the Helpful Hamster | Year 5/6 Wellbeing | Session 1";

// 6 primary emotion segments for the visual emotions wheel.
// Six is enough at classroom distance and stops adjacent labels overlapping.
const WHEEL_EMOTIONS = [
  { name: "Happy",     color: C.SUCCESS },
  { name: "Surprised", color: C.ACCENT },
  { name: "Scared",    color: C.SECONDARY },
  { name: "Sad",       color: C.PRIMARY },
  { name: "Angry",     color: C.ALERT },
  { name: "Calm",      color: C.SUCCESS },
];

// ============================================================
// Teacher notes
// ============================================================

const NOTES_TITLE = [
  "SAY:",
  "- Today we're going to meet Harry the Helpful Hamster.",
  "- Harry helps us notice and name our feelings.",
  "- This is a short 15-20 minute session.",
  "",
  "DO:",
  "- Display the slide as students settle.",
  "- Have your Emotions Wheel chart ready (one per student or display copy).",
  "- Have Harry the Helpful Hamster ready (toy, image, or your school's mascot).",
  "",
  "TEACHER NOTES:",
  "Keep the energy warm and inviting. The aim is for every student to feel safe naming feelings.",
  "",
  "WATCH FOR:",
  "- Students settling in calmly before you begin.",
  "",
  "[Wellbeing: Title | VTLM 2.0: Establishing Purpose]",
].join("\n");

const NOTES_RESOURCES = [
  "SAY:",
  "- Here is everything we need today.",
  "",
  "DO:",
  "- Make sure each student has an Emotions Wheel chart, or display one large copy.",
  "- Have Harry visible (toy, picture, or school mascot wearing spectacles).",
  "- Have mini-whiteboards and markers ready for quick responses.",
  "- Plan a small open space for the role-play moment.",
  "",
  "TEACHER NOTES:",
  "No printed student resources are needed for this short session. Mini-whiteboards and the emotions wheel chart are the only tools.",
  "",
  "WATCH FOR:",
  "- Every student has access to a wheel before the activity begins.",
  "",
  "[Wellbeing: Resources | VTLM 2.0: Preparation]",
].join("\n");

const NOTES_LAUNCH = [
  "SAY:",
  "- This is Harry. Harry is a smart, well-dressed hamster who wears spectacles.",
  "- Why spectacles? Because Harry looks closely at how people feel.",
  "- Harry helps us notice the small signs that show what someone is feeling.",
  "- Quick question: when Harry watches a friend, what might he look at? [face, voice, body, what they're doing]",
  "",
  "DO:",
  "- Hold up your Harry toy or image (or point to the slide character).",
  "- Take 2-3 student responses warmly.",
  "- Connect: 'Today we'll be like Harry -- noticing feelings in ourselves and others.'",
  "",
  "TEACHER NOTES:",
  "Harry is the hook. Keep this short -- about 2 minutes. The character is a friendly entry point, not the lesson itself.",
  "",
  "WATCH FOR:",
  "- Students who name only one cue (e.g. 'face') -- prompt with: 'What else? Voice? Body?'",
  "",
  "[Wellbeing: Launch | VTLM 2.0: Activating Prior Knowledge]",
].join("\n");

const NOTES_LI = [
  "SAY:",
  "- Read the Learning Intention from the slide.",
  "- Read each Success Criterion aloud. We'll come back to these at the end.",
  "",
  "DO:",
  "- Point to each criterion as you read it.",
  "- Leave the slide visible for 15-20 seconds.",
  "",
  "TEACHER NOTES:",
  "Keep this brief. The success criteria stay in the same order on the closing slide so students can self-assess against them.",
  "",
  "WATCH FOR:",
  "- Students unsure about 'name' an emotion -- clarify: 'Use the exact word, like worried or proud.'",
  "",
  "[Wellbeing: LI/SC | VTLM 2.0: Learning Intentions]",
].join("\n");

const NOTES_DISCUSSION = [
  "SAY:",
  "- Why does it help to name our feelings?",
  "- Turn to your partner. Share one emotion you have felt this week and why.",
  "- All emotions are okay. They are part of being human.",
  "",
  "DO:",
  "- Give pairs about 60 seconds.",
  "- Cold call 2-3 students to share one emotion they discussed.",
  "- Acknowledge each emotion warmly. Do not judge.",
  "- Drop in: 'When we can name a feeling, we can talk about it. When we can talk about it, we can take care of it.'",
  "",
  "TEACHER NOTES:",
  "This is a short discussion -- 2 to 3 minutes. The point is to surface that emotions are normal and worth naming. Keep the pace warm and brisk.",
  "",
  "SENSITIVITY ADVISORY:",
  "- What it is: A student may share something heavy (loss, conflict, worry).",
  "- Framing language: 'Thank you for sharing. That's a real feeling, and it makes sense.'",
  "- Watch for: a student who looks upset.",
  "- Protocol: Follow your school's wellbeing referral process if needed.",
  "",
  "WATCH FOR:",
  "- Students sharing only 'happy' or 'good' -- prompt: 'What's a more exact word?'",
  "",
  "[Wellbeing: Discussion | VTLM 2.0: Engaging Learners]",
].join("\n");

const NOTES_WHEEL_IDO = [
  "SAY:",
  "- This is the Emotions Wheel. It helps us find the exact word for what we feel.",
  "- Watch how I use it. I'm going to think of something that happened to me.",
  "- Yesterday, I waited a long time for a parcel and it didn't come. Hmm. I felt... let me look at the wheel.",
  "- I look first at the wider zones. This feels closer to ANGRY than HAPPY.",
  "- Then I look more carefully. Was I angry? Not exactly. I think I was FRUSTRATED.",
  "- Naming it 'frustrated' helps me. It's a smaller, sharper word than 'angry'.",
  "",
  "DO:",
  "- Hold up the Emotions Wheel chart while you talk.",
  "- Point to the segments as you narrow down your example.",
  "- Make every step of your thinking audible -- this is the move students will copy.",
  "",
  "TEACHER NOTES:",
  "The think-aloud is the lesson here. Use a real, low-stakes example from your own week. Modelling 'I started here, then narrowed down' is the move.",
  "",
  "WATCH FOR:",
  "- Don't rush the narrowing step. The thinking move 'big zone first, then exact word' is the part students need to see.",
  "",
  "[Wellbeing: I Do | VTLM 2.0: Explicit Modelling]",
].join("\n");

const NOTES_WEDO = [
  "SAY:",
  "- Now we'll try one together.",
  "- Imagine: you've been working really hard on a project, and your teacher tells you it's brilliant.",
  "- On your wheel, find the zone that fits first.",
  "- Now narrow down. Is it just happy? Or something more exact -- proud, excited, grateful?",
  "- On your mini-whiteboard, write the exact word you'd choose.",
  "",
  "DO:",
  "- Read the scenario from the slide twice.",
  "- Give 20-30 seconds quiet thinking.",
  "- Say: 'Hold up your boards.'",
  "- Scan boards. Pick 3-4 different words to share aloud.",
  "- Affirm: 'There isn't one right answer. The wheel helps us find OUR exact word.'",
  "",
  "CFU CHECKPOINT:",
  "Technique: Mini-whiteboards.",
  "Script:",
  "- Read the scenario, then say: 'Show me the exact emotion word you picked.'",
  "- Scan for: specific words from the wheel (proud, grateful, excited) rather than 'good' or 'fine'.",
  "PROCEED: If most boards show a wheel word, move on.",
  "PIVOT: If many boards show 'good' or 'happy' only, redirect: 'Use your wheel. Look INSIDE the happy zone. What's a sharper word?' Re-ask.",
  "",
  "TEACHER NOTES:",
  "Different students will pick different words. That is the point. The same event can feel different to different people.",
  "",
  "WATCH FOR:",
  "- Students copying their neighbour -- the wheel makes it safe to land somewhere different.",
  "- A student who writes 'embarrassed' or 'nervous' -- honour it. Praise sometimes feels uncomfortable. That is real.",
  "",
  "[Wellbeing: We Do | VTLM 2.0: Guided Practice]",
].join("\n");

const NOTES_CFU = [
  "SAY:",
  "- Listen carefully. I'll read a quick scenario.",
  "- Find the emotion on your wheel. Write the exact word on your whiteboard.",
  "- Scenario: 'You've planned a big day out with a friend, then they cancel at the last minute.'",
  "- Hold up your boards.",
  "",
  "DO:",
  "- Read the scenario twice.",
  "- Allow 15-20 seconds thinking time.",
  "- Scan all boards before discussing.",
  "- Acknowledge the range: disappointed, sad, hurt, frustrated, lonely all sit close together on the wheel.",
  "",
  "CFU CHECKPOINT:",
  "Technique: Show Me Boards.",
  "Script:",
  "- Read the scenario. Wait for boards up.",
  "- Scan for: words from a SAD or ANGRY zone of the wheel (disappointed, hurt, let down, frustrated).",
  "PROCEED: If most boards show a wheel-zone word, move on to the role-play.",
  "PIVOT: If many boards show 'mad' or 'bad' only, prompt: 'Use the wheel. What's the sharper word for that feeling?' Re-ask.",
  "",
  "TEACHER NOTES:",
  "There is no single right answer. The point is using a wheel word, not 'mad' or 'bad'.",
  "",
  "WATCH FOR:",
  "- Students who write more than one word -- excellent. Acknowledge that feelings can mix.",
  "",
  "[Wellbeing: CFU | VTLM 2.0: Check For Understanding]",
].join("\n");

const NOTES_YOUDO = [
  "SAY:",
  "- Now we'll do quick role-plays.",
  "- I'll read a scenario. One volunteer acts it out without words -- just face and body.",
  "- The rest of you watch like Harry. Use your wheel to name the emotion.",
  "- Hold up your whiteboard with the word you choose.",
  "",
  "DO:",
  "- Choose a volunteer for each scenario. Quick rotations -- 30 seconds each.",
  "- After each role-play: 'Show me your word.' Scan boards. Take 1-2 to share.",
  "- Use the three scenarios on the slide, or your own to suit the class.",
  "- Keep it light and warm.",
  "",
  "TEACHER NOTES:",
  "Role-play is the engagement payoff. Pick volunteers who'll have fun with it. About 4 minutes total.",
  "",
  "ENABLING & EXTENDING:",
  "ENABLING PROMPT:",
  "- Task: Give the volunteer a short emotion menu (happy, frustrated, sad) to choose from before acting.",
  "- Extra Notes: Sentence frame for watchers: 'I think they felt ___ because ___.'",
  "EXTENDING PROMPT:",
  "- Task: Watchers name the emotion AND one body sign that gave it away (e.g. crossed arms, slumped shoulders).",
  "",
  "SENSITIVITY ADVISORY:",
  "- What it is: A student might choose not to volunteer. That is fine.",
  "- Framing language: 'Watching like Harry is just as important as acting.'",
  "- Watch for: Students who avoid eye contact when role-play is mentioned.",
  "- Protocol: Never single a student out. Volunteers only.",
  "",
  "WATCH FOR:",
  "- Role-plays drifting into silliness -- gentle reset: 'Harry would notice the small signs. Show us those.'",
  "",
  "[Wellbeing: You Do | VTLM 2.0: Independent Practice]",
].join("\n");

const NOTES_CLOSING = [
  "SAY:",
  "- Let's check our success criteria. Thumbs up, sideways, or down for each one.",
  "- I can name a range of emotions.",
  "- I can use the Emotions Wheel to find an exact feeling word.",
  "- I can notice clues in others that show how they feel.",
  "- Reflection: Share one new emotion word you'll take with you today.",
  "",
  "DO:",
  "- Read each criterion aloud; pause for thumbs.",
  "- Take 3-4 quick shares. Acknowledge each warmly.",
  "- Close with: 'Be like Harry. Notice. Name it. Be kind.'",
  "",
  "TEACHER NOTES:",
  "Connect forward: 'Next time we'll think about what to do once we've named a feeling.' Keep the close warm and short.",
  "",
  "WATCH FOR:",
  "- Sideways or down on any criterion -- note for follow-up next session.",
  "- A student who lingers -- check in quietly after class.",
  "",
  "[Wellbeing: Closing | VTLM 2.0: Review and Reflect]",
].join("\n");

// ============================================================
// Helper: draw Harry the Helpful Hamster (placeholder character)
// ============================================================
//
// Built with PptxGenJS shapes — this is a friendly placeholder. If the school
// has its own Harry image, the teacher can swap this slide visual.

function drawHarry(s, cx, cy, scale) {
  const sc = scale || 1;
  // Body (round, soft)
  s.addShape("ellipse", {
    x: cx - 1.0 * sc, y: cy - 0.20 * sc, w: 2.0 * sc, h: 1.7 * sc,
    fill: { color: "C9A37A" },
    line: { color: "8B6F4E", width: 1.5 },
  });
  // Belly
  s.addShape("ellipse", {
    x: cx - 0.55 * sc, y: cy + 0.25 * sc, w: 1.1 * sc, h: 1.0 * sc,
    fill: { color: "F4E0C5" },
    line: { color: "F4E0C5", width: 0.5 },
  });
  // Head
  s.addShape("ellipse", {
    x: cx - 0.85 * sc, y: cy - 1.10 * sc, w: 1.7 * sc, h: 1.4 * sc,
    fill: { color: "C9A37A" },
    line: { color: "8B6F4E", width: 1.5 },
  });
  // Ears
  s.addShape("ellipse", {
    x: cx - 0.80 * sc, y: cy - 1.30 * sc, w: 0.45 * sc, h: 0.45 * sc,
    fill: { color: "8B6F4E" },
    line: { color: "8B6F4E", width: 1 },
  });
  s.addShape("ellipse", {
    x: cx + 0.35 * sc, y: cy - 1.30 * sc, w: 0.45 * sc, h: 0.45 * sc,
    fill: { color: "8B6F4E" },
    line: { color: "8B6F4E", width: 1 },
  });
  // Inner ears
  s.addShape("ellipse", {
    x: cx - 0.72 * sc, y: cy - 1.22 * sc, w: 0.28 * sc, h: 0.28 * sc,
    fill: { color: "F4C7B3" },
    line: { color: "F4C7B3", width: 0.5 },
  });
  s.addShape("ellipse", {
    x: cx + 0.43 * sc, y: cy - 1.22 * sc, w: 0.28 * sc, h: 0.28 * sc,
    fill: { color: "F4C7B3" },
    line: { color: "F4C7B3", width: 0.5 },
  });
  // Spectacles (the key feature) — two round frames + bridge
  const eyeY = cy - 0.78 * sc;
  const lensR = 0.32 * sc;
  // Left lens
  s.addShape("ellipse", {
    x: cx - 0.55 * sc, y: eyeY - lensR / 2, w: lensR, h: lensR,
    fill: { color: "FFFFFF" },
    line: { color: "1A1A1A", width: 2.5 },
  });
  // Right lens
  s.addShape("ellipse", {
    x: cx + 0.23 * sc, y: eyeY - lensR / 2, w: lensR, h: lensR,
    fill: { color: "FFFFFF" },
    line: { color: "1A1A1A", width: 2.5 },
  });
  // Bridge
  s.addShape("line", {
    x: cx - 0.23 * sc, y: eyeY - 0.02, w: 0.46 * sc, h: 0,
    line: { color: "1A1A1A", width: 2.5 },
  });
  // Eyes (dots inside lenses)
  s.addShape("ellipse", {
    x: cx - 0.42 * sc, y: eyeY - 0.06, w: 0.10 * sc, h: 0.10 * sc,
    fill: { color: "1A1A1A" },
    line: { color: "1A1A1A", width: 0.5 },
  });
  s.addShape("ellipse", {
    x: cx + 0.36 * sc, y: eyeY - 0.06, w: 0.10 * sc, h: 0.10 * sc,
    fill: { color: "1A1A1A" },
    line: { color: "1A1A1A", width: 0.5 },
  });
  // Nose
  s.addShape("ellipse", {
    x: cx - 0.07 * sc, y: cy - 0.40 * sc, w: 0.16 * sc, h: 0.12 * sc,
    fill: { color: "8B3A3A" },
    line: { color: "8B3A3A", width: 0.5 },
  });
  // Smile
  s.addShape("line", {
    x: cx - 0.18 * sc, y: cy - 0.22 * sc, w: 0.36 * sc, h: 0,
    line: { color: "1A1A1A", width: 2 },
  });
}

// ============================================================
// Helper: draw a simple emotions wheel (8 segments)
// ============================================================
//
// PptxGenJS doesn't render arc/pie wedges cleanly. We approximate the wheel
// with a coloured ring of 8 rounded-rect "spokes" arranged radially around
// a centre disc. This reads as a wheel from classroom distance and labels
// each primary emotion clearly.

function drawEmotionsWheel(s, cx, cy, radius) {
  const segments = WHEEL_EMOTIONS.length;
  const innerR = radius * 0.30;
  // Outer ring backdrop
  s.addShape("ellipse", {
    x: cx - radius, y: cy - radius, w: radius * 2, h: radius * 2,
    fill: { color: C.BG_LIGHT },
    line: { color: C.PRIMARY, width: 1.5 },
  });

  // Six wedge labels positioned around the wheel.
  // labelR pushed close to the outer edge so adjacent pills don't overlap.
  for (let i = 0; i < segments; i++) {
    const angle = (i / segments) * 2 * Math.PI - Math.PI / 2;
    const labelR = radius * 0.72;
    const lx = cx + Math.cos(angle) * labelR;
    const ly = cy + Math.sin(angle) * labelR;
    const labelW = 0.95;
    const labelH = 0.38;

    // Coloured pill at each segment
    s.addShape("roundRect", {
      x: lx - labelW / 2, y: ly - labelH / 2, w: labelW, h: labelH, rectRadius: 0.20,
      fill: { color: WHEEL_EMOTIONS[i].color },
      line: { color: WHEEL_EMOTIONS[i].color, width: 0.5 },
    });
    s.addText(WHEEL_EMOTIONS[i].name, {
      x: lx - labelW / 2, y: ly - labelH / 2, w: labelW, h: labelH,
      fontSize: 11, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
  }

  // Centre disc with title
  s.addShape("ellipse", {
    x: cx - innerR, y: cy - innerR, w: innerR * 2, h: innerR * 2,
    fill: { color: C.PRIMARY },
    line: { color: C.PRIMARY, width: 1 },
  });
  s.addText("Emotions\nWheel", {
    x: cx - innerR, y: cy - innerR, w: innerR * 2, h: innerR * 2,
    fontSize: 11, fontFace: FONT_H, color: C.WHITE, bold: true,
    align: "center", valign: "middle", margin: 0,
  });
}

// ============================================================
// Build
// ============================================================

async function build() {
  const pres = new PptxGenJS();
  pres.layout = "LAYOUT_16x9";

  // -- Slide 1: Title --
  titleSlide(
    pres,
    "Exploring Emotions\nwith Harry",
    "Harry the Helpful Hamster",
    "Year 5/6 Wellbeing  |  Session 1  |  15-20 minutes",
    NOTES_TITLE
  );

  // -- Slide 2: Teacher Resources (immediately after title) --
  // No printed PDFs for this 15-20 minute session. The rich resource slide
  // takes a config object so we can list materials, student tools and a
  // small board-setup note without inventing a worksheet.
  addResourceSlide(pres, {
    resources: [],
    manipulatives: [
      "Emotions Wheel chart (one per student, or one large display copy)",
      "Harry the Helpful Hamster (toy, picture, or your school's mascot wearing spectacles)",
    ],
    studentTools: [
      "Mini-whiteboards and markers",
    ],
    boardSetup: [
      "Small open space for short volunteer role-plays",
    ],
  }, T, FOOTER, NOTES_RESOURCES);

  // -- Slide 3: Launch — Meet Harry --
  contentSlide(
    pres,
    "Launch",
    C.SECONDARY,
    "Meet Harry the Helpful Hamster",
    [
      "Harry is smart, well-dressed, and wears spectacles",
      "Harry helps us notice and name our feelings",
      "Today we'll be observers, just like Harry",
    ],
    NOTES_LAUNCH,
    FOOTER,
    (s, lg) => {
      const rX = lg.rightX;
      const rW = lg.rightW;
      const topY = lg.panelTopPadded;
      const cardH = 3.4;

      addCard(s, rX, topY, rW, cardH, { strip: C.SECONDARY, fill: C.WHITE });
      // Harry centred in the right panel
      const harryCx = rX + rW / 2;
      const harryCy = topY + cardH / 2 + 0.10;
      drawHarry(s, harryCx, harryCy, 1.0);
    }
  );

  // -- Slide 4: LI / SC --
  liSlide(
    pres,
    ["We are learning to recognise and name a range of emotions in ourselves and others"],
    [
      "I can name a range of emotions",
      "I can use the Emotions Wheel to find an exact feeling word",
      "I can notice clues in others that show how they feel",
    ],
    NOTES_LI,
    FOOTER
  );

  // -- Slide 5: Discussion — Why it matters --
  pairShareSlide(
    pres,
    "Why Name Our Feelings?",
    [
      "Turn & Talk: share one emotion you have felt this week",
      "Why does it help to put a name to a feeling?",
      "All emotions are okay -- they are part of being human",
    ],
    NOTES_DISCUSSION,
    FOOTER
  );

  // -- Slide 6: I Do — The Emotions Wheel --
  contentSlide(
    pres,
    "I Do",
    C.PRIMARY,
    "Watch How I Use the Wheel",
    [
      "Step 1: Think of something that happened",
      "Step 2: Find the closest zone on the wheel",
      "Step 3: Narrow down to the exact word",
    ],
    NOTES_WHEEL_IDO,
    FOOTER,
    (s, lg) => {
      const rX = lg.rightX;
      const rW = lg.rightW;
      const topY = lg.panelTopPadded;
      const cardH = 3.5;

      addCard(s, rX, topY, rW, cardH, { strip: C.PRIMARY, fill: C.WHITE });
      // Wheel centred in the right card
      const wheelCx = rX + rW / 2;
      const wheelCy = topY + cardH / 2 + 0.10;
      const radius = Math.min(rW / 2 - 0.30, cardH / 2 - 0.30);
      drawEmotionsWheel(s, wheelCx, wheelCy, radius);
    }
  );

  // -- Slide 7: We Do — class uses the wheel together --
  {
    const s = pres.addSlide();
    addTopBar(s, C.SECONDARY);
    addBadge(s, "We Do", { color: C.SECONDARY });
    addTitle(s, "Find the Exact Word");

    // Hero scenario card (the main task)
    const cardY = CONTENT_TOP + 0.10;
    const cardH = 1.85;
    addCard(s, 0.5, cardY, 9, cardH, { strip: C.SECONDARY, fill: C.WHITE });
    s.addText("You've worked really hard on a project. Your teacher tells you it's brilliant.", {
      x: 0.85, y: cardY + 0.18, w: 8.3, h: cardH - 0.36,
      fontSize: 22, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    // Two response cards below
    const respY = cardY + cardH + 0.20;
    const respH = SAFE_BOTTOM - respY - 0.10;
    const colW = (9 - 0.20) / 2;

    // Left: Find the zone
    addCard(s, 0.5, respY, colW, respH, { strip: C.PRIMARY, fill: C.BG_LIGHT });
    s.addText("Find the zone on your wheel", {
      x: 0.65, y: respY + 0.15, w: colW - 0.30, h: 0.45,
      fontSize: 15, fontFace: FONT_H, color: C.PRIMARY, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
    s.addText("Big feeling first. Where does it sit?", {
      x: 0.65, y: respY + 0.65, w: colW - 0.30, h: respH - 0.80,
      fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL,
      align: "center", valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    // Right: Pick the exact word
    addCard(s, 0.5 + colW + 0.20, respY, colW, respH, { strip: C.ACCENT, fill: C.BG_LIGHT });
    s.addText("Pick the exact word", {
      x: 0.65 + colW + 0.20, y: respY + 0.15, w: colW - 0.30, h: 0.45,
      fontSize: 15, fontFace: FONT_H, color: C.ACCENT, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
    s.addText("Mini-whiteboards: write your word.", {
      x: 0.65 + colW + 0.20, y: respY + 0.65, w: colW - 0.30, h: respH - 0.80,
      fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL,
      align: "center", valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_WEDO);
  }

  // -- Slide 8: CFU — quick hinge check --
  cfuSlide(
    pres,
    "CFU",
    "What Would You Feel?",
    "Show Me Boards",
    "You've planned a big day out with a friend. They cancel at the last minute.\n\nUse your Emotions Wheel. Write the exact word on your whiteboard.",
    NOTES_CFU,
    FOOTER
  );

  // -- Slide 9: You Do — role play scenarios --
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "You Do", { color: C.PRIMARY });
    addTitle(s, "Be Like Harry: Read the Feeling");

    // Instruction strip (small — task is the hero, not the steps)
    const stripY = CONTENT_TOP + 0.10;
    const stripH = 0.55;
    s.addShape("roundRect", {
      x: 0.5, y: stripY, w: 9, h: stripH, rectRadius: 0.10,
      fill: { color: C.BG_LIGHT },
      line: { color: C.MUTED, width: 0.5 },
    });
    s.addText("A volunteer acts. The class watches like Harry and writes the emotion on a whiteboard.", {
      x: 0.7, y: stripY, w: 8.6, h: stripH,
      fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
      align: "center", valign: "middle", margin: 0,
    });

    // Three scenario cards (the hero items)
    const startY = stripY + stripH + 0.20;
    const availH = SAFE_BOTTOM - startY - 0.05;
    const gap = 0.15;
    const cardH = (availH - gap * 2) / 3;
    const scenarios = [
      { label: "Scenario 1", text: "You just won a game you've been trying really hard to win.", color: C.SUCCESS },
      { label: "Scenario 2", text: "Your pet has been unwell. The vet has just left after looking at it.", color: C.PRIMARY },
      { label: "Scenario 3", text: "You've been stuck on the same homework problem for ten minutes.", color: C.ALERT },
    ];

    scenarios.forEach((sc, i) => {
      const y = startY + i * (cardH + gap);
      addCard(s, 0.5, y, 9, cardH, { strip: sc.color, fill: C.WHITE });
      // Label pill
      s.addShape("roundRect", {
        x: 0.85, y: y + 0.15, w: 1.4, h: cardH - 0.30, rectRadius: 0.06,
        fill: { color: sc.color },
      });
      s.addText(sc.label, {
        x: 0.85, y: y + 0.15, w: 1.4, h: cardH - 0.30,
        fontSize: 13, fontFace: FONT_H, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      // Scenario text (the hero)
      s.addText(sc.text, {
        x: 2.40, y: y + 0.10, w: 6.50, h: cardH - 0.20,
        fontSize: 18, fontFace: FONT_B, color: C.CHARCOAL,
        align: "left", valign: "middle", margin: 0,
        fit: "shrink", shrinkText: true,
      });
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  }

  // -- Slide 10: Closing --
  closingSlide(
    pres,
    {
      reflectionPrompt: "Share one new emotion word you'll take with you today. Be like Harry: notice, name it, be kind.",
      scItems: [
        "I can name a range of emotions",
        "I can use the Emotions Wheel to find an exact feeling word",
        "I can notice clues in others that show how they feel",
      ],
      selfAssessment: {
        prompt: "Self-assess: thumbs up, sideways, or down for each criterion.",
        options: ["Got it", "Getting there", "Need more practice"],
      },
    },
    NOTES_CLOSING
  );

  // Write PPTX
  fs.mkdirSync(LESSON_FOLDER, { recursive: true });
  await pres.writeFile({ fileName: path.join(LESSON_FOLDER, PPTX_NAME) });
  console.log("PPTX written to", path.join(LESSON_FOLDER, PPTX_NAME));
  console.log("\nBuild complete!");
  console.log("Output folder:", LESSON_FOLDER);
}

build().catch((err) => { console.error(err); process.exit(1); });
