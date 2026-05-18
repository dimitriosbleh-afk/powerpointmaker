"use strict";

// Harry the Helpful Hamster - Session 3
// Year 6 Wellbeing - 15-20 minute session
// Expressing Emotions Like Harry

const path = require("path");
const fs = require("fs");
const PptxGenJS = require("pptxgenjs");
const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addPdfFooter, addBodyText,
  addResourceSlide, makeSessionResource,
} = require("../themes/pdf_helpers");

const T = createTheme("wellbeing", "grade56", weekToVariant(3));
const {
  C, FONT_H, FONT_B,
  titleSlide, liSlide, contentSlide, closingSlide,
  pairShareSlide, reflectionSlide,
  addTopBar, addBadge, addTitle, addCard, addFooter,
  withReveal,
  CONTENT_TOP, SAFE_BOTTOM,
} = T;

const UNIT = "Harry_the_Helpful_Hamster_Session3";
const LESSON_FOLDER = path.join(__dirname, "..", "output", UNIT);
const PPTX_NAME = "Expressing Emotions Like Harry.pptx";
const FOOTER = "Harry the Helpful Hamster | Year 6 Wellbeing | Session 3";
const SESSION_NUMBER = 3;

// Three core strategies for this session. Colours are kept consistent within
// the wellbeing palette and reused on every strategy reference.
const STRAT_I    = { key: "I",       name: "I Statement",    color: C.PRIMARY };
const STRAT_BR   = { key: "Breathe", name: "Deep Breathing", color: C.SECONDARY };
const STRAT_HELP = { key: "Help",    name: "Seek Help",      color: C.ACCENT };

// 12 role-play scenarios paired to a primary strategy. Used by the
// printable Role-Play Cards and the teacher Strategy Guide answer key.
const SCENARIOS = [
  { text: "Your friend keeps interrupting you when you're trying to share an idea in a group.",      strategy: STRAT_I },
  { text: "Your heart is racing and your hands are sweaty just before you have to present to the class.", strategy: STRAT_BR },
  { text: "Someone has been teasing you online for days, and it's starting to get to you.",          strategy: STRAT_HELP },
  { text: "Your sibling went into your room and used your things without asking.",                    strategy: STRAT_I },
  { text: "You lost your temper in PE. You want to calm down before the next round so you can play fairly.", strategy: STRAT_BR },
  { text: "Something at home has been on your mind for days and you can't stop thinking about it.",   strategy: STRAT_HELP },
  { text: "A group member hasn't done their part of the project, and it's due tomorrow.",            strategy: STRAT_I },
  { text: "You're about to start a maths test and your mind has gone blank.",                         strategy: STRAT_BR },
  { text: "A friend has told you they're being bullied and asked you to keep it secret.",            strategy: STRAT_HELP },
  { text: "Your friend keeps borrowing your things and not returning them.",                          strategy: STRAT_I },
  { text: "You feel hot, your fists are clenching, and you want to shout -- but you're in class.",   strategy: STRAT_BR },
  { text: "You haven't understood the last few lessons in a subject and you're feeling lost.",        strategy: STRAT_HELP },
];

// ============================================================
// Resources (printable PDFs)
// ============================================================

const ROLE_PLAY_RES = makeSessionResource(
  SESSION_NUMBER,
  "Role-Play Cards",
  "Print, cut along the lines, and distribute to pairs or small groups."
);
const STRATEGY_GUIDE_RES = makeSessionResource(
  SESSION_NUMBER,
  "Strategy Guide",
  "Teacher reference. Maps each scenario to a suggested strategy."
);

const RESOURCES = [ROLE_PLAY_RES, STRATEGY_GUIDE_RES];

// ============================================================
// Teacher notes
// ============================================================

const NOTES_TITLE = [
  "SAY:",
  "- Welcome back, Harry's helpers.",
  "- Last time we used Harry's wheel to NAME feelings.",
  "- Today we'll learn how to EXPRESS feelings safely -- like Harry would.",
  "- This is a short 15-20 minute session.",
  "",
  "DO:",
  "- Display the title as students settle.",
  "- Have the printed Role-Play Cards cut and ready in stacks.",
  "- Have the Strategy Guide on your desk for quick reference.",
  "",
  "TEACHER NOTES:",
  "Keep the tone warm and respectful. Expressing emotions takes courage. Make it safe to practise.",
  "",
  "WATCH FOR:",
  "- Students settling calmly and ready to listen.",
  "",
  "[Wellbeing: Title | VTLM 2.0: Establishing Purpose]",
].join("\n");

const NOTES_RESOURCES = [
  "SAY:",
  "- Here's what we'll be using today.",
  "",
  "DO:",
  "- Make sure each pair or small group has a stack of Role-Play Cards.",
  "- Keep Harry's Emotions Wheel from last session visible if you have it.",
  "- Have the Strategy Guide on your desk -- not for students.",
  "- Have paper or workbooks ready for the personal reflection.",
  "",
  "TEACHER NOTES:",
  "Two printable PDFs are included: Role-Play Cards (cut up before the lesson) and a Strategy Guide (teacher use only). No worksheets -- students record reflections in their normal book or on a half sheet.",
  "",
  "WATCH FOR:",
  "- Every pair has at least one stack of Role-Play Cards before the role-play activity begins.",
  "",
  "[Wellbeing: Resources | VTLM 2.0: Preparation]",
].join("\n");

const NOTES_LAUNCH = [
  "SAY:",
  "- Last time Harry helped us NAME feelings on the wheel.",
  "- Quick recap on your fingers: 1 finger if you can still remember the six big feelings, 2 fingers if some have slipped, 3 fingers if you'd like a reminder.",
  "- Some of you may remember Harry's wheel. If this feels new, that's okay -- we'll build it together.",
  "- Today Harry will help us with the next step: how do we EXPRESS the feeling in a way that helps, not hurts?",
  "",
  "DO:",
  "- Take a brisk finger check-in scan.",
  "- Acknowledge: 'Naming a feeling is step one. Expressing it well is step two.'",
  "- Link forward: 'Today we'll learn three of Harry's strategies for expressing feelings.'",
  "",
  "TEACHER NOTES:",
  "Keep the launch short -- about 90 seconds. The check-in primes students by connecting to last session's wheel work.",
  "",
  "WATCH FOR:",
  "- A student showing three fingers -- a discreet check-in with them later.",
  "",
  "[Wellbeing: Launch | VTLM 2.0: Activating Prior Knowledge]",
].join("\n");

const NOTES_LI = [
  "SAY:",
  "- Read the Learning Intention together.",
  "- Read each Success Criterion aloud. We'll come back to these at the end.",
  "",
  "DO:",
  "- Point to each criterion as you read it.",
  "- Leave the slide visible for 15-20 seconds.",
  "",
  "TEACHER NOTES:",
  "Keep this brief. The same three criteria appear on the closing slide so students can self-assess against them.",
  "",
  "WATCH FOR:",
  "- A student unsure about 'I statement' -- a quick paraphrase: 'a calm way of saying how you feel without blaming'.",
  "",
  "[Wellbeing: LI/SC | VTLM 2.0: Learning Intentions]",
].join("\n");

const NOTES_STRATEGIES = [
  "SAY:",
  "- Harry has three favourite strategies for expressing feelings safely.",
  "- One: an I Statement -- a calm sentence about how you feel.",
  "- Two: Deep Breathing -- calm your body first, before words.",
  "- Three: Seek Help -- find a trusted person when the feeling is too big to handle alone.",
  "- Different strategies suit different moments. Sometimes you use more than one.",
  "",
  "DO:",
  "- Point to each strategy card as you name it.",
  "- Keep the pace warm -- this is an overview, not a deep lecture.",
  "- Reassure: 'Asking for help is brave, not weak. Harry knows.'",
  "",
  "TEACHER NOTES:",
  "These three strategies are the lesson's spine. The next slides drill into the I Statement, then we apply all three in scenarios and role-play. About 2 minutes.",
  "",
  "WATCH FOR:",
  "- A student who thinks 'Seek Help' is only for big problems -- gently reframe: 'It's for any feeling that feels too big to handle alone.'",
  "",
  "[Wellbeing: Concept Build | VTLM 2.0: Explicit Modelling]",
].join("\n");

const NOTES_IDO_ISTAT = [
  "SAY:",
  "- Watch how I build an I Statement.",
  "- An I Statement has three parts: I FEEL ___ WHEN ___ BECAUSE ___.",
  "- I start with the feeling -- not with what the other person did.",
  "- Situation I'll use: 'My partner keeps interrupting me while I'm sharing an idea.'",
  "- I feel FRUSTRATED -- that's my wheel word. WHEN you interrupt me. BECAUSE I lose my train of thought.",
  "- Notice: I didn't say 'You always interrupt me!' I said how I feel and why. That keeps the conversation open.",
  "",
  "DO:",
  "- Point to each part of the frame as you fill it in.",
  "- Make the think-aloud audible -- this is the move students will copy.",
  "- Keep the example low-stakes; this is modelling, not personal sharing.",
  "",
  "TEACHER NOTES:",
  "The frame I FEEL ___ WHEN ___ BECAUSE ___ is the lesson's main teachable move. Students will reuse it for the rest of the session. About 2-3 minutes.",
  "",
  "WATCH FOR:",
  "- Don't rush. Read each part of the frame slowly. The shape of the sentence is what students need to remember.",
  "",
  "[Wellbeing: I Do | VTLM 2.0: Explicit Modelling]",
].join("\n");

const NOTES_WEDO = [
  "SAY:",
  "- Now let's build one together.",
  "- Read the scenario on the slide with me.",
  "- On your mini-whiteboard or in your book, write an I Statement using Harry's frame: I FEEL ___ WHEN ___ BECAUSE ___.",
  "- Hold up your boards when you're ready.",
  "",
  "DO:",
  "- Read the scenario aloud twice.",
  "- Give 45-60 seconds quiet writing.",
  "- Say: 'Show me your I Statement.'",
  "- Scan boards. Pick 2-3 different responses to share aloud.",
  "- Affirm the range: 'Different students might pick different feeling words. The frame is the same.'",
  "- Click to reveal one strong answer.",
].join("\n");

const NOTES_WEDO_REVEAL = [
  NOTES_WEDO,
  "",
  "CFU CHECKPOINT:",
  "Technique: Show Me Boards.",
  "Script:",
  "- Read the scenario, then say: 'Show me your I Statement.'",
  "- Scan for: all three parts of the frame -- a feeling word, a 'when' part, and a 'because' part.",
  "PROCEED: If most boards show the full frame, move on.",
  "PIVOT: If many boards say things like 'You're so annoying!' or only one part of the frame, model again: 'Start with I FEEL, not YOU. What's the feeling word from Harry's wheel?' Re-ask.",
  "",
  "TEACHER NOTES:",
  "There isn't one right answer. The frame is what we're checking. Affirm any reasonable feeling word.",
  "",
  "WATCH FOR:",
  "- A student writing 'You make me feel...' -- gently redirect to 'I feel...' Owning the feeling is the key move.",
  "",
  "[Wellbeing: We Do | VTLM 2.0: Guided Practice]",
].join("\n");

const NOTES_CFU = [
  "SAY:",
  "- Quick check. Listen carefully.",
  "- Scenario: 'You can feel your face getting hot and your fists clenching after losing an important game.'",
  "- Which strategy fits BEST right now? Write it on your whiteboard.",
  "- Your choices are: I Statement, Deep Breathing, or Seek Help.",
  "",
  "DO:",
  "- Read the scenario twice.",
  "- Allow 15-20 seconds thinking.",
  "- Scan all boards before discussing.",
  "- Reveal: Deep Breathing is the strongest first move -- the body needs to calm before words can come out right.",
  "- Honour other answers: 'I Statement could come AFTER you've calmed your body. Seek Help fits if the feeling stays big.'",
  "",
  "CFU CHECKPOINT:",
  "Technique: Show Me Boards (hinge).",
  "Script:",
  "- Read scenario, wait for boards up.",
  "- Scan for: Deep Breathing (best first move). Other answers are defensible -- accept with a brief 'why?'.",
  "PROCEED: If most boards show Deep Breathing or a reasoned alternative, move on to role-play.",
  "PIVOT: If many show I Statement first, prompt: 'When your body is THAT hot, are the right words easy to find? What does the body need first?' Re-ask.",
  "",
  "TEACHER NOTES:",
  "The hinge here is order. Deep breathing first, then words. Honour layered answers like 'breathe then talk'.",
  "",
  "WATCH FOR:",
  "- A student picking Seek Help -- valid if the feeling is overwhelming. Acknowledge: 'That works too. Harry would say there's no wrong door.'",
  "",
  "[Wellbeing: CFU | VTLM 2.0: Check For Understanding]",
].join("\n");

const NOTES_YOUDO = [
  "SAY:",
  "- Now you'll practise in pairs.",
  "- Take a Role-Play Card from the stack.",
  "- Read it together. Decide which of Harry's three strategies fits best.",
  "- Then PRACTISE it. If it's an I Statement, say the sentence aloud using Harry's frame. If it's Deep Breathing, do a slow 4-in, 4-hold, 4-out together. If it's Seek Help, say WHO you'd ask and WHAT you'd say.",
  "- Swap and try another card. Aim to practise 3-4 cards in the time we have.",
  "",
  "DO:",
  "- Set the timer: about 5-6 minutes for the activity.",
  "- Move between pairs. Don't correct -- ask: 'Why that strategy? How did the sentence sound?'",
  "- Use the Strategy Guide on your desk to help with tricky cards, but accept reasoned alternatives.",
  "",
  "ENABLING & EXTENDING:",
  "ENABLING PROMPT:",
  "- Task: Give the pair the I Statement frame on a sticky note. Limit to two strategies (I Statement or Deep Breathing).",
  "- Extra Notes: Sentence frame -- 'I feel ___ when ___ because ___.'",
  "EXTENDING PROMPT:",
  "- Task: For each card, name TWO strategies that could work. Say which is best FIRST and which works as a back-up.",
  "- Extra Notes: 'Real moments often need more than one strategy. Which order makes sense?'",
  "",
  "SENSITIVITY ADVISORY:",
  "- What it is: Some scenarios touch on bullying, conflict at home, or feeling lost. Real memories may surface.",
  "- Framing language: 'You don't have to share your own story. You're practising how someone in this situation might respond.'",
  "- Watch for: A student who goes quiet on a card or wants to skip it.",
  "- Protocol: Allow students to set a card aside. Quietly check in after the activity. Follow your school's wellbeing referral process if needed.",
  "",
  "WATCH FOR:",
  "- Pairs racing through without actually PRACTISING the strategy -- redirect: 'Say the I Statement out loud. I want to hear it.'",
  "- Pairs stuck on one card -- it's okay to set it aside.",
  "",
  "[Wellbeing: You Do | VTLM 2.0: Independent Practice]",
].join("\n");

const NOTES_REFLECTION = [
  "SAY:",
  "- Now think of a real moment recently when you felt a strong feeling.",
  "- In your book, write three short answers:",
  "  1) What was the feeling? Use Harry's wheel if you need to.",
  "  2) Write an I Statement for that moment. I FEEL ___ WHEN ___ BECAUSE ___.",
  "  3) Which other strategy could also have helped: deep breathing, seeking help, or another I Statement to a different person?",
  "- You don't have to share this. It's for you.",
  "",
  "DO:",
  "- Give about 3-4 minutes of quiet writing.",
  "- Circulate and acknowledge effort without reading over shoulders.",
  "- Remind softly halfway: 'Stick to the frame: I FEEL, WHEN, BECAUSE.'",
  "",
  "TEACHER NOTES:",
  "This is private reflection, not a writing task to mark. The point is for students to apply the frame to their own life. About 3-4 minutes.",
  "",
  "SENSITIVITY ADVISORY:",
  "- What it is: Some students may write about something heavy.",
  "- Framing language: 'This is your private writing. I won't read it unless you'd like me to.'",
  "- Watch for: A student showing signs of distress while writing.",
  "- Protocol: Quietly check in. Follow your school's wellbeing referral process if needed.",
  "",
  "WATCH FOR:",
  "- A student who freezes -- offer the enabling prompt: 'Start with the wheel word. Then plug it into the frame.'",
  "",
  "[Wellbeing: Personal Reflection | VTLM 2.0: Independent Practice]",
].join("\n");

const NOTES_DISCUSSION = [
  "SAY:",
  "- Let's talk together. Why might an I Statement keep a conversation OPEN, when 'You always...' shuts it down?",
  "- When is deep breathing more useful than words?",
  "- Who is one trusted person you could go to when a feeling is too big?",
  "",
  "DO:",
  "- Take 3-4 student responses warmly.",
  "- Acknowledge each one. Do not judge.",
  "- Drop in: 'Expressing a feeling well doesn't make it go away -- it makes it easier for someone else to help.'",
  "",
  "TEACHER NOTES:",
  "A short closing discussion -- 2-3 minutes. Keep it warm. The point is to surface WHY these strategies work in real life.",
  "",
  "WATCH FOR:",
  "- One or two voices dominating -- bring in a quieter student gently: 'What do you think?'",
  "- A student who says they don't have a trusted person -- note for a private follow-up.",
  "",
  "[Wellbeing: Discussion | VTLM 2.0: Engaging Learners]",
].join("\n");

const NOTES_CLOSING = [
  "SAY:",
  "- Let's check our success criteria. Thumbs up, sideways, or down for each one.",
  "- I can name three ways to express my feelings safely.",
  "- I can use an I Statement to say how I feel.",
  "- I can choose a strategy that fits the situation.",
  "- Reflection: share one strategy you'll try this week.",
  "",
  "DO:",
  "- Read each criterion aloud; pause for thumbs.",
  "- Take 3-4 quick shares. Acknowledge each warmly.",
  "- Close with: 'Be like Harry. Notice. Name it. Say it kindly.'",
  "",
  "TEACHER NOTES:",
  "Keep the close warm and short. Note any sideways or down thumbs for follow-up.",
  "",
  "WATCH FOR:",
  "- Sideways or down on any criterion -- note for follow-up next session.",
  "- A student who lingers -- check in quietly after class.",
  "",
  "[Wellbeing: Closing | VTLM 2.0: Review and Reflect]",
].join("\n");

// ============================================================
// Helper: draw Harry the Helpful Hamster
// (same character as Session 1/2 for visual continuity)
// ============================================================

function drawHarry(s, cx, cy, scale) {
  const sc = scale || 1;
  s.addShape("ellipse", {
    x: cx - 1.0 * sc, y: cy - 0.20 * sc, w: 2.0 * sc, h: 1.7 * sc,
    fill: { color: "C9A37A" },
    line: { color: "8B6F4E", width: 1.5 },
  });
  s.addShape("ellipse", {
    x: cx - 0.55 * sc, y: cy + 0.25 * sc, w: 1.1 * sc, h: 1.0 * sc,
    fill: { color: "F4E0C5" },
    line: { color: "F4E0C5", width: 0.5 },
  });
  s.addShape("ellipse", {
    x: cx - 0.85 * sc, y: cy - 1.10 * sc, w: 1.7 * sc, h: 1.4 * sc,
    fill: { color: "C9A37A" },
    line: { color: "8B6F4E", width: 1.5 },
  });
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
  const eyeY = cy - 0.78 * sc;
  const lensR = 0.32 * sc;
  s.addShape("ellipse", {
    x: cx - 0.55 * sc, y: eyeY - lensR / 2, w: lensR, h: lensR,
    fill: { color: "FFFFFF" },
    line: { color: "1A1A1A", width: 2.5 },
  });
  s.addShape("ellipse", {
    x: cx + 0.23 * sc, y: eyeY - lensR / 2, w: lensR, h: lensR,
    fill: { color: "FFFFFF" },
    line: { color: "1A1A1A", width: 2.5 },
  });
  s.addShape("line", {
    x: cx - 0.23 * sc, y: eyeY - 0.02, w: 0.46 * sc, h: 0,
    line: { color: "1A1A1A", width: 2.5 },
  });
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
  s.addShape("ellipse", {
    x: cx - 0.07 * sc, y: cy - 0.40 * sc, w: 0.16 * sc, h: 0.12 * sc,
    fill: { color: "8B3A3A" },
    line: { color: "8B3A3A", width: 0.5 },
  });
  s.addShape("line", {
    x: cx - 0.18 * sc, y: cy - 0.22 * sc, w: 0.36 * sc, h: 0,
    line: { color: "1A1A1A", width: 2 },
  });
}

// ============================================================
// Helper: draw a small strategy icon (purely decorative,
// reinforces what the strategy is about visually)
// ============================================================

function drawStrategyIcon(s, kind, cx, cy, size, color) {
  const r = size / 2;
  // Coloured disc background
  s.addShape("roundRect", {
    x: cx - r, y: cy - r, w: size, h: size, rectRadius: r,
    fill: { color },
    line: { color, width: 0.5 },
  });

  if (kind === "I") {
    // Speech-bubble icon: white rounded rect + a small triangle tail at the bottom
    const bubW = size * 0.62;
    const bubH = size * 0.42;
    s.addShape("roundRect", {
      x: cx - bubW / 2, y: cy - bubH / 2 - size * 0.06, w: bubW, h: bubH, rectRadius: bubH * 0.30,
      fill: { color: "FFFFFF" },
      line: { color: "FFFFFF", width: 0.5 },
    });
    s.addText("I", {
      x: cx - bubW / 2, y: cy - bubH / 2 - size * 0.06, w: bubW, h: bubH,
      fontSize: 18, fontFace: FONT_H, color, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
  } else if (kind === "Breathe") {
    // Concentric circles -- breathing
    s.addShape("ellipse", {
      x: cx - size * 0.30, y: cy - size * 0.30, w: size * 0.60, h: size * 0.60,
      fill: { color: "FFFFFF" },
      line: { color: "FFFFFF", width: 0.5 },
    });
    s.addShape("ellipse", {
      x: cx - size * 0.18, y: cy - size * 0.18, w: size * 0.36, h: size * 0.36,
      fill: { color },
      line: { color, width: 0.5 },
    });
    s.addShape("ellipse", {
      x: cx - size * 0.08, y: cy - size * 0.08, w: size * 0.16, h: size * 0.16,
      fill: { color: "FFFFFF" },
      line: { color: "FFFFFF", width: 0.5 },
    });
  } else if (kind === "Help") {
    // White "?" inside the disc -- ask
    s.addText("?", {
      x: cx - r, y: cy - r, w: size, h: size,
      fontSize: 28, fontFace: FONT_H, color: "FFFFFF", bold: true,
      align: "center", valign: "middle", margin: 0,
    });
  }
}

// ============================================================
// Build PPTX
// ============================================================

async function build() {
  const pres = new PptxGenJS();
  pres.layout = "LAYOUT_16x9";

  // -- Slide 1: Title --
  titleSlide(
    pres,
    "Expressing Emotions\nLike Harry",
    "Harry the Helpful Hamster",
    "Year 6 Wellbeing  |  Session 3  |  15-20 minutes",
    NOTES_TITLE
  );

  // -- Slide 2: Teacher Resources (immediately after title) --
  addResourceSlide(pres, {
    resources: RESOURCES,
    manipulatives: [
      "Harry's Emotions Wheel chart from Session 2 (optional reference at the front)",
      "Role-Play Cards (printed, cut up, one stack per pair or small group)",
    ],
    studentTools: [
      "Mini-whiteboards and markers, OR student books/paper",
      "A book or paper for the personal reflection",
    ],
    boardSetup: [
      "Display the I Statement frame large at the front: I FEEL ___ WHEN ___ BECAUSE ___",
      "Strategy Guide on the teacher's desk -- not for students",
    ],
  }, T, FOOTER, NOTES_RESOURCES);

  // -- Slide 3: Launch -- quick recall check-in --
  contentSlide(
    pres,
    "Launch",
    C.SECONDARY,
    "From Naming to Expressing",
    [
      "Last time: we NAMED feelings on Harry's wheel",
      "Today: we EXPRESS feelings in a way that helps",
      "Quick check on your fingers:",
      "1 = I remember the wheel",
      "2 = some have slipped",
      "3 = I'd like a reminder",
    ],
    NOTES_LAUNCH,
    FOOTER,
    (s, lg) => {
      const rX = lg.rightX;
      const rW = lg.rightW;
      const topY = lg.panelTopPadded;
      const cardH = 3.5;

      addCard(s, rX, topY, rW, cardH, { strip: C.SECONDARY, fill: C.WHITE });
      const harryCx = rX + rW / 2;
      const harryCy = topY + cardH / 2 + 0.10;
      drawHarry(s, harryCx, harryCy, 0.95);
    }
  );

  // -- Slide 4: LI / SC --
  liSlide(
    pres,
    ["We are learning to express our feelings in a healthy way using Harry's three strategies"],
    [
      "I can name three ways to express my feelings safely",
      "I can use an I Statement to say how I feel",
      "I can choose a strategy that fits the situation",
    ],
    NOTES_LI,
    FOOTER
  );

  // -- Slide 5: Three Strategies overview (visual anchor) --
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "Strategies", { color: C.PRIMARY });
    addTitle(s, "Harry's Three Strategies");

    const startY = CONTENT_TOP + 0.10;
    const availH = SAFE_BOTTOM - startY - 0.10;
    const cardW = (9 - 0.30 - 0.30) / 3;
    const cardH = availH;
    const xs = [0.5, 0.5 + cardW + 0.30, 0.5 + (cardW + 0.30) * 2];
    const strategies = [
      {
        kind: "I",
        color: STRAT_I.color,
        name: "I Statement",
        line1: "Say how YOU feel",
        line2: "I feel ___ when ___ because ___",
      },
      {
        kind: "Breathe",
        color: STRAT_BR.color,
        name: "Deep Breathing",
        line1: "Calm your body first",
        line2: "In for 4. Hold for 4. Out for 4.",
      },
      {
        kind: "Help",
        color: STRAT_HELP.color,
        name: "Seek Help",
        line1: "Ask a trusted person",
        line2: "Teacher, family, friend, school counsellor.",
      },
    ];

    strategies.forEach((st, i) => {
      const x = xs[i];
      const y = startY;

      // Card backdrop -- white card with coloured strip
      addCard(s, x, y, cardW, cardH, { strip: st.color, fill: C.WHITE });

      // Coloured header band inside the card
      const headerH = 0.55;
      s.addShape("roundRect", {
        x: x + 0.15, y: y + 0.18, w: cardW - 0.30, h: headerH, rectRadius: 0.08,
        fill: { color: st.color },
      });
      s.addText(st.name, {
        x: x + 0.15, y: y + 0.18, w: cardW - 0.30, h: headerH,
        fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      // Icon disc beneath the header
      const iconCy = y + 0.18 + headerH + 0.55;
      drawStrategyIcon(s, st.kind, x + cardW / 2, iconCy, 0.85, st.color);

      // Two short body lines beneath the icon
      const textY = iconCy + 0.55;
      const textH = cardH - (textY - y) - 0.20;
      s.addText([
        { text: st.line1, options: { fontSize: 14, fontFace: FONT_H, color: C.CHARCOAL, bold: true, breakLine: true } },
        { text: " ", options: { fontSize: 4, breakLine: true } },
        { text: st.line2, options: { fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, italic: true } },
      ], {
        x: x + 0.20, y: textY, w: cardW - 0.40, h: textH,
        valign: "top", align: "center", margin: 0,
        fit: "shrink", shrinkText: true,
      });
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_STRATEGIES);
  }

  // -- Slide 6: I Do -- model an I Statement with the frame --
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "I Do", { color: C.PRIMARY });
    addTitle(s, "Watch Harry Build an I Statement");

    // Hero situation card
    const sitY = CONTENT_TOP + 0.10;
    const sitH = 1.10;
    addCard(s, 0.5, sitY, 9, sitH, { strip: C.PRIMARY, fill: C.WHITE });
    s.addText("Situation: My partner keeps interrupting me while I'm sharing an idea.", {
      x: 0.85, y: sitY + 0.10, w: 8.3, h: sitH - 0.20,
      fontSize: 18, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
      align: "center", valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    // Frame card -- the hero teaching object
    const frY = sitY + sitH + 0.20;
    const frH = SAFE_BOTTOM - frY - 0.10;
    addCard(s, 0.5, frY, 9, frH, { strip: C.ACCENT, fill: C.BG_LIGHT });

    s.addText("I FEEL    when    because", {
      // Placeholder -- we'll redraw below with parts
      x: 0, y: 0, w: 0.01, h: 0.01,
      fontSize: 1, color: C.BG_LIGHT,
    });

    // Three pill-and-blank pairs across the row -- visual sentence frame
    const partY = frY + 0.18;
    const partH = 0.55;
    const pillW = 1.20;
    const blankW = 1.55;
    const gap = 0.18;
    const rowW = pillW * 3 + blankW * 3 + gap * 5; // 3 pills + 3 blanks + 5 gaps
    const startX = 0.5 + (9 - rowW) / 2;

    const parts = [
      { label: "I feel",  blank: "FRUSTRATED" },
      { label: "when",    blank: "you interrupt me" },
      { label: "because", blank: "I lose my idea" },
    ];

    let cx = startX;
    parts.forEach((p) => {
      // Label pill
      s.addShape("roundRect", {
        x: cx, y: partY, w: pillW, h: partH, rectRadius: 0.10,
        fill: { color: C.ACCENT },
      });
      s.addText(p.label, {
        x: cx, y: partY, w: pillW, h: partH,
        fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      cx += pillW + gap;

      // Filled-in blank
      s.addShape("roundRect", {
        x: cx, y: partY, w: blankW, h: partH, rectRadius: 0.10,
        fill: { color: C.WHITE },
        line: { color: C.ACCENT, width: 1 },
      });
      s.addText(p.blank, {
        x: cx + 0.06, y: partY, w: blankW - 0.12, h: partH,
        fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, italic: true, bold: true,
        align: "center", valign: "middle", margin: 0,
        fit: "shrink", shrinkText: true,
      });
      cx += blankW + gap;
    });

    // Harry's tip line
    const tipY = partY + partH + 0.30;
    const tipH = frH - (tipY - frY) - 0.20;
    s.addText([
      { text: "Harry's tip: ", options: { fontSize: 14, fontFace: FONT_H, color: C.ACCENT, bold: true } },
      { text: "Start with the FEELING, not 'you always...'. That keeps the conversation open.", options: { fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, italic: true } },
    ], {
      x: 0.85, y: tipY, w: 8.3, h: tipH,
      valign: "top", align: "center", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO_ISTAT);
  }

  // -- Slide 7-8: We Do with reveal --
  const buildWeDo = () => {
    const s = pres.addSlide();
    addTopBar(s, C.SECONDARY);
    addBadge(s, "We Do", { color: C.SECONDARY });
    addTitle(s, "Build One Together");

    const cardY = CONTENT_TOP + 0.10;
    const cardH = 2.10;
    addCard(s, 0.5, cardY, 9, cardH, { strip: C.SECONDARY, fill: C.WHITE });
    s.addText("Your group member hasn't done their part of the project, and it's due tomorrow.", {
      x: 0.85, y: cardY + 0.18, w: 8.3, h: cardH - 0.36,
      fontSize: 22, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
      align: "center", valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    // Instruction strip
    const promptY = cardY + cardH + 0.22;
    s.addShape("roundRect", {
      x: 0.5, y: promptY, w: 9, h: 0.55, rectRadius: 0.10,
      fill: { color: C.BG_LIGHT },
      line: { color: C.MUTED, width: 0.5 },
    });
    s.addText("Write an I Statement using Harry's frame: I feel ___ when ___ because ___", {
      x: 0.7, y: promptY, w: 8.6, h: 0.55,
      fontSize: 15, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
      align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_WEDO_REVEAL);
    return s;
  };
  withReveal(buildWeDo, (s) => {
    const barY = 4.40;
    const barH = 0.75;
    s.addShape("roundRect", {
      x: 0.5, y: barY, w: 9, h: barH, rectRadius: 0.12,
      fill: { color: STRAT_I.color },
    });
    s.addText("Harry's pick: 'I feel worried when our part isn't done because I want our project to be our best.'", {
      x: 0.6, y: barY, w: 8.8, h: barH,
      fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });
  });

  // -- Slide 9-10: CFU hinge with reveal --
  const buildCfu = () => {
    const s = pres.addSlide();
    addTopBar(s, C.ALERT);
    addBadge(s, "CFU", { color: C.ALERT });
    addTitle(s, "Which Strategy Fits Best?", { color: C.ALERT });

    // Check stamp
    s.addShape("roundRect", {
      x: 7.80, y: 0.20, w: 1.70, h: 0.42, rectRadius: 0.08,
      fill: { color: C.WHITE },
      line: { color: C.ALERT, width: 1.5 },
    });
    s.addText("CHECK", {
      x: 7.80, y: 0.20, w: 1.70, h: 0.42,
      fontSize: 14, fontFace: FONT_B, color: C.ALERT, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    // Technique pill
    const pillY = CONTENT_TOP;
    s.addShape("roundRect", {
      x: 0.5, y: pillY, w: 3.1, h: 0.46, rectRadius: 0.08,
      fill: { color: C.ALERT },
    });
    s.addText("Show Me Boards", {
      x: 0.5, y: pillY, w: 3.1, h: 0.46,
      fontSize: 14, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    // Hero scenario card
    const qY = pillY + 0.46 + 0.18;
    const qH = 1.70;
    addCard(s, 0.5, qY, 9, qH, { strip: C.ALERT, fill: C.WHITE });
    s.addText("You can feel your face getting hot and your fists clenching after losing an important game.", {
      x: 0.85, y: qY + 0.18, w: 8.3, h: qH - 0.36,
      fontSize: 21, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
      align: "center", valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    // Choice strip below the scenario -- visual chooser
    const choiceY = qY + qH + 0.18;
    const choiceH = 0.55;
    const choiceW = (9 - 0.30) / 3;
    const xs = [0.5, 0.5 + choiceW + 0.15, 0.5 + (choiceW + 0.15) * 2];
    const choices = [
      { name: "I Statement",    color: STRAT_I.color },
      { name: "Deep Breathing", color: STRAT_BR.color },
      { name: "Seek Help",      color: STRAT_HELP.color },
    ];
    choices.forEach((ch, i) => {
      s.addShape("roundRect", {
        x: xs[i], y: choiceY, w: choiceW, h: choiceH, rectRadius: 0.10,
        fill: { color: C.WHITE },
        line: { color: ch.color, width: 1.5 },
      });
      s.addText(ch.name, {
        x: xs[i], y: choiceY, w: choiceW, h: choiceH,
        fontSize: 16, fontFace: FONT_H, color: ch.color, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_CFU);
    return s;
  };
  withReveal(buildCfu, (s) => {
    const barY = 4.40;
    const barH = 0.75;
    s.addShape("roundRect", {
      x: 0.5, y: barY, w: 9, h: barH, rectRadius: 0.12,
      fill: { color: STRAT_BR.color },
    });
    s.addText("Harry's pick: DEEP BREATHING first  ->  then words", {
      x: 0.5, y: barY, w: 9, h: barH,
      fontSize: 19, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
  });

  // -- Slide 11: You Do -- role-play cards in pairs --
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "You Do", { color: C.PRIMARY });
    addTitle(s, "Role-Play with Your Partner");

    // Small instruction strip
    const stripY = CONTENT_TOP + 0.05;
    const stripH = 0.55;
    s.addShape("roundRect", {
      x: 0.5, y: stripY, w: 9, h: stripH, rectRadius: 0.10,
      fill: { color: C.BG_LIGHT },
      line: { color: C.MUTED, width: 0.5 },
    });
    s.addText("Take a card. Pick a strategy. PRACTISE it out loud with your partner.", {
      x: 0.7, y: stripY, w: 8.6, h: stripH,
      fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
      align: "center", valign: "middle", margin: 0,
    });

    // Three step cards
    const startY = stripY + stripH + 0.20;
    const availH = SAFE_BOTTOM - startY - 0.05;
    const gap = 0.15;
    const cardH = (availH - gap * 2) / 3;
    const steps = [
      { label: "First", text: "Read the card together. What's happening? How might this person feel?", color: C.SECONDARY },
      { label: "Next",  text: "Choose a strategy: I Statement, Deep Breathing, or Seek Help.",        color: C.PRIMARY   },
      { label: "Then",  text: "Practise it OUT LOUD. Say the words. Do the breathing. Name who you'd ask.", color: C.ACCENT    },
    ];

    steps.forEach((step, i) => {
      const y = startY + i * (cardH + gap);
      addCard(s, 0.5, y, 9, cardH, { strip: step.color, fill: C.WHITE });
      s.addShape("roundRect", {
        x: 0.85, y: y + 0.15, w: 1.3, h: cardH - 0.30, rectRadius: 0.06,
        fill: { color: step.color },
      });
      s.addText(step.label, {
        x: 0.85, y: y + 0.15, w: 1.3, h: cardH - 0.30,
        fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      s.addText(step.text, {
        x: 2.30, y: y + 0.10, w: 6.60, h: cardH - 0.20,
        fontSize: 17, fontFace: FONT_B, color: C.CHARCOAL,
        align: "left", valign: "middle", margin: 0,
        fit: "shrink", shrinkText: true,
      });
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  }

  // -- Slide 12: Personal Reflection --
  reflectionSlide(
    pres,
    "Your Own I Statement",
    [
      "Think of a real moment recently when you felt a strong feeling.",
      "Write an I Statement: I feel ___ when ___ because ___.",
      "Name one other strategy that could also have helped.",
    ],
    NOTES_REFLECTION,
    FOOTER
  );

  // -- Slide 13: Class discussion --
  pairShareSlide(
    pres,
    "Why These Strategies Work",
    [
      "Why might an I Statement keep a conversation open, when 'You always...' shuts it down?",
      "When is deep breathing more useful than words?",
      "Who is one trusted person you could go to when a feeling is too big?",
    ],
    NOTES_DISCUSSION,
    FOOTER
  );

  // -- Slide 14: Closing --
  closingSlide(
    pres,
    {
      reflectionPrompt: "Share one strategy you'll try this week. Be like Harry: notice, name it, say it kindly.",
      scItems: [
        "I can name three ways to express my feelings safely",
        "I can use an I Statement to say how I feel",
        "I can choose a strategy that fits the situation",
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
}

// ============================================================
// Build Role-Play Cards PDF
// ============================================================

function drawRolePlayCard(doc, x, y, w, h, scenario) {
  const accent = "#1B5E3F";
  doc.save();
  doc.roundedRect(x, y, w, h, 8).lineWidth(1.2).strokeColor(accent).stroke();
  doc.rect(x, y, w, 8).fill(accent);
  doc.restore();

  doc.font("Sans-Bold").fontSize(10).fillColor(accent);
  doc.text("Role-Play Card", x + 12, y + 14, { width: w - 24 });

  doc.font("Sans").fontSize(13).fillColor("#1F2937");
  doc.text(scenario, x + 14, y + 34, { width: w - 28, lineGap: 2 });

  doc.font("Sans-Italic").fontSize(9).fillColor("#6B7280");
  doc.text("Choose a strategy. Practise it out loud with your partner.", x + 14, y + h - 22, { width: w - 28 });
}

async function buildRolePlayCardsPdf() {
  const doc = createPdf({ title: "Session 3 Role-Play Cards" });
  const filePath = path.join(LESSON_FOLDER, ROLE_PLAY_RES.fileName);

  let y = addPdfHeader(doc, "Session 3 Role-Play Cards", {
    subtitle: "Expressing Emotions Like Harry | Year 6 Wellbeing | Print, cut, and distribute to pairs.",
    color: "1B5E3F",
    showNameDate: false,
  });

  const margin = 50;
  const contentW = 595.28 - margin * 2;
  const colW = (contentW - 18) / 2;
  const colGap = 18;
  const rowGap = 14;
  const rowH = 165;

  SCENARIOS.forEach((sc, i) => {
    const inPage = i % 6;
    const row = Math.floor(inPage / 2);
    const col = inPage % 2;

    if (i > 0 && inPage === 0) {
      doc.addPage();
      y = addPdfHeader(doc, "Session 3 Role-Play Cards (cont.)", {
        subtitle: "Expressing Emotions Like Harry | Year 6 Wellbeing",
        color: "1B5E3F",
        showNameDate: false,
      });
    }

    const cardX = margin + col * (colW + colGap);
    const cardY = y + row * (rowH + rowGap);
    drawRolePlayCard(doc, cardX, cardY, colW, rowH, sc.text);
  });

  addPdfFooter(doc, "Harry the Helpful Hamster | Year 6 Wellbeing | Session 3 | Role-Play Cards");

  await writePdf(doc, filePath);
  console.log("PDF written to", filePath);
}

// ============================================================
// Build Strategy Guide (teacher answer key) PDF
// ============================================================

async function buildStrategyGuidePdf() {
  const doc = createPdf({ title: "Session 3 Strategy Guide" });
  const filePath = path.join(LESSON_FOLDER, STRATEGY_GUIDE_RES.fileName);

  let y = addPdfHeader(doc, "Session 3 Strategy Guide", {
    subtitle: "Teacher reference. Real moments often need more than one strategy -- accept reasoned alternatives.",
    color: "1B5E3F",
    showNameDate: false,
  });

  y = addBodyText(doc,
    "Listen for the WHY behind the choice. A student who picks Deep Breathing for a conflict scenario is right, as long as they then say 'and then I'd use an I Statement'. The empathy move is naming the order of strategies, not getting one 'correct' answer.",
    y,
    { italic: true, color: "6B7280", fontSize: 10 }
  );
  y += 6;

  // Two-column rows: strategy chip | scenario
  const margin = 50;
  const contentW = 595.28 - margin * 2;
  const stratColW = 110;
  const scenColW = contentW - stratColW - 10;
  const rowGap = 4;

  doc.font("Sans-Bold").fontSize(11).fillColor("#1B5E3F");
  doc.text("Primary strategy", margin, y);
  doc.text("Scenario", margin + stratColW + 10, y);
  y += 18;
  doc.save();
  doc.moveTo(margin, y - 2).lineTo(margin + contentW, y - 2).lineWidth(0.6).strokeColor("#1B5E3F").stroke();
  doc.restore();
  y += 4;

  SCENARIOS.forEach((sc) => {
    doc.font("Sans").fontSize(10);
    const scenH = doc.heightOfString(sc.text, { width: scenColW, lineGap: 1.5 });
    const rowH = Math.max(28, scenH + 14);

    if (y + rowH > 800) {
      doc.addPage();
      y = addPdfHeader(doc, "Session 3 Strategy Guide (cont.)", {
        subtitle: "Teacher reference.",
        color: "1B5E3F",
        showNameDate: false,
      });
    }

    // Strategy chip
    const chipColor = "#" + sc.strategy.color;
    doc.save();
    doc.roundedRect(margin, y, stratColW, 22, 4).fill(chipColor);
    doc.font("Sans-Bold").fontSize(10).fillColor("#FFFFFF");
    doc.text(sc.strategy.name, margin, y + 7, { width: stratColW, align: "center" });
    doc.restore();

    doc.font("Sans").fontSize(10).fillColor("#1F2937");
    doc.text(sc.text, margin + stratColW + 10, y + 4, { width: scenColW, lineGap: 1.5 });

    y += rowH + rowGap;
  });

  addPdfFooter(doc, "Harry the Helpful Hamster | Year 6 Wellbeing | Session 3 | Strategy Guide");

  await writePdf(doc, filePath);
  console.log("PDF written to", filePath);
}

// ============================================================
// Run
// ============================================================

(async () => {
  try {
    fs.mkdirSync(LESSON_FOLDER, { recursive: true });
    await build();
    await buildRolePlayCardsPdf();
    await buildStrategyGuidePdf();
    console.log("\nBuild complete!");
    console.log("Output folder:", LESSON_FOLDER);
  } catch (err) {
    console.error(err);
    process.exit(1);
  }
})();
