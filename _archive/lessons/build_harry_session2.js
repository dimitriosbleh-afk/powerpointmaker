"use strict";

// Harry the Helpful Hamster - Session 2
// Year 6 Wellbeing - 15-20 minute session
// Navigating Emotions with Harry's Wheel

const path = require("path");
const fs = require("fs");
const PptxGenJS = require("pptxgenjs");
const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addPdfFooter, addBodyText,
  addResourceSlide, makeSessionResource,
} = require("../themes/pdf_helpers");

const T = createTheme("wellbeing", "grade56", weekToVariant(2));
const {
  C, FONT_H, FONT_B,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  pairShareSlide, reflectionSlide,
  addTopBar, addBadge, addTitle, addCard, addFooter,
  withReveal,
  CONTENT_TOP, SAFE_BOTTOM,
} = T;

const UNIT = "Harry_the_Helpful_Hamster_Session2";
const LESSON_FOLDER = path.join(__dirname, "..", "output", UNIT);
const PPTX_NAME = "Navigating Emotions with Harrys Wheel.pptx";
const FOOTER = "Harry the Helpful Hamster | Year 6 Wellbeing | Session 2";
const SESSION_NUMBER = 2;

// 6 primary emotion segments — the wheel uses the exact 6 categories
// referenced in the supplied scenario card answers.
const WHEEL_EMOTIONS = [
  { name: "Happy",    color: "F2A93B" }, // warm gold
  { name: "Powerful", color: "8E3B7A" }, // strong purple
  { name: "Peaceful", color: "4E8E6E" }, // calm green
  { name: "Sad",      color: "3D6FA8" }, // soft blue
  { name: "Scared",   color: "6B5DA0" }, // muted violet
  { name: "Mad",      color: "C24A2C" }, // brick red
];

// 12 scenarios supplied by user, with matching emotion.
const SCENARIOS = [
  { text: "Your best friend is moving away to another city, and you won't be able to see them as often anymore.",         emotion: "Sad" },
  { text: "You're lying in bed at night and hear a loud, unfamiliar noise coming from downstairs.",                       emotion: "Scared" },
  { text: "Someone cut in front of you in the canteen line, and you had been waiting patiently for a long time.",          emotion: "Mad" },
  { text: "You learned a new skill that you've been practising for a while, and you feel proud of your accomplishment.",   emotion: "Powerful" },
  { text: "You stood up to a bully who was picking on your friend, and they apologised and left your friend alone.",       emotion: "Powerful" },
  { text: "You're about to give a presentation in front of the whole class, and you're worried about making a mistake.",   emotion: "Scared" },
  { text: "You just found out that you got a high score on a test you worked really hard for.",                            emotion: "Happy" },
  { text: "Your sibling borrowed your favourite toy without asking and accidentally broke it.",                            emotion: "Mad" },
  { text: "You're sitting by a quiet lake, watching the sunset and listening to the gentle sounds of nature around you.",  emotion: "Peaceful" },
  { text: "You just finished a big project you've been working on, and you feel a sense of calm and satisfaction.",        emotion: "Peaceful" },
  { text: "You didn't get the lead role in the school concert that you really wanted and had practised hard for.",         emotion: "Sad" },
  { text: "It's your birthday, and your family surprises you with a party and all your favourite foods.",                  emotion: "Happy" },
];

function emotionColor(name) {
  const e = WHEEL_EMOTIONS.find((w) => w.name === name);
  return e ? e.color : C.PRIMARY;
}

// ============================================================
// Resources (printable PDFs)
// ============================================================

const SCENARIO_CARDS_RES = makeSessionResource(
  SESSION_NUMBER,
  "Scenario Cards",
  "Print, cut along the lines, and distribute to pairs or small groups."
);
const SCENARIO_ANSWERS_RES = makeSessionResource(
  SESSION_NUMBER,
  "Scenario Card Answer Key",
  "Teacher reference. Use to guide pair-share discussion."
);

const RESOURCES = [SCENARIO_CARDS_RES, SCENARIO_ANSWERS_RES];

// ============================================================
// Teacher notes
// ============================================================

const NOTES_TITLE = [
  "SAY:",
  "- Welcome back, Harry's helpers.",
  "- Today we'll use Harry's Emotions Wheel to navigate feelings -- our own and other people's.",
  "- This is a short 15-20 minute session.",
  "",
  "DO:",
  "- Display the title as students settle.",
  "- Have the printed Scenario Cards cut and ready in stacks (about one set per 3-4 students).",
  "- Have the Answer Key on your desk for quick reference.",
  "",
  "TEACHER NOTES:",
  "Keep the tone warm and inviting. The aim is for every student to feel safe naming and discussing emotions.",
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
  "- Make sure each pair or small group has a stack of Scenario Cards.",
  "- Display Harry's Emotions Wheel large at the front, or give each student a small wheel handout.",
  "- Have the Answer Key on your desk -- not for students.",
  "- Have paper or workbooks ready for the personal reflection.",
  "",
  "TEACHER NOTES:",
  "Two printable PDFs are included: Scenario Cards (cut up before the lesson) and an Answer Key (teacher use only). No worksheets -- students record reflections in their normal book or on a half sheet.",
  "",
  "WATCH FOR:",
  "- Every group has at least one set of Scenario Cards before the activity begins.",
  "",
  "[Wellbeing: Resources | VTLM 2.0: Preparation]",
].join("\n");

const NOTES_LAUNCH = [
  "SAY:",
  "- Last time we met Harry. Harry notices feelings.",
  "- Quick check-in: hold up one finger if you're feeling calm, two if you're feeling alert, three if you're feeling a bit unsettled.",
  "- Some of you may remember Harry's wheel from last time. If this feels new, that's okay -- we'll build it together.",
  "- Today Harry will help us look at situations and ask: what might someone be feeling here?",
  "",
  "DO:",
  "- Take a brisk finger check-in scan -- don't single anyone out.",
  "- Acknowledge: 'Lots of different feelings in this room. All of them are okay.'",
  "- Link forward: 'Today we'll use Harry's wheel to put names to the feelings.'",
  "",
  "TEACHER NOTES:",
  "Keep the launch short -- about 90 seconds. The check-in primes students to notice their own state before discussing others'.",
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
  "- A student unsure about 'empathy' -- a quick paraphrase: 'imagining how someone else feels'.",
  "",
  "[Wellbeing: LI/SC | VTLM 2.0: Learning Intentions]",
].join("\n");

const NOTES_WHEEL = [
  "SAY:",
  "- This is Harry's Emotions Wheel.",
  "- Six big feelings: Happy, Powerful, Peaceful, Sad, Scared, Mad.",
  "- Each one is a doorway. Inside the doorway are sharper, more exact feeling words.",
  "- For example, inside MAD you might find frustrated, irritated, jealous.",
  "- Inside HAPPY you might find proud, excited, grateful.",
  "- Harry's tip: start with the big zone, then narrow to the exact word.",
  "",
  "DO:",
  "- Point to each segment as you name it.",
  "- Pause briefly on each emotion and give one example word inside.",
  "- Keep the pace warm -- this is revision, not a long lecture.",
  "",
  "TEACHER NOTES:",
  "Revising the wheel takes about 2-3 minutes. The wheel itself is the visual anchor for the rest of the lesson. Reinforce that all emotions on the wheel are natural and valid.",
  "",
  "WATCH FOR:",
  "- Students who think Mad and Scared are 'bad' -- gently challenge: 'Feelings aren't good or bad. They're information.'",
  "",
  "[Wellbeing: I Do | VTLM 2.0: Explicit Modelling]",
].join("\n");

const NOTES_IDO_MODEL = [
  "SAY:",
  "- Watch how I use the wheel with a situation.",
  "- Situation: 'You waited a long time for a parcel and it didn't come.'",
  "- Step 1 -- what's the BIG feeling? It's not Happy or Peaceful. Closer to Mad.",
  "- Step 2 -- now narrow down. Was I full-on angry? Not quite. I think the sharper word is FRUSTRATED.",
  "- Naming it frustrated helps. It's an exact word -- it tells me, and others, what kind of mad I am.",
  "- That's the move: big zone first, then the exact word inside.",
  "",
  "DO:",
  "- Point to MAD on the wheel as you narrow down.",
  "- Make your thinking audible -- this is the move students will copy.",
  "- Keep the example low-stakes; this is modelling, not personal sharing.",
  "",
  "TEACHER NOTES:",
  "The think-aloud is the lesson. Use a real, low-stakes example from your own week if you want. The move is 'big zone then exact word'.",
  "",
  "WATCH FOR:",
  "- Don't rush the narrowing. The thinking move is the lesson here, not the answer.",
  "",
  "[Wellbeing: I Do | VTLM 2.0: Explicit Modelling]",
].join("\n");

const NOTES_WEDO = [
  "SAY:",
  "- Now let's try one together.",
  "- Read the scenario on the slide with me.",
  "- On your wheel, find the big zone first. Where does this feeling sit?",
  "- On your mini-whiteboard or in your book, write the wheel word you choose.",
  "- Hold up your boards when you're ready.",
  "",
  "DO:",
  "- Read the scenario aloud twice.",
  "- Give 20-30 seconds quiet thinking.",
  "- Say: 'Show me your wheel word.'",
  "- Scan boards. Pick 2-3 different responses to share aloud.",
  "- Affirm the range: 'Different students can pick different words. The wheel helps us find OUR exact one.'",
  "- Click to reveal one strong answer.",
].join("\n");

const NOTES_WEDO_REVEAL = [
  NOTES_WEDO,
  "",
  "CFU CHECKPOINT:",
  "Technique: Show Me Boards.",
  "Script:",
  "- Read the scenario, then say: 'Show me your wheel word.'",
  "- Scan for: a wheel word (Scared, Mad, Sad, Powerful, Happy, Peaceful) or a sharper word inside one of those zones.",
  "PROCEED: If most boards show a wheel-zone word, move on.",
  "PIVOT: If many show 'bad' or 'good' only, prompt: 'Use the wheel. Look INSIDE the big zone. What's a sharper word?' Re-ask.",
  "",
  "TEACHER NOTES:",
  "There isn't one right answer for empathy. The wheel makes it safe to land somewhere different.",
  "",
  "WATCH FOR:",
  "- Students writing two words -- excellent. Feelings often mix.",
  "",
  "[Wellbeing: We Do | VTLM 2.0: Guided Practice]",
].join("\n");

const NOTES_CFU_HINGE = [
  "SAY:",
  "- Quick check. Listen carefully.",
  "- Scenario: 'You stood up to a bully who was picking on your friend, and they apologised and left your friend alone.'",
  "- Use your wheel. Which BIG zone fits best? Write it on your whiteboard.",
  "",
  "DO:",
  "- Read the scenario twice.",
  "- Allow 15 seconds thinking.",
  "- Scan all boards before discussing.",
  "- Reveal: Powerful is the strongest fit -- you used your strength to protect someone.",
  "- Honour other answers: 'Some of you wrote Happy or Peaceful -- those can sit alongside Powerful.'",
  "",
  "CFU CHECKPOINT:",
  "Technique: Show Me Boards (hinge).",
  "Script:",
  "- Read scenario, wait for boards up.",
  "- Scan for: Powerful (best fit). Happy / Peaceful are also defensible -- accept with a brief 'why?'.",
  "PROCEED: If most boards show Powerful or a related strong-zone word, move on to the scenario cards.",
  "PIVOT: If many show Mad, prompt: 'Mad is what we feel WHEN we stand up. AFTER, when they leave us alone, what's the feeling?' Re-ask.",
  "",
  "TEACHER NOTES:",
  "Powerful is the cleanest fit for the answer key, but empathy questions rarely have one answer. Honour reasoned choices.",
  "",
  "WATCH FOR:",
  "- A student picking 'Scared' -- valid! Standing up to a bully can feel scary in the moment. Acknowledge the layered feeling.",
  "",
  "[Wellbeing: CFU | VTLM 2.0: Check For Understanding]",
].join("\n");

const NOTES_YOUDO = [
  "SAY:",
  "- Now you'll do this in pairs or small groups.",
  "- Take a Scenario Card from the stack.",
  "- Read it together. Use Harry's wheel to choose the feeling someone in that situation might have.",
  "- Discuss: WHY did you pick that word? What might help you know?",
  "- Put the card down and take another. Aim to discuss 4-6 cards in the time we have.",
  "",
  "DO:",
  "- Set the timer: about 6-7 minutes for the activity.",
  "- Move between pairs. Don't correct -- ask: 'Tell me why you chose that one.'",
  "- Listen for 'because' -- that's the empathy move.",
  "- Use the Answer Key on your desk to guide tricky cards, but accept reasoned alternatives.",
  "",
  "ENABLING & EXTENDING:",
  "ENABLING PROMPT:",
  "- Task: Give the pair a short menu of THREE wheel words to choose between, rather than the full six.",
  "- Extra Notes: Sentence frame -- 'I think they would feel ___ because ___.'",
  "EXTENDING PROMPT:",
  "- Task: For each card, name TWO emotions that might happen at once, and which is stronger.",
  "- Extra Notes: 'Real feelings often mix. Which one is louder for this person?'",
  "",
  "SENSITIVITY ADVISORY:",
  "- What it is: Some scenarios may bring up real memories (a friend moving away, a broken toy, a missed role).",
  "- Framing language: 'You don't have to share your own story. You're imagining how someone in this situation might feel.'",
  "- Watch for: A student who goes quiet on a card.",
  "- Protocol: Quietly check in after the activity. Follow your school's wellbeing referral process if needed.",
  "",
  "WATCH FOR:",
  "- Pairs racing through without discussing -- redirect: 'Why that word? Convince your partner.'",
  "- Pairs stuck on one card -- it's okay to move on.",
  "",
  "[Wellbeing: You Do | VTLM 2.0: Independent Practice]",
].join("\n");

const NOTES_REFLECTION = [
  "SAY:",
  "- Now think of a time recently when YOU felt a strong emotion.",
  "- Use Harry's wheel to put a name to it.",
  "- In your book, write three short answers:",
  "  1) Which wheel word fits best?",
  "  2) Why did you feel that way?",
  "  3) How did you react -- and could you have responded differently?",
  "- You don't have to share this with anyone. It's for you.",
  "",
  "DO:",
  "- Give about 4-5 minutes of quiet writing.",
  "- Circulate and acknowledge effort without reading over shoulders.",
  "- Remind softly halfway through: 'Use your wheel for the exact word.'",
  "",
  "TEACHER NOTES:",
  "This is private reflection, not a writing task to mark. The point is for students to practise the wheel on themselves. About 4-5 minutes.",
  "",
  "SENSITIVITY ADVISORY:",
  "- What it is: Some students may write about something heavy.",
  "- Framing language: 'This is your private writing. I won't read it unless you'd like me to.'",
  "- Watch for: A student showing signs of distress while writing.",
  "- Protocol: Quietly check in. Follow your school's wellbeing referral process if needed.",
  "",
  "WATCH FOR:",
  "- A student who freezes -- offer the enabling prompt: 'Start with the big zone, then the exact word.'",
  "",
  "[Wellbeing: Personal Reflection | VTLM 2.0: Independent Practice]",
].join("\n");

const NOTES_DISCUSSION = [
  "SAY:",
  "- Let's talk together. How does naming a feeling change what we do with it?",
  "- Did the wheel help you find a sharper word than 'good' or 'bad'?",
  "- How might using Harry's wheel help us understand a friend who's upset?",
  "",
  "DO:",
  "- Take 3-4 student responses warmly.",
  "- Acknowledge each one. Do not judge.",
  "- Drop in: 'When we can name a feeling, we can talk about it. When we can talk about it, we can take care of it -- in ourselves and our friends.'",
  "",
  "TEACHER NOTES:",
  "A short closing discussion -- 2-3 minutes. Keep it warm. The point is to surface the link between naming feelings and looking after them.",
  "",
  "WATCH FOR:",
  "- One or two voices dominating -- bring in a quieter student gently: 'What do you think?'",
  "",
  "[Wellbeing: Discussion | VTLM 2.0: Engaging Learners]",
].join("\n");

const NOTES_CLOSING = [
  "SAY:",
  "- Let's check our success criteria. Thumbs up, sideways, or down for each one.",
  "- I can use Harry's Emotions Wheel to name different emotions.",
  "- I can match an emotion to a situation and explain why.",
  "- I can reflect on a strong emotion I felt and how I responded.",
  "- Reflection: share one thing you learned about your own or someone else's emotions today.",
  "",
  "DO:",
  "- Read each criterion aloud; pause for thumbs.",
  "- Take 3-4 quick shares. Acknowledge each warmly.",
  "- Close with: 'Be like Harry. Notice. Name it. Be kind.'",
  "",
  "TEACHER NOTES:",
  "Connect forward: 'Next time we'll think about how to respond when we notice strong feelings.' Keep the close warm and short.",
  "",
  "WATCH FOR:",
  "- Sideways or down on any criterion -- note for follow-up next session.",
  "- A student who lingers -- check in quietly after class.",
  "",
  "[Wellbeing: Closing | VTLM 2.0: Review and Reflect]",
].join("\n");

// ============================================================
// Helper: draw Harry the Helpful Hamster
// (same character as Session 1 for visual continuity)
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
// Helper: draw Harry's Emotions Wheel (6 segments)
// ============================================================

function drawEmotionsWheel(s, cx, cy, radius) {
  const segments = WHEEL_EMOTIONS.length;
  const innerR = radius * 0.32;

  // Outer ring backdrop
  s.addShape("ellipse", {
    x: cx - radius, y: cy - radius, w: radius * 2, h: radius * 2,
    fill: { color: C.BG_LIGHT },
    line: { color: C.PRIMARY, width: 1.5 },
  });

  // Coloured pill at each segment position
  for (let i = 0; i < segments; i++) {
    const angle = (i / segments) * 2 * Math.PI - Math.PI / 2;
    const labelR = radius * 0.72;
    const lx = cx + Math.cos(angle) * labelR;
    const ly = cy + Math.sin(angle) * labelR;
    const labelW = 1.05;
    const labelH = 0.40;

    s.addShape("roundRect", {
      x: lx - labelW / 2, y: ly - labelH / 2, w: labelW, h: labelH, rectRadius: 0.20,
      fill: { color: WHEEL_EMOTIONS[i].color },
      line: { color: WHEEL_EMOTIONS[i].color, width: 0.5 },
    });
    s.addText(WHEEL_EMOTIONS[i].name, {
      x: lx - labelW / 2, y: ly - labelH / 2, w: labelW, h: labelH,
      fontSize: 12, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
  }

  // Centre disc
  s.addShape("ellipse", {
    x: cx - innerR, y: cy - innerR, w: innerR * 2, h: innerR * 2,
    fill: { color: C.PRIMARY },
    line: { color: C.PRIMARY, width: 1 },
  });
  s.addText("Harry's\nWheel", {
    x: cx - innerR, y: cy - innerR, w: innerR * 2, h: innerR * 2,
    fontSize: 11, fontFace: FONT_H, color: C.WHITE, bold: true,
    align: "center", valign: "middle", margin: 0,
  });
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
    "Navigating Emotions\nwith Harry's Wheel",
    "Harry the Helpful Hamster",
    "Year 6 Wellbeing  |  Session 2  |  15-20 minutes",
    NOTES_TITLE
  );

  // -- Slide 2: Teacher Resources (immediately after title) --
  addResourceSlide(pres, {
    resources: RESOURCES,
    manipulatives: [
      "Harry's Emotions Wheel chart (one large display copy at the front)",
      "Scenario Cards (printed, cut up, one stack per pair or small group)",
    ],
    studentTools: [
      "Mini-whiteboards and markers, OR student books/paper",
      "A book or paper for the personal reflection",
    ],
    boardSetup: [
      "Display Harry's Emotions Wheel large at the front so it stays visible throughout",
      "Answer Key on the teacher's desk -- not for students",
    ],
  }, T, FOOTER, NOTES_RESOURCES);

  // -- Slide 3: Launch -- quick check-in --
  contentSlide(
    pres,
    "Launch",
    C.SECONDARY,
    "How Are You Arriving Today?",
    [
      "1 finger = calm",
      "2 fingers = alert",
      "3 fingers = a bit unsettled",
      "All feelings are okay -- Harry is here to help name them",
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
    ["We are learning to use Harry's Emotions Wheel to name our own and other people's feelings"],
    [
      "I can use Harry's Emotions Wheel to name different emotions",
      "I can match an emotion to a situation and explain why",
      "I can reflect on a strong emotion I felt and how I responded",
    ],
    NOTES_LI,
    FOOTER
  );

  // -- Slide 5: Harry's Wheel (visual anchor) --
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "Revise", { color: C.PRIMARY });
    addTitle(s, "Harry's Emotions Wheel");

    // Right: the wheel (the hero visual)
    const wheelCx = 6.95;
    const wheelCy = (CONTENT_TOP + SAFE_BOTTOM) / 2;
    const wheelRadius = Math.min((SAFE_BOTTOM - CONTENT_TOP) / 2 - 0.05, 1.85);
    drawEmotionsWheel(s, wheelCx, wheelCy, wheelRadius);

    // Left: short legend / explainer
    const cardX = 0.5;
    const cardW = 4.4;
    const cardY = CONTENT_TOP;
    const cardH = SAFE_BOTTOM - CONTENT_TOP;
    addCard(s, cardX, cardY, cardW, cardH, { strip: C.PRIMARY, fill: C.WHITE });
    s.addText("Six big feelings", {
      x: cardX + 0.30, y: cardY + 0.18, w: cardW - 0.40, h: 0.40,
      fontSize: 18, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText([
      { text: "Each one is a doorway.", options: { fontSize: 14, color: C.CHARCOAL, breakLine: true } },
      { text: "Inside the doorway are sharper, more exact feeling words.", options: { fontSize: 14, color: C.CHARCOAL, breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "Harry's tip:", options: { fontSize: 14, color: C.ACCENT, bold: true, breakLine: true } },
      { text: "Big zone first -- then narrow to the exact word.", options: { fontSize: 14, italic: true, color: C.CHARCOAL } },
    ], {
      x: cardX + 0.30, y: cardY + 0.66, w: cardW - 0.40, h: cardH - 0.80,
      fontFace: FONT_B, valign: "top", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_WHEEL);
  }

  // -- Slide 6: I Do -- model how Harry uses the wheel --
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "I Do", { color: C.PRIMARY });
    addTitle(s, "Watch Harry Pick the Exact Word");

    // Hero situation card
    const sitY = CONTENT_TOP + 0.10;
    const sitH = 1.5;
    addCard(s, 0.5, sitY, 9, sitH, { strip: C.PRIMARY, fill: C.WHITE });
    s.addText("Situation: You waited a long time for a parcel and it didn't come.", {
      x: 0.85, y: sitY + 0.18, w: 8.3, h: sitH - 0.36,
      fontSize: 20, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    // Two step cards
    const stepY = sitY + sitH + 0.22;
    const stepH = SAFE_BOTTOM - stepY - 0.10;
    const colW = (9 - 0.20) / 2;

    // Left: Step 1 -- Big zone
    addCard(s, 0.5, stepY, colW, stepH, { strip: emotionColor("Mad"), fill: C.BG_LIGHT });
    s.addText("Step 1 -- Big zone", {
      x: 0.70, y: stepY + 0.18, w: colW - 0.40, h: 0.45,
      fontSize: 16, fontFace: FONT_H, color: emotionColor("Mad"), bold: true,
      align: "left", valign: "middle", margin: 0,
    });
    s.addText("Closer to MAD than Happy.", {
      x: 0.70, y: stepY + 0.65, w: colW - 0.40, h: stepH - 0.80,
      fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL,
      align: "left", valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    // Right: Step 2 -- Exact word
    addCard(s, 0.5 + colW + 0.20, stepY, colW, stepH, { strip: C.ACCENT, fill: C.BG_LIGHT });
    s.addText("Step 2 -- Exact word", {
      x: 0.70 + colW + 0.20, y: stepY + 0.18, w: colW - 0.40, h: 0.45,
      fontSize: 16, fontFace: FONT_H, color: C.ACCENT, bold: true,
      align: "left", valign: "middle", margin: 0,
    });
    s.addText("Not full-on angry. Sharper word: FRUSTRATED.", {
      x: 0.70 + colW + 0.20, y: stepY + 0.65, w: colW - 0.40, h: stepH - 0.80,
      fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL,
      align: "left", valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO_MODEL);
  }

  // -- Slide 7-8: We Do with reveal --
  const buildWeDo = () => {
    const s = pres.addSlide();
    addTopBar(s, C.SECONDARY);
    addBadge(s, "We Do", { color: C.SECONDARY });
    addTitle(s, "Try One Together");

    const cardY = CONTENT_TOP + 0.10;
    const cardH = 2.10;
    addCard(s, 0.5, cardY, 9, cardH, { strip: C.SECONDARY, fill: C.WHITE });
    s.addText("Someone cut in front of you in the canteen line, and you had been waiting patiently for a long time.", {
      x: 0.85, y: cardY + 0.18, w: 8.3, h: cardH - 0.36,
      fontSize: 22, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
      align: "center", valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    // Instruction strip (smaller than the task)
    const promptY = cardY + cardH + 0.22;
    s.addShape("roundRect", {
      x: 0.5, y: promptY, w: 9, h: 0.55, rectRadius: 0.10,
      fill: { color: C.BG_LIGHT },
      line: { color: C.MUTED, width: 0.5 },
    });
    s.addText("Use Harry's wheel. Write the exact word on your whiteboard.", {
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
      fill: { color: emotionColor("Mad") },
    });
    s.addText("Harry's pick: MAD  ->  exact word: ANNOYED or FRUSTRATED", {
      x: 0.5, y: barY, w: 9, h: barH,
      fontSize: 19, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
  });

  // -- Slide 9-10: CFU hinge with reveal --
  const buildCfu = () => {
    const s = pres.addSlide();
    addTopBar(s, C.ALERT);
    addBadge(s, "CFU", { color: C.ALERT });
    addTitle(s, "Which Big Zone Fits?", { color: C.ALERT });

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
    const qY = pillY + 0.46 + 0.20;
    const qH = SAFE_BOTTOM - qY - 0.90; // leave room for reveal bar
    addCard(s, 0.5, qY, 9, qH, { strip: C.ALERT, fill: C.WHITE });
    s.addText("You stood up to a bully who was picking on your friend, and they apologised and left your friend alone.", {
      x: 0.85, y: qY + 0.20, w: 8.3, h: qH - 0.40,
      fontSize: 22, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
      align: "center", valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_CFU_HINGE);
    return s;
  };
  withReveal(buildCfu, (s) => {
    const barY = 4.40;
    const barH = 0.75;
    s.addShape("roundRect", {
      x: 0.5, y: barY, w: 9, h: barH, rectRadius: 0.12,
      fill: { color: emotionColor("Powerful") },
    });
    s.addText("Harry's pick: POWERFUL  ->  exact word: PROUD or BRAVE", {
      x: 0.5, y: barY, w: 9, h: barH,
      fontSize: 19, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
  });

  // -- Slide 11: You Do -- scenario cards in pairs --
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "You Do", { color: C.PRIMARY });
    addTitle(s, "Scenario Cards with Your Partner");

    // Small instruction strip
    const stripY = CONTENT_TOP + 0.05;
    const stripH = 0.55;
    s.addShape("roundRect", {
      x: 0.5, y: stripY, w: 9, h: stripH, rectRadius: 0.10,
      fill: { color: C.BG_LIGHT },
      line: { color: C.MUTED, width: 0.5 },
    });
    s.addText("Take a card. Read it together. Choose a wheel word. Tell your partner WHY.", {
      x: 0.7, y: stripY, w: 8.6, h: stripH,
      fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
      align: "center", valign: "middle", margin: 0,
    });

    // Three step cards (the hero items)
    const startY = stripY + stripH + 0.20;
    const availH = SAFE_BOTTOM - startY - 0.05;
    const gap = 0.15;
    const cardH = (availH - gap * 2) / 3;
    const steps = [
      { label: "First",  text: "Read the card together. What's happening?",           color: C.SECONDARY },
      { label: "Next",   text: "Use Harry's wheel. Find the big zone, then the exact word.", color: C.PRIMARY },
      { label: "Then",   text: "Tell your partner WHY you picked that word.",         color: C.ACCENT },
    ];

    steps.forEach((step, i) => {
      const y = startY + i * (cardH + gap);
      addCard(s, 0.5, y, 9, cardH, { strip: step.color, fill: C.WHITE });
      // Label pill
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
        fontSize: 18, fontFace: FONT_B, color: C.CHARCOAL,
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
    "Your Own Feeling",
    [
      "Think of a time recently when you felt a strong emotion.",
      "Use Harry's wheel to find the exact word.",
      "Write: what was the feeling, why, and how did you react?",
    ],
    NOTES_REFLECTION,
    FOOTER
  );

  // -- Slide 13: Class discussion --
  pairShareSlide(
    pres,
    "Talking About Feelings",
    [
      "How does naming a feeling change what we do with it?",
      "Did Harry's wheel help you find a sharper word than 'good' or 'bad'?",
      "How might Harry's wheel help us understand a friend who's upset?",
    ],
    NOTES_DISCUSSION,
    FOOTER
  );

  // -- Slide 14: Closing --
  closingSlide(
    pres,
    {
      reflectionPrompt: "Share one thing you learned about your own or someone else's emotions today. Be like Harry: notice, name it, be kind.",
      scItems: [
        "I can use Harry's Emotions Wheel to name different emotions",
        "I can match an emotion to a situation and explain why",
        "I can reflect on a strong emotion I felt and how I responded",
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
// Build Scenario Cards PDF
// ============================================================

function drawScenarioCard(doc, x, y, w, h, scenario, opts) {
  const o = opts || {};
  const accent = "#1B5E3F"; // soft wellbeing green border
  doc.save();
  // Card border
  doc.roundedRect(x, y, w, h, 8).lineWidth(1.2).strokeColor(accent).stroke();
  // Top accent strip
  doc.rect(x, y, w, 8).fill(accent);
  doc.restore();

  // Card label
  doc.font("Sans-Bold").fontSize(10).fillColor(accent);
  doc.text("Scenario Card", x + 12, y + 14, { width: w - 24 });

  // Body text
  doc.font("Sans").fontSize(13).fillColor("#1F2937");
  doc.text(scenario, x + 14, y + 34, { width: w - 28, lineGap: 2 });

  // Footer prompt
  doc.font("Sans-Italic").fontSize(9).fillColor("#6B7280");
  doc.text("Use Harry's Wheel. Choose a feeling. Tell your partner why.", x + 14, y + h - 22, { width: w - 28 });
}

async function buildScenarioCardsPdf() {
  const doc = createPdf({ title: "Session 2 Scenario Cards" });
  const filePath = path.join(LESSON_FOLDER, SCENARIO_CARDS_RES.fileName);

  let y = addPdfHeader(doc, "Session 2 Scenario Cards", {
    subtitle: "Harry's Wheel | Year 6 Wellbeing | Print, cut, and distribute to pairs or small groups.",
    color: "1B5E3F",
    showNameDate: false,
  });

  // Card grid: 2 columns x 3 rows = 6 per page, 12 cards = 2 pages
  const margin = 50;
  const contentW = 595.28 - margin * 2;
  const colW = (contentW - 18) / 2;
  const colGap = 18;
  const rowGap = 14;
  const rowH = 165;

  SCENARIOS.forEach((sc, i) => {
    const pageIndex = Math.floor(i / 6);
    const inPage = i % 6;
    const row = Math.floor(inPage / 2);
    const col = inPage % 2;

    if (i > 0 && inPage === 0) {
      doc.addPage();
      y = addPdfHeader(doc, "Session 2 Scenario Cards (cont.)", {
        subtitle: "Harry's Wheel | Year 6 Wellbeing",
        color: "1B5E3F",
        showNameDate: false,
      });
    }

    const cardX = margin + col * (colW + colGap);
    const cardY = y + row * (rowH + rowGap);
    drawScenarioCard(doc, cardX, cardY, colW, rowH, sc.text);
  });

  addPdfFooter(doc, "Harry the Helpful Hamster | Year 6 Wellbeing | Session 2 | Scenario Cards");

  await writePdf(doc, filePath);
  console.log("PDF written to", filePath);
}

// ============================================================
// Build Scenario Card Answer Key PDF
// ============================================================

async function buildAnswerKeyPdf() {
  const doc = createPdf({ title: "Session 2 Scenario Card Answer Key" });
  const filePath = path.join(LESSON_FOLDER, SCENARIO_ANSWERS_RES.fileName);

  let y = addPdfHeader(doc, "Session 2 Scenario Card Answer Key", {
    subtitle: "Teacher reference. Empathy questions rarely have one right answer.",
    color: "1B5E3F",
    showNameDate: false,
  });

  // Intro
  y = addBodyText(doc,
    "Many scenarios hold more than one feeling at once. Listen for students who say 'mostly ___ but also ___'. That is the empathy move.",
    y,
    { italic: true, color: "6B7280", fontSize: 10 }
  );
  y += 6;

  // Two-column rows: emotion | scenario
  const margin = 50;
  const contentW = 595.28 - margin * 2;
  const emoColW = 90;
  const scenColW = contentW - emoColW - 10;
  const rowGap = 4;

  // Column headers
  doc.font("Sans-Bold").fontSize(11).fillColor("#1B5E3F");
  doc.text("Emotion", margin, y);
  doc.text("Scenario", margin + emoColW + 10, y);
  y += 18;
  doc.save();
  doc.moveTo(margin, y - 2).lineTo(margin + contentW, y - 2).lineWidth(0.6).strokeColor("#1B5E3F").stroke();
  doc.restore();
  y += 4;

  SCENARIOS.forEach((sc) => {
    // Measure scenario text height
    doc.font("Sans").fontSize(10);
    const scenH = doc.heightOfString(sc.text, { width: scenColW, lineGap: 1.5 });
    const rowH = Math.max(28, scenH + 14);

    // Page break if needed
    if (y + rowH > 800) {
      doc.addPage();
      y = addPdfHeader(doc, "Session 2 Scenario Card Answer Key (cont.)", {
        subtitle: "Teacher reference.",
        color: "1B5E3F",
        showNameDate: false,
      });
    }

    // Emotion chip
    const chipColor = (function () {
      const e = WHEEL_EMOTIONS.find((w) => w.name === sc.emotion);
      return e ? "#" + e.color : "#1B5E3F";
    })();
    doc.save();
    doc.roundedRect(margin, y, emoColW, 22, 4).fill(chipColor);
    doc.font("Sans-Bold").fontSize(11).fillColor("#FFFFFF");
    doc.text(sc.emotion, margin, y + 6, { width: emoColW, align: "center" });
    doc.restore();

    // Scenario text
    doc.font("Sans").fontSize(10).fillColor("#1F2937");
    doc.text(sc.text, margin + emoColW + 10, y + 4, { width: scenColW, lineGap: 1.5 });

    y += rowH + rowGap;
  });

  addPdfFooter(doc, "Harry the Helpful Hamster | Year 6 Wellbeing | Session 2 | Answer Key");

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
    await buildScenarioCardsPdf();
    await buildAnswerKeyPdf();
    console.log("\nBuild complete!");
    console.log("Output folder:", LESSON_FOLDER);
  } catch (err) {
    console.error(err);
    process.exit(1);
  }
})();
