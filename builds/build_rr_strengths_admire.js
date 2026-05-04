"use strict";

// Respectful Relationships - Strengths I Admire
// Grade 5/6 Wellbeing - 30+ minute session
// Activity 4: Identifying admired character strengths in fictional and real people

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

// Theme — match the unit's existing palette (variant 1, used in Activity 1).
const T = createTheme("wellbeing", "grade56", weekToVariant(1));
const {
  C, FONT_H, FONT_B,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  pairShareSlide,
  addTopBar, addBadge, addTitle, addCard, addFooter,
  CONTENT_TOP, SAFE_BOTTOM,
} = T;

// Output paths
const UNIT = "RR_Strengths_I_Admire";
const LESSON_FOLDER = path.join(__dirname, "..", "output", UNIT);
const PPTX_NAME = "Strengths I Admire.pptx";
const FOOTER = "Respectful Relationships | Grade 5/6 Wellbeing";
const SESSION = 2;
const RES_DIR = path.join(LESSON_FOLDER, getSessionResourceFolder(SESSION));

// Resources
const PLANNER_RESOURCE = makeSessionResource(
  SESSION,
  "Strengths I Admire Planner",
  "Two-part planner: a fictional character you admire, then someone from everyday life. Lists actions and matches them to character strengths."
);
const RESOURCE_ITEMS = [PLANNER_RESOURCE];

fs.mkdirSync(RES_DIR, { recursive: true });

// ============================================================
// Strengths reference list (matches Activity 1 handout wording)
// ============================================================

const STRENGTHS = [
  "Kindness", "Courage", "Honesty", "Fairness",
  "Humour", "Persistence", "Gratitude", "Curiosity",
  "Teamwork", "Leadership", "Forgiveness", "Self-Control",
  "Creativity", "Love of Learning", "Hope", "Caring",
];

// ============================================================
// Teacher Notes
// ============================================================

const NOTES_TITLE = [
  "SAY:",
  "- Today we're thinking about people we admire and the strengths they show.",
  "- We'll look at a fictional character first, then someone from your real life.",
  "",
  "DO:",
  "- Display slide as students settle.",
  "- Have the Strengths I Admire Planner ready to distribute.",
  "- Have the Character Strengths Handout from Session 1 visible (display copy or each student's copy).",
  "",
  "TEACHER NOTES:",
  "This is a 30-minute session that builds directly on the character strengths concept introduced in Session 1. Students apply the strengths language to people they admire.",
  "",
  "WATCH FOR:",
  "- Students transitioning calmly from the previous subject.",
  "",
  "[General: Title | VTLM 2.0: Establishing Purpose]",
].join("\n");

const NOTES_RESOURCES = [
  "SAY:",
  "- Here's what we'll use today.",
  "",
  "DO:",
  "- Print one Strengths I Admire Planner per student before the session.",
  "- Have Session 1 Character Strengths Handout available as a reference.",
  "- Plan board space for two columns: ACTIONS and STRENGTHS.",
  "",
  "TEACHER NOTES:",
  "Students need their Session 1 Character Strengths Handout as a reference list. The 16 strengths are also shown on the launch slide so all students can see them, but the handout gives kid-friendly definitions.",
  "",
  "WATCH FOR:",
  "- Make sure every student has the planner before the You Do slides.",
  "",
  "[General: Resources | VTLM 2.0: Preparation]",
].join("\n");

const NOTES_LAUNCH = [
  "SAY:",
  "- Last session we learned about character strengths -- the qualities inside a person.",
  "- Some of you may remember the 16 strengths on our handout.",
  "- Today we're going to spot those strengths in people we admire.",
  "- Quick warm-up: Turn to your partner. Name one character strength from the list and one tiny example of what it looks like in action.",
  "",
  "DO:",
  "- Give pairs 30-45 seconds.",
  "- Cold call 2-3 pairs to share.",
  "- Record one clean ACTION | STRENGTH pair on the board to model the language.",
  "",
  "TEACHER NOTES:",
  "This launch reactivates the strengths language from Session 1 and bridges into today's focus on admired people. Keep it brisk.",
  "",
  "WATCH FOR:",
  "- Students naming talents (running, drawing) instead of strengths -- redirect: 'That's a talent. What's the quality inside the person?'",
  "",
  "[Wellbeing: Launch | VTLM 2.0: Activating Prior Knowledge]",
].join("\n");

const NOTES_LI = [
  "SAY:",
  "- Read the Learning Intention from the slide.",
  "- Read each Success Criterion aloud -- we'll come back to these at the end.",
  "",
  "DO:",
  "- Point to each SC as you read it.",
  "- Leave the slide visible for 15-20 seconds.",
  "",
  "TEACHER NOTES:",
  "SC1 is the floor (everyone can name an admired person and one thing they do). SC2 is the core target (matching admirable actions to strengths). SC3 extends into the role-model idea.",
  "",
  "WATCH FOR:",
  "- Students unsure what 'admire' means -- offer: 'Someone you look up to. Someone whose qualities you'd like to grow in yourself.'",
  "",
  "[General: LI/SC | VTLM 2.0: Learning Intentions]",
].join("\n");

const NOTES_IDO = [
  "SAY:",
  "- Watch how I do this first.",
  "- I'm going to think of a fictional character I admire. I'll write down what they DO -- their admirable actions.",
  "- Then I'll look at the strengths list and pick the strengths that match each action.",
  "- Example character: a young hero in a story.",
  "- Action 1: stood up to a bully even though they were scared. Strength: COURAGE.",
  "- Action 2: stayed beside an upset friend at lunch. Strength: KINDNESS, CARING.",
  "- Action 3: kept practising even when training was hard. Strength: PERSISTENCE.",
  "",
  "DO:",
  "- Use a character from a class novel or film if you have one. Otherwise use the generic hero example shown.",
  "- Write each ACTION on the board, then ask the class which STRENGTH matches.",
  "- Confirm and write the strength next to the action.",
  "- Make the ACTION-to-STRENGTH move visible -- this is the thinking students will copy.",
  "",
  "TEACHER NOTES:",
  "The think-aloud is the lesson here. Make every step of your thinking audible. Notice that one action can match more than one strength.",
  "",
  "WATCH FOR:",
  "- Don't rush the matching step. The thinking move 'what kind of person does this?' is the part you're modelling.",
  "",
  "[General: I Do | VTLM 2.0: Explicit Modelling]",
].join("\n");

const NOTES_WEDO = [
  "SAY:",
  "- Now we'll do one together. Let's think of a fictional character we all know.",
  "- Choose a character from a class text, film, or story everyone knows.",
  "- Ask: What is one admirable action this character does? [Take 2-3 contributions; record on board.]",
  "- Ask: Which strength does that action show? [Cold call, then check with class.]",
  "- Build a class list: ACTION | STRENGTH.",
  "",
  "DO:",
  "- Pick a character your class has read about, or use a teacher-selected one from the school's list.",
  "- Take contributions on mini-whiteboards: 'Write one action this character did that you admire.'",
  "- Scan whiteboards. Pick 3 strong actions to put on the board.",
  "- Then take strength suggestions for each action; allow more than one strength per action.",
  "",
  "CFU CHECKPOINT:",
  "Technique: Mini-whiteboards.",
  "Script:",
  "- Say: Write one strength this character shows. Hold up your board.",
  "- Scan for: strengths from the list (kindness, courage etc.) rather than talents (running, singing).",
  "PROCEED: If most boards show a strength from the list, move on.",
  "PIVOT: If many boards show a talent or a personality word ('cool', 'funny'), redirect: 'Look at the strengths list. Which one matches what they DID?' Re-ask.",
  "",
  "TEACHER NOTES:",
  "If your class has not yet shared a class character, use a film or story you know everyone has seen at school. The point is shared knowledge, not the specific character.",
  "",
  "ENABLING & EXTENDING:",
  "ENABLING PROMPT:",
  "- Task: Provide a 4-strength menu (kindness, courage, honesty, persistence). Students point to or write the matching one.",
  "- Extra Notes: Sentence stem: 'When the character ___, that shows ___.'",
  "EXTENDING PROMPT:",
  "- Task: Students name TWO strengths the character shows in the same action and explain how the strengths work together (e.g. courage + kindness).",
  "",
  "WATCH FOR:",
  "- Students stuck on the same one or two strengths -- nudge them to try a less common one (humour, hope, fairness).",
  "",
  "[General: We Do | VTLM 2.0: Guided Practice]",
].join("\n");

const NOTES_CFU = [
  "SAY:",
  "- Listen carefully. I'll read a sentence about someone.",
  "- On your whiteboard, write ACTION or STRENGTH.",
  "- Sentence: 'She is really good at scoring goals in soccer.'",
  "- Hold up your boards.",
  "",
  "DO:",
  "- Read the sentence twice.",
  "- Give 5-10 seconds thinking time.",
  "- Scan all boards before revealing the answer on the next slide.",
  "",
  "CFU CHECKPOINT:",
  "Technique: Show Me Boards.",
  "Script:",
  "- Read the sentence. Wait for boards up.",
  "- Scan for: ACTION (correct -- this describes a talent or skill, not a character strength).",
  "PROCEED: If most boards show ACTION, move on after the reveal.",
  "PIVOT: If many boards show STRENGTH, replay the contrast: 'Being good at soccer is a talent. A character strength is the quality inside the person -- like persistence at training, or fairness on the field.' Re-ask with: 'She kept training even after a tough loss.'",
  "",
  "TEACHER NOTES:",
  "This hinge question separates talent from character strength -- a confusion that came up in Session 1. Use the answer reveal slide to confirm and explain.",
  "",
  "WATCH FOR:",
  "- Students writing both -- accept the discussion this opens up.",
  "",
  "[General: CFU | VTLM 2.0: Check For Understanding]",
].join("\n");

const NOTES_CFU_REVEAL = [
  "SAY:",
  "- The answer is ACTION -- because being good at soccer is a talent, not a strength.",
  "- A character strength would be HOW she plays: with fairness, persistence, or teamwork.",
  "- Talents are what we can DO. Strengths are who we ARE inside.",
  "",
  "DO:",
  "- Pause briefly to let students compare to their own boards.",
  "- Quickly poll: 'Hands up if you wrote ACTION.' Acknowledge.",
  "",
  "TEACHER NOTES:",
  "Reinforce the talent vs strength distinction one more time before the You Do.",
  "",
  "WATCH FOR:",
  "- Students saying it's both -- that is a higher-level answer; honour it: 'Yes, she might also show persistence -- but the sentence as written only tells us about the talent.'",
  "",
  "[General: CFU Reveal | VTLM 2.0: Feedback]",
].join("\n");

const NOTES_YOUDO_FICTIONAL = [
  "SAY:",
  "- Your turn. Use Part A of your planner.",
  "- Choose a fictional character you admire -- from a book, film, or show.",
  "- Write 2 or 3 admirable actions they do.",
  "- For each action, pick a strength from the list.",
  "- You have about 4 minutes.",
  "",
  "DO:",
  "- Distribute the Strengths I Admire Planner if not already handed out.",
  "- Display the strengths list on the board (slide content) for easy reference.",
  "- Circulate. Prompt hesitant students with: 'What's a story you've enjoyed? Who in it does something you admire?'",
  "",
  "TEACHER NOTES:",
  "Students may name characters that aren't widely known -- that is fine. They are working with their own admiration, not the teacher's. The strengths list is the anchor.",
  "",
  "ENABLING & EXTENDING:",
  "ENABLING PROMPT:",
  "- Task: Offer a short menu of class-known characters from a recent class text. Students pick one.",
  "- Extra Notes: Provide a sentence frame: 'When ___, they show ___.'",
  "EXTENDING PROMPT:",
  "- Task: After the planner, students explain in writing why this strength matters in real life, not just in the story.",
  "",
  "WATCH FOR:",
  "- Students writing only the character's name without an action -- prompt: 'What did they DO?'",
  "- Students using the same strength three times -- nudge them to try a different one.",
  "",
  "[General: You Do | VTLM 2.0: Independent Practice]",
].join("\n");

const NOTES_PAIRSHARE = [
  "SAY:",
  "- In your group of 3 or 4, take turns sharing your character.",
  "- Say: who they are, one admirable action, and the strength it shows.",
  "- Listeners: notice -- did the same strength come up in different groups?",
  "- Each group will pick one character to share with the whole class.",
  "",
  "DO:",
  "- Form groups of 3 or 4.",
  "- Give about 4 minutes for the small-group share.",
  "- Then 2-3 minutes for one volunteer per group to share with the class.",
  "- Record characters and strengths on the board in two columns: CHARACTER | STRENGTH.",
  "",
  "TEACHER NOTES:",
  "The class board notes become the bridge to Part B. They also show the breadth of strengths students value -- name this aloud at the end of the share.",
  "",
  "WATCH FOR:",
  "- Quiet groups -- prompt: 'Who'd like to start? Just one sentence.'",
  "- Sensitivity: a chosen character may carry meaning for a student (a fictional character who has lost someone, a brave hero in a tough story). Acknowledge warmly without probing.",
  "",
  "[General: Pair-Share | VTLM 2.0: Social Learning]",
].join("\n");

const NOTES_YOUDO_REAL = [
  "SAY:",
  "- Now we shift from fiction to real life.",
  "- Use Part B of your planner.",
  "- Choose someone you know from your everyday life -- family, friend, neighbour, coach, teacher.",
  "- You can keep their name private if you want -- write a code (like 'M') instead of their name.",
  "- Write 2 or 3 admirable actions and the strengths they show.",
  "- You have about 4 minutes.",
  "",
  "DO:",
  "- Re-display the strengths list.",
  "- Circulate quietly. Some students may pause -- they may be thinking of someone they have lost or are processing tough feelings about.",
  "- Don't pressure students to share names. Names being private is fine.",
  "",
  "TEACHER NOTES:",
  "This is the heart of the activity. The strengths students name in real people are usually the strengths they value most. Notice which ones come up.",
  "",
  "ENABLING & EXTENDING:",
  "ENABLING PROMPT:",
  "- Task: For students stuck choosing a person, offer the prompt: 'Think of someone who has helped you when things were hard.'",
  "- Extra Notes: Sentence frame: 'When ___, they show ___.'",
  "EXTENDING PROMPT:",
  "- Task: Students add: 'A strength I'd like to grow in myself by learning from this person is ___.'",
  "",
  "SENSITIVITY ADVISORY:",
  "- What it is: This activity may surface feelings about people students have lost, miss, or have complicated relationships with.",
  "- Framing language: 'You don't have to share if it feels private. Writing it on your planner is enough.'",
  "- Watch for: tearfulness, withdrawal, or refusal -- offer a quiet alternative (write about a fictional person instead).",
  "- Protocol: Follow your school's wellbeing referral process if a student is visibly distressed.",
  "",
  "WATCH FOR:",
  "- Students who name famous people they don't know personally -- gently redirect: 'Try someone you actually know, even if just a little.'",
  "- Students who can't think of anyone -- offer: 'A teacher? A coach? A neighbour?'",
  "",
  "[General: You Do | VTLM 2.0: Independent Practice]",
].join("\n");

const NOTES_ROLE_MODELS = [
  "SAY:",
  "- The people we admire shape us. We call them ROLE MODELS.",
  "- Here's something important: now that you're in the older year levels, you are role models too.",
  "- Younger students watch you. They notice how you act in the yard, in the corridor, on the bus.",
  "- Quick think: When you're around younger students, which strength do you most want to show?",
  "- Turn & Talk: Tell your partner one strength you want younger students to see in you.",
  "",
  "DO:",
  "- Give 60 seconds for Turn & Talk.",
  "- Cold call 3-4 students to share.",
  "- Record their strengths on the board next to the words ROLE MODEL.",
  "",
  "TEACHER NOTES:",
  "This is the concept-bridge from the source activity: 'Part of the responsibility of reaching higher year levels is realising that you will have this influence.' Make this responsibility feel honourable, not heavy.",
  "",
  "WATCH FOR:",
  "- Students naming lots of strengths at once -- ask them to choose just one to show this week.",
  "- Students saying 'I'm not a role model' -- challenge gently: 'You're older than someone in this school. They notice.'",
  "",
  "[General: Application | VTLM 2.0: Transfer]",
].join("\n");

const NOTES_CLOSING = [
  "SAY:",
  "- Let's check our success criteria. Thumbs up, sideways, or down for each one.",
  "- SC1: I can name a person I admire and one thing they do.",
  "- SC2: I can match admirable actions to character strengths.",
  "- SC3: I can name a strength I want to show as a role model.",
  "- Reflection: Which strength from today do you most want to grow in yourself this term?",
  "- Tell your partner one tiny way you'll show that strength this week.",
  "",
  "DO:",
  "- Read each SC aloud; pause for thumbs.",
  "- Allow 60 seconds for Turn & Talk.",
  "- Acknowledge progress: name the breadth of strengths the class identified today.",
  "",
  "TEACHER NOTES:",
  "Connect back to Session 1 if useful: 'Last session we learned the strengths. Today we noticed them in the people we admire. Next, we'll think about the strengths in ourselves.'",
  "",
  "WATCH FOR:",
  "- Students showing sideways or down on SC2 -- note for a quick revisit.",
  "- Students who can't choose a strength to grow -- offer: 'Pick one that already feels close.'",
  "",
  "[General: Closing | VTLM 2.0: Review and Reflect]",
].join("\n");

// ============================================================
// Helper: draw the 16 strengths as a compact reference grid
// ============================================================

function drawStrengthsGrid(s, x, y, w, h) {
  const cols = 4;
  const rows = 4;
  const gap = 0.08;
  const cellW = (w - gap * (cols - 1)) / cols;
  const cellH = (h - gap * (rows - 1)) / rows;

  STRENGTHS.forEach((name, i) => {
    const r = Math.floor(i / cols);
    const c = i % cols;
    const cx = x + c * (cellW + gap);
    const cy = y + r * (cellH + gap);
    s.addShape("roundRect", {
      x: cx, y: cy, w: cellW, h: cellH, rectRadius: 0.06,
      fill: { color: C.BG_LIGHT },
      line: { color: C.ACCENT, width: 0.75 },
    });
    s.addText(name, {
      x: cx, y: cy, w: cellW, h: cellH,
      fontSize: 11, fontFace: FONT_B, color: C.PRIMARY, bold: true,
      align: "center", valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });
  });
}

// ============================================================
// Build function
// ============================================================

async function build() {
  const pres = new PptxGenJS();
  pres.layout = "LAYOUT_16x9";

  // -- Slide 1: Title --
  titleSlide(
    pres,
    "Strengths\nI Admire",
    "Respectful Relationships",
    "Grade 5/6 Wellbeing  |  Session 2",
    NOTES_TITLE
  );

  // -- Slide 2: Teacher Resources (immediately after title) --
  addResourceSlide(pres, RESOURCE_ITEMS, T, FOOTER, NOTES_RESOURCES);

  // -- Slide 3: Launch --
  contentSlide(
    pres,
    "Launch",
    C.SECONDARY,
    "Spot a Strength",
    [
      "Last session: we learned 16 character strengths",
      "Today: we'll spot those strengths in people we admire",
      "Turn & Talk: name one strength + a tiny example",
    ],
    NOTES_LAUNCH,
    FOOTER,
    (s, lg) => {
      const rX = lg.rightX;
      const rW = lg.rightW;
      const topY = lg.panelTopPadded;
      const cardH = 3.4;

      addCard(s, rX, topY, rW, cardH, { strip: C.PRIMARY, fill: C.WHITE });
      s.addText("16 Character Strengths", {
        x: rX + 0.15, y: topY + 0.10, w: rW - 0.3, h: 0.32,
        fontSize: 13, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", margin: 0,
      });
      drawStrengthsGrid(s, rX + 0.15, topY + 0.50, rW - 0.30, cardH - 0.65);
    }
  );

  // -- Slide 4: LI / SC --
  liSlide(
    pres,
    ["We are learning to identify the character strengths we admire in others, and to match admirable actions to those strengths"],
    [
      "I can name a person I admire and one thing they do",
      "I can match admirable actions to character strengths",
      "I can name a strength I want to show as a role model",
    ],
    NOTES_LI,
    FOOTER
  );

  // -- Slide 5: I Do — modelling actions → strengths --
  contentSlide(
    pres,
    "I Do",
    C.PRIMARY,
    "Watch How I Match Actions to Strengths",
    [
      "Step 1: Pick a fictional character you admire",
      "Step 2: Write 2-3 admirable actions",
      "Step 3: Match each action to a strength",
    ],
    NOTES_IDO,
    FOOTER,
    (s, lg) => {
      const rX = lg.rightX;
      const rW = lg.rightW;
      const topY = lg.panelTopPadded;
      const cardH = 3.5;

      addCard(s, rX, topY, rW, cardH, { strip: C.ACCENT, fill: C.WHITE });
      s.addText("Worked Example: A young hero in a story", {
        x: rX + 0.15, y: topY + 0.10, w: rW - 0.30, h: 0.32,
        fontSize: 12, fontFace: FONT_H, color: C.ACCENT, bold: true,
        align: "center", margin: 0,
      });

      // Three rows: ACTION → STRENGTH
      const rowsY = topY + 0.55;
      const rowH = 0.78;
      const arrowW = 0.45;
      const cellPad = 0.10;
      const totalW = rW - 0.30;
      const cellW = (totalW - arrowW) / 2;
      const examples = [
        { action: "Stood up to a bully",        strength: "COURAGE" },
        { action: "Stayed beside an upset friend", strength: "KINDNESS" },
        { action: "Kept practising when tired",  strength: "PERSISTENCE" },
      ];
      examples.forEach((ex, i) => {
        const ry = rowsY + i * (rowH + 0.10);

        // Action cell
        s.addShape("roundRect", {
          x: rX + 0.15, y: ry, w: cellW, h: rowH, rectRadius: 0.06,
          fill: { color: C.BG_LIGHT },
          line: { color: C.MUTED, width: 0.75 },
        });
        s.addText(ex.action, {
          x: rX + 0.15 + cellPad, y: ry, w: cellW - cellPad * 2, h: rowH,
          fontSize: 11, fontFace: FONT_B, color: C.CHARCOAL,
          align: "center", valign: "middle", margin: 0,
          fit: "shrink", shrinkText: true,
        });

        // Arrow
        s.addText(">", {
          x: rX + 0.15 + cellW, y: ry, w: arrowW, h: rowH,
          fontSize: 22, fontFace: FONT_H, color: C.ACCENT, bold: true,
          align: "center", valign: "middle", margin: 0,
        });

        // Strength cell
        s.addShape("roundRect", {
          x: rX + 0.15 + cellW + arrowW, y: ry, w: cellW, h: rowH, rectRadius: 0.06,
          fill: { color: C.PRIMARY },
        });
        s.addText(ex.strength, {
          x: rX + 0.15 + cellW + arrowW, y: ry, w: cellW, h: rowH,
          fontSize: 13, fontFace: FONT_H, color: C.WHITE, bold: true,
          align: "center", valign: "middle", margin: 0,
        });
      });
    }
  );

  // -- Slide 6: We Do — class works through a shared character --
  contentSlide(
    pres,
    "We Do",
    C.SECONDARY,
    "Try One Together",
    [
      "Pick a character we all know",
      "What is one admirable action they do?",
      "Which strength does that show?",
      "Mini-whiteboards: write one strength",
    ],
    NOTES_WEDO,
    FOOTER,
    (s, lg) => {
      const rX = lg.rightX;
      const rW = lg.rightW;
      const topY = lg.panelTopPadded;
      const cardH = 3.4;

      addCard(s, rX, topY, rW, cardH, { strip: C.SECONDARY, fill: C.WHITE });
      s.addText("Class Builds Together", {
        x: rX + 0.15, y: topY + 0.10, w: rW - 0.30, h: 0.32,
        fontSize: 13, fontFace: FONT_H, color: C.SECONDARY, bold: true,
        align: "center", margin: 0,
      });

      // Two-column header
      const headerY = topY + 0.55;
      const headerH = 0.42;
      const colW = (rW - 0.30 - 0.10) / 2;

      s.addShape("roundRect", {
        x: rX + 0.15, y: headerY, w: colW, h: headerH, rectRadius: 0.06,
        fill: { color: C.PRIMARY },
      });
      s.addText("ACTION", {
        x: rX + 0.15, y: headerY, w: colW, h: headerH,
        fontSize: 13, fontFace: FONT_H, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      s.addShape("roundRect", {
        x: rX + 0.15 + colW + 0.10, y: headerY, w: colW, h: headerH, rectRadius: 0.06,
        fill: { color: C.ACCENT },
      });
      s.addText("STRENGTH", {
        x: rX + 0.15 + colW + 0.10, y: headerY, w: colW, h: headerH,
        fontSize: 13, fontFace: FONT_H, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      // Three blank rows for board build
      const rowsTop = headerY + headerH + 0.10;
      const rowH = 0.50;
      for (let i = 0; i < 3; i++) {
        const ry = rowsTop + i * (rowH + 0.08);
        s.addShape("roundRect", {
          x: rX + 0.15, y: ry, w: colW, h: rowH, rectRadius: 0.06,
          fill: { color: C.WHITE },
          line: { color: C.MUTED, width: 0.75 },
        });
        s.addShape("roundRect", {
          x: rX + 0.15 + colW + 0.10, y: ry, w: colW, h: rowH, rectRadius: 0.06,
          fill: { color: C.WHITE },
          line: { color: C.MUTED, width: 0.75 },
        });
      }
    }
  );

  // -- Slide 7: CFU Hinge question --
  cfuSlide(
    pres,
    "CFU",
    "Action or Strength?",
    "Show Me Boards",
    "She is really good at scoring goals in soccer.\n\nWrite ACTION or STRENGTH on your whiteboard.",
    NOTES_CFU,
    FOOTER
  );

  // -- Slide 8: CFU Reveal --
  {
    const s = pres.addSlide();
    addTopBar(s, C.ALERT);
    addBadge(s, "CFU Reveal", { color: C.ALERT });
    addTitle(s, "Action or Strength?", { color: C.ALERT });

    // Big answer card
    const cardY = CONTENT_TOP + 0.20;
    const cardH = 1.6;
    addCard(s, 0.5, cardY, 9, cardH, { strip: C.ALERT, fill: C.WHITE });
    s.addText("ACTION (it's a talent)", {
      x: 0.5, y: cardY, w: 9, h: cardH,
      fontSize: 36, fontFace: FONT_H, color: C.ALERT, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    // Explanation card
    const explY = cardY + cardH + 0.20;
    const explH = SAFE_BOTTOM - explY - 0.05;
    addCard(s, 0.5, explY, 9, explH, { strip: C.PRIMARY, fill: C.BG_LIGHT });
    s.addText([
      { text: "Talent: ", options: { fontSize: 18, color: C.PRIMARY, bold: true } },
      { text: "what someone can DO (scoring goals).", options: { fontSize: 18, color: C.CHARCOAL, breakLine: true } },
      { text: "", options: { fontSize: 8, breakLine: true } },
      { text: "Strength: ", options: { fontSize: 18, color: C.PRIMARY, bold: true } },
      { text: "who someone IS inside (fairness, persistence, teamwork).", options: { fontSize: 18, color: C.CHARCOAL } },
    ], {
      x: 0.85, y: explY + 0.20, w: 8.3, h: explH - 0.40,
      fontFace: FONT_B, valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_CFU_REVEAL);
  }

  // -- Slide 9: You Do (Part A) — fictional character --
  contentSlide(
    pres,
    "You Do",
    C.PRIMARY,
    "Part A: A Character You Admire",
    [
      "First: Pick a fictional character you admire",
      "Next: Write 2-3 admirable actions",
      "Then: Match each action to a strength",
      "Use Part A of your planner. About 4 minutes.",
    ],
    NOTES_YOUDO_FICTIONAL,
    FOOTER,
    (s, lg) => {
      const rX = lg.rightX;
      const rW = lg.rightW;
      const topY = lg.panelTopPadded;
      const cardH = 3.4;

      addCard(s, rX, topY, rW, cardH, { strip: C.ACCENT, fill: C.WHITE });
      s.addText("Strengths List", {
        x: rX + 0.15, y: topY + 0.10, w: rW - 0.3, h: 0.32,
        fontSize: 13, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", margin: 0,
      });
      drawStrengthsGrid(s, rX + 0.15, topY + 0.50, rW - 0.30, cardH - 0.65);
    }
  );

  // -- Slide 10: Pair-Share / Group share --
  pairShareSlide(
    pres,
    "Share with Your Group",
    [
      "Take turns: name your character, one action, one strength",
      "Listen for: did the same strength come up in different groups?",
      "Pick one character to share with the whole class",
    ],
    NOTES_PAIRSHARE,
    FOOTER
  );

  // -- Slide 11: You Do (Part B) — someone from everyday life --
  contentSlide(
    pres,
    "You Do",
    C.PRIMARY,
    "Part B: Someone from Everyday Life",
    [
      "Choose someone you know -- family, friend, coach, teacher",
      "Names can be private (write a code, like 'M')",
      "Write 2-3 actions and the strengths they show",
      "Use Part B of your planner. About 4 minutes.",
    ],
    NOTES_YOUDO_REAL,
    FOOTER,
    (s, lg) => {
      const rX = lg.rightX;
      const rW = lg.rightW;
      const topY = lg.panelTopPadded;
      const cardH = 3.4;

      addCard(s, rX, topY, rW, cardH, { strip: C.SUCCESS, fill: C.WHITE });
      s.addText("Strengths List", {
        x: rX + 0.15, y: topY + 0.10, w: rW - 0.3, h: 0.32,
        fontSize: 13, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", margin: 0,
      });
      drawStrengthsGrid(s, rX + 0.15, topY + 0.50, rW - 0.30, cardH - 0.65);
    }
  );

  // -- Slide 12: Role Models bridge --
  contentSlide(
    pres,
    "Connect",
    C.SUCCESS,
    "You Are Role Models Too",
    [
      "The people we admire are role models",
      "Now you are in the older year levels...",
      "Younger students notice how you act",
      "Turn & Talk: which strength do you most want them to see?",
    ],
    NOTES_ROLE_MODELS,
    FOOTER,
    (s, lg) => {
      const rX = lg.rightX;
      const rW = lg.rightW;
      const topY = lg.panelTopPadded;
      const cardH = 3.3;

      addCard(s, rX, topY, rW, cardH, { strip: C.SUCCESS, fill: C.WHITE });
      s.addText("ROLE MODEL", {
        x: rX + 0.15, y: topY + 0.18, w: rW - 0.30, h: 0.45,
        fontSize: 18, fontFace: FONT_H, color: C.SUCCESS, bold: true,
        align: "center", margin: 0,
      });

      // Two arrows showing "older student → younger student"
      const sceneY = topY + 0.85;
      const sceneH = cardH - 1.05;
      const figW = 1.2;
      const olderX = rX + 0.4;
      const youngerX = rX + rW - figW - 0.4;
      const figTop = sceneY + 0.10;
      const headR = 0.20;

      // Older student (taller)
      s.addShape("roundRect", {
        x: olderX + figW / 2 - headR, y: figTop, w: headR * 2, h: headR * 2,
        rectRadius: headR,
        fill: { color: C.PRIMARY },
        line: { color: C.PRIMARY, width: 1 },
      });
      s.addShape("roundRect", {
        x: olderX + figW / 2 - 0.20, y: figTop + headR * 2, w: 0.40, h: 0.85,
        rectRadius: 0.10,
        fill: { color: C.PRIMARY },
      });
      s.addText("YOU", {
        x: olderX, y: figTop + headR * 2 + 0.95, w: figW, h: 0.30,
        fontSize: 12, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      // Younger student (shorter)
      const youngerTop = figTop + 0.30;
      const youngerHeadR = 0.16;
      s.addShape("roundRect", {
        x: youngerX + figW / 2 - youngerHeadR, y: youngerTop, w: youngerHeadR * 2, h: youngerHeadR * 2,
        rectRadius: youngerHeadR,
        fill: { color: C.ACCENT },
        line: { color: C.ACCENT, width: 1 },
      });
      s.addShape("roundRect", {
        x: youngerX + figW / 2 - 0.16, y: youngerTop + youngerHeadR * 2, w: 0.32, h: 0.65,
        rectRadius: 0.08,
        fill: { color: C.ACCENT },
      });
      s.addText("Younger", {
        x: youngerX, y: youngerTop + youngerHeadR * 2 + 0.75, w: figW, h: 0.30,
        fontSize: 12, fontFace: FONT_H, color: C.ACCENT, bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      // Eyes-watching arrow from younger to older (arrowhead points at YOU)
      const arrowY = figTop + 0.40;
      const arrowStartX = youngerX + 0.10;
      const arrowEndX = olderX + figW - 0.10;
      s.addShape("line", {
        x: arrowEndX, y: arrowY,
        w: arrowStartX - arrowEndX, h: 0,
        line: { color: C.SUCCESS, width: 2, beginArrowType: "triangle" },
      });
      s.addText("watches", {
        x: olderX + figW, y: arrowY - 0.30, w: youngerX - (olderX + figW), h: 0.28,
        fontSize: 11, fontFace: FONT_B, color: C.SUCCESS, italic: true,
        align: "center", valign: "middle", margin: 0,
      });
    }
  );

  // -- Slide 13: Closing --
  closingSlide(
    pres,
    {
      reflectionPrompt: "Which strength from today do you most want to grow in yourself this term? Tell your partner one tiny way you'll show it this week.",
      scItems: [
        "I can name a person I admire and one thing they do",
        "I can match admirable actions to character strengths",
        "I can name a strength I want to show as a role model",
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

  // ============================================================
  // Companion PDF: Strengths I Admire Planner
  // ============================================================
  {
    const doc = createPdf({ title: "Strengths I Admire Planner" });
    let y = addPdfHeader(doc, "Session 2 Strengths I Admire Planner", {
      subtitle: "Match admirable actions to character strengths",
      color: C.PRIMARY,
      lessonInfo: "Grade 5/6 Wellbeing - Respectful Relationships",
    });

    y = addTipBox(doc,
      "A character strength is who someone IS inside -- like kindness, courage, fairness, or persistence. Use the strengths list from Session 1 to help you match.",
      y,
      { color: C.ACCENT }
    );

    // ----- Strengths reference strip (single line list) -----
    y = addSectionHeading(doc, "Strengths reference", y, { color: C.ACCENT });
    doc.fontSize(9.5).font("Sans").fillColor(hex(C.CHARCOAL));
    doc.text(STRENGTHS.join("  •  "), PAGE.MARGIN, y, {
      width: PAGE.CONTENT_W,
      lineGap: 2,
    });
    y = doc.y + 14;

    // ----- Part A: Fictional Character -----
    y = addSectionHeading(doc, "Part A: A fictional character I admire", y, { color: C.PRIMARY });

    doc.fontSize(10).font("Sans-Bold").fillColor(hex(C.PRIMARY));
    doc.text("Character name:", PAGE.MARGIN, y);
    y += 14;
    y = addLinedArea(doc, y, 1, { lineSpacing: 22 });

    // Action / Strength two-column table for Part A
    y = drawActionStrengthTable(doc, y, 3, C.PRIMARY);

    y += 6;

    // ----- Part B: Everyday Person -----
    if (y + 200 > PAGE.H - PAGE.MARGIN - 40) {
      doc.addPage();
      y = PAGE.MARGIN;
    }

    y = addSectionHeading(doc, "Part B: Someone from everyday life I admire", y, { color: C.SECONDARY });

    doc.fontSize(10).font("Sans").fillColor(hex(C.CHARCOAL));
    doc.text("(Their name -- or a private code like 'M' if you'd prefer to keep it private):", PAGE.MARGIN, y, {
      width: PAGE.CONTENT_W,
    });
    y = doc.y + 8;
    y = addLinedArea(doc, y, 1, { lineSpacing: 22 });

    y = drawActionStrengthTable(doc, y, 3, C.SECONDARY);

    y += 8;

    // ----- Reflection -----
    if (y + 90 > PAGE.H - PAGE.MARGIN - 40) {
      doc.addPage();
      y = PAGE.MARGIN;
    }

    y = addSectionHeading(doc, "My reflection", y, { color: C.ACCENT });
    y = addBodyText(doc,
      "A strength I want to grow in myself by learning from these people:",
      y,
      { fontSize: 10 }
    );
    y = addLinedArea(doc, y, 3, { lineSpacing: 24 });

    y = addBodyText(doc,
      "Why does this strength matter to me?",
      y,
      { fontSize: 10 }
    );
    y = addLinedArea(doc, y, 3, { lineSpacing: 24 });

    y = addBodyText(doc,
      "One tiny way I will show that strength this week (around younger students or anywhere):",
      y,
      { fontSize: 10 }
    );
    y = addLinedArea(doc, y, 3, { lineSpacing: 24 });

    addPdfFooter(doc, "Respectful Relationships | Grade 5/6 | Session 2");
    const plannerPath = path.join(RES_DIR, "Session 2 Strengths I Admire Planner.pdf");
    await writePdf(doc, plannerPath);
    console.log("PDF written:", plannerPath);
  }

  console.log("\nBuild complete!");
  console.log("Output folder:", LESSON_FOLDER);
}

// ============================================================
// PDF helper: action / strength two-column table
// ============================================================

function drawActionStrengthTable(doc, y, rowCount, headerColor) {
  const colGap = 10;
  const colW = (PAGE.CONTENT_W - colGap) / 2;
  const rowH = 36;
  const headerH = 22;

  // Headers
  doc.save();
  doc.rect(PAGE.MARGIN, y, colW, headerH).fillColor(hex(headerColor)).fill();
  doc.rect(PAGE.MARGIN + colW + colGap, y, colW, headerH).fillColor(hex(headerColor)).fill();
  doc.restore();

  doc.fontSize(11).font("Sans-Bold").fillColor("#FFFFFF");
  doc.text("ACTION I admire", PAGE.MARGIN, y + 6, { width: colW, align: "center" });
  doc.text("STRENGTH it shows", PAGE.MARGIN + colW + colGap, y + 6, { width: colW, align: "center" });

  let cy = y + headerH;

  for (let i = 0; i < rowCount; i++) {
    // Cell borders
    doc.save();
    doc.lineWidth(0.8).strokeColor("#888888");
    doc.rect(PAGE.MARGIN, cy, colW, rowH).stroke();
    doc.rect(PAGE.MARGIN + colW + colGap, cy, colW, rowH).stroke();
    doc.restore();

    // Light writing line inside each cell (about 2/3 down)
    doc.save();
    doc.lineWidth(0.4).strokeColor("#BBBBBB");
    doc.moveTo(PAGE.MARGIN + 6, cy + rowH - 8)
      .lineTo(PAGE.MARGIN + colW - 6, cy + rowH - 8)
      .stroke();
    doc.moveTo(PAGE.MARGIN + colW + colGap + 6, cy + rowH - 8)
      .lineTo(PAGE.MARGIN + colW + colGap + colW - 6, cy + rowH - 8)
      .stroke();
    doc.restore();

    cy += rowH;
  }

  return cy + 6;
}

build().catch((err) => { console.error(err); process.exit(1); });
