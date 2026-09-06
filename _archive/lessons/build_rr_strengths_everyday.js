"use strict";

// Respectful Relationships — Year 5/6
// "Using our strengths in everyday life" (15+ minutes)
// Single-session deck. No generated PDFs — scenarios + character strengths
// handout are supplied by the teacher from the Resilience, Rights and
// Respectful Relationships unit (Topic: Personal Strengths).

const path = require("path");
const fs = require("fs");
const PptxGenJS = require("pptxgenjs");
const { createTheme, weekToVariant } = require("../themes/factory");
const { addRichResourceSlide } = require("../themes/pdf_helpers");

const T = createTheme("wellbeing", "grade56", weekToVariant(2));
const {
  C, FONT_H, FONT_B,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  pairShareSlide, scenarioSlide,
} = T;

const UNIT = "RR_Strengths_Everyday";
const LESSON_FOLDER = path.join(__dirname, "..", "output", UNIT);
const PPTX_NAME = "Using Our Strengths in Everyday Life.pptx";
const FOOTER = "Using our strengths in everyday life | Year 5/6 Respectful Relationships";

fs.mkdirSync(LESSON_FOLDER, { recursive: true });

// ════════════════════════════════════════════════════════════════════
// Teacher notes
// ════════════════════════════════════════════════════════════════════

const NOTES_TITLE = [
  "SAY:",
  "- Last time we thought about people we admire and the character strengths they show",
  "- Today we take it one step further -- we work out which strengths help with real, everyday challenges",
  "- This is a short, sharp activity. About 15 minutes",
  "",
  "DO:",
  "- Have the Character Strengths handout out on each table",
  "- Have the cut-up Everyday Strengths Scenarios ready (two per pair)",
  "",
  "TEACHER NOTES:",
  "Builds on the prior Strengths I Admire session. Today moves from admiring strengths in others to choosing strengths for ourselves to use in everyday situations.",
  "",
  "WATCH FOR:",
  "- Students confusing talents (e.g. good at football) with character strengths (e.g. perseverance) -- redirect if needed",
  "",
  "[General: Title | VTLM 2.0: Establishing Purpose]",
].join("\n");

const NOTES_RESOURCES = [
  "SAY:",
  "- I will hand out two scenario cards to each pair",
  "- Keep your Character Strengths handout out -- you will use it the whole lesson",
  "",
  "DO:",
  "- Pre-cut the Everyday Strengths Scenarios so there are enough for two per pair",
  "- Place the Character Strengths handout on every desk",
  "- Have a whiteboard / butcher's paper space ready for capturing class feedback",
  "",
  "TEACHER NOTES:",
  "No new printed resources are generated for this lesson. The two handouts come from the Personal Strengths topic in the Resilience, Rights and Respectful Relationships unit. If your scenarios pack is missing, write three or four short everyday challenges on sticky notes and use those instead.",
  "",
  "WATCH FOR:",
  "- Make sure each pair has TWO different scenarios so feedback time covers a wider range of strengths",
  "",
  "[General: Resources | VTLM 2.0: Preparing Resources]",
].join("\n");

const NOTES_LAUNCH = [
  "SAY:",
  "- Last lesson we thought of a person we admire and the character strengths they show",
  "- Quick share -- turn to your partner. Say the name of one person you admire and ONE character strength they have. Ten seconds each",
  "- Today we are flipping it around. Instead of admiring strengths in someone else, we are going to choose strengths WE could use in everyday situations",
  "",
  "DO:",
  "- Cue a sharp 20-second Turn and Tell",
  "- Cold call two students to share the strength they named",
  "",
  "TEACHER NOTES:",
  "Quick prior-knowledge bridge. Keep it short -- this is a 15-minute lesson and the main activity is the scenarios.",
  "",
  "WATCH FOR:",
  "- Students naming talents (good at art, fast runner) -- gently redirect to character strengths (creativity, perseverance)",
  "",
  "[General: Launch | VTLM 2.0: Activating Prior Knowledge]",
].join("\n");

const NOTES_LI = [
  "SAY:",
  "- Read the learning intention together",
  "- Read each I can statement",
  "",
  "DO:",
  "- Choral read the LI",
  "- Thumbs up if you remember one strength from last lesson",
  "",
  "TEACHER NOTES:",
  "Three success criteria build from naming a strength, to matching it to a situation, to giving advice another student could actually use.",
  "",
  "WATCH FOR:",
  "- Students who cannot recall any strength names -- direct them to the Character Strengths handout",
  "",
  "[General: LI/SC | VTLM 2.0: Clear Learning Intention]",
].join("\n");

const NOTES_RECAP = [
  "SAY:",
  "- Quick reminder before we start. A character strength is something about HOW you act, not what you are good at",
  "- A talent is something you can DO well. A character strength is something you ARE",
  "- Kindness, honesty, perseverance, curiosity, fairness -- these are character strengths",
  "- Drawing, swimming, maths -- these are talents",
  "- Today we use character strengths to help with everyday challenges",
  "",
  "DO:",
  "- Point to the Character Strengths handout on desks",
  "- Ask: Tell me one strength from the list. Cold call three students",
  "",
  "TEACHER NOTES:",
  "Brief recap, not a re-teach. Anchors the rest of the lesson in the agreed strengths vocabulary so students all use the same words.",
  "",
  "WATCH FOR:",
  "- Students who name a talent again -- prompt: That is a talent. What strength helps someone keep getting better at it?",
  "",
  "[General: Concept Recap | VTLM 2.0: Connecting to Prior Knowledge]",
].join("\n");

const NOTES_IDO = [
  "SAY:",
  "- Watch how I work through a scenario before you try one",
  "- The scenario: A new student has joined the class. They sit alone at lunch every day. Mia notices",
  "- I am going to think aloud. Which character strength would best help Mia in this everyday challenge?",
  "- I look at the strengths list. Kindness fits. So does courage -- it takes a bit of bravery to walk up to someone you do not know",
  "- I think the best fit is kindness with a bit of courage. Why? Because the situation is about caring for someone AND doing something a little brave",
  "- My advice for Mia: Use your kindness. Walk over, say hi, ask if they want to sit with you. The courage part is just taking the first step",
  "",
  "DO:",
  "- Point to kindness and courage on the Character Strengths handout as you name them",
  "- Slow down on the advice -- model giving SPECIFIC advice, not just naming a feeling",
  "",
  "TEACHER NOTES:",
  "This is the I Do. Show the thinking move: read the scenario, pick a strength, justify the choice, then turn it into practical advice. Resist the urge to ask students for their ideas yet -- this is teacher modelling.",
  "",
  "WATCH FOR:",
  "- Students wanting to call out -- hold them off, You will get your turn in a moment",
  "",
  "[General: I Do | VTLM 2.0: Explicit Modelling]",
].join("\n");

const NOTES_WEDO = [
  "SAY:",
  "- Now we do one together",
  "- Scenario: Jay forgot his homework AGAIN. He is starting to feel like he just cannot keep on top of school stuff",
  "- Turn to your partner. Which character strength would best help Jay? You have 30 seconds",
  "- Come back -- show me on your fingers the number on the strengths list, or call out the strength",
  "- Why that strength? Tell your partner the reason",
  "- What advice would you give Jay? Whisper it to your partner",
  "",
  "DO:",
  "- Set a 30-second timer for the strength choice",
  "- Cold call three pairs after each prompt",
  "- Capture two or three suggested strengths on the board (e.g. perseverance, self-control, hope)",
  "",
  "CFU CHECKPOINT:",
  "Technique: Turn and Tell + Cold Call",
  "Script:",
  "- Tell your partner the strength and ONE reason",
  "- Scan for: pairs who can name BOTH a strength and a reason in their own words",
  "PROCEED:",
  "- If most pairs can do both, move to the hinge CFU",
  "PIVOT:",
  "- Most likely issue: pairs naming a strength but no reason. Re-model the reason step using kindness and Mia, then re-check with a new pair",
  "",
  "TEACHER NOTES:",
  "We Do is shared thinking, not teacher monologue. Use the Character Strengths handout in students hands -- this is the support that keeps everyone in the game.",
  "",
  "WATCH FOR:",
  "- Pairs offering judgement (Jay should be more organised) instead of strength-based advice -- redirect: Which strength helps him build that?",
  "",
  "[General: We Do | VTLM 2.0: Guided Practice]",
].join("\n");

const NOTES_CFU = [
  "SAY:",
  "- Quick check before pairs do their own scenarios",
  "- Read the action on the slide. Pick the character strength that BEST matches it. Show me on your fingers: 1, 2, 3 or 4",
  "- Three, two, one -- show me",
  "",
  "DO:",
  "- Wait for all hands to be up before scanning",
  "- Identify the most common answer",
  "- Reveal the answer and a one-line reason",
  "",
  "CFU CHECKPOINT:",
  "Technique: Show Me Boards (fingers)",
  "Script:",
  "- Read the action. Pick the strength that BEST fits. Fingers up on 3",
  "- Scan for: at least 80% choosing perseverance",
  "PROCEED:",
  "- If 80%+ show perseverance, move on to You Do",
  "PIVOT:",
  "- Most likely misconception: students pick honesty (because the student is being honest about struggling) instead of perseverance (the strength that helps them KEEP trying). Reteach: honesty is naming the truth. Perseverance is what gets you back up. Re-check with a different action: A student stays calm when a teammate yells at them in soccer. Which strength?",
  "",
  "TEACHER NOTES:",
  "Hinge question. Wrong answers are diagnostic -- they tell you whether students are confusing related strengths.",
  "",
  "WATCH FOR:",
  "- Students checking what their friends choose -- count down sharply so fingers go up at the same time",
  "",
  "[General: CFU | VTLM 2.0: Checking for Understanding]",
].join("\n");

const NOTES_CFU_REVEAL = [
  "SAY:",
  "- The best fit here is perseverance",
  "- The student is feeling frustrated but keeps trying. That is the strength of perseverance in action",
  "- Honesty is close, but honesty is about telling the truth. Perseverance is about not giving up",
  "",
  "DO:",
  "- Click to reveal the answer",
  "- Briefly acknowledge close answers (honesty, hope, self-control) so students who chose those still feel heard",
  "",
  "TEACHER NOTES:",
  "Reveal only after the show-me. Do not let students change their answers before the reveal.",
  "",
  "WATCH FOR:",
  "- Students who insist on a different answer -- ask them to explain, then connect it back to perseverance",
  "",
  "[General: CFU Reveal | VTLM 2.0: Feedback]",
].join("\n");

const NOTES_YOUDO = [
  "SAY:",
  "- Now your turn. You and your partner have TWO scenarios",
  "- For each scenario: pick the best character strength, agree on a reason, and prepare ONE piece of advice you could give that character",
  "- Use the Character Strengths handout. It is your toolkit",
  "- You have six minutes. Be ready to share back",
  "",
  "DO:",
  "- Hand out the cut-up Everyday Strengths Scenarios -- two per pair",
  "- Set a 6-minute timer visible on the board",
  "- Circulate. Listen for pairs who can justify their choice in their own words -- those pairs will share back first",
  "",
  "TEACHER NOTES:",
  "Pairs may pick different strengths for the same scenario -- that is exactly the point. There is no single right answer; the reasoning matters more than the label.",
  "",
  "ENABLING & EXTENDING:",
  "ENABLING PROMPT:",
  "- Task: Use the handout. Point to a strength that feels like it fits. Then tell your partner WHY",
  "- Extra Notes: Pair a less confident student with a confident partner. Provide a sentence starter: This person needs ___ because ___",
  "EXTENDING PROMPT:",
  "- Task: For one of your scenarios, pick TWO strengths that work together. Explain how they help in different ways",
  "- Extra Notes: Push students to give specific, doable advice (the exact words or action) rather than general feelings",
  "",
  "WATCH FOR:",
  "- Pairs naming a strength but giving advice that does not use that strength -- prompt: How does THAT strength help here?",
  "- Pairs finishing too quickly -- ask for a second strength that also fits",
  "",
  "[General: You Do | VTLM 2.0: Independent Practice]",
].join("\n");

const NOTES_SHARE = [
  "SAY:",
  "- Time is up. Let us share",
  "- I will call on three pairs. For each one: What scenario did you have? Which strength did you choose? Why? And what was your advice?",
  "- Listen for similarities and differences -- it is okay to disagree",
  "",
  "DO:",
  "- Call on three pre-selected pairs first (the ones whose reasoning you heard while circulating)",
  "- Capture strengths chosen on the board as they speak",
  "- Ask one follow-up question: Did anyone choose a DIFFERENT strength for the same kind of situation?",
  "",
  "TEACHER NOTES:",
  "This is the feedback step from the source method. It is where the class sees that different people choose different strengths for the same challenge -- and that is valid. Strengths are personal.",
  "",
  "WATCH FOR:",
  "- A single strength dominating (e.g. kindness for everything) -- prompt the class: What other strengths could help here too?",
  "",
  "[General: Share Back | VTLM 2.0: Making Thinking Visible]",
].join("\n");

const NOTES_CLOSING = [
  "SAY:",
  "- Quick review. Did we meet our learning intentions today?",
  "- Look at the I can statements. Thumbs up, sideways or down for each one",
  "- One last thought: pick ONE character strength you could use THIS WEEK -- at home, in the yard, in class. Whisper it to your partner",
  "",
  "DO:",
  "- Run thumbs check against each I can in turn",
  "- Note which criterion got the most sideways/down thumbs -- that tells you what to revisit next session",
  "- Acknowledge the class: Today you identified strengths AND gave each other real advice. That is using strengths in everyday life",
  "",
  "TEACHER NOTES:",
  "Megaprompt section 52: the closing reviews the success criteria with a self-assessment routine and acknowledges progress. The whisper share at the end seeds tomorrow's transfer.",
  "",
  "WATCH FOR:",
  "- Sideways thumbs on SC2 or SC3 -- plan a quick 5-minute warm-up next session with one more scenario",
  "",
  "[General: Closing | VTLM 2.0: Review and Reflection]",
].join("\n");

// ════════════════════════════════════════════════════════════════════
// Build deck
// ════════════════════════════════════════════════════════════════════

async function build() {
  const pres = new PptxGenJS();
  pres.layout = "LAYOUT_16x9";
  pres.title = "Using Our Strengths in Everyday Life — Year 5/6 RR";

  // ── Slide 1: Title ──
  titleSlide(
    pres,
    "Using Our Strengths in Everyday Life",
    "Choosing character strengths for real situations",
    "Year 5/6  |  Respectful Relationships  |  15+ minutes",
    NOTES_TITLE
  );

  // ── Slide 2: Teacher Resources ──
  addRichResourceSlide(
    pres,
    {
      resources: [],
      studentTools: [
        "Character Strengths handout (one per student, from Activity 1)",
        "Everyday Strengths Scenarios (pre-cut, two per pair)",
        "Whiteboards or inquiry books for jotting strengths",
      ],
      boardSetup: [
        "Space on the board to capture strengths suggested during share-back",
        "Timer visible (for the 6-minute pair task)",
      ],
      routineIcons: [
        "Turn and Tell",
        "Show Me (fingers)",
        "Pair work",
      ],
    },
    { C, FONT_H, FONT_B },
    FOOTER,
    NOTES_RESOURCES
  );

  // ── Slide 3: Launch ──
  pairShareSlide(
    pres,
    "Strengths warm-up",
    [
      "Tell your partner: one person you admire AND one character strength they show.",
      "Today we flip it around: which strengths could WE use in everyday situations?",
    ],
    NOTES_LAUNCH,
    FOOTER
  );

  // ── Slide 4: LI & SC ──
  liSlide(
    pres,
    "I can identify the character strengths that help with everyday challenges.",
    [
      "I can name a character strength from the list.",
      "I can match a character strength to an everyday situation and say why.",
      "I can give a piece of advice that uses that strength in action.",
    ],
    NOTES_LI,
    FOOTER
  );

  // ── Slide 5: Concept recap (strengths vs talents) ──
  contentSlide(
    pres,
    "Quick recap",
    C.SECONDARY,
    "Character strengths vs talents",
    [
      "A talent is something you DO well — drawing, footy, maths.",
      "A character strength is HOW you act — kindness, perseverance, courage, honesty, fairness, curiosity.",
      "Today we use character strengths to help with everyday challenges.",
    ],
    NOTES_RECAP,
    FOOTER
  );

  // ── Slide 6: I Do — modelled scenario ──
  scenarioSlide(
    pres,
    "I Do",
    "Watch how I choose a strength",
    "A new student has joined the class. They sit alone at lunch every day. Mia notices.",
    [
      "Best strength: kindness (with a little courage).",
      "Advice for Mia: walk over, say hi, ask if they want to sit with you.",
    ],
    NOTES_IDO,
    FOOTER
  );

  // ── Slide 7: We Do — shared scenario ──
  scenarioSlide(
    pres,
    "We Do",
    "Let's work one out together",
    "Jay forgot his homework AGAIN. He is starting to feel like he just cannot keep on top of school stuff.",
    [
      "Which character strength would best help Jay?",
      "Why that strength? What advice would you give him?",
    ],
    NOTES_WEDO,
    FOOTER
  );

  // ── Slide 8: CFU hinge ──
  cfuSlide(
    pres,
    "Check",
    "Pick the best-fit strength",
    "Show Me (fingers)",
    "A student is struggling with long division. They feel frustrated but keep practising every day.\n\nWhich strength fits best?\n   1) Honesty     2) Perseverance     3) Curiosity     4) Fairness",
    NOTES_CFU,
    FOOTER
  );

  // ── Slide 9: CFU reveal ──
  cfuSlide(
    pres,
    "Check — answer",
    "Best-fit strength: Perseverance",
    "Reveal",
    "Perseverance — they feel frustrated but keep trying.\n\nHonesty is about telling the truth. Perseverance is about not giving up.",
    NOTES_CFU_REVEAL,
    FOOTER
  );

  // ── Slide 10: You Do — pairs work two scenarios ──
  pairShareSlide(
    pres,
    "You Do — with your partner",
    [
      "Take your TWO scenario cards. For each one: choose the best-fit strength.",
      "Agree on a reason for your choice using your Character Strengths handout.",
      "Prepare ONE piece of advice you could give that character. Be ready to share.",
    ],
    NOTES_YOUDO,
    FOOTER
  );

  // ── Slide 11: Share back ──
  pairShareSlide(
    pres,
    "Share what you found",
    [
      "What strengths did your pair choose? Why?",
      "Did any pair choose a DIFFERENT strength for a similar situation? What was their reason?",
    ],
    NOTES_SHARE,
    FOOTER
  );

  // ── Slide 12: Closing ──
  closingSlide(
    pres,
    {
      reflectionPrompt: "Pick ONE character strength you could use THIS WEEK. Whisper it to your partner.",
      scItems: [
        "I can name a character strength from the list.",
        "I can match a character strength to an everyday situation and say why.",
        "I can give a piece of advice that uses that strength in action.",
      ],
      selfAssessment: {
        prompt: "Thumbs check on each I can statement",
        options: ["Thumbs up", "Thumbs sideways", "Thumbs down"],
      },
      takeaways: [
        "Different people pick different strengths for the same challenge — and that is okay.",
        "Strengths are most useful when they turn into specific actions.",
      ],
    },
    NOTES_CLOSING
  );

  // ── Write PPTX ──
  const pptxPath = path.join(LESSON_FOLDER, PPTX_NAME);
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to", pptxPath);
}

build().catch((err) => {
  console.error(err);
  process.exit(1);
});
