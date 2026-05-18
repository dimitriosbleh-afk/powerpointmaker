"use strict";

// Respectful Relationships - Role-Playing Strengths
// Grade 6 Wellbeing - 15+ minute session (Activity 3 in the strengths sequence)
// Students plan and perform short role plays that show character strengths in action.

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

// Theme — match the existing RR Character Strengths series variant for cohesion
const T = createTheme("wellbeing", "grade56", weekToVariant(1));
const {
  C, FONT_H, FONT_B,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  pairShareSlide,
  addTopBar, addBadge, addTitle, addCard, addFooter, addTextOnShape,
  CONTENT_TOP, SAFE_BOTTOM,
} = T;

// Output paths
const UNIT = "RR_Role_Play_Strengths";
const LESSON_FOLDER = path.join(__dirname, "..", "output", UNIT);
const PPTX_NAME = "Role-Playing Strengths.pptx";
const FOOTER = "Respectful Relationships | Grade 6 Wellbeing";
const SESSION = 3;
const RES_DIR = path.join(LESSON_FOLDER, getSessionResourceFolder(SESSION));

// One new resource. Character Strengths Handout is reused from Activity 1
// (already in the unit) and is not regenerated here.
const PLANNER_RESOURCE = makeSessionResource(
  SESSION,
  "Role Play Planning Card",
  "Group-of-four planner: setting, four roles, the strength each character shows, and what they will say or do."
);
const RESOURCE_ITEMS = [PLANNER_RESOURCE];

fs.mkdirSync(RES_DIR, { recursive: true });

// ============================================================
// Teacher Notes
// ============================================================

const NOTES_TITLE = [
  "SAY:",
  "- Today we are going to bring character strengths to life by acting them out.",
  "- You will work in groups of four to plan a short scene that shows your character using a strength.",
  "",
  "DO:",
  "- Display slide as students arrive.",
  "- Have Role Play Planning Cards ready, plus the Character Strengths Handout from Activity 1 within easy reach.",
  "",
  "TEACHER NOTES:",
  "This is a short 15-20 minute session. The thinking happens in planning and watching, not just in performing. Keep the energy warm and low-stakes.",
  "",
  "WATCH FOR:",
  "- Students who are nervous about performing -- reassure them that the audience is on their side.",
  "",
  "[General: Title | VTLM 2.0: Establishing Purpose]",
].join("\n");

const NOTES_RESOURCES = [
  "SAY:",
  "- These are the materials we need for the session.",
  "",
  "DO:",
  "- Print one Role Play Planning Card per group of four before the session.",
  "- Have the Character Strengths Handout from Activity 1 ready as a reference for each group.",
  "- Clear a small open space in the room for performing scenes.",
  "",
  "TEACHER NOTES:",
  "Groups need the Activity 1 handout open beside them while planning. The planning card structures the scene; the handout reminds them which strengths are available.",
  "",
  "WATCH FOR:",
  "- Every group has a planner and access to a strengths list before planning starts.",
  "",
  "[General: Resources | VTLM 2.0: Preparation]",
].join("\n");

const NOTES_LAUNCH = [
  "SAY:",
  "- Last session you matched character strengths to everyday challenges in pairs.",
  "- Today we go one step further. We are going to act those strengths out so we can see what they look like in real moments.",
  "- Quick recall: name one character strength you remember from last session. Whisper it to your partner.",
  "- Ask one or two pairs to share. Record their words on the board.",
  "",
  "DO:",
  "- Give 30 seconds for the whisper-share.",
  "- Cold call two pairs to share a strength word.",
  "- Write the words on the board as a live word bank for today.",
  "",
  "TEACHER NOTES:",
  "This launch activates the prior session and signals the shift from talking about strengths to showing them. Keep it brisk -- aim for about 90 seconds.",
  "",
  "WATCH FOR:",
  "- Students who cannot recall a strength -- direct them to the Character Strengths Handout from Activity 1.",
  "",
  "[General: Launch | VTLM 2.0: Activating Prior Knowledge]",
].join("\n");

const NOTES_VOCAB = [
  "SAY:",
  "- Two words to keep in mind today.",
  "- A role play is a short pretend scene where we act out a situation.",
  "- A character strength is part of who someone is on the inside -- like kindness, courage, fairness, or tolerance.",
  "- Quick check: thumbs up if a role play is a scene we plan and act out.",
  "",
  "DO:",
  "- Read the two definitions aloud, pointing to each card.",
  "- Ask students to repeat the word 'tolerance' chorally if it is new -- this is one of the strengths in today's examples.",
  "",
  "TEACHER NOTES:",
  "Keep this short. The aim is shared vocabulary, not a full vocab lesson. Tolerance and fairness are the two strengths the source examples use, so they get a quick spotlight.",
  "",
  "WATCH FOR:",
  "- Students confusing role play with reading a script -- emphasise that students invent the lines themselves.",
  "",
  "[General: Vocabulary | VTLM 2.0: Explicit Explanation]",
].join("\n");

const NOTES_LI = [
  "SAY:",
  "- Read the learning intention from the slide.",
  "- Read each success criterion aloud. We will check ourselves against these at the end.",
  "",
  "DO:",
  "- Point to each success criterion as you read it.",
  "- Leave the slide visible for 15-20 seconds so students can read it themselves.",
  "",
  "TEACHER NOTES:",
  "SC1 is the ultra-achievable floor (everyone in a group of four can name and use one strength). SC2 is the core target the exit reflection assesses. SC3 stretches students who finish quickly into noticing strengths in others.",
  "",
  "WATCH FOR:",
  "- Students unsure what 'show in action' means -- the I Do slide makes this concrete.",
  "",
  "[General: LI/SC | VTLM 2.0: Learning Intentions]",
].join("\n");

const NOTES_IDO_FRAME = [
  "SAY:",
  "- A strong role play has four parts. We call this the SRSS frame.",
  "- Setting -- where is the scene happening?",
  "- Roles -- who are the four characters?",
  "- Strengths -- which strength does each character show?",
  "- Situation -- what is the problem or moment they are working through?",
  "- The audience will be trying to spot each character's strength, so the strength has to be visible in what the character says or does.",
  "",
  "DO:",
  "- Point to each part of the SRSS frame as you read it.",
  "- Say: 'If your audience cannot guess your strength, your scene needs to show it more clearly.'",
  "",
  "TEACHER NOTES:",
  "The SRSS frame gives students a planning scaffold so they do not just freestyle. The next slide shows a worked example using this frame.",
  "",
  "WATCH FOR:",
  "- Students treating role play as a comedy show -- redirect to the strength being shown.",
  "",
  "[General: I Do | VTLM 2.0: Explicit Explanation]",
].join("\n");

const NOTES_IDO_EXAMPLE = [
  "SAY:",
  "- Here is one worked example before you plan your own.",
  "- Setting: at the park. Roles: four friends. Strengths: tolerance, fairness, humour, leadership. Situation: they all want to play different games.",
  "- Watch how each character's strength shows up in what they say.",
  "- Read the lines on the right with expression.",
  "- Ask: which line shows tolerance? Which line shows fairness?",
  "",
  "DO:",
  "- Read the four sample lines with different voices.",
  "- Point to each line and identify the strength being shown.",
  "- Highlight that the strength is shown by ACTION or WORDS, not announced ('Hi, I am being tolerant!').",
  "",
  "TEACHER NOTES:",
  "This example uses the TOLERANCE situation from the source coaching prompts. It is a model only -- groups choose their own scenarios. The example is closely matched but uses different strengths than the We Do, so students still need to think.",
  "",
  "MISCONCEPTIONS:",
  "- Misconception: A character can announce their strength by naming it ('I am being kind').",
  "  Why: Students often equate showing with telling.",
  "  Impact: Scenes become flat narration instead of action.",
  "  Quick correction: 'Show me kindness without using the word kind. What might that look like?'",
  "",
  "WATCH FOR:",
  "- Students who copy the example word for word -- nudge them toward their own situation.",
  "",
  "[General: I Do | VTLM 2.0: Worked Example]",
].join("\n");

const NOTES_WEDO = [
  "SAY:",
  "- Let's plan one together using the SRSS frame.",
  "- Setting: at home. Roles: four family members. Situation: working out whose job it should be to wash the dishes.",
  "- What four strengths could each family member show in this scene?",
  "- Turn to your partner: which strength would help here? Why?",
  "- Cold call two pairs. Record their suggestions on the board next to each role.",
  "- Build the planning out loud as students contribute.",
  "",
  "DO:",
  "- Use the Activity 1 Character Strengths Handout list as a reference.",
  "- Co-construct the scene on the board using the right-hand frame from the slide.",
  "- Once strengths are chosen, ask one student to suggest a line of dialogue that shows fairness in action.",
  "",
  "CFU CHECKPOINT:",
  "Technique: Turn & Talk with Cold Call.",
  "Script:",
  "- Say: With your partner, choose two strengths that could help the family solve this fairly. 30 seconds.",
  "- Cold call two pairs. Listen for strengths that suit the situation (fairness, teamwork, leadership, kindness).",
  "- Scan for: students naming a strength AND giving a brief reason.",
  "PROCEED: If most pairs name a relevant strength and a reason, move to the spot-the-strength CFU.",
  "PIVOT: If pairs name a strength but cannot say why, model: 'Fairness fits because the job needs to be shared evenly -- here is what a fair character might say.' Re-ask one new pair.",
  "",
  "TEACHER NOTES:",
  "This is the FAIRNESS situation from the source coaching prompts. Use it deliberately so the class shares one worked example before group planning. Different strengths from the I Do keeps the thinking active.",
  "",
  "ENABLING & EXTENDING:",
  "ENABLING PROMPT:",
  "- Task: Students who struggle choose just one strength from a three-word menu on the board (fairness, teamwork, kindness) and explain who in the family would show it.",
  "- Extra Notes: Sentence stem: 'The ___ might show ___ by saying ___.'",
  "EXTENDING PROMPT:",
  "- Task: Students suggest a TENSION line -- a line a character might say if they were not using their strength -- and a strength line that turns the moment around.",
  "",
  "WATCH FOR:",
  "- Students suggesting unrelated strengths -- ask: 'How would that help with the dishes?'",
  "- Students who only name one strength for the whole group -- push for one per character.",
  "",
  "[General: We Do | VTLM 2.0: Scaffold Practice]",
].join("\n");

const NOTES_CFU = [
  "SAY:",
  "- Quick check before you plan your own scene.",
  "- Read the short scene on the slide.",
  "- Which strength is Jordan showing? Hold up your fingers: 1 for Honesty, 2 for Fairness, 3 for Courage, 4 for Tolerance.",
  "- After fingers are up, ask one student to justify.",
  "",
  "DO:",
  "- Read the scene aloud calmly.",
  "- Pause 5 seconds for thinking time before fingers up.",
  "- Scan the room. Look for at least 80% choosing 3 (Courage).",
  "",
  "CFU CHECKPOINT:",
  "Technique: Finger Voting.",
  "Script:",
  "- Say: 1 to 4 fingers, on three. Three.",
  "- Scan for: most fingers on 3.",
  "PROCEED: If most students show 3 and can justify, move to the You Do.",
  "PIVOT: If many show 2 or 4 (fairness or tolerance), explain: 'Jordan is doing something difficult and right when no one else will -- that is courage. Fairness would be about sharing or treating people equally.' Ask a fresh scene-spotting question.",
  "",
  "TEACHER NOTES:",
  "This hinge question tests whether students can read a strength from an action rather than a label. The wrong options are plausible: students who confuse 'standing up' with fairness or tolerance will pick 2 or 4.",
  "",
  "WATCH FOR:",
  "- Students changing their finger count after seeing peers -- ask them to commit before scanning.",
  "",
  "[General: CFU | VTLM 2.0: Checking for Understanding]",
].join("\n");

const NOTES_YOUDO_TASK = [
  "SAY:",
  "- Now it is your turn.",
  "- In your group of four, plan a short role play using the SRSS frame.",
  "- Each of you picks one strength from the Character Strengths Handout.",
  "- Your scene should show all four strengths in action.",
  "- Read the situation options on the slide -- choose one or invent your own.",
  "- You will have about 5 minutes to plan. Aim for a scene that lasts about 1 minute when you perform it.",
  "",
  "DO:",
  "- Group students into fours (use existing classroom groupings).",
  "- Distribute one Role Play Planning Card per group.",
  "- Make sure each group has the Character Strengths Handout from Activity 1.",
  "- Circulate. Push groups to commit to four DIFFERENT strengths, not the same one for everyone.",
  "",
  "TEACHER NOTES:",
  "This slide sets up the task. The next slide gives the planning structure. Allow students to invent their own situation if the listed ones do not fit, but they must keep the four-strengths rule.",
  "",
  "ENABLING & EXTENDING:",
  "ENABLING PROMPT:",
  "- Task: Group uses the FAIRNESS-dishes scenario from the We Do and chooses four strengths that fit. The teacher visits to check the four strengths are different.",
  "- Extra Notes: Pair this group with a confident scribe for the planner.",
  "EXTENDING PROMPT:",
  "- Task: Group plans a scene where one character starts WITHOUT using their strength and learns to use it during the scene. Audience identifies the shift.",
  "",
  "WATCH FOR:",
  "- Groups choosing the same strength for everyone -- redirect to four different strengths.",
  "- Off-task behaviour during planning -- timer and gentle check-ins.",
  "",
  "[General: You Do | VTLM 2.0: Independent Practice]",
].join("\n");

const NOTES_YOUDO_PLAN = [
  "SAY:",
  "- Use your Planning Card to organise the scene.",
  "- Setting: where does the scene happen?",
  "- Roles: who is each character?",
  "- Strengths: which strength does each character show? Four different strengths -- one per person.",
  "- Situation: what is the problem or moment?",
  "- Lines: what could each character say or do that shows their strength?",
  "- Five minutes to plan. Two minutes to rehearse.",
  "",
  "DO:",
  "- Start a visible timer for 5 minutes of planning.",
  "- Circulate. Read planners over shoulders. Ask one targeted question per group.",
  "- After 5 minutes, give a 2-minute rehearsal window before performances.",
  "",
  "TEACHER NOTES:",
  "Planning is the heart of the lesson. Resist the urge to over-coach -- ask one good question, then move on. Look for the alignment between strength named and what the character is actually doing.",
  "",
  "WATCH FOR:",
  "- Groups skipping the lines column -- the lines are where the strength becomes visible.",
  "- Groups treating the planner as a script to read out word for word -- it is a scaffold, not a script.",
  "",
  "[General: You Do | VTLM 2.0: Guided Independence]",
].join("\n");

const NOTES_SHARE = [
  "SAY:",
  "- Groups perform one at a time. Audience watches silently.",
  "- Audience: try to spot each character's strength as they perform.",
  "- After each scene we will ask three questions. Be ready to share what you noticed.",
  "- Read the three audience questions from the slide.",
  "",
  "DO:",
  "- Invite each group to perform in turn (about 1 minute per scene).",
  "- After each scene, run the three discussion questions briskly.",
  "- Validate every strength named by the audience even if it is not the one the actor planned -- this often reveals a second strength.",
  "- Keep transitions warm with a quick clap before the next group.",
  "",
  "TEACHER NOTES:",
  "These three questions are taken from the source activity. They move the thinking from spotting (Where) to value (How was it useful) to transfer (Real life). Aim for one or two voices per question, not a long discussion.",
  "",
  "SENSITIVITY ADVISORY:",
  "- What it is: Some scenes may touch on real classroom moments, conflicts, or family situations.",
  "- Framing language: 'Scenes are pretend. We are not naming anyone real in the class.'",
  "- Watch for: A student looking uncomfortable during a peer's scene.",
  "- Protocol: Pause and check in privately. Allow opt-out from audience or performance.",
  "",
  "WATCH FOR:",
  "- Audience members who go silent -- direct a specific question to a specific student.",
  "- Performers being teased or laughed at -- pause and reset class norms.",
  "",
  "[General: We Do / Share | VTLM 2.0: Demonstrate and Apply]",
].join("\n");

const NOTES_CLOSING = [
  "SAY:",
  "- Look back at the success criteria. Show me with thumbs how you went.",
  "- Read each criterion aloud and pause for thumbs.",
  "- Turn and tell your partner: which character strength did YOU show in your role play today?",
  "- Challenge: notice one strength in someone in your group on the way out and tell them.",
  "",
  "DO:",
  "- Read each SC aloud. Pause for thumbs up, sideways, or down.",
  "- Allow 30 seconds for the Turn & Talk.",
  "- Acknowledge the courage it takes to perform in front of peers.",
  "",
  "TEACHER NOTES:",
  "The closing turns the spotlight back on the student. Self-assessment uses the same three plain SC bullets that opened the lesson.",
  "",
  "WATCH FOR:",
  "- Students who show thumbs down on SC2 -- note for follow-up in the next session.",
  "",
  "[General: Closing | VTLM 2.0: Review and Reflect]",
].join("\n");

// ============================================================
// Common LI / SC strings (referenced in slide + closing)
// ============================================================

const LI_TEXT = "We are learning to show what character strengths look like in action through role play.";
const SC_ITEMS = [
  "I can name one character strength and what it might look like.",
  "I can show a character strength in action in our group role play.",
  "I can spot strengths in other groups' role plays and explain how the strength helped.",
];

// ============================================================
// Build function
// ============================================================

async function build() {
  const pres = new PptxGenJS();
  pres.layout = "LAYOUT_16x9";

  // -- Slide 1: Title --
  titleSlide(
    pres,
    "Role-Playing\nStrengths",
    "Respectful Relationships",
    "Grade 6 Wellbeing  |  Session 3",
    NOTES_TITLE
  );

  // -- Slide 2: Teacher Resources (immediately after title) --
  addResourceSlide(pres, RESOURCE_ITEMS, T, FOOTER, NOTES_RESOURCES);

  // -- Slide 3: Launch (connect prior session to today) --
  contentSlide(
    pres,
    "Launch",
    C.SECONDARY,
    "From Talking About Strengths to Showing Them",
    [
      "Last session: matched strengths to everyday challenges",
      "Today: act those strengths out in a short scene",
      "Whisper to your partner: one strength you remember",
    ],
    NOTES_LAUNCH,
    FOOTER,
    (s, lg) => {
      const rX = lg.rightX;
      const rW = lg.rightW;
      const topY = lg.panelTopPadded;
      const cardH = 3.3;

      addCard(s, rX, topY, rW, cardH, { strip: C.PRIMARY, fill: C.WHITE });

      s.addText("From Knowing to Doing", {
        x: rX + 0.15, y: topY + 0.12, w: rW - 0.3, h: 0.36,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true, align: "center", margin: 0,
      });

      // Two pill stages with an arrow between
      const stageY = topY + 0.7;
      const stageH = 0.85;
      const stageW = (rW - 0.6) / 2;

      // Stage 1
      s.addShape("roundRect", {
        x: rX + 0.15, y: stageY, w: stageW, h: stageH, rectRadius: 0.10,
        fill: { color: C.BG_LIGHT }, line: { color: C.PRIMARY, width: 1 },
      });
      s.addText("Last Session", {
        x: rX + 0.15, y: stageY + 0.08, w: stageW, h: 0.26,
        fontSize: 11, fontFace: FONT_H, color: C.PRIMARY, bold: true, align: "center", margin: 0,
      });
      s.addText("Match strengths to scenarios", {
        x: rX + 0.20, y: stageY + 0.36, w: stageW - 0.1, h: 0.42,
        fontSize: 10.5, fontFace: FONT_B, color: C.CHARCOAL, align: "center", valign: "middle", margin: 0,
      });

      // Arrow
      s.addShape("rightTriangle", {
        x: rX + 0.15 + stageW + 0.05, y: stageY + stageH / 2 - 0.18, w: 0.30, h: 0.36,
        fill: { color: C.ACCENT }, line: { color: C.ACCENT, width: 0 },
        rotate: 90,
      });

      // Stage 2
      s.addShape("roundRect", {
        x: rX + 0.15 + stageW + 0.30, y: stageY, w: stageW, h: stageH, rectRadius: 0.10,
        fill: { color: C.ACCENT },
      });
      s.addText("Today", {
        x: rX + 0.15 + stageW + 0.30, y: stageY + 0.08, w: stageW, h: 0.26,
        fontSize: 11, fontFace: FONT_H, color: C.WHITE, bold: true, align: "center", margin: 0,
      });
      s.addText("Show strengths in action", {
        x: rX + 0.20 + stageW + 0.30, y: stageY + 0.36, w: stageW - 0.1, h: 0.42,
        fontSize: 10.5, fontFace: FONT_B, color: C.WHITE, align: "center", valign: "middle", margin: 0,
      });

      // Class word bank prompt
      s.addText("Class word bank: write recalled strengths on the board", {
        x: rX + 0.15, y: topY + cardH - 0.85, w: rW - 0.3, h: 0.7,
        fontSize: 11, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
        align: "center", valign: "middle", margin: 0,
      });
    }
  );

  // -- Slide 4: Key Vocabulary --
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "Vocabulary", { color: C.PRIMARY });
    addTitle(s, "Two Words for Today");

    const cardY = CONTENT_TOP + 0.20;
    const cardH = 3.0;
    const cardW = 4.30;
    const gap = 0.40;
    const startX = (10 - (cardW * 2 + gap)) / 2;

    // Card 1: Role Play
    addCard(s, startX, cardY, cardW, cardH, { strip: C.PRIMARY, fill: C.WHITE });
    s.addText("Role Play", {
      x: startX + 0.2, y: cardY + 0.18, w: cardW - 0.4, h: 0.55,
      fontSize: 26, fontFace: FONT_H, color: C.PRIMARY, bold: true, align: "center", margin: 0,
    });
    // Simple visual: two stick figures
    {
      const sceneY = cardY + 0.95;
      const cx = startX + cardW / 2;
      // Figure A
      s.addShape("roundRect", {
        x: cx - 0.60, y: sceneY, w: 0.30, h: 0.30, rectRadius: 0.15,
        line: { color: C.PRIMARY, width: 1.5 }, fill: { color: C.WHITE },
      });
      s.addShape("line", { x: cx - 0.45, y: sceneY + 0.30, w: 0, h: 0.5, line: { color: C.PRIMARY, width: 1.5 } });
      s.addShape("line", { x: cx - 0.45, y: sceneY + 0.50, w: -0.20, h: 0, line: { color: C.PRIMARY, width: 1.5 } });
      s.addShape("line", { x: cx - 0.45, y: sceneY + 0.50, w: 0.20, h: 0, line: { color: C.PRIMARY, width: 1.5 } });
      s.addShape("line", { x: cx - 0.45, y: sceneY + 0.80, w: -0.15, h: 0.30, line: { color: C.PRIMARY, width: 1.5 } });
      s.addShape("line", { x: cx - 0.45, y: sceneY + 0.80, w: 0.15, h: 0.30, line: { color: C.PRIMARY, width: 1.5 } });
      // Figure B
      s.addShape("roundRect", {
        x: cx + 0.30, y: sceneY, w: 0.30, h: 0.30, rectRadius: 0.15,
        line: { color: C.ACCENT, width: 1.5 }, fill: { color: C.WHITE },
      });
      s.addShape("line", { x: cx + 0.45, y: sceneY + 0.30, w: 0, h: 0.5, line: { color: C.ACCENT, width: 1.5 } });
      s.addShape("line", { x: cx + 0.45, y: sceneY + 0.50, w: -0.20, h: 0, line: { color: C.ACCENT, width: 1.5 } });
      s.addShape("line", { x: cx + 0.45, y: sceneY + 0.50, w: 0.20, h: 0, line: { color: C.ACCENT, width: 1.5 } });
      s.addShape("line", { x: cx + 0.45, y: sceneY + 0.80, w: -0.15, h: 0.30, line: { color: C.ACCENT, width: 1.5 } });
      s.addShape("line", { x: cx + 0.45, y: sceneY + 0.80, w: 0.15, h: 0.30, line: { color: C.ACCENT, width: 1.5 } });
      // Speech curve
      s.addShape("roundRect", {
        x: cx - 0.20, y: sceneY - 0.15, w: 0.4, h: 0.22, rectRadius: 0.10,
        fill: { color: C.BG_LIGHT }, line: { color: C.MUTED, width: 0.8 },
      });
    }
    s.addText("A short pretend scene where we act out a situation", {
      x: startX + 0.2, y: cardY + cardH - 0.70, w: cardW - 0.4, h: 0.6,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
      align: "center", valign: "middle", margin: 0,
    });

    // Card 2: Character Strength
    const c2X = startX + cardW + gap;
    addCard(s, c2X, cardY, cardW, cardH, { strip: C.ACCENT, fill: C.WHITE });
    s.addText("Character Strength", {
      x: c2X + 0.2, y: cardY + 0.18, w: cardW - 0.4, h: 0.55,
      fontSize: 23, fontFace: FONT_H, color: C.ACCENT, bold: true, align: "center", margin: 0,
    });
    // Simple visual: heart-on-figure (who you are inside)
    {
      const cx = c2X + cardW / 2;
      const sceneY = cardY + 0.95;
      s.addShape("roundRect", {
        x: cx - 0.20, y: sceneY, w: 0.40, h: 0.40, rectRadius: 0.20,
        line: { color: C.PRIMARY, width: 1.5 }, fill: { color: C.WHITE },
      });
      s.addShape("line", { x: cx, y: sceneY + 0.40, w: 0, h: 0.55, line: { color: C.PRIMARY, width: 1.5 } });
      s.addShape("line", { x: cx, y: sceneY + 0.62, w: -0.30, h: 0, line: { color: C.PRIMARY, width: 1.5 } });
      s.addShape("line", { x: cx, y: sceneY + 0.62, w: 0.30, h: 0, line: { color: C.PRIMARY, width: 1.5 } });
      s.addShape("line", { x: cx, y: sceneY + 0.95, w: -0.18, h: 0.30, line: { color: C.PRIMARY, width: 1.5 } });
      s.addShape("line", { x: cx, y: sceneY + 0.95, w: 0.18, h: 0.30, line: { color: C.PRIMARY, width: 1.5 } });
      // Heart shape (rounded square with rotated diamond)
      s.addShape("heart", {
        x: cx - 0.13, y: sceneY + 0.50, w: 0.26, h: 0.22,
        fill: { color: C.ACCENT }, line: { color: C.ACCENT, width: 0 },
      });
    }
    s.addText("Part of WHO you are: kindness, courage, fairness, tolerance...", {
      x: c2X + 0.2, y: cardY + cardH - 0.70, w: cardW - 0.4, h: 0.6,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
      align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_VOCAB);
  }

  // -- Slide 5: LI / SC --
  liSlide(
    pres,
    [LI_TEXT],
    SC_ITEMS,
    NOTES_LI,
    FOOTER
  );

  // -- Slide 6: I Do - SRSS Frame --
  contentSlide(
    pres,
    "I Do",
    C.PRIMARY,
    "What a Strong Role Play Has",
    [
      "Setting -- where the scene happens",
      "Roles -- who each character is",
      "Strengths -- the strength each character shows",
      "Situation -- the moment or problem they're working through",
      "Audience must SEE the strength in what each character says or does",
    ],
    NOTES_IDO_FRAME,
    FOOTER,
    (s, lg) => {
      const rX = lg.rightX;
      const rW = lg.rightW;
      const topY = lg.panelTopPadded;
      const cardH = 3.3;

      addCard(s, rX, topY, rW, cardH, { strip: C.ACCENT, fill: C.WHITE });

      s.addText("The SRSS Frame", {
        x: rX + 0.15, y: topY + 0.10, w: rW - 0.3, h: 0.32,
        fontSize: 14, fontFace: FONT_H, color: C.ACCENT, bold: true, align: "center", margin: 0,
      });

      const rows = [
        { letter: "S", word: "Setting",   detail: "Where?" },
        { letter: "R", word: "Roles",     detail: "Who? (x4)" },
        { letter: "S", word: "Strengths", detail: "Which strength each?" },
        { letter: "S", word: "Situation", detail: "What's happening?" },
      ];

      const rowStartY = topY + 0.55;
      const rowH = 0.58;
      const rowGap = 0.08;
      const letterW = 0.55;

      rows.forEach((r, i) => {
        const rY = rowStartY + i * (rowH + rowGap);
        // Letter box
        s.addShape("roundRect", {
          x: rX + 0.15, y: rY, w: letterW, h: rowH, rectRadius: 0.08,
          fill: { color: C.PRIMARY },
        });
        s.addText(r.letter, {
          x: rX + 0.15, y: rY, w: letterW, h: rowH,
          fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true,
          align: "center", valign: "middle", margin: 0,
        });
        // Word
        s.addText(r.word, {
          x: rX + 0.15 + letterW + 0.12, y: rY + 0.04, w: rW - letterW - 0.45, h: 0.28,
          fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true, valign: "middle", margin: 0,
        });
        // Detail
        s.addText(r.detail, {
          x: rX + 0.15 + letterW + 0.12, y: rY + 0.30, w: rW - letterW - 0.45, h: 0.24,
          fontSize: 11, fontFace: FONT_B, color: C.CHARCOAL, italic: true, valign: "top", margin: 0,
        });
      });
    }
  );

  // -- Slide 7: I Do - Worked Example (TOLERANCE) --
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "I Do", { color: C.PRIMARY });
    addTitle(s, "Worked Example: Friends at the Park");

    // Left scene-setup card
    const leftX = 0.5;
    const leftW = 4.4;
    const cardY = CONTENT_TOP;
    const cardH = 3.4;

    addCard(s, leftX, cardY, leftW, cardH, { strip: C.PRIMARY, fill: C.WHITE });
    s.addText("Scene Setup", {
      x: leftX + 0.20, y: cardY + 0.12, w: leftW - 0.4, h: 0.30,
      fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0,
    });

    const setupRows = [
      { label: "Setting:",   value: "At the park" },
      { label: "Roles:",     value: "Four friends" },
      { label: "Situation:", value: "They want to play different games" },
      { label: "Strengths:", value: "Tolerance, Fairness, Humour, Leadership" },
    ];

    setupRows.forEach((row, i) => {
      const rY = cardY + 0.55 + i * 0.66;
      s.addText(row.label, {
        x: leftX + 0.20, y: rY, w: 1.3, h: 0.30,
        fontSize: 13, fontFace: FONT_H, color: C.ACCENT, bold: true, valign: "top", margin: 0,
      });
      s.addText(row.value, {
        x: leftX + 0.20, y: rY + 0.28, w: leftW - 0.4, h: 0.36,
        fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, valign: "top", margin: 0,
      });
    });

    // Right - sample dialogue lines, each labelled with its strength
    const rightX = 5.2;
    const rightW = 4.3;
    addCard(s, rightX, cardY, rightW, cardH, { strip: C.ACCENT, fill: C.WHITE });
    s.addText("What Each Character Shows", {
      x: rightX + 0.15, y: cardY + 0.12, w: rightW - 0.30, h: 0.30,
      fontSize: 13.5, fontFace: FONT_H, color: C.ACCENT, bold: true, align: "center", margin: 0,
    });

    const lines = [
      { tag: "Tolerance",  text: "\"I don't love that game, but I'll give it a go for you.\"" },
      { tag: "Fairness",   text: "\"Let's each pick one game and play them in turn.\"" },
      { tag: "Humour",     text: "\"Loser of round one carries the bag!\" (smiles)" },
      { tag: "Leadership", text: "\"Okay -- soccer first, then handball. Sound good?\"" },
    ];

    const lineStartY = cardY + 0.52;
    const lineH = 0.65;
    const lineGap = 0.05;

    lines.forEach((ln, i) => {
      const lY = lineStartY + i * (lineH + lineGap);
      // Tag pill
      s.addShape("roundRect", {
        x: rightX + 0.15, y: lY + 0.05, w: 1.1, h: 0.30, rectRadius: 0.06,
        fill: { color: C.PRIMARY },
      });
      s.addText(ln.tag, {
        x: rightX + 0.15, y: lY + 0.05, w: 1.1, h: 0.30,
        fontSize: 10.5, fontFace: FONT_B, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      // Line text
      s.addText(ln.text, {
        x: rightX + 0.15 + 1.20, y: lY, w: rightW - 1.45, h: lineH,
        fontSize: 11.5, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
        valign: "middle", margin: 0,
      });
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO_EXAMPLE);
  }

  // -- Slide 8: We Do - Plan Together (FAIRNESS-dishes) --
  contentSlide(
    pres,
    "We Do",
    C.SECONDARY,
    "Plan One Together: The Dishes",
    [
      "Setting: at home",
      "Roles: four family members",
      "Situation: working out whose job it should be to wash the dishes",
      "Turn & Talk: which strengths would help here?",
    ],
    NOTES_WEDO,
    FOOTER,
    (s, lg) => {
      const rX = lg.rightX;
      const rW = lg.rightW;
      const topY = lg.panelTopPadded;
      const cardH = 3.3;

      addCard(s, rX, topY, rW, cardH, { strip: C.SECONDARY, fill: C.WHITE });

      s.addText("Class Plan", {
        x: rX + 0.15, y: topY + 0.10, w: rW - 0.3, h: 0.30,
        fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true, align: "center", margin: 0,
      });

      // Four character rows, each with a blank for the strength
      const charRows = [
        "Family member 1: __________",
        "Family member 2: __________",
        "Family member 3: __________",
        "Family member 4: __________",
      ];
      const rowStartY = topY + 0.55;
      const rowH = 0.42;
      const rowGap = 0.10;

      charRows.forEach((label, i) => {
        const rY = rowStartY + i * (rowH + rowGap);
        s.addShape("roundRect", {
          x: rX + 0.15, y: rY, w: rW - 0.30, h: rowH, rectRadius: 0.08,
          fill: { color: C.BG_LIGHT }, line: { color: C.MUTED, width: 0.8 },
        });
        s.addText(label, {
          x: rX + 0.25, y: rY, w: rW - 0.50, h: rowH,
          fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, valign: "middle", margin: 0,
        });
      });

      // Hint banner
      const hintY = rowStartY + 4 * (rowH + rowGap) + 0.02;
      const hintH = cardH - (hintY - topY) - 0.15;
      if (hintH > 0.30) {
        s.addShape("roundRect", {
          x: rX + 0.15, y: hintY, w: rW - 0.30, h: hintH, rectRadius: 0.06,
          fill: { color: C.ACCENT },
        });
        s.addText("Build it on the board as students share", {
          x: rX + 0.15, y: hintY, w: rW - 0.30, h: hintH,
          fontSize: 11, fontFace: FONT_B, color: C.WHITE, italic: true,
          align: "center", valign: "middle", margin: 0,
        });
      }
    }
  );

  // -- Slide 9: CFU - Spot the Strength --
  cfuSlide(
    pres,
    "CFU",
    "Spot the Strength",
    "Finger Voting",
    "Jordan sees a new student sitting alone at lunch. The other kids ignore them. Jordan walks over, sits down, and starts a chat. Which strength is Jordan showing?  1: Honesty   2: Fairness   3: Courage   4: Tolerance",
    NOTES_CFU,
    FOOTER
  );

  // -- Slide 10: You Do - Group Task setup --
  contentSlide(
    pres,
    "You Do",
    C.ACCENT,
    "Your Group's Role Play",
    [
      "First: get into your group of four",
      "Next: each person picks ONE strength from the handout",
      "Then: plan a short scene where each character shows their strength",
      "Aim for a 1-minute scene with all four strengths visible",
    ],
    NOTES_YOUDO_TASK,
    FOOTER,
    (s, lg) => {
      const rX = lg.rightX;
      const rW = lg.rightW;
      const topY = lg.panelTopPadded;
      const cardH = 3.3;

      addCard(s, rX, topY, rW, cardH, { strip: C.PRIMARY, fill: C.WHITE });

      s.addText("Pick a Situation", {
        x: rX + 0.15, y: topY + 0.10, w: rW - 0.3, h: 0.30,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true, align: "center", margin: 0,
      });

      const ideas = [
        "Friends deciding which game to play",
        "Family sharing a household job fairly",
        "A new student joining your group",
        "Working out a small disagreement at lunch",
        "Or invent your own everyday moment",
      ];

      const ideaStartY = topY + 0.55;
      const ideaH = 0.45;
      const ideaGap = 0.06;

      ideas.forEach((idea, i) => {
        const iY = ideaStartY + i * (ideaH + ideaGap);
        s.addShape("roundRect", {
          x: rX + 0.15, y: iY, w: rW - 0.30, h: ideaH, rectRadius: 0.06,
          fill: { color: C.BG_LIGHT }, line: { color: C.MUTED, width: 0.8 },
        });
        // bullet dot
        s.addShape("roundRect", {
          x: rX + 0.28, y: iY + ideaH / 2 - 0.07, w: 0.14, h: 0.14, rectRadius: 0.07,
          fill: { color: C.ACCENT },
        });
        s.addText(idea, {
          x: rX + 0.52, y: iY, w: rW - 0.70, h: ideaH,
          fontSize: 11.5, fontFace: FONT_B, color: C.CHARCOAL, valign: "middle", margin: 0,
        });
      });
    }
  );

  // -- Slide 11: You Do - Plan the Scene --
  {
    const s = pres.addSlide();
    addTopBar(s, C.ACCENT);
    addBadge(s, "You Do", { color: C.ACCENT });
    addTitle(s, "Plan the Scene");

    // Instruction strip
    s.addText("Use your Planning Card. 5 minutes to plan, 2 minutes to rehearse.", {
      x: 0.5, y: CONTENT_TOP, w: 9, h: 0.35,
      fontSize: 13, fontFace: FONT_B, color: C.PRIMARY, italic: true, margin: 0,
    });

    // 4 numbered planning prompts in a 2x2 grid
    const gridY = CONTENT_TOP + 0.50;
    const gridH = SAFE_BOTTOM - gridY - 0.05;
    const cols = 2;
    const rows = 2;
    const gap = 0.18;
    const cellW = (9 - gap * (cols - 1)) / cols;
    const cellH = (gridH - gap * (rows - 1)) / rows;

    const prompts = [
      { num: "1", label: "Setting",   detail: "Where does the scene happen?" },
      { num: "2", label: "Roles",     detail: "Who are the four characters?" },
      { num: "3", label: "Strengths", detail: "Which strength does EACH character show? Four different strengths." },
      { num: "4", label: "Lines",     detail: "What might each character SAY or DO that shows their strength?" },
    ];

    prompts.forEach((p, i) => {
      const r = Math.floor(i / cols);
      const c = i % cols;
      const cx = 0.5 + c * (cellW + gap);
      const cy = gridY + r * (cellH + gap);

      addCard(s, cx, cy, cellW, cellH, { strip: C.ACCENT, fill: C.WHITE });

      // Number badge
      s.addShape("roundRect", {
        x: cx + 0.15, y: cy + 0.15, w: 0.55, h: 0.55, rectRadius: 0.28,
        fill: { color: C.PRIMARY },
      });
      s.addText(p.num, {
        x: cx + 0.15, y: cy + 0.15, w: 0.55, h: 0.55,
        fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      // Label
      s.addText(p.label, {
        x: cx + 0.85, y: cy + 0.18, w: cellW - 1.00, h: 0.36,
        fontSize: 16, fontFace: FONT_H, color: C.PRIMARY, bold: true, valign: "middle", margin: 0,
      });

      // Detail
      s.addText(p.detail, {
        x: cx + 0.20, y: cy + 0.80, w: cellW - 0.40, h: cellH - 0.95,
        fontSize: 12.5, fontFace: FONT_B, color: C.CHARCOAL, valign: "top", margin: 0,
      });
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO_PLAN);
  }

  // -- Slide 12: Share & Discuss --
  // After each scene the class works through the three source questions
  // (Where did you see it / How was it useful / In real life). The
  // "perform one at a time, audience watches silently" routine lives in
  // teacher notes -- pairShareSlide caps questions to 3 at this year band,
  // so all three source questions are preserved here.
  pairShareSlide(
    pres,
    "After Each Scene: Spot the Strengths",
    [
      "Where did you see the strength in action?",
      "How was the strength useful to the person or to the other characters?",
      "In real life, is it different? How could this strength be useful in real life?",
    ],
    NOTES_SHARE,
    FOOTER
  );

  // -- Slide 13: Closing --
  closingSlide(
    pres,
    {
      reflectionPrompt: "Which character strength did YOU show in your role play today? Tell your partner one moment you showed it.",
      scItems: SC_ITEMS,
      selfAssessment: {
        prompt: "Self-assess: thumbs up, sideways, or down for each criterion.",
        options: ["Got it", "Getting there", "Need more practice"],
      },
    },
    NOTES_CLOSING
  );

  // ============================================================
  // Write PPTX
  // ============================================================
  fs.mkdirSync(LESSON_FOLDER, { recursive: true });
  await pres.writeFile({ fileName: path.join(LESSON_FOLDER, PPTX_NAME) });
  console.log("PPTX written to", path.join(LESSON_FOLDER, PPTX_NAME));

  // ============================================================
  // Companion PDF: Role Play Planning Card
  // ============================================================
  {
    const doc = createPdf({ title: "Role Play Planning Card" });
    let y = addPdfHeader(doc, "Session 3 Role Play Planning Card", {
      subtitle: "Group of four: plan a short scene that shows each person's character strength",
      color: C.PRIMARY,
      lessonInfo: "Grade 6 Wellbeing - Respectful Relationships",
    });

    y = addTipBox(doc,
      "A strong role play uses the SRSS frame: Setting, Roles, Strengths, Situation. Your audience will try to spot each character's strength, so the strength has to show in what your character says or does.",
      y, { color: C.ACCENT });

    // Section 1: Scene setup
    y = addSectionHeading(doc, "Step 1: Set the scene", y, { color: C.PRIMARY });

    y = addBodyText(doc, "Setting (where does the scene happen?)", y, { fontSize: 10 });
    y = addLinedArea(doc, y, 1, { lineSpacing: 24 });
    y += 4;

    y = addBodyText(doc, "Situation (what is the moment or problem?)", y, { fontSize: 10 });
    y = addLinedArea(doc, y, 2, { lineSpacing: 22 });
    y += 6;

    // Section 2: Roles + strengths table
    y = addSectionHeading(doc, "Step 2: Roles and strengths", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Each person picks ONE different strength. Use the Character Strengths Handout from Activity 1 if you need a list.", y, { fontSize: 9.5, italic: true });

    // 4-row table: Name | Role | Strength
    const tableX = PAGE.MARGIN;
    const tableW = PAGE.CONTENT_W;
    const colNameW = 110;
    const colRoleW = 150;
    const colStrengthW = tableW - colNameW - colRoleW;
    const headH = 22;
    const rowH = 32;

    // Header
    doc.save();
    doc.rect(tableX, y, tableW, headH).fill(hex(C.PRIMARY));
    doc.restore();
    doc.fontSize(10).font("Sans-Bold").fillColor(hex(C.WHITE));
    doc.text("Name", tableX + 8, y + 6, { width: colNameW - 12 });
    doc.text("Role (who are they?)", tableX + colNameW + 8, y + 6, { width: colRoleW - 12 });
    doc.text("Strength", tableX + colNameW + colRoleW + 8, y + 6, { width: colStrengthW - 12 });

    y += headH;

    for (let i = 0; i < 4; i++) {
      doc.save();
      doc.rect(tableX, y, tableW, rowH)
        .lineWidth(0.8).strokeColor(hex(C.MUTED)).stroke();
      doc.moveTo(tableX + colNameW, y).lineTo(tableX + colNameW, y + rowH).stroke();
      doc.moveTo(tableX + colNameW + colRoleW, y).lineTo(tableX + colNameW + colRoleW, y + rowH).stroke();
      doc.restore();
      y += rowH;
    }
    y += 10;

    // Section 3: Lines
    y = addSectionHeading(doc, "Step 3: Lines that show the strength", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Write one line each character might SAY or DO that shows their strength.", y, { fontSize: 10 });

    for (let i = 0; i < 4; i++) {
      doc.fontSize(10).font("Sans-Bold").fillColor(hex(C.ACCENT));
      doc.text(`Character ${i + 1}:`, PAGE.MARGIN, y);
      y += 14;
      y = addLinedArea(doc, y, 1, { lineSpacing: 22 });
      y += 4;
    }

    addPdfFooter(doc, "Respectful Relationships | Grade 6 | Session 3");
    const plannerPath = path.join(RES_DIR, "Session 3 Role Play Planning Card.pdf");
    await writePdf(doc, plannerPath);
    console.log("PDF written:", plannerPath);
  }

  console.log("\nBuild complete!");
  console.log("Output folder:", LESSON_FOLDER);
}

build().catch((err) => { console.error(err); process.exit(1); });
