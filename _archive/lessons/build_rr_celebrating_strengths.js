"use strict";

// Respectful Relationships - Celebrating Strengths (Activity 5)
// Grade 5/6 Wellbeing - 30+ minute session
// Students create an advertisement promoting a character strength

const path = require("path");
const fs = require("fs");
const PptxGenJS = require("pptxgenjs");
const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf,
  addPdfHeader, addSectionHeading, addBodyText, addTipBox,
  addPdfFooter,
  addResourceSlide, makeSessionResource,
  getSessionResourceFolder,
  PAGE, hex,
} = require("../themes/pdf_helpers");

// Theme — same variant as Session 1 (unit cohesion)
const T = createTheme("wellbeing", "grade56", weekToVariant(1));
const {
  C, FONT_H, FONT_B,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  pairShareSlide, scenarioSlide,
  withReveal,
  addTopBar, addBadge, addTitle, addCard, addFooter, addTextOnShape,
  composeNotes,
  CONTENT_TOP, SAFE_BOTTOM,
} = T;

// Output paths
const UNIT = "RR_Celebrating_Strengths";
const LESSON_FOLDER = path.join(__dirname, "..", "output", UNIT);
const PPTX_NAME = "Celebrating Strengths.pptx";
const FOOTER = "Respectful Relationships | Grade 5/6 Wellbeing";
const SESSION = 5;
const RES_DIR = path.join(LESSON_FOLDER, getSessionResourceFolder(SESSION));

// Resources
const CARDS_RESOURCE = makeSessionResource(
  SESSION,
  "Character Strengths Cards",
  "Lucky dip cards - one strength per card. Cut up for students to select."
);

const PLANNING_RESOURCE = makeSessionResource(
  SESSION,
  "Advertisement Planning Frame",
  "Simple planner to sketch key images and messages before creating their poster."
);

const RESOURCE_ITEMS = [CARDS_RESOURCE, PLANNING_RESOURCE];

fs.mkdirSync(RES_DIR, { recursive: true });

// ============================================================
// Character Strengths (same set as Activity 1 handout)
// ============================================================

const STRENGTHS = [
  { name: "Kindness", desc: "Caring about others and helping when it matters." },
  { name: "Courage", desc: "Doing the right thing even when it's scary or hard." },
  { name: "Honesty", desc: "Telling the truth and being real about who you are." },
  { name: "Fairness", desc: "Treating everyone evenly and giving people a fair go." },
  { name: "Humour", desc: "Bringing lightness, smiles, and laughter to others." },
  { name: "Persistence", desc: "Keeping going when something is difficult." },
  { name: "Gratitude", desc: "Noticing good things and saying thank you." },
  { name: "Curiosity", desc: "Wanting to learn, explore, and ask questions." },
  { name: "Teamwork", desc: "Working well with others to get things done." },
  { name: "Leadership", desc: "Helping a group move in a good direction." },
  { name: "Forgiveness", desc: "Letting go of hurt instead of holding a grudge." },
  { name: "Self-Control", desc: "Managing your feelings and actions carefully." },
  { name: "Creativity", desc: "Coming up with new ideas or different ways of doing things." },
  { name: "Love of Learning", desc: "Enjoying finding out new things." },
  { name: "Hope", desc: "Expecting good things and working toward them." },
  { name: "Caring", desc: "Looking out for how others feel and what they need." },
];

// ============================================================
// Teacher Notes
// ============================================================

const NOTES_TITLE = composeNotes({
  say: [
    "Today we're going to celebrate character strengths by creating an advertisement that promotes a strength and shows why it matters.",
    "Last session you acted out strengths in role plays. Today you're making a poster that sells a strength to the world.",
  ],
  do: [
    "Display slide as students settle.",
    "Have Character Strengths Cards cut up and placed in a container for the lucky dip.",
    "Have the Advertisement Planning Frame ready to distribute.",
  ],
  teacherNotes: "This is Activity 5 in the Respectful Relationships unit on character strengths. Students have already defined character strengths (Activity 1), explored them through scenarios (Activities 2-3), and acted them in role plays (Activity 4). Today they create an advertisement promoting a single strength.",
  watchFor: [
    "Students recalling their role play strength from last session - use this to build the bridge.",
  ],
  tag: "[General: Title | VTLM 2.0: Establishing Purpose]",
});

const NOTES_RESOURCES = composeNotes({
  say: [
    "These are the materials for this session.",
  ],
  do: [
    "Print and cut up the Character Strengths Cards before the session.",
    "Print one Advertisement Planning Frame per student (or per pair if working in pairs).",
    "Have A3 paper, coloured pencils, and markers available for poster creation.",
  ],
  teacherNotes: "The Character Strengths Cards are the same 16 strengths from Activity 1. Students select one from the lucky dip. The Planning Frame gives structure before they create their final poster on A3 paper.",
  watchFor: [
    "Ensure cards are cut up and in a container ready for the lucky dip.",
  ],
  tag: "[General: Resources | VTLM 2.0: Preparation]",
});

const NOTES_HOOK = composeNotes({
  say: [
    "Last session you acted out character strengths in role plays. Turn to your partner: which strength did your group show? How did the audience know?",
    "Cold call 2-3 students: What strength did you see someone else's group act out? How did you spot it?",
  ],
  do: [
    "Allow 30 seconds for partner chat.",
    "Cold call and capture 3-4 strengths on the board (e.g. tolerance, fairness, teamwork).",
    "Bridge: Today we're creating something that PROMOTES a strength - like an advertisement convincing people why it matters.",
  ],
  teacherNotes: "This hook connects to Activity 4 (role plays). Keep it brisk - 2-3 minutes maximum. The board list provides a visual anchor for the new task.",
  watchFor: [
    "Students who were absent for the role play - invite them to share a strength they admire in someone instead.",
  ],
  tag: "[General: Hook | VTLM 2.0: Prior Knowledge]",
});

const NOTES_LI = composeNotes({
  say: [
    "Read the Learning Intention aloud.",
    "Read each Success Criterion - we'll come back to these at the end.",
  ],
  do: [
    "Point to each SC as you read it.",
    "Leave the slide visible for 10-15 seconds.",
  ],
  teacherNotes: "SC1 is the floor (everyone can explain what their strength means). SC2 is the core (connecting the strength to real life). SC3 extends into creative communication through the advertisement.",
  watchFor: [
    "Students who look uncertain about 'contribute to a good life' - this will be modelled on the next slide.",
  ],
  tag: "[General: LI/SC | VTLM 2.0: Learning Intentions]",
});

const NOTES_IDO = composeNotes({
  say: [
    "Watch me first. I've pulled a card from the lucky dip and it says... PERSISTENCE.",
    "Step 1: What does persistence mean? Keeping going when things are hard.",
    "Step 2: How does persistence contribute to a good life? If you give up every time something is tough, you never grow. Persistence helps you finish hard tasks, reach goals, and feel proud of yourself.",
    "Step 3: What makes a good advertisement? Strong images, a catchy slogan, clear reasons why someone should want this.",
  ],
  do: [
    "Model your thinking out loud - show how you brainstorm the connection between the strength and living well.",
    "Sketch quick advertisement ideas on the board: a bold slogan, one key image idea, 2-3 reasons why this strength matters.",
    "Keep the model deliberately quick and sketchy - show that the thinking matters more than the art.",
  ],
  teacherNotes: "The coaching point from the source material: creating a sense of relationship between character strengths and living well is important as it encourages children to value positive characteristics. Model this thinking explicitly.",
  watchFor: [
    "Students who conflate the strength with a talent - redirect gently if needed.",
    "Students already generating ideas for their own card - let them hold onto those ideas.",
  ],
  tag: "[General: I Do | VTLM 2.0: Explicit Modelling]",
});

const NOTES_IDO_AD_FEATURES = composeNotes({
  say: [
    "A strong advertisement has these parts: a bold headline or slogan, a key image that grabs attention, clear reasons or messages about why this matters, and it leaves people wanting to have this thing.",
    "Think about ads you see every day - they don't have 50 words on them. They have one big idea and powerful images.",
  ],
  do: [
    "Point to each feature on the slide as you name it.",
    "Refer back to your board sketch from the previous slide as an example.",
  ],
  teacherNotes: "This slide provides the success criteria for a strong advertisement without being prescriptive about layout. Students need the 'what makes it good' before they design their own.",
  watchFor: [
    "Students who may over-focus on artistic quality rather than the MESSAGE - emphasise that the thinking matters most.",
  ],
  tag: "[General: I Do | VTLM 2.0: Explicit Explanation]",
});

const NOTES_CFU = composeNotes({
  say: [
    "Quick check before we start creating.",
  ],
  do: [
    "Read the question aloud.",
    "Students show response on mini-whiteboards or share with partner.",
    "Scan for students who understand the concept of 'contribute to a good life'.",
  ],
  cfu: {
    technique: "Mini-whiteboard or Turn and Tell",
    script: [
      "On your whiteboard (or tell your partner): If your strength is COURAGE, write one way courage could help someone live a good life.",
      "Look for answers like: standing up for a friend, trying new things, being honest even when it's hard.",
    ],
    scanFor: "Students connecting the strength to a real-life benefit or situation, not just defining the word.",
    proceed: "If most students connect the strength to a life benefit, move to the lucky dip.",
    pivot: "If students only repeat the definition, model the thinking again with a different strength: 'KINDNESS means caring about others - but HOW does it help someone live well? It means you have good friendships, people trust you, you feel good about yourself.' Re-ask with FAIRNESS.",
  },
  teacherNotes: "This CFU checks the core thinking move needed for the advertisement task. Students who can make this connection are ready to work independently.",
  watchFor: [
    "Students who write only 'it helps people' without specifics - push for a concrete example.",
  ],
  tag: "[General: CFU | VTLM 2.0: Checking for Understanding]",
});

const NOTES_RETEACH = composeNotes({
  say: [
    "Let me show you another way to think about this.",
    "Imagine someone with the strength of TEAMWORK. Picture them at school, at home, at sport. What's their life like because of teamwork?",
    "They have people who enjoy working with them. They get more done. They solve problems together. Their life is better because they know how to work with others.",
  ],
  do: [
    "Draw a simple mind map on the board: TEAMWORK in the centre, then arrows to 'good friendships', 'solves problems', 'achieves more'.",
    "Invite one student to add a branch for a different strength.",
  ],
  teacherNotes: "Skip this slide if the CFU showed 80%+ understanding. This re-teach uses a mind-map visual instead of the linear thinking from the I Do. The different entry point helps students who need to see the connections mapped spatially.",
  watchFor: [
    "Students who now make a connection - they're ready to proceed.",
  ],
  tag: "[General: Optional Re-teach | VTLM 2.0: Differentiated Instruction]",
});

const NOTES_LUCKY_DIP = composeNotes({
  say: [
    "Time for the lucky dip. You'll reach into the container and pull out one character strength card.",
    "This is YOUR strength to promote. You're going to make an advertisement that shows why this strength is amazing and what it can provide in someone's life.",
    "You can work individually or with a partner who drew the same or a similar strength.",
  ],
  do: [
    "Circulate with the container for the lucky dip.",
    "Distribute the Advertisement Planning Frame.",
    "Allow students who drew the same strength to pair up if they wish.",
    "Give students 2-3 minutes on the Planning Frame before they move to A3 paper.",
  ],
  teacherNotes: "The lucky dip adds engagement and removes choice paralysis. Some students may want to swap - allow one swap if needed, but encourage them to work with what they get. The planning frame structures their thinking before they commit to the poster.",
  watchFor: [
    "Students who look disappointed with their strength - prompt them: 'This is your challenge - make me believe this is the best strength in the world.'",
    "Students who skip the planning frame and jump to drawing - redirect them to plan first.",
  ],
  tag: "[General: You Do | VTLM 2.0: Independent Practice]",
});

const NOTES_CREATE = composeNotes({
  say: [
    "You have about 15 minutes to design and create your advertisement poster.",
    "Remember: bold headline, strong images, clear messages about WHY this strength matters for a good life.",
    "Think about your audience - make them want to HAVE this strength.",
  ],
  do: [
    "Distribute A3 paper, coloured pencils, and markers.",
    "Circulate and conference with students/pairs.",
    "Conference questions: 'Tell me why this strength matters.' 'What's your slogan?' 'What will your audience think when they see this?'",
    "Give a 5-minute warning before the gallery walk.",
  ],
  enabling: {
    text: "ENABLING PROMPT:",
    bullets: [
      "Task: Students who struggle to connect the strength to life benefits use the sentence stem on their planning frame: '[Strength] helps you ___ because ___.'",
      "Extra Notes: Pair with a confident partner. Offer to brainstorm 3 life benefits together before they start the poster.",
    ],
  },
  teacherNotes: "This is the main creative work time. Keep energy up but allow focus. Avoid over-directing the artistic choices - the thinking behind the messages is what matters.",
  watchFor: [
    "Students spending all their time on borders/decoration rather than content - redirect to the messages.",
    "Students who finish quickly - challenge them to add a 'testimonial' quote from an imaginary person whose life is better because of this strength.",
  ],
  tag: "[General: You Do | VTLM 2.0: Independent Practice]",
});

const NOTES_GALLERY = composeNotes({
  say: [
    "Time for our gallery walk. Display your poster on your desk or the wall.",
    "As you walk around, look at each poster and think: What's the strongest message? What convinced you?",
    "When we come back together, each person or pair will introduce their poster in 30 seconds or less.",
  ],
  do: [
    "Arrange a gallery walk path around the room.",
    "Allow 3-4 minutes for viewing.",
    "Invite each student/pair to present their poster briefly (30 seconds each).",
    "After each presentation, invite one observation from the class: 'What message stood out to you?'",
  ],
  teacherNotes: "The gallery walk and presentations provide accountability and celebration. Keep presentations brief - 30 seconds each. Display the posters on the class wall afterwards as a reference for the unit.",
  watchFor: [
    "Students who are reluctant to present - offer the option to have a partner introduce their poster.",
    "Positive peer feedback - encourage specific observations rather than generic 'that's good'.",
  ],
  tag: "[General: Share | VTLM 2.0: Collaborative Learning]",
});

const NOTES_CLOSING = composeNotes({
  say: [
    "Let's look back at our success criteria. Thumbs up, sideways, or down for each.",
    "SC1: I can explain what my character strength means.",
    "SC2: I can describe how a character strength helps someone live a good life.",
    "SC3: I can create an advertisement that promotes the advantages of a strength.",
    "Turn and Talk: Has this activity helped you see how a character strength can contribute to a good life? Tell your partner one thing you now understand better.",
  ],
  do: [
    "Read each SC aloud; pause for thumbs.",
    "Allow 40 seconds for Turn and Talk.",
    "Cold call 2 students: 'What do you now understand better about character strengths and living well?'",
  ],
  teacherNotes: "The reflection directly addresses the coaching point: creating a sense of relationship between character strengths and living well encourages children to value positive characteristics in themselves and others.",
  watchFor: [
    "Students showing sideways/down on SC2 - note for revisiting the strengths-to-life connection in later sessions.",
    "Celebrate specific poster messages that showed deep thinking about how strengths contribute to life.",
  ],
  tag: "[General: Closing | VTLM 2.0: Review and Reflect]",
});

// ============================================================
// Build function
// ============================================================

async function build() {
  const pres = new PptxGenJS();
  pres.layout = "LAYOUT_16x9";

  // -- Slide 1: Title --
  titleSlide(
    pres,
    "Celebrating\nStrengths",
    "Respectful Relationships",
    "Grade 5/6 Wellbeing  |  Session 5  |  30+ minutes",
    NOTES_TITLE
  );

  // -- Slide 2: Teacher Resources --
  addResourceSlide(pres, {
    resources: RESOURCE_ITEMS,
    manipulatives: [
      "Character Strengths Cards in a container (lucky dip)",
    ],
    studentTools: [
      "A3 paper for poster creation",
      "Coloured pencils, markers, textas",
      "Advertisement Planning Frame (one per student or pair)",
    ],
    boardSetup: [
      "Space to display student posters for gallery walk",
      "Teacher board for modelling advertisement brainstorm",
    ],
  }, T, FOOTER, NOTES_RESOURCES);

  // -- Slide 3: Hook --
  pairShareSlide(
    pres,
    "Connect to Last Session",
    [
      "Which strength did your group show in the role play?",
      "How did the audience know which strength it was?",
      "What made the strength useful in your scene?",
    ],
    NOTES_HOOK,
    FOOTER
  );

  // -- Slide 4: LI / SC --
  liSlide(
    pres,
    ["We are learning to identify how a particular character strength can contribute to a good life"],
    [
      "I can explain what my character strength means",
      "I can describe how a character strength helps someone live a good life",
      "I can create an advertisement that promotes the advantages of a strength",
    ],
    NOTES_LI,
    FOOTER
  );

  // -- Slide 5: I Do - Modelling the thinking --
  contentSlide(
    pres,
    "I Do",
    C.PRIMARY,
    "How Does a Strength Help You Live Well?",
    [
      "Step 1: What does the strength mean?",
      "Step 2: How does it help in real life?",
      "Step 3: Why would someone WANT this strength?",
    ],
    NOTES_IDO,
    FOOTER,
    (s, lg) => {
      // Right column: Teacher's worked example for PERSISTENCE
      const rX = lg.rightX;
      const rW = lg.rightW;
      const topY = lg.panelTopPadded;
      const cardH = 3.4;

      addCard(s, rX, topY, rW, cardH, { strip: C.ACCENT, fill: C.WHITE });

      s.addText("Teacher Example: PERSISTENCE", {
        x: rX + 0.12, y: topY + 0.1, w: rW - 0.24, h: 0.35,
        fontSize: 12, fontFace: FONT_H, color: C.ACCENT, bold: true, align: "center", margin: 0,
      });

      // Meaning
      s.addShape("roundRect", {
        x: rX + 0.15, y: topY + 0.55, w: rW - 0.3, h: 0.55,
        rectRadius: 0.06,
        fill: { color: C.BG_LIGHT },
        line: { color: C.MUTED, width: 0.5 },
      });
      s.addText("Keeping going when things are hard", {
        x: rX + 0.25, y: topY + 0.58, w: rW - 0.5, h: 0.5,
        fontSize: 11, fontFace: FONT_B, color: C.CHARCOAL, valign: "middle", margin: 0,
      });

      // How it helps
      s.addShape("roundRect", {
        x: rX + 0.15, y: topY + 1.2, w: rW - 0.3, h: 1.05,
        rectRadius: 0.06,
        fill: { color: C.BG_LIGHT },
        line: { color: C.MUTED, width: 0.5 },
      });
      s.addText([
        { text: "You finish hard tasks", options: { bullet: true, breakLine: true, fontSize: 11 } },
        { text: "You reach your goals", options: { bullet: true, breakLine: true, fontSize: 11 } },
        { text: "You feel proud of yourself", options: { bullet: true, breakLine: true, fontSize: 11 } },
        { text: "You grow stronger each time", options: { bullet: true, breakLine: false, fontSize: 11 } },
      ], {
        x: rX + 0.25, y: topY + 1.25, w: rW - 0.5, h: 0.95,
        fontFace: FONT_B, color: C.CHARCOAL, valign: "top", margin: 0,
      });

      // Slogan idea
      s.addShape("roundRect", {
        x: rX + 0.15, y: topY + 2.4, w: rW - 0.3, h: 0.75,
        rectRadius: 0.06,
        fill: { color: C.ACCENT },
        line: { color: C.ACCENT, width: 0.5 },
      });
      s.addText("Slogan idea:\n\"Never give up. Never give in.\nPersistence wins.\"", {
        x: rX + 0.25, y: topY + 2.45, w: rW - 0.5, h: 0.65,
        fontSize: 12, fontFace: FONT_H, color: C.WHITE, bold: true, align: "center", valign: "middle", margin: 0,
      });
    }
  );

  // -- Slide 6: I Do - What makes a good advertisement --
  contentSlide(
    pres,
    "I Do",
    C.PRIMARY,
    "What Makes a Strong Advertisement?",
    [
      "A bold headline or slogan",
      "A key image that grabs attention",
      "Clear reasons WHY this strength matters",
      "Leaves people thinking: I want that!",
    ],
    NOTES_IDO_AD_FEATURES,
    FOOTER,
    (s, lg) => {
      // Right column: simple visual of a poster layout
      const rX = lg.rightX;
      const rW = lg.rightW;
      const topY = lg.panelTopPadded;
      const cardH = 3.4;

      // Poster mockup frame
      addCard(s, rX, topY, rW, cardH, { strip: C.SECONDARY, fill: C.WHITE });
      s.addText("Poster Layout", {
        x: rX + 0.1, y: topY + 0.08, w: rW - 0.2, h: 0.28,
        fontSize: 10, fontFace: FONT_B, color: C.MUTED, italic: true, align: "center", margin: 0,
      });

      // Headline area
      s.addShape("roundRect", {
        x: rX + 0.2, y: topY + 0.4, w: rW - 0.4, h: 0.55,
        rectRadius: 0.06,
        fill: { color: C.PRIMARY },
        line: { color: C.PRIMARY, width: 0.5 },
      });
      s.addText("BOLD HEADLINE", {
        x: rX + 0.3, y: topY + 0.45, w: rW - 0.6, h: 0.45,
        fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true, align: "center", valign: "middle", margin: 0,
      });

      // Image area
      s.addShape("roundRect", {
        x: rX + 0.2, y: topY + 1.05, w: rW - 0.4, h: 1.2,
        rectRadius: 0.06,
        fill: { color: C.BG_LIGHT },
        line: { color: C.MUTED, width: 0.8, dashType: "dash" },
      });
      s.addText("KEY IMAGE", {
        x: rX + 0.3, y: topY + 1.45, w: rW - 0.6, h: 0.4,
        fontSize: 13, fontFace: FONT_B, color: C.MUTED, align: "center", valign: "middle", margin: 0,
      });

      // Messages area
      s.addShape("roundRect", {
        x: rX + 0.2, y: topY + 2.35, w: rW - 0.4, h: 0.85,
        rectRadius: 0.06,
        fill: { color: C.BG_LIGHT },
        line: { color: C.ACCENT, width: 0.8 },
      });
      s.addText("WHY THIS\nSTRENGTH MATTERS", {
        x: rX + 0.3, y: topY + 2.45, w: rW - 0.6, h: 0.7,
        fontSize: 12, fontFace: FONT_B, color: C.ACCENT, bold: true, align: "center", valign: "middle", margin: 0,
      });
    }
  );

  // -- Slide 7: CFU --
  cfuSlide(
    pres,
    "Check",
    "Can You Make the Connection?",
    {
      technique: "Mini-whiteboard or Turn and Tell",
      question: "If your strength is COURAGE, write one way courage could help someone live a good life.",
    },
    NOTES_CFU,
    FOOTER
  );

  // -- Slide 8: Optional Re-teach --
  contentSlide(
    pres,
    "Optional Re-teach",
    C.SECONDARY,
    "Another Way to See It",
    [
      "Pick a strength. Put it in the middle.",
      "Ask: What does life look like WITH this strength?",
      "Draw arrows to real-life benefits.",
      "Example: TEAMWORK -> good friendships, solves problems, achieves more",
    ],
    NOTES_RETEACH,
    FOOTER,
    (s, lg) => {
      // Right column: mind map visual
      const rX = lg.rightX;
      const rW = lg.rightW;
      const topY = lg.panelTopPadded;
      const cardH = 3.2;

      addCard(s, rX, topY, rW, cardH, { strip: C.SUCCESS, fill: C.WHITE });

      // Centre circle
      const cx = rX + rW / 2;
      const cy = topY + cardH / 2;
      const circR = 0.55;
      s.addShape("roundRect", {
        x: cx - circR, y: cy - circR, w: circR * 2, h: circR * 2,
        rectRadius: circR,
        fill: { color: C.PRIMARY },
        line: { color: C.PRIMARY, width: 1 },
      });
      s.addText("TEAMWORK", {
        x: cx - circR, y: cy - circR, w: circR * 2, h: circR * 2,
        fontSize: 9, fontFace: FONT_H, color: C.WHITE, bold: true, align: "center", valign: "middle", margin: 0,
      });

      // Branches
      const branches = [
        { label: "Good\nfriendships", dx: -1.1, dy: -0.85 },
        { label: "Solves\nproblems", dx: 1.2, dy: -0.35 },
        { label: "Achieves\nmore", dx: 0.4, dy: 1.05 },
      ];

      branches.forEach((b) => {
        const bx = cx + b.dx;
        const by = cy + b.dy;

        // Line from center to branch
        s.addShape("line", {
          x: cx, y: cy,
          w: b.dx, h: b.dy,
          line: { color: C.ACCENT, width: 1.5 },
        });

        // Branch bubble
        s.addShape("roundRect", {
          x: bx - 0.5, y: by - 0.25, w: 1.0, h: 0.5,
          rectRadius: 0.1,
          fill: { color: C.BG_LIGHT },
          line: { color: C.ACCENT, width: 0.8 },
        });
        s.addText(b.label, {
          x: bx - 0.5, y: by - 0.25, w: 1.0, h: 0.5,
          fontSize: 9, fontFace: FONT_B, color: C.CHARCOAL, align: "center", valign: "middle", margin: 0,
        });
      });
    }
  );

  // -- Slide 9: Lucky Dip + Task Setup (You Do) --
  contentSlide(
    pres,
    "You Do",
    C.ACCENT,
    "Your Artistic Commission",
    [
      "First: Pick a strength from the lucky dip",
      "Next: Use your Planning Frame to brainstorm",
      "Then: Create your advertisement poster",
      "Make the world WANT your strength!",
    ],
    NOTES_LUCKY_DIP,
    FOOTER,
    (s, lg) => {
      // Right column: planning frame preview
      const rX = lg.rightX;
      const rW = lg.rightW;
      const topY = lg.panelTopPadded;
      const cardH = 3.3;

      addCard(s, rX, topY, rW, cardH, { strip: C.ACCENT, fill: C.WHITE });

      s.addText("Planning Frame", {
        x: rX + 0.1, y: topY + 0.08, w: rW - 0.2, h: 0.28,
        fontSize: 11, fontFace: FONT_H, color: C.ACCENT, bold: true, align: "center", margin: 0,
      });

      // Mini version of planning frame sections
      const sections = [
        { label: "My strength:", h: 0.4 },
        { label: "What it means:", h: 0.5 },
        { label: "How it helps in life:", h: 0.7 },
        { label: "My slogan:", h: 0.4 },
        { label: "Key image idea:", h: 0.7 },
      ];

      let secY = topY + 0.4;
      sections.forEach((sec) => {
        s.addShape("roundRect", {
          x: rX + 0.15, y: secY, w: rW - 0.3, h: sec.h,
          rectRadius: 0.04,
          fill: { color: C.BG_LIGHT },
          line: { color: C.MUTED, width: 0.5 },
        });
        s.addText(sec.label, {
          x: rX + 0.22, y: secY + 0.04, w: rW - 0.44, h: 0.22,
          fontSize: 9, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
        });
        secY += sec.h + 0.06;
      });
    }
  );

  // -- Slide 10: Create Time --
  contentSlide(
    pres,
    "You Do",
    C.ACCENT,
    "Design & Create",
    [
      "Bold headline or slogan",
      "Key images that grab attention",
      "Clear messages: WHY does this strength matter?",
      "Make your audience think: I want that!",
    ],
    NOTES_CREATE,
    FOOTER,
    (s, lg) => {
      // Right column: timer/reminder visual
      const rX = lg.rightX;
      const rW = lg.rightW;
      const topY = lg.panelTopPadded;

      // Time reminder
      addTextOnShape(s, "15 minutes", {
        x: rX + 0.3, y: topY + 0.1, w: rW - 0.6, h: 0.55, rectRadius: 0.1,
        fill: { color: C.PRIMARY },
      }, {
        fontSize: 18, fontFace: FONT_H, color: C.WHITE,
        bold: true, align: "center", valign: "middle", margin: 0,
      });

      // Checklist
      const checkY = topY + 0.85;
      const checks = [
        "Headline that grabs attention",
        "Image that shows the strength",
        "Reasons it helps in life",
        "Makes people WANT it",
      ];

      checks.forEach((item, i) => {
        const iy = checkY + i * 0.55;
        // Checkbox
        s.addShape("roundRect", {
          x: rX + 0.2, y: iy + 0.08, w: 0.28, h: 0.28,
          rectRadius: 0.04,
          fill: { color: C.WHITE },
          line: { color: C.ACCENT, width: 1.2 },
        });
        s.addText(item, {
          x: rX + 0.58, y: iy, w: rW - 0.78, h: 0.44,
          fontSize: 11, fontFace: FONT_B, color: C.CHARCOAL, valign: "middle", margin: 0,
        });
      });
    }
  );

  // -- Slide 11: Gallery Walk --
  contentSlide(
    pres,
    "Gallery Walk",
    C.SUCCESS,
    "Share & Celebrate",
    [
      "Display your poster",
      "Walk around and view each advertisement",
      "Think: What message stood out? What convinced you?",
      "Introduce your poster (30 seconds each)",
    ],
    NOTES_GALLERY,
    FOOTER,
    (s, lg) => {
      // Right column: gallery walk visual cue
      const rX = lg.rightX;
      const rW = lg.rightW;
      const topY = lg.panelTopPadded;

      addCard(s, rX, topY, rW, 2.8, { strip: C.SUCCESS, fill: C.WHITE });

      s.addText("As you walk, look for:", {
        x: rX + 0.15, y: topY + 0.15, w: rW - 0.3, h: 0.35,
        fontSize: 12, fontFace: FONT_H, color: C.SUCCESS, bold: true, margin: 0,
      });

      const prompts = [
        "The strongest slogan",
        "The most convincing reason",
        "An image that grabs you",
        "A strength you want to have",
      ];

      prompts.forEach((p, i) => {
        const py = topY + 0.6 + i * 0.5;
        s.addShape("roundRect", {
          x: rX + 0.2, y: py, w: 0.22, h: 0.22,
          rectRadius: 0.11,
          fill: { color: C.ACCENT },
          line: { color: C.ACCENT, width: 0.5 },
        });
        s.addText(p, {
          x: rX + 0.52, y: py - 0.04, w: rW - 0.72, h: 0.3,
          fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, valign: "middle", margin: 0,
        });
      });
    }
  );

  // -- Slide 12: Closing --
  closingSlide(
    pres,
    {
      reflectionPrompt: "Has this activity helped you see how a character strength can contribute to a good life? Tell your partner one thing you understand better now.",
      scItems: [
        "I can explain what my character strength means",
        "I can describe how a character strength helps someone live a good life",
        "I can create an advertisement that promotes the advantages of a strength",
      ],
      selfAssessment: {
        prompt: "Thumbs up, sideways, or down for each success criterion.",
        options: ["Got it", "Getting there", "Need help"],
      },
    },
    NOTES_CLOSING
  );

  // -- Slide 13: Resources --
  addResourceSlide(pres, RESOURCE_ITEMS, T, FOOTER, NOTES_RESOURCES);

  // Write PPTX
  fs.mkdirSync(LESSON_FOLDER, { recursive: true });
  await pres.writeFile({ fileName: path.join(LESSON_FOLDER, PPTX_NAME) });
  console.log("PPTX written to", path.join(LESSON_FOLDER, PPTX_NAME));

  // ============================================================
  // Companion PDF 1: Character Strengths Cards (Lucky Dip)
  // ============================================================
  {
    const doc = createPdf({ title: "Character Strengths Cards" });
    let y = addPdfHeader(doc, "Session 5 Character Strengths Cards", {
      subtitle: "Cut into individual cards for the lucky dip",
      color: C.PRIMARY,
      lessonInfo: "Grade 5/6 Wellbeing - Respectful Relationships",
    });

    y += 10;

    // 2 columns x 8 rows of cards with cutting guides
    const colCount = 2;
    const colGap = 16;
    const colW = (PAGE.CONTENT_W - colGap * (colCount - 1)) / colCount;
    const cardH = 78;
    const cardGap = 8;

    let cardY = y;
    STRENGTHS.forEach((st, i) => {
      const col = i % colCount;
      if (col === 0 && i > 0) {
        cardY += cardH + cardGap;
      }

      // Page break if the next row won't fit
      if (col === 0 && cardY + cardH > PAGE.H - PAGE.MARGIN) {
        doc.addPage();
        cardY = PAGE.MARGIN;
      }

      const cx = PAGE.MARGIN + col * (colW + colGap);

      doc.save();
      doc.roundedRect(cx, cardY, colW, cardH, 6)
        .lineWidth(1.5)
        .dash(4, { space: 3 })
        .stroke(hex(C.MUTED));
      doc.undash();
      doc.restore();

      // Strength name
      doc.fontSize(16).font("Sans-Bold").fillColor(hex(C.PRIMARY));
      doc.text(st.name, cx + 10, cardY + 12, { width: colW - 20, align: "center" });

      // Description
      doc.fontSize(10).font("Sans").fillColor(hex("4B5563"));
      doc.text(st.desc, cx + 10, cardY + 38, { width: colW - 20, align: "center" });

      // Small scissors icon cue
      doc.fontSize(8).font("Sans").fillColor(hex(C.MUTED));
      doc.text("cut", cx + colW - 28, cardY + cardH - 14, { width: 20, align: "right" });
    });

    addPdfFooter(doc);
    const cardsPath = path.join(RES_DIR, `${CARDS_RESOURCE.name}.pdf`);
    await writePdf(doc, cardsPath);
    console.log("PDF written to", cardsPath);
  }

  // ============================================================
  // Companion PDF 2: Advertisement Planning Frame
  // ============================================================
  {
    const doc = createPdf({ title: "Advertisement Planning Frame" });
    let y = addPdfHeader(doc, "Session 5 Advertisement Planning Frame", {
      subtitle: "Plan your advertisement before creating your poster",
      color: C.ACCENT,
      lessonInfo: "Grade 5/6 Wellbeing - Respectful Relationships",
    });

    y += 8;

    // Section 1: My Strength
    doc.save();
    doc.roundedRect(PAGE.MARGIN, y, PAGE.CONTENT_W, 50, 4)
      .lineWidth(1)
      .stroke(hex(C.PRIMARY));
    doc.restore();
    doc.fontSize(11).font("Sans-Bold").fillColor(hex(C.PRIMARY));
    doc.text("My character strength:", PAGE.MARGIN + 10, y + 8, { width: PAGE.CONTENT_W - 20 });
    doc.fontSize(10).font("Sans").fillColor(hex("6B7280"));
    doc.text("(Write the strength from your lucky dip card)", PAGE.MARGIN + 10, y + 28, { width: PAGE.CONTENT_W - 20 });
    y += 60;

    // Section 2: What it means
    doc.save();
    doc.roundedRect(PAGE.MARGIN, y, PAGE.CONTENT_W, 70, 4)
      .lineWidth(1)
      .stroke(hex(C.PRIMARY));
    doc.restore();
    doc.fontSize(11).font("Sans-Bold").fillColor(hex(C.PRIMARY));
    doc.text("What does this strength mean?", PAGE.MARGIN + 10, y + 8, { width: PAGE.CONTENT_W - 20 });
    // Lines for writing
    for (let i = 0; i < 2; i++) {
      const lineY = y + 32 + i * 18;
      doc.moveTo(PAGE.MARGIN + 10, lineY)
        .lineTo(PAGE.MARGIN + PAGE.CONTENT_W - 10, lineY)
        .lineWidth(0.5)
        .stroke(hex(C.MUTED));
    }
    y += 80;

    // Section 3: How it helps someone live a good life
    doc.save();
    doc.roundedRect(PAGE.MARGIN, y, PAGE.CONTENT_W, 110, 4)
      .lineWidth(1)
      .stroke(hex(C.ACCENT));
    doc.restore();
    doc.fontSize(11).font("Sans-Bold").fillColor(hex(C.ACCENT));
    doc.text("How does this strength help someone live a good life?", PAGE.MARGIN + 10, y + 8, { width: PAGE.CONTENT_W - 20 });
    doc.fontSize(9).font("Sans-Italic").fillColor(hex("6B7280"));
    doc.text("Think: What can someone DO because they have this strength? How are their relationships, goals, and daily life better?", PAGE.MARGIN + 10, y + 24, { width: PAGE.CONTENT_W - 20 });
    // Lines
    for (let i = 0; i < 4; i++) {
      const lineY = y + 48 + i * 16;
      doc.moveTo(PAGE.MARGIN + 10, lineY)
        .lineTo(PAGE.MARGIN + PAGE.CONTENT_W - 10, lineY)
        .lineWidth(0.5)
        .stroke(hex(C.MUTED));
    }
    y += 120;

    // Section 4: My slogan or headline
    doc.save();
    doc.roundedRect(PAGE.MARGIN, y, PAGE.CONTENT_W, 65, 4)
      .lineWidth(1.5)
      .stroke(hex(C.PRIMARY));
    doc.restore();
    doc.fontSize(11).font("Sans-Bold").fillColor(hex(C.PRIMARY));
    doc.text("My slogan or headline:", PAGE.MARGIN + 10, y + 8, { width: PAGE.CONTENT_W - 20 });
    doc.fontSize(9).font("Sans-Italic").fillColor(hex("6B7280"));
    doc.text("Something short, bold, and catchy that grabs attention.", PAGE.MARGIN + 10, y + 24, { width: PAGE.CONTENT_W - 20 });
    // Large writing line
    const sloganLineY = y + 46;
    doc.moveTo(PAGE.MARGIN + 10, sloganLineY)
      .lineTo(PAGE.MARGIN + PAGE.CONTENT_W - 10, sloganLineY)
      .lineWidth(0.8)
      .stroke(hex(C.MUTED));
    y += 76;

    // Section 5: Key image sketch
    doc.save();
    doc.roundedRect(PAGE.MARGIN, y, PAGE.CONTENT_W, 180, 4)
      .lineWidth(1)
      .stroke(hex(C.ACCENT));
    doc.restore();
    doc.fontSize(11).font("Sans-Bold").fillColor(hex(C.ACCENT));
    doc.text("Sketch your key image idea here:", PAGE.MARGIN + 10, y + 8, { width: PAGE.CONTENT_W - 20 });
    doc.fontSize(9).font("Sans-Italic").fillColor(hex("6B7280"));
    doc.text("What will people SEE on your poster? Quick sketch - it doesn't need to be perfect.", PAGE.MARGIN + 10, y + 24, { width: PAGE.CONTENT_W - 20 });
    y += 190;

    // Section 6: Key messages
    doc.save();
    doc.roundedRect(PAGE.MARGIN, y, PAGE.CONTENT_W, 90, 4)
      .lineWidth(1)
      .stroke(hex(C.PRIMARY));
    doc.restore();
    doc.fontSize(11).font("Sans-Bold").fillColor(hex(C.PRIMARY));
    doc.text("Key messages (2-3 reasons why this strength matters):", PAGE.MARGIN + 10, y + 8, { width: PAGE.CONTENT_W - 20 });

    // Sentence stem help
    doc.fontSize(9).font("Sans-Italic").fillColor(hex("6B7280"));
    doc.text("[Strength] helps you ___ because ___.", PAGE.MARGIN + 10, y + 24, { width: PAGE.CONTENT_W - 20 });
    for (let i = 0; i < 3; i++) {
      const lineY = y + 44 + i * 16;
      doc.moveTo(PAGE.MARGIN + 10, lineY)
        .lineTo(PAGE.MARGIN + PAGE.CONTENT_W - 10, lineY)
        .lineWidth(0.5)
        .stroke(hex(C.MUTED));
    }

    addPdfFooter(doc);
    const planPath = path.join(RES_DIR, `${PLANNING_RESOURCE.name}.pdf`);
    await writePdf(doc, planPath);
    console.log("PDF written to", planPath);
  }

  console.log("\nBuild complete. Output folder:", LESSON_FOLDER);
}

build().catch((err) => { console.error(err); process.exit(1); });
