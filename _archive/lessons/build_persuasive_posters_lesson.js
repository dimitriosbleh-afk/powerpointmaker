"use strict";

// Persuasive Posters - Year 4 Literacy (single session, 60 min)
// Focus: how persuasive posters use headings, images and colour to convince the reader.
// Mixed readiness. Mini-whiteboards available. No source poster supplied -
// the deck and the printed sheet use built wireframe-style poster mockups only.
// Standalone lesson (no unit cohesion constraint). Variant 1 (Plum & Honey).

const pptxgen = require("pptxgenjs");
const path = require("path");

const { createTheme } = require("../themes/factory");
const T = createTheme("literacy", "grade34", 1);
const {
  C, FONT_H, FONT_B,
  withReveal, addRevealAnswerBar,
  titleSlide, liSlide, vocabSlide, annotatedModelSlide, compareVisualSlide,
  cfuSlide, exitTicketSlide, closingSlide,
} = T;

const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addLinedArea,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
  addPosterMockupPdf, PAGE, hex,
} = require("../themes/pdf_helpers");

// ---------------------------------------------------------------------------
// Wireframe poster mockups (structured previewSpec objects).
// The SAME spec object renders on the slides (annotatedModelSlide /
// compareVisualSlide previewSpec) AND on the printed sheet (addPosterMockupPdf),
// so students analyse the identical wireframe on screen and on paper.
// Poster accent colours are chosen deliberately and independently of the theme
// so that COLOUR is a real variable students can compare across posters.
// No real poster is imitated - these are clean structure/feature wireframes.
// ---------------------------------------------------------------------------

const SAVE_BLUE = "1E6F9F";  // Save Water - cool blue
const FAIR_TEAL = "1F8A70";  // School Fair - fresh teal
const READ_ORANGE = "C25A1E"; // Readathon (strong) - warm orange
const READ_GREY = "8A8F98";  // Readathon (weak) - dull grey
const WALK_GREEN = "2E7D32"; // Walk to School - healthy green

// School Fair - the launch poster (students discover its purpose).
const POSTER_FAIR = {
  accent: FAIR_TEAL,
  components: [
    { kind: "masthead", text: "SCHOOL FAIR SATURDAY", scale: 1.1 },
    { kind: "hero", mode: "photo", scale: 2.0 },
    { kind: "cta", text: "Come and join the fun", scale: 0.9 },
  ],
};

// Save Water - the I Do poster (teacher labels the 3 tools). Reused as the
// worked example on the printed sheet so the resource matches the slides.
const POSTER_SAVE = {
  accent: SAVE_BLUE,
  components: [
    { kind: "masthead", text: "SAVE EVERY DROP", scale: 1.1 },
    { kind: "hero", mode: "photo", scale: 2.0 },
    { kind: "cta", text: "Turn off the tap while you brush", scale: 0.9 },
  ],
};

// Readathon, strong design - big heading, book image, warm inviting colour.
const POSTER_READ_STRONG = {
  accent: READ_ORANGE,
  components: [
    { kind: "masthead", text: "GET LOST IN A BOOK", scale: 1.15 },
    { kind: "hero", mode: "photo", scale: 1.9 },
    { kind: "cta", text: "Join the Readathon on Friday", scale: 0.85 },
  ],
};

// Readathon, weak design - tiny heading, no image, wall of grey text.
const POSTER_READ_WEAK = {
  accent: READ_GREY,
  pageFill: "F2F3F5",
  components: [
    { kind: "heading", text: "readathon information", scale: 0.8 },
    { kind: "textBlock", count: 5, scale: 2.6 },
    { kind: "caption", text: "see the office for more details", scale: 0.7 },
  ],
};

// Walk to School - the You Do poster (new content students analyse).
const POSTER_WALK = {
  accent: WALK_GREEN,
  components: [
    { kind: "masthead", text: "WALK TO SCHOOL WEEK", scale: 1.1 },
    { kind: "hero", mode: "photo", scale: 2.0 },
    { kind: "cta", text: "Leave the car at home", scale: 0.9 },
  ],
};

// ---------------------------------------------------------------------------
// Resources
// ---------------------------------------------------------------------------

const SESSION_NUMBER = 1;
const FOOTER = "Persuasive Posters | Year 4 Literacy";
const OUT_DIR = "output/Persuasive_Posters_Lesson";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const DETECTIVE_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Poster Detective",
  "Analyse a persuasive poster: worked example, find how the heading, image and colour persuade, plus a design challenge."
);
const ANSWER_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Answer Key",
  "Sample answers and teacher look-fors for the Poster Detective sheet."
);

const DETECTIVE_PDF_PATH = path.join(OUT_DIR, DETECTIVE_RESOURCE.fileName);
const ANSWER_PDF_PATH = path.join(OUT_DIR, ANSWER_RESOURCE.fileName);

const RESOURCE_CONFIG = {
  resources: [DETECTIVE_RESOURCE, ANSWER_RESOURCE],
  studentTools: ["Mini-whiteboards and markers (one per student)"],
  boardSetup: ["Display deck on the board", "Have the Poster Detective sheet printed, one per student"],
  manipulatives: ["No source poster needed - the deck uses built wireframe examples"],
};

// ---------------------------------------------------------------------------
// Teacher notes (Glance Format, megaprompt v11.0 sections 45-47)
// ---------------------------------------------------------------------------

const NOTES_TITLE =
  "Title slide. Display as students arrive. Today: how persuasive posters use headings, images and colour to convince the reader.";

const NOTES_RESOURCES =
  "Print the Poster Detective sheet, one per student. Keep the Answer Key for yourself. " +
  "Whiteboards and markers out for every student. No source poster needed - the deck uses built wireframe examples.";

const NOTES_LAUNCH = T.composeGlanceNotes({
  answer: "open - listen for what the poster wants you to DO (come to the fair) and what made it stand out (big words, the picture, the bright colour).",
  beats: [
    "SHOW the poster. SAY: Posters are everywhere, and every one wants you to DO something.",
    "ASK: What does this poster want you to do? 15 sec, turn and tell. EXPECT: come to the school fair. ACCEPT: go to the fair on Saturday.",
    "ASK: What made you notice it first - the big words, the picture, or the colour? Boards up, 15 sec. EXPECT: any of the three, with a reason.",
    "SCAN boards. 80%+ naming a feature -> next slide. Less -> point to the big heading and the bright colour, re-ask.",
  ],
  trap: "saying only what the poster says, not what it wants you to do. Fix: ask 'what does it want YOU to do?', student answers with an action.",
  prep: "Hook activates prior knowledge - students already read posters. Bridges to today: naming the three tools posters use. New learning, mixed readiness.",
  tag: "[Launch | Attention and prior knowledge | HITS 2, 7]",
});

const NOTES_LI_SC = T.composeGlanceNotes({
  beats: [
    "POINT to the learning intention. SAY: Today we learn how posters use three tools - heading, image and colour - to convince us.",
    "SAY: Read the I can statements with me. Choral read.",
    "SAY: The first one everyone can do - point to the heading, the image and the colour on a poster.",
  ],
  prep: "SC1 achievable by all (point to the 3 features). SC2 is the core target and the exit ticket (explain how a feature persuades). SC3 stretches to judging which feature convinces most.",
  tag: "[LI/SC | Planning made visible | HITS 1]",
});

const NOTES_VOCAB = T.composeGlanceNotes({
  answer: "persuade means to make someone want to do or believe something.",
  beats: [
    "SAY: Our key word today is persuade. Choral say it - persuade.",
    "SAY: To persuade is to make someone WANT to do or believe something. A poster persuades you to act.",
    "ASK: Use it - tell your partner one thing a poster once tried to persuade you to do. 20 sec. EXPECT: a real action, like buy something or come to an event.",
  ],
  trap: "treating persuade as just 'telling'. Fix: a poster does not only tell, it makes you WANT to; student re-says the meaning.",
  prep: "One key word, persuade, is the keystone for the whole lesson. Kept concrete; the poster examples carry the rest.",
  tag: "[Vocabulary | Knowledge and memory | HITS 3]",
});

const NOTES_IDO = T.composeGlanceNotes({
  answer: "the heading grabs you with big bold words, the image shows the problem and makes you care, the colour sets a mood that fits the message.",
  beats: [
    "SHOW the Save Water poster. SAY: Watch how I find the three tools this poster uses to persuade me.",
    "MODEL the heading. SAY: First the heading - SAVE EVERY DROP in huge bold letters. I cannot miss the main message.",
    "MODEL the image. SAY: The image is a dripping tap. It shows the problem and makes me care about wasted water.",
    "MODEL the colour. SAY: The whole poster is cool blue. Blue feels like water - clean and calm. The colour fits the message.",
    "ASK: Which tool tells you the main message? 10 sec, choral. EXPECT: the heading.",
  ],
  trap: "thinking the picture IS the message. Fix: point to the heading words, student names the heading as the message-carrier.",
  stretch: "if you could keep only ONE tool, which would you keep and why?",
  help: "give the three tool words on cards (heading, image, colour), student points to each one on the poster.",
  prep: "Core model. The wireframe poster is the anchor; each feature card names one tool and its job. Keep the poster up for the CFU.",
  tag: "[I Do | Explicit teaching | HITS 3, 4]",
});

const NOTES_CFU = T.composeGlanceNotes({
  answer: "H - the heading. Big bold words carry the main message and tell you what to do. The image shows the problem or feeling; colour sets the mood.",
  beats: [
    "SAY: Quick check. On the Save Water poster, three tools do three jobs.",
    "ASK: Which tool's main job is to tell you the MAIN MESSAGE? Show H, I or C on your board. 20 sec, boards up. EXPECT: H.",
    "SCAN boards, back row first. 80%+ show H -> reveal. Less -> cover the image and colour, read only the heading, re-ask.",
  ],
  trap: "picking I - thinking the picture is the message. Fix: cover the image, show the heading still tells you what to do, student re-answers H.",
  stretch: "what job does the IMAGE do instead of carrying the message?",
  help: "point to each part as you name its job, student chooses the heading.",
  prep: "Hinge on feature jobs (underpins SC2). Wrong answers map to misconceptions: I = picture-is-message, C = colour-is-message.",
  tag: "[CFU | Formative assessment | HITS 7, 8]",
});

const NOTES_CFU_REVEAL = T.composeGlanceNotes({
  beats: [
    "REVEAL after every board is up. SAY: The heading. Big bold words carry the message and tell you what to do.",
    "SAY: The image shows the problem and the colour sets the mood - important helpers, but the heading carries the message.",
  ],
  prep: "Reveal only after boards are scanned, so students commit first.",
  tag: "[CFU Reveal | Feedback | HITS 8]",
});

const NOTES_WEDO = T.composeGlanceNotes({
  answer: "open - Poster A convinces more (big clear heading, a book image, warm inviting colour). The feature doing the work is usually the heading or the image.",
  beats: [
    "SHOW both Readathon posters. SAY: Same message, two designs. Your turn with your partner.",
    "ASK: Which poster convinces you MORE? Boards: A or B. 20 sec. EXPECT: A.",
    "ASK: Which feature does the work? Tell your partner: 'A convinces me more because its ...'. 30 sec. EXPECT: names heading, image or colour with a reason.",
    "SCAN boards, back row first. 80%+ choose A with a reason -> reveal. Less -> hold the two headings side by side, re-ask.",
  ],
  trap: "picking B because it has 'more information'. Fix: more words is not more convincing; point to B's tiny heading and missing image, student re-judges.",
  stretch: "what ONE change would make poster B convince you? Name it.",
  help: "frame on the board: 'A convinces me more because its ___.' Student fills the blank with heading, image or colour.",
  prep: "Guided comparison - the visual stays on screen. Both posters carry the SAME message so students judge design, not content. Fade labels, not the poster.",
  tag: "[We Do | Supported application | HITS 5, 7]",
});

const NOTES_WEDO_REVEAL = T.composeGlanceNotes({
  beats: [
    "REVEAL after boards scanned. SAY: Most of us chose A. Its big bold heading and the book image pull you straight in.",
    "SAY: Warm orange feels friendly and exciting. Poster B hides its message in small grey text with no picture.",
  ],
  prep: "The reveal is one strong reading. Accept any answer that names a feature and gives a reason.",
  tag: "[We Do Reveal | Feedback | HITS 8]",
});

const NOTES_YOUDO = T.composeGlanceNotes({
  answer: "open - for the Walk to School poster, listen for: the heading grabs attention, the image shows the idea, the green colour feels healthy and outdoorsy.",
  beats: [
    "COLLECT the Poster Detective sheet. SAY: On your own now. Read the worked example, then be a detective on the Walk to School poster.",
    "SAY: For each tool - heading, image, colour - write HOW it helps persuade the reader.",
    "CIRCULATE, back row first. Ask: 'Read me your reason. How does that feature convince the reader?'",
  ],
  trap: "just describing the feature ('it is green') without the job. Fix: prompt 'how does green HELP persuade?', student adds the reason.",
  stretch: "the challenge on page 2: which feature convinces you most, and one design change that makes the poster even stronger.",
  help: "the sheet's worked example and the frame 'This feature helps persuade because ...' keep SC1 and SC2 reachable.",
  prep: "Independent application on a NEW poster (Walk to School). The sheet shows the SAME wireframe posters students saw on screen.",
  tag: "[You Do | Mastery and application | HITS 4, 10]",
});

const NOTES_EXIT = T.composeGlanceNotes({
  answer: "open - one feature named (heading, image or colour) plus HOW it persuades, using the frame 'This feature helps persuade because ...'.",
  beats: [
    "SHOW the prompt. SAY: Last task, on your own. Pick ONE feature from any poster you know.",
    "TIME 3 minutes. COLLECT boards or slips.",
  ],
  trap: "naming a feature but not how it persuades. Note who for tomorrow.",
  prep: "Assesses SC2 (explain how a feature persuades). The SC number stays in notes, not on the slide face.",
  tag: "[Exit Ticket | Evidence of learning | HITS 8]",
});

const NOTES_CLOSING = T.composeGlanceNotes({
  answer: "open - listen for a feature and its job, e.g. 'the heading grabs you so you read the message'.",
  beats: [
    "POINT to the I can statements. SAY: Show me thumbs for each one - up, sideways or down.",
    "ASK: Which poster tool convinces YOU most? Turn and tell one reason. 30 sec. EXPECT: a feature plus a reason.",
    "SAY: You can now spot the heading, image and colour AND say how each one persuades. That is a real reader's eye.",
  ],
  prep: "Revisits the three success criteria and the big idea: posters persuade with heading, image and colour working together.",
  tag: "[Closing | Review and reflect | HITS 9]",
});

// ---------------------------------------------------------------------------
// Success criteria (used on LI/SC and Closing - same order, plain "I can")
// ---------------------------------------------------------------------------

const LI = "We are learning how persuasive posters use headings, images and colour to convince the reader.";
const SC = [
  "I can point to the heading, the image and the colour on a poster",
  "I can explain how the heading, image or colour helps persuade the reader",
  "I can say which feature convinces me most and why",
];

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Persuasive Posters - Year 4 Literacy";

  // SLIDE 1 - Title
  titleSlide(
    pres,
    "How Posters Persuade You",
    "Headings, images and colour that make you want to act",
    "Year 4 Literacy  |  Persuasive Texts",
    NOTES_TITLE
  );

  // SLIDE 2 - Teacher Resources
  addResourceSlide(pres, RESOURCE_CONFIG, T, FOOTER, NOTES_RESOURCES);

  // SLIDE 3 - Launch (discover the poster's purpose)
  annotatedModelSlide(
    pres,
    "Launch",
    "What Is This Poster Trying To Do?",
    [
      { text: "Read it like a detective", role: "header" },
      "What does this poster want you to DO?",
      "What made you notice it first?",
      "Tell your partner one reason.",
    ],
    "School Fair poster",
    [
      { label: "Words", detail: "What do the big words say?", color: C.SECONDARY },
      { label: "Picture", detail: "What picture can you see?", color: C.ACCENT },
      { label: "Colour", detail: "What colour stands out most?", color: C.PRIMARY },
    ],
    NOTES_LAUNCH,
    FOOTER,
    {
      badgeColor: C.SECONDARY,
      leftW: 3.5,
      sourceType: "Persuasive poster (example)",
      previewSpec: POSTER_FAIR,
      previewAccent: FAIR_TEAL,
    }
  );

  // SLIDE 4 - LI / SC
  liSlide(pres, [LI], SC, NOTES_LI_SC, FOOTER);

  // SLIDE 5 - Vocabulary (persuade)
  vocabSlide(
    pres,
    "persuade",
    "verb",
    "To persuade means to make someone want to do or believe something. A persuasive poster is built to persuade you to act.",
    "The poster tries to persuade you to save water.",
    NOTES_VOCAB,
    FOOTER,
    { routine: "Say it. Show a persuading face. Use it in a sentence." }
  );

  // SLIDE 6 - I Do (teacher labels the 3 tools on the Save Water poster)
  annotatedModelSlide(
    pres,
    "I Do",
    "Three Tools That Persuade",
    [
      { text: "Watch me find the tools", role: "header" },
      "I find the heading, the image and the colour.",
      "For each one I ask: how does it convince me?",
    ],
    "Save Water poster",
    [
      { label: "Heading", detail: "SAVE EVERY DROP in big bold letters. You cannot miss the main message.", color: C.PRIMARY },
      { label: "Image", detail: "A photo of a dripping tap. It shows the problem and makes you care.", color: C.SECONDARY },
      { label: "Colour", detail: "Cool blue feels like water - clean and calm. It fits the message.", color: C.ACCENT },
    ],
    NOTES_IDO,
    FOOTER,
    {
      badgeColor: C.PRIMARY,
      badgeW: 1.5,
      leftW: 3.2,
      sourceType: "Persuasive poster (example)",
      previewSpec: POSTER_SAVE,
      previewAccent: SAVE_BLUE,
    }
  );

  // SLIDE 7 + 8 - CFU (hinge on feature jobs, with reveal)
  withReveal(
    () => cfuSlide(
      pres,
      "CFU",
      "Which Tool Carries The Message?",
      "Show Me Boards: H, I or C",
      "On the Save Water poster, which tool's main job is to tell you the main message?",
      NOTES_CFU,
      FOOTER
    ),
    (slide) => {
      addRevealAnswerBar(
        slide,
        "H - the heading. Big bold words carry the message and tell you what to do. The image shows the problem; colour sets the mood.",
        { label: "Answer", color: C.SUCCESS, showTickAndFix: false, fontSize: 14, y: 4.05, h: 0.95 }
      );
      slide.addNotes(NOTES_CFU_REVEAL);
    }
  );

  // SLIDE 9 + 10 - We Do (compare two posters, with reveal)
  withReveal(
    () => compareVisualSlide(
      pres,
      "We Do",
      "Which Poster Convinces You More?",
      "With your partner, on your boards: which poster convinces you more, A or B? Which feature does the work?",
      { panelTitle: "Poster A", title: "Readathon", previewSpec: POSTER_READ_STRONG, previewAccent: READ_ORANGE, strip: C.SECONDARY },
      { panelTitle: "Poster B", title: "Readathon", previewSpec: POSTER_READ_WEAK, previewAccent: READ_GREY, strip: C.MUTED },
      NOTES_WEDO,
      FOOTER,
      { badgeFill: C.SUCCESS }
    ),
    (slide) => {
      addRevealAnswerBar(
        slide,
        "Poster A. Its big bold heading and book image pull you in, and warm colour feels friendly. Poster B hides the message in small grey text with no picture.",
        { label: "Answer", color: C.SUCCESS, showTickAndFix: false, fontSize: 12.5, y: 4.2, h: 0.82 }
      );
      slide.addNotes(NOTES_WEDO_REVEAL);
    }
  );

  // SLIDE 11 - You Do (analyse a new poster on the printed sheet)
  annotatedModelSlide(
    pres,
    "You Do",
    "Be a Poster Detective",
    [
      { text: "Your turn on the sheet", role: "header" },
      "First: read the worked example.",
      "Next: find the heading, image and colour.",
      "Then: write HOW each one persuades.",
      { text: "Frame: This feature helps persuade because ...", role: "emphasis" },
    ],
    "Walk to School Week poster",
    [
      { label: "Heading", detail: "How does it grab you?", color: C.PRIMARY },
      { label: "Image", detail: "What does it show?", color: C.SECONDARY },
      { label: "Colour", detail: "What mood does it set?", color: C.ACCENT },
    ],
    NOTES_YOUDO,
    FOOTER,
    {
      badgeColor: C.PRIMARY,
      badgeW: 1.5,
      leftW: 3.5,
      sourceType: "Use your Poster Detective sheet",
      previewSpec: POSTER_WALK,
      previewAccent: WALK_GREEN,
    }
  );

  // SLIDE 12 - Exit Ticket
  exitTicketSlide(
    pres,
    ["Pick ONE feature (heading, image or colour) on a poster you know. Explain how it helps persuade the reader. Use: This feature helps persuade because ..."],
    NOTES_EXIT,
    FOOTER,
    { assessesSc: 2, badgeColor: C.ACCENT }
  );

  // SLIDE 13 - Closing
  closingSlide(
    pres,
    {
      reflectionPrompt: "Turn and tell: which poster tool convinces YOU most, and why?",
      scItems: SC,
      selfAssessment: {
        prompt: "Show on your fingers: 1 (need help) to 5 (could teach it)",
        options: ["1-2", "3", "4-5"],
      },
    },
    NOTES_CLOSING
  );

  // -------------------------------------------------------------------------
  // PDF: Poster Detective (student worksheet)
  // -------------------------------------------------------------------------
  const ws = createPdf({ title: DETECTIVE_RESOURCE.name });

  let wy = addPdfHeader(ws, "Poster Detective", {
    color: WALK_GREEN,
    subtitle: "Find the three tools every persuasive poster uses: heading, image and colour.",
    lessonInfo: "Year 4 Literacy | Persuasive Posters",
    showNameDate: true,
  });

  wy = addTipBox(ws,
    "Every persuasive poster uses three tools: a HEADING, an IMAGE and COLOUR. Each tool has a job. A detective finds each tool and works out HOW it convinces the reader.",
    wy, { color: WALK_GREEN });

  // Worked example: the Save Water poster (same wireframe students saw on screen)
  wy = addSectionHeading(ws, "Watch me first (worked example)", wy, { color: SAVE_BLUE });
  {
    const posterW = 132;
    const posterH = 118;
    const posterX = PAGE.MARGIN;
    const textX = posterX + posterW + 16;
    const textW = PAGE.CONTENT_W - posterW - 16;
    addPosterMockupPdf(ws, posterX, wy, posterW, posterH, POSTER_SAVE);
    const lines = [
      ["Heading", "SAVE EVERY DROP in big bold letters. Job: grab your eyes with the main message."],
      ["Image", "a photo of a dripping tap. Job: show the problem and make you care."],
      ["Colour", "cool blue. Job: feel like water, clean and calm - it fits the message."],
    ];
    let ty = wy + 2;
    lines.forEach(([label, detail]) => {
      ws.fontSize(10).font("Sans-Bold").fillColor(hex(SAVE_BLUE));
      ws.text(label + " - ", textX, ty, { width: textW, continued: true });
      ws.font("Sans").fillColor(hex("2D3142"));
      ws.text(detail, { width: textW });
      ty = ws.y + 6;
    });
    wy = Math.max(wy + posterH, ty) + 12;
  }

  // Student task: the Walk to School poster + three labelled reason boxes
  wy = addSectionHeading(ws, "Your turn - be a detective", wy, { color: WALK_GREEN });
  wy = addBodyText(ws,
    "Look at the Walk to School Week poster. Find each tool, then write HOW it helps persuade the reader.",
    wy, { fontSize: 10, italic: true });
  {
    const posterW = 200;
    const posterH = 148;
    const posterX = PAGE.MARGIN + (PAGE.CONTENT_W - posterW) / 2;
    addPosterMockupPdf(ws, posterX, wy, posterW, posterH, POSTER_WALK);
    wy = wy + posterH + 12;
  }

  const reasonBox = (doc, label, y) => {
    doc.fontSize(11).font("Sans-Bold").fillColor(hex(WALK_GREEN));
    doc.text(label + " - how does it help persuade?", PAGE.MARGIN, y, { width: PAGE.CONTENT_W });
    return addLinedArea(doc, y + 18, 2, { lineSpacing: 24 }) + 8;
  };
  wy = reasonBox(ws, "Heading", wy);
  wy = reasonBox(ws, "Image", wy);
  wy = reasonBox(ws, "Colour", wy);

  addPdfFooter(ws, "Poster Detective | Year 4 Literacy - Page 1");

  // Page 2 - challenge / extension (deepens the same concept)
  ws.addPage();
  let wy2 = addPdfHeader(ws, "Poster Designer (Challenge)", {
    color: WALK_GREEN,
    subtitle: "Judge the poster, then improve it.",
    lessonInfo: "Year 4 Literacy | Persuasive Posters",
    showNameDate: false,
  });

  wy2 = addSectionHeading(ws, "Which tool convinces YOU most?", wy2, { color: C.PRIMARY });
  wy2 = addBodyText(ws,
    "Write the tool (heading, image or colour), then explain why on the lines below.",
    wy2, { fontSize: 10, italic: true });
  wy2 = addBodyText(ws, "The tool that convinces me most is ...", wy2, { fontSize: 12 });
  wy2 = addLinedArea(ws, wy2 + 4, 3, { lineSpacing: 26 }) + 10;

  wy2 = addSectionHeading(ws, "Design challenge", wy2, { color: SAVE_BLUE });
  wy2 = addBodyText(ws,
    "You are the poster designer. Write ONE change that would make the Walk to School poster even more convincing, and say why it would work.",
    wy2, { fontSize: 10, italic: true });
  wy2 = addLinedArea(ws, wy2 + 4, 4, { lineSpacing: 26 });

  addPdfFooter(ws, "Poster Detective | Year 4 Literacy - Page 2");

  // -------------------------------------------------------------------------
  // PDF: Answer Key
  // -------------------------------------------------------------------------
  const ak = createPdf({ title: ANSWER_RESOURCE.name });

  let ay = addPdfHeader(ak, "Poster Detective - Answer Key", {
    color: WALK_GREEN,
    subtitle: "Sample answers. Many good answers are possible.",
    lessonInfo: "Year 4 Literacy | Persuasive Posters",
    showNameDate: false,
  });

  ay = addTipBox(ak,
    "These are samples, not the only answers. Look for the feature named AND how it persuades, not just what it looks like.",
    ay, { color: WALK_GREEN });

  ay = addSectionHeading(ak, "Walk to School poster - sample answers", ay, { color: WALK_GREEN });
  ay = addBodyText(ak, "Heading - WALK TO SCHOOL WEEK in big bold letters. It grabs your eyes and tells you exactly what to do.", ay, { fontSize: 11 });
  ay = addBodyText(ak, "Image - a photo of children walking (or shoes on a path). It shows the idea and makes walking look fun and normal.", ay, { fontSize: 11 });
  ay = addBodyText(ak, "Colour - fresh green. Green feels healthy and outdoorsy, so it fits the message about walking.", ay, { fontSize: 11 });
  ay += 4;

  ay = addSectionHeading(ak, "Challenge - teacher look-fors", ay, { color: C.PRIMARY });
  ay = addBodyText(ak, "- Names one feature (heading, image or colour) and gives a reason, not just a description.", ay, { fontSize: 10 });
  ay = addBodyText(ak, "- The design change is specific (bigger heading, add a clear picture, brighter colour) and linked to persuading the reader.", ay, { fontSize: 10 });
  ay = addBodyText(ak, "- Uses the frame: This feature helps persuade because ...", ay, { fontSize: 10 });

  addPdfFooter(ak, "Poster Detective Answer Key | Year 4 Literacy");

  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Persuasive Posters.pptx` }),
    writePdf(ws, DETECTIVE_PDF_PATH),
    writePdf(ak, ANSWER_PDF_PATH),
  ]);

  console.log("PPTX written to " + `${OUT_DIR}/Persuasive Posters.pptx`);
  console.log("Done: " + DETECTIVE_RESOURCE.name);
  console.log("Done: " + ANSWER_RESOURCE.name);
}

build().catch((err) => { console.error(err); process.exit(1); });
