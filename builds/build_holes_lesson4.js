"use strict";

// Holes Unit - Lesson 4 (Week 2, Session 1): Persuasive Structure -- Should Camp Green Lake Be Shut Down?
// Year 5/6 Literacy
// Focus: refresh persuasive structure (introduction with position) + introduce the unit's debate topic
// Anchor: Holes by Louis Sachar, Chapters 5-7 (camp conditions: heat, lack of water, dangerous work)

const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");

const { createTheme, weekToVariant } = require("../themes/factory");
const T = createTheme("literacy", "grade56", weekToVariant(1));
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  addTopBar, addBadge, addTitle, addCard, addFooter,
  withReveal,
  titleSlide, liSlide, contentSlide,
  cfuSlide, closingSlide,
  modellingSlide, vocabSlide, quoteSlide,
} = T;

const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addLinedArea,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const SESSION_NUMBER = 4;
const FOOTER = "Persuasive Structure | Lesson 4 | Week 2 | Year 5/6 Literacy | Holes";
const OUT_DIR = "output/Holes_Lesson4_Persuasive_Structure";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const PLAN_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Persuasive Plan Camp Green Lake",
  "Student template: pick a position and plan a persuasive paragraph introduction."
);
const MENTOR_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Mentor Persuasive Introduction",
  "Annotated model persuasive introduction -- hook, position and signpost to reasons."
);
const RESOURCE_ITEMS = [PLAN_RESOURCE, MENTOR_RESOURCE];
const PLAN_PDF_PATH = path.join(OUT_DIR, PLAN_RESOURCE.fileName);
const MENTOR_PDF_PATH = path.join(OUT_DIR, MENTOR_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Welcome to Week 2 -- persuasive writing
- This week's debate: Should Camp Green Lake be SHUT DOWN?
- Four sessions to plan and write a short persuasive piece
- We will revisit ideas you learned 4 months ago and bring them back together

DO:
- Display title slide
- Briefly preview the week: today structure, then arguments, then emotive language, then conclusion / counter-argument
- Have the novel ready -- a quick re-read from Ch 5-7 to remind students of camp conditions

TEACHER NOTES:
Students learned persuasive writing earlier in the year (about 4 months ago). Expect them to remember some structure but not all techniques. Today's lesson REFRESHES the structure and locks in the unit's topic.

WATCH FOR:
- Students who remember terms like "introduction" or "argument" -- celebrate and connect to today's slides
- Students with no recall -- reassure: "We will rebuild it together"

[Literacy: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- Two resources today
- The ${PLAN_RESOURCE.name} -- pick a side and plan your introduction
- The ${MENTOR_RESOURCE.name} -- a model introduction with annotations

DO:
- Print the plan (one per student)
- Print the mentor (one per student or pair)
- Have the novel handy for a short re-read

[Literacy: Resources | VTLM 2.0: Student Resources]`;

const NOTES_HOOK = `SAY:
- A quick refresh from Holes
- I am going to read a SHORT bit from Chapters 5 to 7 -- what Camp Green Lake is REALLY like
- Listen for the DETAILS that make this place rough: the heat, the water, the holes the boys dig

DO:
- Read a SHORT passage from Holes Chapter 5, 6 or 7 -- the descriptions of camp life (the heat, the canteen water, digging holes the size of the shovel)
- 1 to 2 minutes of reading is enough
- Quick partner share: "Would YOU want to be sent to Camp Green Lake?"
- Cold call 2-3 students -- expect almost all to say no
- Briefly: "Today we ARGUE that case. We persuade"

TEACHER NOTES:
The launch sets up the persuasive topic by activating real concrete details from the novel. Students will naturally lean toward "shut it down" but some will see the argument for "keep it open" (it might teach the boys responsibility). Both sides are valid -- you will let them choose.

WATCH FOR:
- Students who note specific camp conditions -- celebrate (these become arguments later)
- Students who say "Stanley likes it" -- redirect: "Did he WANT to go? Or did he have to?"

[Literacy: Text Launch | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_LI_SC = `SAY:
- Read the LI with me
- We are learning to WRITE a persuasive introduction with three parts
- Three "I can" statements

DO:
- Choral read LI and SCs
- Brief: "An introduction has 3 jobs: HOOK the reader, state your POSITION, signpost your REASONS"

TEACHER NOTES:
SC1 -- one clear sentence of position -- achievable for all. SC2 (target) -- hook + position + 3 reasons named. SC3 -- a strong rhetorical hook (question / statement / fact). Exit ticket targets SC2.

WATCH FOR:
- Students who think they need to write the whole essay today -- redirect: "Just the OPENING paragraph"

[Literacy: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_VOCAB_INHUMANE = `SAY:
- One word today: INHUMANE
- An adjective. It means cruel -- not fitting for a human being to experience
- Used a lot in persuasive writing about how people are treated

DO:
- Choral say INHUMANE
- Quick example: "Forcing children to dig holes in 35-degree heat all day is INHUMANE"
- Ask: "What other situations might be called inhumane?" -- 30 seconds partner share

TEACHER NOTES:
Inhumane is a strong word students can use as evidence-of-feeling in their arguments. Pair it with concrete examples from the novel.

WATCH FOR:
- Students using it accurately -- celebrate
- Students using it for small things ("missing the bus is inhumane") -- redirect: "Save it for big cruelty"

[Literacy: Vocabulary | VTLM 2.0: Vocabulary Instruction]`;

const NOTES_STRUCTURE = `SAY:
- The structure for a persuasive introduction
- Three parts: HOOK, POSITION, SIGNPOST
- Look at the slide. Each part has a job

DO:
- Choral read each row
- Quick check: "Which one is the FIRST sentence?" (HOOK)
- Quick check: "Which one tells me which SIDE you are on?" (POSITION)
- Quick check: "Which one previews the reasons?" (SIGNPOST)

TEACHER NOTES:
This is the structural backbone for the week. Students will use it today AND next session (when they write body arguments).

WATCH FOR:
- Students who confuse hook with position -- redirect: "The hook GRABS. The position SAYS WHICH SIDE"
- Students who try to write all 3 parts in one sentence -- allow it for one strong writer, redirect most

[Literacy: Structure | VTLM 2.0: Explicit Teaching / Structure]`;

const NOTES_IDO = `SAY:
- Watch me write a persuasive introduction
- The topic: Should Camp Green Lake be shut down?
- I will pick the SHUT-DOWN side for the model -- you can pick either side
- Watch my CHOICES: hook first, then position, then signpost

DO:
- Display the I Do slide
- Read the model aloud
- Highlight each part:
  - HOOK (the rhetorical question)
  - POSITION (the clear "I believe" sentence)
  - SIGNPOST (the three reasons named)
- After reading: "Notice -- only ONE position sentence. Only THREE reasons signposted. Tight"

TEACHER NOTES:
The model picks the shut-down side. Acknowledge in passing that students may pick the OTHER side. Tomorrow they write body paragraphs using these signposted reasons.

WATCH FOR:
- Students who notice the rhetorical question -- celebrate
- Students who try to memorise the sentence -- redirect: "Your hook will be different. Your three reasons will be different"

[Literacy: I Do | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_WEDO = `SAY:
- Together. We plan a persuasive introduction for the OTHER side
- The topic: Camp Green Lake should STAY OPEN -- it teaches boys responsibility and consequences
- We are not picking that side because we believe it -- we are practising the STRUCTURE

DO:
- Display the We Do slide
- Cold call for HOOK ideas (a rhetorical question? a surprising statement?)
- Build a POSITION sentence together
- Cold call for THREE reasons (responsibility, consequences, hard work, time outdoors)
- Aim: full hook + position written on the board

TEACHER NOTES:
We Do uses the OTHER side from I Do -- this teaches that the STRUCTURE works no matter the position. Some students will argue passionately. Channel that energy back into the STRUCTURE.

WATCH FOR:
- Students who refuse to argue this side -- offer them a different example: "We are practising the structure. Try a topic you DO agree with -- you should always get 10 hours of sleep"
- Students who name strong reasons -- celebrate; some will use these in their writing

[Literacy: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_CFU_BUILD = `SAY:
- Quick check. Two introductions about whether kids should have homework. Which ONE follows the structure -- hook, position, signpost?
- Read carefully
- Show me A or B on your fingers

DO:
- Display both
- Show Me Fingers
- Scan: most students should choose B
- Cold call 1-2 students: "Why B?"

TEACHER NOTES:
A is one long sentence with no clear structure. B has a hook (question), a position ("homework should be limited") and a signpost ("three reasons"). B is the stronger introduction.

WATCH FOR:
- Students who choose B and name the three parts -- ready
- Students who choose A because it has "more reasons" -- redirect: "Reasons go in the BODY. The introduction signposts them"

[Literacy: CFU | VTLM 2.0: Formative Assessment]`;

const NOTES_CFU_REVEAL = `SAY:
- Stronger introduction: B
- B has all three: HOOK (the question), POSITION ("homework should be limited"), SIGNPOST ("three reasons")
- A has reasons mashed in but no structure

DO:
- Display the reveal banner
- Read B aloud
- Point to each of the three structural parts

CFU CHECKPOINT:
Technique: Show Me Fingers (1 for A, 2 for B)
Script:
- "Which introduction follows hook + position + signpost?"
- Scan for: most students choose B (>=80%)
PROCEED (>=80%): Release to You Do.
PIVOT (<80%): Use the optional re-teach slide that follows -- it shows the structure as three SEPARATE sentences side by side.

[Literacy: CFU Reveal | VTLM 2.0: Formative Assessment]`;

const NOTES_RETEACH = `SAY:
- Quick second look at the structure
- Watch me BUILD an introduction one sentence at a time
- Topic: school uniforms should be banned
- Sentence 1 -- the HOOK. A question: "Why do we wear the same colour every day?"
- Sentence 2 -- the POSITION. Clear: "School uniforms should be banned."
- Sentence 3 -- the SIGNPOST. The three reasons: "Uniforms are uncomfortable, expensive and stop students from showing who they are."
- Three sentences. Three jobs. One introduction

DO:
- Display the re-teach slide
- Build LIVE on the board, sentence by sentence
- Use a different topic so it does not feel like a repeat
- Re-check: ask students to write ONE position sentence on their boards for "Phones should be banned in schools"

TEACHER NOTES:
OPTIONAL. Use only if CFU was below 80%. Different approach: three SEPARATE sentences, built one at a time, on a different topic. Slows the structure down.

WATCH FOR:
- Students who write a clear position sentence ("Phones should be banned in schools" or "Phones should be allowed in schools") -- ready
- Students who write a reason ("phones are distracting") -- redirect: "That is a REASON. The position SAYS WHICH SIDE"

[Literacy: Optional Re-teach | VTLM 2.0: Reteach with Different Approach]`;

const NOTES_YOUDO = `SAY:
- Your turn
- Pick your SIDE: should Camp Green Lake be shut down? YES or NO
- Plan your introduction: hook, position, signpost three reasons
- Then write the introduction. Three to four sentences only
- 15 minutes. I will circulate

DO:
- Display the You Do slide
- Distribute plan and mentor
- Circulate -- prioritise students who needed the re-teach
- Quick conferences: "Read me your position sentence. Which side are you on?"

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the mentor sheet. Look at the three highlighted parts. Use the sentence frames on the back. Pick whichever side feels easier
- Extra Notes: Frames give every student a structure that hits SC1 and SC2

EXTENDING PROMPT:
- Task: After your introduction, write your THREE reasons as topic sentences for the body paragraphs you will write next lesson. ("Firstly..." / "Secondly..." / "Most importantly...")
- Extra Notes: This gets confident writers ahead for tomorrow

TEACHER NOTES:
You Do is the same topic for everyone (Camp Green Lake) -- the choice is which SIDE. Tomorrow students write body paragraphs from the reasons they signposted today, so capture good thinking in the plan.

WATCH FOR:
- Students changing sides mid-paragraph -- redirect: "Pick ONE side. Stick with it"
- Students writing a long hook -- prompt: "One sentence. Make me curious"
- Students with strong signposted reasons -- celebrate; flag for tomorrow

[Literacy: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket
- On your plan sheet, write ONE position sentence for this topic: Should students have a 10-minute brain break every hour?
- One clear sentence. Use "I believe..." or "I think..." or simply make the claim
- 2 minutes

DO:
- 2 minutes silent
- Collect

TEACHER NOTES:
Exit ticket targets SC1 (position sentence). One sentence is enough.

WATCH FOR:
- Students writing a reason instead of a position -- collect; flag

[Literacy: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Self-check
- Show on fingers: 1 to 5 for each SC
- Partner share: which side did you pick on Camp Green Lake -- and what is your strongest reason?

DO:
- Run fingers check for each SC
- 60 seconds partner share
- Briefly: "Tomorrow we take your three reasons and turn them into BODY paragraphs"

TEACHER NOTES:
This sets up Lesson 5. Collect the strongest reasons you hear -- you may use them as examples.

[Literacy: Closing | VTLM 2.0: Review and Reflect]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Holes -- Lesson 4 -- Persuasive Structure";

  // SLIDE 1 -- Title
  titleSlide(
    pres,
    "Should Camp Green Lake Be Shut Down?",
    "Persuasive Writing -- Introduction Structure",
    "Lesson 4  |  Week 2  |  Year 5/6 Literacy",
    NOTES_TITLE
  );

  // SLIDE 2 -- Resources
  addResourceSlide(
    pres,
    RESOURCE_ITEMS,
    T,
    FOOTER,
    NOTES_RESOURCES
  );

  // SLIDE 3 -- Text Launch (re-read of Ch 5-7)
  quoteSlide(
    pres,
    "Re-read",
    "Holes -- Chapters 5 to 7 (camp life)",
    "Heat. Dust. A hole every day, every boy, the same size as the shovel.",
    "Ch. 5-7, camp conditions",
    "Partner talk: would YOU want to be sent to Camp Green Lake? Name ONE reason.",
    NOTES_HOOK,
    FOOTER
  );

  // SLIDE 4 -- LI / SC
  liSlide(
    pres,
    [
      "We are learning to write a persuasive INTRODUCTION using HOOK + POSITION + SIGNPOST",
    ],
    [
      "I can write one clear sentence that states my position",
      "I can write a hook, a position and a signpost of three reasons",
      "I can write a strong rhetorical hook (a question, a surprising fact, or a bold statement)",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // SLIDE 5 -- Vocab
  vocabSlide(
    pres,
    "inhumane",
    "adjective",
    "Cruel -- treating people in a way that is not fitting for a human being. Used a lot in persuasive writing about how people are treated.",
    "Forcing children to dig holes all day in extreme heat is inhumane.",
    NOTES_VOCAB_INHUMANE,
    FOOTER
  );

  // SLIDE 6 -- Structure overview (the three parts)
  contentSlide(
    pres,
    "Structure",
    C.SECONDARY,
    "Persuasive Introduction -- 3 Parts",
    [
      "HOOK:        first sentence. Grabs the reader (a question, a surprising fact, a bold statement)",
      "POSITION:  one clear sentence. WHICH SIDE are you on? (\"I believe Camp Green Lake should be shut down.\")",
      "SIGNPOST:  one sentence that previews your three reasons (\"There are three reasons: it is unsafe, it is inhumane, and it does not work.\")",
    ],
    NOTES_STRUCTURE,
    FOOTER
  );

  // SLIDE 7 -- I Do
  modellingSlide(
    pres,
    "I Do -- Watch Me",
    "Model Introduction: Shut Down Camp Green Lake",
    "Topic:\n\nShould Camp Green Lake be shut down?\n\nMy side: YES, shut it down.\n\nWatch the three parts:\n\n- HOOK (rhetorical question)\n- POSITION (clear stand)\n- SIGNPOST (three reasons named)",
    "\"Would you send your worst enemy to a place where children dig holes in 35-degree heat with one canteen of water a day? I strongly believe Camp Green Lake must be shut down. There are three reasons: the camp is unsafe, the work is inhumane, and -- worst of all -- it does not actually help the boys learn anything.\"\n\n(My choices:\n- HOOK: rhetorical question\n- POSITION: \"must be shut down\"\n- SIGNPOST: three reasons named\n- bonus: emotive language already at work: \"35-degree heat\", \"worst of all\")",
    NOTES_IDO,
    FOOTER
  );

  // SLIDE 8 -- We Do
  {
    const s = pres.addSlide();
    addTopBar(s, C.SECONDARY);
    addBadge(s, "We Do", { color: C.SECONDARY });
    addTitle(s, "Together: The OTHER Side -- Keep Camp Green Lake Open");

    const cardY = CONTENT_TOP;
    const cardH = 2.0;
    addCard(s, 0.5, cardY, 9, cardH, { strip: C.SECONDARY, fill: C.WHITE });
    s.addText("Plan together -- one introduction for the OTHER side", {
      x: 0.75, y: cardY + 0.10, w: 8.5, h: 0.30,
      fontSize: 14, fontFace: FONT_B, color: C.SECONDARY, bold: true, margin: 0,
    });
    s.addText("- HOOK:        what surprising fact or question could grab the reader? (e.g. \"What if hard work could change a boy's life?\")\n- POSITION:  \"Camp Green Lake should stay open.\"\n- SIGNPOST: three possible reasons -- teaches responsibility, builds resilience, gives second chance", {
      x: 0.75, y: cardY + 0.45, w: 8.5, h: cardH - 0.55,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    const tipY = cardY + cardH + 0.18;
    const tipH = SAFE_BOTTOM - tipY;
    addCard(s, 0.5, tipY, 9, tipH, { strip: C.ACCENT, fill: C.BG_CARD });
    s.addText("Why are we writing the OTHER side?", {
      x: 0.75, y: tipY + 0.10, w: 8.5, h: 0.30,
      fontSize: 13, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("- Practising the STRUCTURE -- it works on any topic\n- You will pick YOUR side in You Do\n- Strong writers can argue any side", {
      x: 0.75, y: tipY + 0.42, w: 8.5, h: tipH - 0.55,
      fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_WEDO);
  }

  // SLIDE 9 + 10 -- CFU
  function buildCfuBase() {
    const slide = pres.addSlide();
    addTopBar(slide, C.ALERT);
    addBadge(slide, "CFU", { color: C.ALERT });
    addTitle(slide, "Which Introduction Follows HOOK + POSITION + SIGNPOST?", { color: C.ALERT });

    const stampW = 1.3;
    slide.addShape("roundRect", {
      x: 9.5 - stampW, y: 0.20, w: stampW, h: 0.32, rectRadius: 0.08,
      fill: { color: C.WHITE }, line: { color: C.ALERT, width: 1.5 },
    });
    slide.addText("CHECK", {
      x: 9.5 - stampW, y: 0.20, w: stampW, h: 0.32,
      fontSize: 11, fontFace: FONT_B, color: C.ALERT, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    slide.addShape("roundRect", {
      x: 0.5, y: CONTENT_TOP, w: 3.2, h: 0.40, rectRadius: 0.08,
      fill: { color: C.ALERT },
    });
    slide.addText("Show Me Fingers: 1 (A) or 2 (B)", {
      x: 0.5, y: CONTENT_TOP, w: 3.2, h: 0.40,
      fontSize: 12, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    const cardY = CONTENT_TOP + 0.55;
    const cardH = 1.5;
    addCard(slide, 0.5, cardY, 9, cardH, { strip: C.MUTED, fill: C.WHITE });
    slide.addText("A", {
      x: 0.7, y: cardY + 0.10, w: 0.5, h: 0.4,
      fontSize: 22, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0,
    });
    slide.addText("\"Homework is bad because it takes too much time and stops you spending time with your family and is boring and most kids don't want to do it so it should be banned.\"", {
      x: 1.2, y: cardY + 0.12, w: 8.0, h: cardH - 0.24,
      fontSize: 12, fontFace: FONT_H, color: C.CHARCOAL, italic: true, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    const cardBY = cardY + cardH + 0.18;
    const cardBH = SAFE_BOTTOM - cardBY - 0.55;
    addCard(slide, 0.5, cardBY, 9, cardBH, { strip: C.SECONDARY, fill: C.BG_CARD });
    slide.addText("B", {
      x: 0.7, y: cardBY + 0.10, w: 0.5, h: 0.4,
      fontSize: 22, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0,
    });
    slide.addText("\"Should we really spend more time on schoolwork after school ends? Homework should be limited to one short task per night. There are three reasons: it eats family time, it adds stress, and it does not actually improve learning.\"", {
      x: 1.2, y: cardBY + 0.12, w: 8.0, h: cardBH - 0.24,
      fontSize: 12, fontFace: FONT_H, color: C.CHARCOAL, italic: true, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    addFooter(slide, FOOTER);
    slide.addNotes(NOTES_CFU_BUILD);
    return slide;
  }

  withReveal(
    buildCfuBase,
    (slide) => {
      const revealY = 4.68;
      addCard(slide, 0.5, revealY, 9, 0.42, { fill: C.SUCCESS });
      slide.addText("Stronger introduction: B  --  has HOOK (question), POSITION (\"should be limited\") and SIGNPOST (three reasons)", {
        x: 0.5, y: revealY, w: 9, h: 0.42,
        fontSize: 12, fontFace: FONT_B, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      slide.addNotes(NOTES_CFU_REVEAL);
    }
  );

  // SLIDE 11 -- Optional Re-teach
  modellingSlide(
    pres,
    "Optional Re-teach",
    "Three Sentences. Three Jobs.",
    "Topic: school uniforms should be banned\n\nWatch me build it one sentence at a time. Each sentence has ONE job.",
    "Sentence 1 -- HOOK (a question):\n\"Why do we wear the same colour every day?\"\n\nSentence 2 -- POSITION (clear stand):\n\"School uniforms should be banned.\"\n\nSentence 3 -- SIGNPOST (three reasons):\n\"They are uncomfortable, expensive and stop students from showing who they are.\"\n\nYour turn: write ONE position sentence on your board for \"Phones in schools\".",
    NOTES_RETEACH,
    FOOTER
  );

  // SLIDE 12 -- You Do
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "You Do", { color: C.PRIMARY, w: 1.5 });
    addTitle(s, "Write YOUR Persuasive Introduction");

    addCard(s, 0.5, CONTENT_TOP, 9, 1.85, { strip: C.PRIMARY, fill: C.WHITE });
    s.addText("Topic:  Should Camp Green Lake be shut down?", {
      x: 0.75, y: CONTENT_TOP + 0.10, w: 8.4, h: 0.30,
      fontSize: 14, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("Pick YOUR side:\n- YES, shut it down\n- NO, keep it open\n\nWrite an INTRODUCTION only (3 to 4 sentences).", {
      x: 0.75, y: CONTENT_TOP + 0.45, w: 8.4, h: 1.30,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    const tipY = CONTENT_TOP + 2.00;
    const tipH = SAFE_BOTTOM - tipY;
    addCard(s, 0.5, tipY, 9, tipH, { strip: C.ACCENT, fill: C.BG_CARD });
    s.addText("Writing Time: 15 minutes", {
      x: 0.75, y: tipY + 0.10, w: 5, h: 0.30,
      fontSize: 13, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("First:    Use the Plan -- HOOK + POSITION + 3 REASONS\nNext:    Write the introduction (3 to 4 sentences)\nThen:    Read it back aloud -- does it sound persuasive?\nTomorrow: we turn your 3 reasons into 3 body paragraphs", {
      x: 0.75, y: tipY + 0.42, w: 8.4, h: tipH - 0.55,
      fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  }

  // SLIDE 13 -- Exit Ticket
  contentSlide(
    pres,
    "Exit Ticket",
    C.ACCENT,
    "One Position Sentence -- a 10-Minute Brain Break Every Hour?",
    [
      "Should students have a 10-minute brain break every hour?",
      "Write ONE clear position sentence",
      "Use \"I believe...\" or \"I think...\" or make the claim straight",
      "2 minutes -- drop on my desk",
    ],
    NOTES_EXIT,
    FOOTER
  );

  // SLIDE 14 -- Closing
  closingSlide(
    pres,
    {
      reflectionPrompt: "Partner share: which side did you pick on Camp Green Lake -- and what is your strongest reason?",
      scItems: [
        "I can write one clear sentence that states my position",
        "I can write a hook, a position and a signpost of three reasons",
        "I can write a strong rhetorical hook (a question, a surprising fact, or a bold statement)",
      ],
      selfAssessment: {
        prompt: "Show on your fingers: 1 (need help) to 5 (could teach it)",
        options: ["1-2", "3", "4-5"],
      },
    },
    NOTES_CLOSING
  );

  // ---- PDF: Persuasive Plan -------------------------------------------------
  const plan = createPdf({ title: PLAN_RESOURCE.name });
  let planY = addPdfHeader(plan, "Persuasive Plan -- Camp Green Lake", {
    color: C.PRIMARY,
    subtitle: "Pick your side, then plan your introduction (HOOK + POSITION + SIGNPOST)",
    lessonInfo: "Lesson 4 | Week 2 | Year 5/6 Literacy | Holes",
    showNameDate: true,
  });

  planY = addTipBox(plan, "Today you write the INTRODUCTION only -- 3 to 4 sentences. Tomorrow you write the body. Plan first, then write.", planY, { color: C.PRIMARY });

  planY = addSectionHeading(plan, "Step 1 -- Pick your side", planY, { color: C.PRIMARY });
  planY = addBodyText(plan, "Should Camp Green Lake be shut down? Circle ONE:  YES, shut it down  /  NO, keep it open", planY, { fontSize: 11 });
  planY += 4;

  planY = addSectionHeading(plan, "Step 2 -- HOOK (first sentence)", planY, { color: C.PRIMARY });
  planY = addBodyText(plan, "A question? A surprising fact? A bold statement? Grab the reader.", planY, { fontSize: 10, italic: true });
  planY = addLinedArea(plan, planY, 2, { lineSpacing: 22 });
  planY += 4;

  planY = addSectionHeading(plan, "Step 3 -- POSITION (one clear sentence)", planY, { color: C.SECONDARY });
  planY = addBodyText(plan, "\"I believe...\" or \"...should be...\" -- which side are you on?", planY, { fontSize: 10, italic: true });
  planY = addLinedArea(plan, planY, 2, { lineSpacing: 22 });
  planY += 4;

  planY = addSectionHeading(plan, "Step 4 -- SIGNPOST (three reasons)", planY, { color: C.ACCENT });
  planY = addBodyText(plan, "List the THREE reasons you will use in your body paragraphs tomorrow.", planY, { fontSize: 10, italic: true });
  planY = addBodyText(plan, "Reason 1:", planY, { fontSize: 11 });
  planY = addLinedArea(plan, planY, 1, { lineSpacing: 18 });
  planY = addBodyText(plan, "Reason 2:", planY, { fontSize: 11 });
  planY = addLinedArea(plan, planY, 1, { lineSpacing: 18 });
  planY = addBodyText(plan, "Reason 3 (your strongest):", planY, { fontSize: 11 });
  planY = addLinedArea(plan, planY, 1, { lineSpacing: 18 });

  addPdfFooter(plan, "Lesson 4 | Persuasive Plan -- Page 1");

  plan.addPage();
  let planY2 = addPdfHeader(plan, "Persuasive Introduction -- Write Here", {
    color: C.PRIMARY,
    subtitle: "3 to 4 sentences only. Tomorrow you write the body.",
    lessonInfo: "Lesson 4 | Week 2 | Year 5/6 Literacy",
    showNameDate: false,
  });

  planY2 = addTipBox(plan, "Use your plan. Open with the hook. State your position clearly. Signpost the three reasons. Read it back -- does it sound persuasive?", planY2, { color: C.SECONDARY });

  planY2 = addLinedArea(plan, planY2, 14, { lineSpacing: 22 });

  addPdfFooter(plan, "Lesson 4 | Persuasive Introduction -- Page 2");

  plan.addPage();
  let planY3 = addPdfHeader(plan, "Exit Ticket", {
    color: C.ACCENT,
    subtitle: "One position sentence -- a 10-minute brain break every hour",
    lessonInfo: "Lesson 4 | Week 2",
    showNameDate: false,
  });

  planY3 = addTipBox(plan, "Write ONE clear position sentence. Should students have a 10-minute brain break every hour? Make your claim.", planY3, { color: C.ACCENT });

  planY3 = addLinedArea(plan, planY3, 3, { lineSpacing: 24 });

  addPdfFooter(plan, "Lesson 4 | Exit Ticket");

  // ---- PDF: Mentor Persuasive Introduction ----------------------------------
  const mp = createPdf({ title: MENTOR_RESOURCE.name });
  let mpY = addPdfHeader(mp, "Mentor Persuasive Introduction -- Annotated", {
    color: C.PRIMARY,
    subtitle: "Camp Green Lake should be shut down -- a model introduction",
    lessonInfo: "Lesson 4 | Week 2 | Year 5/6 Literacy",
    showNameDate: false,
  });

  mpY = addTipBox(mp, "This is a model introduction. It uses HOOK + POSITION + SIGNPOST. Use the PATTERN -- your hook, position and reasons will be different.", mpY, { color: C.PRIMARY });

  mpY = addSectionHeading(mp, "Model introduction", mpY, { color: C.PRIMARY });
  mpY = addBodyText(mp, "\"Would you send your worst enemy to a place where children dig holes in 35-degree heat with one canteen of water a day? I strongly believe Camp Green Lake must be shut down. There are three reasons: the camp is unsafe, the work is inhumane, and -- worst of all -- it does not actually help the boys learn anything.\"", mpY, { fontSize: 12 });
  mpY += 14;

  mpY = addSectionHeading(mp, "Annotations -- HOOK", mpY, { color: C.PRIMARY });
  mpY = addBodyText(mp, "\"Would you send your worst enemy to a place where children dig holes in 35-degree heat...\" -- a rhetorical question. Pulls the reader in by asking them to imagine.", mpY, { fontSize: 10 });
  mpY += 6;

  mpY = addSectionHeading(mp, "Annotations -- POSITION", mpY, { color: C.SECONDARY });
  mpY = addBodyText(mp, "\"I strongly believe Camp Green Lake must be shut down.\" -- one clear sentence. No wobble. The reader knows which side this writer is on.", mpY, { fontSize: 10 });
  mpY += 6;

  mpY = addSectionHeading(mp, "Annotations -- SIGNPOST", mpY, { color: C.ACCENT });
  mpY = addBodyText(mp, "\"There are three reasons: the camp is unsafe, the work is inhumane, and -- worst of all -- it does not actually help the boys learn anything.\" -- previews the three body paragraphs that come next.", mpY, { fontSize: 10 });
  mpY += 6;

  mpY = addSectionHeading(mp, "Annotations -- bonus moves", mpY, { color: C.PRIMARY });
  mpY = addBodyText(mp, "Concrete detail in the hook (\"35-degree heat\", \"one canteen of water\") -- this makes the persuasive case feel real.", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "Emphasis phrase: \"-- worst of all --\" -- builds order from weakest to strongest reason.", mpY, { fontSize: 10 });
  mpY += 8;

  mpY = addSectionHeading(mp, "Sentence frames you can borrow", mpY, { color: C.PRIMARY });
  mpY = addBodyText(mp, "HOOK: \"Would you...?\" / \"Imagine if...\" / \"Did you know that...?\" / \"It is shocking that...\"", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "POSITION: \"I strongly believe [X] should be [Y].\" / \"[X] must be [Y].\"", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "SIGNPOST: \"There are three reasons: ___, ___, and ___.\" / \"This is true for three reasons.\"", mpY, { fontSize: 10 });

  addPdfFooter(mp, "Lesson 4 | Mentor Persuasive Introduction -- REFERENCE");

  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Holes_Lesson4.pptx` }),
    writePdf(plan, PLAN_PDF_PATH),
    writePdf(mp, MENTOR_PDF_PATH),
  ]);

  console.log("PPTX written to " + `${OUT_DIR}/Holes_Lesson4.pptx`);
  console.log("Done: " + PLAN_RESOURCE.name);
  console.log("Done: " + MENTOR_RESOURCE.name);
}

build().catch(console.error);
