"use strict";

// Holes Unit - Lesson 6 (Week 2, Session 3): Persuasive Punch -- Emotive Language + Rhetorical Devices
// Year 5/6 Literacy
// Focus: add emotive words and rhetorical devices (question, repetition, rule of three) to a PEEL paragraph
// Anchor: Holes by Louis Sachar -- camp life details for emotive examples

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

const SESSION_NUMBER = 6;
const FOOTER = "Emotive + Rhetorical | Lesson 6 | Week 2 | Year 5/6 Literacy | Holes";
const OUT_DIR = "output/Holes_Lesson6_Emotive_Rhetorical";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const PLAN_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Emotive and Rhetorical Upgrade",
  "Student template: upgrade a PEEL paragraph with emotive words and one rhetorical device."
);
const MENTOR_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Mentor Emotive Paragraph",
  "Annotated model showing emotive language and three rhetorical devices in action."
);
const RESOURCE_ITEMS = [PLAN_RESOURCE, MENTOR_RESOURCE];
const PLAN_PDF_PATH = path.join(OUT_DIR, PLAN_RESOURCE.fileName);
const MENTOR_PDF_PATH = path.join(OUT_DIR, MENTOR_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Welcome back. You have your PEEL paragraph from yesterday
- Today: add PUNCH. Emotive language and rhetorical devices
- The structure stays the same. We upgrade the WORDS

DO:
- Display title slide
- Read 2-3 strong evidence sentences from yesterday's exit ticket (no names)
- Have student PEEL paragraphs from yesterday on desks

TEACHER NOTES:
The PEEL structure stays. We add EMOTIVE words to evoke feeling, and ONE rhetorical device (question / repetition / rule of three). Three devices in one paragraph is too many -- pick ONE.

WATCH FOR:
- Students who think this lesson means rewriting from scratch -- reassure: "Upgrade, not rewrite"

[Literacy: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- Two resources today
- The ${PLAN_RESOURCE.name} -- upgrade your paragraph from yesterday
- The ${MENTOR_RESOURCE.name} -- annotated model showing emotive and rhetorical moves

DO:
- Print the plan (one per student)
- Print the mentor (one per student or pair)
- Have yesterday's PEEL paragraphs ready

[Literacy: Resources | VTLM 2.0: Student Resources]`;

const NOTES_HOOK = `SAY:
- Quick partner share
- 30 seconds: read ONE sentence from your PEEL paragraph aloud to your partner
- Partner job: name ONE word that could be more EMOTIVE

DO:
- 30 seconds reading + listening
- Cold call 2-3 partners: "What word did you flag?"
- Capture the upgrade words on the board (small -> tiny? bad -> cruel? hot -> burning?)
- Briefly: "Same meaning. STRONGER feeling. That is emotive language"

TEACHER NOTES:
This launch reuses yesterday's writing. It also gets students to listen for soft words. Use the upgrade words you collect throughout the lesson.

WATCH FOR:
- Students who flag "good" / "bad" / "hot" -- celebrate; these are easy upgrade targets
- Students who refuse to share -- pair them with a more confident partner

[Literacy: Hook | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_LI_SC = `SAY:
- Read the LI with me
- We are learning to add EMOTIVE language and ONE rhetorical device to a persuasive paragraph
- Three "I can" statements

DO:
- Choral read LI and SCs
- Brief: "Emotive = stir feeling. Rhetorical = a move the writer makes to persuade"

TEACHER NOTES:
SC1 -- replace one soft word with one emotive word. Achievable. SC2 (target) -- three emotive upgrades + one rhetorical device. SC3 -- the device is well chosen for the point.

WATCH FOR:
- Students who think emotive = exaggerated -- redirect: "Strong, not silly"

[Literacy: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_VOCAB_EMOTIVE = `SAY:
- One word today: EMOTIVE
- An adjective. It means causing strong feeling
- Emotive language uses words that make the reader FEEL something
- "Tired" is mild. "Exhausted" is emotive. "Broken" is even more emotive

DO:
- Choral say EMOTIVE
- Quick ladder on the board:
  - mild:  tired
  - emotive:  exhausted
  - emotive +:  broken
- Ask: "Give me an emotive word for 'angry'" -- 30 seconds partner share
- Take 3 examples (furious / livid / outraged)

TEACHER NOTES:
This vocabulary AND its examples become the toolkit for the lesson. Keep the ladder on the board.

WATCH FOR:
- Students using a thesaurus word they do not know -- prompt: "Read that out loud. Does it sound natural?"

[Literacy: Vocabulary | VTLM 2.0: Vocabulary Instruction]`;

const NOTES_DEVICES = `SAY:
- Three rhetorical devices we will use today
- ONE per paragraph -- not all three
- Look at the slide

DO:
- Display the rhetorical devices slide
- Choral read each row
- Brief example for each:
  - Rhetorical question: "Would you send your worst enemy to such a place?"
  - Repetition: "Day after day. Hole after hole. Boy after boy."
  - Rule of three: "It is unsafe, inhumane and unjustified."
- Ask: "Which device sounds best to YOU?" -- 30 seconds partner share

TEACHER NOTES:
Three devices keep the cognitive load manageable. Students pick ONE for their own writing. If a confident writer wants to use two, allow it -- but the YOU DO target is ONE.

WATCH FOR:
- Students who confuse rhetorical question with a real question -- redirect: "You are not waiting for an answer. You are making the reader think"

[Literacy: Devices | VTLM 2.0: Explicit Teaching / Strategies]`;

const NOTES_IDO = `SAY:
- Watch me UPGRADE yesterday's I Do paragraph
- Same point: Camp Green Lake is unsafe
- Same PEEL shape
- I will add THREE emotive upgrades and ONE rhetorical device

DO:
- Display the I Do slide
- Read yesterday's paragraph (left column) aloud
- Read TODAY's upgraded paragraph (right column) aloud
- Point to each upgrade in turn:
  - emotive word 1
  - emotive word 2
  - emotive word 3
  - rhetorical device

TEACHER NOTES:
The model shows the SAME paragraph, BEFORE and AFTER the upgrade. This makes the move explicit. Students will do the same to their own paragraph in You Do.

WATCH FOR:
- Students who notice the structure did not change -- celebrate
- Students who notice the emotive words pulling more weight -- celebrate

[Literacy: I Do | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_WEDO = `SAY:
- Together. Different paragraph
- Take this draft -- "Camp Green Lake is inhumane. The boys have very little water. This is not fair. The camp should close."
- We upgrade it together. Find 3 soft words. Then add ONE rhetorical device

DO:
- Display the We Do slide
- Cold call for upgrade candidates (very little -> meagre / one canteen)
- Capture upgrades on the board
- Cold call for ONE device (question / repetition / rule of three)
- Read the upgraded paragraph aloud

TEACHER NOTES:
Different content from I Do. Keep the build to 6-7 minutes. The students drive the upgrade choices.

WATCH FOR:
- Students upgrading "not fair" to "unjust" / "inhumane" / "shameful" -- celebrate
- Students who upgrade EVERY word -- redirect: "Three is enough. One device is enough"

[Literacy: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_CFU_BUILD = `SAY:
- Quick check. Two versions of the SAME sentence
- Which one uses emotive language and a rhetorical device better?

DO:
- Display both
- Show Me Fingers
- Cold call 1-2 students: "Why?"

TEACHER NOTES:
A uses one mild adjective and no device. B uses three emotive words AND the rule of three. B is the stronger upgrade.

WATCH FOR:
- Students who choose B and name the rule of three -- ready
- Students who choose A because it is "simpler" -- redirect: "Simple is fine for evidence. Punch is what makes persuasive WORK"

[Literacy: CFU | VTLM 2.0: Formative Assessment]`;

const NOTES_CFU_REVEAL = `SAY:
- Stronger sentence: B
- B has three emotive words AND uses the rule of three
- A has one mild adjective and no device

DO:
- Display the reveal banner
- Read B with feeling

CFU CHECKPOINT:
Technique: Show Me Fingers (1 for A, 2 for B)
Script:
- "Which uses emotive + rhetorical?"
- Scan for: most students choose B (>=80%)
PROCEED (>=80%): Release to You Do.
PIVOT (<80%): Use the optional re-teach slide that follows -- it walks through the upgrade move-by-move.

[Literacy: CFU Reveal | VTLM 2.0: Formative Assessment]`;

const NOTES_RETEACH = `SAY:
- Quick re-teach -- the UPGRADE move, step by step
- Start with a soft sentence: "The boys have it bad at Camp Green Lake."
- Step 1: pick THREE soft words. "Bad" is the obvious one
- Step 2: replace with emotive. "Bad" -> "brutal"
- Step 3: add a device. Rule of three: "...brutal, exhausting and unjust."
- Final: "The boys have it brutal, exhausting and unjust at Camp Green Lake."

DO:
- Display the re-teach slide
- Work live on the board
- Show each step
- Re-check: students take the soft sentence "Stanley is not happy at camp" and do the same upgrade on their boards

TEACHER NOTES:
OPTIONAL. Use only if CFU was below 80%. Different approach: rather than comparing two finished sentences, show the upgrade as a STEP-BY-STEP move.

WATCH FOR:
- Students who upgrade with the rule of three -- ready
- Students who write a question instead -- accept; question is also valid

[Literacy: Optional Re-teach | VTLM 2.0: Reteach with Different Approach]`;

const NOTES_YOUDO = `SAY:
- Your turn. Take yesterday's PEEL paragraph
- Goal: 3 emotive upgrades + 1 rhetorical device
- Cross out the soft word. Write the emotive word above it
- Decide on the device. Add it as a new sentence OR work it into an existing sentence
- 15 minutes. I will circulate

DO:
- Display the You Do slide
- Distribute the upgrade plan and mentor
- Have students mark yesterday's paragraph (cross-out + above)
- Circulate -- focus on students who chose ONE device clearly

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the emotive word bank on the back of the mentor sheet. Pick 3 words to upgrade. Pick ONE device using the cheat sheet
- Extra Notes: Students still hit SC1 and SC2 -- word bank lowers the load

EXTENDING PROMPT:
- Task: After the upgrade, write a brand-new SECOND PEEL paragraph using a different rhetorical device from today
- Extra Notes: This gets confident writers building two body paragraphs ready for tomorrow

TEACHER NOTES:
Most students will upgrade by writing the paragraph FRESH on the page 2 lined area of the new template. Some will mark up yesterday's draft. Both are valid -- the move matters more than the format.

WATCH FOR:
- Students who pick a thesaurus word they cannot pronounce -- redirect: "Pick a word you would actually say"
- Students who use ALL three devices -- redirect: "One is plenty"

[Literacy: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket
- Take this soft sentence: "The Warden is not nice to the boys."
- Upgrade it with TWO emotive words AND ONE rhetorical device
- 2 minutes

DO:
- 2 minutes silent
- Collect

TEACHER NOTES:
Exit ticket targets SC2. The Warden's cruelty is a clear emotive opportunity.

WATCH FOR:
- Students using strong emotive words (cruel / heartless / merciless) -- celebrate
- Students adding the rule of three (cruel, heartless and greedy) -- celebrate

[Literacy: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Self-check
- Show on fingers 1 to 5
- Partner share: which device did you use, and where?

DO:
- Run fingers check
- 30 seconds partner share
- Briefly: "Tomorrow: COUNTER-ARGUMENT and CONCLUSION -- the last persuasive lesson before publish"

TEACHER NOTES:
Sets up Lesson 7. Tomorrow is the final persuasive lesson of Week 2.

[Literacy: Closing | VTLM 2.0: Review and Reflect]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Holes -- Lesson 6 -- Emotive + Rhetorical";

  // SLIDE 1 -- Title
  titleSlide(
    pres,
    "Emotive Language + Rhetorical Devices",
    "Add Punch to Your Persuasive Writing",
    "Lesson 6  |  Week 2  |  Year 5/6 Literacy",
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

  // SLIDE 3 -- Launch
  contentSlide(
    pres,
    "Launch",
    C.PRIMARY,
    "Read Aloud and Listen for Soft Words",
    [
      "30 seconds: read ONE sentence from yesterday's PEEL paragraph aloud to your partner",
      "Partner job: name ONE word that could be MORE EMOTIVE",
      "I will collect upgrades on the board",
    ],
    NOTES_HOOK,
    FOOTER
  );

  // SLIDE 4 -- LI / SC
  liSlide(
    pres,
    [
      "We are learning to upgrade a persuasive paragraph using EMOTIVE language and ONE rhetorical device",
    ],
    [
      "I can replace one soft word with one emotive word",
      "I can make three emotive upgrades and add one rhetorical device",
      "I can choose a rhetorical device that fits my point",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // SLIDE 5 -- Vocab
  vocabSlide(
    pres,
    "emotive",
    "adjective",
    "Causing strong feeling. Emotive language uses words that make the reader FEEL something. 'Tired' is mild. 'Exhausted' is emotive. 'Broken' is even more emotive.",
    "The boys' faces were exhausted, blistered and broken.",
    NOTES_VOCAB_EMOTIVE,
    FOOTER
  );

  // SLIDE 6 -- Three rhetorical devices
  contentSlide(
    pres,
    "Devices",
    C.SECONDARY,
    "Three Rhetorical Devices -- Pick ONE",
    [
      "Rhetorical question:    a question the writer is not waiting to have answered  -  \"Would you send your worst enemy to such a place?\"",
      "Repetition:                       repeating a word or phrase to drive a point home  -  \"Day after day. Hole after hole. Boy after boy.\"",
      "Rule of three:                 a list of three for emphasis  -  \"It is unsafe, inhumane and unjustified.\"",
    ],
    NOTES_DEVICES,
    FOOTER
  );

  // SLIDE 7 -- I Do (BEFORE / AFTER)
  modellingSlide(
    pres,
    "I Do -- Watch Me",
    "Upgrade in Action: BEFORE -> AFTER",
    "BEFORE (yesterday's paragraph):\n\nFirstly, Camp Green Lake puts the boys in real danger every single day. In Holes, the boys dig holes five feet deep in 35-degree heat with deadly yellow-spotted lizards. This shows the camp is exhausting and risky. For this reason, Camp Green Lake should close.\n\n(Solid -- but a bit flat.)",
    "AFTER (today's upgrade):\n\nIs it acceptable to send children to a place that puts their lives at risk every single day? Firstly, Camp Green Lake exposes the boys to brutal, exhausting and deadly conditions. In Holes, the boys dig holes five feet deep in punishing 35-degree heat, surrounded by venomous yellow-spotted lizards. This proves the camp is not just hard -- it is dangerous. For this reason, Camp Green Lake must be shut down immediately.\n\nUpgrades:\n- emotive: brutal, exhausting, deadly, punishing, venomous, dangerous\n- device: rule of three (brutal, exhausting, deadly)\n- bonus: rhetorical question opening",
    NOTES_IDO,
    FOOTER
  );

  // SLIDE 8 -- We Do
  {
    const s = pres.addSlide();
    addTopBar(s, C.SECONDARY);
    addBadge(s, "We Do", { color: C.SECONDARY });
    addTitle(s, "Together: Upgrade This Soft Paragraph");

    const cardY = CONTENT_TOP;
    const cardH = 1.85;
    addCard(s, 0.5, cardY, 9, cardH, { strip: C.SECONDARY, fill: C.WHITE });
    s.addText("BEFORE:", {
      x: 0.75, y: cardY + 0.10, w: 8.5, h: 0.30,
      fontSize: 14, fontFace: FONT_B, color: C.SECONDARY, bold: true, margin: 0,
    });
    s.addText("\"Camp Green Lake is inhumane. The boys have very little water. This is not fair. The camp should close.\"", {
      x: 0.75, y: cardY + 0.42, w: 8.5, h: cardH - 0.55,
      fontSize: 14, fontFace: FONT_H, color: C.CHARCOAL, italic: true, margin: 0,
    });

    const tipY = cardY + cardH + 0.18;
    const tipH = SAFE_BOTTOM - tipY;
    addCard(s, 0.5, tipY, 9, tipH, { strip: C.ACCENT, fill: C.BG_CARD });
    s.addText("Together on the board:", {
      x: 0.75, y: tipY + 0.10, w: 8.5, h: 0.30,
      fontSize: 13, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("- Find 3 soft words to UPGRADE (very little -> meagre / not fair -> unjust)\n- Pick ONE rhetorical device (question? repetition? rule of three?)\n- Read the AFTER version aloud as a class", {
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
    addTitle(slide, "Which Sentence Uses Emotive + Rhetorical?", { color: C.ALERT });

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
    const cardH = 1.20;
    addCard(slide, 0.5, cardY, 9, cardH, { strip: C.MUTED, fill: C.WHITE });
    slide.addText("A", {
      x: 0.7, y: cardY + 0.10, w: 0.5, h: 0.4,
      fontSize: 22, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0,
    });
    slide.addText("\"The camp is bad for the boys.\"", {
      x: 1.2, y: cardY + 0.18, w: 8.0, h: cardH - 0.30,
      fontSize: 20, fontFace: FONT_H, color: C.CHARCOAL, italic: true, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    const cardBY = cardY + cardH + 0.18;
    const cardBH = SAFE_BOTTOM - cardBY - 0.55;
    addCard(slide, 0.5, cardBY, 9, cardBH, { strip: C.SECONDARY, fill: C.BG_CARD });
    slide.addText("B", {
      x: 0.7, y: cardBY + 0.10, w: 0.5, h: 0.4,
      fontSize: 22, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0,
    });
    slide.addText("\"The camp is brutal, exhausting and unjust -- a place built to break boys.\"", {
      x: 1.2, y: cardBY + 0.14, w: 8.0, h: cardBH - 0.24,
      fontSize: 18, fontFace: FONT_H, color: C.CHARCOAL, italic: true, valign: "middle", margin: 0,
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
      slide.addText("Stronger: B  --  three emotive words + rule of three + the closer \"built to break boys\"", {
        x: 0.5, y: revealY, w: 9, h: 0.42,
        fontSize: 13, fontFace: FONT_B, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      slide.addNotes(NOTES_CFU_REVEAL);
    }
  );

  // SLIDE 11 -- Optional Re-teach
  modellingSlide(
    pres,
    "Optional Re-teach",
    "The Upgrade Move -- Step by Step",
    "Soft sentence:\n\n\"The boys have it bad at Camp Green Lake.\"\n\nMy steps:\n- Step 1:  find soft words\n- Step 2:  replace with emotive\n- Step 3:  add ONE device",
    "Step 1 -- find soft words:\n\n\"The boys have it BAD at Camp Green Lake.\"\n\nStep 2 -- replace with emotive:\n\n\"BAD\" -> \"brutal\"\n\nStep 3 -- add a device. Try rule of three:\n\n\"brutal, exhausting and unjust\"\n\nFINAL:\n\n\"The boys have it brutal, exhausting and unjust at Camp Green Lake.\"\n\nYour turn -- upgrade on your board:\n\n\"Stanley is not happy at camp.\"",
    NOTES_RETEACH,
    FOOTER
  );

  // SLIDE 12 -- You Do
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "You Do", { color: C.PRIMARY, w: 1.5 });
    addTitle(s, "Upgrade Yesterday's PEEL Paragraph");

    addCard(s, 0.5, CONTENT_TOP, 9, 1.85, { strip: C.PRIMARY, fill: C.WHITE });
    s.addText("Your goal:", {
      x: 0.75, y: CONTENT_TOP + 0.10, w: 8.4, h: 0.30,
      fontSize: 14, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("- 3 emotive upgrades (cross out soft, write emotive above)\n- 1 rhetorical device (question, repetition or rule of three)\n- Same PEEL shape -- do not change the structure", {
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
    s.addText("First:    Mark up yesterday's paragraph -- circle 3 soft words\nNext:    Write emotive replacements ABOVE the line\nThen:    Pick ONE device. Add it to a new or rewritten sentence\nFinally: Rewrite the FULL upgraded paragraph neatly on page 2", {
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
    "Upgrade This Sentence",
    [
      "Soft: \"The Warden is not nice to the boys.\"",
      "Upgrade with TWO emotive words AND ONE rhetorical device",
      "Write the FULL upgraded sentence",
      "2 minutes -- drop on my desk",
    ],
    NOTES_EXIT,
    FOOTER
  );

  // SLIDE 14 -- Closing
  closingSlide(
    pres,
    {
      reflectionPrompt: "Partner share: which device did you use, and where?",
      scItems: [
        "I can replace one soft word with one emotive word",
        "I can make three emotive upgrades and add one rhetorical device",
        "I can choose a rhetorical device that fits my point",
      ],
      selfAssessment: {
        prompt: "Show on your fingers: 1 (need help) to 5 (could teach it)",
        options: ["1-2", "3", "4-5"],
      },
    },
    NOTES_CLOSING
  );

  // ---- PDF: Emotive and Rhetorical Upgrade Template -------------------------
  const plan = createPdf({ title: PLAN_RESOURCE.name });
  let planY = addPdfHeader(plan, "Emotive + Rhetorical Upgrade", {
    color: C.PRIMARY,
    subtitle: "Upgrade yesterday's PEEL paragraph -- 3 emotive words + 1 device",
    lessonInfo: "Lesson 6 | Week 2 | Year 5/6 Literacy | Holes",
    showNameDate: true,
  });

  planY = addTipBox(plan, "Same paragraph. Same shape. Stronger words. Pick ONE rhetorical device. Do not change the PEEL structure.", planY, { color: C.PRIMARY });

  planY = addSectionHeading(plan, "Step 1 -- Mark up yesterday's paragraph", planY, { color: C.PRIMARY });
  planY = addBodyText(plan, "Place yesterday's paragraph next to you. Circle 3 soft words. Write the emotive upgrade ABOVE the circled word.", planY, { fontSize: 10, italic: true });
  planY += 6;

  planY = addSectionHeading(plan, "Step 2 -- Choose your device", planY, { color: C.PRIMARY });
  planY = addBodyText(plan, "Tick ONE: rhetorical question  /  repetition  /  rule of three", planY, { fontSize: 11 });
  planY += 4;

  planY = addSectionHeading(plan, "Step 3 -- Draft the device", planY, { color: C.SECONDARY });
  planY = addBodyText(plan, "Write your device here before you put it in the paragraph.", planY, { fontSize: 10, italic: true });
  planY = addLinedArea(plan, planY, 3, { lineSpacing: 22 });
  planY += 6;

  planY = addSectionHeading(plan, "Step 4 -- Emotive word bank (steal from this list)", planY, { color: C.ACCENT });
  planY = addBodyText(plan, "Tired -> exhausted / drained / broken", planY, { fontSize: 10 });
  planY = addBodyText(plan, "Bad -> brutal / cruel / inhumane / unjust", planY, { fontSize: 10 });
  planY = addBodyText(plan, "Hot -> burning / scorching / blistering", planY, { fontSize: 10 });
  planY = addBodyText(plan, "Sad -> heartbroken / devastated / hopeless", planY, { fontSize: 10 });
  planY = addBodyText(plan, "Mean -> heartless / merciless / ruthless", planY, { fontSize: 10 });
  planY = addBodyText(plan, "Hard -> punishing / gruelling / relentless", planY, { fontSize: 10 });

  addPdfFooter(plan, "Lesson 6 | Upgrade Plan -- Page 1");

  plan.addPage();
  let planY2 = addPdfHeader(plan, "Upgraded PEEL Paragraph -- Write Here", {
    color: C.PRIMARY,
    subtitle: "Same shape. Stronger words. One rhetorical device.",
    lessonInfo: "Lesson 6 | Week 2 | Year 5/6 Literacy",
    showNameDate: false,
  });

  planY2 = addTipBox(plan, "Rewrite the FULL upgraded paragraph neatly. Use your emotive replacements. Drop in your device.", planY2, { color: C.SECONDARY });

  planY2 = addLinedArea(plan, planY2, 14, { lineSpacing: 22 });

  addPdfFooter(plan, "Lesson 6 | Upgraded Paragraph -- Page 2");

  plan.addPage();
  let planY3 = addPdfHeader(plan, "Exit Ticket", {
    color: C.ACCENT,
    subtitle: "Upgrade: The Warden is not nice to the boys",
    lessonInfo: "Lesson 6 | Week 2",
    showNameDate: false,
  });

  planY3 = addTipBox(plan, "Soft: \"The Warden is not nice to the boys.\" Upgrade with TWO emotive words AND ONE rhetorical device. Write the full upgraded sentence.", planY3, { color: C.ACCENT });

  planY3 = addLinedArea(plan, planY3, 4, { lineSpacing: 22 });

  addPdfFooter(plan, "Lesson 6 | Exit Ticket");

  // ---- PDF: Mentor Emotive Paragraph ----------------------------------------
  const mp = createPdf({ title: MENTOR_RESOURCE.name });
  let mpY = addPdfHeader(mp, "Mentor Emotive + Rhetorical -- Annotated", {
    color: C.PRIMARY,
    subtitle: "Camp Green Lake is unsafe -- the upgraded paragraph + word bank + device cheat sheet",
    lessonInfo: "Lesson 6 | Week 2 | Year 5/6 Literacy",
    showNameDate: false,
  });

  mpY = addTipBox(mp, "This model shows the SAME PEEL paragraph after an emotive + rhetorical upgrade. Use the PATTERN. Your paragraph is different.", mpY, { color: C.PRIMARY });

  mpY = addSectionHeading(mp, "Model upgraded paragraph", mpY, { color: C.PRIMARY });
  mpY = addBodyText(mp, "Is it acceptable to send children to a place that puts their lives at risk every single day? Firstly, Camp Green Lake exposes the boys to brutal, exhausting and deadly conditions. In Holes, the boys dig holes five feet deep in punishing 35-degree heat, surrounded by venomous yellow-spotted lizards. This proves the camp is not just hard -- it is dangerous. For this reason, Camp Green Lake must be shut down immediately.", mpY, { fontSize: 12 });
  mpY += 14;

  mpY = addSectionHeading(mp, "Annotations -- emotive upgrades", mpY, { color: C.PRIMARY });
  mpY = addBodyText(mp, "\"brutal, exhausting and deadly\" -- three emotive words in a row.", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "\"punishing 35-degree heat\" -- emotive adjective added to a fact.", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "\"venomous yellow-spotted lizards\" -- emotive adjective added to the evidence.", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "\"dangerous\" -- pulls the explanation tighter.", mpY, { fontSize: 10 });
  mpY += 8;

  mpY = addSectionHeading(mp, "Annotations -- rhetorical devices", mpY, { color: C.SECONDARY });
  mpY = addBodyText(mp, "Rhetorical question opener: \"Is it acceptable to send children to a place that puts their lives at risk every single day?\" -- the reader cannot say yes.", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "Rule of three (within the topic sentence): \"brutal, exhausting and deadly\" -- locks emotional weight into one phrase.", mpY, { fontSize: 10 });
  mpY += 8;

  mpY = addSectionHeading(mp, "Emotive word bank", mpY, { color: C.ACCENT });
  mpY = addBodyText(mp, "Tired -> exhausted, drained, broken, depleted", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "Bad -> brutal, cruel, inhumane, unjust, vicious", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "Hot -> burning, scorching, blistering, searing, punishing", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "Sad -> heartbroken, devastated, hopeless, crushed", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "Mean -> heartless, merciless, ruthless, callous", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "Hard -> punishing, gruelling, relentless, unbearable", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "Risky -> dangerous, deadly, life-threatening, perilous", mpY, { fontSize: 10 });
  mpY += 8;

  mpY = addSectionHeading(mp, "Device cheat sheet", mpY, { color: C.PRIMARY });
  mpY = addBodyText(mp, "Rhetorical question: ask the reader something they cannot answer with yes. Frame: \"Is it [acceptable/fair/reasonable] to ___?\"", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "Repetition: repeat a key word or short phrase. Frame: \"___ after ___. ___ after ___. ___ after ___.\"", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "Rule of three: list three emotive words. Frame: \"It is ___, ___ and ___.\"", mpY, { fontSize: 10 });

  addPdfFooter(mp, "Lesson 6 | Mentor Emotive + Rhetorical -- REFERENCE + WORD BANK");

  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Holes_Lesson6.pptx` }),
    writePdf(plan, PLAN_PDF_PATH),
    writePdf(mp, MENTOR_PDF_PATH),
  ]);

  console.log("PPTX written to " + `${OUT_DIR}/Holes_Lesson6.pptx`);
  console.log("Done: " + PLAN_RESOURCE.name);
  console.log("Done: " + MENTOR_RESOURCE.name);
}

build().catch(console.error);
