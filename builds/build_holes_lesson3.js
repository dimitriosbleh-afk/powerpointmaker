"use strict";

// Holes Unit - Lesson 3 (Week 1, Session 3): Narrative Tension -- Short Story with Setting + Character + Trouble
// Year 5/6 Literacy
// Focus: build tension through pacing, short sentences and stakes
// Anchor: Holes by Louis Sachar, Chapter 28 (the yellow-spotted lizards) -- a classic tension scene

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

const SESSION_NUMBER = 3;
const FOOTER = "Narrative Tension | Lesson 3 | Week 1 | Year 5/6 Literacy | Holes";
const OUT_DIR = "output/Holes_Lesson3_Tension_Lizards";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const PLAN_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Tension Short Story Plan",
  "Student template: plan a short tension scene with setting, character and a clear trouble."
);
const MENTOR_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Mentor Tension Scene",
  "Annotated model tension scene -- shows short sentences, sensory detail and rising stakes."
);
const RESOURCE_ITEMS = [PLAN_RESOURCE, MENTOR_RESOURCE];
const PLAN_PDF_PATH = path.join(OUT_DIR, PLAN_RESOURCE.fileName);
const MENTOR_PDF_PATH = path.join(OUT_DIR, MENTOR_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Last lesson of Week 1 -- narrative writing
- We learned SETTING. We learned CHARACTER. Today we add TENSION
- Tension = the reader has to keep reading. Something is at stake

DO:
- Display title slide
- Read 2-3 of yesterday's strong showing sentences aloud (no names) to celebrate
- Have the novel ready for the Ch 28 read-aloud (the yellow-spotted lizards)

TEACHER NOTES:
This is the bridge lesson into Week 2 (persuasive) -- students finish Week 1 with one complete short narrative they can be proud of. The anchor today is the famous yellow-spotted lizard scene -- a model of how Sachar uses short sentences and small details to build dread.

WATCH FOR:
- Students excited to write a "scary" story -- redirect gently: "Not scary for the sake of scary. We want TENSION -- the reader leans forward"

[Literacy: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- Two resources today
- The ${PLAN_RESOURCE.name} -- your tension story plan and writing page
- The ${MENTOR_RESOURCE.name} -- the annotated tension scene

DO:
- Print the plan (one per student)
- Print the mentor (one per student or pair)
- Have the novel ready for the Ch 28 read-aloud
- Optional: keep yesterday's character drafts handy -- students may want to use the same character

[Literacy: Resources | VTLM 2.0: Student Resources]`;

const NOTES_HOOK = `SAY:
- I am going to read a SHORT extract from Holes -- Chapter 28
- This is the famous scene with the yellow-spotted lizards
- Listen for the SHORT sentences. Listen for how Sachar slows the moment down
- After: what made it feel tense?

DO:
- Read a short extract from Holes Chapter 28 -- the part where Stanley realises a lizard is on him, or watches them gather
- 1 to 2 pages is enough
- After reading, pause 5 seconds
- Cold call 2-3 students: "What made that feel TENSE?"

TEACHER NOTES:
The yellow-spotted lizard scene is iconic and accessible without context -- students will feel the tension even if they have not read the lead-up. Students will name: short sentences, the lizard is still, Stanley does not move, the heat, the quiet. Capture 2-3 strong responses on the board.

WATCH FOR:
- Students who name SHORT SENTENCES -- celebrate; this is the technique we teach today
- Students who name a SMALL detail (one yellow spot, a tongue flicker) -- celebrate

[Literacy: Text Launch | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_LI_SC = `SAY:
- Read the LI with me
- We are learning to build TENSION in a short narrative
- Three "I can" statements

DO:
- Choral read the LI and SCs
- Brief: "Tension is the reader leaning FORWARD. We make them want the next line"

TEACHER NOTES:
SC1 (foundation) -- one short sentence used to slow a moment. Achievable for all. SC2 (target) -- short sentence + a small detail + stakes. SC3 -- builds tension across the whole paragraph. Exit ticket targets SC2.

WATCH FOR:
- Students who think tension = action -- redirect: "Tension is BEFORE the action. It is the waiting"

[Literacy: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_VOCAB_DREAD = `SAY:
- One key word today: DREAD
- A noun -- a heavy, sick feeling that something bad is coming
- It is not fear of right now. It is fear of the next moment

DO:
- Choral say DREAD
- Quick comparison: "Fear = something scary is here. Dread = something scary is about to happen"
- Ask: "Think of a moment when you felt dread" -- 30 seconds partner share

TEACHER NOTES:
Dread is a precise word that helps students name what tension creates. Be careful with the partner share -- some students may share real fears; keep it light and move on if needed.

WATCH FOR:
- Students confusing dread with fear -- redirect: "Dread is the WAIT. Fear is the now"

[Literacy: Vocabulary | VTLM 2.0: Vocabulary Instruction]`;

const NOTES_IDO = `SAY:
- Watch me write a short tension scene
- The setup: a child alone in the house. They hear a sound downstairs
- I am going to use:
  - short sentences
  - a small detail
  - one moment of stillness
- I will NOT write the ending. Tension lives in the WAIT

DO:
- Display the I Do slide
- Read the model aloud SLOWLY, pausing after short sentences
- Point to the SHORT sentences with your finger as you read
- After reading: "Notice -- I stopped at the worst moment. That is tension"

TEACHER NOTES:
The child-alone-in-the-house is a classic tension setup. We deliberately end on the wait -- not the resolution -- so students learn that tension lives before the climax. They will write similar moments in You Do.

WATCH FOR:
- Students who notice the short sentences -- celebrate
- Students who ask "what happens next?" -- celebrate; that means the tension worked

[Literacy: I Do | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_WEDO = `SAY:
- Together. We plan a tension scene about: someone walking home through a quiet park at dusk
- Different setup from the I Do
- I will collect your ideas on the board

DO:
- Display the We Do slide
- Cold call for: SETTING detail, CHARACTER body action, SMALL DETAIL, what they FEAR
- Build a sequence of 3 SHORT sentences together
- Aim: 3 short, well-paced sentences on the board

TEACHER NOTES:
Quiet park at dusk is different from the I Do (house alone) and from the You Do (3 setups later). Keep the build to 5-6 minutes. Push students to USE PERIODS -- they will want to write long compound sentences.

WATCH FOR:
- Students who write a long sentence -- redirect: "Cut it. Where can a period go?"
- Students who give a precise small detail (a piece of paper on the swing, a single car door slam) -- celebrate

[Literacy: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_CFU_BUILD = `SAY:
- Quick check. Two openings to a tension scene in a school at night. Which one builds more tension?
- A: "Maya walked through the dark, empty school corridor. She was very scared because she could hear strange noises and she didn't know what they were and she just wanted to get out as fast as possible."
- B: "Maya stopped. The corridor was empty. A door clicked shut somewhere behind her."
- Show me A or B on your fingers

DO:
- Display both
- Show Me Fingers (1 for A, 2 for B)
- Scan: most students should choose B
- Cold call 1-2 students: "Why B?"

TEACHER NOTES:
A is one long, breathless sentence that tells us everything at once. B uses three short sentences that drip-feed information and slow the moment. B is the stronger tension opening. If many chose A, pivot to the re-teach slide.

WATCH FOR:
- Students who choose B and articulate "short sentences" -- ready for You Do
- Students who choose A because it has more words -- redirect: "More words = less tension. Cut, don't add"

[Literacy: CFU | VTLM 2.0: Formative Assessment]`;

const NOTES_CFU_REVEAL = `SAY:
- Stronger tension opening: B
- B is THREE short sentences. Each one is small. Together they build the dread
- A is one sentence trying to do everything at once. The tension leaks out

DO:
- Display the reveal banner
- Read B aloud SLOWLY -- pause between sentences
- Read A aloud at normal pace -- it sounds rushed

CFU CHECKPOINT:
Technique: Show Me Fingers (1 for A, 2 for B)
Script:
- "Hold up 1 for A, 2 for B. Which builds more tension?"
- Scan for: most students choose B (>=80%)
PROCEED (>=80%): Release to You Do.
PIVOT (<80%): Use the optional re-teach slide that follows -- it takes ONE long sentence and breaks it apart in front of the class.

WATCH FOR:
- Students who self-correct -- celebrate

[Literacy: CFU Reveal | VTLM 2.0: Formative Assessment]`;

const NOTES_RETEACH = `SAY:
- Second look at tension and short sentences
- Watch -- I will take a LONG, breathless sentence and BREAK it apart
- Long: "The footsteps were getting closer and closer and Stanley could feel his heart pounding and he didn't know whether to run or to hide and he kept looking around but he couldn't see anything in the dark."
- Where can I put a period? Watch me cut

DO:
- Display the re-teach slide
- Live-edit on the board: add periods, delete connectives
- Show the result:
  "The footsteps came closer. Stanley's heart pounded. Run or hide? He looked around. Nothing in the dark."
- Re-check: ask students to break ONE long sentence on their boards (give them one)

TEACHER NOTES:
OPTIONAL slide. Use only if CFU missed B. Different approach: instead of comparing two openings, transform one breathless sentence into a paced one. Show the cuts on the board so students see the surgery.

WATCH FOR:
- Students who use periods on their boards -- ready
- Students who still use "and" or "so" -- prompt: "Where could a period go?"

[Literacy: Optional Re-teach | VTLM 2.0: Reteach with Different Approach]`;

const NOTES_YOUDO = `SAY:
- Your turn. Write a SHORT tension scene -- 6 to 10 sentences
- Pick ONE setup:
  - A child alone in the house hears a sound downstairs
  - Someone gets lost in a busy market
  - A swimmer realises they have gone too far out
  - A child opens a door they were told never to open
  - Your own tension setup
- Use the Tension Plan
- 18 minutes. I will circulate

DO:
- Display the You Do slide
- Distribute the plan and mentor
- Circulate -- prioritise students who needed the re-teach
- Quick conferences: "Read me your first 3 sentences. Count the periods"

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the mentor scene. Look at the short sentences. Pick the FIRST setup (child in house, sound downstairs). Use the sentence frames on the back of the mentor sheet
- Extra Notes: These students still hit SC1 and SC2 -- the frame shows them how to pace

EXTENDING PROMPT:
- Task: After your tension scene, do NOT resolve it. Instead, write ONE final short sentence that opens up the dread further. ("The handle turned." / "The footsteps stopped.")
- Extra Notes: Tension lives in the WAIT. The strongest writers stop short

TEACHER NOTES:
The five setups all share the SKILL: short sentences, small detail, stakes. Active circulation is the formative assessment. Push for 6-10 sentences, not 20.

WATCH FOR:
- Students writing a long story -- redirect: "Just the tense MOMENT. Not before, not after"
- Students writing a single long sentence -- prompt: "Cut. Where can a period go?"
- Students who land a strong short final sentence -- celebrate publicly

[Literacy: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket
- On your plan sheet, write ONE tension MOMENT -- 3 short sentences
- The setup: a phone rings at 3am
- 3 short sentences. No explanation. No ending. Just the moment

DO:
- Display the exit ticket slide
- 2 minutes silent
- Collect on the way out
- Note SC2 achievement for evidence

TEACHER NOTES:
Exit ticket targets SC2. Three sentences = enough to demonstrate pacing.

WATCH FOR:
- Students who write one long sentence -- still collect; flag
- Students who use 3 short sentences -- evidence of SC2

[Literacy: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Last lesson of Week 1 finishing now
- Self-check: show on your fingers, 1 to 5, for each success criterion
- Then partner share: what is the BEST short sentence you wrote today?

DO:
- Run fingers check for each SC
- 60 seconds partner share -- give them a moment to enjoy what they wrote
- Briefly: "Next week is PERSUASIVE writing. Same novel, different job: convince a reader"

TEACHER NOTES:
End Week 1 well. Acknowledge that students have written THREE short narrative pieces -- a setting, a character and a tension scene -- and built the skills they will publish in Week 3.

WATCH FOR:
- Students who name their BEST short sentence -- celebrate
- Students who say "I don't know" -- prompt them to find one

[Literacy: Closing | VTLM 2.0: Review and Reflect / End of Week 1]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Holes -- Lesson 3 -- Narrative Tension";

  // SLIDE 1 -- Title
  titleSlide(
    pres,
    "Tension: Make the Reader Lean Forward",
    "The Yellow-Spotted Lizards -- Holes Ch. 28",
    "Lesson 3  |  Week 1  |  Year 5/6 Literacy",
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

  // SLIDE 3 -- Hook / Text Launch (Ch 28 read-aloud)
  quoteSlide(
    pres,
    "Read Aloud",
    "Holes -- Chapter 28 (the yellow-spotted lizards)",
    "Listen for the SHORT sentences. Listen for what Sachar slows down.",
    "Ch. 28, the lizard scene",
    "Partner talk: what made that feel TENSE? Name one specific technique you noticed.",
    NOTES_HOOK,
    FOOTER
  );

  // SLIDE 4 -- LI / SC
  liSlide(
    pres,
    [
      "We are learning to build TENSION in a short narrative using short sentences, small details and clear stakes",
    ],
    [
      "I can use at least one short sentence to slow a tense moment",
      "I can include a small specific detail that builds dread",
      "I can pace tension across the whole paragraph using short and long sentences",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // SLIDE 5 -- Vocab
  vocabSlide(
    pres,
    "dread",
    "noun",
    "A heavy, sick feeling that something bad is about to happen. Different from fear -- dread is the WAIT.",
    "A wave of dread settled over Stanley as the lizard's tongue flickered.",
    NOTES_VOCAB_DREAD,
    FOOTER
  );

  // SLIDE 6 -- I Do
  modellingSlide(
    pres,
    "I Do -- Watch Me",
    "Model Tension Scene: A Child Alone in the House",
    "Setup:\n\nA child alone in the house. They hear a sound downstairs.\n\nMy techniques:\n- short sentences\n- one small specific detail\n- a moment of stillness\n- stop on the WAIT (no ending)",
    "She heard the back door click.\n\nShe sat very still on the stairs. The hallway clock ticked. One. Two. Three.\n\nFootsteps. Not the cat's. Heavier. Slower. Closer to the bottom of the stairs.\n\nShe held her breath.\n\n(My choices:\n- short sentences: 4 words / 3 words / 2 words\n- small detail: clock ticking\n- stillness: \"sat very still\"\n- stop on the wait: \"held her breath\")",
    NOTES_IDO,
    FOOTER
  );

  // SLIDE 7 -- We Do
  {
    const s = pres.addSlide();
    addTopBar(s, C.SECONDARY);
    addBadge(s, "We Do", { color: C.SECONDARY });
    addTitle(s, "Together: Quiet Park at Dusk");

    const cardY = CONTENT_TOP;
    const cardH = 2.0;
    addCard(s, 0.5, cardY, 9, cardH, { strip: C.SECONDARY, fill: C.WHITE });
    s.addText("Plan together -- 3 short sentences", {
      x: 0.75, y: cardY + 0.10, w: 8.5, h: 0.30,
      fontSize: 14, fontFace: FONT_B, color: C.SECONDARY, bold: true, margin: 0,
    });
    s.addText("- SETTING:  one small specific detail (a single glove on a bench? a coat over a fence? a swing creaking?)\n- CHARACTER:  what they do with their body (slow steps? hand to pocket? glance back?)\n- SOUND:  one small sound that does not fit (a single car door slam? a branch breaking?)\n- STAKES:  what do they FEAR is happening? (do not name it -- imply it)", {
      x: 0.75, y: cardY + 0.45, w: 8.5, h: cardH - 0.55,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    const tipY = cardY + cardH + 0.18;
    const tipH = SAFE_BOTTOM - tipY;
    addCard(s, 0.5, tipY, 9, tipH, { strip: C.ACCENT, fill: C.BG_CARD });
    s.addText("Together on the board: 3 SHORT sentences", {
      x: 0.75, y: tipY + 0.10, w: 8.5, h: 0.30,
      fontSize: 13, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("- Each sentence under 12 words\n- Use a period -- not a comma\n- Try one as a class: \"The park was empty. A single glove lay on the bench. Maya stopped walking.\"", {
      x: 0.75, y: tipY + 0.42, w: 8.5, h: tipH - 0.55,
      fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_WEDO);
  }

  // SLIDE 8 + 9 -- CFU: Which opening builds tension? (reveal pair)
  function buildCfuBase() {
    const slide = pres.addSlide();
    addTopBar(slide, C.ALERT);
    addBadge(slide, "CFU", { color: C.ALERT });
    addTitle(slide, "Which Opening Builds More Tension?", { color: C.ALERT });

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
    slide.addText("\"Maya walked through the dark, empty school corridor. She was very scared because she could hear strange noises and she didn't know what they were and she just wanted to get out as fast as possible.\"", {
      x: 1.2, y: cardY + 0.10, w: 8.0, h: cardH - 0.20,
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
    slide.addText("\"Maya stopped. The corridor was empty. A door clicked shut somewhere behind her.\"", {
      x: 1.2, y: cardBY + 0.14, w: 8.0, h: cardBH - 0.24,
      fontSize: 16, fontFace: FONT_H, color: C.CHARCOAL, italic: true, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    addFooter(slide, FOOTER);
    slide.addNotes(NOTES_CFU_BUILD);
    return slide;
  }

  withReveal(
    buildCfuBase,
    (slide) => {
      const revealY = 4.90;
      addCard(slide, 0.5, revealY, 9, 0.42, { fill: C.SUCCESS });
      slide.addText("Stronger tension: B  --  three short sentences, one small detail, the worst is yet to come", {
        x: 0.5, y: revealY, w: 9, h: 0.42,
        fontSize: 13, fontFace: FONT_B, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      slide.addNotes(NOTES_CFU_REVEAL);
    }
  );

  // SLIDE 10 -- Optional Re-teach
  modellingSlide(
    pres,
    "Optional Re-teach",
    "Break the Breathless Sentence Apart",
    "Long, breathless:\n\n\"The footsteps were getting closer and closer and Stanley could feel his heart pounding and he didn't know whether to run or to hide and he kept looking around but he couldn't see anything in the dark.\"\n\nWhere can a PERIOD go?",
    "After cutting:\n\n\"The footsteps came closer. Stanley's heart pounded. Run or hide? He looked around. Nothing in the dark.\"\n\nFive short sentences. The reader leans forward.\n\nYour turn: break this sentence on your board --\n\"The lizard was sitting on the shovel handle and Stanley didn't dare move and his hand was so close he could see the yellow spots.\"",
    NOTES_RETEACH,
    FOOTER
  );

  // SLIDE 11 -- You Do
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "You Do", { color: C.PRIMARY, w: 1.5 });
    addTitle(s, "Write Your Own Short Tension Scene");

    addCard(s, 0.5, CONTENT_TOP, 9, 1.95, { strip: C.PRIMARY, fill: C.WHITE });
    s.addText("Pick ONE setup:", {
      x: 0.75, y: CONTENT_TOP + 0.10, w: 8.4, h: 0.30,
      fontSize: 14, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("- A child alone in the house hears a sound downstairs\n- Someone gets lost in a busy market\n- A swimmer realises they have gone too far out\n- A child opens a door they were told never to open\n- Your own tension setup", {
      x: 0.75, y: CONTENT_TOP + 0.45, w: 8.4, h: 1.40,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    const tipY = CONTENT_TOP + 2.10;
    const tipH = SAFE_BOTTOM - tipY;
    addCard(s, 0.5, tipY, 9, tipH, { strip: C.ACCENT, fill: C.BG_CARD });
    s.addText("Writing Time: 18 minutes  |  Target: 6 to 10 sentences", {
      x: 0.75, y: tipY + 0.10, w: 8.5, h: 0.30,
      fontSize: 13, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("First:    Use the Plan -- setting + character + trouble + small detail\nNext:    Write 6 to 10 sentences. Use AT LEAST 3 short sentences\nThen:    STOP on the wait. Do not resolve the moment\nFinally: Read it aloud quietly -- count your periods", {
      x: 0.75, y: tipY + 0.42, w: 8.4, h: tipH - 0.55,
      fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  }

  // SLIDE 12 -- Exit Ticket
  contentSlide(
    pres,
    "Exit Ticket",
    C.ACCENT,
    "Three Short Sentences -- the Phone Rings at 3am",
    [
      "The setup: a phone rings at 3am",
      "Write THREE short sentences -- no more, no less",
      "Stop on the wait. No ending",
      "2 minutes -- drop on my desk",
    ],
    NOTES_EXIT,
    FOOTER
  );

  // SLIDE 13 -- Closing (end of Week 1)
  closingSlide(
    pres,
    {
      reflectionPrompt: "Partner share: what is the BEST short sentence you wrote today?",
      scItems: [
        "I can use at least one short sentence to slow a tense moment",
        "I can include a small specific detail that builds dread",
        "I can pace tension across the whole paragraph using short and long sentences",
      ],
      selfAssessment: {
        prompt: "Show on your fingers: 1 (need help) to 5 (could teach it)",
        options: ["1-2", "3", "4-5"],
      },
    },
    NOTES_CLOSING
  );

  // ---- PDF: Tension Story Plan ----------------------------------------------
  const plan = createPdf({ title: PLAN_RESOURCE.name });
  let planY = addPdfHeader(plan, "Tension Scene Plan", {
    color: C.PRIMARY,
    subtitle: "Setting + Character + Trouble + Small Detail = Tension",
    lessonInfo: "Lesson 3 | Week 1 | Year 5/6 Literacy | Holes",
    showNameDate: true,
  });

  planY = addTipBox(plan, "Tension is the reader leaning forward. Today you write a SHORT scene (6 to 10 sentences) using short sentences, one small detail, and clear stakes. Plan first.", planY, { color: C.PRIMARY });

  planY = addSectionHeading(plan, "Step 1 -- Pick your setup", planY, { color: C.PRIMARY });
  planY = addBodyText(plan, "Circle ONE: child alone in house / lost in market / too far out at sea / door you mustn't open / your own:", planY, { fontSize: 11 });
  planY = addLinedArea(plan, planY, 1, { lineSpacing: 18 });
  planY += 4;

  planY = addSectionHeading(plan, "Step 2 -- Setting (one small detail)", planY, { color: C.PRIMARY });
  planY = addBodyText(plan, "What one small SPECIFIC thing tells us where we are? (a clock ticking, a glove on a bench, a single light on)", planY, { fontSize: 10, italic: true });
  planY = addLinedArea(plan, planY, 2, { lineSpacing: 18 });
  planY += 4;

  planY = addSectionHeading(plan, "Step 3 -- Character action", planY, { color: C.PRIMARY });
  planY = addBodyText(plan, "What is your character DOING with their body? (sitting still, walking slowly, holding breath)", planY, { fontSize: 10, italic: true });
  planY = addLinedArea(plan, planY, 2, { lineSpacing: 18 });
  planY += 4;

  planY = addSectionHeading(plan, "Step 4 -- The trouble (what is happening)", planY, { color: C.SECONDARY });
  planY = addBodyText(plan, "What is going wrong? (a sound? a person? a danger? something missing?)", planY, { fontSize: 10, italic: true });
  planY = addLinedArea(plan, planY, 2, { lineSpacing: 18 });

  addPdfFooter(plan, "Lesson 3 | Tension Plan -- Page 1");

  plan.addPage();
  let planY2 = addPdfHeader(plan, "Tension Scene -- Write Here", {
    color: C.PRIMARY,
    subtitle: "6 to 10 sentences. At least 3 short sentences. Stop on the wait.",
    lessonInfo: "Lesson 3 | Week 1 | Year 5/6 Literacy",
    showNameDate: false,
  });

  planY2 = addTipBox(plan, "Use your plan. Vary sentence length. Drop in your small detail early. STOP on the wait -- do not resolve. Read aloud quietly when you finish and count your periods.", planY2, { color: C.SECONDARY });

  planY2 = addLinedArea(plan, planY2, 14, { lineSpacing: 22 });

  addPdfFooter(plan, "Lesson 3 | Tension Scene -- Page 2");

  plan.addPage();
  let planY3 = addPdfHeader(plan, "Exit Ticket", {
    color: C.ACCENT,
    subtitle: "Three short sentences -- the phone rings at 3am",
    lessonInfo: "Lesson 3 | Week 1",
    showNameDate: false,
  });

  planY3 = addTipBox(plan, "Write THREE short sentences. The phone rings at 3am. Stop on the wait. Do not finish the moment.", planY3, { color: C.ACCENT });

  planY3 = addLinedArea(plan, planY3, 4, { lineSpacing: 22 });

  addPdfFooter(plan, "Lesson 3 | Exit Ticket");

  // ---- PDF: Mentor Tension Scene --------------------------------------------
  const mp = createPdf({ title: MENTOR_RESOURCE.name });
  let mpY = addPdfHeader(mp, "Mentor Tension Scene -- Annotated", {
    color: C.PRIMARY,
    subtitle: "Child alone in the house -- a model that pauses on the wait",
    lessonInfo: "Lesson 3 | Week 1 | Year 5/6 Literacy",
    showNameDate: false,
  });

  mpY = addTipBox(mp, "This model builds tension through SHORT SENTENCES, ONE small detail, STILLNESS and STOPPING on the wait. Use the PATTERN -- your scene is different.", mpY, { color: C.PRIMARY });

  mpY = addSectionHeading(mp, "Model scene", mpY, { color: C.PRIMARY });
  mpY = addBodyText(mp, "She heard the back door click. She sat very still on the stairs. The hallway clock ticked. One. Two. Three. Footsteps. Not the cat's. Heavier. Slower. Closer to the bottom of the stairs. She held her breath.", mpY, { fontSize: 12 });
  mpY += 12;

  mpY = addSectionHeading(mp, "Annotations -- short sentences (the engine of tension)", mpY, { color: C.PRIMARY });
  mpY = addBodyText(mp, "\"One. Two. Three.\" -- three sentences of ONE word each. Count the ticks with the character.", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "\"Footsteps. Not the cat's. Heavier. Slower.\" -- four short sentences in a row. Each one adds dread.", mpY, { fontSize: 10 });
  mpY += 6;

  mpY = addSectionHeading(mp, "Annotations -- small specific detail", mpY, { color: C.SECONDARY });
  mpY = addBodyText(mp, "\"The hallway clock ticked.\" -- a specific object (a clock, not just 'a sound') with a specific verb (ticked, not made a noise). The detail is more powerful than the noise.", mpY, { fontSize: 10 });
  mpY += 6;

  mpY = addSectionHeading(mp, "Annotations -- stillness", mpY, { color: C.ACCENT });
  mpY = addBodyText(mp, "\"She sat very still on the stairs.\" / \"She held her breath.\" -- the character is the calm centre while the trouble grows around her.", mpY, { fontSize: 10 });
  mpY += 6;

  mpY = addSectionHeading(mp, "Annotations -- stop on the wait", mpY, { color: C.ALERT });
  mpY = addBodyText(mp, "The scene ENDS on \"She held her breath\". The reader does not get to find out who is at the bottom of the stairs. That is the technique.", mpY, { fontSize: 10 });
  mpY += 8;

  mpY = addSectionHeading(mp, "Sentence frames you can borrow", mpY, { color: C.PRIMARY });
  mpY = addBodyText(mp, "[Character] heard [small sound].", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "[Character] [stillness action]. [One small detail in the place].", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "[One-word sentence]. [One-word sentence]. [One-word sentence].", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "[The thing that is wrong]. Not [the safe explanation]. [One more detail]. [Closer / louder / slower].", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "[Final short action -- holding breath, not moving, staring].", mpY, { fontSize: 10 });

  addPdfFooter(mp, "Lesson 3 | Mentor Tension Scene -- REFERENCE");

  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Holes_Lesson3.pptx` }),
    writePdf(plan, PLAN_PDF_PATH),
    writePdf(mp, MENTOR_PDF_PATH),
  ]);

  console.log("PPTX written to " + `${OUT_DIR}/Holes_Lesson3.pptx`);
  console.log("Done: " + PLAN_RESOURCE.name);
  console.log("Done: " + MENTOR_RESOURCE.name);
}

build().catch(console.error);
