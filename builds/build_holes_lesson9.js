"use strict";

// Holes Unit - Lesson 9 (Week 3, Session 2): Draft the Publish Piece
// Year 5/6 Literacy
// Focus: write a full first draft of the publish piece using yesterday's plan

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
  modellingSlide,
} = T;

const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addLinedArea,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const SESSION_NUMBER = 9;
const FOOTER = "Draft Publish Piece | Lesson 9 | Week 3 | Year 5/6 Literacy | Holes";
const OUT_DIR = "output/Holes_Lesson9_Draft";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const DRAFT_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Publish Piece Draft Booklet",
  "Student draft template -- one page per section, single-sided, plenty of space to write a full first draft."
);
const RESOURCE_ITEMS = [DRAFT_RESOURCE];
const DRAFT_PDF_PATH = path.join(OUT_DIR, DRAFT_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Today: DRAFT day
- Take your plan from yesterday
- Take your starting piece from Weeks 1 + 2
- Combine, polish, draft -- one full first draft of your publish piece

DO:
- Display title slide
- Have student plans + original pieces ready on desks
- Have draft booklets ready to distribute

TEACHER NOTES:
Today is heavy on student writing time. Aim for 35-40 minutes of writing in the lesson. The teaching is short and front-loaded -- one quick model, then students go.

WATCH FOR:
- Students who lost their plan -- have spare plans available
- Students who arrive without their starting piece -- offer to dig out from your photographs of earlier work, or have them start fresh using the plan

[Literacy: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- One resource today: the Publish Piece Draft Booklet
- A4 booklet with space for the full draft
- Use lined paper, single-sided, single column

DO:
- Distribute the draft booklet
- Have student plans + Week 1-2 work on desks
- Have the Success Criteria checklist (from Lesson 8) on desks as well

[Literacy: Resources | VTLM 2.0: Student Resources]`;

const NOTES_HOOK = `SAY:
- 60-second warm-up
- Open your plan from yesterday
- Read your ONE specific improvement for each section
- Whisper to yourself: "Today I will..."

DO:
- 60 seconds silent reading
- Brief: "You are not starting from scratch. You have a plan. You have a draft. Today: stitch them together."

TEACHER NOTES:
The warm-up gets students re-oriented quickly. Many will have forgotten exactly what they planned. Re-reading the plan locks intention before the pencil moves.

WATCH FOR:
- Students who cannot find their plan -- 30-second teacher conference to rebuild a quick plan from their existing piece

[Literacy: Hook | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_LI_SC = `SAY:
- Read the LI with me
- We are learning to draft our publish piece using yesterday's plan
- Three "I can" statements

DO:
- Choral read LI and SCs
- Brief: "First DRAFT. Not perfect. Tomorrow we EDIT"

TEACHER NOTES:
SC1 -- write at least 200 words. Achievable for all. SC2 (target) -- follow the planned structure section by section. SC3 -- incorporate at least 2 of the specific improvements planned yesterday.

WATCH FOR:
- Students who polish each sentence before moving on -- redirect: "Draft now. Polish tomorrow"

[Literacy: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO = `SAY:
- Quick model -- watch me start a draft
- I will read my plan, then write the FIRST paragraph live on the board
- Notice: I do not stop to polish. I write FAST

DO:
- Display the I Do slide
- Read the plan aloud
- Write the first paragraph LIVE on the board (or whiteboard, or visualiser)
- Talk through the choices as you write
- After the first paragraph: "I will keep going. So will you"

TEACHER NOTES:
This is a 3-4 minute model -- short. The point is to show pace. Students who watch carefully will start their own draft faster.

WATCH FOR:
- Students who notice the speed -- celebrate
- Students who try to copy your words -- redirect: "Your draft is different"

[Literacy: I Do | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_WEDO = `SAY:
- Together. One quick group decision
- For students drafting NARRATIVE: how will you OPEN your piece? With the setting? With the character? With action?
- For students drafting PERSUASIVE: how will you OPEN? Rhetorical question? Bold statement? Fact?
- 30 seconds partner share
- Then: STAND UP if you have decided your opener

DO:
- 30 seconds partner share
- "Stand up if you have a clear opener"
- Cold call 3-4 students -- name their opener

TEACHER NOTES:
This is a tiny We Do -- the lesson is for INDEPENDENT drafting. The mini-decision unsticks students who would otherwise stare at a blank page.

WATCH FOR:
- Students who name a clear opener -- ready to write
- Students still sitting -- 30-second teacher conference to pick an opener with them

[Literacy: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_CFU_BUILD = `SAY:
- Quick check before we draft. Two strategies for starting a draft fast. Which one will get you writing more today?

DO:
- Display both
- Show Me Fingers
- Cold call 1-2 students

TEACHER NOTES:
A is procrastination disguised as preparation. B is the strategy successful drafters use: skip the perfect opening if it stalls you, write the easiest section first, come back later. B is the right answer.

WATCH FOR:
- Students who pick A and explain "I need it to be right" -- redirect: "Right comes later. Words on the page comes now"

[Literacy: CFU | VTLM 2.0: Formative Assessment]`;

const NOTES_CFU_REVEAL = `SAY:
- Stronger drafting strategy: B
- If the opening stalls you, SKIP IT. Write your favourite section first. Come back for the opener
- Done is better than perfect

DO:
- Display the reveal banner
- Set a 35-minute writing block timer

CFU CHECKPOINT:
Technique: Show Me Fingers
Script:
- "Which strategy gets you writing more?"
- Scan for: most students choose B (>=80%)
PROCEED (>=80%): Release to drafting block.
PIVOT (<80%): Use the optional re-teach slide -- the start-anywhere strategy.

[Literacy: CFU Reveal | VTLM 2.0: Formative Assessment]`;

const NOTES_RETEACH = `SAY:
- Quick reminder -- you can start your draft ANYWHERE on the page
- Not the introduction
- Not the title
- The section you are MOST confident about
- Then go back for the rest

DO:
- Display the re-teach slide
- Walk through the example
- Re-check: ask students to point to the section of their plan they will start with TODAY

TEACHER NOTES:
OPTIONAL. Use only if CFU was under 80%. Different approach: explicit start-anywhere strategy with a worked example.

[Literacy: Optional Re-teach | VTLM 2.0: Reteach with Different Approach]`;

const NOTES_YOUDO = `SAY:
- Drafting block. 35 minutes
- Target: a FULL first draft -- around 250 words
- Use your plan
- Use your Week 1-2 piece as your starting point
- Do NOT stop to polish. We polish tomorrow

DO:
- Display the You Do slide
- Distribute draft booklets (if not already done)
- Start the timer (35 minutes)
- Circulate -- conference briefly with 6-10 students
- Half-time check at 15 minutes: "How many words so far? Show me on your fingers"

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Start by copying your strongest paragraph from Weeks 1-2 into the draft booklet. Then add the next section. Then the next
- Extra Notes: For students who freeze on the blank page, copying their existing strong work is a legitimate start. The polish happens tomorrow

EXTENDING PROMPT:
- Task: After your draft is complete, GO BACK to your weakest section and rewrite it once. (Cross out the first version. Write the second version below)
- Extra Notes: This is the SC3 lift -- self-revision in progress

TEACHER NOTES:
This is the largest writing block of the unit. Protect it fiercely. Use the half-time word count check to nudge anyone falling behind.

WATCH FOR:
- Students who stop after 5 sentences -- prompt: "Next section. Don't stop"
- Students who hit 300+ words -- celebrate; flag for editing
- Students rewriting the same sentence repeatedly -- prompt: "Draft now. Tomorrow we polish"

[Literacy: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket
- Read your draft to yourself quietly
- On the last page of the booklet, write ONE sentence: what is the BEST line in your draft, and ONE thing you want to fix tomorrow?

DO:
- 3 minutes silent reading + writing
- Collect

TEACHER NOTES:
This exit gives you a forward look for tomorrow's editing focus. Students who name a specific fix are ready for the edit lesson.

WATCH FOR:
- Students who name a specific best line + specific fix -- evidence of SC2 + SC3
- Students who write "all of it" or "nothing" -- prompt for specificity

[Literacy: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Self-check
- Show on fingers 1 to 5
- Partner share: read ONE favourite line from your draft

DO:
- Run fingers check
- 60 seconds partner share
- Briefly: "Tomorrow we EDIT and REVISE. Bring your draft. Bring fresh eyes"

TEACHER NOTES:
End on the favourite-line share -- a tiny celebration of progress. Tomorrow is editing.

[Literacy: Closing | VTLM 2.0: Review and Reflect]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Holes -- Lesson 9 -- Draft Publish Piece";

  // SLIDE 1 -- Title
  titleSlide(
    pres,
    "Draft Day",
    "Write Your Full First Draft -- 250 Words on One Page",
    "Lesson 9  |  Week 3  |  Year 5/6 Literacy",
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

  // SLIDE 3 -- Hook
  contentSlide(
    pres,
    "Launch",
    C.PRIMARY,
    "60-Second Plan Warm-Up",
    [
      "Open yesterday's plan",
      "Read your ONE specific improvement for each section",
      "Whisper to yourself: \"Today I will...\"",
    ],
    NOTES_HOOK,
    FOOTER
  );

  // SLIDE 4 -- LI / SC
  liSlide(
    pres,
    [
      "We are learning to draft our publish piece using yesterday's plan",
    ],
    [
      "I can write at least 200 words of my first draft",
      "I can follow my plan section by section",
      "I can include at least 2 of the specific improvements I planned",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // SLIDE 5 -- I Do
  modellingSlide(
    pres,
    "I Do -- Watch Me",
    "Start the Draft -- Read Plan, Then Write Fast",
    "My plan says:\n\n- Choice: NARRATIVE\n- Starting piece: my Lesson 1 setting\n- First specific improvement: replace 3 weak verbs in the setting paragraph\n- Layout: title top, setting / character / tension stacked",
    "I read my plan once. Then I WRITE. Not perfect. Fast.\n\nFirst paragraph (live):\n\n\"There is no lake at Camp Green Lake -- there has not been one for over a hundred years. The ground is cracked and pale, and the heat shimmers above it like a second sky. A faded sign creaks in the wind. Nothing else moves. The place feels barren -- like it has forgotten how to breathe.\"\n\nI did not polish. I kept going. So will you.",
    NOTES_IDO,
    FOOTER
  );

  // SLIDE 6 -- We Do (decide opener)
  contentSlide(
    pres,
    "We Do",
    C.SECONDARY,
    "Decide Your Opener -- Then Stand Up",
    [
      "NARRATIVE: setting? character action? a line of dialogue?",
      "PERSUASIVE: rhetorical question? bold statement? surprising fact?",
      "30 seconds partner share -- name your opener",
      "STAND UP if you have decided",
    ],
    NOTES_WEDO,
    FOOTER
  );

  // SLIDE 7 + 8 -- CFU
  function buildCfuBase() {
    const slide = pres.addSlide();
    addTopBar(slide, C.ALERT);
    addBadge(slide, "CFU", { color: C.ALERT });
    addTitle(slide, "Which Drafting Strategy Gets You Writing More?", { color: C.ALERT });

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
    const cardH = 1.4;
    addCard(slide, 0.5, cardY, 9, cardH, { strip: C.MUTED, fill: C.WHITE });
    slide.addText("A", {
      x: 0.7, y: cardY + 0.10, w: 0.5, h: 0.4,
      fontSize: 22, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0,
    });
    slide.addText("\"I need the perfect opening sentence before I start anything else. I'll spend 10 minutes rewriting it until it is right. Then I'll start paragraph 2.\"", {
      x: 1.2, y: cardY + 0.12, w: 8.0, h: cardH - 0.24,
      fontSize: 14, fontFace: FONT_H, color: C.CHARCOAL, italic: true, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    const cardBY = cardY + cardH + 0.18;
    const cardBH = SAFE_BOTTOM - cardBY - 0.55;
    addCard(slide, 0.5, cardBY, 9, cardBH, { strip: C.SECONDARY, fill: C.BG_CARD });
    slide.addText("B", {
      x: 0.7, y: cardBY + 0.10, w: 0.5, h: 0.4,
      fontSize: 22, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0,
    });
    slide.addText("\"If my opening stalls me, I'll skip it. I'll write the easiest section first. I can come back for the opener later. The point is to keep moving.\"", {
      x: 1.2, y: cardBY + 0.12, w: 8.0, h: cardBH - 0.24,
      fontSize: 14, fontFace: FONT_H, color: C.CHARCOAL, italic: true, valign: "middle", margin: 0,
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
      slide.addText("Stronger strategy: B  --  start anywhere, keep moving, polish tomorrow", {
        x: 0.5, y: revealY, w: 9, h: 0.42,
        fontSize: 14, fontFace: FONT_B, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      slide.addNotes(NOTES_CFU_REVEAL);
    }
  );

  // SLIDE 9 -- Optional Re-teach
  modellingSlide(
    pres,
    "Optional Re-teach",
    "You Can Start ANYWHERE in Your Draft",
    "Stuck on the opening?\n\nThat is normal. Skip it.\n\nWhich section of your plan are you most CONFIDENT about?\n\nWrite that first.",
    "Example -- a narrative writer who is stuck on the setting:\n\n1. Skip the setting for now\n2. Write the tension scene (their favourite from Lesson 3)\n3. Write the character (using last week's piece)\n4. Come back for the setting LAST\n\nThe reader sees the order on the page.\nThe writer does not have to write in order.\n\nPoint to your plan -- which section will YOU start with?",
    NOTES_RETEACH,
    FOOTER
  );

  // SLIDE 10 -- You Do (long writing block)
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "You Do", { color: C.PRIMARY, w: 1.5 });
    addTitle(s, "Drafting Block -- 35 Minutes");

    addCard(s, 0.5, CONTENT_TOP, 9, 1.95, { strip: C.PRIMARY, fill: C.WHITE });
    s.addText("Target: a FULL first draft -- around 250 words", {
      x: 0.75, y: CONTENT_TOP + 0.10, w: 8.4, h: 0.30,
      fontSize: 14, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("- Use your plan from yesterday\n- Use your Week 1-2 starting piece\n- Do NOT stop to polish -- we polish tomorrow\n- Skip the opener if it stalls you. Come back for it last", {
      x: 0.75, y: CONTENT_TOP + 0.42, w: 8.4, h: 1.35,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    const tipY = CONTENT_TOP + 2.10;
    const tipH = SAFE_BOTTOM - tipY;
    addCard(s, 0.5, tipY, 9, tipH, { strip: C.ACCENT, fill: C.BG_CARD });
    s.addText("Half-time check: at 15 minutes, show me your word count on your fingers", {
      x: 0.75, y: tipY + 0.10, w: 8.5, h: 0.30,
      fontSize: 13, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("Stuck? -- look at the optional re-teach slide.   /   Already finished? -- rewrite your weakest section once.", {
      x: 0.75, y: tipY + 0.42, w: 8.4, h: tipH - 0.55,
      fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  }

  // SLIDE 11 -- Exit Ticket
  contentSlide(
    pres,
    "Exit Ticket",
    C.ACCENT,
    "Best Line + One Fix for Tomorrow",
    [
      "Read your draft to yourself quietly",
      "On the last page of the booklet, write ONE sentence:",
      "What is the BEST line in your draft -- and ONE thing you want to FIX tomorrow?",
      "3 minutes",
    ],
    NOTES_EXIT,
    FOOTER
  );

  // SLIDE 12 -- Closing
  closingSlide(
    pres,
    {
      reflectionPrompt: "Partner share: read ONE favourite line from your draft.",
      scItems: [
        "I can write at least 200 words of my first draft",
        "I can follow my plan section by section",
        "I can include at least 2 of the specific improvements I planned",
      ],
      selfAssessment: {
        prompt: "Show on your fingers: 1 (need help) to 5 (could teach it)",
        options: ["1-2", "3", "4-5"],
      },
    },
    NOTES_CLOSING
  );

  // ---- PDF: Draft Booklet ---------------------------------------------------
  const draft = createPdf({ title: DRAFT_RESOURCE.name });
  let dY = addPdfHeader(draft, "Publish Piece -- First Draft", {
    color: C.PRIMARY,
    subtitle: "Around 250 words. Single-sided. Draft fast -- polish tomorrow.",
    lessonInfo: "Lesson 9 | Week 3 | Year 5/6 Literacy | Holes",
    showNameDate: true,
  });

  dY = addTipBox(draft, "Use this booklet for your full first draft. Write neatly enough to read tomorrow, but don't polish. We edit and revise tomorrow.", dY, { color: C.PRIMARY });

  dY = addSectionHeading(draft, "Genre and title", dY, { color: C.PRIMARY });
  dY = addBodyText(draft, "Circle ONE:  NARRATIVE  /  PERSUASIVE", dY, { fontSize: 11 });
  dY = addBodyText(draft, "Title:", dY, { fontSize: 11 });
  dY = addLinedArea(draft, dY, 1, { lineSpacing: 24 });
  dY += 6;

  dY = addSectionHeading(draft, "Draft -- write your full piece below", dY, { color: C.PRIMARY });
  dY = addLinedArea(draft, dY, 14, { lineSpacing: 22 });

  addPdfFooter(draft, "Lesson 9 | First Draft -- Page 1");

  draft.addPage();
  let dY2 = addPdfHeader(draft, "First Draft -- continued", {
    color: C.PRIMARY,
    subtitle: "Around 250 words target",
    lessonInfo: "Lesson 9 | Week 3 | Year 5/6 Literacy",
    showNameDate: false,
  });

  dY2 = addLinedArea(draft, dY2, 22, { lineSpacing: 22 });

  addPdfFooter(draft, "Lesson 9 | First Draft -- Page 2");

  draft.addPage();
  let dY3 = addPdfHeader(draft, "Exit Ticket -- Best Line + One Fix", {
    color: C.ACCENT,
    subtitle: "What is the best line in your draft, and one thing to fix tomorrow?",
    lessonInfo: "Lesson 9 | Week 3",
    showNameDate: false,
  });

  dY3 = addTipBox(draft, "Read your draft quietly. Write one sentence below.", dY3, { color: C.ACCENT });

  dY3 = addLinedArea(draft, dY3, 4, { lineSpacing: 22 });

  addPdfFooter(draft, "Lesson 9 | Exit Ticket");

  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Holes_Lesson9.pptx` }),
    writePdf(draft, DRAFT_PDF_PATH),
  ]);

  console.log("PPTX written to " + `${OUT_DIR}/Holes_Lesson9.pptx`);
  console.log("Done: " + DRAFT_RESOURCE.name);
}

build().catch(console.error);
