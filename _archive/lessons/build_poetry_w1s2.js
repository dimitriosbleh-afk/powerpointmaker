"use strict";

// Poetry Unit (Year 6) - Week 1 Session 2: Drafting a Three-Line Poem
// Form: Three-Line Poem (Line 1 subject / Line 2 location / Line 3 action).
// Catch-up design: opens with a 20-second recap of the form so a student who
// missed Session 1 can still draft today.
// Unit-wide look: literacy / grade56 / variant 3 ("Ink & Paper").

const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");

const { createTheme } = require("../themes/factory");
const T = createTheme("literacy", "grade56", 3);
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  addTopBar, addBadge, addTitle, addCard, addFooter,
  withReveal, addRevealAnswerBar,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  exitTicketSlide, modellingSlide, boardBuildSlide,
} = T;

const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addLinedArea,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const SESSION_NUMBER = 2;
const FOOTER = "Poetry | Week 1 Session 2 | Year 6 Literacy";
const OUT_DIR = "output/Poetry_W1S2_Draft_ThreeLine";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const PLAN_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Three-Line Poem Plan",
  "Planning template: subject, location, action boxes plus a word bank to draft from."
);
const RESOURCE_ITEMS = [PLAN_RESOURCE];
const PLAN_PDF_PATH = path.join(OUT_DIR, PLAN_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Today we stop being just readers and become poets. We write our own three-line poem.
- Do not worry about making it perfect. Poets always start with a draft.
- If you were away last session, that is fine -- I will remind you of the shape in a moment.

DO:
- Display the title slide.
- Set a calm, can-do tone. Everyone writes today.

TEACHER NOTES:
This is the first time students write their own poem this year. Keep the task small and the modelling generous.

WATCH FOR:
- Students who freeze on a blank page -- reassure them the plan template breaks it into three small steps.

[Literacy: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RECAP = `SAY:
- Quick reminder of the three-line poem, in twenty seconds.
- Line one is the subject -- what the poem is about, in a word or two.
- Line two is the location -- where the subject is.
- Line three is the action -- what the subject does.
- That is the whole recipe. Subject, where, what it does.

DO:
- Display the recap and read the example poem aloud.
- Point to each line and name its job once more.
- This is your catch-up moment for anyone who missed Session 1.

TEACHER NOTES:
The recap activates prior knowledge and lets absent students rejoin. Keep it brief -- the real work is the drafting that follows.

WATCH FOR:
- Students who already remember the shape -- invite them to whisper the three jobs to a partner.

[Literacy: Launch / Recap | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me. Today we WRITE a three-line poem using the shape.
- The first "I can" is for everyone -- write three lines that follow subject, location, action.

DO:
- Choral read the LI and success criteria.
- Stress: the goal today is a DRAFT, not a finished piece.

TEACHER NOTES:
SC1 (all): write three lines in the subject-location-action shape. SC2 (target): choose strong, precise words. SC3 (depth): add an image (a simile, metaphor or personification). Exit ticket targets SC1.

WATCH FOR:
- Students who try to write a long poem -- redirect: "Three lines only today."

[Literacy: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO = `SAY:
- Watch me write a three-line poem. I will think out loud so you can hear my choices.
- First my subject. I will choose an eagle, because I can picture it clearly. Line one: "The eagle".
- Now where is it? High above the mountain. Line two: "High above the mountain".
- Now what does it do? I could write "flies", but "soars" gives a stronger picture. Line three: "Soars".
- Notice I changed flies to soars. Choosing the strongest word is a poet's job.

DO:
- Display the model and read it aloud.
- Point to each line and name its job.
- Make the word-choice decision visible: write "flies" then cross it out and write "soars".

TEACHER NOTES:
The think-aloud models both the shape and word choice. Word choice is SC2, so make the "flies to soars" swap a clear teaching beat.

WATCH FOR:
- Students who copy the eagle poem -- redirect: "Your subject will be different. Pick something YOU can picture."

[Literacy: I Do | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_WEDO = `SAY:
- Now we build one together. Help me out.
- First, give me a subject -- something we can all picture. Call one out.
- Good. Now where is it? Give me the location for line two.
- And what does it do? Give me a strong action word for line three.
- Let us read our class poem back. Does each line do its job?

DO:
- Take suggestions and build the poem live in the blank space.
- Push for a strong action verb -- offer two options and let the class choose.
- Read the finished class poem aloud together.

TEACHER NOTES:
Guided practice. Channel the energy into the three jobs. If a student offers a whole sentence, help them trim it to a poem line.

WATCH FOR:
- Students who offer a weak verb (goes, is) -- prompt: "What is a stronger word for that?"

[Literacy: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_CFU = `SAY:
- Quick check. Here is a three-line poem that does NOT follow the shape. Something is wrong.
- Read it carefully. Which line is the problem, and what is missing?
- Tell your partner first, then we will fix it together.

DO:
- Display the broken poem.
- Partner talk for thirty seconds.
- Cold call: what is wrong? Then reveal the fix.

CFU CHECKPOINT:
Technique: Partner talk then cold call
Script:
- "Which line breaks the shape, and how do we fix it?"
- Scan for: students notice line three is a location, not an action.
PROCEED (>=80%): Release to You Do.
PIVOT (<80%): Use the optional re-teach slide that follows -- it rebuilds the shape one line at a time with a new subject.

TEACHER NOTES:
The broken poem has two location-type lines and no action. The fix is to give line three a real action word.

WATCH FOR:
- Students who cannot spot the problem -- point to line three and ask "is that what it DOES?"

[Literacy: CFU | VTLM 2.0: Formative Assessment]`;

const NOTES_CFU_REVEAL = `SAY:
- The problem was line three -- "near the cold water" is another location, not an action.
- A three-line poem needs an action in line three. We changed it to "waits".
- Now each line does its own job: subject, location, action.

DO:
- Display the fixed poem.
- Reread it so students hear the corrected shape.

[Literacy: CFU Reveal | VTLM 2.0: Formative Assessment]`;

const NOTES_RETEACH = `SAY:
- Let us rebuild the shape slowly with a brand new subject.
- Line one, the subject. I will pick a candle. Just "The candle".
- Line two, where is it? "On the windowsill".
- Line three, what does it do? A candle flickers. "Flickers".
- Three lines, three jobs. Subject, where, action. Now you try one on your whiteboard.

DO:
- Display the re-teach slide and build the candle poem line by line.
- Re-check: students write a three-line poem about a dog on their whiteboards.

TEACHER NOTES:
OPTIONAL. Use only if the CFU was below 80%. Different approach: a fresh subject built one line at a time with an immediate whiteboard re-check.

WATCH FOR:
- Students whose line three is still a location -- prompt: "What does the dog DO?"

[Literacy: Optional Re-teach | VTLM 2.0: Reteach with Different Approach]`;

const NOTES_YOUDO = `SAY:
- Your turn to write your own three-line poem.
- Use your plan: choose a subject you can picture, then its location, then a strong action.
- Read it back to yourself. Does each line do its job? Is line three your strongest word?
- Fifteen minutes. I will come around to hear them.

DO:
- Distribute the Three-Line Poem Plan.
- Circulate. Prioritise students who needed the re-teach.
- Quick conferences: "Read me your action line. Could the word be even stronger?"

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the word bank on the plan. Pick a subject, a location and an action from the lists, then write your three lines.
- Extra Notes: The word bank removes the blank-page block and still meets SC1.
EXTENDING PROMPT:
- Task: Add one image to a line -- turn the action into a simile or personification. Then write a SECOND three-line poem on a different subject.
- Extra Notes: Connects today's drafting back to the imagery skill from Session 1.

TEACHER NOTES:
Students keep their drafts -- they publish them in Session 4. Note any strong action verbs to celebrate.

WATCH FOR:
- Students who write full sentences -- help them trim to short poem lines.

[Literacy: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket. I will give you a subject and a location. You write line three -- the action.
- The subject is a storm. The location is over the ocean. Write a strong action line.
- One line. Make the word powerful. Two minutes.

DO:
- Two minutes, silent.
- Collect, or have students read theirs to a partner as they leave.

TEACHER NOTES:
Exit ticket targets SC1 -- writing a line that does its job. Strong verbs (rages, builds, crashes) show students understand the action line.

WATCH FOR:
- Students who write a location again -- flag for a quick reminder next session.

[Literacy: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Self-check. Show me on your fingers, one to five, for each "I can".
- Then share your favourite line with a partner and say why you chose that word.

DO:
- Run the fingers check for each success criterion.
- Sixty seconds partner share.
- Preview: "Next session we write a haiku, counting the beats."

TEACHER NOTES:
Acknowledge that everyone wrote a poem today. Note students who rated word choice low for follow-up.

[Literacy: Closing | VTLM 2.0: Review and Reflect]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Poetry Week 1 Session 2 -- Drafting a Three-Line Poem";

  // SLIDE 1 -- Title
  titleSlide(
    pres,
    "Write Your Own Three-Line Poem",
    "Subject, Location, Action -- one clear picture",
    "Week 1 Session 2  |  Year 6 Literacy  |  Poetry",
    NOTES_TITLE
  );

  // SLIDE 2 -- Resources
  addResourceSlide(pres, RESOURCE_ITEMS, T, FOOTER, NOTES_RESOURCES());

  // SLIDE 3 -- Recap / Launch (catch-up friendly)
  modellingSlide(
    pres,
    "Quick Recap",
    "The Three-Line Poem -- in 20 Seconds",
    "Three lines. Three jobs:\n\n- Line 1 = SUBJECT\n   (a word or two)\n\n- Line 2 = LOCATION\n   (where it is)\n\n- Line 3 = ACTION\n   (what it does)\n\nMissed last session? This is all you need.",
    "The tree\n\nIn the middle of the paddock\n\nStands",
    NOTES_RECAP,
    FOOTER
  );

  // SLIDE 4 -- LI / SC
  liSlide(
    pres,
    [
      "We are learning to WRITE a three-line poem using the shape: subject, location, action",
    ],
    [
      "I can write three lines that follow subject, location and action",
      "I can choose strong, precise words for each line",
      "I can add an image to one line (a simile, metaphor or personification)",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // SLIDE 5 -- I Do (teacher drafts a model, think-aloud)
  modellingSlide(
    pres,
    "I Do",
    "Watch Me Draft: Choosing Strong Words",
    "My thinking:\n\n1. Subject I can picture\n   -> an eagle\n\n2. Where is it?\n   -> high above the mountain\n\n3. What does it DO?\n   -> flies... no, SOARS\n   (the stronger word wins)",
    "The eagle\n\nHigh above the mountain\n\nSoars",
    NOTES_IDO,
    FOOTER
  );

  // SLIDE 6 -- We Do (build together)
  boardBuildSlide(
    pres,
    "We Do",
    "Build a Three-Line Poem Together",
    "Call out each line. Subject, then where, then a strong action.",
    NOTES_WEDO,
    FOOTER,
    {
      promptText: "Teacher writes the class poem here as students call out each line.",
      prefilledHints: ["Subject", "Where?", "Action"],
      badgeColor: C.SECONDARY,
    }
  );

  // SLIDE 7 + 8 -- CFU (fix the broken poem) with reveal
  withReveal(
    () => cfuSlide(
      pres,
      "CFU",
      "Fix the Shape",
      "Partner talk, then we fix it",
      "This poem breaks the shape. Which line is wrong?\n\nThe frog\nOn the green lily pad\nNear the cold water",
      NOTES_CFU,
      FOOTER
    ),
    (slide) => {
      addRevealAnswerBar(
        slide,
        "Line 3 was another location. Fixed: \"Waits\" -- now line 3 is an action.",
        { color: C.SUCCESS, label: "Fixed", showTickAndFix: false }
      );
      slide.addNotes(NOTES_CFU_REVEAL);
    }
  );

  // SLIDE 9 -- Optional Re-teach
  modellingSlide(
    pres,
    "Optional Re-teach",
    "Build It One Line at a Time",
    "New subject, built slowly:\n\n- Line 1 SUBJECT\n   -> The candle\n\n- Line 2 LOCATION\n   -> On the windowsill\n\n- Line 3 ACTION\n   -> Flickers\n\nNow you try one about a dog.",
    "The candle\n\nOn the windowsill\n\nFlickers",
    NOTES_RETEACH,
    FOOTER
  );

  // SLIDE 10 -- You Do
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "You Do", { color: C.PRIMARY, w: 1.5 });
    addTitle(s, "Write Your Three-Line Poem");

    addCard(s, 0.5, CONTENT_TOP, 9, 1.7, { strip: C.PRIMARY, fill: C.WHITE });
    s.addText("Use your plan:", {
      x: 0.75, y: CONTENT_TOP + 0.12, w: 8.4, h: 0.32,
      fontSize: 16, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("First:   Choose a subject you can picture (Line 1).\nNext:   Add where it is (Line 2), then a strong action (Line 3).\nThen:   Read it back -- is line 3 your strongest word?", {
      x: 0.75, y: CONTENT_TOP + 0.52, w: 8.4, h: 1.05,
      fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    const tipY = CONTENT_TOP + 1.85;
    const tipH = SAFE_BOTTOM - tipY;
    addCard(s, 0.5, tipY, 9, tipH, { strip: C.ACCENT, fill: C.BG_CARD });
    s.addText("Choices for everyone", {
      x: 0.75, y: tipY + 0.10, w: 8.4, h: 0.30,
      fontSize: 14, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("- Stuck? Use the word bank on your plan to pick a subject, location and action.\n- Ready for more? Turn one line into an image, then write a second poem on a new subject.", {
      x: 0.75, y: tipY + 0.44, w: 8.4, h: tipH - 0.55,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  }

  // SLIDE 11 -- Exit Ticket
  exitTicketSlide(
    pres,
    [
      "Subject: a storm.  Location: over the ocean.",
      "Write LINE 3 -- the action. Make the word powerful.",
    ],
    NOTES_EXIT,
    FOOTER,
    { assessesSc: 1, title: "Write the Action Line" }
  );

  // SLIDE 12 -- Closing
  closingSlide(
    pres,
    {
      reflectionPrompt: "Share your favourite line with a partner. Why did you choose that word?",
      scItems: [
        "I can write three lines that follow subject, location and action",
        "I can choose strong, precise words for each line",
        "I can add an image to one line (a simile, metaphor or personification)",
      ],
      selfAssessment: {
        prompt: "Show on your fingers: 1 (need help) to 5 (could teach it)",
        options: ["1-2", "3", "4-5"],
      },
    },
    NOTES_CLOSING
  );

  // ---- PDF: Three-Line Poem Plan -------------------------------------------
  const pl = createPdf({ title: PLAN_RESOURCE.name });
  let py = addPdfHeader(pl, "Three-Line Poem Plan", {
    color: C.PRIMARY,
    subtitle: "Subject, then where it is, then what it does.",
    lessonInfo: "Week 1 Session 2 | Year 6 Literacy | Poetry",
    showNameDate: true,
  });

  py = addTipBox(pl, "Plan first, then write. Choose a subject you can really picture in your mind.", py, { color: C.PRIMARY });

  py = addSectionHeading(pl, "Line 1 -- Subject (a word or two)", py, { color: C.PRIMARY });
  py = addLinedArea(pl, py, 1, { lineSpacing: 26 });
  py += 4;
  py = addSectionHeading(pl, "Line 2 -- Location (where is it?)", py, { color: C.SECONDARY });
  py = addLinedArea(pl, py, 1, { lineSpacing: 26 });
  py += 4;
  py = addSectionHeading(pl, "Line 3 -- Action (what does it do? choose a strong word)", py, { color: C.ACCENT });
  py = addLinedArea(pl, py, 1, { lineSpacing: 26 });
  py += 8;

  py = addSectionHeading(pl, "Word bank -- borrow if you are stuck", py, { color: C.PRIMARY });
  py = addBodyText(pl, "Subjects: the moon, the river, a wolf, the old gate, a candle, the storm, a seed, the train", py, { fontSize: 11 });
  py = addBodyText(pl, "Locations: in the dark forest, above the rooftops, beside the still pond, across the empty field", py, { fontSize: 11 });
  py = addBodyText(pl, "Strong actions: soars, crashes, waits, glows, races, whispers, trembles, drifts, towers, melts", py, { fontSize: 11 });

  addPdfFooter(pl, "Week 1 Session 2 | Three-Line Poem Plan -- Page 1");

  pl.addPage();
  let py2 = addPdfHeader(pl, "My Three-Line Poem -- Write It Here", {
    color: C.PRIMARY,
    subtitle: "Three lines. One clear picture.",
    lessonInfo: "Week 1 Session 2 | Year 6 Literacy",
    showNameDate: false,
  });
  py2 = addTipBox(pl, "Write your poem on the lines. Then read it aloud and check each line does its job.", py2, { color: C.SECONDARY });
  py2 = addLinedArea(pl, py2, 3, { lineSpacing: 30 });
  py2 += 10;
  py2 = addSectionHeading(pl, "Challenge -- write a second poem (new subject)", py2, { color: C.ACCENT });
  py2 = addLinedArea(pl, py2, 3, { lineSpacing: 30 });

  addPdfFooter(pl, "Week 1 Session 2 | My Three-Line Poem -- Page 2");

  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Poetry_W1S2.pptx` }),
    writePdf(pl, PLAN_PDF_PATH),
  ]);

  console.log("PPTX written to " + `${OUT_DIR}/Poetry_W1S2.pptx`);
  console.log("Done: " + PLAN_RESOURCE.name);
}

function NOTES_RESOURCES() {
  return `SAY:
- One sheet today: your Three-Line Poem Plan.
- It has a box for each line and a word bank if you get stuck.
- Keep your draft -- we publish it on Friday.

DO:
- Print the Three-Line Poem Plan, one per student.
- Have the Poets Toolbox from Session 1 nearby for reference.

[Literacy: Resources | VTLM 2.0: Enabling Learning]`;
}

build().catch((err) => {
  console.error(err);
  process.exit(1);
});
