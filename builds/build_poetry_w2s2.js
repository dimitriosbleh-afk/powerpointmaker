"use strict";

// Poetry Unit (Year 6) - Week 2 Session 2: Colour Poem -- Read, Analyse and Draft
// Form: sensory colour poem ("Colour is the [sense] of [a specific thing]" x5 senses).
// Combined read + draft session. Supplied "Orange" lines used exactly as excerpts.
// Curriculum spine: sensory language and vivid imagery (VC2E5LE04 / VCELT355).
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
  exitTicketSlide, modellingSlide, boardBuildSlide, quoteSlide,
} = T;

const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addLinedArea,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const SESSION_NUMBER = 2;
const FOOTER = "Poetry | Week 2 Session 2 | Year 6 Literacy";
const OUT_DIR = "output/Poetry_W2S2_ColourPoem";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const PLAN_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Colour Poem Plan",
  "Colour poem planning template -- one line per sense -- with sentence frames and sense word banks."
);
const RESOURCE_ITEMS = [PLAN_RESOURCE];
const PLAN_PDF_PATH = path.join(OUT_DIR, PLAN_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Today's poem shape is a colour poem. We take one colour and describe it with all five senses.
- It is a chance to be really imaginative -- what does a colour taste like? sound like?
- We will read part of one, see how it works, then write our own.

DO:
- Display the title slide.
- Have the example colour poem lines ready to read.

TEACHER NOTES:
This session both analyses and drafts. The heart of it is sensory imagery -- sights, sounds, smells, tastes and feelings.

WATCH FOR:
- Students who think a colour can only be "seen" -- that is the fun: stretch it to all five senses.

[Literacy: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- One sheet today: your Colour Poem Plan. It has a line for each sense and some sentence starters.
- There is a word bank of sense words to help you.
- Keep your draft -- we publish it next session.

DO:
- Print the Colour Poem Plan, one per student.
- Have the Poets Toolbox from Week 1 nearby.

[Literacy: Resources | VTLM 2.0: Enabling Learning]`;

const NOTES_LAUNCH = `SAY:
- Listen to part of a poem all about the colour orange.
- "Orange is the sun after a summer day. Orange is the taste of a pizza that just came out of the oven. Orange is the feeling inside you when you accomplish something."
- Notice -- it is not just what orange looks like. It is what it tastes like, and even what it feels like inside.
- That is a colour poem: one colour, all five senses.

DO:
- Read the excerpt aloud, with warmth.
- Quick partner share: "Which line surprised you? Which sense was it using?"
- Cold call two students.

TEACHER NOTES:
These are exact lines from the supplied Orange poem, used as a short excerpt. Focus students on the senses each line uses.

WATCH FOR:
- Students who enjoy the surprising lines (taste, feeling) -- celebrate, those are the strongest images.

[Literacy: Launch | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me. Today we read a colour poem, then write our own using the five senses.
- The first "I can" is for everyone -- write a line about how a colour looks.

DO:
- Choral read the LI and success criteria.
- List the five senses on the board: see, hear, smell, taste, feel.

TEACHER NOTES:
SC1 (all): write sense lines for a colour. SC2 (target): use specific, vivid images (not "nice" or "good"). SC3 (depth): include a simile or metaphor in a line. Exit ticket targets SC2.

WATCH FOR:
- Students who write "blue is good" -- that has no picture; push for a specific image.

[Literacy: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO = `SAY:
- Watch me write a colour poem about green. I will use a different sense for each line.
- Sound: "Green is the sound of leaves in the wind." Can you hear it?
- Smell: "Green is the smell of fresh-cut grass."
- Taste: "Green is the taste of a crisp apple."
- Sight and feel to finish. Notice I chose SPECIFIC things -- not just "green is nice", but a real picture.

DO:
- Display the model and read it aloud.
- Point to the sense each line uses.
- Highlight the difference between "nice" and a specific image like "fresh-cut grass".

TEACHER NOTES:
The teaching point is specificity. Vague words make no picture; specific sensory details do. Model the choice openly.

WATCH FOR:
- Students who copy green -- redirect: "Your colour will be different. Pick a colour you have a strong feeling about."

[Literacy: I Do | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_WEDO = `SAY:
- Let us build a colour poem together. Our colour is red.
- One line for each sense. What does red SOUND like? SMELL like? TASTE like?
- Push past the first idea to a specific picture. Not "red is loud" -- red sounds like what, exactly?

DO:
- Build the poem live, one sense per line.
- For each line, ask "what specific thing?" to sharpen the image.
- Read the finished class poem aloud.

TEACHER NOTES:
Guided practice in sensory imagery. The work is turning a vague idea into a specific picture.

WATCH FOR:
- Students stuck on taste or sound for a colour -- offer two options and let the class choose.

[Literacy: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_CFU = `SAY:
- Quick check. Both lines are about the colour yellow. Which one paints a stronger picture?
- Read them carefully. Which makes you actually SEE or SMELL something?
- Show me on your fingers: one for A, two for B.

DO:
- Display both lines.
- Show Me Fingers.
- Scan: most should choose B.

CFU CHECKPOINT:
Technique: Show Me Fingers (1 for A, 2 for B)
Script:
- "Which line paints a stronger picture?"
- Scan for: students choose B and can say which sense it uses.
PROCEED (>=80%): Release to You Do.
PIVOT (<80%): Use the optional re-teach slide that follows -- it turns three vague lines into specific sensory images.

TEACHER NOTES:
A is vague ("nice") with no picture. B uses smell and a specific image (warm summer sand). B is far stronger.

WATCH FOR:
- Students who pick A -- ask "what do you picture when you read A?" The answer is nothing.

[Literacy: CFU | VTLM 2.0: Formative Assessment]`;

const NOTES_CFU_REVEAL = `SAY:
- B is the stronger image: "Yellow smells like warm summer sand."
- It uses a sense (smell) and a specific picture. A just says "nice" -- no picture at all.

DO:
- Display the reveal.
- Ask students which sense B uses (smell).

[Literacy: CFU Reveal | VTLM 2.0: Formative Assessment]`;

const NOTES_RETEACH = `SAY:
- Let us turn vague lines into pictures. The trick is to add a SPECIFIC thing.
- "Blue is calm." -- no picture. Better: "Blue is the calm of still water at dawn."
- "Blue is cold." -- okay. Better: "Blue tastes like ice on a winter morning."
- See how adding a specific thing makes a picture appear? That is your job.

DO:
- Display the re-teach slide and improve each line together.
- Re-check: students improve "Red is hot" into a specific sensory line on whiteboards.

TEACHER NOTES:
OPTIONAL. Use only if the CFU was below 80%. Different approach: a before-and-after that shows exactly how to add specific detail.

WATCH FOR:
- Students whose new line is still vague -- ask "a specific thing -- what exactly?"

[Literacy: Optional Re-teach | VTLM 2.0: Reteach with Different Approach]`;

const NOTES_YOUDO = `SAY:
- Your turn. Choose a colour you have a strong feeling about.
- Write one line for each sense: looks, sounds, smells, tastes, feels.
- For each line, picture a SPECIFIC thing. Make me see it, hear it, smell it.
- Fifteen minutes. I will come around.

DO:
- Distribute the Colour Poem Plan.
- Circulate. Prioritise students who needed the re-teach.
- Conference: "Read me your smell line. What specific thing do you picture?"

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the sentence frames -- "[Colour] looks like ___, sounds like ___" -- and the sense word bank.
- Extra Notes: Frames give every student a way in and still meet SC1.
EXTENDING PROMPT:
- Task: Turn one line into a simile or metaphor, and write a longer colour poem with extra lines, like the Orange poem.
- Extra Notes: Connects sensory writing to the imagery tools from Week 1.

TEACHER NOTES:
Students keep their colour poem draft to publish next session. Note vivid, surprising lines to share.

WATCH FOR:
- Students who only use sight -- nudge them to try taste or sound for at least one line.

[Literacy: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket. The colour is blue.
- Write ONE line that uses smell OR sound. Make it a specific picture.
- For example: blue smells like rain on hot footpaths. One line. Two minutes.

DO:
- Two minutes, silent.
- Collect.

TEACHER NOTES:
Exit ticket targets SC2 -- a specific, vivid sensory image. A concrete picture (not "nice") shows the skill.

WATCH FOR:
- Vague lines with no picture -- a sign to revisit specificity.

[Literacy: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Self-check. Show me on your fingers, one to five, for each "I can".
- Then read your favourite line to a partner and say which sense it uses.

DO:
- Run the fingers check for each success criterion.
- Sixty seconds partner share.
- Preview: "Next session we publish both poems -- our cinquain and our colour poem."

TEACHER NOTES:
Note students whose images are still vague so you can support them in publishing.

[Literacy: Closing | VTLM 2.0: Review and Reflect]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Poetry Week 2 Session 2 -- Colour Poem";

  // SLIDE 1 -- Title
  titleSlide(
    pres,
    "The Colour Poem",
    "One colour, all five senses -- read it, then write it",
    "Week 2 Session 2  |  Year 6 Literacy  |  Poetry",
    NOTES_TITLE
  );

  // SLIDE 2 -- Resources
  addResourceSlide(pres, RESOURCE_ITEMS, T, FOOTER, NOTES_RESOURCES);

  // SLIDE 3 -- Launch (read the supplied Orange excerpt)
  quoteSlide(
    pres,
    "Read Aloud",
    "A Colour Poem -- Orange",
    "Orange is the sun after a summer day.\nOrange is the taste of a pizza that just came out of the oven.\nOrange is the feeling inside you when you accomplish something.",
    "an excerpt",
    "Which line surprised you? Which sense was it using?",
    NOTES_LAUNCH,
    FOOTER
  );

  // SLIDE 4 -- LI / SC
  liSlide(
    pres,
    [
      "We are learning to read a colour poem, then write our own using all five senses",
    ],
    [
      "I can write a line about how a colour looks",
      "I can use specific, vivid images for different senses",
      "I can include a simile or metaphor in one line",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // SLIDE 5 -- I Do (the colour poem form + clean model)
  modellingSlide(
    pres,
    "I Do",
    "Build a Colour Poem -- One Sense Per Line",
    "Take one colour.\nWrite a line for each sense:\n\n- Looks like...\n- Sounds like...\n- Smells like...\n- Tastes like...\n- Feels like...\n\nChoose SPECIFIC things,\nnot just \"nice\".",
    "Green is the sound of\nleaves in the wind.\n\nGreen is the smell of\nfresh-cut grass.\n\nGreen is the taste of\na crisp apple.",
    NOTES_IDO,
    FOOTER
  );

  // SLIDE 6 -- We Do (build a colour poem together)
  boardBuildSlide(
    pres,
    "We Do",
    "Build a Colour Poem Together -- Red",
    "One sense per line. Push to a SPECIFIC picture each time.",
    NOTES_WEDO,
    FOOTER,
    {
      promptText: "Teacher drafts each line here. For every line ask: what specific thing?",
      prefilledHints: ["Looks", "Sounds", "Smells", "Tastes", "Feels"],
      badgeColor: C.SECONDARY,
    }
  );

  // SLIDE 7 + 8 -- CFU (stronger image) with reveal
  withReveal(
    () => cfuSlide(
      pres,
      "CFU",
      "Which Paints a Stronger Picture?",
      "Show Me Fingers: 1 (A) or 2 (B)",
      "Both lines are about yellow.\n\nA)  Yellow is nice.\n\nB)  Yellow smells like warm summer sand.",
      NOTES_CFU,
      FOOTER
    ),
    (slide) => {
      addRevealAnswerBar(
        slide,
        "B is stronger -- it uses a sense (smell) and a specific picture.",
        { color: C.SUCCESS, label: "Answer", showTickAndFix: false }
      );
      slide.addNotes(NOTES_CFU_REVEAL);
    }
  );

  // SLIDE 9 -- Optional Re-teach
  modellingSlide(
    pres,
    "Optional Re-teach",
    "Vague to Vivid -- Add a Specific Thing",
    "The trick: add a SPECIFIC thing.\n\n- \"Blue is calm.\"\n   -> no picture\n\n- \"Blue is the calm of\n   still water at dawn.\"\n   -> now I can see it\n\nYour turn: improve \"Red is hot\".",
    "Before:\nBlue is cold.\n\nAfter:\nBlue tastes like ice on\na winter morning.",
    NOTES_RETEACH,
    FOOTER
  );

  // SLIDE 10 -- You Do
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "You Do", { color: C.PRIMARY, w: 1.5 });
    addTitle(s, "Write Your Colour Poem");

    addCard(s, 0.5, CONTENT_TOP, 9, 1.7, { strip: C.PRIMARY, fill: C.WHITE });
    s.addText("Use your Colour Poem Plan:", {
      x: 0.75, y: CONTENT_TOP + 0.12, w: 8.4, h: 0.32,
      fontSize: 16, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("First:   Choose a colour you have a strong feeling about.\nNext:   Write one line for each sense -- looks, sounds, smells, tastes, feels.\nThen:   For each line, picture a SPECIFIC thing.", {
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
    s.addText("- Stuck? Use the sentence frames and the sense word bank on your plan.\n- Ready for more? Add a simile or metaphor, and write extra lines like the Orange poem.", {
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
      "The colour is: blue.",
      "Write ONE line that uses smell OR sound. Make it a specific picture.",
    ],
    NOTES_EXIT,
    FOOTER,
    { assessesSc: 2, title: "One Vivid Sense Line" }
  );

  // SLIDE 12 -- Closing
  closingSlide(
    pres,
    {
      reflectionPrompt: "Read your favourite line to a partner. Which sense does it use?",
      scItems: [
        "I can write a line about how a colour looks",
        "I can use specific, vivid images for different senses",
        "I can include a simile or metaphor in one line",
      ],
      selfAssessment: {
        prompt: "Show on your fingers: 1 (need help) to 5 (could teach it)",
        options: ["1-2", "3", "4-5"],
      },
    },
    NOTES_CLOSING
  );

  // ---- PDF: Colour Poem Plan ------------------------------------------------
  const pl = createPdf({ title: PLAN_RESOURCE.name });
  let py = addPdfHeader(pl, "Colour Poem Plan -- All Five Senses", {
    color: C.PRIMARY,
    subtitle: "One colour. One line per sense. Make each one a specific picture.",
    lessonInfo: "Week 2 Session 2 | Year 6 Literacy | Poetry",
    showNameDate: true,
  });

  py = addTipBox(pl, "My colour is: ____________.  Write a line for each sense. Picture a specific thing every time.", py, { color: C.PRIMARY });

  py = addSectionHeading(pl, "Looks like...", py, { color: C.PRIMARY });
  py = addLinedArea(pl, py, 1, { lineSpacing: 24 });
  py = addSectionHeading(pl, "Sounds like...", py, { color: C.SECONDARY });
  py = addLinedArea(pl, py, 1, { lineSpacing: 24 });
  py = addSectionHeading(pl, "Smells like...", py, { color: C.PRIMARY });
  py = addLinedArea(pl, py, 1, { lineSpacing: 24 });
  py = addSectionHeading(pl, "Tastes like...", py, { color: C.SECONDARY });
  py = addLinedArea(pl, py, 1, { lineSpacing: 24 });
  py = addSectionHeading(pl, "Feels like...", py, { color: C.ACCENT });
  py = addLinedArea(pl, py, 1, { lineSpacing: 24 });
  py += 8;

  py = addSectionHeading(pl, "Sentence frames + sense words", py, { color: C.PRIMARY });
  py = addBodyText(pl, "Frames: [Colour] looks like ___.  [Colour] is the sound of ___.  [Colour] tastes like ___.", py, { fontSize: 11 });
  py = addBodyText(pl, "Sound: roar, whisper, hum, crash, sizzle, rustle.  Smell: smoke, rain, salt, blossom, fresh bread.", py, { fontSize: 11 });
  py = addBodyText(pl, "Taste: honey, mint, lemon, warm toast.  Feel: warm sand, cold glass, soft moss, sharp wind.", py, { fontSize: 11 });

  addPdfFooter(pl, "Week 2 Session 2 | Colour Poem Plan -- Page 1");

  pl.addPage();
  let py2 = addPdfHeader(pl, "My Colour Poem -- Write It Here", {
    color: C.PRIMARY,
    subtitle: "Read it aloud. Can a reader see, hear and smell your colour?",
    lessonInfo: "Week 2 Session 2 | Year 6 Literacy",
    showNameDate: false,
  });
  py2 = addTipBox(pl, "Write your colour poem here. Challenge: add a simile or metaphor in one line.", py2, { color: C.SECONDARY });
  py2 = addLinedArea(pl, py2, 8, { lineSpacing: 26 });

  addPdfFooter(pl, "Week 2 Session 2 | My Colour Poem -- Page 2");

  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Poetry_W2S2.pptx` }),
    writePdf(pl, PLAN_PDF_PATH),
  ]);

  console.log("PPTX written to " + `${OUT_DIR}/Poetry_W2S2.pptx`);
  console.log("Done: " + PLAN_RESOURCE.name);
}

build().catch((err) => {
  console.error(err);
  process.exit(1);
});
