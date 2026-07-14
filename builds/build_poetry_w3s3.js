"use strict";

// Poetry Unit (Year 6) - Week 3 Session 3: Drafting a Tanka
// Form: Tanka (5-7-5-7-7 syllables, unrhymed; often tells a tiny story).
// Catch-up design: opens with a 20-second recap of the tanka shape.
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

const SESSION_NUMBER = 3;
const FOOTER = "Poetry | Week 3 Session 3 | Year 6 Literacy";
const OUT_DIR = "output/Poetry_W3S3_Draft_Tanka";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const PLAN_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Tanka Plan",
  "Tanka planning template with beat boxes (5-7-5-7-7) and a tiny-story prompt."
);
const RESOURCE_ITEMS = [PLAN_RESOURCE];
const PLAN_PDF_PATH = path.join(OUT_DIR, PLAN_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Today we write a tanka -- five lines that count beats, like a haiku with two extra lines.
- A tanka often tells a tiny story or captures one small moment.
- If you missed last session, do not worry -- I will remind you how a tanka works.

DO:
- Display the title slide.
- Have the example tanka ready to read and clap.

TEACHER NOTES:
Counting beats is the core skill, plus shaping a small moment. Plan to clap a lot.

WATCH FOR:
- Students who try to rhyme a tanka -- remind them it counts beats and does not rhyme.

[Literacy: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RECAP = `SAY:
- Quick reminder, in twenty seconds. A tanka has five lines and counts beats.
- Five, seven, five, seven, seven. Thirty-one beats in all. No rhyming.
- Clap the example with me to feel the beats.

DO:
- Display the recap and clap each line with the class.
- Write 5, 7, 5, 7, 7 beside the lines.
- This is your catch-up moment for anyone who missed Session 1.

TEACHER NOTES:
Re-establish the clapping routine and the 5-7-5-7-7 pattern before drafting.

WATCH FOR:
- Students who clap words, not beats -- model a multi-beat word slowly.

[Literacy: Launch / Recap | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me. Today we WRITE a tanka with the 5-7-5-7-7 beat.
- The first "I can" is for everyone -- write five lines about one moment.

DO:
- Choral read the LI and success criteria.
- Remind students the two seven-beat lines are the trickiest.

TEACHER NOTES:
SC1 (all): write a five-line tanka about a moment. SC2 (target): match the 5-7-5-7-7 beat. SC3 (depth): the poem turns or tells a tiny story. Exit ticket targets SC2.

WATCH FOR:
- Students who write a list rather than a moment -- nudge them toward one small scene.

[Literacy: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO = `SAY:
- Watch me write a tanka. I start with one small moment -- a frosty morning in the garden.
- Line one, five beats: "Morning frost has come". Clap it: morn-ing-frost-has-come. Five.
- Line two, seven beats: "The garden glitters with white". Clap: seven.
- Lines three, four, five build the little moment -- a bird, then it flies away.
- Notice the tanka tells a tiny story: frost, a bird arrives, the bird leaves.

DO:
- Display the model and read it aloud.
- Clap each line to prove the counts.
- Point out the small story across the five lines.

TEACHER NOTES:
Model both the counting and the tiny-story shape. The two seven-beat lines are where students will need the most help.

WATCH FOR:
- Students who think each line must be a new idea -- show how the five lines hold ONE moment.

[Literacy: I Do | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_WEDO = `SAY:
- Let us build a tanka together about a sudden thunderstorm.
- Line one, five beats. Call out ideas and we will clap to check.
- Then seven, five, seven, seven. We will tell the little story of the storm arriving.

DO:
- Build the tanka live, clapping each line to test the count.
- Fix any line that is over or under by swapping a word.
- Read the finished class tanka aloud.

TEACHER NOTES:
Guided practice in counting and shaping a moment. Keep the five lines as one connected scene.

WATCH FOR:
- Students who break the seven-beat lines -- help them stretch a short idea to seven.

[Literacy: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_CFU = `SAY:
- Quick check. Clap this tanka line and count the beats.
- "It pecks at the frozen ground." How many beats?
- Show me on your fingers.

DO:
- Display the line.
- Show Me Fingers for the syllable count.
- Scan: most should show 7.

CFU CHECKPOINT:
Technique: Show Me Fingers (number of beats)
Script:
- "How many beats in this line?"
- Scan for: most students show seven.
PROCEED (>=80%): Release to You Do.
PIVOT (<80%): Use the optional re-teach slide that follows -- stretching a line to exactly seven beats.

TEACHER NOTES:
It-pecks-at-the-fro-zen-ground is seven beats. This is a tanka line 2 or 4.

WATCH FOR:
- Students who count six (saying "frozen" as one beat) -- clap "fro-zen" slowly.

[Literacy: CFU | VTLM 2.0: Formative Assessment]`;

const NOTES_CFU_REVEAL = `SAY:
- Seven beats: it-pecks-at-the-fro-zen-ground.
- The tricky word is "frozen" -- two beats, fro-zen.
- A seven-beat line like this is line 2, 4 or 5 of a tanka.

DO:
- Display the reveal.
- Clap the line once more, stretching "fro-zen".

[Literacy: CFU Reveal | VTLM 2.0: Formative Assessment]`;

const NOTES_RETEACH = `SAY:
- The seven-beat lines are the hardest. Here is how to grow a short line to seven.
- Start with five beats: "The waves crash on rocks". Clap -- five.
- Add a describing word or two: "The grey waves crash on jagged rocks". Clap -- now seven.
- Adding precise detail fills the beats AND paints a better picture.

DO:
- Display the re-teach slide and grow the line together.
- Re-check: students grow "Rain falls on the street" to seven beats on whiteboards.

TEACHER NOTES:
OPTIONAL. Use only if the CFU was below 80%. Different approach: grow a short line to seven by adding precise detail, which serves both the count and the imagery.

WATCH FOR:
- Students who add filler words ("very", "really") -- steer them to descriptive words that add a picture.

[Literacy: Optional Re-teach | VTLM 2.0: Reteach with Different Approach]`;

const NOTES_YOUDO = `SAY:
- Your turn to write a tanka. Pick one small moment -- a storm, a pet, a sunset, a memory.
- Draft five lines, then clap each one: five, seven, five, seven, seven.
- Try to let the five lines tell one tiny story.
- Fifteen minutes. I will come and clap them with you.

DO:
- Distribute the Tanka Plan with its beat boxes.
- Circulate. Prioritise students who needed the re-teach.
- Conference by clapping a line together and counting.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the counted word bank. Build the five-beat lines first, then grow them into the seven-beat lines.
- Extra Notes: Pre-counted words let students focus on the moment, not the maths.
EXTENDING PROMPT:
- Task: Make your tanka TURN -- something changes between the start and the end (a sound breaks the calm, a feeling shifts). Add an image too.
- Extra Notes: The turn is what makes a tanka feel like a tiny story.

TEACHER NOTES:
Students keep their tanka draft to publish in Session 4. Celebrate clever fixes that hit the count and a clear small moment.

WATCH FOR:
- Students stuck on the seven-beat lines -- use the grow-a-line strategy from the re-teach.

[Literacy: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket. Write ONE tanka line that has exactly SEVEN beats.
- It can be about anything. Clap it to be sure before you hand it in.
- One line, seven beats. Two minutes.

DO:
- Two minutes, silent.
- As students finish, have them clap the line to a partner to check.

TEACHER NOTES:
Exit ticket targets SC2 -- matching the seven-beat count, the hardest part of a tanka.

WATCH FOR:
- Lines with six or eight beats -- note who needs more counting practice.

[Literacy: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Self-check. Show me on your fingers, one to five, for each "I can".
- Then read your tanka to a partner and have them clap to check your beats.

DO:
- Run the fingers check for each success criterion.
- Sixty seconds partner share with clapping.
- Preview: "Next session we publish both poems -- our rhyming poem and our tanka."

TEACHER NOTES:
Note any students still struggling with the count so you can pair them well for publishing.

[Literacy: Closing | VTLM 2.0: Review and Reflect]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Poetry Week 3 Session 3 -- Drafting a Tanka";

  // SLIDE 1 -- Title
  titleSlide(
    pres,
    "Write Your Own Tanka",
    "Five lines, 5 - 7 - 5 - 7 - 7 beats -- a tiny story in sound",
    "Week 3 Session 3  |  Year 6 Literacy  |  Poetry",
    NOTES_TITLE
  );

  // SLIDE 2 -- Resources
  addResourceSlide(pres, RESOURCE_ITEMS, T, FOOTER, NOTES_RESOURCES());

  // SLIDE 3 -- Recap / Launch
  modellingSlide(
    pres,
    "Quick Recap",
    "The Tanka -- in 20 Seconds",
    "Five lines. Count the beats:\n\n- 5 - 7 - 5 - 7 - 7\n- 31 beats in all\n- No rhyming\n- Often a tiny story\n\nMissed last session?\nThis is all you need.",
    "Crash at two A.M.\n(5)\nI opened my bedroom door\n(7)\nA white cat ran by\n(5)\nStartled by the clanging fall\n(7)\nOf the treat jar's metal lid\n(7)",
    NOTES_RECAP,
    FOOTER
  );

  // SLIDE 4 -- LI / SC
  liSlide(
    pres,
    [
      "We are learning to WRITE a tanka about one moment using the 5 - 7 - 5 - 7 - 7 beat",
    ],
    [
      "I can write five lines about one moment",
      "I can match the beats: 5 - 7 - 5 - 7 - 7",
      "I can make my tanka turn or tell a tiny story",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // SLIDE 5 -- I Do
  modellingSlide(
    pres,
    "I Do",
    "Watch Me Draft a Tanka",
    "My thinking:\n\n1. One small moment\n   -> a frosty morning\n\n2. Draft, then CLAP\n   each line\n\n3. Tell a tiny story:\n   frost, a bird, it flies\n\n5 - 7 - 5 - 7 - 7.",
    "Morning frost has come\n(5)\nThe garden glitters with white\n(7)\nA small bird hops near\n(5)\nIt pecks at the frozen ground\n(7)\nThen lifts into grey, cold air\n(7)",
    NOTES_IDO,
    FOOTER
  );

  // SLIDE 6 -- We Do
  boardBuildSlide(
    pres,
    "We Do",
    "Build a Tanka Together -- A Sudden Storm",
    "Call out each line. We clap to check: 5 - 7 - 5 - 7 - 7. Tell the little story of the storm.",
    NOTES_WEDO,
    FOOTER,
    {
      promptText: "Teacher drafts each line here. Clap together to test the count, then fix.",
      prefilledHints: ["5", "7", "5", "7", "7"],
      badgeColor: C.SECONDARY,
    }
  );

  // SLIDE 7 + 8 -- CFU (count the beats) with reveal
  withReveal(
    () => cfuSlide(
      pres,
      "CFU",
      "Count the Beats",
      "Show Me Fingers: how many beats?",
      "Clap this tanka line and count the beats.\n\nIt pecks at the frozen ground",
      NOTES_CFU,
      FOOTER
    ),
    (slide) => {
      addRevealAnswerBar(
        slide,
        "7 beats: it - pecks - at - the - fro - zen - ground. (\"frozen\" is two beats.)",
        { color: C.SUCCESS, label: "Answer", showTickAndFix: false }
      );
      slide.addNotes(NOTES_CFU_REVEAL);
    }
  );

  // SLIDE 9 -- Optional Re-teach
  modellingSlide(
    pres,
    "Optional Re-teach",
    "Grow a Line to Seven Beats",
    "Start with five beats:\n\"The waves crash on rocks\" (5)\n\nAdd precise detail:\n\"The grey waves crash on\njagged rocks\" (7)\n\nDetail fills the beats AND\npaints a better picture.",
    "Your turn:\n\nGrow this to 7 beats --\n\"Rain falls on the street\" (5)\n\nAdd describing words.",
    NOTES_RETEACH,
    FOOTER
  );

  // SLIDE 10 -- You Do
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "You Do", { color: C.PRIMARY, w: 1.5 });
    addTitle(s, "Write Your Tanka");

    addCard(s, 0.5, CONTENT_TOP, 9, 1.7, { strip: C.PRIMARY, fill: C.WHITE });
    s.addText("Use your Tanka Plan:", {
      x: 0.75, y: CONTENT_TOP + 0.12, w: 8.4, h: 0.32,
      fontSize: 16, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("First:   Pick one small moment.\nNext:   Draft five lines, then clap each -- 5, 7, 5, 7, 7.\nThen:   Let the five lines tell one tiny story.", {
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
    s.addText("- Stuck? Use the counted word bank -- build the 5-beat lines first, then grow the 7s.\n- Ready for more? Make your tanka TURN -- something changes from start to end.", {
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
      "Write ONE tanka line with exactly SEVEN beats.",
      "About anything. Clap it to check before you hand it in.",
    ],
    NOTES_EXIT,
    FOOTER,
    { assessesSc: 2, title: "A Seven-Beat Line" }
  );

  // SLIDE 12 -- Closing
  closingSlide(
    pres,
    {
      reflectionPrompt: "Read your tanka to a partner. Have them clap to check your beats.",
      scItems: [
        "I can write five lines about one moment",
        "I can match the beats: 5 - 7 - 5 - 7 - 7",
        "I can make my tanka turn or tell a tiny story",
      ],
      selfAssessment: {
        prompt: "Show on your fingers: 1 (need help) to 5 (could teach it)",
        options: ["1-2", "3", "4-5"],
      },
    },
    NOTES_CLOSING
  );

  // ---- PDF: Tanka Plan ------------------------------------------------------
  const pl = createPdf({ title: PLAN_RESOURCE.name });
  let py = addPdfHeader(pl, "Tanka Plan -- Count the Beats", {
    color: C.PRIMARY,
    subtitle: "Five lines, 5 - 7 - 5 - 7 - 7. One small moment.",
    lessonInfo: "Week 3 Session 3 | Year 6 Literacy | Poetry",
    showNameDate: true,
  });

  py = addTipBox(pl, "My moment is: ____________.  Draft each line, then clap it and write the beats in the box.", py, { color: C.PRIMARY });

  py = addSectionHeading(pl, "Line 1 -- 5 beats", py, { color: C.PRIMARY });
  py = addLinedArea(pl, py, 1, { lineSpacing: 24 });
  py = addSectionHeading(pl, "Line 2 -- 7 beats", py, { color: C.SECONDARY });
  py = addLinedArea(pl, py, 1, { lineSpacing: 24 });
  py = addSectionHeading(pl, "Line 3 -- 5 beats", py, { color: C.PRIMARY });
  py = addLinedArea(pl, py, 1, { lineSpacing: 24 });
  py = addSectionHeading(pl, "Line 4 -- 7 beats", py, { color: C.SECONDARY });
  py = addLinedArea(pl, py, 1, { lineSpacing: 24 });
  py = addSectionHeading(pl, "Line 5 -- 7 beats", py, { color: C.ACCENT });
  py = addLinedArea(pl, py, 1, { lineSpacing: 24 });
  py += 8;

  py = addSectionHeading(pl, "Counted word bank + a tip", py, { color: C.PRIMARY });
  py = addBodyText(pl, "1 beat: storm, rain, frost, leaf, dusk, light.  2 beats: thunder, morning, garden, silence, shadow.", py, { fontSize: 11 });
  py = addBodyText(pl, "3 beats: suddenly, distant hills, butterfly, waterfall.", py, { fontSize: 11 });
  py = addBodyText(pl, "Tip: grow a 5-beat line to 7 by adding precise describing words.", py, { fontSize: 11, italic: true });

  addPdfFooter(pl, "Week 3 Session 3 | Tanka Plan -- Page 1");

  pl.addPage();
  let py2 = addPdfHeader(pl, "My Tanka -- Write It Here", {
    color: C.PRIMARY,
    subtitle: "Five lines. 5 - 7 - 5 - 7 - 7. One tiny story.",
    lessonInfo: "Week 3 Session 3 | Year 6 Literacy",
    showNameDate: false,
  });
  py2 = addTipBox(pl, "Write your tanka. Read it aloud and clap each line to check the beats.", py2, { color: C.SECONDARY });
  py2 = addLinedArea(pl, py2, 5, { lineSpacing: 28 });
  py2 += 10;
  py2 = addSectionHeading(pl, "Challenge -- make it TURN (something changes)", py2, { color: C.ACCENT });
  py2 = addLinedArea(pl, py2, 5, { lineSpacing: 28 });

  addPdfFooter(pl, "Week 3 Session 3 | My Tanka -- Page 2");

  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Poetry_W3S3.pptx` }),
    writePdf(pl, PLAN_PDF_PATH),
  ]);

  console.log("PPTX written to " + `${OUT_DIR}/Poetry_W3S3.pptx`);
  console.log("Done: " + PLAN_RESOURCE.name);
}

function NOTES_RESOURCES() {
  return `SAY:
- One sheet today: your Tanka Plan. It has a beat box for each line and a counted word bank.
- Keep your draft -- we publish it next session.

DO:
- Print the Tanka Plan, one per student.
- Have the Poets Toolbox nearby.

[Literacy: Resources | VTLM 2.0: Enabling Learning]`;
}

build().catch((err) => {
  console.error(err);
  process.exit(1);
});
