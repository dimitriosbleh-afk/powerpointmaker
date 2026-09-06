"use strict";

// Poetry Unit (Year 6) - Week 1 Session 3: Drafting a Haiku
// Form: Haiku (5-7-5 syllables, usually about nature).
// Catch-up design: opens with a 20-second recap of the haiku shape.
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
const FOOTER = "Poetry | Week 1 Session 3 | Year 6 Literacy";
const OUT_DIR = "output/Poetry_W1S3_Draft_Haiku";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const PLAN_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Haiku Plan",
  "Haiku planning template with syllable-counting boxes (5 - 7 - 5) and a nature word bank."
);
const RESOURCE_ITEMS = [PLAN_RESOURCE];
const PLAN_PDF_PATH = path.join(OUT_DIR, PLAN_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Today we write a haiku -- a tiny nature poem from Japan.
- The fun of a haiku is the puzzle: fitting your picture into exactly five, seven, five beats.
- If you missed last session, do not worry. I will remind you how a haiku works.

DO:
- Display the title slide.
- Keep the tone playful -- a haiku is a word puzzle, not a test.

TEACHER NOTES:
Counting syllables is the core skill today. Plan to clap a lot, together and out loud.

WATCH FOR:
- Students who think they cannot count syllables -- the chin-drop trick on the re-teach slide helps.

[Literacy: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RECAP = `SAY:
- Quick reminder of the haiku, in twenty seconds.
- Three lines, all about nature, and we count the beats in the words -- the syllables.
- Line one has five beats, line two has seven, line three has five. Five, seven, five.
- Clap the example with me to feel the beats.

DO:
- Display the recap and clap each line with the class.
- Write 5, 7, 5 beside the three lines.
- This is your catch-up moment for anyone who missed Session 1.

TEACHER NOTES:
Re-establish the clapping routine before drafting. Absent students rejoin here.

WATCH FOR:
- Students who clap words, not beats -- model "drag-on-fly" as three beats.

[Literacy: Launch / Recap | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me. Today we WRITE a haiku with the 5-7-5 beat.
- Everyone can reach the first "I can" -- write three lines about nature.

DO:
- Choral read the LI and success criteria.
- Remind students this is a draft, and counting beats takes a few tries.

TEACHER NOTES:
SC1 (all): write a three-line nature haiku. SC2 (target): match the 5-7-5 syllable count. SC3 (depth): include an image. Exit ticket targets SC2 (counting beats).

WATCH FOR:
- Students who chase the count and forget the picture -- remind them the nature image matters too.

[Literacy: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO = `SAY:
- Watch me write a haiku. I start with a small nature moment -- a single leaf falling.
- Line one, five beats. "One leaf lets go now." Let me clap it: one-leaf-lets-go-now. Five.
- Line two, seven beats. "Spinning to the cold, wet ground." Clap: spin-ning-to-the-cold-wet-ground. Seven.
- Line three, five beats. "Autumn says hello." Clap: au-tumn-says-hel-lo. Five.
- See how I counted and adjusted? That is the haiku puzzle.

DO:
- Display the model and read it aloud.
- Clap each line with the class to prove the counts.
- Show one fix: if a line is too long, model dropping or swapping a word.

TEACHER NOTES:
Make the counting visible and physical. The point is the process: draft, clap, adjust. "Autumn says hello" also personifies autumn -- mention it lightly to connect to imagery.

WATCH FOR:
- Students who think the words must rhyme -- clarify: haiku do not rhyme, they count beats.

[Literacy: I Do | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_WEDO = `SAY:
- Let us build a haiku together about the rain.
- Line one, five beats. Call out ideas and we will clap to check.
- Line two, seven beats. What is the rain doing? Let us count it.
- Line three, five beats to finish the picture. Clap the whole thing back.

DO:
- Build the haiku live, clapping each line to test the count.
- If a line is over or under, fix it together by swapping a word.
- Read the finished class haiku aloud.

TEACHER NOTES:
Guided practice in the counting-and-adjusting cycle. Keep the picture about nature.

WATCH FOR:
- Students who add an extra describing word that breaks the count -- model trimming back to seven.

[Literacy: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_CFU = `SAY:
- Quick check. Here is a haiku line. Clap it and count the beats.
- "The bright morning sun." How many beats?
- Show me on your fingers.

DO:
- Display the line.
- Show Me Fingers for the syllable count.
- Scan: most should show 5.

CFU CHECKPOINT:
Technique: Show Me Fingers (number of syllables)
Script:
- "How many beats in this line?"
- Scan for: most students show five fingers.
PROCEED (>=80%): Release to You Do.
PIVOT (<80%): Use the optional re-teach slide that follows -- the chin-drop trick for counting syllables.

TEACHER NOTES:
The-bright-morn-ing-sun is five syllables. This line would work as line 1 or line 3 of a haiku.

WATCH FOR:
- Students who count four (saying "morning" as one beat) -- clap "morn-ing" slowly.

[Literacy: CFU | VTLM 2.0: Formative Assessment]`;

const NOTES_CFU_REVEAL = `SAY:
- Five beats: the-bright-morn-ing-sun.
- The tricky word is "morning" -- it is two beats, morn-ing.
- A five-beat line like this fits line one or line three of a haiku.

DO:
- Display the reveal.
- Clap the line once more, stretching "morn-ing".

[Literacy: CFU Reveal | VTLM 2.0: Formative Assessment]`;

const NOTES_RETEACH = `SAY:
- Here is a trick for counting beats. Put your hand flat under your chin.
- Every time your chin drops to say a word part, that is one beat.
- Try "but-ter-fly". Three chin drops, three beats.
- Now try "wa-ter-fall". Three. Try "sky". One. The chin does not lie.

DO:
- Display the re-teach slide and demonstrate the chin-drop method.
- Practise three words together as a class.
- Re-check: students count the beats in "the silver moonlight" on their fingers.

TEACHER NOTES:
OPTIONAL. Use only if the CFU was below 80%. Different approach: a physical chin-drop method instead of clapping, which helps students who lose the beat when clapping.

WATCH FOR:
- Students still unsure -- have them whisper the word slowly with a hand on the chin.

[Literacy: Optional Re-teach | VTLM 2.0: Reteach with Different Approach]`;

const NOTES_YOUDO = `SAY:
- Your turn to write a haiku. Pick one small thing from nature -- a snail, the wind, a puddle.
- Draft your three lines, then clap each one and check: five, seven, five.
- If a line is over, swap or drop a word. If it is short, add one.
- Fifteen minutes. I will come and clap them with you.

DO:
- Distribute the Haiku Plan with its syllable boxes.
- Circulate. Prioritise students who needed the re-teach.
- Conference by clapping a line together and counting.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the nature word bank. Build line one to five beats first using the counted words, then line three, then the seven-beat middle line.
- Extra Notes: Pre-counted words let students focus on the picture, not the maths.
EXTENDING PROMPT:
- Task: Add an image -- make one line a simile or personification (like "the wind tiptoes past"). Then try a second haiku for a different season.
- Extra Notes: Links the counting task back to the imagery skill.

TEACHER NOTES:
Students keep their haiku draft for publishing in Session 4. Celebrate clever fixes that hit the count.

WATCH FOR:
- Students stuck on the seven-beat line -- the middle line is hardest; help them stretch a five-beat idea.

[Literacy: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket. Write ONE haiku line that has exactly five beats.
- It can be about anything in nature. Clap it to be sure before you hand it in.
- One line, five beats. Two minutes.

DO:
- Two minutes, silent.
- As students finish, have them clap the line to a partner to check.

TEACHER NOTES:
Exit ticket targets SC2 -- matching a syllable count. A correct five-beat line shows the core haiku skill.

WATCH FOR:
- Lines with six or four beats -- note who needs more counting practice next week.

[Literacy: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Self-check. Show me on your fingers, one to five, for each "I can".
- Then read your haiku to a partner and have them clap to check your beats.

DO:
- Run the fingers check for each success criterion.
- Sixty seconds partner share with clapping.
- Preview: "Next session we publish both poems -- our three-line poem and our haiku."

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
  pres.title = "Poetry Week 1 Session 3 -- Drafting a Haiku";

  // SLIDE 1 -- Title
  titleSlide(
    pres,
    "Write Your Own Haiku",
    "A tiny nature poem -- five, seven, five beats",
    "Week 1 Session 3  |  Year 6 Literacy  |  Poetry",
    NOTES_TITLE
  );

  // SLIDE 2 -- Resources
  addResourceSlide(pres, RESOURCE_ITEMS, T, FOOTER, NOTES_RESOURCES());

  // SLIDE 3 -- Recap / Launch (catch-up friendly)
  modellingSlide(
    pres,
    "Quick Recap",
    "The Haiku -- in 20 Seconds",
    "Three lines about nature.\nCount the beats (syllables):\n\n- Line 1 = 5 beats\n- Line 2 = 7 beats\n- Line 3 = 5 beats\n\nNo rhyming needed.\nMissed last session? This is all you need.",
    "Winter is coming.\n(5)\n\nSnow will be arriving soon.\n(7)\n\nWe should rake the leaves.\n(5)",
    NOTES_RECAP,
    FOOTER
  );

  // SLIDE 4 -- LI / SC
  liSlide(
    pres,
    [
      "We are learning to WRITE a haiku about nature using the 5 - 7 - 5 beat",
    ],
    [
      "I can write three lines about something in nature",
      "I can match the beats: 5 in line 1, 7 in line 2, 5 in line 3",
      "I can add an image so a line paints a clear picture",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // SLIDE 5 -- I Do (teacher drafts a haiku, counting aloud)
  modellingSlide(
    pres,
    "I Do",
    "Watch Me Draft: Counting the Beats",
    "My thinking:\n\n1. A small nature moment\n   -> one leaf falling\n\n2. Draft, then CLAP\n   each line to check\n\n3. Too long? Swap or\n   drop a word\n\n5 - 7 - 5.",
    "One leaf lets go now\n(5)\n\nSpinning to the cold, wet ground\n(7)\n\nAutumn says hello\n(5)",
    NOTES_IDO,
    FOOTER
  );

  // SLIDE 6 -- We Do (build a haiku together)
  boardBuildSlide(
    pres,
    "We Do",
    "Build a Haiku Together -- About the Rain",
    "Call out each line. We clap to check the beats: 5 - 7 - 5.",
    NOTES_WEDO,
    FOOTER,
    {
      promptText: "Teacher drafts each line here. Clap together to test the count, then fix.",
      prefilledHints: ["5", "7", "5"],
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
      "Clap this line and count the beats.\n\nThe bright morning sun",
      NOTES_CFU,
      FOOTER
    ),
    (slide) => {
      addRevealAnswerBar(
        slide,
        "5 beats: the - bright - morn - ing - sun. (\"morning\" is two beats.)",
        { color: C.SUCCESS, label: "Answer", showTickAndFix: false }
      );
      slide.addNotes(NOTES_CFU_REVEAL);
    }
  );

  // SLIDE 9 -- Optional Re-teach (chin-drop trick)
  modellingSlide(
    pres,
    "Optional Re-teach",
    "The Chin-Drop Trick for Counting Beats",
    "Put a flat hand under your chin.\n\nEvery time your chin drops,\nthat is one beat.\n\nTry these together:\n\n- but-ter-fly = 3\n- wa-ter-fall = 3\n- sky = 1",
    "Now you try:\n\nThe silver moonlight\n\nHow many beats?\nUse the chin-drop to check.",
    NOTES_RETEACH,
    FOOTER
  );

  // SLIDE 10 -- You Do
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "You Do", { color: C.PRIMARY, w: 1.5 });
    addTitle(s, "Write Your Haiku");

    addCard(s, 0.5, CONTENT_TOP, 9, 1.7, { strip: C.PRIMARY, fill: C.WHITE });
    s.addText("Use your Haiku Plan:", {
      x: 0.75, y: CONTENT_TOP + 0.12, w: 8.4, h: 0.32,
      fontSize: 16, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("First:   Pick one small thing in nature.\nNext:   Draft three lines, then clap each one -- 5, 7, 5.\nThen:   Too long? Swap or drop a word. Too short? Add one.", {
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
    s.addText("- Stuck? Use the counted nature word bank to build line 1 and line 3 first.\n- Ready for more? Add an image (a simile or personification), then try a second haiku for another season.", {
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
      "Write ONE haiku line with exactly FIVE beats.",
      "About anything in nature. Clap it to check before you hand it in.",
    ],
    NOTES_EXIT,
    FOOTER,
    { assessesSc: 2, title: "A Five-Beat Line" }
  );

  // SLIDE 12 -- Closing
  closingSlide(
    pres,
    {
      reflectionPrompt: "Read your haiku to a partner. Have them clap to check your beats.",
      scItems: [
        "I can write three lines about something in nature",
        "I can match the beats: 5 in line 1, 7 in line 2, 5 in line 3",
        "I can add an image so a line paints a clear picture",
      ],
      selfAssessment: {
        prompt: "Show on your fingers: 1 (need help) to 5 (could teach it)",
        options: ["1-2", "3", "4-5"],
      },
    },
    NOTES_CLOSING
  );

  // ---- PDF: Haiku Plan ------------------------------------------------------
  const pl = createPdf({ title: PLAN_RESOURCE.name });
  let py = addPdfHeader(pl, "Haiku Plan -- Count the Beats", {
    color: C.PRIMARY,
    subtitle: "A nature poem in five, seven, five beats.",
    lessonInfo: "Week 1 Session 3 | Year 6 Literacy | Poetry",
    showNameDate: true,
  });

  py = addTipBox(pl, "Draft a line, then clap it and write the number of beats in the box. Aim for 5 - 7 - 5.", py, { color: C.PRIMARY });

  py = addSectionHeading(pl, "Line 1 -- 5 beats", py, { color: C.PRIMARY });
  py = addLinedArea(pl, py, 1, { lineSpacing: 26 });
  py += 4;
  py = addSectionHeading(pl, "Line 2 -- 7 beats", py, { color: C.SECONDARY });
  py = addLinedArea(pl, py, 1, { lineSpacing: 26 });
  py += 4;
  py = addSectionHeading(pl, "Line 3 -- 5 beats", py, { color: C.ACCENT });
  py = addLinedArea(pl, py, 1, { lineSpacing: 26 });
  py += 8;

  py = addSectionHeading(pl, "Nature word bank (beats counted for you)", py, { color: C.PRIMARY });
  py = addBodyText(pl, "1 beat: moon, rain, frost, breeze, snow, leaf, dusk", py, { fontSize: 11 });
  py = addBodyText(pl, "2 beats: sunset, river, blossom, shadow, morning, ocean, silence", py, { fontSize: 11 });
  py = addBodyText(pl, "3 beats: butterfly, waterfall, dragonfly, distant hills, falling snow", py, { fontSize: 11 });

  addPdfFooter(pl, "Week 1 Session 3 | Haiku Plan -- Page 1");

  pl.addPage();
  let py2 = addPdfHeader(pl, "My Haiku -- Write It Here", {
    color: C.PRIMARY,
    subtitle: "Three lines. 5 - 7 - 5. Clap to check.",
    lessonInfo: "Week 1 Session 3 | Year 6 Literacy",
    showNameDate: false,
  });
  py2 = addTipBox(pl, "Write your haiku. Read it aloud and clap each line to be sure of the beats.", py2, { color: C.SECONDARY });
  py2 = addLinedArea(pl, py2, 3, { lineSpacing: 30 });
  py2 += 10;
  py2 = addSectionHeading(pl, "Challenge -- a second haiku (different season)", py2, { color: C.ACCENT });
  py2 = addLinedArea(pl, py2, 3, { lineSpacing: 30 });

  addPdfFooter(pl, "Week 1 Session 3 | My Haiku -- Page 2");

  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Poetry_W1S3.pptx` }),
    writePdf(pl, PLAN_PDF_PATH),
  ]);

  console.log("PPTX written to " + `${OUT_DIR}/Poetry_W1S3.pptx`);
  console.log("Done: " + PLAN_RESOURCE.name);
}

function NOTES_RESOURCES() {
  return `SAY:
- One sheet today: your Haiku Plan.
- It has a beat box for each line and a nature word bank with the beats already counted.
- Keep your draft -- we publish it next session.

DO:
- Print the Haiku Plan, one per student.
- Have the Poets Toolbox from Session 1 nearby for reference.

[Literacy: Resources | VTLM 2.0: Enabling Learning]`;
}

build().catch((err) => {
  console.error(err);
  process.exit(1);
});
