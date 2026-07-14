"use strict";

// Poetry Unit (Year 6) - Week 5 Session 2: Drafting a Verse Poem
// Form: Verse (poetry with a steady beat; rhyme ABAB), with at least one image.
// Pulls together the unit's two big skills: sound (beat + rhyme) and imagery.
// Catch-up: opens with a 20-second recap of verse.
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
const FOOTER = "Poetry | Week 5 Session 2 | Year 6 Literacy";
const OUT_DIR = "output/Poetry_W5S2_Draft_Verse";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const PLAN_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Verse Plan",
  "Verse planning template (ABAB, with a beat) plus rhyme families and an image prompt."
);
const RESOURCE_ITEMS = [PLAN_RESOURCE];
const PLAN_PDF_PATH = path.join(OUT_DIR, PLAN_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Today we write our own verse -- a poem with a steady beat, a rhyme, and a picture.
- This is the big one: it brings together sound AND imagery, everything we have learned.
- If you missed last session, do not worry -- I will remind you what verse is.

DO:
- Display the title slide.
- Read a short verse aloud to set the beat in the room.

TEACHER NOTES:
This is the unit's most complete writing task. Keep it to four lines so students can manage beat, rhyme and image together.

WATCH FOR:
- Students who try to do everything at once -- build it one layer at a time.

[Literacy: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RECAP = `SAY:
- Quick reminder, in twenty seconds. Verse is poetry with a steady beat you can feel.
- Our verse will rhyme ABAB -- lines one and three rhyme, lines two and four rhyme.
- And it will hold at least one image -- a simile, metaphor or personification.

DO:
- Display the recap and read the example with a steady beat.
- Point to the ABAB labels and the image.
- This is your catch-up moment for anyone who missed Session 1.

TEACHER NOTES:
Re-establish verse: steady beat, ABAB rhyme, and an image. Tap the beat together.

WATCH FOR:
- Students who forget the beat -- read aloud and tap to bring it back.

[Literacy: Launch / Recap | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me. Today we WRITE a verse with a beat, an ABAB rhyme, and an image.
- The first "I can" is for everyone -- write two lines with a steady beat.

DO:
- Choral read the LI and success criteria.
- Remind students to read aloud as they write, to keep the beat.

TEACHER NOTES:
SC1 (all): two lines with a steady beat. SC2 (target): a four-line ABAB verse. SC3 (depth): include an image that adds to the feeling. Exit ticket targets SC1 and SC3.

WATCH FOR:
- Students who get the rhyme but lose the beat -- reading aloud is the fix.

[Literacy: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO = `SAY:
- Watch me write a verse. I build it one layer at a time.
- Line one, with a steady beat: "The morning light begins to creep". Tap it -- da-DUM, da-DUM.
- Line two rhymes later with line four, and I add an image: "And spills like honey down the hill". Spills like honey -- a simile.
- Line three rhymes with line one (creep): "The valley, soft and still asleep". Asleep -- and that personifies the valley.
- Line four rhymes with line two (hill): "Lies hushed and golden, calm and still". ABAB, with a beat and pictures.

DO:
- Display the model and read it with the beat.
- Build it layer by layer: beat first, then rhyme, then the image.
- Point to the simile and the personification.

TEACHER NOTES:
Model building in layers so students do not try to control everything at once. Highlight the images that add feeling.

WATCH FOR:
- Students who copy the model -- redirect: "Your picture will be different."

[Literacy: I Do | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_WEDO = `SAY:
- Let us write a verse together about the night sky.
- Line one with a steady beat. Line two with an image. Line three rhymes with line one. Line four rhymes with line two.
- We will read each line aloud to keep the beat.

DO:
- Build the verse live, reading aloud after each line.
- Add at least one image and keep the ABAB rhyme.
- Read the whole verse with the beat at the end.

TEACHER NOTES:
Guided practice combining beat, rhyme and image. Keep reading aloud so the rhythm stays steady.

WATCH FOR:
- Students who stuff in extra words -- read aloud to find the bumpy beat.

[Literacy: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_CFU = `SAY:
- Quick check. Line one of a verse ends in "night". Line three must rhyme with line one and keep the beat.
- Read both options. Which line three works as verse?
- Show me on your fingers: one for A, two for B.

DO:
- Display the line and both options.
- Show Me Fingers.
- Scan: most should choose B.

CFU CHECKPOINT:
Technique: Show Me Fingers (1 for A, 2 for B)
Script:
- "Which line three rhymes with night AND keeps a steady beat?"
- Scan for: students choose B and can say it rhymes and has an image.
PROCEED (>=80%): Release to You Do.
PIVOT (<80%): Use the optional re-teach slide that follows -- counting beats to keep a steady rhythm.

TEACHER NOTES:
A does not rhyme with "night" and has no image. B rhymes (light), keeps the beat, and adds an image (silver light). B is the answer.

WATCH FOR:
- Students who pick A because it is true -- remind them verse needs rhyme AND beat.

[Literacy: CFU | VTLM 2.0: Formative Assessment]`;

const NOTES_CFU_REVEAL = `SAY:
- The answer is B: "The moon spills silver light." It rhymes with night, keeps the beat, and paints a picture.
- A was a true sentence but it did not rhyme or hold a steady beat.

DO:
- Display the reveal.
- Read the rhyming pair aloud with the beat.

[Literacy: CFU Reveal | VTLM 2.0: Formative Assessment]`;

const NOTES_RETEACH = `SAY:
- To keep a steady beat, count the beats in each line. Aim for about the same number in matching lines.
- "The morning light begins to creep" -- tap it: eight beats.
- "The valley, soft and still asleep" -- tap it: eight beats too. They match, so they feel even.
- If a line feels lumpy, count the beats and add or trim a word.

DO:
- Display the re-teach slide and count beats together.
- Re-check: students count the beats in a given line and compare it to another.

TEACHER NOTES:
OPTIONAL. Use only if the CFU was below 80%. Different approach: count the beats per line so matching lines feel even.

WATCH FOR:
- Students whose lines are wildly different lengths -- counting beats makes the unevenness visible.

[Literacy: Optional Re-teach | VTLM 2.0: Reteach with Different Approach]`;

const NOTES_YOUDO = `SAY:
- Your turn to write a verse. Build it one layer at a time.
- Write two lines with a steady beat, then make lines three and four rhyme back (ABAB).
- Slip in at least one image -- a simile, metaphor or personification.
- Read it aloud as you go. Fifteen minutes.

DO:
- Distribute the Verse Plan with its rhyme families and image prompt.
- Circulate. Prioritise students who needed the re-teach.
- Conference: "Read me your verse. Can I feel the beat? Where is your picture?"

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the given first line and a rhyme family. Write line three to rhyme with it, keeping a steady beat.
- Extra Notes: Two beat-steady, rhyming lines meet SC1.
EXTENDING PROMPT:
- Task: Write eight lines (two verses), and use TWO different images that build one clear feeling.
- Extra Notes: This reaches the depth criterion -- sound and imagery creating effect.

TEACHER NOTES:
This is the unit's final original poem before publishing and the booklet. Capture strong verses to celebrate.

WATCH FOR:
- Students who lose the beat for the rhyme -- read aloud and tap to rebalance.

[Literacy: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket. Write TWO lines of verse with a steady beat.
- Include ONE image -- a simile, metaphor or personification.
- Read them aloud to check the beat. Two minutes.

DO:
- Two minutes, silent.
- As students finish, have them read the lines to a partner.

TEACHER NOTES:
Exit ticket targets SC1 and SC3 -- a steady beat and an image, the two skills the whole unit has built.

WATCH FOR:
- Lines with an image but no beat, or a beat but no image -- note which skill to support in publishing.

[Literacy: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Self-check. Show me on your fingers, one to five, for each "I can".
- Then read your verse to a partner with the beat, and point out your image.

DO:
- Run the fingers check for each success criterion.
- Sixty seconds partner share.
- Preview: "Next session is a workshop -- a chance to finish and publish any poems you still have."

TEACHER NOTES:
This sets up the catch-up workshop. Note which students have unfinished poems to prioritise.

[Literacy: Closing | VTLM 2.0: Review and Reflect]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Poetry Week 5 Session 2 -- Drafting a Verse";

  // SLIDE 1 -- Title
  titleSlide(
    pres,
    "Write Your Own Verse",
    "A steady beat, an ABAB rhyme, and a picture -- everything together",
    "Week 5 Session 2  |  Year 6 Literacy  |  Poetry",
    NOTES_TITLE
  );

  // SLIDE 2 -- Resources
  addResourceSlide(pres, RESOURCE_ITEMS, T, FOOTER, NOTES_RESOURCES());

  // SLIDE 3 -- Recap / Launch
  modellingSlide(
    pres,
    "Quick Recap",
    "Verse -- in 20 Seconds",
    "Verse = poetry with a\nsteady beat.\n\n- Rhyme ABAB (lines 1 & 3,\n   lines 2 & 4)\n- At least one image\n- A beat you can feel\n\nMissed last session?\nThis is all you need.",
    "I wondered lonely as a cloud   (A)\nThat floats on high o'er vales and hills,   (B)\nWhen all at once I saw a crowd,   (A)\nA host of golden daffodils.   (B)",
    NOTES_RECAP,
    FOOTER
  );

  // SLIDE 4 -- LI / SC
  liSlide(
    pres,
    [
      "We are learning to WRITE a verse with a steady beat, an ABAB rhyme, and an image",
    ],
    [
      "I can write two lines with a steady beat",
      "I can write a four-line verse that rhymes ABAB",
      "I can include an image that adds to the feeling",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // SLIDE 5 -- I Do
  modellingSlide(
    pres,
    "I Do",
    "Build It One Layer at a Time",
    "My thinking:\n\n1. BEAT first -- tap line 1.\n\n2. Add an IMAGE in line 2\n   (\"spills like honey\").\n\n3. RHYME: line 3 with line 1,\n   line 4 with line 2 (ABAB).\n\nBeat + rhyme + picture.",
    "The morning light begins to creep,   (A)\n\nAnd spills like honey down the hill,   (B)\n\nThe valley, soft and still asleep,   (A)\n\nLies hushed and golden, calm and still.   (B)",
    NOTES_IDO,
    FOOTER
  );

  // SLIDE 6 -- We Do
  boardBuildSlide(
    pres,
    "We Do",
    "Build a Verse Together -- The Night Sky",
    "Beat first. Add an image. Lines 1 & 3 rhyme, lines 2 & 4 rhyme. Read each line aloud.",
    NOTES_WEDO,
    FOOTER,
    {
      promptText: "Teacher drafts each line here. Read aloud after each one to keep a steady beat.",
      prefilledHints: ["A", "B", "A", "B"],
      badgeColor: C.SECONDARY,
    }
  );

  // SLIDE 7 + 8 -- CFU with reveal
  withReveal(
    () => cfuSlide(
      pres,
      "CFU",
      "Which Line Three Works as Verse?",
      "Show Me Fingers: 1 (A) or 2 (B)",
      "Line 1:  The owl is calling through the night,\n\nA)  The stars are very pretty up there.\n\nB)  The moon spills silver light.",
      NOTES_CFU,
      FOOTER
    ),
    (slide) => {
      addRevealAnswerBar(
        slide,
        "B -- it rhymes with \"night\" (light), keeps the beat, and paints a picture.",
        { color: C.SUCCESS, label: "Answer", showTickAndFix: false }
      );
      slide.addNotes(NOTES_CFU_REVEAL);
    }
  );

  // SLIDE 9 -- Optional Re-teach
  modellingSlide(
    pres,
    "Optional Re-teach",
    "Count the Beats to Stay Steady",
    "Count the beats in each line.\nMatching lines should feel\nabout the same.\n\n\"The morning light begins\nto creep\" -- 8 beats.\n\n\"The valley, soft and still\nasleep\" -- 8 beats. They match.",
    "Your turn:\n\nCount the beats in this line --\n\"The river runs beside\nthe trees\"\n\nDoes it match 8?",
    NOTES_RETEACH,
    FOOTER
  );

  // SLIDE 10 -- You Do
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "You Do", { color: C.PRIMARY, w: 1.5 });
    addTitle(s, "Write Your Verse");

    addCard(s, 0.5, CONTENT_TOP, 9, 1.7, { strip: C.PRIMARY, fill: C.WHITE });
    s.addText("Use your Verse Plan -- build in layers:", {
      x: 0.75, y: CONTENT_TOP + 0.12, w: 8.4, h: 0.32,
      fontSize: 16, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("First:   Two lines with a steady beat.\nNext:   Lines 3 and 4 rhyme back (ABAB).\nThen:   Slip in at least one image. Read it aloud.", {
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
    s.addText("- Stuck? Use the given first line and a rhyme family on your plan.\n- Ready for more? Write eight lines with TWO images that build one clear feeling.", {
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
      "Write TWO lines of verse with a steady beat.",
      "Include ONE image -- a simile, metaphor or personification.",
    ],
    NOTES_EXIT,
    FOOTER,
    { assessesSc: 3, title: "Two Lines With a Beat and a Picture" }
  );

  // SLIDE 12 -- Closing
  closingSlide(
    pres,
    {
      reflectionPrompt: "Read your verse to a partner with the beat, and point out your image.",
      scItems: [
        "I can write two lines with a steady beat",
        "I can write a four-line verse that rhymes ABAB",
        "I can include an image that adds to the feeling",
      ],
      selfAssessment: {
        prompt: "Show on your fingers: 1 (need help) to 5 (could teach it)",
        options: ["1-2", "3", "4-5"],
      },
    },
    NOTES_CLOSING
  );

  // ---- PDF: Verse Plan ------------------------------------------------------
  const pl = createPdf({ title: PLAN_RESOURCE.name });
  let py = addPdfHeader(pl, "Verse Plan -- Beat, Rhyme and Picture", {
    color: C.PRIMARY,
    subtitle: "Four lines, ABAB, with a steady beat and an image.",
    lessonInfo: "Week 5 Session 2 | Year 6 Literacy | Poetry",
    showNameDate: true,
  });

  py = addTipBox(pl, "Build it in layers: beat first, then rhyme, then slip in a picture. Read aloud as you go.", py, { color: C.PRIMARY });

  py = addSectionHeading(pl, "Starter line -- use it or change it", py, { color: C.PRIMARY });
  py = addBodyText(pl, "Line 1:  The evening sky begins to ________,", py, { fontSize: 12 });
  py += 6;

  py = addSectionHeading(pl, "Rhyme families", py, { color: C.PRIMARY });
  py = addBodyText(pl, "-ow / glow: glow, slow, flow, below, snow", py, { fontSize: 11 });
  py = addBodyText(pl, "-ight: light, night, bright, sight, white", py, { fontSize: 11 });
  py = addBodyText(pl, "-eep: creep, sleep, deep, sweep, asleep", py, { fontSize: 11 });
  py += 8;

  py = addSectionHeading(pl, "Image prompt -- pick one to slip in", py, { color: C.SECONDARY });
  py = addBodyText(pl, "Simile: ... like ___.   Metaphor: the ___ is a ___.   Personification: the ___ whispered / marched / slept.", py, { fontSize: 11 });
  py += 6;

  py = addSectionHeading(pl, "Line 2 (B) -- add your image here", py, { color: C.PRIMARY });
  py = addLinedArea(pl, py, 1, { lineSpacing: 24 });
  py = addSectionHeading(pl, "Line 3 (A) -- rhymes with line 1", py, { color: C.SECONDARY });
  py = addLinedArea(pl, py, 1, { lineSpacing: 24 });
  py = addSectionHeading(pl, "Line 4 (B) -- rhymes with line 2", py, { color: C.ACCENT });
  py = addLinedArea(pl, py, 1, { lineSpacing: 24 });

  addPdfFooter(pl, "Week 5 Session 2 | Verse Plan -- Page 1");

  pl.addPage();
  let py2 = addPdfHeader(pl, "My Verse -- Write It Here", {
    color: C.PRIMARY,
    subtitle: "Read it aloud. Can you feel the beat? Where is your picture?",
    lessonInfo: "Week 5 Session 2 | Year 6 Literacy",
    showNameDate: false,
  });
  py2 = addTipBox(pl, "Write your finished verse. Label the lines A, B, A, B and underline your image.", py2, { color: C.SECONDARY });
  py2 = addLinedArea(pl, py2, 4, { lineSpacing: 28 });
  py2 += 10;
  py2 = addSectionHeading(pl, "Challenge -- eight lines, two images, one feeling", py2, { color: C.ACCENT });
  py2 = addLinedArea(pl, py2, 4, { lineSpacing: 28 });

  addPdfFooter(pl, "Week 5 Session 2 | My Verse -- Page 2");

  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Poetry_W5S2.pptx` }),
    writePdf(pl, PLAN_PDF_PATH),
  ]);

  console.log("PPTX written to " + `${OUT_DIR}/Poetry_W5S2.pptx`);
  console.log("Done: " + PLAN_RESOURCE.name);
}

function NOTES_RESOURCES() {
  return `SAY:
- One sheet today: your Verse Plan. It has a starter line, rhyme families and an image prompt.
- Keep your draft -- you may publish it in our workshop next session.

DO:
- Print the Verse Plan, one per student.
- Have the Poets Toolbox nearby.

[Literacy: Resources | VTLM 2.0: Enabling Learning]`;
}

build().catch((err) => {
  console.error(err);
  process.exit(1);
});
