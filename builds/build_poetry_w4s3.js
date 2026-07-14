"use strict";

// Poetry Unit (Year 6) - Week 4 Session 3: Drafting a Limerick
// Form: limerick (5 lines, AABBA, bouncy rhythm, funny).
// Catch-up: opens with a 20-second recap of the limerick shape.
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
const FOOTER = "Poetry | Week 4 Session 3 | Year 6 Literacy";
const OUT_DIR = "output/Poetry_W4S3_Draft_Limerick";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const PLAN_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Limerick Plan",
  "Limerick planning template (5 lines, AABBA) with a starter line, rhyme families and a rhythm guide."
);
const RESOURCE_ITEMS = [PLAN_RESOURCE];
const PLAN_PDF_PATH = path.join(OUT_DIR, PLAN_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Today we write a limerick -- five bouncy lines that end with a silly twist.
- Lots of famous limericks start with "There was a...". We can use that too.
- If you missed last session, do not worry -- I will remind you how a limerick works.

DO:
- Display the title slide.
- Read a limerick with a strong bounce to set the mood.

TEACHER NOTES:
The rhythm is the joy and the challenge. Read examples aloud often so students feel the da-da-DUM bounce.

WATCH FOR:
- Students who write five lines but lose the bounce -- reading aloud fixes this.

[Literacy: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RECAP = `SAY:
- Quick reminder, in twenty seconds. A limerick has five lines.
- Lines one, two and five rhyme. Lines three and four rhyme. That is AABBA.
- Lines three and four are shorter. The whole thing bounces along.

DO:
- Display the recap and read the example with a strong rhythm.
- Point to the AABBA labels and the two short middle lines.
- This is your catch-up moment for anyone who missed Session 1.

TEACHER NOTES:
Re-establish AABBA and the bouncy rhythm before drafting.

WATCH FOR:
- Students who forget lines 3 and 4 are shorter -- show the shape on the board.

[Literacy: Launch / Recap | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me. Today we WRITE a limerick with the AABBA pattern and a bouncy beat.
- The first "I can" is for everyone -- write lines one and two that rhyme.

DO:
- Choral read the LI and success criteria.
- Remind students to read aloud as they write, to hear the bounce.

TEACHER NOTES:
SC1 (all): write a rhyming opening pair. SC2 (target): a full AABBA limerick. SC3 (depth): a clever twist in the last line and a steady bounce. Exit ticket targets SC1.

WATCH FOR:
- Students who chase the rhyme and lose the beat -- reading aloud is the fix.

[Literacy: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO = `SAY:
- Watch me write a limerick. I will start the classic way: "There was a young girl from the coast".
- Line two must rhyme with coast and bounce: "Who loved to eat jam on her toast." Coast and toast rhyme.
- Lines three and four are shorter and rhyme with each other: "She ate so much jam, she turned into a clam."
- Line five rhymes with one and two again, and lands the joke: "And now she's the sea's proudest boast."
- Coast, toast, boast -- AABBA, with a silly ending.

DO:
- Display the model and read it with a big bounce.
- Show the rhyme choices live for lines two and five.
- Tap the strong beats so students feel the rhythm.

TEACHER NOTES:
Model the AABBA and the bounce together. Reading aloud with rhythm is the heart of the teaching.

WATCH FOR:
- Students who make lines 3 and 4 too long -- show they are the short, quick lines.

[Literacy: I Do | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_WEDO = `SAY:
- Let us write a limerick together. We will start "There was a..." and pick a place that is easy to rhyme.
- Line one and two rhyme and bounce. Lines three and four are short and rhyme. Line five lands the joke.
- We will read each line aloud to keep the bounce.

DO:
- Build the limerick live, reading aloud after each line.
- List rhymes for the opening word and choose ones that fit.
- Read the whole limerick with a big bounce at the end.

TEACHER NOTES:
Guided practice in AABBA and rhythm. Keep reading aloud so the beat stays bouncy.

WATCH FOR:
- Students whose line 5 does not rhyme with lines 1 and 2 -- point back to the A rhymes.

[Literacy: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_CFU = `SAY:
- Quick check on the limerick pattern. Which lines must rhyme TOGETHER in a limerick?
- Think about AABBA.
- Show me on your fingers: one for A, two for B.

DO:
- Display both options.
- Show Me Fingers.
- Scan: most should choose B.

CFU CHECKPOINT:
Technique: Show Me Fingers (1 for A, 2 for B)
Script:
- "Which lines rhyme together in a limerick?"
- Scan for: students choose B -- lines 1, 2, 5 rhyme and lines 3, 4 rhyme.
PROCEED (>=80%): Release to You Do.
PIVOT (<80%): Use the optional re-teach slide that follows -- the limerick shape and bounce.

TEACHER NOTES:
A is wrong (1, 2, 3 do not all rhyme). B is correct: lines 1, 2, 5 share one rhyme; lines 3, 4 share another. AABBA.

WATCH FOR:
- Students who pick A -- walk back through the example's rhymes.

[Literacy: CFU | VTLM 2.0: Formative Assessment]`;

const NOTES_CFU_REVEAL = `SAY:
- The answer is B. Lines 1, 2 and 5 rhyme together, and lines 3 and 4 rhyme together.
- That is the AABBA pattern that makes a limerick bounce.

DO:
- Display the reveal.
- Point to the two rhyme groups in the example.

[Literacy: CFU Reveal | VTLM 2.0: Formative Assessment]`;

const NOTES_RETEACH = `SAY:
- Let us feel the shape. The long lines -- one, two and five -- carry the main rhyme and the bounce.
- The short lines -- three and four -- are quick and rhyme with each other.
- Clap with me: da-da-DUM, da-da-DUM. Long, long, short, short, long.
- Try the opening together: "There was an old man with a beard". Now clap its bounce.

DO:
- Display the re-teach slide showing the long and short lines.
- Clap the rhythm together.
- Re-check: students clap the bounce of a given limerick line.

TEACHER NOTES:
OPTIONAL. Use only if the CFU was below 80%. Different approach: feel the long-short-short-long shape physically by clapping.

WATCH FOR:
- Students who make every line the same length -- show the two short middle lines.

[Literacy: Optional Re-teach | VTLM 2.0: Reteach with Different Approach]`;

const NOTES_YOUDO = `SAY:
- Your turn to write a limerick. Start with "There was a..." if it helps.
- Lines one, two and five rhyme. Lines three and four are short and rhyme.
- Read every line aloud as you go to keep the bounce. Make the last line funny.
- Fifteen minutes. I will come around.

DO:
- Distribute the Limerick Plan with its starter line and rhyme families.
- Circulate. Prioritise students who needed the re-teach.
- Conference: "Read me your first two lines. Do they bounce?"

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the starter line and a rhyme family on the plan. Just complete lines one and two so they rhyme and bounce.
- Extra Notes: A rhyming, bouncy opening pair meets SC1.
EXTENDING PROMPT:
- Task: Write a full AABBA limerick about a fantastical creature, with a twist in the last line. Then try a second one.
- Extra Notes: Pushes confident writers to control the whole pattern and the joke.

TEACHER NOTES:
Students keep their limerick draft to publish in Session 4. Celebrate strong bounces and funny endings.

WATCH FOR:
- Students whose rhythm wobbles -- have them read aloud and tap the beats.

[Literacy: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket. Write the FIRST TWO LINES of a limerick.
- Lines one and two must rhyme and bounce. Start with "There was a..." if you like.
- Two lines. Read them aloud. Two minutes.

DO:
- Two minutes, silent.
- As students finish, have them read the lines to a partner to check the bounce.

TEACHER NOTES:
Exit ticket targets SC1 -- a rhyming, bouncy opening pair, the foundation of the limerick.

WATCH FOR:
- Lines that rhyme but do not bounce -- note who needs more rhythm practice.

[Literacy: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Self-check. Show me on your fingers, one to five, for each "I can".
- Then read your limerick to a partner with a big bouncy voice.

DO:
- Run the fingers check for each success criterion.
- Sixty seconds partner share -- bouncy reading.
- Preview: "Next session we publish both funny poems -- our clerihew and our limerick."

TEACHER NOTES:
Note any students still wobbling on the rhythm so you can pair them well for publishing.

[Literacy: Closing | VTLM 2.0: Review and Reflect]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Poetry Week 4 Session 3 -- Drafting a Limerick";

  // SLIDE 1 -- Title
  titleSlide(
    pres,
    "Write Your Own Limerick",
    "Five bouncy lines, AABBA, with a silly twist",
    "Week 4 Session 3  |  Year 6 Literacy  |  Poetry",
    NOTES_TITLE
  );

  // SLIDE 2 -- Resources
  addResourceSlide(pres, RESOURCE_ITEMS, T, FOOTER, NOTES_RESOURCES());

  // SLIDE 3 -- Recap / Launch
  modellingSlide(
    pres,
    "Quick Recap",
    "The Limerick -- in 20 Seconds",
    "Five lines. Rhyme AABBA.\n\n- Lines 1, 2, 5 rhyme\n- Lines 3, 4 rhyme (shorter)\n- A bouncy beat\n- A funny ending\n\nMissed last session?\nThis is all you need.",
    "there WAS a young FELLow named HALL\nwho FELL in the SPRING in the FALL.\n'twould have BEEN a sad THING\nhad he DIED in the SPRING,\nbut he DIDn't—he DIED in the FALL.",
    NOTES_RECAP,
    FOOTER
  );

  // SLIDE 4 -- LI / SC
  liSlide(
    pres,
    [
      "We are learning to WRITE a limerick -- five bouncy lines in an AABBA pattern with a funny ending",
    ],
    [
      "I can write lines one and two that rhyme",
      "I can write a full AABBA limerick",
      "I can keep a bouncy beat and end with a twist",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // SLIDE 5 -- I Do
  modellingSlide(
    pres,
    "I Do",
    "Watch Me Build the Bounce",
    "My thinking:\n\n1. Start \"There was a...\"\n\n2. Lines 1, 2, 5 share one\n   rhyme: coast, toast, boast\n\n3. Lines 3, 4 are short and\n   rhyme: jam, clam\n\n4. Land the joke in line 5.",
    "There was a young girl from the coast,   (A)\nWho loved to eat jam on her toast.   (A)\nShe ate so much jam,   (B)\nShe turned into a clam,   (B)\nAnd now she's the sea's proudest boast.   (A)",
    NOTES_IDO,
    FOOTER
  );

  // SLIDE 6 -- We Do
  boardBuildSlide(
    pres,
    "We Do",
    "Build a Limerick Together -- Start \"There Was A...\"",
    "Lines 1, 2, 5 rhyme. Lines 3, 4 are short and rhyme. Read each line aloud to keep the bounce.",
    NOTES_WEDO,
    FOOTER,
    {
      promptText: "Teacher drafts each line here. Read aloud after each one to keep the bouncy beat.",
      prefilledHints: ["A", "A", "B", "B", "A"],
      badgeColor: C.SECONDARY,
    }
  );

  // SLIDE 7 + 8 -- CFU with reveal
  withReveal(
    () => cfuSlide(
      pres,
      "CFU",
      "Which Lines Rhyme Together?",
      "Show Me Fingers: 1 (A) or 2 (B)",
      "In a limerick (AABBA):\n\nA)  Lines 1, 2 and 3 all rhyme.\n\nB)  Lines 1, 2, 5 rhyme; lines 3, 4 rhyme.",
      NOTES_CFU,
      FOOTER
    ),
    (slide) => {
      addRevealAnswerBar(
        slide,
        "B -- lines 1, 2, 5 share one rhyme; lines 3, 4 share another. That is AABBA.",
        { color: C.SUCCESS, label: "Answer", showTickAndFix: false }
      );
      slide.addNotes(NOTES_CFU_REVEAL);
    }
  );

  // SLIDE 9 -- Optional Re-teach
  modellingSlide(
    pres,
    "Optional Re-teach",
    "Feel the Shape -- Long, Long, Short, Short, Long",
    "The long lines (1, 2, 5)\ncarry the main rhyme and\nthe bounce.\n\nThe short lines (3, 4) are\nquick and rhyme together.\n\nClap with me:\nda-da-DUM, da-da-DUM.",
    "Try the opening together:\n\n\"There was an old man\nwith a beard\"\n\nClap its bounce.",
    NOTES_RETEACH,
    FOOTER
  );

  // SLIDE 10 -- You Do
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "You Do", { color: C.PRIMARY, w: 1.5 });
    addTitle(s, "Write Your Limerick");

    addCard(s, 0.5, CONTENT_TOP, 9, 1.7, { strip: C.PRIMARY, fill: C.WHITE });
    s.addText("Use your Limerick Plan:", {
      x: 0.75, y: CONTENT_TOP + 0.12, w: 8.4, h: 0.32,
      fontSize: 16, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("First:   Write lines 1, 2 (they rhyme and bounce).\nNext:   Write the two short lines 3, 4 (they rhyme).\nThen:   Line 5 rhymes with 1 and 2, and lands the joke.", {
      x: 0.75, y: CONTENT_TOP + 0.52, w: 8.4, h: 1.05,
      fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    const tipY = CONTENT_TOP + 1.85;
    const tipH = SAFE_BOTTOM - tipY;
    addCard(s, 0.5, tipY, 9, tipH, { strip: C.ACCENT, fill: C.BG_CARD });
    s.addText("Read every line aloud to keep the bounce", {
      x: 0.75, y: tipY + 0.10, w: 8.4, h: 0.30,
      fontSize: 14, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("- Stuck? Use the starter line and a rhyme family on your plan.\n- Ready for more? Write about a fantastical creature with a twist, then try a second limerick.", {
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
      "Write the FIRST TWO LINES of a limerick.",
      "They must rhyme and bounce. Start with \"There was a...\" if you like.",
    ],
    NOTES_EXIT,
    FOOTER,
    { assessesSc: 1, title: "Start the Bounce" }
  );

  // SLIDE 12 -- Closing
  closingSlide(
    pres,
    {
      reflectionPrompt: "Read your limerick to a partner with a big bouncy voice.",
      scItems: [
        "I can write lines one and two that rhyme",
        "I can write a full AABBA limerick",
        "I can keep a bouncy beat and end with a twist",
      ],
      selfAssessment: {
        prompt: "Show on your fingers: 1 (need help) to 5 (could teach it)",
        options: ["1-2", "3", "4-5"],
      },
    },
    NOTES_CLOSING
  );

  // ---- PDF: Limerick Plan ---------------------------------------------------
  const pl = createPdf({ title: PLAN_RESOURCE.name });
  let py = addPdfHeader(pl, "Limerick Plan -- Five Bouncy Lines", {
    color: C.PRIMARY,
    subtitle: "5 lines, AABBA. Read each line aloud to keep the bounce.",
    lessonInfo: "Week 4 Session 3 | Year 6 Literacy | Poetry",
    showNameDate: true,
  });

  py = addTipBox(pl, "Lines 1, 2, 5 rhyme. Lines 3, 4 are short and rhyme. End with a twist.", py, { color: C.PRIMARY });

  py = addSectionHeading(pl, "Starter line -- use it or change it", py, { color: C.PRIMARY });
  py = addBodyText(pl, "Line 1:  There was a young ________ from ________,", py, { fontSize: 12 });
  py += 6;

  py = addSectionHeading(pl, "Rhyme families for line endings", py, { color: C.PRIMARY });
  py = addBodyText(pl, "-oast: coast, toast, boast, roast, most", py, { fontSize: 11 });
  py = addBodyText(pl, "-ee: tree, sea, bee, free, three, knee", py, { fontSize: 11 });
  py = addBodyText(pl, "-ain: Spain, rain, train, plain, brain, chain", py, { fontSize: 11 });
  py += 8;

  py = addSectionHeading(pl, "Line 2 -- rhymes with line 1 (A)", py, { color: C.SECONDARY });
  py = addLinedArea(pl, py, 1, { lineSpacing: 24 });
  py = addSectionHeading(pl, "Lines 3 and 4 -- short, and rhyme together (B, B)", py, { color: C.PRIMARY });
  py = addLinedArea(pl, py, 2, { lineSpacing: 22 });
  py = addSectionHeading(pl, "Line 5 -- rhymes with 1 and 2, lands the joke (A)", py, { color: C.ACCENT });
  py = addLinedArea(pl, py, 1, { lineSpacing: 24 });

  addPdfFooter(pl, "Week 4 Session 3 | Limerick Plan -- Page 1");

  pl.addPage();
  let py2 = addPdfHeader(pl, "My Limerick -- Write It Here", {
    color: C.PRIMARY,
    subtitle: "Read it aloud with a big bounce.",
    lessonInfo: "Week 4 Session 3 | Year 6 Literacy",
    showNameDate: false,
  });
  py2 = addTipBox(pl, "Write your finished limerick. Label the lines A, A, B, B, A at the side.", py2, { color: C.SECONDARY });
  py2 = addLinedArea(pl, py2, 5, { lineSpacing: 28 });
  py2 += 10;
  py2 = addSectionHeading(pl, "Challenge -- a limerick about a fantastical creature", py2, { color: C.ACCENT });
  py2 = addLinedArea(pl, py2, 5, { lineSpacing: 28 });

  addPdfFooter(pl, "Week 4 Session 3 | My Limerick -- Page 2");

  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Poetry_W4S3.pptx` }),
    writePdf(pl, PLAN_PDF_PATH),
  ]);

  console.log("PPTX written to " + `${OUT_DIR}/Poetry_W4S3.pptx`);
  console.log("Done: " + PLAN_RESOURCE.name);
}

function NOTES_RESOURCES() {
  return `SAY:
- One sheet today: your Limerick Plan. It has a starter line, rhyme families and a spot for each line.
- Keep your draft -- we publish it next session.

DO:
- Print the Limerick Plan, one per student.
- Have the Poets Toolbox nearby.

[Literacy: Resources | VTLM 2.0: Enabling Learning]`;
}

build().catch((err) => {
  console.error(err);
  process.exit(1);
});
