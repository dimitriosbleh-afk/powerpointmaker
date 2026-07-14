"use strict";

// Poetry Unit (Year 6) - Week 3 Session 2: Drafting a Rhyming Poem
// Form: rhyming poem in AABB couplets.
// Catch-up design: opens with a 20-second recap of rhyme + AABB.
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
const FOOTER = "Poetry | Week 3 Session 2 | Year 6 Literacy";
const OUT_DIR = "output/Poetry_W3S2_Draft_Rhyming";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const PLAN_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Rhyming Poem Plan",
  "Rhyming poem planning template with rhyming word families and an AABB frame."
);
const RESOURCE_ITEMS = [PLAN_RESOURCE];
const PLAN_PDF_PATH = path.join(OUT_DIR, PLAN_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Today we write a rhyming poem. We get to play with matching sounds at the ends of lines.
- The trick is to make the rhyme fit the meaning, not just grab any word that rhymes.
- If you missed last session, do not worry -- I will remind you how rhyme works.

DO:
- Display the title slide.
- Keep it playful -- rhyming is a game of sounds.

TEACHER NOTES:
The challenge is choosing rhymes that still make sense. Model that openly today.

WATCH FOR:
- Students who force a silly rhyme that breaks the meaning -- that is the main thing to coach.

[Literacy: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RECAP = `SAY:
- Quick reminder, in twenty seconds. Rhyme is when two words end with the same sound.
- When line 1 and line 2 rhyme, then line 3 and line 4 rhyme, we call that AABB.
- Listen: blow and below rhyme. Sky and high rhyme. Two matching pairs.

DO:
- Display the recap and read the example aloud, leaning into the rhymes.
- Point to the A, A, B, B labels.
- This is your catch-up moment for anyone who missed Session 1.

TEACHER NOTES:
Re-establish rhyme and the AABB pattern before drafting. Absent students rejoin here.

WATCH FOR:
- Students who already remember -- have them whisper a rhyming pair to a partner.

[Literacy: Launch / Recap | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me. Today we WRITE a rhyming poem in an AABB pattern.
- The first "I can" is for everyone -- write two lines that rhyme.

DO:
- Choral read the LI and success criteria.
- Remind students the rhyme must still make sense.

TEACHER NOTES:
SC1 (all): write a rhyming pair. SC2 (target): write an AABB poem where the rhymes make sense. SC3 (depth): add an image and keep a steady rhythm. Exit ticket targets SC1.

WATCH FOR:
- Students who chase the rhyme and lose the meaning -- this is the key coaching point.

[Literacy: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO = `SAY:
- Watch me write a rhyming poem. I think out loud so you hear my choices.
- Line one: "The autumn wind begins to blow". It ends in "blow".
- Now line two must rhyme with "blow" AND make sense. Below, snow, glow, slow... "below" fits: "It sends the dry leaves down below."
- Lines three and four make a new pair. "sky" and "high": "The trees stand bare against the sky, and wave their branches up so high."
- See? I picked rhymes that still made a picture. That is the skill.

DO:
- Display the model and read it aloud.
- Show the rhyme choice live -- list options for "blow", then choose the one that fits.
- Label the lines A, A, B, B.

TEACHER NOTES:
Make the rhyme-hunting visible. The teaching point is choosing a rhyme that keeps the meaning, not the first word that fits the sound.

WATCH FOR:
- Students who think any rhyme will do -- show how a forced rhyme can wreck the line.

[Literacy: I Do | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_WEDO = `SAY:
- Let us write a rhyming pair together about the sea.
- Give me a first line. Now, what does it end with? Let us list words that rhyme with that.
- Which rhyming word also makes sense? Let us finish line two together.

DO:
- Build a couplet live.
- For the rhyme, list three or four options and let the class pick the one that fits the meaning.
- Read the couplet aloud.

TEACHER NOTES:
Guided practice in the rhyme-and-meaning balance. Keep the rhyme list visible so students see the choice.

WATCH FOR:
- Students who pick a rhyme that does not fit -- read the line back so they hear it wobble.

[Literacy: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_CFU = `SAY:
- Quick check. Here is the first line of a couplet. Which second line rhymes AND makes sense?
- Read both options carefully.
- Show me on your fingers: one for A, two for B.

DO:
- Display the line and both options.
- Show Me Fingers.
- Scan: most should choose B.

CFU CHECKPOINT:
Technique: Show Me Fingers (1 for A, 2 for B)
Script:
- "Which line rhymes with rain AND makes sense?"
- Scan for: students choose B and can say it rhymes and fits.
PROCEED (>=80%): Release to You Do.
PIVOT (<80%): Use the optional re-teach slide that follows -- it shows how to use rhyming word families.

TEACHER NOTES:
A makes sense but does not rhyme with "rain". B rhymes (windowpane) and keeps the meaning. B is the answer.

WATCH FOR:
- Students who pick A because they like it -- remind them the task needs BOTH rhyme and sense.

[Literacy: CFU | VTLM 2.0: Formative Assessment]`;

const NOTES_CFU_REVEAL = `SAY:
- The answer is B: "It taps a song on my windowpane." It rhymes with "rain" and paints a picture.
- A made sense but did not rhyme. A good rhyming line needs both.

DO:
- Display the reveal.
- Read the couplet aloud so the rhyme lands.

[Literacy: CFU Reveal | VTLM 2.0: Formative Assessment]`;

const NOTES_RETEACH = `SAY:
- When you need a rhyme, think in word families -- groups of words that share an ending.
- The "-ain" family: rain, train, plain, chain, gain, lane, brain.
- The "-ight" family: light, night, bright, sight, flight, might.
- Pick your end word, then run through its family to find one that fits your meaning.

DO:
- Display the re-teach slide and read each word family.
- Re-check: I say "stay"; you call out words in its family (day, play, way, bay).

TEACHER NOTES:
OPTIONAL. Use only if the CFU was below 80%. Different approach: a word-family strategy that makes finding rhymes systematic.

WATCH FOR:
- Students who still pick non-rhyming words -- read the end sounds aloud together.

[Literacy: Optional Re-teach | VTLM 2.0: Reteach with Different Approach]`;

const NOTES_YOUDO = `SAY:
- Your turn to write a rhyming poem. Aim for at least one AABB pair -- four lines if you can.
- Write a line, find a rhyme that fits, then build the next pair.
- Read it aloud. Does it rhyme? Does it still make sense?
- Fifteen minutes. I will come around.

DO:
- Distribute the Rhyming Poem Plan with its word families.
- Circulate. Prioritise students who needed the re-teach.
- Conference: "Read me your pair. Does the rhyme still make a picture?"

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the given first line and the rhyming word families. Just complete the rhyme for line two.
- Extra Notes: One strong rhyming pair meets SC1.
EXTENDING PROMPT:
- Task: Write eight lines (two more pairs), or try an ABAB pattern where every other line rhymes.
- Extra Notes: Pushes confident writers into a harder rhyme scheme.

TEACHER NOTES:
Students keep their rhyming poem draft to publish in Session 4. Note rhymes that still paint a picture.

WATCH FOR:
- Students who break the rhythm by stuffing in words -- read it aloud to find the bumpy beat.

[Literacy: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket. Write a rhyming couplet -- two lines that rhyme AND make sense.
- It can be about anything. Read it back to check the rhyme. Two minutes.

DO:
- Two minutes, silent.
- As students finish, have them read the couplet to a partner.

TEACHER NOTES:
Exit ticket targets SC1 -- a rhyming pair that makes sense. Both rhyme and meaning must be present.

WATCH FOR:
- Couplets that rhyme but do not make sense -- note who needs the meaning coaching next time.

[Literacy: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Self-check. Show me on your fingers, one to five, for each "I can".
- Then read your best rhyming pair to a partner.

DO:
- Run the fingers check for each success criterion.
- Sixty seconds partner share.
- Preview: "Next session we write a tanka, counting beats like a haiku's bigger cousin."

TEACHER NOTES:
Note students who found rhyme-and-meaning hard so you can support them next.

[Literacy: Closing | VTLM 2.0: Review and Reflect]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Poetry Week 3 Session 2 -- Drafting a Rhyming Poem";

  // SLIDE 1 -- Title
  titleSlide(
    pres,
    "Write Your Own Rhyming Poem",
    "Matching sounds in an AABB pattern -- make the rhyme make sense",
    "Week 3 Session 2  |  Year 6 Literacy  |  Poetry",
    NOTES_TITLE
  );

  // SLIDE 2 -- Resources
  addResourceSlide(pres, RESOURCE_ITEMS, T, FOOTER, NOTES_RESOURCES());

  // SLIDE 3 -- Recap / Launch
  modellingSlide(
    pres,
    "Quick Recap",
    "Rhyme and AABB -- in 20 Seconds",
    "Rhyme = two words end\nwith the same sound.\n\nAABB = line 1 and 2 rhyme,\nthen line 3 and 4 rhyme.\n\nMake the rhyme make sense.\nMissed last session?\nThis is all you need.",
    "The autumn wind begins to blow,   (A)\nIt sends the dry leaves down below.   (A)\nThe trees stand bare against the sky,   (B)\nAnd wave their branches up so high.   (B)",
    NOTES_RECAP,
    FOOTER
  );

  // SLIDE 4 -- LI / SC
  liSlide(
    pres,
    [
      "We are learning to WRITE a rhyming poem in an AABB pattern, where the rhymes still make sense",
    ],
    [
      "I can write two lines that rhyme",
      "I can write an AABB poem where the rhymes make sense",
      "I can add an image and keep a steady rhythm",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // SLIDE 5 -- I Do
  modellingSlide(
    pres,
    "I Do",
    "Watch Me Find a Rhyme That Fits",
    "My thinking:\n\n1. Write line 1, ending\n   in \"blow\".\n\n2. Rhymes with blow:\n   below, snow, glow, slow.\n\n3. Which one keeps the\n   meaning? -> below.\n\nRhyme AND sense.",
    "The autumn wind begins to blow,\n\nIt sends the dry leaves down below.\n\nThe trees stand bare against the sky,\n\nAnd wave their branches up so high.",
    NOTES_IDO,
    FOOTER
  );

  // SLIDE 6 -- We Do
  boardBuildSlide(
    pres,
    "We Do",
    "Build a Rhyming Couplet Together -- The Sea",
    "Give me line 1. List words that rhyme with its last word. Pick the one that fits.",
    NOTES_WEDO,
    FOOTER,
    {
      promptText: "Teacher writes line 1, lists rhyming options, then the class chooses line 2.",
      prefilledHints: ["Line A", "rhymes with A"],
      badgeColor: C.SECONDARY,
    }
  );

  // SLIDE 7 + 8 -- CFU with reveal
  withReveal(
    () => cfuSlide(
      pres,
      "CFU",
      "Which Line Completes the Rhyme?",
      "Show Me Fingers: 1 (A) or 2 (B)",
      "Line 1:  I love to watch the falling rain,\n\nA)  It always makes me happy.\n\nB)  It taps a song on my windowpane.",
      NOTES_CFU,
      FOOTER
    ),
    (slide) => {
      addRevealAnswerBar(
        slide,
        "B -- it rhymes with \"rain\" (windowpane) AND makes sense.",
        { color: C.SUCCESS, label: "Answer", showTickAndFix: false }
      );
      slide.addNotes(NOTES_CFU_REVEAL);
    }
  );

  // SLIDE 9 -- Optional Re-teach
  modellingSlide(
    pres,
    "Optional Re-teach",
    "Use Rhyming Word Families",
    "Pick your end word.\nThen run through its family:\n\n- \"-ain\": rain, train, plain,\n   chain, gain, lane, brain\n\n- \"-ight\": light, night, bright,\n   sight, flight, might\n\nFind one that fits your meaning.",
    "Your turn:\n\nI say \"stay\".\nCall out its family:\nday, play, way, bay...",
    NOTES_RETEACH,
    FOOTER
  );

  // SLIDE 10 -- You Do
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "You Do", { color: C.PRIMARY, w: 1.5 });
    addTitle(s, "Write Your Rhyming Poem");

    addCard(s, 0.5, CONTENT_TOP, 9, 1.7, { strip: C.PRIMARY, fill: C.WHITE });
    s.addText("Use your Rhyming Poem Plan:", {
      x: 0.75, y: CONTENT_TOP + 0.12, w: 8.4, h: 0.32,
      fontSize: 16, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("First:   Write a line, then find a rhyme that fits the meaning.\nNext:   Build the next pair the same way (AABB).\nThen:   Read it aloud -- does it rhyme AND make sense?", {
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
    s.addText("- Stuck? Use the given first line and the rhyming word families on your plan.\n- Ready for more? Write eight lines, or try ABAB where every other line rhymes.", {
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
      "Write a rhyming couplet -- two lines that rhyme AND make sense.",
      "Read it back to check the rhyme.",
    ],
    NOTES_EXIT,
    FOOTER,
    { assessesSc: 1, title: "Write a Rhyming Couplet" }
  );

  // SLIDE 12 -- Closing
  closingSlide(
    pres,
    {
      reflectionPrompt: "Read your best rhyming pair to a partner.",
      scItems: [
        "I can write two lines that rhyme",
        "I can write an AABB poem where the rhymes make sense",
        "I can add an image and keep a steady rhythm",
      ],
      selfAssessment: {
        prompt: "Show on your fingers: 1 (need help) to 5 (could teach it)",
        options: ["1-2", "3", "4-5"],
      },
    },
    NOTES_CLOSING
  );

  // ---- PDF: Rhyming Poem Plan -----------------------------------------------
  const pl = createPdf({ title: PLAN_RESOURCE.name });
  let py = addPdfHeader(pl, "Rhyming Poem Plan -- AABB", {
    color: C.PRIMARY,
    subtitle: "Make the rhyme fit the meaning.",
    lessonInfo: "Week 3 Session 2 | Year 6 Literacy | Poetry",
    showNameDate: true,
  });

  py = addTipBox(pl, "Write a line, choose your end word, then find a rhyme from the same family that still makes sense.", py, { color: C.PRIMARY });

  py = addSectionHeading(pl, "Rhyming word families -- borrow if you are stuck", py, { color: C.PRIMARY });
  py = addBodyText(pl, "-ay: day, play, way, stay, bay, grey, sway", py, { fontSize: 11 });
  py = addBodyText(pl, "-ight: light, night, bright, sight, flight, might", py, { fontSize: 11 });
  py = addBodyText(pl, "-ound: round, sound, ground, found, mound, bound", py, { fontSize: 11 });
  py = addBodyText(pl, "-ee: sea, tree, free, bee, key, three, breeze", py, { fontSize: 11 });
  py += 8;

  py = addSectionHeading(pl, "Start me off -- finish this pair", py, { color: C.SECONDARY });
  py = addBodyText(pl, "Line A:  The morning sun begins to rise,", py, { fontSize: 12 });
  py = addBodyText(pl, "Line A (rhymes with rise -- skies, eyes, flies...):", py, { fontSize: 10, italic: true });
  py = addLinedArea(pl, py, 1, { lineSpacing: 24 });

  addPdfFooter(pl, "Week 3 Session 2 | Rhyming Poem Plan -- Page 1");

  pl.addPage();
  let py2 = addPdfHeader(pl, "My Rhyming Poem -- Write It Here", {
    color: C.PRIMARY,
    subtitle: "AABB. Read it aloud -- rhyme AND sense.",
    lessonInfo: "Week 3 Session 2 | Year 6 Literacy",
    showNameDate: false,
  });
  py2 = addTipBox(pl, "Write your rhyming poem. Label the lines A, A, B, B at the side.", py2, { color: C.SECONDARY });
  py2 = addLinedArea(pl, py2, 4, { lineSpacing: 28 });
  py2 += 10;
  py2 = addSectionHeading(pl, "Challenge -- four more lines, or try ABAB", py2, { color: C.ACCENT });
  py2 = addLinedArea(pl, py2, 4, { lineSpacing: 28 });

  addPdfFooter(pl, "Week 3 Session 2 | My Rhyming Poem -- Page 2");

  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Poetry_W3S2.pptx` }),
    writePdf(pl, PLAN_PDF_PATH),
  ]);

  console.log("PPTX written to " + `${OUT_DIR}/Poetry_W3S2.pptx`);
  console.log("Done: " + PLAN_RESOURCE.name);
}

function NOTES_RESOURCES() {
  return `SAY:
- One sheet today: your Rhyming Poem Plan. It has rhyming word families and a line to start you off.
- Keep your draft -- we publish it later this week.

DO:
- Print the Rhyming Poem Plan, one per student.
- Have the Poets Toolbox nearby.

[Literacy: Resources | VTLM 2.0: Enabling Learning]`;
}

build().catch((err) => {
  console.error(err);
  process.exit(1);
});
