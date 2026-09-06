"use strict";

// Poetry Unit (Year 6) - Week 5 Session 1: Reading -- Verse (beat + imagery together)
// Form: Verse (poetry that contains a beat; the supplied example rhymes ABAB).
// Curriculum spine: VC2E6LE04 -- explain how authors use sound AND imagery to create effect.
// Supplied example used exactly (including the supplied wording "I wondered lonely as a cloud").
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
  exitTicketSlide, modellingSlide, vocabSlide, quoteSlide,
} = T;

const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addLinedArea,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const SESSION_NUMBER = 1;
const FOOTER = "Poetry | Week 5 Session 1 | Year 6 Literacy";
const OUT_DIR = "output/Poetry_W5S1_Reading_Verse";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const ANALYSIS_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Verse Analysis",
  "The verse printed -- mark the beat, label the rhyme, circle the image and note its effect."
);
const RESOURCE_ITEMS = [ANALYSIS_RESOURCE];
const ANALYSIS_PDF_PATH = path.join(OUT_DIR, ANALYSIS_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- This week we pull it all together. Today's poem is a verse -- poetry with a strong, steady beat.
- We will listen for the beat and the rhyme, AND look for the images, and see how they work together.
- This is everything we have learned about sound and pictures, in one poem.

DO:
- Display the title slide.
- Have the verse ready to read with a clear rhythm.

TEACHER NOTES:
This session combines sound (beat and rhyme) with imagery, the heart of the curriculum focus. Read the verse aloud with feeling.

WATCH FOR:
- Students who notice only the rhyme -- nudge them to also feel the beat and see the picture.

[Literacy: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- One sheet today: the Verse Analysis. The verse is printed for you to mark up.
- You will tap the beat, label the rhyme, and circle the image.

DO:
- Print the analysis sheet, one per student.
- Have the Poets Toolbox handy and highlighters ready.

[Literacy: Resources | VTLM 2.0: Enabling Learning]`;

const NOTES_LAUNCH = `SAY:
- Listen to this famous verse. Feel the steady beat as I read.
- "I wondered lonely as a cloud, that floats on high o'er vales and hills, when all at once I saw a crowd, a host of golden daffodils."
- Did you feel the beat, like a gentle walking pace? And did you see the picture -- a field of golden flowers?
- Today we look at how the SOUND and the PICTURE work together.

DO:
- Read the verse aloud with a steady, gentle rhythm.
- Quick partner share: "What did you picture? Could you feel the beat?"
- Cold call two students.

TEACHER NOTES:
Supplied example, used exactly as provided. Let students feel the beat and see the image before analysing.

WATCH FOR:
- Students who picture the daffodils -- celebrate; that image is the poem's heart.

[Literacy: Launch | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me. Today we explain how a poet uses BEAT, RHYME and IMAGES together to create a feeling.
- The first "I can" is for everyone -- clap the beat and find the rhyming words.

DO:
- Choral read the LI and success criteria.
- Note: today we connect everything -- sound and pictures.

TEACHER NOTES:
SC1 (all): clap the beat and find rhymes. SC2 (target): spot an image (a simile). SC3 (depth): explain the feeling the sound and image create. Exit ticket targets SC2 and SC3.

WATCH FOR:
- Students who can find features but not explain the effect -- the re-teach supports the "feeling" step.

[Literacy: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_VOCAB = `SAY:
- One word for today: rhythm.
- Rhythm is the beat you feel when you read a poem aloud -- like a heartbeat or footsteps.
- When a poem has a strong, steady rhythm, we call it verse.

DO:
- Choral say "rhythm".
- Tap the beat of the verse together so students feel it in their bodies.

TEACHER NOTES:
Rhythm is the sound focus today. Pair it with imagery so students see how the two combine.

WATCH FOR:
- Students who confuse rhythm with rhyme -- rhythm is the BEAT; rhyme is matching SOUNDS.

[Literacy: Vocabulary | VTLM 2.0: Vocabulary Instruction]`;

const NOTES_IDO = `SAY:
- Let us look closely at the verse. Watch me notice three things working together.
- First, the BEAT. Tap with me -- da-DUM, da-DUM. A steady, gentle pace, like walking.
- Second, the RHYME. Cloud and crowd rhyme; hills and daffodils rhyme. That is ABAB.
- Third, the IMAGE. "Lonely as a cloud" -- a simile. And "a host of golden daffodils" -- a bright picture.
- Together, the gentle beat and the golden image make the poem feel calm and full of wonder.

DO:
- Display the verse and tap the beat.
- Underline the rhyming line-ends and label ABAB.
- Circle "lonely as a cloud" and "golden daffodils".

TEACHER NOTES:
The key move is showing how sound and imagery TOGETHER create the feeling. Name all three: beat, rhyme, image.

WATCH FOR:
- Students who treat the features separately -- keep linking them to the overall feeling.

[Literacy: I Do | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_WEDO = `SAY:
- Your turn to help with the sound. Let us name the rhyme scheme and feel the beat together.
- Read the last word of each line. Which lines rhyme? Lines 1 and 3, then lines 2 and 4.
- Tap the beat with me. Is it steady?
- Tell your partner the rhyme scheme, then we will check.

DO:
- Display the verse again.
- Read each line-end word and tap the beat.
- Cold call for the rhyme scheme, then reveal ABAB.

TEACHER NOTES:
Guided practice on the sound layer. cloud/crowd (lines 1, 3) and hills/daffodils (lines 2, 4) make ABAB.

WATCH FOR:
- Students who expect AABB -- show that here the rhymes alternate, ABAB.

[Literacy: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_REVEAL = `SAY:
- The rhyme scheme is ABAB -- the rhymes alternate. Cloud and crowd; hills and daffodils.
- And the beat is steady and gentle. That calm beat suits the peaceful scene.

DO:
- Display the reveal.
- Tap the steady beat once more.

[Literacy: We Do Reveal | VTLM 2.0: Guided Practice]`;

const NOTES_CFU = `SAY:
- Quick check on the image. The poet writes "lonely as a cloud".
- Is that a simile or a metaphor?
- Show me on your fingers: one for simile, two for metaphor.

DO:
- Display the phrase.
- Show Me Fingers.
- Scan: most should choose simile (1).

CFU CHECKPOINT:
Technique: Show Me Fingers (1 simile, 2 metaphor)
Script:
- "Is 'lonely as a cloud' a simile or a metaphor?"
- Scan for: students choose simile and point to the word "as".
PROCEED (>=80%): Release to You Do.
PIVOT (<80%): Use the optional re-teach slide that follows -- it shows how sound and imagery build a feeling.

TEACHER NOTES:
It is a simile -- it compares using "as". The comparison makes the speaker feel small and drifting, which adds to the calm, wondering mood.

WATCH FOR:
- Students who say metaphor -- remind them: "as" or "like" means a simile.

[Literacy: CFU | VTLM 2.0: Formative Assessment]`;

const NOTES_CFU_REVEAL = `SAY:
- It is a simile -- "lonely as a cloud" compares using the word "as".
- That picture of a drifting cloud makes the poem feel calm and full of quiet wonder.

DO:
- Display the reveal.
- Link the image back to the feeling.

[Literacy: CFU Reveal | VTLM 2.0: Formative Assessment]`;

const NOTES_RETEACH = `SAY:
- Let us see how sound and image build a feeling, step by step.
- The BEAT is gentle and steady -- that already feels calm, not rushed.
- The RHYME is soft and regular -- it feels settled, not jagged.
- The IMAGE is a drifting cloud and golden flowers -- peaceful and bright.
- Put them together and the poem feels calm and full of wonder. That is the author's effect.

DO:
- Display the re-teach slide showing the three layers building one feeling.
- Re-check: ask students to name the feeling and point to ONE thing that creates it.

TEACHER NOTES:
OPTIONAL. Use only if the CFU was below 80%. Different approach: stack the three layers (beat, rhyme, image) to show how they combine into a feeling.

WATCH FOR:
- Students who name a feeling but no evidence -- ask "which part of the poem makes you feel that?"

[Literacy: Optional Re-teach | VTLM 2.0: Reteach with Different Approach]`;

const NOTES_YOUDO = `SAY:
- Your turn to be a poetry detective one last time. The verse is on your sheet.
- Tap the beat and mark the strong beats. Label the rhyme scheme. Circle the image.
- Then write ONE sentence: what feeling does the poem create, and what helps make it?
- Work quietly first, then compare with a partner.

DO:
- Distribute the analysis sheet.
- Circulate. Prioritise students who needed the re-teach.
- Conference: "What feeling? Now show me one thing in the poem that makes it."

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Just circle the image and tell a partner what you picture and how it feels.
- Extra Notes: Naming the image and a feeling meets SC2.
EXTENDING PROMPT:
- Task: Explain how TWO different things -- the beat and the image -- work together to create the feeling.
- Extra Notes: This reaches the depth criterion: sound and imagery creating effect.

TEACHER NOTES:
Students keep the marked-up sheet. This is the analysis skill that underpins their own writing this week.

WATCH FOR:
- Students who describe features but not the effect -- push them to the feeling.

[Literacy: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket. Find the simile in the verse and write it down.
- Then write the feeling it creates, and why.
- Two minutes.

DO:
- Two minutes, silent.
- Collect.

TEACHER NOTES:
Exit ticket targets SC2 and SC3 -- naming the image and explaining its effect, the core of this week's reading goal.

WATCH FOR:
- Students who name the image but not the feeling -- the next session's writing builds this further.

[Literacy: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Self-check. Show me on your fingers, one to five, for each "I can".
- Then tell a partner: which part of the verse did you like best -- the beat, the rhyme or the picture?

DO:
- Run the fingers check for each success criterion.
- Sixty seconds partner share.
- Preview: "Next session we write our own verse -- with a beat and a picture."

TEACHER NOTES:
This reading analysis sets up the final piece of original writing in the unit.

[Literacy: Closing | VTLM 2.0: Review and Reflect]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Poetry Week 5 Session 1 -- Reading: Verse";

  // SLIDE 1 -- Title
  titleSlide(
    pres,
    "When Sound and Picture Meet",
    "Reading Verse -- Beat, Rhyme and Imagery Together",
    "Week 5 Session 1  |  Year 6 Literacy  |  Poetry",
    NOTES_TITLE
  );

  // SLIDE 2 -- Resources
  addResourceSlide(pres, RESOURCE_ITEMS, T, FOOTER, NOTES_RESOURCES);

  // SLIDE 3 -- Launch (read the verse)
  quoteSlide(
    pres,
    "Read Aloud",
    "A Verse",
    "I wondered lonely as a cloud\nThat floats on high o'er vales and hills,\nWhen all at once I saw a crowd,\nA host of golden daffodils.",
    "",
    "What did you picture? Could you feel the steady beat?",
    NOTES_LAUNCH,
    FOOTER
  );

  // SLIDE 4 -- LI / SC
  liSlide(
    pres,
    [
      "We are learning to explain how a poet uses BEAT, RHYME and IMAGES together to create a feeling",
    ],
    [
      "I can clap the beat and find the rhyming words",
      "I can spot an image (a simile) the poet uses",
      "I can explain the feeling the sound and image create",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // SLIDE 5 -- Vocab: rhythm
  vocabSlide(
    pres,
    "rhythm",
    "noun",
    "The beat you feel when you read a poem aloud -- like a heartbeat or footsteps. A poem with a strong, steady beat is called verse.",
    "I wondered lonely as a cloud, that floats on high o'er vales and hills.",
    NOTES_VOCAB,
    FOOTER,
    { routine: "Say it together. Then tap the beat of the verse with me." }
  );

  // SLIDE 6 -- I Do (three layers: beat, rhyme, image)
  modellingSlide(
    pres,
    "I Do",
    "Three Things Working Together",
    "Notice three things:\n\n1. BEAT -- tap it.\n   Steady, gentle, like walking.\n\n2. RHYME -- cloud / crowd,\n   hills / daffodils. ABAB.\n\n3. IMAGE -- \"lonely as a\n   cloud\" (simile), golden\n   daffodils.",
    "I wondered lonely as a cloud   (A)\n\nThat floats on high o'er vales and hills,   (B)\n\nWhen all at once I saw a crowd,   (A)\n\nA host of golden daffodils.   (B)",
    NOTES_IDO,
    FOOTER
  );

  // SLIDE 7 + 8 -- We Do (rhyme scheme + beat) with reveal
  withReveal(
    () => cfuSlide(
      pres,
      "We Do",
      "Name the Rhyme Scheme, Feel the Beat",
      "Which lines rhyme? Is the beat steady?",
      "Read the last word of each line.\nLines 1 and 3: cloud / crowd.\nLines 2 and 4: hills / daffodils.\n\nWhat is the rhyme scheme?",
      NOTES_WEDO,
      FOOTER
    ),
    (slide) => {
      addRevealAnswerBar(
        slide,
        "ABAB -- the rhymes alternate. The steady, gentle beat suits the calm scene.",
        { color: C.SUCCESS, label: "Answer", showTickAndFix: false }
      );
      slide.addNotes(NOTES_WEDO_REVEAL);
    }
  );

  // SLIDE 9 + 10 -- CFU (simile vs metaphor) with reveal
  withReveal(
    () => cfuSlide(
      pres,
      "CFU",
      "Simile or Metaphor?",
      "Show Me Fingers: 1 (simile) or 2 (metaphor)",
      "The poet writes:\n\n\"lonely as a cloud\"\n\nIs that a simile or a metaphor?",
      NOTES_CFU,
      FOOTER
    ),
    (slide) => {
      addRevealAnswerBar(
        slide,
        "Simile -- it compares using \"as\". The drifting cloud adds to the calm, wondering feeling.",
        { color: C.SUCCESS, label: "Answer", showTickAndFix: false }
      );
      slide.addNotes(NOTES_CFU_REVEAL);
    }
  );

  // SLIDE 11 -- Optional Re-teach
  modellingSlide(
    pres,
    "Optional Re-teach",
    "How the Layers Build a Feeling",
    "Stack the three layers:\n\n- BEAT: gentle, steady\n   -> feels calm\n\n- RHYME: soft, regular\n   -> feels settled\n\n- IMAGE: drifting cloud,\n   golden flowers\n   -> feels peaceful, bright",
    "Together they make the\npoem feel:\n\nCALM and full of WONDER.\n\nThat is the author's effect.",
    NOTES_RETEACH,
    FOOTER
  );

  // SLIDE 12 -- You Do
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "You Do", { color: C.PRIMARY, w: 1.5 });
    addTitle(s, "Poetry Detective -- Sound and Picture");

    addCard(s, 0.5, CONTENT_TOP, 9, 1.7, { strip: C.PRIMARY, fill: C.WHITE });
    s.addText("On your Verse Analysis sheet:", {
      x: 0.75, y: CONTENT_TOP + 0.12, w: 8.4, h: 0.32,
      fontSize: 16, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("First:   Tap the beat and mark the strong beats. Label the rhyme scheme.\nNext:   Circle the image (the simile).\nThen:   Write ONE sentence: what feeling does it create, and what helps make it?", {
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
    s.addText("- Just circling the image and naming a feeling is a great start.\n- Ready for more? Explain how the beat AND the image work together to create the feeling.", {
      x: 0.75, y: tipY + 0.44, w: 8.4, h: tipH - 0.55,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  }

  // SLIDE 13 -- Exit Ticket
  exitTicketSlide(
    pres,
    [
      "Find the simile in the verse and write it down.",
      "Then write the feeling it creates, and why.",
    ],
    NOTES_EXIT,
    FOOTER,
    { assessesSc: 3, title: "The Image and Its Feeling" }
  );

  // SLIDE 14 -- Closing
  closingSlide(
    pres,
    {
      reflectionPrompt: "Which part of the verse did you like best -- the beat, the rhyme or the picture?",
      scItems: [
        "I can clap the beat and find the rhyming words",
        "I can spot an image (a simile) the poet uses",
        "I can explain the feeling the sound and image create",
      ],
      selfAssessment: {
        prompt: "Show on your fingers: 1 (need help) to 5 (could teach it)",
        options: ["1-2", "3", "4-5"],
      },
    },
    NOTES_CLOSING
  );

  // ---- PDF: Verse Analysis --------------------------------------------------
  const an = createPdf({ title: ANALYSIS_RESOURCE.name });
  let ay = addPdfHeader(an, "Verse Analysis -- Sound and Picture", {
    color: C.PRIMARY,
    subtitle: "Mark the beat. Label the rhyme. Circle the image. Name the feeling.",
    lessonInfo: "Week 5 Session 1 | Year 6 Literacy | Poetry",
    showNameDate: true,
  });

  ay = addTipBox(an, "Read the verse aloud first and tap the beat. Then mark it up like a poetry detective.", ay, { color: C.PRIMARY });

  ay = addSectionHeading(an, "The Verse", ay, { color: C.PRIMARY });
  ay = addBodyText(an, "I wondered lonely as a cloud", ay, { fontSize: 14 });
  ay = addBodyText(an, "That floats on high o'er vales and hills,", ay, { fontSize: 14 });
  ay = addBodyText(an, "When all at once I saw a crowd,", ay, { fontSize: 14 });
  ay = addBodyText(an, "A host of golden daffodils.", ay, { fontSize: 14 });
  ay += 8;
  ay = addBodyText(an, "1. Tap the beat and mark the strong beats above the words.", ay, { fontSize: 10, italic: true });
  ay = addBodyText(an, "2. Label the rhyme scheme at the side (it alternates -- ABAB).", ay, { fontSize: 10, italic: true });
  ay = addBodyText(an, "3. Circle the image (the simile).", ay, { fontSize: 10, italic: true });
  ay += 10;

  ay = addSectionHeading(an, "What feeling does the verse create? What helps make it?", ay, { color: C.SECONDARY });
  ay = addLinedArea(an, ay, 4, { lineSpacing: 24 });

  addPdfFooter(an, "Week 5 Session 1 | Verse Analysis -- Page 1");

  an.addPage();
  let ay2 = addPdfHeader(an, "Image and Feeling -- Exit Ticket", {
    color: C.ACCENT,
    subtitle: "Sound and imagery create the feeling.",
    lessonInfo: "Week 5 Session 1 | Year 6 Literacy",
    showNameDate: false,
  });
  ay2 = addTipBox(an, "Find the image, then explain the feeling it creates.", ay2, { color: C.ACCENT });
  ay2 = addBodyText(an, "The simile in the verse is:", ay2, { fontSize: 11 });
  ay2 = addLinedArea(an, ay2, 1, { lineSpacing: 24 });
  ay2 += 6;
  ay2 = addBodyText(an, "The feeling it creates, and why:", ay2, { fontSize: 11 });
  ay2 = addLinedArea(an, ay2, 4, { lineSpacing: 24 });
  addPdfFooter(an, "Week 5 Session 1 | Exit Ticket");

  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Poetry_W5S1.pptx` }),
    writePdf(an, ANALYSIS_PDF_PATH),
  ]);

  console.log("PPTX written to " + `${OUT_DIR}/Poetry_W5S1.pptx`);
  console.log("Done: " + ANALYSIS_RESOURCE.name);
}

build().catch((err) => {
  console.error(err);
  process.exit(1);
});
