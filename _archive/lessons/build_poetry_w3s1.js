"use strict";

// Poetry Unit (Year 6) - Week 3 Session 1: Reading -- Rhyme and Rhythm (Rhyming poem + Tanka)
// Forms introduced: Rhyming poem (AABB couplets) + Tanka (5-7-5-7-7, unrhymed).
// Curriculum spine: VCELT344 -- relationship between words, sounds and imagery.
// Supplied examples used exactly. Catch-up: each form is fully taught in its own session.
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
const FOOTER = "Poetry | Week 3 Session 1 | Year 6 Literacy";
const OUT_DIR = "output/Poetry_W3S1_Reading_Rhyme_Tanka";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const ANALYSIS_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Rhyme and Tanka Analysis",
  "Two poems printed -- mark the rhyme pairs on one, count the tanka's beats on the other."
);
const RESOURCE_ITEMS = [ANALYSIS_RESOURCE];
const ANALYSIS_PDF_PATH = path.join(OUT_DIR, ANALYSIS_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- This week we look at how poems use SOUND. Two shapes today: a rhyming poem and a tanka.
- A rhyming poem plays with matching end-sounds. A tanka counts beats, like a haiku's bigger cousin.
- We are readers today -- we will write each one later in the week.

DO:
- Display the title slide.
- Have both example poems ready to read aloud.

TEACHER NOTES:
Today analyses two contrasting sound shapes: one rhymes, one counts beats and does not rhyme. Keep that contrast clear.

WATCH FOR:
- Students who assume all poems rhyme -- the tanka is a good surprise.

[Literacy: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- One sheet today: the Rhyme and Tanka Analysis. Both poems are printed for you to mark up.
- You will circle rhyming pairs on one, and count beats on the other.

DO:
- Print the analysis sheet, one per student.
- Have the Poets Toolbox handy and highlighters ready.

[Literacy: Resources | VTLM 2.0: Enabling Learning]`;

const NOTES_LAUNCH = `SAY:
- Listen to this poem. Notice the end of each line -- do any words sound alike?
- "The freshening feel of an ocean breeze, the colours of change in the leaves on the trees..."
- Breeze and trees -- hear how they chime? That matching sound is rhyme.
- Let us look closely at how this poet uses it.

DO:
- Read the rhyming poem aloud, leaning into the rhymes.
- Quick partner share: "Which words rhymed? How did the rhyme make it feel?"
- Cold call two students.

TEACHER NOTES:
Supplied example, used exactly. Let students hear the rhyme before naming the pattern (AABB).

WATCH FOR:
- Students who catch breeze/trees and by/why -- celebrate, that is the rhyme scheme.

[Literacy: Launch | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me. Today we notice how poems use sound -- rhyme and beat.
- The first "I can" is for everyone -- find two words that rhyme.

DO:
- Choral read the LI and success criteria.
- Note the two shapes: one rhymes, one counts beats.

TEACHER NOTES:
SC1 (all): find rhyming words. SC2 (target): name a poem's rhyme pattern (AABB) and count a tanka's beats. SC3 (depth): explain how the sound shapes the feeling. Exit ticket targets SC1 and SC2.

WATCH FOR:
- Students who mix up the two forms -- the CFU later checks they can tell them apart.

[Literacy: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_VOCAB = `SAY:
- One word for today: rhyme.
- Rhyme is when two words end with the same sound, like cat and hat, or breeze and trees.
- Poets often rhyme the ends of lines. When line 1 and line 2 rhyme, we label them A and A.

DO:
- Choral say "rhyme".
- Quick game: I say a word, you call out one that rhymes. Try "light", "day", "round".

TEACHER NOTES:
Rhyme is the key sound device this week. The A/A/B/B labelling makes the pattern visible.

WATCH FOR:
- Students who offer words that start the same instead of end the same -- redirect to the END sound.

[Literacy: Vocabulary | VTLM 2.0: Vocabulary Instruction]`;

const NOTES_IDO = `SAY:
- Let us look closely at the rhyming poem. Watch how I find its pattern.
- I read the LAST word of each line. Line 1 ends in "breeze". Line 2 ends in "trees". They rhyme.
- Line 3 ends in "by". Line 4 ends in "why". They rhyme too.
- So the first pair matches, then the second pair matches. We call that AABB.

DO:
- Display the poem and underline the last word of each line.
- Draw the A, A, B, B labels down the side.
- Read it once more so students hear the chiming pairs.

TEACHER NOTES:
The skill is reading line-end words and hearing the match. Make the AABB labels visible on the board.

WATCH FOR:
- Students who think every line must rhyme with every other -- show that it is pairs.

[Literacy: I Do | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_WEDO = `SAY:
- Your turn to help. Here is a new four-line poem. Let us find its rhyme pattern together.
- Read the last word of each line with me. Which words rhyme?
- Tell your partner the pattern, then we will check.

DO:
- Display the second rhyming poem.
- Read each line-end word aloud with the class.
- Cold call for the rhyming pairs, then reveal the AABB labels.

TEACHER NOTES:
Guided practice in spotting a rhyme scheme. hill/still and call/all are the pairs.

WATCH FOR:
- Students who find one pair but miss the other -- prompt them to check lines 3 and 4.

[Literacy: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_REVEAL = `SAY:
- The pairs are hill and still, then call and all.
- So this poem is also AABB -- two matching pairs.

DO:
- Display the reveal.
- Point to each rhyming pair.

[Literacy: We Do Reveal | VTLM 2.0: Guided Practice]`;

const NOTES_TANKA = `SAY:
- Now a completely different shape: a tanka. It does NOT rhyme. Instead it counts beats.
- A tanka has five lines: five, seven, five, seven, seven beats. Thirty-one beats in all.
- Listen and clap with me. "Crash at two A.M." -- five beats.
- It is like a haiku with two extra lines, and it often tells a tiny story.

DO:
- Display the tanka and read it aloud.
- Clap each line, writing the counts 5, 7, 5, 7, 7 beside them.
- Point out there is no rhyme -- the music comes from the beats.

TEACHER NOTES:
Supplied example, used exactly. The contrast with the rhyming poem is the teaching point: two ways to make sound-music.

WATCH FOR:
- Students who try to find rhymes in the tanka -- there are none; the pattern is beats.

[Literacy: I Do / We Do | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_CFU = `SAY:
- Quick check. Which statement describes a TANKA?
- Read both. Think about what we just learned.
- Show me on your fingers: one for A, two for B.

DO:
- Display both options.
- Show Me Fingers.
- Scan: most should choose B.

CFU CHECKPOINT:
Technique: Show Me Fingers (1 for A, 2 for B)
Script:
- "Which describes a tanka?"
- Scan for: students choose B and can say a tanka counts beats, not rhymes.
PROCEED (>=80%): Release to You Do.
PIVOT (<80%): Use the optional re-teach slide that follows -- it lines up rhyme and beat side by side.

TEACHER NOTES:
A describes a rhyming poem. B describes a tanka (5-7-5-7-7, unrhymed). Students must separate the two forms before drafting them later this week.

WATCH FOR:
- Students who pick A -- remind them: the tanka counts beats and does not rhyme.

[Literacy: CFU | VTLM 2.0: Formative Assessment]`;

const NOTES_CFU_REVEAL = `SAY:
- A tanka counts beats and does not rhyme, so the answer is B.
- A described a rhyming poem -- matching end-sounds in an AABB pattern.

DO:
- Display the reveal.
- Restate the two shapes once more.

[Literacy: CFU Reveal | VTLM 2.0: Formative Assessment]`;

const NOTES_RETEACH = `SAY:
- Let us sort the two ideas. Rhyme is about MATCHING END-SOUNDS. Beat is about COUNTING SYLLABLES.
- Rhyming poem: cat / hat -- the ends match. We do not count beats.
- Tanka: five, seven, five, seven, seven -- we count beats, and the ends do NOT need to match.
- Two different kinds of music. Which one are you listening for?

DO:
- Display the re-teach slide with the two columns.
- Re-check: read a couplet -- do these rhyme? Then read a tanka line -- how many beats?

TEACHER NOTES:
OPTIONAL. Use only if the CFU was below 80%. Different approach: a two-column sort that separates "matching sounds" from "counting beats".

WATCH FOR:
- Students who still blend them -- anchor on "rhyme = ends match, tanka = beats counted".

[Literacy: Optional Re-teach | VTLM 2.0: Reteach with Different Approach]`;

const NOTES_YOUDO = `SAY:
- Your turn to be a poetry detective. Both poems are on your sheet.
- On the rhyming poem, circle the rhyming pairs and label them A, A, B, B.
- On the tanka, clap each line and write the number of beats. Check it is 5, 7, 5, 7, 7.
- Work quietly first, then compare with a partner.

DO:
- Distribute the analysis sheet.
- Circulate. Prioritise students who needed the re-teach.
- Conference: "Read me your rhyming pair. Now clap me a tanka line."

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Just circle the two rhyming pairs on the rhyming poem and read them to a partner.
- Extra Notes: Finding one pair meets SC1; the rest can come with support.
EXTENDING PROMPT:
- Task: Explain how the rhyme or the beat changes the FEELING of each poem -- which sounds gentle, which sounds surprising?
- Extra Notes: This reaches the depth criterion about sound shaping meaning.

TEACHER NOTES:
Students keep the marked-up sheet -- it feeds the drafting sessions this week. Capture good noticings to share.

WATCH FOR:
- Students who count beats on the rhyming poem -- redirect each form to its own kind of music.

[Literacy: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket. Two quick tasks.
- One: write a word that rhymes with "night", and one that rhymes with "blue".
- Two: how many beats are in this tanka line: "A white cat ran by"?
- Two minutes.

DO:
- Two minutes, silent.
- Collect.

TEACHER NOTES:
Exit ticket targets SC1 (rhyme) and SC2 (beat count -- the line has five beats). Shows students can use both kinds of sound.

WATCH FOR:
- Students who rhyme well but miscount, or vice versa -- note which skill to support before drafting.

[Literacy: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Self-check. Show me on your fingers, one to five, for each "I can".
- Then tell a partner: which poem's sound did you like better -- the rhyme or the beat -- and why?

DO:
- Run the fingers check for each success criterion.
- Sixty seconds partner share.
- Preview: "Next session we write our own rhyming poem."

TEACHER NOTES:
Note which sound device students preferred and which they found harder, to plan support.

[Literacy: Closing | VTLM 2.0: Review and Reflect]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Poetry Week 3 Session 1 -- Reading: Rhyme and Rhythm";

  // SLIDE 1 -- Title
  titleSlide(
    pres,
    "The Music of Poems",
    "Reading for Sound -- Rhyming Poems and Tanka",
    "Week 3 Session 1  |  Year 6 Literacy  |  Poetry",
    NOTES_TITLE
  );

  // SLIDE 2 -- Resources
  addResourceSlide(pres, RESOURCE_ITEMS, T, FOOTER, NOTES_RESOURCES);

  // SLIDE 3 -- Launch (read the rhyming poem)
  quoteSlide(
    pres,
    "Read Aloud",
    "A Rhyming Poem",
    "The freshening feel of an ocean breeze,\nThe colours of change in the leaves on the trees,\nThe feeling of peace as the days go by,\nLife's a dazzling puzzle and we don't know why.",
    "",
    "Which words rhymed? How did the rhyme make the poem feel?",
    NOTES_LAUNCH,
    FOOTER
  );

  // SLIDE 4 -- LI / SC
  liSlide(
    pres,
    [
      "We are learning to notice how poems use SOUND -- rhyme and beat -- to make their music",
    ],
    [
      "I can find two words that rhyme",
      "I can name a poem's rhyme pattern (AABB) and count a tanka's beats",
      "I can explain how the sound shapes the feeling of a poem",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // SLIDE 5 -- Vocab: rhyme
  vocabSlide(
    pres,
    "rhyme",
    "noun / verb",
    "When two words end with the same sound, like cat and hat, or breeze and trees.",
    "The freshening feel of an ocean breeze, the colours of change in the leaves on the trees.",
    NOTES_VOCAB,
    FOOTER,
    { routine: "Say it together. I say a word, you call one that rhymes." }
  );

  // SLIDE 6 -- I Do (rhyme scheme of the poem)
  modellingSlide(
    pres,
    "I Do",
    "Find the Rhyme Pattern",
    "Read the LAST word of each line:\n\n- breeze\n- trees   } these rhyme (A, A)\n\n- by\n- why   } these rhyme (B, B)\n\nTwo matching pairs.\nThe pattern is AABB.",
    "...ocean breeze,   (A)\n\n...on the trees,   (A)\n\n...the days go by,   (B)\n\n...we don't know why.   (B)",
    NOTES_IDO,
    FOOTER
  );

  // SLIDE 7 + 8 -- We Do (find the scheme of a new poem) with reveal
  withReveal(
    () => modellingSlide(
      pres,
      "We Do",
      "Together: Find the Rhyme Pattern",
      "Read the last word of\neach line.\n\nWhich words rhyme?\n\nWhat is the pattern --\nis it AABB?",
      "The moon climbs high above the hill,\n\nThe night is calm, the world is still,\n\nA single owl begins to call,\n\nThen silence settles over all.",
      NOTES_WEDO,
      FOOTER
    ),
    (slide) => {
      addRevealAnswerBar(
        slide,
        "Pairs: hill / still, then call / all. Pattern: AABB.",
        { color: C.SUCCESS, label: "Answer", showTickAndFix: false }
      );
      slide.addNotes(NOTES_WEDO_REVEAL);
    }
  );

  // SLIDE 9 -- I Do / We Do: the tanka shape
  modellingSlide(
    pres,
    "Notice",
    "A Tanka -- Counts Beats, Does Not Rhyme",
    "Five lines. Count the beats:\n\n- Line 1 = 5\n- Line 2 = 7\n- Line 3 = 5\n- Line 4 = 7\n- Line 5 = 7\n\n31 beats. No rhyme.\nOften tells a tiny story.",
    "Crash at two A.M.\n(5)\nI opened my bedroom door\n(7)\nA white cat ran by\n(5)\nStartled by the clanging fall\n(7)\nOf the treat jar's metal lid\n(7)",
    NOTES_TANKA,
    FOOTER
  );

  // SLIDE 10 + 11 -- CFU (tell the forms apart) with reveal
  withReveal(
    () => cfuSlide(
      pres,
      "CFU",
      "Which Describes a Tanka?",
      "Show Me Fingers: 1 (A) or 2 (B)",
      "A)  It rhymes in an AABB pattern.\n\nB)  It counts 5-7-5-7-7 beats and does not rhyme.",
      NOTES_CFU,
      FOOTER
    ),
    (slide) => {
      addRevealAnswerBar(
        slide,
        "B describes a tanka -- it counts beats (5-7-5-7-7) and does not rhyme.",
        { color: C.SUCCESS, label: "Answer", showTickAndFix: false }
      );
      slide.addNotes(NOTES_CFU_REVEAL);
    }
  );

  // SLIDE 12 -- Optional Re-teach
  modellingSlide(
    pres,
    "Optional Re-teach",
    "Two Kinds of Music -- Sort Them",
    "RHYME = matching end-sounds\n\n- cat / hat\n- bright / night\n- We do NOT count beats.\n\nTANKA = counting beats\n\n- 5 - 7 - 5 - 7 - 7\n- The ends do NOT\n   need to match.",
    "Rhyme question:\nDo \"hill\" and \"still\" rhyme?\n\nTanka question:\nHow many beats in\n\"A white cat ran by\"?",
    NOTES_RETEACH,
    FOOTER
  );

  // SLIDE 13 -- You Do
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "You Do", { color: C.PRIMARY, w: 1.5 });
    addTitle(s, "Poetry Detective -- Sound");

    addCard(s, 0.5, CONTENT_TOP, 9, 1.7, { strip: C.PRIMARY, fill: C.WHITE });
    s.addText("On your Analysis sheet:", {
      x: 0.75, y: CONTENT_TOP + 0.12, w: 8.4, h: 0.32,
      fontSize: 16, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("First:   On the rhyming poem, circle the rhyming pairs and label them A, A, B, B.\nNext:   On the tanka, clap each line and write the beats.\nThen:   Check the tanka is 5, 7, 5, 7, 7.", {
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
    s.addText("- Just finding the two rhyming pairs is a great start.\n- Ready for more? Explain how the rhyme or the beat changes each poem's feeling.", {
      x: 0.75, y: tipY + 0.44, w: 8.4, h: tipH - 0.55,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  }

  // SLIDE 14 -- Exit Ticket
  exitTicketSlide(
    pres,
    [
      "Write a word that rhymes with night, and one that rhymes with blue.",
      "How many beats are in this tanka line: A white cat ran by?",
    ],
    NOTES_EXIT,
    FOOTER,
    { assessesSc: 2, title: "Rhyme and Beat Check" }
  );

  // SLIDE 15 -- Closing
  closingSlide(
    pres,
    {
      reflectionPrompt: "Which poem's sound did you like better -- the rhyme or the beat -- and why?",
      scItems: [
        "I can find two words that rhyme",
        "I can name a poem's rhyme pattern (AABB) and count a tanka's beats",
        "I can explain how the sound shapes the feeling of a poem",
      ],
      selfAssessment: {
        prompt: "Show on your fingers: 1 (need help) to 5 (could teach it)",
        options: ["1-2", "3", "4-5"],
      },
    },
    NOTES_CLOSING
  );

  // ---- PDF: Rhyme and Tanka Analysis ----------------------------------------
  const an = createPdf({ title: ANALYSIS_RESOURCE.name });
  let ay = addPdfHeader(an, "Rhyme and Tanka Analysis", {
    color: C.PRIMARY,
    subtitle: "Mark the rhyme pairs on one poem. Count the beats on the other.",
    lessonInfo: "Week 3 Session 1 | Year 6 Literacy | Poetry",
    showNameDate: true,
  });

  ay = addTipBox(an, "Read each poem aloud first, then mark it up like a poetry detective.", ay, { color: C.PRIMARY });

  ay = addSectionHeading(an, "Poem 1 -- A Rhyming Poem", ay, { color: C.PRIMARY });
  ay = addBodyText(an, "The freshening feel of an ocean breeze,", ay, { fontSize: 13 });
  ay = addBodyText(an, "The colours of change in the leaves on the trees,", ay, { fontSize: 13 });
  ay = addBodyText(an, "The feeling of peace as the days go by,", ay, { fontSize: 13 });
  ay = addBodyText(an, "Life's a dazzling puzzle and we don't know why.", ay, { fontSize: 13 });
  ay += 6;
  ay = addBodyText(an, "Circle the rhyming pairs. Label each line A, A, B or B.", ay, { fontSize: 10, italic: true });
  ay = addLinedArea(an, ay, 1, { lineSpacing: 22 });
  ay += 6;

  ay = addSectionHeading(an, "Poem 2 -- A Tanka (5 - 7 - 5 - 7 - 7)", ay, { color: C.SECONDARY });
  ay = addBodyText(an, "Crash at two A.M.", ay, { fontSize: 13 });
  ay = addBodyText(an, "I opened my bedroom door", ay, { fontSize: 13 });
  ay = addBodyText(an, "A white cat ran by", ay, { fontSize: 13 });
  ay = addBodyText(an, "Startled by the clanging fall", ay, { fontSize: 13 });
  ay = addBodyText(an, "Of the treat jar's metal lid", ay, { fontSize: 13 });
  ay += 6;
  ay = addBodyText(an, "Clap each line. Write the beats next to it. Does it rhyme? (No -- it counts beats.)", ay, { fontSize: 10, italic: true });
  ay = addLinedArea(an, ay, 2, { lineSpacing: 22 });

  addPdfFooter(an, "Week 3 Session 1 | Rhyme and Tanka Analysis -- Page 1");

  an.addPage();
  let ay2 = addPdfHeader(an, "Sound and Feeling -- Exit Ticket", {
    color: C.ACCENT,
    subtitle: "Rhyme and beat.",
    lessonInfo: "Week 3 Session 1 | Year 6 Literacy",
    showNameDate: false,
  });
  ay2 = addTipBox(an, "Two quick tasks to show what you can hear in a poem.", ay2, { color: C.ACCENT });
  ay2 = addBodyText(an, "1. A word that rhymes with night: ______________   A word that rhymes with blue: ______________", ay2, { fontSize: 11 });
  ay2 += 8;
  ay2 = addBodyText(an, "2. How many beats in this tanka line: \"A white cat ran by\"?  ______", ay2, { fontSize: 11 });
  ay2 += 12;
  ay2 = addSectionHeading(an, "Challenge", ay2, { color: C.SECONDARY });
  ay2 = addBodyText(an, "How does the rhyme or the beat change the feeling of each poem?", ay2, { fontSize: 11, italic: true });
  ay2 = addLinedArea(an, ay2, 3, { lineSpacing: 24 });
  addPdfFooter(an, "Week 3 Session 1 | Exit Ticket");

  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Poetry_W3S1.pptx` }),
    writePdf(an, ANALYSIS_PDF_PATH),
  ]);

  console.log("PPTX written to " + `${OUT_DIR}/Poetry_W3S1.pptx`);
  console.log("Done: " + ANALYSIS_RESOURCE.name);
}

build().catch((err) => {
  console.error(err);
  process.exit(1);
});
