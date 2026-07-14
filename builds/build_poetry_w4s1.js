"use strict";

// Poetry Unit (Year 6) - Week 4 Session 1: Reading -- Funny Poems (Clerihew + Limerick)
// Forms introduced: Clerihew (4 lines, AABB, about a person, funny) + Limerick (5 lines, AABBA, bouncy).
// Curriculum spine: VCELA339 -- authors play with structure and language for humour and effect.
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
const FOOTER = "Poetry | Week 4 Session 1 | Year 6 Literacy";
const OUT_DIR = "output/Poetry_W4S1_Reading_Funny";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const ANALYSIS_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Funny Poems Analysis",
  "A clerihew and a limerick printed -- mark the rhyme scheme and underline what makes each funny."
);
const RESOURCE_ITEMS = [ANALYSIS_RESOURCE];
const ANALYSIS_PDF_PATH = path.join(OUT_DIR, ANALYSIS_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- This week our poems are meant to make people laugh. Two funny shapes: the clerihew and the limerick.
- Both use rhyme, but in different patterns, and both have a punchline feel.
- We are readers today -- we will write each one later in the week.

DO:
- Display the title slide.
- Have both example poems ready to read with a bit of comic timing.

TEACHER NOTES:
Humour is the hook this week. The learning is how rhyme pattern and rhythm create the funny effect.

WATCH FOR:
- Students who love the silliness -- channel it into noticing HOW the poem is built.

[Literacy: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- One sheet today: the Funny Poems Analysis. Both poems are printed for you to mark up.
- You will label the rhyme scheme and underline the funny part.

DO:
- Print the analysis sheet, one per student.
- Have the Poets Toolbox handy and highlighters ready.

[Literacy: Resources | VTLM 2.0: Enabling Learning]`;

const NOTES_LAUNCH = `SAY:
- Listen to this little poem about an art teacher. It is a clerihew.
- "Our art teacher, Mr. Shaw, really knows how to draw. But his awful paintings have caused many faintings."
- Did the ending surprise you? The joke lands right at the end.
- Let us find out how a clerihew is built so we can write our own.

DO:
- Read the clerihew aloud with comic timing -- pause before the last line.
- Quick partner share: "What made it funny? Which words rhymed?"
- Cold call two students.

TEACHER NOTES:
Supplied example, used exactly. Let students enjoy the joke, then notice the rhyme (Shaw/draw, paintings/faintings).

WATCH FOR:
- Students who notice the surprise ending -- that is the clerihew's comic engine.

[Literacy: Launch | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me. Today we notice how funny poems use rhyme and surprise.
- The first "I can" is for everyone -- find the rhyming words in a funny poem.

DO:
- Choral read the LI and success criteria.
- Note the two shapes: clerihew (4 lines) and limerick (5 lines).

TEACHER NOTES:
SC1 (all): find the rhyming words. SC2 (target): label a rhyme scheme (AABB or AABBA). SC3 (depth): explain how the structure makes it funny. Exit ticket targets SC2.

WATCH FOR:
- Students who mix up the two forms -- the CFU later checks they can tell them apart.

[Literacy: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_VOCAB = `SAY:
- One phrase for today: rhyme scheme.
- A rhyme scheme is the pattern of rhymes in a poem. We write it with letters.
- Lines that rhyme get the same letter. So if lines 1 and 2 rhyme, they are A and A.

DO:
- Choral say "rhyme scheme".
- Quick demo on the board: read a poem's line-ends and label them with letters.

TEACHER NOTES:
Rhyme scheme labelling (AABB, AABBA) is the analysis tool for both forms this week.

WATCH FOR:
- Students who give every line a new letter -- remind them: same sound, same letter.

[Literacy: Vocabulary | VTLM 2.0: Vocabulary Instruction]`;

const NOTES_IDO_CLERIHEW = `SAY:
- Let us look closely at the clerihew. Watch how I find its pattern and its joke.
- Four lines. Line 1 names a person: "Mr. Shaw". Line 2 ends with "draw" -- it rhymes with Shaw.
- Lines 3 and 4: "paintings" and "faintings" rhyme. So the scheme is AABB.
- And the rule is it should be FUNNY. The joke is that his paintings are so bad people faint.

DO:
- Display the clerihew and underline the rhyming line-ends.
- Label the lines A, A, B, B.
- Point to where the joke lands -- the last line.

TEACHER NOTES:
Three things define a clerihew: four lines, AABB, names a person, and is funny. Make all four visible.

WATCH FOR:
- Students who miss that line 1 names a person -- that is a defining feature.

[Literacy: I Do | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_IDO_LIMERICK = `SAY:
- Now a limerick -- five lines, and a bouncy rhythm you can almost dance to.
- Listen for the rhymes: Hall, fall and fall all rhyme -- lines one, two and five.
- The two short middle lines rhyme too: thing and spring.
- So the scheme is AABBA. The capital letters in the poem show the strong beats -- the bounce.

DO:
- Display the limerick and read it with a strong bouncy rhythm.
- Underline the rhyming line-ends and label them A, A, B, B, A.
- Tap the strong beats so students feel the rhythm.

TEACHER NOTES:
Supplied example, used exactly (the capitals mark the stressed beats). The teaching point is AABBA plus the da-da-DUM bounce.

WATCH FOR:
- Students who try to make all five lines rhyme -- show the two short lines have their own rhyme.

[Literacy: I Do | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_WEDO = `SAY:
- Your turn to help. Let us work out the limerick's rhyme scheme together.
- Read the last word of each line. Which lines rhyme with each other?
- Tell your partner the pattern using letters, then we will check.

DO:
- Display the limerick again.
- Read each line-end word with the class.
- Cold call for which lines rhyme, then reveal AABBA.

TEACHER NOTES:
Guided practice in labelling a rhyme scheme. Lines 1, 2, 5 share one rhyme; lines 3, 4 share another.

WATCH FOR:
- Students who only spot the long-line rhyme -- prompt them to check the two short lines.

[Literacy: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_REVEAL = `SAY:
- The pattern is AABBA. Lines 1, 2 and 5 rhyme (Hall, fall, fall). Lines 3 and 4 rhyme (thing, spring).
- That is what makes a limerick bounce the way it does.

DO:
- Display the reveal.
- Point to each rhyming group.

[Literacy: We Do Reveal | VTLM 2.0: Guided Practice]`;

const NOTES_CFU = `SAY:
- Quick check. Read this poem. Is it a clerihew or a limerick?
- Count the lines and look at the rhyme pattern.
- Show me on your fingers: one for clerihew, two for limerick.

DO:
- Display the mystery poem.
- Show Me Fingers.
- Scan: most should choose clerihew (1).

CFU CHECKPOINT:
Technique: Show Me Fingers (1 clerihew, 2 limerick)
Script:
- "Clerihew or limerick?"
- Scan for: students count four lines and AABB, choosing clerihew.
PROCEED (>=80%): Release to You Do.
PIVOT (<80%): Use the optional re-teach slide that follows -- it lines up the two forms side by side.

TEACHER NOTES:
The mystery poem has four lines and an AABB scheme, so it is a clerihew. A limerick would have five lines and AABBA.

WATCH FOR:
- Students who guess from the humour alone -- redirect to counting lines and the rhyme pattern.

[Literacy: CFU | VTLM 2.0: Formative Assessment]`;

const NOTES_CFU_REVEAL = `SAY:
- It is a clerihew -- four lines, AABB, and it names a person.
- A limerick would have five lines and the AABBA pattern.

DO:
- Display the reveal.
- Restate the two shapes once more.

[Literacy: CFU Reveal | VTLM 2.0: Formative Assessment]`;

const NOTES_RETEACH = `SAY:
- Let us sort the two funny forms side by side.
- A clerihew: FOUR lines, rhyme AABB, names a person, and is funny.
- A limerick: FIVE lines, rhyme AABBA, a bouncy rhythm, and is funny.
- The quickest test? Count the lines. Four is a clerihew. Five is a limerick.

DO:
- Display the re-teach slide with the two columns.
- Re-check: show a short poem, ask students to count lines and name the form.

TEACHER NOTES:
OPTIONAL. Use only if the CFU was below 80%. Different approach: a two-column comparison with "count the lines" as the quick test.

WATCH FOR:
- Students who still blend them -- anchor on the line count.

[Literacy: Optional Re-teach | VTLM 2.0: Reteach with Different Approach]`;

const NOTES_YOUDO = `SAY:
- Your turn to be a poetry detective. Both funny poems are on your sheet.
- On each poem, label the rhyme scheme with letters at the side.
- Then underline the part that makes it funny.
- Work quietly first, then compare with a partner.

DO:
- Distribute the analysis sheet.
- Circulate. Prioritise students who needed the re-teach.
- Conference: "Read me the rhyme scheme. Now show me the funny part."

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Just circle the rhyming pairs on the clerihew and label them A, A, B, B.
- Extra Notes: Finding the rhymes meets SC1.
EXTENDING PROMPT:
- Task: Explain HOW each poem makes you laugh -- is it a surprise ending, a silly picture, or the bouncy rhythm?
- Extra Notes: This reaches the depth criterion about structure creating humour.

TEACHER NOTES:
Students keep the marked-up sheet -- it feeds the drafting sessions this week.

WATCH FOR:
- Students who label a five-line poem as a clerihew -- redirect to counting lines.

[Literacy: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket. Two quick tasks.
- One: which poem has FIVE lines and rhymes AABBA -- a clerihew or a limerick?
- Two: name one thing that makes a clerihew funny.
- Two minutes.

DO:
- Two minutes, silent.
- Collect.

TEACHER NOTES:
Exit ticket targets SC2 -- knowing the two forms. The limerick is the five-line AABBA poem.

WATCH FOR:
- Students who confuse the forms -- note who needs a reminder before drafting.

[Literacy: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Self-check. Show me on your fingers, one to five, for each "I can".
- Then tell a partner: which poem made you laugh more, and what made it funny?

DO:
- Run the fingers check for each success criterion.
- Sixty seconds partner share.
- Preview: "Next session we write our own clerihew -- a funny four-line poem about a person."

TEACHER NOTES:
Note which form students found clearer to plan support for the drafting sessions.

[Literacy: Closing | VTLM 2.0: Review and Reflect]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Poetry Week 4 Session 1 -- Reading: Funny Poems";

  // SLIDE 1 -- Title
  titleSlide(
    pres,
    "Poems That Make You Laugh",
    "Reading for Humour -- Clerihews and Limericks",
    "Week 4 Session 1  |  Year 6 Literacy  |  Poetry",
    NOTES_TITLE
  );

  // SLIDE 2 -- Resources
  addResourceSlide(pres, RESOURCE_ITEMS, T, FOOTER, NOTES_RESOURCES);

  // SLIDE 3 -- Launch (read the clerihew)
  quoteSlide(
    pres,
    "Read Aloud",
    "A Clerihew",
    "Our art teacher, Mr. Shaw,\nReally knows how to draw.\nBut his awful paintings\nHave caused many faintings.",
    "",
    "What made it funny? Which words rhymed?",
    NOTES_LAUNCH,
    FOOTER
  );

  // SLIDE 4 -- LI / SC
  liSlide(
    pres,
    [
      "We are learning to notice how funny poems use rhyme, rhythm and surprise to make us laugh",
    ],
    [
      "I can find the rhyming words in a funny poem",
      "I can label a rhyme scheme (AABB or AABBA)",
      "I can explain how the structure makes the poem funny",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // SLIDE 5 -- Vocab: rhyme scheme
  vocabSlide(
    pres,
    "rhyme scheme",
    "noun",
    "The pattern of rhymes in a poem, written with letters. Lines that rhyme get the same letter.",
    "Lines 1 and 2 rhyme (A, A); lines 3 and 4 rhyme (B, B). That is AABB.",
    NOTES_VOCAB,
    FOOTER,
    { routine: "Say it together. Then label a poem's line-ends with letters." }
  );

  // SLIDE 6 -- I Do (clerihew)
  modellingSlide(
    pres,
    "I Do",
    "How a Clerihew Works",
    "Four lines. The rules:\n\n- Line 1 names a PERSON\n- Line 2 rhymes with the\n   person's name (A, A)\n- Lines 3 and 4 rhyme (B, B)\n- It should be FUNNY\n\nScheme: AABB.",
    "Our art teacher, Mr. Shaw,   (A)\n\nReally knows how to draw.   (A)\n\nBut his awful paintings   (B)\n\nHave caused many faintings.   (B)",
    NOTES_IDO_CLERIHEW,
    FOOTER
  );

  // SLIDE 7 -- I Do (limerick)
  modellingSlide(
    pres,
    "I Do",
    "How a Limerick Works",
    "Five lines, bouncy rhythm.\n\n- Lines 1, 2 and 5 rhyme (A)\n- Lines 3 and 4 rhyme (B)\n- Lines 3 and 4 are shorter\n\nScheme: AABBA.\nThe CAPITALS show the\nstrong beats -- the bounce.",
    "there WAS a young FELLow named HALL\nwho FELL in the SPRING in the FALL.\n'twould have BEEN a sad THING\nhad he DIED in the SPRING,\nbut he DIDn't—he DIED in the FALL.",
    NOTES_IDO_LIMERICK,
    FOOTER
  );

  // SLIDE 8 + 9 -- We Do (find the limerick scheme) with reveal
  withReveal(
    () => cfuSlide(
      pres,
      "We Do",
      "Find the Limerick's Rhyme Scheme",
      "Which lines rhyme with each other?",
      "Read the last word of each line.\nLines 1, 2, 5: Hall / fall / fall.\nLines 3, 4: thing / spring.\n\nWhat is the rhyme scheme?",
      NOTES_WEDO,
      FOOTER
    ),
    (slide) => {
      addRevealAnswerBar(
        slide,
        "AABBA -- lines 1, 2, 5 rhyme; lines 3, 4 rhyme.",
        { color: C.SUCCESS, label: "Answer", showTickAndFix: false }
      );
      slide.addNotes(NOTES_WEDO_REVEAL);
    }
  );

  // SLIDE 10 + 11 -- CFU (clerihew or limerick?) with reveal
  withReveal(
    () => cfuSlide(
      pres,
      "CFU",
      "Clerihew or Limerick?",
      "Show Me Fingers: 1 (clerihew) or 2 (limerick)",
      "My cousin, a girl named Lou,\nKept a very loud kangaroo.\nIt bounced on the bed,\nAnd stood on its head.\n\nHow many lines? What pattern?",
      NOTES_CFU,
      FOOTER
    ),
    (slide) => {
      addRevealAnswerBar(
        slide,
        "Clerihew -- four lines, AABB, and it names a person (Lou).",
        { color: C.SUCCESS, label: "Answer", showTickAndFix: false }
      );
      slide.addNotes(NOTES_CFU_REVEAL);
    }
  );

  // SLIDE 12 -- Optional Re-teach
  modellingSlide(
    pres,
    "Optional Re-teach",
    "Two Funny Forms -- Sort Them",
    "CLERIHEW\n\n- 4 lines\n- rhyme AABB\n- names a person\n- funny\n\nLIMERICK\n\n- 5 lines\n- rhyme AABBA\n- bouncy rhythm\n- funny",
    "Quick test:\n\nCount the lines.\n\n4 lines = clerihew\n5 lines = limerick",
    NOTES_RETEACH,
    FOOTER
  );

  // SLIDE 13 -- You Do
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "You Do", { color: C.PRIMARY, w: 1.5 });
    addTitle(s, "Poetry Detective -- Funny Poems");

    addCard(s, 0.5, CONTENT_TOP, 9, 1.7, { strip: C.PRIMARY, fill: C.WHITE });
    s.addText("On your Analysis sheet:", {
      x: 0.75, y: CONTENT_TOP + 0.12, w: 8.4, h: 0.32,
      fontSize: 16, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("First:   Label the rhyme scheme on each poem with letters.\nNext:   Underline the part that makes it funny.\nThen:   Check: the clerihew is AABB, the limerick is AABBA.", {
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
    s.addText("- Just circling the rhyming pairs is a great start.\n- Ready for more? Explain HOW each poem makes you laugh -- surprise, silly picture, or rhythm.", {
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
      "Which poem has FIVE lines and rhymes AABBA -- a clerihew or a limerick?",
      "Name one thing that makes a clerihew funny.",
    ],
    NOTES_EXIT,
    FOOTER,
    { assessesSc: 2, title: "Know Your Funny Forms" }
  );

  // SLIDE 15 -- Closing
  closingSlide(
    pres,
    {
      reflectionPrompt: "Which poem made you laugh more, and what made it funny?",
      scItems: [
        "I can find the rhyming words in a funny poem",
        "I can label a rhyme scheme (AABB or AABBA)",
        "I can explain how the structure makes the poem funny",
      ],
      selfAssessment: {
        prompt: "Show on your fingers: 1 (need help) to 5 (could teach it)",
        options: ["1-2", "3", "4-5"],
      },
    },
    NOTES_CLOSING
  );

  // ---- PDF: Funny Poems Analysis --------------------------------------------
  const an = createPdf({ title: ANALYSIS_RESOURCE.name });
  let ay = addPdfHeader(an, "Funny Poems Analysis", {
    color: C.PRIMARY,
    subtitle: "Label the rhyme scheme. Underline what makes each poem funny.",
    lessonInfo: "Week 4 Session 1 | Year 6 Literacy | Poetry",
    showNameDate: true,
  });

  ay = addTipBox(an, "Read each poem aloud first. Then mark the rhyme scheme and the joke.", ay, { color: C.PRIMARY });

  ay = addSectionHeading(an, "Poem 1 -- A Clerihew (4 lines, AABB)", ay, { color: C.PRIMARY });
  ay = addBodyText(an, "Our art teacher, Mr. Shaw,", ay, { fontSize: 13 });
  ay = addBodyText(an, "Really knows how to draw.", ay, { fontSize: 13 });
  ay = addBodyText(an, "But his awful paintings", ay, { fontSize: 13 });
  ay = addBodyText(an, "Have caused many faintings.", ay, { fontSize: 13 });
  ay += 6;
  ay = addBodyText(an, "Label each line A, A, B, B. Underline the funny part.", ay, { fontSize: 10, italic: true });
  ay = addLinedArea(an, ay, 1, { lineSpacing: 22 });
  ay += 6;

  ay = addSectionHeading(an, "Poem 2 -- A Limerick (5 lines, AABBA)", ay, { color: C.SECONDARY });
  ay = addBodyText(an, "there WAS a young FELLow named HALL", ay, { fontSize: 13 });
  ay = addBodyText(an, "who FELL in the SPRING in the FALL.", ay, { fontSize: 13 });
  ay = addBodyText(an, "'twould have BEEN a sad THING", ay, { fontSize: 13 });
  ay = addBodyText(an, "had he DIED in the SPRING,", ay, { fontSize: 13 });
  ay = addBodyText(an, "but he DIDn't—he DIED in the FALL.", ay, { fontSize: 13 });
  ay += 6;
  ay = addBodyText(an, "Label each line A, A, B, B, A. Which lines bounce together?", ay, { fontSize: 10, italic: true });
  ay = addLinedArea(an, ay, 1, { lineSpacing: 22 });

  addPdfFooter(an, "Week 4 Session 1 | Funny Poems Analysis -- Page 1");

  an.addPage();
  let ay2 = addPdfHeader(an, "Funny Forms -- Exit Ticket", {
    color: C.ACCENT,
    subtitle: "Show what you know about the two funny shapes.",
    lessonInfo: "Week 4 Session 1 | Year 6 Literacy",
    showNameDate: false,
  });
  ay2 = addTipBox(an, "Two quick questions about clerihews and limericks.", ay2, { color: C.ACCENT });
  ay2 = addBodyText(an, "1. Which poem has FIVE lines and rhymes AABBA? (Circle one)   Clerihew   /   Limerick", ay2, { fontSize: 11 });
  ay2 += 8;
  ay2 = addBodyText(an, "2. Name one thing that makes a clerihew funny:", ay2, { fontSize: 11 });
  ay2 = addLinedArea(an, ay2, 2, { lineSpacing: 24 });
  ay2 += 8;
  ay2 = addSectionHeading(an, "Challenge", ay2, { color: C.SECONDARY });
  ay2 = addBodyText(an, "How does the surprise ending or the bouncy rhythm make you laugh?", ay2, { fontSize: 11, italic: true });
  ay2 = addLinedArea(an, ay2, 3, { lineSpacing: 24 });
  addPdfFooter(an, "Week 4 Session 1 | Exit Ticket");

  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Poetry_W4S1.pptx` }),
    writePdf(an, ANALYSIS_PDF_PATH),
  ]);

  console.log("PPTX written to " + `${OUT_DIR}/Poetry_W4S1.pptx`);
  console.log("Done: " + ANALYSIS_RESOURCE.name);
}

build().catch((err) => {
  console.error(err);
  process.exit(1);
});
