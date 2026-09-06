"use strict";

// Poetry Unit (Year 6) - Week 1 Session 1: Reading Poems -- Shape and Sound
// Forms introduced: Three-Line Poem (subject / location / action) + Haiku (5-7-5)
// Curriculum spine: VC2E5LE04 / VC2E6LE04 -- imagery (simile, metaphor, personification) and sound devices.
// First poetry lesson of the year -> heavy modelling, exemplars used exactly as supplied.
// Unit-wide look: literacy / grade56 / variant 3 ("Ink & Paper") -- fixed for all 19 sessions.

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
const FOOTER = "Poetry | Week 1 Session 1 | Year 6 Literacy";
const OUT_DIR = "output/Poetry_W1S1_Reading_Poems";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const ANALYSIS_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Poem Analysis",
  "Two example poems printed with space to mark the shape and circle an image."
);
const TOOLBOX_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Poets Toolbox Reference",
  "Keep-for-the-unit reference: imagery (simile, metaphor, personification) and sound devices."
);
const RESOURCE_ITEMS = [ANALYSIS_RESOURCE, TOOLBOX_RESOURCE];
const ANALYSIS_PDF_PATH = path.join(OUT_DIR, ANALYSIS_RESOURCE.fileName);
const TOOLBOX_PDF_PATH = path.join(OUT_DIR, TOOLBOX_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Welcome to our poetry unit. Over the next few weeks we are going to read poems and write our own.
- A poem can paint a whole picture in just a few words, and that is the magic we are going to learn.
- Today we are just READERS. We look closely at two short poems and notice how they are built.
- Some of you may have written poems before, and some of you may feel this is brand new. Either way is fine -- we build it together.

DO:
- Display the title slide.
- Have the two example poems ready to read aloud.
- Keep the mood calm and curious. This is the first poetry lesson of the year.

TEACHER NOTES:
This session sets up the whole unit. Students only READ and ANALYSE today. They draft their own three-line poem and haiku in the next two sessions.

WATCH FOR:
- Students who say they are not good at poetry -- reassure: noticing is the only job today.

[Literacy: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- Two things on your desk today.
- The Poem Analysis sheet has our two poems printed, with space to mark them up.
- The Poets Toolbox is a reference card. Keep it -- we use it every week of this unit.

DO:
- Print the Poem Analysis sheet, one per student.
- Print the Poets Toolbox, one per student to keep in their poetry folder.
- Have highlighters or coloured pencils ready for marking the poems.

[Literacy: Resources | VTLM 2.0: Enabling Learning]`;

const NOTES_LAUNCH = `SAY:
- Listen to this. It is only three lines long. Close your eyes if it helps you picture it.
- "The tree. In the middle of the paddock. Stands."
- Now tell me -- what did you SEE in your mind? A big tree, all alone, out in a wide empty field?
- That whole picture came from just five or six words. That is what poems do.

DO:
- Read the three-line poem slowly, twice.
- Quick partner share: "What did you picture? Tell your partner in one sentence."
- Cold call two or three students and collect their pictures warmly.

TEACHER NOTES:
The launch shows the unit's big idea immediately: poems make pictures with very few words. Do not analyse yet -- just let them feel the image.

WATCH FOR:
- Students who add their own detail (a hot day, no other trees) -- celebrate, that is exactly the imagining poems invite.

[Literacy: Text Launch | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- Today we notice how a poem uses its SHAPE and its IMAGES to paint a picture.
- Three "I can" statements. The first one everyone can do -- read a poem and say what you picture.

DO:
- Choral read the LI and the three success criteria.
- Point out: today is about NOTICING, not writing our own yet.

TEACHER NOTES:
SC1 (all): read aloud and say what you picture. SC2 (target): find the parts of a poem's shape. SC3 (depth): spot an image or sound and explain its effect. Exit ticket targets SC2 and SC3.

WATCH FOR:
- Students who think they must write a poem today -- redirect: "Just reading and noticing today."

[Literacy: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_VOCAB = `SAY:
- One word to know for the whole unit: imagery.
- Imagery means the pictures that words make in your mind. Sights, sounds, smells, even feelings.
- When the poet wrote "the wings of the dragonfly sparkle like blue gold", your mind made a picture. That is imagery doing its job.

DO:
- Choral say "imagery".
- Ask students to whisper to a partner one thing they can picture right now from any poem today.

TEACHER NOTES:
Imagery is the curriculum spine for this unit. Keep returning to it: "What picture did that make?"

WATCH FOR:
- Students who think imagery only means a drawing -- redirect: "It is the picture words make in your HEAD."

[Literacy: Vocabulary | VTLM 2.0: Vocabulary Instruction]`;

const NOTES_IDO_THREE = `SAY:
- Let us look at this three-line poem together. Watch how I work out its shape.
- Line one is just "The tree". That is the SUBJECT -- what the poem is about, in a word or two.
- Line two is "In the middle of the paddock". That tells me WHERE the subject is. The location.
- Line three is "Stands". That is the ACTION -- what the subject is doing.
- So the shape is simple: subject, then where, then what it does. Three lines, one clear picture.

DO:
- Display the slide and read the poem aloud.
- Point to each line as you name its job: subject, location, action.
- Say the pattern back once more so students hear it as a recipe they could follow.

TEACHER NOTES:
This is the form students will draft in the next session, so make the three jobs crystal clear. Keep the poem itself large -- it is the hero of the slide.

WATCH FOR:
- Students who muddle location and action -- reread line two then line three and ask "where, or what it does?"

[Literacy: I Do | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_IDO_HAIKU = `SAY:
- Here is a different shape: a haiku. Haiku are tiny poems, usually about nature, from Japan.
- A haiku has three lines, and it counts BEATS in the words. We call those beats syllables.
- Watch me clap the first line: "Win-ter-is-com-ing". Five claps. Five syllables.
- Line two: "Snow-will-be-ar-riv-ing-soon". Seven. Line three: "We-should-rake-the-leaves". Five again.
- So a haiku is five, seven, five. Notice it is about the seasons -- that is very haiku.

DO:
- Display the haiku and read it aloud once for meaning.
- Clap each line slowly with the class, counting syllables out loud.
- Write 5, 7, 5 on the board beside the three lines.

TEACHER NOTES:
Clapping syllables is the key skill students need to draft a haiku next session. Do it physically -- hands, not just counting in heads.

WATCH FOR:
- Students who clap whole words instead of beats -- model a two-beat word like "win-ter" slowly.

[Literacy: I Do | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_WEDO = `SAY:
- Your turn to help with a second haiku. We do this one together.
- Read it with me: "Underneath the moon, the wings of the dragonfly sparkle like blue gold."
- First job: clap each line and count the beats. Are they five, seven, five?
- Second job: find the words that paint a picture. Which words made a picture in your head?

DO:
- Read the haiku aloud together.
- Clap each line as a class and check the syllable count.
- Cold call for the image: which words made a picture? Steer them to "sparkle like blue gold".
- Then reveal the answer slide.

TEACHER NOTES:
This We Do practises both unit skills at once: counting syllables and spotting imagery. "Sparkle like blue gold" is a simile because it uses the word "like".

WATCH FOR:
- Students who find the image but cannot name it -- that is fine today, naming comes in the next slide.

[Literacy: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_REVEAL = `SAY:
- Let us check. Five, seven, five -- yes, it is a true haiku.
- And the picture words: "sparkle like blue gold". Because it uses the word "like", that is a special kind of image called a simile.
- A simile compares two things using "like" or "as".

DO:
- Display the reveal.
- Reread "sparkle like blue gold" and underline the word "like".

TEACHER NOTES:
This connects the noticing to a named technique (simile) without overloading. Similes are revisited in the CFU next.

WATCH FOR:
- Students who now spot "like" or "as" elsewhere -- celebrate, they are ready for the CFU.

[Literacy: We Do Reveal | VTLM 2.0: Guided Practice]`;

const NOTES_CFU = `SAY:
- Quick check on similes. A simile compares two things using "like" or "as".
- Read both options. Which one is a simile?
- Show me on your fingers: one for A, two for B.

DO:
- Display both options.
- Show Me Fingers.
- Scan: most students should choose B.
- Cold call one student: "How do you know B is the simile?"

CFU CHECKPOINT:
Technique: Show Me Fingers (1 for A, 2 for B)
Script:
- "Which one is a simile?"
- Scan for: most students choose B and can point to the word "like".
PROCEED (>=80%): Move to You Do.
PIVOT (<80%): Use the optional re-teach slide that follows -- it shows simile, metaphor and personification as three simple picture-making moves.

TEACHER NOTES:
A is a plain fact with no comparison. B compares the moon to a lamp using "like", so B is the simile.

WATCH FOR:
- Students who pick A because it sounds "poetic" -- redirect: "A simile needs the word like or as."

[Literacy: CFU | VTLM 2.0: Formative Assessment]`;

const NOTES_CFU_REVEAL = `SAY:
- The simile is B: "The moon glows like a lamp." It compares the moon to a lamp using "like".
- A just tells us a fact. B makes a picture by comparing.

DO:
- Display the reveal banner.
- Underline the word "like" in option B.

[Literacy: CFU Reveal | VTLM 2.0: Formative Assessment]`;

const NOTES_RETEACH = `SAY:
- Let us slow down and look at three ways poets paint pictures. These are our imagery tools.
- A simile compares using "like" or "as": "the clouds are like cotton wool".
- A metaphor says one thing IS another: "the clouds are cotton wool". No "like", it just says it is.
- Personification gives a human action to something that is not human: "the clouds marched across the sky".
- Same clouds, three different pictures. That is a poet choosing tools.

DO:
- Display the re-teach slide.
- Read each example aloud and act it out a little (march for personification).
- Re-check: say "the wind ___". Ask students to finish it as a simile, then as personification, on whiteboards.

TEACHER NOTES:
OPTIONAL. Use only if the CFU was below 80%. Different approach: three named tools side by side with one shared subject (clouds) so the contrast is obvious.

WATCH FOR:
- Students who write a simile when asked for personification -- point back to the "human action" idea.

[Literacy: Optional Re-teach | VTLM 2.0: Reteach with Different Approach]`;

const NOTES_YOUDO = `SAY:
- Your turn to be a poetry detective. On your Poem Analysis sheet you have both poems.
- Job one: mark the SHAPE. On the three-line poem, label the subject, the location and the action.
- Job two: find ONE image. Circle the words that painted a picture in your head, and write what you saw beside it.
- Work quietly first, then you can compare with a partner.

DO:
- Distribute the Poem Analysis sheet and the Poets Toolbox.
- Circulate. Prioritise students who needed the re-teach.
- Quick conferences: "Read me your image. What did you picture?"

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the Poets Toolbox card. Just circle one set of picture-words in either poem and tell your partner what you saw.
- Extra Notes: This still hits SC1 and SC3 without needing to label every part.
EXTENDING PROMPT:
- Task: Find a simile, a metaphor OR a personification across the two poems, name it, and explain the effect: what does it make you feel or see?
- Extra Notes: Pushes confident students into the depth criterion.

TEACHER NOTES:
Students keep the marked-up sheet -- it feeds the drafting sessions that follow. Capture strong noticings to share.

WATCH FOR:
- Students who label the haiku with subject/location/action -- redirect: that shape is the three-line poem; the haiku's shape is its 5-7-5 beats.

[Literacy: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket. On the bottom of your sheet, write ONE image from today's poems.
- Copy the words that made the picture, then write what you saw in your mind.
- One image is enough. Two minutes.

DO:
- Two minutes, silent.
- Collect the sheets.

TEACHER NOTES:
Exit ticket targets SC2 and SC3: spotting an image and explaining what it shows. Use it to group students for the drafting sessions.

WATCH FOR:
- Students who copy a whole line with no picture words -- next session, model choosing the strongest words.

[Literacy: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Quick self-check. Show me on your fingers, one to five, for each "I can" statement.
- Then tell your partner: which poem painted the clearest picture for you, and why?

DO:
- Run the fingers check for each success criterion.
- Sixty seconds partner share.
- Preview: "Next session we write our OWN three-line poem."

TEACHER NOTES:
This closes the reading session and sets up the first drafting session. Note any students who rated themselves low on picturing images.

[Literacy: Closing | VTLM 2.0: Review and Reflect]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Poetry Week 1 Session 1 -- Reading Poems: Shape and Sound";

  // SLIDE 1 -- Title
  titleSlide(
    pres,
    "Painting Pictures with Words",
    "Reading Poems: Shape and Sound -- Three-Line Poems and Haiku",
    "Week 1 Session 1  |  Year 6 Literacy  |  Poetry",
    NOTES_TITLE
  );

  // SLIDE 2 -- Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, T, FOOTER, NOTES_RESOURCES);

  // SLIDE 3 -- Text Launch (read the three-line poem)
  quoteSlide(
    pres,
    "Read Aloud",
    "A Three-Line Poem",
    "The tree\nIn the middle of the paddock\nStands",
    "",
    "Close your eyes. What do you picture? Tell your partner in one sentence.",
    NOTES_LAUNCH,
    FOOTER
  );

  // SLIDE 4 -- LI / SC
  liSlide(
    pres,
    [
      "We are learning to notice how a poem uses its SHAPE and its IMAGES to paint a picture in just a few words",
    ],
    [
      "I can read a poem aloud and say what I picture",
      "I can find the parts of a poem's shape and say what each line does",
      "I can spot an image a poet uses and explain what it makes me see",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // SLIDE 5 -- Vocab: imagery
  vocabSlide(
    pres,
    "imagery",
    "noun",
    "The pictures that words make in your mind -- sights, sounds, smells and feelings.",
    "The wings of the dragonfly sparkle like blue gold.",
    NOTES_VOCAB,
    FOOTER,
    { routine: "Say it together. Then tell a partner one thing you can picture." }
  );

  // SLIDE 6 -- I Do: three-line poem shape (poem is the hero)
  modellingSlide(
    pres,
    "I Do",
    "The Shape of a Three-Line Poem",
    "Every line has ONE job:\n\n- Line 1 = SUBJECT\n   (what it is about, in a word or two)\n\n- Line 2 = LOCATION\n   (where the subject is)\n\n- Line 3 = ACTION\n   (what the subject does)",
    "The tree\n\nIn the middle of the paddock\n\nStands",
    NOTES_IDO_THREE,
    FOOTER
  );

  // SLIDE 7 -- I Do: haiku shape 5-7-5 (poem is the hero)
  modellingSlide(
    pres,
    "I Do",
    "The Shape of a Haiku -- Count the Beats",
    "A haiku is a tiny nature poem.\n\nThree lines. Count the beats (syllables):\n\n- Line 1 = 5 beats\n- Line 2 = 7 beats\n- Line 3 = 5 beats\n\nClap each word with me.",
    "Winter is coming.\n(5)\n\nSnow will be arriving soon.\n(7)\n\nWe should rake the leaves.\n(5)",
    NOTES_IDO_HAIKU,
    FOOTER
  );

  // SLIDE 8 + 9 -- We Do (count beats + find the image) with reveal
  withReveal(
    () => modellingSlide(
      pres,
      "We Do",
      "Together: A Second Haiku",
      "Two jobs, together:\n\n1. Clap each line.\n   Is it 5 - 7 - 5?\n\n2. Find the picture words.\n   Which words made a picture\n   in your head?",
      "Underneath the moon\n\nThe wings of the dragonfly\n\nSparkle like blue gold",
      NOTES_WEDO,
      FOOTER
    ),
    (slide) => {
      addRevealAnswerBar(
        slide,
        "5 - 7 - 5. Image: \"sparkle like blue gold\" -- a simile (it uses \"like\").",
        { color: C.SUCCESS, label: "Notice", showTickAndFix: false }
      );
      slide.addNotes(NOTES_WEDO_REVEAL);
    }
  );

  // SLIDE 10 + 11 -- CFU: simile hinge, with reveal
  withReveal(
    () => cfuSlide(
      pres,
      "CFU",
      "Which One Is a Simile?",
      "Show Me Fingers: 1 (A) or 2 (B)",
      "A simile compares two things using \"like\" or \"as\".\n\nA)  The moon is bright tonight.\n\nB)  The moon glows like a lamp.",
      NOTES_CFU,
      FOOTER
    ),
    (slide) => {
      addRevealAnswerBar(
        slide,
        "B is the simile -- it compares the moon to a lamp using \"like\".",
        { color: C.SUCCESS, label: "Answer", showTickAndFix: false }
      );
      slide.addNotes(NOTES_CFU_REVEAL);
    }
  );

  // SLIDE 12 -- Optional Re-teach: three imagery tools
  modellingSlide(
    pres,
    "Optional Re-teach",
    "Three Ways Poets Paint a Picture",
    "Same clouds. Three tools:\n\n- SIMILE compares with\n   \"like\" or \"as\"\n\n- METAPHOR says it IS\n   the other thing\n\n- PERSONIFICATION gives it\n   a human action",
    "Simile:\nThe clouds are like cotton wool.\n\nMetaphor:\nThe clouds are cotton wool.\n\nPersonification:\nThe clouds marched across the sky.",
    NOTES_RETEACH,
    FOOTER
  );

  // SLIDE 13 -- You Do
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "You Do", { color: C.PRIMARY, w: 1.5 });
    addTitle(s, "Poetry Detective");

    addCard(s, 0.5, CONTENT_TOP, 9, 1.7, { strip: C.PRIMARY, fill: C.WHITE });
    s.addText("On your Poem Analysis sheet:", {
      x: 0.75, y: CONTENT_TOP + 0.12, w: 8.4, h: 0.32,
      fontSize: 16, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("First:   Mark the SHAPE -- label the subject, location and action on the three-line poem.\nNext:   Find ONE image -- circle the words that painted a picture.\nThen:   Write what you saw beside it.", {
      x: 0.75, y: CONTENT_TOP + 0.52, w: 8.4, h: 1.05,
      fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    const tipY = CONTENT_TOP + 1.85;
    const tipH = SAFE_BOTTOM - tipY;
    addCard(s, 0.5, tipY, 9, tipH, { strip: C.ACCENT, fill: C.BG_CARD });
    s.addText("Stuck? Use your Poets Toolbox", {
      x: 0.75, y: tipY + 0.10, w: 8.4, h: 0.30,
      fontSize: 14, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("- Just circle one set of picture-words and tell a partner what you saw.\n- Ready for more? Name the image -- simile, metaphor or personification -- and explain its effect.", {
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
      "Write ONE image from today's poems.",
      "Copy the words that made the picture, then write what you saw in your mind.",
    ],
    NOTES_EXIT,
    FOOTER,
    { assessesSc: 3, title: "One Image You Could Picture" }
  );

  // SLIDE 15 -- Closing
  closingSlide(
    pres,
    {
      reflectionPrompt: "Partner share: which poem painted the clearest picture for you, and why?",
      scItems: [
        "I can read a poem aloud and say what I picture",
        "I can find the parts of a poem's shape and say what each line does",
        "I can spot an image a poet uses and explain what it makes me see",
      ],
      selfAssessment: {
        prompt: "Show on your fingers: 1 (need help) to 5 (could teach it)",
        options: ["1-2", "3", "4-5"],
      },
    },
    NOTES_CLOSING
  );

  // ---- PDF: Poem Analysis ---------------------------------------------------
  const an = createPdf({ title: ANALYSIS_RESOURCE.name });
  let ay = addPdfHeader(an, "Poem Analysis -- Reading Like a Poet", {
    color: C.PRIMARY,
    subtitle: "Mark the shape. Circle an image. Write what you picture.",
    lessonInfo: "Week 1 Session 1 | Year 6 Literacy | Poetry",
    showNameDate: true,
  });

  ay = addTipBox(an, "You are a poetry detective today. Read each poem aloud first, then mark it up.", ay, { color: C.PRIMARY });

  ay = addSectionHeading(an, "Poem 1 -- A Three-Line Poem", ay, { color: C.PRIMARY });
  ay = addBodyText(an, "The tree", ay, { fontSize: 14 });
  ay = addBodyText(an, "In the middle of the paddock", ay, { fontSize: 14 });
  ay = addBodyText(an, "Stands", ay, { fontSize: 14 });
  ay += 6;
  ay = addBodyText(an, "Label each line: which one is the SUBJECT, the LOCATION, the ACTION?", ay, { fontSize: 10, italic: true });
  ay = addLinedArea(an, ay, 3, { lineSpacing: 22 });
  ay += 6;

  ay = addSectionHeading(an, "Poem 2 -- A Haiku (5 - 7 - 5)", ay, { color: C.SECONDARY });
  ay = addBodyText(an, "Underneath the moon", ay, { fontSize: 14 });
  ay = addBodyText(an, "The wings of the dragonfly", ay, { fontSize: 14 });
  ay = addBodyText(an, "Sparkle like blue gold", ay, { fontSize: 14 });
  ay += 6;
  ay = addBodyText(an, "Clap each line. Write the number of beats (syllables) next to it. Circle the words that paint a picture.", ay, { fontSize: 10, italic: true });
  ay = addLinedArea(an, ay, 3, { lineSpacing: 22 });

  addPdfFooter(an, "Week 1 Session 1 | Poem Analysis -- Page 1");

  an.addPage();
  let ay2 = addPdfHeader(an, "Your Image -- Exit Ticket", {
    color: C.ACCENT,
    subtitle: "One image from today's poems.",
    lessonInfo: "Week 1 Session 1 | Year 6 Literacy",
    showNameDate: false,
  });
  ay2 = addTipBox(an, "Copy the words that made a picture. Then write what you saw in your mind.", ay2, { color: C.ACCENT });
  ay2 = addBodyText(an, "The words that made a picture:", ay2, { fontSize: 11 });
  ay2 = addLinedArea(an, ay2, 2, { lineSpacing: 24 });
  ay2 += 6;
  ay2 = addBodyText(an, "What I pictured in my mind:", ay2, { fontSize: 11 });
  ay2 = addLinedArea(an, ay2, 3, { lineSpacing: 24 });
  addPdfFooter(an, "Week 1 Session 1 | Exit Ticket");

  // ---- PDF: Poets Toolbox (keep-for-the-unit reference) ---------------------
  const tb = createPdf({ title: TOOLBOX_RESOURCE.name });
  let ty = addPdfHeader(tb, "The Poet's Toolbox", {
    color: C.PRIMARY,
    subtitle: "Keep this in your poetry folder -- we use it every week.",
    lessonInfo: "Year 6 Literacy | Poetry Unit",
    showNameDate: false,
  });

  ty = addTipBox(tb, "Poets paint pictures (imagery) and play with sound. Here are the tools.", ty, { color: C.PRIMARY });

  ty = addSectionHeading(tb, "Imagery -- making pictures", ty, { color: C.PRIMARY });
  ty = addBodyText(tb, "Simile -- compares two things using \"like\" or \"as\".  Example: her smile was like sunshine.", ty, { fontSize: 11 });
  ty = addBodyText(tb, "Metaphor -- says one thing IS another.  Example: the playground was a zoo.", ty, { fontSize: 11 });
  ty = addBodyText(tb, "Personification -- gives a human action to something not human.  Example: the wind whispered.", ty, { fontSize: 11 });
  ty += 8;

  ty = addSectionHeading(tb, "Sound -- playing with how it sounds", ty, { color: C.SECONDARY });
  ty = addBodyText(tb, "Rhyme -- words that end with the same sound.  Example: cat / hat.", ty, { fontSize: 11 });
  ty = addBodyText(tb, "Rhythm (beat) -- the pattern of strong and soft beats when you read it aloud.", ty, { fontSize: 11 });
  ty = addBodyText(tb, "Syllables -- the beats inside a word. Clap them: but-ter-fly = 3.", ty, { fontSize: 11 });
  ty = addBodyText(tb, "Alliteration -- words near each other starting with the same sound.  Example: slippery snake.", ty, { fontSize: 11 });
  ty += 8;

  ty = addSectionHeading(tb, "Ask yourself", ty, { color: C.ACCENT });
  ty = addBodyText(tb, "What picture do I want the reader to see? Which tool will help me paint it?", ty, { fontSize: 11, italic: true });

  addPdfFooter(tb, "Poetry Unit | The Poet's Toolbox -- REFERENCE");

  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Poetry_W1S1.pptx` }),
    writePdf(an, ANALYSIS_PDF_PATH),
    writePdf(tb, TOOLBOX_PDF_PATH),
  ]);

  console.log("PPTX written to " + `${OUT_DIR}/Poetry_W1S1.pptx`);
  console.log("Done: " + ANALYSIS_RESOURCE.name);
  console.log("Done: " + TOOLBOX_RESOURCE.name);
}

build().catch((err) => {
  console.error(err);
  process.exit(1);
});
