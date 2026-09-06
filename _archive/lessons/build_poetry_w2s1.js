"use strict";

// Poetry Unit (Year 6) - Week 2 Session 1: Cinquain -- Read, Analyse and Draft
// Form: word-count cinquain (1-2-3-4-1 words). Extension: syllable cinquain (2-4-6-8-2).
// Combined read + draft session (per the planner). Supplied examples used exactly.
// Catch-up design: the form is taught in full each session; no dependency on Week 1.
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

const SESSION_NUMBER = 1;
const FOOTER = "Poetry | Week 2 Session 1 | Year 6 Literacy";
const OUT_DIR = "output/Poetry_W2S1_Cinquain";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const PLAN_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Cinquain Plan",
  "Cinquain planning template (1-2-3-4-1 words) with a word bank and a challenge syllable version."
);
const RESOURCE_ITEMS = [PLAN_RESOURCE];
const PLAN_PDF_PATH = path.join(OUT_DIR, PLAN_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- This week we meet two new poem shapes. Today it is the cinquain -- a five-line poem.
- We will read one, work out how it is built, then write our own. All in one lesson.
- If poetry still feels new, that is fine. We build every step together.

DO:
- Display the title slide.
- Have the example cinquain ready to read aloud.

TEACHER NOTES:
This session both analyses and drafts a cinquain. Keep the form simple: count the words per line.

WATCH FOR:
- Students who expect rhyming -- a cinquain does not rhyme; it counts words.

[Literacy: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RECAP = `SAY:
- One sheet today: your Cinquain Plan. It has a box for each of the five lines and a word bank.
- The back has a challenge version for anyone who wants a harder puzzle.
- Keep your draft -- we publish it later this week.

DO:
- Print the Cinquain Plan, one per student.
- Have the Poets Toolbox from Week 1 nearby.

[Literacy: Resources | VTLM 2.0: Enabling Learning]`;

const NOTES_LAUNCH = `SAY:
- Listen to this little poem. It is only five lines, and it is about a seabird.
- "Seabird. Spray-spattered. High-sky soaring. I'd like to fly. Wings."
- Can you feel how it grows then shrinks? One word, a bit more, a bit more, then back to one.
- That shape is a cinquain. Let us find out how it is built.

DO:
- Read the cinquain aloud, slowly.
- Quick partner share: "What did you picture? What is the shape doing?"
- Cold call two students.

TEACHER NOTES:
This is a supplied example, used exactly. Let students feel the expanding-then-contracting shape before naming the rule.

WATCH FOR:
- Students who notice the poem grows then shrinks -- celebrate, that is the cinquain shape.

[Literacy: Launch | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me. Today we read a cinquain, work out its shape, and write one.
- The first "I can" is for everyone -- count the words in each line of a cinquain.

DO:
- Choral read the LI and success criteria.
- Point out the shape: 1, 2, 3, 4, then back to 1 word.

TEACHER NOTES:
SC1 (all): follow the 1-2-3-4-1 word pattern. SC2 (target): choose precise describing and action words. SC3 (depth): the title and last word connect in meaning, or use the syllable version. Exit ticket targets SC1.

WATCH FOR:
- Students worried about counting -- the plan has a box per line to make it concrete.

[Literacy: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO = `SAY:
- Here is how a cinquain is built. Five lines, and we count the words.
- Line one is one word -- the title. Watch me: "Tiger".
- Line two is two words that describe it: "Striped, silent".
- Line three is three words showing action: "Stalking through grass".
- Line four is four words showing a feeling: "I feel its power".
- Line five is one word, like the title again: "Hunter".
- See the shape grow and shrink? One, two, three, four, one.

DO:
- Display the model and read it aloud.
- Point to each line and count the words on your fingers.
- Name each line's job: title, describe, action, feeling, final word.

TEACHER NOTES:
Use the clean Tiger model so the word counts are obvious. Count words out loud for every line -- this is the core skill.

WATCH FOR:
- Students who write phrases longer than the count -- model trimming to exactly the right number of words.

[Literacy: I Do | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_WEDO = `SAY:
- Let us build a cinquain together. Our title will be one word -- let us choose "Storm".
- Line two, two describing words. Call some out.
- Line three, three action words. Line four, four feeling words. Line five, one word like the title.
- We will count every line together as we go.

DO:
- Build the cinquain live, counting words for each line.
- Push for strong, precise words, not just any words.
- Read the finished class cinquain aloud.

TEACHER NOTES:
Guided practice in counting and word choice at once. If a line is over, trim it together.

WATCH FOR:
- Students who offer a full sentence for line four -- help them shape exactly four feeling words.

[Literacy: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_CFU = `SAY:
- Quick check. Line three of a cinquain needs exactly THREE words showing action.
- Read both options. Which one fits line three?
- Show me on your fingers: one for A, two for B.

DO:
- Display both options.
- Show Me Fingers.
- Scan: most should choose B.

CFU CHECKPOINT:
Technique: Show Me Fingers (1 for A, 2 for B)
Script:
- "Which option fits line three -- three action words?"
- Scan for: students count and choose B.
PROCEED (>=80%): Release to You Do.
PIVOT (<80%): Use the optional re-teach slide that follows -- it counts each line of a fresh cinquain on fingers.

TEACHER NOTES:
A is one word. B is three words (tumbling, down, fast). B fits line three.

WATCH FOR:
- Students who pick A because it "sounds poetic" -- redirect to counting the words.

[Literacy: CFU | VTLM 2.0: Formative Assessment]`;

const NOTES_CFU_REVEAL = `SAY:
- Line three needs three action words, so the answer is B: "Tumbling down fast".
- A is only one word -- that would fit line one or line five, not line three.

DO:
- Display the reveal.
- Count the three words in B on your fingers.

[Literacy: CFU Reveal | VTLM 2.0: Formative Assessment]`;

const NOTES_RETEACH = `SAY:
- Let us count a whole cinquain together, line by line, on our fingers.
- Line one: "Rain" -- one finger. Line two: "Soft, grey" -- two.
- Line three: "Falling on rooftops" -- three. Line four: "It calms my mind" -- four.
- Line five: "Drizzle" -- one. One, two, three, four, one. That is the shape.

DO:
- Display the re-teach slide and count each line on fingers with the class.
- Re-check: students write line two (two describing words) for a cinquain about the sun.

TEACHER NOTES:
OPTIONAL. Use only if the CFU was below 80%. Different approach: a fully counted model plus an immediate single-line write.

WATCH FOR:
- Students who still miscount -- have them touch a finger for each word as they whisper it.

[Literacy: Optional Re-teach | VTLM 2.0: Reteach with Different Approach]`;

const NOTES_YOUDO = `SAY:
- Your turn to write a cinquain. Choose a one-word title you can picture.
- Then build the shape: two describing words, three action words, four feeling words, one final word.
- Count every line. The plan has a box for each one.
- Fifteen minutes. I will come around.

DO:
- Distribute the Cinquain Plan.
- Circulate. Prioritise students who needed the re-teach.
- Conference by counting a line together.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the word bank. Pick your title, then choose describing, action and feeling words from the lists.
- Extra Notes: The word bank removes the blank-page block and still meets SC1.
EXTENDING PROMPT:
- Task: Try the challenge cinquain on the back -- count SYLLABLES instead: 2, 4, 6, 8, 2.
- Extra Notes: This is the traditional cinquain. It connects to the syllable counting from haiku week.

TEACHER NOTES:
Students keep their cinquain draft to publish in Session 3. Note strong titles and final words.

WATCH FOR:
- Students who write five long lines -- remind them to count and trim each line.

[Literacy: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket. The cinquain title is "Ocean".
- Write LINE 4 -- four words that show a feeling about the ocean.
- Exactly four words. Two minutes.

DO:
- Two minutes, silent.
- Collect.

TEACHER NOTES:
Exit ticket targets SC1 -- following the word count. Exactly four feeling words shows students understand line four.

WATCH FOR:
- Lines that describe rather than feel -- a sign to revisit the difference next session.

[Literacy: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Self-check. Show me on your fingers, one to five, for each "I can".
- Then read your cinquain to a partner and have them check the word counts.

DO:
- Run the fingers check for each success criterion.
- Sixty seconds partner share with counting.
- Preview: "Next session we write a colour poem using all five senses."

TEACHER NOTES:
Note students who found the counting hard so you can support them in the next form.

[Literacy: Closing | VTLM 2.0: Review and Reflect]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Poetry Week 2 Session 1 -- Cinquain";

  // SLIDE 1 -- Title
  titleSlide(
    pres,
    "The Cinquain",
    "A five-line poem that grows then shrinks -- read it, then write it",
    "Week 2 Session 1  |  Year 6 Literacy  |  Poetry",
    NOTES_TITLE
  );

  // SLIDE 2 -- Resources
  addResourceSlide(pres, RESOURCE_ITEMS, T, FOOTER, NOTES_RECAP);

  // SLIDE 3 -- Launch (read the supplied cinquain)
  quoteSlide(
    pres,
    "Read Aloud",
    "A Cinquain",
    "Seabird\nSpray-spattered\nHigh-sky soaring\nI'd like to fly\nWings",
    "",
    "What did you picture? What is the shape of the poem doing as it goes?",
    NOTES_LAUNCH,
    FOOTER
  );

  // SLIDE 4 -- LI / SC
  liSlide(
    pres,
    [
      "We are learning to read a cinquain, work out its shape, and write our own",
    ],
    [
      "I can follow the cinquain shape: 1, 2, 3, 4, 1 words",
      "I can choose precise describing, action and feeling words",
      "I can make my title and final word connect in meaning",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // SLIDE 5 -- I Do (the cinquain form + clean model)
  modellingSlide(
    pres,
    "I Do",
    "How a Cinquain Is Built -- Count the Words",
    "Five lines. Count the words:\n\n- Line 1 = 1 word (title)\n- Line 2 = 2 words (describe)\n- Line 3 = 3 words (action)\n- Line 4 = 4 words (a feeling)\n- Line 5 = 1 word (like the title)",
    "Tiger\n\nStriped, silent\n\nStalking through grass\n\nI feel its power\n\nHunter",
    NOTES_IDO,
    FOOTER
  );

  // SLIDE 6 -- We Do (build a cinquain together)
  boardBuildSlide(
    pres,
    "We Do",
    "Build a Cinquain Together -- Title: Storm",
    "Call out each line. We count the words: 1 - 2 - 3 - 4 - 1.",
    NOTES_WEDO,
    FOOTER,
    {
      promptText: "Teacher drafts each line here. Count the words for every line with the class.",
      prefilledHints: ["1", "2", "3", "4", "1"],
      badgeColor: C.SECONDARY,
    }
  );

  // SLIDE 7 + 8 -- CFU (word count) with reveal
  withReveal(
    () => cfuSlide(
      pres,
      "CFU",
      "Which Fits Line Three?",
      "Show Me Fingers: 1 (A) or 2 (B)",
      "Line 3 needs THREE words showing action.\n\nA)  Falling\n\nB)  Tumbling down fast",
      NOTES_CFU,
      FOOTER
    ),
    (slide) => {
      addRevealAnswerBar(
        slide,
        "B fits line 3 -- three action words: tumbling / down / fast.",
        { color: C.SUCCESS, label: "Answer", showTickAndFix: false }
      );
      slide.addNotes(NOTES_CFU_REVEAL);
    }
  );

  // SLIDE 9 -- Optional Re-teach
  modellingSlide(
    pres,
    "Optional Re-teach",
    "Count a Whole Cinquain on Your Fingers",
    "One word for each finger:\n\n- Rain (1)\n- Soft, grey (2)\n- Falling on rooftops (3)\n- It calms my mind (4)\n- Drizzle (1)\n\nNow you write line 2 for a\ncinquain about the sun.",
    "Rain\n\nSoft, grey\n\nFalling on rooftops\n\nIt calms my mind\n\nDrizzle",
    NOTES_RETEACH,
    FOOTER
  );

  // SLIDE 10 -- You Do
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "You Do", { color: C.PRIMARY, w: 1.5 });
    addTitle(s, "Write Your Cinquain");

    addCard(s, 0.5, CONTENT_TOP, 9, 1.7, { strip: C.PRIMARY, fill: C.WHITE });
    s.addText("Use your Cinquain Plan:", {
      x: 0.75, y: CONTENT_TOP + 0.12, w: 8.4, h: 0.32,
      fontSize: 16, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("First:   Choose a one-word title you can picture.\nNext:   Build the shape -- 2 describing, 3 action, 4 feeling words.\nThen:   Finish with one word, and count every line.", {
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
    s.addText("- Stuck? Use the word bank on your plan to choose each line's words.\n- Ready for more? Try the challenge cinquain: count SYLLABLES (2 - 4 - 6 - 8 - 2).", {
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
      "The cinquain title is: Ocean.",
      "Write LINE 4 -- four words that show a feeling. Exactly four words.",
    ],
    NOTES_EXIT,
    FOOTER,
    { assessesSc: 1, title: "Write the Feeling Line" }
  );

  // SLIDE 12 -- Closing
  closingSlide(
    pres,
    {
      reflectionPrompt: "Read your cinquain to a partner. Have them check the word counts.",
      scItems: [
        "I can follow the cinquain shape: 1, 2, 3, 4, 1 words",
        "I can choose precise describing, action and feeling words",
        "I can make my title and final word connect in meaning",
      ],
      selfAssessment: {
        prompt: "Show on your fingers: 1 (need help) to 5 (could teach it)",
        options: ["1-2", "3", "4-5"],
      },
    },
    NOTES_CLOSING
  );

  // ---- PDF: Cinquain Plan ---------------------------------------------------
  const pl = createPdf({ title: PLAN_RESOURCE.name });
  let py = addPdfHeader(pl, "Cinquain Plan -- Count the Words", {
    color: C.PRIMARY,
    subtitle: "Five lines: 1, 2, 3, 4, 1 words.",
    lessonInfo: "Week 2 Session 1 | Year 6 Literacy | Poetry",
    showNameDate: true,
  });

  py = addTipBox(pl, "Write each line, then count the words in the box. Choose words that paint a picture.", py, { color: C.PRIMARY });

  py = addSectionHeading(pl, "Line 1 -- 1 word (the title)", py, { color: C.PRIMARY });
  py = addLinedArea(pl, py, 1, { lineSpacing: 24 });
  py = addSectionHeading(pl, "Line 2 -- 2 words (describe it)", py, { color: C.SECONDARY });
  py = addLinedArea(pl, py, 1, { lineSpacing: 24 });
  py = addSectionHeading(pl, "Line 3 -- 3 words (action)", py, { color: C.PRIMARY });
  py = addLinedArea(pl, py, 1, { lineSpacing: 24 });
  py = addSectionHeading(pl, "Line 4 -- 4 words (a feeling)", py, { color: C.SECONDARY });
  py = addLinedArea(pl, py, 1, { lineSpacing: 24 });
  py = addSectionHeading(pl, "Line 5 -- 1 word (like the title)", py, { color: C.ACCENT });
  py = addLinedArea(pl, py, 1, { lineSpacing: 24 });
  py += 8;

  py = addSectionHeading(pl, "Word bank -- borrow if you are stuck", py, { color: C.PRIMARY });
  py = addBodyText(pl, "Describing: silent, golden, jagged, restless, gentle, ancient, frozen, glowing", py, { fontSize: 11 });
  py = addBodyText(pl, "Action: racing, crashing, drifting, towering, whispering, melting, soaring, trembling", py, { fontSize: 11 });
  py = addBodyText(pl, "Feeling: it fills me, I feel so, makes me want, leaves me feeling", py, { fontSize: 11 });

  addPdfFooter(pl, "Week 2 Session 1 | Cinquain Plan -- Page 1");

  pl.addPage();
  let py2 = addPdfHeader(pl, "My Cinquain + Challenge", {
    color: C.PRIMARY,
    subtitle: "Write your cinquain, then try the syllable version.",
    lessonInfo: "Week 2 Session 1 | Year 6 Literacy",
    showNameDate: false,
  });
  py2 = addTipBox(pl, "Write your finished cinquain here. Read it aloud and count each line.", py2, { color: C.SECONDARY });
  py2 = addLinedArea(pl, py2, 5, { lineSpacing: 26 });
  py2 += 10;
  py2 = addSectionHeading(pl, "Challenge -- syllable cinquain (2 - 4 - 6 - 8 - 2 beats)", py2, { color: C.ACCENT });
  py2 = addBodyText(pl, "Same five lines, but count BEATS not words. Clap to check.", py2, { fontSize: 10, italic: true });
  py2 = addLinedArea(pl, py2, 5, { lineSpacing: 26 });

  addPdfFooter(pl, "Week 2 Session 1 | My Cinquain -- Page 2");

  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Poetry_W2S1.pptx` }),
    writePdf(pl, PLAN_PDF_PATH),
  ]);

  console.log("PPTX written to " + `${OUT_DIR}/Poetry_W2S1.pptx`);
  console.log("Done: " + PLAN_RESOURCE.name);
}

build().catch((err) => {
  console.error(err);
  process.exit(1);
});
