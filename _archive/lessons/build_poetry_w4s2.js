"use strict";

// Poetry Unit (Year 6) - Week 4 Session 2: Drafting a Clerihew
// Form: clerihew (4 lines, AABB, names a person, funny).
// Catch-up: opens with a 20-second recap. Sensitivity: keep it kind -- funny, not mean.
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
const FOOTER = "Poetry | Week 4 Session 2 | Year 6 Literacy";
const OUT_DIR = "output/Poetry_W4S2_Draft_Clerihew";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const PLAN_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Clerihew Plan",
  "Clerihew planning template (4 lines, AABB) with name-rhyme families and a kindness reminder."
);
const RESOURCE_ITEMS = [PLAN_RESOURCE];
const PLAN_PDF_PATH = path.join(OUT_DIR, PLAN_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Today we write a clerihew -- a funny four-line poem about a person.
- The trick is the rhyme: line two has to rhyme with the person's name.
- One important rule for us: we keep it KIND. Funny, never mean.

DO:
- Display the title slide.
- Set the tone: clever and kind humour, not put-downs.

TEACHER NOTES:
The form is easy and fun. The main thing to manage is kindness -- steer students to made-up people, characters, or themselves.

WATCH FOR:
- Students who want to write about a classmate unkindly -- redirect early to a safe subject.

[Literacy: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RECAP = `SAY:
- Quick reminder, in twenty seconds. A clerihew has four lines and rhymes AABB.
- Line one names a person. Line two rhymes with the person's name.
- Lines three and four rhyme together. And it should make people smile.

DO:
- Display the recap and read the example aloud.
- Point to the AABB labels and the name in line one.
- This is your catch-up moment for anyone who missed Session 1.

TEACHER NOTES:
Re-establish the four features: 4 lines, AABB, names a person, funny.

WATCH FOR:
- Students who forget line 2 must rhyme with the NAME -- highlight that link.

[Literacy: Launch / Recap | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me. Today we WRITE a clerihew about a person, and keep it kind.
- The first "I can" is for everyone -- name a person and rhyme line two with their name.

DO:
- Choral read the LI and success criteria.
- Remind students: kind and funny, about a made-up person, a character, or themselves.

TEACHER NOTES:
SC1 (all): name a person and make line 2 rhyme with the name. SC2 (target): a full AABB clerihew that is funny. SC3 (depth): a surprise or clever twist in the last line. Exit ticket targets SC1.

WATCH FOR:
- Students who pick a real classmate -- redirect to a safe, kind subject.

[Literacy: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO = `SAY:
- Watch me write a clerihew. I will use a made-up neighbour so it stays kind and funny.
- Line one names him: "My neighbour, old Mr. McGee".
- Line two must rhyme with McGee. Tree, sea, bee, free... "tree" makes a funny picture: "Once got his beard stuck in a tree."
- Lines three and four rhyme: "He hung there all day, with nothing to say." Day and say rhyme.
- See the shape? Name, rhyme the name, then a rhyming pair -- and a silly picture.

DO:
- Display the model and read it with comic timing.
- Show the name-rhyme hunt live: list rhymes for McGee, then choose one.
- Label the lines A, A, B, B.

TEACHER NOTES:
Model with a clearly made-up person so kindness is built in. The key move is finding a rhyme for the name.

WATCH FOR:
- Students who cannot rhyme their chosen name -- the re-teach has name-rhyme families.

[Literacy: I Do | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_WEDO = `SAY:
- Let us write a clerihew together about a made-up character -- say, a clumsy wizard named Lee.
- Line one names him. Line two must rhyme with Lee -- tree, sea, key, bee...
- Then a rhyming pair for lines three and four, with a funny picture.
- We will keep it silly and kind.

DO:
- Build the clerihew live.
- List rhymes for the name, then choose one that makes a funny picture.
- Read the finished class clerihew aloud.

TEACHER NOTES:
Guided practice with a safe, made-up subject. The name rhyme is the key challenge.

WATCH FOR:
- Students who lose the AABB -- read it back so they hear the pairs.

[Literacy: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_CFU = `SAY:
- Quick check. Our person is named Sue. Line two MUST rhyme with Sue.
- Read both options. Which line two works as a clerihew line?
- Show me on your fingers: one for A, two for B.

DO:
- Display both options.
- Show Me Fingers.
- Scan: most should choose B.

CFU CHECKPOINT:
Technique: Show Me Fingers (1 for A, 2 for B)
Script:
- "Which line two rhymes with Sue?"
- Scan for: students choose B (blue rhymes with Sue).
PROCEED (>=80%): Release to You Do.
PIVOT (<80%): Use the optional re-teach slide that follows -- name-rhyme families.

TEACHER NOTES:
A is funny but does not rhyme with Sue. B rhymes (blue) and keeps it silly. B is the answer.

WATCH FOR:
- Students who pick A because they like it -- remind them line 2 must rhyme with the name.

[Literacy: CFU | VTLM 2.0: Formative Assessment]`;

const NOTES_CFU_REVEAL = `SAY:
- The answer is B: "She painted her cat bright blue." Blue rhymes with Sue.
- A was funny but did not rhyme with the name -- so it breaks the clerihew rule.

DO:
- Display the reveal.
- Read the rhyming pair aloud.

[Literacy: CFU Reveal | VTLM 2.0: Formative Assessment]`;

const NOTES_RETEACH = `SAY:
- The hardest part is rhyming the name. Here is the trick: list the name's rhyme family first.
- "Sue": blue, two, zoo, glue, flew, true.
- "Fred": bed, red, head, said, bread.
- "Claire": hair, bear, chair, air, square.
- Pick a name, list its rhymes, THEN write a funny line two that ends on one of them.

DO:
- Display the re-teach slide and read each name-rhyme family.
- Re-check: I give the name "Jack"; students list its rhymes (back, pack, snack, track).

TEACHER NOTES:
OPTIONAL. Use only if the CFU was below 80%. Different approach: build the rhyme family for the name before writing the line.

WATCH FOR:
- Students who choose a hard-to-rhyme name -- suggest swapping to an easier one.

[Literacy: Optional Re-teach | VTLM 2.0: Reteach with Different Approach]`;

const NOTES_YOUDO = `SAY:
- Your turn to write a clerihew. Choose a KIND, safe subject: a made-up person, a book character, a famous explorer, or even yourself.
- Line one names them. Line two rhymes with the name. Lines three and four rhyme and are funny.
- Keep it kind -- we want people to laugh WITH, not laugh at.
- Fifteen minutes. I will come around.

DO:
- Distribute the Clerihew Plan with its name-rhyme families.
- Circulate. Prioritise students who needed the re-teach.
- Conference: "Read me your name line. What rhymes with that name?"

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Choose a name from the plan with its rhymes ready, then write a funny line two that ends on a rhyme.
- Extra Notes: A name plus a rhyming line two meets SC1.
EXTENDING PROMPT:
- Task: Write a second clerihew, or write one about a famous person from history with a clever twist in the last line.
- Extra Notes: Connects humour to knowledge and a surprise ending.

TEACHER NOTES:
Students keep their clerihew draft to publish in Session 4. Keep all subjects kind and safe.

WATCH FOR:
- Students who drift into unkind humour about a real peer -- redirect to a made-up or famous subject.

[Literacy: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket. Write the FIRST TWO LINES of a clerihew.
- Line one names a person. Line two must rhyme with that person's name.
- Two lines. Keep it kind. Two minutes.

DO:
- Two minutes, silent.
- Collect.

TEACHER NOTES:
Exit ticket targets SC1 -- naming a person and rhyming line two with the name, the trickiest part of the form.

WATCH FOR:
- Lines where line 2 does not rhyme with the name -- note who needs the rhyme-family support.

[Literacy: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Self-check. Show me on your fingers, one to five, for each "I can".
- Then read your clerihew to a partner -- did it make them smile?

DO:
- Run the fingers check for each success criterion.
- Sixty seconds partner share.
- Preview: "Next session we write a limerick -- five bouncy lines."

TEACHER NOTES:
Note students who found the name rhyme hard so you can support them next.

[Literacy: Closing | VTLM 2.0: Review and Reflect]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Poetry Week 4 Session 2 -- Drafting a Clerihew";

  // SLIDE 1 -- Title
  titleSlide(
    pres,
    "Write Your Own Clerihew",
    "A funny four-line poem about a person -- and we keep it kind",
    "Week 4 Session 2  |  Year 6 Literacy  |  Poetry",
    NOTES_TITLE
  );

  // SLIDE 2 -- Resources
  addResourceSlide(pres, RESOURCE_ITEMS, T, FOOTER, NOTES_RESOURCES());

  // SLIDE 3 -- Recap / Launch
  modellingSlide(
    pres,
    "Quick Recap",
    "The Clerihew -- in 20 Seconds",
    "Four lines. Rhyme AABB.\n\n- Line 1 names a PERSON\n- Line 2 rhymes with the\n   name\n- Lines 3 and 4 rhyme\n- Make it FUNNY and KIND\n\nMissed last session?\nThis is all you need.",
    "Our art teacher, Mr. Shaw,   (A)\nReally knows how to draw.   (A)\nBut his awful paintings   (B)\nHave caused many faintings.   (B)",
    NOTES_RECAP,
    FOOTER
  );

  // SLIDE 4 -- LI / SC
  liSlide(
    pres,
    [
      "We are learning to WRITE a kind, funny clerihew -- four lines, AABB, about a person",
    ],
    [
      "I can name a person and rhyme line two with their name",
      "I can write a full AABB clerihew that is funny",
      "I can end with a surprise or a clever twist",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // SLIDE 5 -- I Do
  modellingSlide(
    pres,
    "I Do",
    "Watch Me Rhyme the Name",
    "My thinking:\n\n1. Name a made-up person\n   -> Mr. McGee\n\n2. Rhymes with McGee:\n   tree, sea, bee, free\n\n3. Pick a FUNNY one\n   -> tree\n\n4. Lines 3 and 4 rhyme.",
    "My neighbour, old Mr. McGee,   (A)\n\nOnce got his beard stuck in a tree.   (A)\n\nHe hung there all day,   (B)\n\nWith nothing to say.   (B)",
    NOTES_IDO,
    FOOTER
  );

  // SLIDE 6 -- We Do
  boardBuildSlide(
    pres,
    "We Do",
    "Build a Clerihew Together -- A Wizard Named Lee",
    "Line 1 names Lee. List rhymes for Lee. Pick a funny one for line 2. Then a rhyming pair.",
    NOTES_WEDO,
    FOOTER,
    {
      promptText: "Teacher writes the name line, lists name rhymes, then the class builds the funny lines.",
      prefilledHints: ["Name (A)", "rhymes (A)", "B", "B"],
      badgeColor: C.SECONDARY,
    }
  );

  // SLIDE 7 + 8 -- CFU with reveal
  withReveal(
    () => cfuSlide(
      pres,
      "CFU",
      "Which Line Two Rhymes With Sue?",
      "Show Me Fingers: 1 (A) or 2 (B)",
      "Line 1:  My cousin, a girl named Sue,\n\nA)  Loved to bake a chocolate cake.\n\nB)  Painted her cat bright blue.",
      NOTES_CFU,
      FOOTER
    ),
    (slide) => {
      addRevealAnswerBar(
        slide,
        "B -- \"blue\" rhymes with Sue. (A is funny but does not rhyme with the name.)",
        { color: C.SUCCESS, label: "Answer", showTickAndFix: false }
      );
      slide.addNotes(NOTES_CFU_REVEAL);
    }
  );

  // SLIDE 9 -- Optional Re-teach
  modellingSlide(
    pres,
    "Optional Re-teach",
    "List the Name's Rhyme Family First",
    "Pick a name, then list\nits rhymes BEFORE you write:\n\n- Sue: blue, two, zoo, glue\n- Fred: bed, red, head, said\n- Claire: hair, bear, chair, air\n\nThen write a funny line 2\nthat ends on a rhyme.",
    "Your turn:\n\nThe name is \"Jack\".\nList its rhymes:\nback, pack, snack, track...",
    NOTES_RETEACH,
    FOOTER
  );

  // SLIDE 10 -- You Do
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "You Do", { color: C.PRIMARY, w: 1.5 });
    addTitle(s, "Write Your Clerihew");

    addCard(s, 0.5, CONTENT_TOP, 9, 1.7, { strip: C.PRIMARY, fill: C.WHITE });
    s.addText("Choose a KIND subject -- a made-up person, a character, or yourself:", {
      x: 0.75, y: CONTENT_TOP + 0.12, w: 8.4, h: 0.32,
      fontSize: 15, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("First:   Line 1 names the person.\nNext:   Line 2 rhymes with the name.\nThen:   Lines 3 and 4 rhyme and add the funny picture.", {
      x: 0.75, y: CONTENT_TOP + 0.52, w: 8.4, h: 1.05,
      fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    const tipY = CONTENT_TOP + 1.85;
    const tipH = SAFE_BOTTOM - tipY;
    addCard(s, 0.5, tipY, 9, tipH, { strip: C.ACCENT, fill: C.BG_CARD });
    s.addText("Keep it kind -- laugh WITH, not AT", {
      x: 0.75, y: tipY + 0.10, w: 8.4, h: 0.30,
      fontSize: 14, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("- Stuck? Use a name with its rhymes ready on your plan.\n- Ready for more? Write a second one, or one about a famous person with a twist ending.", {
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
      "Write the FIRST TWO LINES of a clerihew.",
      "Line 1 names a person. Line 2 must rhyme with that name. Keep it kind.",
    ],
    NOTES_EXIT,
    FOOTER,
    { assessesSc: 1, title: "Name It and Rhyme It" }
  );

  // SLIDE 12 -- Closing
  closingSlide(
    pres,
    {
      reflectionPrompt: "Read your clerihew to a partner. Did it make them smile?",
      scItems: [
        "I can name a person and rhyme line two with their name",
        "I can write a full AABB clerihew that is funny",
        "I can end with a surprise or a clever twist",
      ],
      selfAssessment: {
        prompt: "Show on your fingers: 1 (need help) to 5 (could teach it)",
        options: ["1-2", "3", "4-5"],
      },
    },
    NOTES_CLOSING
  );

  // ---- PDF: Clerihew Plan ---------------------------------------------------
  const pl = createPdf({ title: PLAN_RESOURCE.name });
  let py = addPdfHeader(pl, "Clerihew Plan -- Four Funny Lines", {
    color: C.PRIMARY,
    subtitle: "4 lines, AABB. Name a person, rhyme the name, make it kind and funny.",
    lessonInfo: "Week 4 Session 2 | Year 6 Literacy | Poetry",
    showNameDate: true,
  });

  py = addTipBox(pl, "Keep it kind -- choose a made-up person, a book character, or yourself. Laugh WITH, not AT.", py, { color: C.PRIMARY });

  py = addSectionHeading(pl, "Name-rhyme families -- borrow a name", py, { color: C.PRIMARY });
  py = addBodyText(pl, "Sue: blue, two, zoo, glue, flew, true", py, { fontSize: 11 });
  py = addBodyText(pl, "Fred: bed, red, head, said, bread, sled", py, { fontSize: 11 });
  py = addBodyText(pl, "Claire: hair, bear, chair, air, square, stare", py, { fontSize: 11 });
  py = addBodyText(pl, "Jack: back, pack, snack, track, black, quack", py, { fontSize: 11 });
  py += 8;

  py = addSectionHeading(pl, "Line 1 -- name a person", py, { color: C.PRIMARY });
  py = addLinedArea(pl, py, 1, { lineSpacing: 24 });
  py = addSectionHeading(pl, "Line 2 -- rhyme with the name", py, { color: C.SECONDARY });
  py = addLinedArea(pl, py, 1, { lineSpacing: 24 });
  py = addSectionHeading(pl, "Line 3 -- start the joke", py, { color: C.PRIMARY });
  py = addLinedArea(pl, py, 1, { lineSpacing: 24 });
  py = addSectionHeading(pl, "Line 4 -- rhyme with line 3, land the joke", py, { color: C.ACCENT });
  py = addLinedArea(pl, py, 1, { lineSpacing: 24 });

  addPdfFooter(pl, "Week 4 Session 2 | Clerihew Plan -- Page 1");

  pl.addPage();
  let py2 = addPdfHeader(pl, "My Clerihew -- Write It Here", {
    color: C.PRIMARY,
    subtitle: "Read it aloud. Did it make someone smile?",
    lessonInfo: "Week 4 Session 2 | Year 6 Literacy",
    showNameDate: false,
  });
  py2 = addTipBox(pl, "Write your finished clerihew. Label the lines A, A, B, B at the side.", py2, { color: C.SECONDARY });
  py2 = addLinedArea(pl, py2, 4, { lineSpacing: 28 });
  py2 += 10;
  py2 = addSectionHeading(pl, "Challenge -- a second clerihew, or a famous person with a twist", py2, { color: C.ACCENT });
  py2 = addLinedArea(pl, py2, 4, { lineSpacing: 28 });

  addPdfFooter(pl, "Week 4 Session 2 | My Clerihew -- Page 2");

  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Poetry_W4S2.pptx` }),
    writePdf(pl, PLAN_PDF_PATH),
  ]);

  console.log("PPTX written to " + `${OUT_DIR}/Poetry_W4S2.pptx`);
  console.log("Done: " + PLAN_RESOURCE.name);
}

function NOTES_RESOURCES() {
  return `SAY:
- One sheet today: your Clerihew Plan. It has name-rhyme families and a line for each part.
- There is a kindness reminder at the top -- funny, never mean.
- Keep your draft -- we publish it later this week.

DO:
- Print the Clerihew Plan, one per student.
- Have the Poets Toolbox nearby.

[Literacy: Resources | VTLM 2.0: Enabling Learning]`;
}

build().catch((err) => {
  console.error(err);
  process.exit(1);
});
