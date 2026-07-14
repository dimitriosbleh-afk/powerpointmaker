"use strict";

// Poetry Unit (Year 6) - Week 1 Session 4: Publishing Both Poems
// Focus: VC2E6LY10 -- re-read and edit using agreed criteria, then publish neatly.
// Edits the three-line poem (S2) and haiku (S3) drafts; flexible -- publish whatever is finished.
// Catch-up design: "publish whatever you have"; the editing checklist recurs all unit.
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
  exitTicketSlide, modellingSlide,
} = T;

const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addLinedArea,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const SESSION_NUMBER = 4;
const FOOTER = "Poetry | Week 1 Session 4 | Year 6 Literacy";
const OUT_DIR = "output/Poetry_W1S4_Publish";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const PUBLISH_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Publish Your Poems",
  "Editing checklist plus two bordered final-copy frames to write and illustrate the published poems."
);
const RESOURCE_ITEMS = [PUBLISH_RESOURCE];
const PUBLISH_PDF_PATH = path.join(OUT_DIR, PUBLISH_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Today we publish. That means we make our poems the best they can be, then write a neat, proud final copy.
- Real poets edit before they publish. We will do the same: check, improve, then make it beautiful.
- If you only finished one poem, that is fine. Publish what you have and we keep building.

DO:
- Display the title slide.
- Have students' drafts from Sessions 2 and 3 ready.

TEACHER NOTES:
This is an editing-and-presenting session. The thinking work is the editing against agreed criteria; the neat copy is the reward.

WATCH FOR:
- Students who think publishing just means copying neatly -- stress that we IMPROVE first.

[Literacy: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_LAUNCH = `SAY:
- What is the difference between a draft and a published poem? Look at these two.
- The draft is rough -- it has the idea, but the words could be sharper and it is not laid out for a reader.
- The published one has been checked, the best words chosen, and it is set out neatly.
- Today we move our drafts to published. The secret is our editing checklist.

DO:
- Display the draft-versus-published comparison.
- Talk through one or two improvements between the two.
- Connect to the checklist coming up.

TEACHER NOTES:
The launch makes "publishing" concrete: it is editing plus neat presentation, not just copying. This activates the agreed-criteria idea.

WATCH FOR:
- Students who prefer their rough draft -- reassure: editing keeps your ideas, it just sharpens them.

[Literacy: Launch | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me. Today we EDIT our poems against a checklist, then publish them.
- The first "I can" is for everyone -- check a poem against the list.

DO:
- Choral read the LI and success criteria.
- Point to the checklist slide that follows -- these are our agreed criteria.

TEACHER NOTES:
SC1 (all): check a poem against the criteria. SC2 (target): improve at least one word or line. SC3 (depth): explain why a change makes the poem stronger. Exit ticket targets SC2.

WATCH FOR:
- Students who tick every box without changing anything -- prompt them to find one real improvement.

[Literacy: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_CHECKLIST = `SAY:
- Here are our agreed criteria -- the checklist every poet uses before publishing.
- One: does it follow the form? The right lines, beats or shape.
- Two: is there at least one strong image -- a simile, metaphor or personification?
- Three: are the best words chosen -- precise and full of picture?
- Four: read it aloud. Does it sound right?
- Five: is it neat -- spelling, capitals and layout?

DO:
- Display the checklist and read each point aloud.
- Tell students they will use this exact list on their own poems shortly.

TEACHER NOTES:
This checklist is the agreed criteria for the whole unit. Students will meet it again every publishing session, which supports catch-up.

WATCH FOR:
- Students who rush past "read it aloud" -- it is the most powerful check; insist on it.

[Literacy: Content / Criteria | VTLM 2.0: Explicit Teaching / Success Criteria]`;

const NOTES_IDO = `SAY:
- Watch me edit a draft using the checklist. Here is my rough three-line poem.
- "The bird. In the tree. Goes." It follows the shape, but the words are weak.
- Checklist point three -- best words. "Bird" could be "magpie". "Goes" is dull -- let me make it "sings".
- Checklist point two -- an image. I will make line two paint more: "high in the swaying gum".
- Read it aloud: "The magpie, high in the swaying gum, sings." Much stronger. Same idea, sharper words.

DO:
- Display the before and after.
- Make each edit visible, pointing to the checklist point that prompted it.
- Read the published version aloud with feeling.

TEACHER NOTES:
Model that editing keeps the original idea and improves the words. Tie each change to a checklist point so the process is repeatable.

WATCH FOR:
- Students who think editing means starting over -- show that we kept the same poem, just improved it.

[Literacy: I Do | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_WEDO = `SAY:
- Now we edit one together. Here is a haiku draft. Let us run the checklist.
- Does it have 5-7-5 beats? Clap and check. Is there a clear picture? Which word could be stronger?
- Tell me one change that would make it better. We will improve it together.

DO:
- Display the draft haiku.
- Work the checklist as a class -- count beats, find the image, hunt for a stronger word.
- Take one improvement from the class, then reveal an improved version.

CFU CHECKPOINT:
Technique: Cold call against the checklist
Script:
- "Which checklist point can make this poem stronger, and how?"
- Scan for: students name a point and suggest a real change.
PROCEED (>=80%): Release to You Do.
PIVOT (<80%): Re-model one edit slowly, naming the checklist point, then take a second class suggestion before releasing.

TEACHER NOTES:
The draft has the right beats but a weak last line. Guide students to sharpen the final image.

WATCH FOR:
- Students who only check spelling -- widen them to word choice and imagery.

[Literacy: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_REVEAL = `SAY:
- Here is our improved haiku. We kept the beats, but the last line now paints a clearer picture.
- Notice we did not rewrite the whole thing -- one strong change made it better.

DO:
- Display the improved version.
- Reread it aloud and point to the line that changed.

[Literacy: We Do Reveal | VTLM 2.0: Guided Practice]`;

const NOTES_YOUDO = `SAY:
- Your turn to publish. Take your three-line poem and your haiku from this week.
- Run the checklist on each one. Make at least one real improvement.
- Then write your neat final copy in the frames and add a small illustration.
- If you only have one poem, publish that one beautifully. Quality, not quantity.

DO:
- Distribute the Publish Your Poems sheet.
- Circulate. Help students find one strong improvement per poem before they write the neat copy.
- Encourage reading aloud before the final copy goes down.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the checklist one point at a time. Pick ONE word to make stronger, then write your neat copy.
- Extra Notes: One genuine improvement meets SC1 and SC2 without overwhelming the student.
EXTENDING PROMPT:
- Task: Improve two lines and write a sentence explaining why each change makes the poem stronger.
- Extra Notes: The explanation hits the depth criterion (SC3).

TEACHER NOTES:
Published poems go into the class poetry folder. They build toward the booklet in Week 5. Publish whatever each student has finished.

WATCH FOR:
- Students who make the neat copy first -- redirect: improve first, then write the final version.

[Literacy: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket. Write down ONE change you made to a poem today, and why it made it stronger.
- For example: I changed "goes" to "sings" because it gives a clearer sound.
- One change, one reason. Two minutes.

DO:
- Two minutes, silent.
- Collect.

TEACHER NOTES:
Exit ticket targets SC2 and touches SC3 -- making and justifying an improvement. This shows students understand editing, not just copying.

WATCH FOR:
- Students who cannot name a change -- a sign they copied without editing; support next session.

[Literacy: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Self-check. Show me on your fingers, one to five, for each "I can".
- Then read one published poem to a partner with your proudest voice.

DO:
- Run the fingers check for each success criterion.
- Sixty seconds partner share -- a proud read-aloud.
- Celebrate: "You have written and published your first poems this year."

TEACHER NOTES:
Close Week 1 by acknowledging real progress. Keep published poems safe for the Week 5 booklet.

[Literacy: Closing | VTLM 2.0: Review and Reflect]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Poetry Week 1 Session 4 -- Publishing Both Poems";

  // SLIDE 1 -- Title
  titleSlide(
    pres,
    "Publish Your Poems",
    "Edit against the checklist, then make a proud final copy",
    "Week 1 Session 4  |  Year 6 Literacy  |  Poetry",
    NOTES_TITLE
  );

  // SLIDE 2 -- Resources
  addResourceSlide(pres, RESOURCE_ITEMS, T, FOOTER, NOTES_RESOURCES());

  // SLIDE 3 -- Launch (draft vs published)
  modellingSlide(
    pres,
    "Launch",
    "Draft to Published -- What Changes?",
    "A DRAFT has the idea,\nbut the words can be sharper\nand it is not laid out\nfor a reader.\n\nA PUBLISHED poem has been\nchecked, improved and set\nout neatly.\n\nToday we move from one\nto the other.",
    "Draft:\nthe bird\nin the tree\ngoes\n\nPublished:\nThe magpie\nHigh in the swaying gum\nSings",
    NOTES_LAUNCH,
    FOOTER
  );

  // SLIDE 4 -- LI / SC
  liSlide(
    pres,
    [
      "We are learning to EDIT our poems against a checklist and then publish them neatly",
    ],
    [
      "I can check a poem against our editing checklist",
      "I can improve at least one word or line",
      "I can explain why a change makes the poem stronger",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // SLIDE 5 -- Content: the editing checklist (agreed criteria)
  contentSlide(
    pres,
    "Our Checklist",
    C.SECONDARY,
    "Before You Publish -- Agreed Criteria",
    [
      "Form: does it follow the shape? (lines, beats or pattern)",
      "Image: is there at least one simile, metaphor or personification?",
      "Words: are the best, most precise words chosen?",
      "Sound: read it aloud -- does it sound right?",
      "Neat: spelling, capitals and layout tidy?",
    ],
    NOTES_CHECKLIST,
    FOOTER
  );

  // SLIDE 6 -- I Do (edit a draft against the checklist)
  modellingSlide(
    pres,
    "I Do",
    "Watch Me Edit With the Checklist",
    "My draft:\n\"The bird. In the tree. Goes.\"\n\nChecklist fixes:\n\n- Words: bird -> magpie,\n   goes -> sings\n\n- Image: line 2 paints\n   more -> swaying gum\n\nSame idea, sharper words.",
    "Before:\nThe bird\nIn the tree\nGoes\n\nAfter:\nThe magpie\nHigh in the swaying gum\nSings",
    NOTES_IDO,
    FOOTER
  );

  // SLIDE 7 + 8 -- We Do (edit a haiku together) with reveal
  withReveal(
    () => modellingSlide(
      pres,
      "We Do",
      "Together: Improve This Haiku",
      "Run the checklist:\n\n1. Beats 5 - 7 - 5? Clap it.\n\n2. A clear picture?\n\n3. A stronger word?\n\nName ONE change that\nwould make it better.",
      "Draft haiku:\n\nThe little garden\nWet flowers after the rain\nThey look quite pretty",
      NOTES_WEDO,
      FOOTER
    ),
    (slide) => {
      addRevealAnswerBar(
        slide,
        "Stronger last line: \"Petals catch the light\" -- a clearer picture, still 5 beats.",
        { color: C.SUCCESS, label: "Improved", showTickAndFix: false }
      );
      slide.addNotes(NOTES_WEDO_REVEAL);
    }
  );

  // SLIDE 9 -- You Do
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "You Do", { color: C.PRIMARY, w: 1.5 });
    addTitle(s, "Publish Your Poems");

    addCard(s, 0.5, CONTENT_TOP, 9, 1.7, { strip: C.PRIMARY, fill: C.WHITE });
    s.addText("Take your three-line poem and your haiku:", {
      x: 0.75, y: CONTENT_TOP + 0.12, w: 8.4, h: 0.32,
      fontSize: 16, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("First:   Run the checklist on each poem.\nNext:   Make at least one real improvement.\nThen:   Write your neat final copy and add a small illustration.", {
      x: 0.75, y: CONTENT_TOP + 0.52, w: 8.4, h: 1.05,
      fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    const tipY = CONTENT_TOP + 1.85;
    const tipH = SAFE_BOTTOM - tipY;
    addCard(s, 0.5, tipY, 9, tipH, { strip: C.ACCENT, fill: C.BG_CARD });
    s.addText("Only finished one? Publish that one beautifully.", {
      x: 0.75, y: tipY + 0.10, w: 8.4, h: 0.30,
      fontSize: 14, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("- Stuck? Improve just ONE word, then write your neat copy.\n- Ready for more? Improve two lines and explain why each change helps.", {
      x: 0.75, y: tipY + 0.44, w: 8.4, h: tipH - 0.55,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  }

  // SLIDE 10 -- Exit Ticket
  exitTicketSlide(
    pres,
    [
      "Write ONE change you made to a poem today.",
      "Then write why it made the poem stronger.",
    ],
    NOTES_EXIT,
    FOOTER,
    { assessesSc: 2, title: "One Change, One Reason" }
  );

  // SLIDE 11 -- Closing
  closingSlide(
    pres,
    {
      reflectionPrompt: "Read one published poem to a partner with your proudest voice.",
      scItems: [
        "I can check a poem against our editing checklist",
        "I can improve at least one word or line",
        "I can explain why a change makes the poem stronger",
      ],
      selfAssessment: {
        prompt: "Show on your fingers: 1 (need help) to 5 (could teach it)",
        options: ["1-2", "3", "4-5"],
      },
    },
    NOTES_CLOSING
  );

  // ---- PDF: Publish Your Poems ----------------------------------------------
  const pb = createPdf({ title: PUBLISH_RESOURCE.name });
  let pby = addPdfHeader(pb, "Publish Your Poems", {
    color: C.PRIMARY,
    subtitle: "Edit with the checklist, then write your proud final copy.",
    lessonInfo: "Week 1 Session 4 | Year 6 Literacy | Poetry",
    showNameDate: true,
  });

  pby = addTipBox(pb, "Improve first, then write neatly. Read each poem aloud before your final copy.", pby, { color: C.PRIMARY });

  pby = addSectionHeading(pb, "Editing checklist -- tick when done", pby, { color: C.SECONDARY });
  pby = addBodyText(pb, "[  ]  Form: it follows the shape (lines, beats or pattern)", pby, { fontSize: 11 });
  pby = addBodyText(pb, "[  ]  Image: at least one simile, metaphor or personification", pby, { fontSize: 11 });
  pby = addBodyText(pb, "[  ]  Words: the best, most precise words chosen", pby, { fontSize: 11 });
  pby = addBodyText(pb, "[  ]  Sound: I read it aloud and it sounds right", pby, { fontSize: 11 });
  pby = addBodyText(pb, "[  ]  Neat: spelling, capitals and layout tidy", pby, { fontSize: 11 });
  pby += 10;

  pby = addSectionHeading(pb, "Published poem 1 -- Three-Line Poem", pby, { color: C.PRIMARY });
  pby = drawFinalCopyFrame(pb, pby, 150);
  pby += 10;

  pby = addSectionHeading(pb, "Published poem 2 -- Haiku", pby, { color: C.ACCENT });
  pby = drawFinalCopyFrame(pb, pby, 150);

  addPdfFooter(pb, "Week 1 Session 4 | Publish Your Poems");

  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Poetry_W1S4.pptx` }),
    writePdf(pb, PUBLISH_PDF_PATH),
  ]);

  console.log("PPTX written to " + `${OUT_DIR}/Poetry_W1S4.pptx`);
  console.log("Done: " + PUBLISH_RESOURCE.name);
}

// A bordered frame for the neat final copy + a small illustration corner.
function drawFinalCopyFrame(doc, y, height) {
  const left = doc.page.margins.left;
  const width = doc.page.width - doc.page.margins.left - doc.page.margins.right;
  doc.save();
  doc.lineWidth(1).strokeColor("#888888")
    .roundedRect(left, y, width, height, 8).stroke();
  doc.restore();
  doc.font("Sans").fontSize(8).fillColor("#777777")
    .text("Write your published poem here. Add a small illustration.", left + 10, y + 8, { width: width - 20 });
  doc.fillColor("#000000");
  return y + height + 6;
}

function NOTES_RESOURCES() {
  return `SAY:
- One sheet today: Publish Your Poems.
- The top has our editing checklist. Below are two frames for your neat final copies.
- Keep your published poems safe -- they go in our poetry booklet in Week 5.

DO:
- Print Publish Your Poems, one per student.
- Have students' drafts from Sessions 2 and 3, plus the Poets Toolbox, ready.
- Have coloured pencils ready for the small illustrations.

[Literacy: Resources | VTLM 2.0: Enabling Learning]`;
}

build().catch((err) => {
  console.error(err);
  process.exit(1);
});
