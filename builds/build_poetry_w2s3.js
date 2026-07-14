"use strict";

// Poetry Unit (Year 6) - Week 2 Session 3: Publishing Both Poems (Cinquain + Colour Poem)
// Focus: VC2E6LY10 -- re-read and edit using the agreed criteria, then publish neatly.
// Catch-up design: "publish whatever you have finished"; the editing checklist recurs all unit.
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

const SESSION_NUMBER = 3;
const FOOTER = "Poetry | Week 2 Session 3 | Year 6 Literacy";
const OUT_DIR = "output/Poetry_W2S3_Publish";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const PUBLISH_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Publish Your Poems",
  "Editing checklist plus two bordered final-copy frames for the cinquain and the colour poem."
);
const RESOURCE_ITEMS = [PUBLISH_RESOURCE];
const PUBLISH_PDF_PATH = path.join(OUT_DIR, PUBLISH_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Today we publish this week's two poems -- the cinquain and the colour poem.
- Publishing means we improve first, then write a proud, neat final copy.
- If you only finished one poem, that is fine. Publish what you have.

DO:
- Display the title slide.
- Have students' cinquain and colour poem drafts ready.

TEACHER NOTES:
Editing-and-presenting session. The same agreed criteria from Week 1 apply, so the routine is familiar.

WATCH FOR:
- Students who think publishing is just copying -- stress that we improve first.

[Literacy: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- One sheet today: Publish Your Poems. The top is our editing checklist; below are two frames.
- Keep your published poems safe -- they go in our poetry booklet in Week 5.

DO:
- Print Publish Your Poems, one per student.
- Have this week's drafts plus the Poets Toolbox ready.
- Have coloured pencils for small illustrations.

[Literacy: Resources | VTLM 2.0: Enabling Learning]`;

const NOTES_LAUNCH = `SAY:
- A quick reminder of what publishing means. Look at these two.
- The draft has good ideas but loose words. The published poem has been checked and the best words chosen.
- We do not start again -- we keep our idea and sharpen it. Our checklist shows us how.

DO:
- Display the draft-versus-published comparison.
- Talk through one improvement between the two.
- Connect to the checklist coming up.

TEACHER NOTES:
The launch re-establishes the publishing routine and is a catch-up point for anyone new to the unit.

WATCH FOR:
- Students attached to a rough draft -- reassure: editing keeps your ideas, it sharpens them.

[Literacy: Launch | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me. Today we edit our poems against the checklist, then publish them.
- The first "I can" is for everyone -- check a poem against the list.

DO:
- Choral read the LI and success criteria.
- Point to the checklist slide -- our agreed criteria.

TEACHER NOTES:
SC1 (all): check a poem against the criteria. SC2 (target): improve at least one word or line. SC3 (depth): explain why a change makes the poem stronger. Exit ticket targets SC2.

WATCH FOR:
- Students who tick boxes without changing anything -- push for one real improvement.

[Literacy: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_CHECKLIST = `SAY:
- Our agreed criteria -- the same checklist we always use before publishing.
- Does it follow the form? Is there a strong image? Are the best words chosen?
- Read it aloud -- does it sound right? Is it neat?
- For a cinquain, "form" means the word counts. For a colour poem, it means a line per sense.

DO:
- Display the checklist and read each point aloud.
- Remind students they will use this exact list on their own poems.

TEACHER NOTES:
Same criteria as Week 1, which supports catch-up and consistency.

WATCH FOR:
- Students who skip "read it aloud" -- insist on it; it catches the most.

[Literacy: Content / Criteria | VTLM 2.0: Explicit Teaching / Success Criteria]`;

const NOTES_IDO = `SAY:
- Watch me edit a cinquain using the checklist. Here is my rough one.
- Form check: counts are right. But the words are weak -- "big" and "goes" make no picture.
- Words check: "big" becomes "towering", "goes" becomes "thunders".
- Image check: line four can show a feeling more strongly. Read it aloud -- much better. Same shape, stronger words.

DO:
- Display the before and after.
- Make each edit visible and tie it to a checklist point.
- Read the published version aloud.

TEACHER NOTES:
Model that the form stays the same while the words sharpen. Tie each change to a checklist point.

WATCH FOR:
- Students who break the word count while editing -- remind them to recount after a change.

[Literacy: I Do | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_WEDO = `SAY:
- Now we edit a colour poem line together. Here is a draft line.
- Run the checklist: is it a specific picture? Is there a stronger word? Could we add an image?
- Tell me one change that would make it stronger.

DO:
- Display the draft line.
- Work the checklist as a class to sharpen it.
- Take one improvement, then reveal an improved version.

CFU CHECKPOINT:
Technique: Cold call against the checklist
Script:
- "Which checklist point can make this line stronger, and how?"
- Scan for: students name a point and suggest a real change.
PROCEED (>=80%): Release to You Do.
PIVOT (<80%): Re-model one edit slowly, naming the checklist point, then take a second class suggestion.

TEACHER NOTES:
The draft line is vague. Guide students to add a specific sensory image.

WATCH FOR:
- Students who only fix spelling -- widen them to imagery and word choice.

[Literacy: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_REVEAL = `SAY:
- Here is our improved line. We kept the idea but made the picture specific.
- One strong change, not a rewrite.

DO:
- Display the improved line.
- Reread it aloud.

[Literacy: We Do Reveal | VTLM 2.0: Guided Practice]`;

const NOTES_YOUDO = `SAY:
- Your turn to publish. Take your cinquain and your colour poem from this week.
- Run the checklist on each. Make at least one real improvement.
- Then write your neat final copy in the frames and add a small illustration.
- Only have one poem? Publish that one beautifully.

DO:
- Distribute the Publish Your Poems sheet.
- Circulate. Help each student find one strong improvement before the neat copy.
- Encourage reading aloud before the final copy.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the checklist one point at a time. Pick ONE word to make stronger, then write your neat copy.
- Extra Notes: One genuine improvement meets SC1 and SC2.
EXTENDING PROMPT:
- Task: Improve two lines and write a sentence explaining why each change makes the poem stronger.
- Extra Notes: The explanation hits the depth criterion.

TEACHER NOTES:
Published poems join the class poetry folder, building toward the Week 5 booklet. Publish whatever each student has finished.

WATCH FOR:
- Students who copy first -- redirect: improve first, then write the final version.

[Literacy: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket. Write down ONE change you made to a poem today, and why it made it stronger.
- One change, one reason. Two minutes.

DO:
- Two minutes, silent.
- Collect.

TEACHER NOTES:
Exit ticket targets SC2 and touches SC3 -- making and justifying an improvement.

WATCH FOR:
- Students who cannot name a change -- a sign they copied without editing.

[Literacy: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Self-check. Show me on your fingers, one to five, for each "I can".
- Then read one published poem to a partner with your proudest voice.

DO:
- Run the fingers check for each success criterion.
- Sixty seconds partner share -- a proud read-aloud.
- Celebrate: "Two more published poems for your collection."

TEACHER NOTES:
Close Week 2 by acknowledging the growing poetry folder. Keep published poems safe for the booklet.

[Literacy: Closing | VTLM 2.0: Review and Reflect]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Poetry Week 2 Session 3 -- Publishing Both Poems";

  // SLIDE 1 -- Title
  titleSlide(
    pres,
    "Publish Your Poems",
    "Edit the cinquain and colour poem, then make a proud final copy",
    "Week 2 Session 3  |  Year 6 Literacy  |  Poetry",
    NOTES_TITLE
  );

  // SLIDE 2 -- Resources
  addResourceSlide(pres, RESOURCE_ITEMS, T, FOOTER, NOTES_RESOURCES);

  // SLIDE 3 -- Launch (draft vs published)
  modellingSlide(
    pres,
    "Launch",
    "Draft to Published -- a Reminder",
    "A DRAFT has good ideas,\nbut the words can be sharper.\n\nA PUBLISHED poem has been\nchecked, improved and set\nout neatly.\n\nWe keep our idea -- we just\nsharpen it.",
    "Draft:\nblue\ncold and quiet\nit feels nice\n\nPublished:\nBlue\nstill, silver\nit calms my mind",
    NOTES_LAUNCH,
    FOOTER
  );

  // SLIDE 4 -- LI / SC
  liSlide(
    pres,
    [
      "We are learning to EDIT our poems against the checklist and then publish them neatly",
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
      "Form: does it follow the shape? (word counts, or a line per sense)",
      "Image: is there at least one simile, metaphor or personification?",
      "Words: are the best, most precise words chosen?",
      "Sound: read it aloud -- does it sound right?",
      "Neat: spelling, capitals and layout tidy?",
    ],
    NOTES_CHECKLIST,
    FOOTER
  );

  // SLIDE 6 -- I Do (edit a cinquain against the checklist)
  modellingSlide(
    pres,
    "I Do",
    "Watch Me Edit a Cinquain",
    "My draft cinquain:\ncounts are right,\nbut the words are weak.\n\nChecklist fixes:\n\n- Words: big -> towering,\n   goes -> thunders\n\n- Recount after each change.",
    "Before:\nStorm\nbig, loud\nit goes fast\nit makes me scared\nRain\n\nAfter:\nStorm\ntowering, wild\nit thunders closer\nit fills me with awe\nTempest",
    NOTES_IDO,
    FOOTER
  );

  // SLIDE 7 + 8 -- We Do (edit a colour poem line) with reveal
  withReveal(
    () => modellingSlide(
      pres,
      "We Do",
      "Together: Improve This Colour Line",
      "Run the checklist:\n\n1. A specific picture?\n\n2. A stronger word?\n\n3. Could we add an image\n   (a simile or metaphor)?\n\nName ONE change that\nwould make it better.",
      "Draft line:\n\nYellow is a happy colour\nthat looks really bright.",
      NOTES_WEDO,
      FOOTER
    ),
    (slide) => {
      addRevealAnswerBar(
        slide,
        "Improved: \"Yellow is the bright laugh of the morning sun.\" -- a specific image.",
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
    s.addText("Take your cinquain and your colour poem:", {
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
    lessonInfo: "Week 2 Session 3 | Year 6 Literacy | Poetry",
    showNameDate: true,
  });

  pby = addTipBox(pb, "Improve first, then write neatly. Read each poem aloud before your final copy.", pby, { color: C.PRIMARY });

  pby = addSectionHeading(pb, "Editing checklist -- tick when done", pby, { color: C.SECONDARY });
  pby = addBodyText(pb, "[  ]  Form: it follows the shape (word counts, or a line per sense)", pby, { fontSize: 11 });
  pby = addBodyText(pb, "[  ]  Image: at least one simile, metaphor or personification", pby, { fontSize: 11 });
  pby = addBodyText(pb, "[  ]  Words: the best, most precise words chosen", pby, { fontSize: 11 });
  pby = addBodyText(pb, "[  ]  Sound: I read it aloud and it sounds right", pby, { fontSize: 11 });
  pby = addBodyText(pb, "[  ]  Neat: spelling, capitals and layout tidy", pby, { fontSize: 11 });
  pby += 10;

  pby = addSectionHeading(pb, "Published poem 1 -- Cinquain", pby, { color: C.PRIMARY });
  pby = drawFinalCopyFrame(pb, pby, 150);
  pby += 10;

  pby = addSectionHeading(pb, "Published poem 2 -- Colour Poem", pby, { color: C.ACCENT });
  pby = drawFinalCopyFrame(pb, pby, 150);

  addPdfFooter(pb, "Week 2 Session 3 | Publish Your Poems");

  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Poetry_W2S3.pptx` }),
    writePdf(pb, PUBLISH_PDF_PATH),
  ]);

  console.log("PPTX written to " + `${OUT_DIR}/Poetry_W2S3.pptx`);
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

build().catch((err) => {
  console.error(err);
  process.exit(1);
});
