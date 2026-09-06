"use strict";

// Poetry Unit (Year 6) - Week 4 Session 4: Publishing Both Poems (Clerihew + Limerick)
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

const SESSION_NUMBER = 4;
const FOOTER = "Poetry | Week 4 Session 4 | Year 6 Literacy";
const OUT_DIR = "output/Poetry_W4S4_Publish";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const PUBLISH_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Publish Your Poems",
  "Editing checklist plus two bordered final-copy frames for the clerihew and the limerick."
);
const RESOURCE_ITEMS = [PUBLISH_RESOURCE];
const PUBLISH_PDF_PATH = path.join(OUT_DIR, PUBLISH_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Today we publish this week's two funny poems -- the clerihew and the limerick.
- Publishing means we improve first, then write a proud, neat final copy.
- If you only finished one poem, that is fine. Publish what you have.

DO:
- Display the title slide.
- Have students' clerihew and limerick drafts ready.

TEACHER NOTES:
Editing-and-presenting session. The same agreed criteria apply; for funny poems, read aloud is the key check for rhythm and the joke.

WATCH FOR:
- Students who think publishing is just copying -- stress that we improve first.

[Literacy: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- One sheet today: Publish Your Poems. The top is our editing checklist; below are two frames.
- Keep your published poems safe -- they go in our poetry booklet next week.

DO:
- Print Publish Your Poems, one per student.
- Have this week's drafts plus the Poets Toolbox ready.
- Have coloured pencils for small illustrations.

[Literacy: Resources | VTLM 2.0: Enabling Learning]`;

const NOTES_LAUNCH = `SAY:
- A quick reminder of what publishing means. Look at these two.
- The draft is funny but the rhythm stumbles. The published one has been smoothed so it bounces.
- We keep the joke -- we just make it land better. Our checklist shows us how.

DO:
- Display the draft-versus-published comparison.
- Read both aloud so students hear the smoother rhythm.
- Connect to the checklist coming up.

TEACHER NOTES:
For funny poems, "read it aloud" is the most powerful check. The launch makes that concrete.

WATCH FOR:
- Students attached to a clunky draft -- reassure: editing keeps the joke, it sharpens the timing.

[Literacy: Launch | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me. Today we edit our funny poems against the checklist, then publish them.
- The first "I can" is for everyone -- check a poem against the list.

DO:
- Choral read the LI and success criteria.
- Point to the checklist slide -- our agreed criteria.

TEACHER NOTES:
SC1 (all): check a poem against the criteria. SC2 (target): improve at least one word or line. SC3 (depth): explain why a change makes the poem stronger or funnier. Exit ticket targets SC2.

WATCH FOR:
- Students who tick boxes without changing anything -- push for one real improvement.

[Literacy: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_CHECKLIST = `SAY:
- Our agreed criteria -- the same checklist we always use before publishing.
- Does it follow the form? For a clerihew, that is AABB and naming a person. For a limerick, AABBA.
- Is the rhyme strong? Are the best words chosen? Read it aloud -- does it bounce? Is it neat and kind?

DO:
- Display the checklist and read each point aloud.
- Remind students they will use this exact list on their own poems.

TEACHER NOTES:
Same criteria as every week. For funny poems, the "sound" check is about rhythm and the joke landing.

WATCH FOR:
- Students who skip "read it aloud" -- insist on it; it catches a stumbling bounce.

[Literacy: Content / Criteria | VTLM 2.0: Explicit Teaching / Success Criteria]`;

const NOTES_IDO = `SAY:
- Watch me edit a clerihew using the checklist. Here is my rough one.
- Form check: it rhymes AABB and names a person. Good.
- Sound check: read aloud, line three is a bit clunky -- too many small words.
- Words check: I tighten it so the joke lands faster. Read it again -- snappier and funnier.

DO:
- Display the before and after.
- Read both aloud so students hear the joke land better.
- Tie each change to a checklist point.

TEACHER NOTES:
Model tightening for comic timing. For funny poems, fewer, sharper words often make the joke land.

WATCH FOR:
- Students who add MORE words -- show how trimming can make a joke funnier.

[Literacy: I Do | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_WEDO = `SAY:
- Now we edit a limerick line together. Here is a draft line that breaks the bounce.
- Run the checklist: does it keep the rhythm? Does it rhyme? Could a word be sharper?
- Tell me one change that would make it bounce better.

DO:
- Display the draft line.
- Read it aloud to hear the stumble, then fix it as a class.
- Take one improvement, then reveal an improved version.

CFU CHECKPOINT:
Technique: Cold call against the checklist
Script:
- "Which checklist point can make this line bounce better, and how?"
- Scan for: students name a point and suggest a real change.
PROCEED (>=80%): Release to You Do.
PIVOT (<80%): Re-model one edit slowly, reading aloud, then take a second class suggestion.

TEACHER NOTES:
The draft line has an extra word that breaks the rhythm. Guide students to trim it so the bounce returns.

WATCH FOR:
- Students who fix the meaning but ignore the rhythm -- read aloud after each change.

[Literacy: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_REVEAL = `SAY:
- Here is our improved line. We trimmed a word and the bounce came back.
- One small change, big difference to the rhythm.

DO:
- Display the improved line.
- Read it aloud so the bounce is clear.

[Literacy: We Do Reveal | VTLM 2.0: Guided Practice]`;

const NOTES_YOUDO = `SAY:
- Your turn to publish. Take your clerihew and your limerick from this week.
- Run the checklist on each. Read them ALOUD -- does the joke land? Does it bounce?
- Make at least one real improvement, then write your neat final copy and add a small illustration.
- Only have one poem? Publish that one beautifully.

DO:
- Distribute the Publish Your Poems sheet.
- Circulate. Encourage reading aloud to test the rhythm and the joke.
- Help each student find one strong improvement before the neat copy.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Read your poem aloud with a partner and fix ONE word together, then write your neat copy.
- Extra Notes: One genuine improvement meets SC1 and SC2.
EXTENDING PROMPT:
- Task: Improve two lines for sharper timing and write a sentence explaining why each change makes it funnier.
- Extra Notes: The explanation hits the depth criterion.

TEACHER NOTES:
Published poems join the class poetry folder, building toward the Week 5 booklet. Publish whatever each student has finished. Keep all subjects kind.

WATCH FOR:
- Students who copy first -- redirect: improve first, then write the final version.

[Literacy: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket. Write down ONE change you made to a poem today, and why it made it better or funnier.
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
- Then read one published poem to a partner with your funniest voice.

DO:
- Run the fingers check for each success criterion.
- Sixty seconds partner share -- a funny read-aloud.
- Celebrate: "Your poetry collection is nearly ready to become a booklet."

TEACHER NOTES:
Close Week 4 by acknowledging the growing folder. Keep published poems safe for next week's booklet.

[Literacy: Closing | VTLM 2.0: Review and Reflect]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Poetry Week 4 Session 4 -- Publishing Both Poems";

  // SLIDE 1 -- Title
  titleSlide(
    pres,
    "Publish Your Poems",
    "Edit the clerihew and limerick, then make a proud final copy",
    "Week 4 Session 4  |  Year 6 Literacy  |  Poetry",
    NOTES_TITLE
  );

  // SLIDE 2 -- Resources
  addResourceSlide(pres, RESOURCE_ITEMS, T, FOOTER, NOTES_RESOURCES);

  // SLIDE 3 -- Launch (draft vs published)
  modellingSlide(
    pres,
    "Launch",
    "Draft to Published -- a Reminder",
    "A DRAFT can be funny\nbut stumble in rhythm.\n\nA PUBLISHED poem has been\nchecked so the joke lands\nand the beat bounces.\n\nWe keep the joke --\nwe just sharpen the timing.",
    "Draft (clerihew):\nMy friend, a boy we call Ben,\nhe really likes to use a pen\n\nPublished:\nMy friend, a boy named Ben,\nLives happily in a pen.",
    NOTES_LAUNCH,
    FOOTER
  );

  // SLIDE 4 -- LI / SC
  liSlide(
    pres,
    [
      "We are learning to EDIT our funny poems against the checklist and then publish them neatly",
    ],
    [
      "I can check a poem against our editing checklist",
      "I can improve at least one word or line",
      "I can explain why a change makes the poem stronger or funnier",
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
      "Form: does it follow the shape? (clerihew AABB, or limerick AABBA)",
      "Image / joke: is the funny part clear, and is it kind?",
      "Words: are the best, most precise words chosen?",
      "Sound: read it aloud -- does it bounce and does the joke land?",
      "Neat: spelling, capitals and layout tidy?",
    ],
    NOTES_CHECKLIST,
    FOOTER
  );

  // SLIDE 6 -- I Do (edit a clerihew against the checklist)
  modellingSlide(
    pres,
    "I Do",
    "Watch Me Sharpen the Timing",
    "My draft rhymes AABB,\nbut line 3 is clunky --\ntoo many small words.\n\nChecklist fixes:\n\n- Read aloud: the joke is\n   slow to land.\n\n- Trim the words so the\n   punchline is snappy.",
    "Before:\n...and then he tried to draw,\nbut it was really not that good at all,\n\nAfter:\n...and then he tried to draw,\nbut it was the worst we ever saw,",
    NOTES_IDO,
    FOOTER
  );

  // SLIDE 7 + 8 -- We Do (edit a limerick line) with reveal
  withReveal(
    () => modellingSlide(
      pres,
      "We Do",
      "Together: Fix the Bounce",
      "Run the checklist:\n\n1. Does it keep the\n   rhythm? Read aloud.\n\n2. Does it still rhyme?\n\n3. A sharper word?\n\nName ONE change that\nwould make it bounce.",
      "Draft line (breaks the bounce):\n\nWho really liked to eat a great\nbig lot of cheese",
      NOTES_WEDO,
      FOOTER
    ),
    (slide) => {
      addRevealAnswerBar(
        slide,
        "Trimmed: \"Who loved to nibble cheese\" -- the bounce returns.",
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
    s.addText("Take your clerihew and your limerick:", {
      x: 0.75, y: CONTENT_TOP + 0.12, w: 8.4, h: 0.32,
      fontSize: 16, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("First:   Run the checklist on each poem.\nNext:   Read ALOUD -- does the joke land? Make one real improvement.\nThen:   Write your neat final copy and add a small illustration.", {
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
    s.addText("- Stuck? Read it aloud with a partner and fix ONE word.\n- Ready for more? Sharpen two lines and explain why each makes it funnier.", {
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
      "Then write why it made the poem better or funnier.",
    ],
    NOTES_EXIT,
    FOOTER,
    { assessesSc: 2, title: "One Change, One Reason" }
  );

  // SLIDE 11 -- Closing
  closingSlide(
    pres,
    {
      reflectionPrompt: "Read one published poem to a partner with your funniest voice.",
      scItems: [
        "I can check a poem against our editing checklist",
        "I can improve at least one word or line",
        "I can explain why a change makes the poem stronger or funnier",
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
    lessonInfo: "Week 4 Session 4 | Year 6 Literacy | Poetry",
    showNameDate: true,
  });

  pby = addTipBox(pb, "Improve first, then write neatly. Read each poem aloud -- does the joke land and bounce?", pby, { color: C.PRIMARY });

  pby = addSectionHeading(pb, "Editing checklist -- tick when done", pby, { color: C.SECONDARY });
  pby = addBodyText(pb, "[  ]  Form: it follows the shape (clerihew AABB, or limerick AABBA)", pby, { fontSize: 11 });
  pby = addBodyText(pb, "[  ]  Joke: the funny part is clear -- and kind", pby, { fontSize: 11 });
  pby = addBodyText(pb, "[  ]  Words: the best, most precise words chosen", pby, { fontSize: 11 });
  pby = addBodyText(pb, "[  ]  Sound: I read it aloud and it bounces", pby, { fontSize: 11 });
  pby = addBodyText(pb, "[  ]  Neat: spelling, capitals and layout tidy", pby, { fontSize: 11 });
  pby += 10;

  pby = addSectionHeading(pb, "Published poem 1 -- Clerihew", pby, { color: C.PRIMARY });
  pby = drawFinalCopyFrame(pb, pby, 150);
  pby += 10;

  pby = addSectionHeading(pb, "Published poem 2 -- Limerick", pby, { color: C.ACCENT });
  pby = drawFinalCopyFrame(pb, pby, 150);

  addPdfFooter(pb, "Week 4 Session 4 | Publish Your Poems");

  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Poetry_W4S4.pptx` }),
    writePdf(pb, PUBLISH_PDF_PATH),
  ]);

  console.log("PPTX written to " + `${OUT_DIR}/Poetry_W4S4.pptx`);
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
