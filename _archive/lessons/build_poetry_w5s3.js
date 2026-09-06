"use strict";

// Poetry Unit (Year 6) - Week 5 Session 3: Catch-Up Publishing Workshop
// A flexible workshop: students audit their poetry folder and finish/publish any
// unfinished poems from the whole unit, using the agreed editing checklist.
// This is the unit's dedicated catch-up session -- no new form, no dependency on
// having attended every prior session.
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
const FOOTER = "Poetry | Week 5 Session 3 | Year 6 Literacy";
const OUT_DIR = "output/Poetry_W5S3_CatchUp_Workshop";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const FOLDER_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Poetry Folder Check",
  "A tracker of the unit's poems, the editing checklist, and a final-copy frame to finish one off."
);
const RESOURCE_ITEMS = [FOLDER_RESOURCE];
const FOLDER_PDF_PATH = path.join(OUT_DIR, FOLDER_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Today is a workshop. No new poem shape -- this is your time to finish and publish.
- Over the unit we have tried nine poem shapes. Today you catch up on any you did not finish.
- Whether you have missed a session or just want to polish, this is your chance.

DO:
- Display the title slide.
- Have every student's poetry folder out, plus the editing checklist.

TEACHER NOTES:
This is the dedicated catch-up session. It removes the pressure of missed sessions: students simply finish what they can, to a high standard.

WATCH FOR:
- Students unsure what to work on -- the folder check helps them choose.

[Literacy: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- One sheet today: the Poetry Folder Check. It lists the unit's poems so you can see what is done.
- It also has the editing checklist and a frame to publish one more poem.

DO:
- Print the Poetry Folder Check, one per student.
- Have all drafts and published poems from the unit available.
- Have the Poets Toolbox nearby.

[Literacy: Resources | VTLM 2.0: Enabling Learning]`;

const NOTES_LAUNCH = `SAY:
- Let us take stock. Open your poetry folder and look at what you have.
- Which poems are published and proud? Which are still drafts? Which did you not get to?
- Tick them off on your folder check. We are not behind -- we are getting ready to make a booklet.

DO:
- Display the folder audit prompt.
- Give two minutes for students to sort their folder and tick the check sheet.
- Reassure anyone with gaps: today is exactly for that.

TEACHER NOTES:
The self-audit activates a clear goal. It is the heart of the catch-up design: students see what they have and choose a priority, rather than feeling behind.

WATCH FOR:
- Students who feel they have "nothing" -- help them find one draft to finish; one published poem is a success.

[Literacy: Launch | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me. Today we finish and publish poems from our unit.
- The first "I can" is for everyone -- choose a poem to work on.

DO:
- Choral read the LI and success criteria.
- Stress: the goal is at least one MORE published poem.

TEACHER NOTES:
SC1 (all): choose a poem to finish. SC2 (target): publish at least one more poem using the checklist. SC3 (depth): give a partner kind, useful feedback. Exit ticket targets SC2.

WATCH FOR:
- Students who aim too big -- one finished, polished poem beats three half-done ones.

[Literacy: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_MENU = `SAY:
- Here are the nine shapes we have tried this unit. Use this as your menu.
- Maybe you never finished your haiku, or your colour poem needs a stronger image.
- Pick ONE to be your priority today. Then use the checklist to publish it.

DO:
- Display the poem menu and the checklist together.
- Help students choose a realistic priority.
- Remind them the checklist is the same one we always use.

TEACHER NOTES:
The menu makes the whole unit visible so students can choose. Keep the choice small -- one priority poem.

WATCH FOR:
- Students drawn to start something brand new -- redirect to finishing an existing draft first.

[Literacy: Content / Menu | VTLM 2.0: Explicit Teaching / Success Criteria]`;

const NOTES_IDO = `SAY:
- Watch me finish a poem quickly and well. Here is a half-done haiku draft.
- First, the checklist. Form: it is short of beats on line two. I add a word to reach seven.
- Image: it has no picture yet, so I make "the flower" into "the bright red poppy".
- Read aloud: it works. In two minutes I turned a draft into a published poem.

DO:
- Display the before and after.
- Run the checklist out loud as you fix the poem.
- Show how quickly a draft becomes publishable with focused edits.

TEACHER NOTES:
Model efficient, focused finishing. The point is that publishing one poem well is achievable in the session.

WATCH FOR:
- Students who think finishing means rewriting -- show that small checklist fixes are enough.

[Literacy: I Do | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_YOUDO = `SAY:
- Your turn. This is your workshop time. Pick your priority poem and publish it.
- Run the checklist. Make the fixes. Write your neat final copy in the frame.
- When you finish one, start the next. Quality over quantity.
- You have plenty of time. I will come around and help.

DO:
- Distribute the Poetry Folder Check (with its final-copy frame).
- Circulate for the whole block. Prioritise students with the most catching up to do.
- Quick conferences using the checklist: "Which point will make this one stronger?"

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Choose ONE short poem (a three-line poem or a haiku). Fix one thing with the checklist, then write it neatly.
- Extra Notes: One published poem is a real success and meets SC2.
EXTENDING PROMPT:
- Task: Publish two or three poems to a high standard, then plan the order for your booklet next session.
- Extra Notes: Gets confident students ready for the booklet build.

TEACHER NOTES:
This long work block is the catch-up. Every student should leave with at least one more published poem. Keep published poems safe for the booklet.

WATCH FOR:
- Students who drift -- give them a clear next step: "Finish this one, then show me."

[Literacy: You Do | VTLM 2.0: Supported Application]`;

const NOTES_PEER = `SAY:
- Quick partner check. Swap one published poem with a partner.
- Read your partner's poem aloud to them, gently. Then say ONE kind thing and ONE helpful idea.
- Use the checklist words: form, image, words, sound, neat.

DO:
- Pair students up.
- Model the sentence frames: "I really liked..." and "One idea: maybe try...".
- Keep feedback kind and specific.

CFU CHECKPOINT:
Technique: Partner read-aloud and feedback
Script:
- "Give one kind comment and one helpful idea using the checklist."
- Scan for: feedback that is specific and kind, not just "it's good".
PROCEED (>=80%): Students apply one piece of feedback, then continue publishing.
PIVOT (<80%): Re-model a specific kind comment and a checklist-based suggestion, then have pairs try again.

TEACHER NOTES:
Peer feedback gives every poem a reader and builds the depth criterion. Keep it warm and useful.

WATCH FOR:
- Vague feedback ("it's nice") -- prompt for a checklist point.

[Literacy: We Do / Peer Feedback | VTLM 2.0: Formative Assessment]`;

const NOTES_EXIT = `SAY:
- Exit ticket. Write down which poems you have now published.
- Then write your goal for next session: what will your booklet include?
- Two minutes.

DO:
- Two minutes, silent.
- Collect, and note who still needs time so you can support them in the booklet session.

TEACHER NOTES:
Exit ticket targets SC2 and sets up the booklet. Use it to see who is ready and who needs more support.

WATCH FOR:
- Students with no published poems yet -- plan to give them early support next session.

[Literacy: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Self-check. Show me on your fingers, one to five, for each "I can".
- Then tell a partner: which poem are you most proud of so far, and why?

DO:
- Run the fingers check for each success criterion.
- Sixty seconds partner share.
- Preview: "Next session we turn your published poems into a real booklet."

TEACHER NOTES:
Celebrate the catching up. Everyone should leave with a clearer, fuller folder ready for the booklet.

[Literacy: Closing | VTLM 2.0: Review and Reflect]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Poetry Week 5 Session 3 -- Catch-Up Publishing Workshop";

  // SLIDE 1 -- Title
  titleSlide(
    pres,
    "Poetry Workshop",
    "Finish and publish your poems -- your catch-up time",
    "Week 5 Session 3  |  Year 6 Literacy  |  Poetry",
    NOTES_TITLE
  );

  // SLIDE 2 -- Resources
  addResourceSlide(pres, RESOURCE_ITEMS, T, FOOTER, NOTES_RESOURCES);

  // SLIDE 3 -- Launch (folder self-audit)
  modellingSlide(
    pres,
    "Take Stock",
    "Open Your Poetry Folder",
    "Look at what you have:\n\n- Which poems are\n   PUBLISHED and proud?\n\n- Which are still DRAFTS?\n\n- Which did you not\n   get to yet?\n\nTick them on your sheet.",
    "We are not behind.\n\nWe are getting ready\nto make a booklet.\n\nToday you choose ONE\npriority poem to publish.",
    NOTES_LAUNCH,
    FOOTER
  );

  // SLIDE 4 -- LI / SC
  liSlide(
    pres,
    [
      "We are finishing and publishing poems from our unit, using our editing checklist",
    ],
    [
      "I can choose a poem to finish",
      "I can publish at least one more poem using the checklist",
      "I can give a partner kind, useful feedback",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // SLIDE 5 -- Content: the poem menu + checklist
  contentSlide(
    pres,
    "Your Menu",
    C.SECONDARY,
    "Nine Poems This Unit -- Pick One to Finish",
    [
      "Week 1: three-line poem, haiku",
      "Week 2: cinquain, colour poem",
      "Week 3: rhyming poem, tanka",
      "Week 4: clerihew, limerick",
      "Week 5: verse",
      "Checklist: form, image, words, sound, neat",
    ],
    NOTES_MENU,
    FOOTER
  );

  // SLIDE 6 -- I Do (finish a poem quickly with the checklist)
  modellingSlide(
    pres,
    "I Do",
    "Watch Me Finish One Quickly",
    "A half-done haiku draft.\nRun the checklist:\n\n- Form: line 2 is short\n   -> add a word to reach 7\n\n- Image: \"the flower\"\n   -> \"the bright red poppy\"\n\nRead aloud -- published!",
    "Before:\nThe garden is calm\nA flower opens slowly\nMorning light arrives\n\nAfter:\nThe garden is calm\nThe bright red poppy opens\nMorning light arrives",
    NOTES_IDO,
    FOOTER
  );

  // SLIDE 7 -- You Do (main work block)
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "You Do", { color: C.PRIMARY, w: 1.5 });
    addTitle(s, "Workshop Time -- Publish Your Poems");

    addCard(s, 0.5, CONTENT_TOP, 9, 1.7, { strip: C.PRIMARY, fill: C.WHITE });
    s.addText("Your priority poem:", {
      x: 0.75, y: CONTENT_TOP + 0.12, w: 8.4, h: 0.32,
      fontSize: 16, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("First:   Pick ONE poem to finish.\nNext:   Run the checklist and make the fixes.\nThen:   Write your neat final copy. Finished one? Start the next.", {
      x: 0.75, y: CONTENT_TOP + 0.52, w: 8.4, h: 1.05,
      fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    const tipY = CONTENT_TOP + 1.85;
    const tipH = SAFE_BOTTOM - tipY;
    addCard(s, 0.5, tipY, 9, tipH, { strip: C.ACCENT, fill: C.BG_CARD });
    s.addText("Quality over quantity", {
      x: 0.75, y: tipY + 0.10, w: 8.4, h: 0.30,
      fontSize: 14, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("- Catching up? Finish ONE short poem well -- a three-line poem or haiku.\n- Already done? Polish two or three, then plan the order for your booklet.", {
      x: 0.75, y: tipY + 0.44, w: 8.4, h: tipH - 0.55,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  }

  // SLIDE 8 -- Peer feedback (We Do / CFU)
  cfuSlide(
    pres,
    "Partner Check",
    "Read It Aloud -- Give Kind, Useful Feedback",
    "Swap one poem. Read it aloud to your partner.",
    "Say ONE kind thing and ONE helpful idea.\n\n\"I really liked...\"\n\"One idea: maybe try...\"\n\nUse the checklist words: form, image, words, sound, neat.",
    NOTES_PEER,
    FOOTER
  );

  // SLIDE 9 -- Exit Ticket
  exitTicketSlide(
    pres,
    [
      "Which poems have you now published? List them.",
      "Your goal for next session: what will your booklet include?",
    ],
    NOTES_EXIT,
    FOOTER,
    { assessesSc: 2, title: "Where Are You Up To?" }
  );

  // SLIDE 10 -- Closing
  closingSlide(
    pres,
    {
      reflectionPrompt: "Which poem are you most proud of so far, and why?",
      scItems: [
        "I can choose a poem to finish",
        "I can publish at least one more poem using the checklist",
        "I can give a partner kind, useful feedback",
      ],
      selfAssessment: {
        prompt: "Show on your fingers: 1 (need help) to 5 (could teach it)",
        options: ["1-2", "3", "4-5"],
      },
    },
    NOTES_CLOSING
  );

  // ---- PDF: Poetry Folder Check ---------------------------------------------
  const fc = createPdf({ title: FOLDER_RESOURCE.name });
  let fy = addPdfHeader(fc, "Poetry Folder Check", {
    color: C.PRIMARY,
    subtitle: "See what you have. Choose one to finish. Publish it.",
    lessonInfo: "Week 5 Session 3 | Year 6 Literacy | Poetry",
    showNameDate: true,
  });

  fy = addTipBox(fc, "Tick each poem: done, draft, or not yet. Then circle the ONE you will publish today.", fy, { color: C.PRIMARY });

  fy = addSectionHeading(fc, "My poems this unit", fy, { color: C.PRIMARY });
  fy = addBodyText(fc, "[  ] Three-line poem      [  ] Haiku           [  ] Cinquain", fy, { fontSize: 11 });
  fy = addBodyText(fc, "[  ] Colour poem          [  ] Rhyming poem     [  ] Tanka", fy, { fontSize: 11 });
  fy = addBodyText(fc, "[  ] Clerihew             [  ] Limerick         [  ] Verse", fy, { fontSize: 11 });
  fy += 8;

  fy = addSectionHeading(fc, "Editing checklist for the poem I am finishing", fy, { color: C.SECONDARY });
  fy = addBodyText(fc, "[  ]  Form: it follows the shape", fy, { fontSize: 11 });
  fy = addBodyText(fc, "[  ]  Image: at least one simile, metaphor or personification", fy, { fontSize: 11 });
  fy = addBodyText(fc, "[  ]  Words: the best, most precise words chosen", fy, { fontSize: 11 });
  fy = addBodyText(fc, "[  ]  Sound: I read it aloud and it sounds right", fy, { fontSize: 11 });
  fy = addBodyText(fc, "[  ]  Neat: spelling, capitals and layout tidy", fy, { fontSize: 11 });

  addPdfFooter(fc, "Week 5 Session 3 | Poetry Folder Check -- Page 1");

  fc.addPage();
  let fy2 = addPdfHeader(fc, "Publish One More -- Final Copy", {
    color: C.PRIMARY,
    subtitle: "Write your neat, published poem in the frame.",
    lessonInfo: "Week 5 Session 3 | Year 6 Literacy",
    showNameDate: false,
  });
  fy2 = addTipBox(fc, "Improve with the checklist first, then write your neat final copy. Add a small illustration.", fy2, { color: C.SECONDARY });
  fy2 = drawFinalCopyFrame(fc, fy2, 320);

  addPdfFooter(fc, "Week 5 Session 3 | Final Copy");

  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Poetry_W5S3.pptx` }),
    writePdf(fc, FOLDER_PDF_PATH),
  ]);

  console.log("PPTX written to " + `${OUT_DIR}/Poetry_W5S3.pptx`);
  console.log("Done: " + FOLDER_RESOURCE.name);
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
