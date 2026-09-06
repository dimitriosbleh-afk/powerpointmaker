"use strict";

// Poetry Unit (Year 6) - Week 5 Session 4: Create the Poetry Booklet
// The culminating session: students assemble their published poems into a personal
// poetry booklet (cover, contents, poems, about-the-poet) and celebrate the unit.
// Booklet Kit PDF provides the cover, contents and about-the-poet pages.
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
const FOOTER = "Poetry | Week 5 Session 4 | Year 6 Literacy";
const OUT_DIR = "output/Poetry_W5S4_Booklet";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const KIT_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Poetry Booklet Kit",
  "Cover page, contents page and an about-the-poet page to assemble a personal poetry booklet."
);
const RESOURCE_ITEMS = [KIT_RESOURCE];
const KIT_PDF_PATH = path.join(OUT_DIR, KIT_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- This is it -- the last session of our poetry unit. Today we make a real booklet of your poems.
- You have written and published poems in nine different shapes. Now we gather them into a book.
- By the end of today, you will hold your own poetry book that you made.

DO:
- Display the title slide.
- Have all published poems and the Booklet Kit ready.

TEACHER NOTES:
A celebration and assembly session. The thinking work is selecting, ordering and presenting; the joy is holding a finished book.

WATCH FOR:
- Students with fewer poems -- a small booklet is still a real achievement; help them feel proud.

[Literacy: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- One kit today: the Poetry Booklet Kit. It has a cover page, a contents page and an about-the-poet page.
- Your published poems become the inside pages.

DO:
- Print the Booklet Kit, one per student.
- Have every student's published poems gathered and ready to bind or staple.
- Have coloured pencils for cover art.

[Literacy: Resources | VTLM 2.0: Enabling Learning]`;

const NOTES_LAUNCH = `SAY:
- What makes a real poetry book? Let us look at the parts.
- A cover with a title and the poet's name. A contents page that lists the poems.
- The poems themselves, in an order the poet chooses. And often a page about the poet.
- Today you build all of those parts for YOUR book.

DO:
- Display the parts of a poetry book.
- Pass around a real poetry book if you have one, pointing to each part.
- Connect each part to the kit pages students have.

TEACHER NOTES:
The launch makes "a real book" concrete by naming its parts, which become the build plan.

WATCH FOR:
- Students who want to dive into cover art first -- fine, but make sure the poems get ordered too.

[Literacy: Launch | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me. Today we assemble our poems into a booklet we are proud to share.
- The first "I can" is for everyone -- put your poems in order with a cover.

DO:
- Choral read the LI and success criteria.
- Stress: a finished booklet, however many poems, is the goal.

TEACHER NOTES:
SC1 (all): order poems with a cover. SC2 (target): present the booklet neatly with a contents page. SC3 (depth): write an about-the-poet page and choose a proudest poem with a reason. Exit ticket targets SC3 (reflection).

WATCH FOR:
- Students who fuss over perfection -- a neat, complete booklet beats an unfinished perfect one.

[Literacy: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_PARTS = `SAY:
- Here is your booklet build plan -- the parts to put together.
- Cover: a title you choose, your name, and some cover art.
- Contents: a list of your poems, in order.
- The poems: your published copies, in the order you like best.
- About the poet: a short page about you as a writer.

DO:
- Display the build plan.
- Suggest an order: maybe favourite first, or by mood, or by shape.
- Remind students to use their best published copies.

TEACHER NOTES:
The plan keeps the assembly organised. Let students choose the order -- that is an authorial decision.

WATCH FOR:
- Students unsure how to order -- offer a simple rule: start with your proudest poem.

[Literacy: Content / Plan | VTLM 2.0: Explicit Teaching / Success Criteria]`;

const NOTES_IDO = `SAY:
- Watch me put a booklet together. First the cover -- a title, my name, and a simple drawing.
- Then I choose an order. I will put my favourite poem first so it grabs the reader.
- I fill in the contents page to match that order.
- Last, the about-the-poet page: my name, the poem I most enjoyed writing, and one thing I learned.

DO:
- Display the model booklet layout.
- Show the cover, the ordered poems, the contents and the about-the-poet page.
- Make the ordering decision out loud.

TEACHER NOTES:
Model the assembly steps in order so students have a clear path. Keep it achievable.

WATCH FOR:
- Students who skip the contents -- show how it helps a reader find each poem.

[Literacy: I Do | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_YOUDO = `SAY:
- Your turn. Build your poetry booklet.
- Design your cover. Choose the order of your poems. Fill in the contents. Write your about-the-poet page.
- Use your best published copies. This is your book -- make it one you are proud of.
- You have the whole block. I will come around.

DO:
- Distribute the Booklet Kit.
- Circulate. Help students order their poems and complete each part.
- Support students with fewer poems to make a small, proud booklet.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the kit pages as they are. Put your poems in any order, add a cover title and your name.
- Extra Notes: A complete, ordered booklet with a cover meets SC1 and SC2.
EXTENDING PROMPT:
- Task: Add an illustrated title page, write a short welcome to your reader, or include a brand-new poem.
- Extra Notes: Pushes confident students to present like a published poet.

TEACHER NOTES:
Every student should leave with a finished booklet, however many poems it holds. Celebrate the achievement.

WATCH FOR:
- Students who run out of time on cover art -- make sure poems and contents are done first.

[Literacy: You Do | VTLM 2.0: Supported Application]`;

const NOTES_GALLERY = `SAY:
- Time to celebrate. We are going to share our books in a poetry gallery.
- Place your open booklet on your desk. Then walk quietly and read others' poems.
- Leave a kind sticky-note comment on one book you enjoyed.
- When you read your own favourite aloud, use your proudest voice.

DO:
- Set up the gallery -- books open on desks.
- Give students sticky notes for kind comments.
- Invite a few volunteers to read a poem aloud to the class.

TEACHER NOTES:
A gallery walk gives every booklet a reader and ends the unit on a celebratory note. Keep comments kind and specific.

WATCH FOR:
- Students reluctant to share -- let them choose just one poem, or share with a partner first.

[Literacy: We Do / Celebration | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Last reflection of the unit. Which poem in your booklet are you most proud of, and why?
- Then write one thing you can do now as a poet that you could not do five weeks ago.
- Two minutes.

DO:
- Two minutes, silent.
- Collect, or have students read their reflection to a partner.

TEACHER NOTES:
Exit ticket targets SC3 and reflects on the whole unit. It captures growth in the poet's own words.

WATCH FOR:
- Students who cannot name a proud poem -- help them see the progress in their folder.

[Literacy: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Self-check. Show me on your fingers, one to five, for each "I can".
- Look back at the whole unit. You have painted pictures with words and played with sound.
- You are poets now -- you have your own book to prove it. Well done.

DO:
- Run the fingers check for each success criterion.
- Briefly revisit the two big skills: imagery (pictures) and sound (rhyme, rhythm, beat).
- Celebrate the finished booklets.

TEACHER NOTES:
Close the unit by naming the two transferable skills students have built -- imagery and sound -- which carry into all their writing.

[Literacy: Closing | VTLM 2.0: Review and Reflect]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Poetry Week 5 Session 4 -- Create the Poetry Booklet";

  // SLIDE 1 -- Title
  titleSlide(
    pres,
    "Make Your Poetry Booklet",
    "Gather your poems into a book you are proud to share",
    "Week 5 Session 4  |  Year 6 Literacy  |  Poetry",
    NOTES_TITLE
  );

  // SLIDE 2 -- Resources
  addResourceSlide(pres, RESOURCE_ITEMS, T, FOOTER, NOTES_RESOURCES);

  // SLIDE 3 -- Launch (parts of a poetry book)
  modellingSlide(
    pres,
    "Launch",
    "What Makes a Real Poetry Book?",
    "Every poetry book has parts:\n\n- A COVER with a title\n   and the poet's name\n- A CONTENTS page\n- The POEMS, in an order\n   the poet chooses\n- An ABOUT THE POET page\n\nToday you build them all.",
    "Cover\n   |\nContents\n   |\nYour poems (in order)\n   |\nAbout the poet",
    NOTES_LAUNCH,
    FOOTER
  );

  // SLIDE 4 -- LI / SC
  liSlide(
    pres,
    [
      "We are assembling our published poems into a poetry booklet we are proud to share",
    ],
    [
      "I can put my poems in order with a cover",
      "I can present my booklet neatly with a contents page",
      "I can write an about-the-poet page and choose my proudest poem",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // SLIDE 5 -- Content: the build plan
  contentSlide(
    pres,
    "Build Plan",
    C.SECONDARY,
    "The Parts of Your Booklet",
    [
      "Cover: your title, your name, cover art",
      "Contents: a list of your poems, in order",
      "Poems: your best published copies, in the order you choose",
      "About the poet: a short page about you as a writer",
    ],
    NOTES_PARTS,
    FOOTER
  );

  // SLIDE 6 -- I Do (assemble a booklet)
  modellingSlide(
    pres,
    "I Do",
    "Watch Me Put a Booklet Together",
    "My steps:\n\n1. COVER -- title, name,\n   a simple drawing.\n\n2. ORDER -- favourite poem\n   first to grab the reader.\n\n3. CONTENTS -- list them\n   in that order.\n\n4. ABOUT THE POET page.",
    "Cover: \"Words I Made\"\n   by Sam\n\nContents:\n1. My Haiku\n2. The Storm (verse)\n3. My Limerick\n\nAbout the poet:\nI loved writing the verse...",
    NOTES_IDO,
    FOOTER
  );

  // SLIDE 7 -- You Do (assemble)
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "You Do", { color: C.PRIMARY, w: 1.5 });
    addTitle(s, "Build Your Poetry Booklet");

    addCard(s, 0.5, CONTENT_TOP, 9, 1.7, { strip: C.PRIMARY, fill: C.WHITE });
    s.addText("Use your Booklet Kit and your published poems:", {
      x: 0.75, y: CONTENT_TOP + 0.12, w: 8.4, h: 0.32,
      fontSize: 15, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("First:   Design your cover (title, name, art).\nNext:   Choose the order of your poems, then fill in the contents.\nThen:   Write your about-the-poet page.", {
      x: 0.75, y: CONTENT_TOP + 0.52, w: 8.4, h: 1.05,
      fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    const tipY = CONTENT_TOP + 1.85;
    const tipH = SAFE_BOTTOM - tipY;
    addCard(s, 0.5, tipY, 9, tipH, { strip: C.ACCENT, fill: C.BG_CARD });
    s.addText("This is YOUR book -- make it one you are proud of", {
      x: 0.75, y: tipY + 0.10, w: 8.4, h: 0.30,
      fontSize: 14, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("- Fewer poems? A small booklet is still a real book. Make it neat and proud.\n- Ready for more? Add an illustrated title page, a welcome to your reader, or a new poem.", {
      x: 0.75, y: tipY + 0.44, w: 8.4, h: tipH - 0.55,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  }

  // SLIDE 8 -- Celebration (gallery walk)
  contentSlide(
    pres,
    "Celebrate",
    C.SECONDARY,
    "Poetry Gallery -- Share Your Books",
    [
      "Place your open booklet on your desk",
      "Walk quietly and read others' poems",
      "Leave one kind sticky-note comment",
      "Read your favourite poem aloud in your proudest voice",
    ],
    NOTES_GALLERY,
    FOOTER
  );

  // SLIDE 9 -- Exit Ticket / reflection
  exitTicketSlide(
    pres,
    [
      "Which poem in your booklet are you most proud of, and why?",
      "Write one thing you can do now as a poet that you could not do five weeks ago.",
    ],
    NOTES_EXIT,
    FOOTER,
    { assessesSc: 3, title: "Proud to Be a Poet" }
  );

  // SLIDE 10 -- Closing
  closingSlide(
    pres,
    {
      reflectionPrompt: "You are poets now -- you have your own book to prove it. What will you write next?",
      scItems: [
        "I can put my poems in order with a cover",
        "I can present my booklet neatly with a contents page",
        "I can write an about-the-poet page and choose my proudest poem",
      ],
      selfAssessment: {
        prompt: "Show on your fingers: 1 (need help) to 5 (could teach it)",
        options: ["1-2", "3", "4-5"],
      },
    },
    NOTES_CLOSING
  );

  // ---- PDF: Poetry Booklet Kit ----------------------------------------------
  const kit = createPdf({ title: KIT_RESOURCE.name });

  // Page 1 -- Cover
  let ky = addPdfHeader(kit, "My Poetry Booklet", {
    color: C.PRIMARY,
    subtitle: "Front cover -- design it to make a reader want to open your book.",
    lessonInfo: "Week 5 Session 4 | Year 6 Literacy | Poetry",
    showNameDate: false,
  });
  ky = addBodyText(kit, "Title of my booklet:", ky, { fontSize: 12 });
  ky = addLinedArea(kit, ky, 1, { lineSpacing: 26 });
  ky = addBodyText(kit, "By (poet's name):", ky, { fontSize: 12 });
  ky = addLinedArea(kit, ky, 1, { lineSpacing: 26 });
  ky += 6;
  ky = addBodyText(kit, "Cover art -- draw something that fits your poems:", ky, { fontSize: 10, italic: true });
  ky = drawFrameBox(kit, ky, 300, "Draw your cover art here.");
  addPdfFooter(kit, "Week 5 Session 4 | Booklet Kit -- Cover");

  // Page 2 -- Contents
  kit.addPage();
  let ky2 = addPdfHeader(kit, "Contents", {
    color: C.SECONDARY,
    subtitle: "List your poems in the order they appear.",
    lessonInfo: "Week 5 Session 4 | Year 6 Literacy",
    showNameDate: false,
  });
  ky2 = addTipBox(kit, "Write the title of each poem in order. Start with your proudest one if you like.", ky2, { color: C.SECONDARY });
  for (let i = 1; i <= 9; i += 1) {
    ky2 = addBodyText(kit, `${i}.`, ky2, { fontSize: 12 });
    ky2 = addLinedArea(kit, ky2, 1, { lineSpacing: 20 });
  }
  addPdfFooter(kit, "Week 5 Session 4 | Booklet Kit -- Contents");

  // Page 3 -- About the Poet
  kit.addPage();
  let ky3 = addPdfHeader(kit, "About the Poet", {
    color: C.PRIMARY,
    subtitle: "A short page about you as a writer.",
    lessonInfo: "Week 5 Session 4 | Year 6 Literacy",
    showNameDate: false,
  });
  ky3 = addBodyText(kit, "My name:", ky3, { fontSize: 12 });
  ky3 = addLinedArea(kit, ky3, 1, { lineSpacing: 24 });
  ky3 = addBodyText(kit, "The poem I most enjoyed writing was... because...", ky3, { fontSize: 12 });
  ky3 = addLinedArea(kit, ky3, 2, { lineSpacing: 24 });
  ky3 = addBodyText(kit, "One thing I learned about writing poems:", ky3, { fontSize: 12 });
  ky3 = addLinedArea(kit, ky3, 2, { lineSpacing: 24 });
  ky3 += 6;
  ky3 = addBodyText(kit, "A small self-portrait:", ky3, { fontSize: 10, italic: true });
  ky3 = drawFrameBox(kit, ky3, 150, "Draw yourself here.");
  addPdfFooter(kit, "Week 5 Session 4 | Booklet Kit -- About the Poet");

  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Poetry_W5S4.pptx` }),
    writePdf(kit, KIT_PDF_PATH),
  ]);

  console.log("PPTX written to " + `${OUT_DIR}/Poetry_W5S4.pptx`);
  console.log("Done: " + KIT_RESOURCE.name);
}

// A bordered drawing/writing frame with a small grey caption.
function drawFrameBox(doc, y, height, caption) {
  const left = doc.page.margins.left;
  const width = doc.page.width - doc.page.margins.left - doc.page.margins.right;
  doc.save();
  doc.lineWidth(1).strokeColor("#888888")
    .roundedRect(left, y, width, height, 8).stroke();
  doc.restore();
  if (caption) {
    doc.font("Sans").fontSize(8).fillColor("#777777")
      .text(String(caption), left + 10, y + 8, { width: width - 20 });
    doc.fillColor("#000000");
  }
  return y + height + 6;
}

build().catch((err) => {
  console.error(err);
  process.exit(1);
});
