"use strict";

// Holes Unit - Lesson 11 (Week 3, Session 4 -- FINAL): Publish & Celebrate
// Year 5/6 Literacy
// Focus: write the final neat one-page publish piece + design + display + celebrate

const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");

const { createTheme, weekToVariant } = require("../themes/factory");
const T = createTheme("literacy", "grade56", weekToVariant(1));
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  addTopBar, addBadge, addTitle, addCard, addFooter,
  withReveal,
  titleSlide, liSlide, contentSlide,
  cfuSlide, closingSlide,
  modellingSlide,
} = T;

const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addLinedArea,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const SESSION_NUMBER = 11;
const FOOTER = "Publish + Celebrate | Lesson 11 | Week 3 | Year 5/6 Literacy | Holes";
const OUT_DIR = "output/Holes_Lesson11_Publish_Celebrate";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const PUBLISH_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Final Publish Page Template",
  "Single-sheet A4 publish template -- title block, writing area, visual feature area."
);
const CELEBRATE_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Unit Reflection and Celebration",
  "End-of-unit reflection card -- best line, biggest improvement, what I would write next."
);
const RESOURCE_ITEMS = [PUBLISH_RESOURCE, CELEBRATE_RESOURCE];
const PUBLISH_PDF_PATH = path.join(OUT_DIR, PUBLISH_RESOURCE.fileName);
const CELEBRATE_PDF_PATH = path.join(OUT_DIR, CELEBRATE_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Today: the FINAL lesson of our Holes unit
- You PUBLISH your piece -- write the final neat copy, add the visual feature, get it ready for display
- Then we CELEBRATE -- this is real work and you have done it

DO:
- Display title slide as students settle
- Have student edited drafts on desks
- Have publish templates ready to distribute
- Have colour materials ready (textas, coloured pencils, coloured paper for backing if available)

TEACHER NOTES:
This is the celebration lesson. Two halves: 30 minutes focused publish work, then 15-20 minutes of share + display + acknowledgement. Make the celebration genuine.

WATCH FOR:
- Students nervous about presenting -- offer the partner option, never force whole-class reading
- Students still editing during the publish block -- gently redirect: "Today is for the FINAL copy. The polish is done"

[Literacy: Title | VTLM 2.0: Establishing Purpose / Celebration]`;

const NOTES_RESOURCES = `SAY:
- Two resources today
- The ${PUBLISH_RESOURCE.name} -- your single-sheet final page
- The ${CELEBRATE_RESOURCE.name} -- your end-of-unit reflection

DO:
- Print publish template (one per student -- single-sided A4)
- Print celebrate sheet (one per student)
- Have colour materials ready
- Optional: have a few photos of last year's published work to show the standard

[Literacy: Resources | VTLM 2.0: Student Resources]`;

const NOTES_HOOK = `SAY:
- 60-second connection back
- Look at your edited draft
- Read it aloud one last time
- Whisper to yourself: "This is the piece going on the wall"

DO:
- 60 seconds silent re-read
- Briefly show where pieces will go on display (point to the wall / corridor)
- Cold call 2-3 students: "What is ONE thing you are looking forward to people seeing?"

TEACHER NOTES:
This is a quick momentum-builder. Get students reading their final draft once more before they commit to the neat copy.

WATCH FOR:
- Students who notice one more thing to fix -- allow ONE last-minute edit; then commit
- Students who are proud -- celebrate publicly

[Literacy: Hook | VTLM 2.0: Establishing Purpose]`;

const NOTES_LI_SC = `SAY:
- Read the LI with me
- We are learning to publish our piece and celebrate the unit
- Three "I can" statements

DO:
- Choral read LI and SCs
- Brief: "Three jobs today: final neat copy, visual feature, share with the class"

TEACHER NOTES:
SC1 -- write the final neat copy. Achievable. SC2 (target) -- neat copy + one visual feature (title, illustration, or quote box). SC3 -- present one paragraph aloud to a partner or table.

WATCH FOR:
- Students with poor handwriting concerned about neatness -- offer typed alternative if a device is available

[Literacy: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_FEATURES = `SAY:
- The three visual features you can add
- Pick ONE that fits your piece

DO:
- Display the visual features slide
- Read each row
- Briefly: "ONE feature, well done. Not three half-done"

TEACHER NOTES:
The three feature options keep the visual element manageable. Narrative pieces often benefit from an illustration or a decorated title. Persuasive pieces benefit from a pull-out quote box or a strong title.

WATCH FOR:
- Students wanting to add ALL features -- redirect: "One. Well done"

[Literacy: Structure | VTLM 2.0: Choice and Voice]`;

const NOTES_IDO = `SAY:
- Quick model -- watch me lay out a publish page
- Three zones on the slide
- Title at the top. Writing in the middle. Visual feature at the bottom (or side)

DO:
- Display the I Do slide
- Walk through the three zones on a sample layout
- Point out the spacing
- Brief: "Use the template. The zones are already on it"

TEACHER NOTES:
The model is brief (2-3 minutes). Most students will use the supplied template -- the zones are already set.

WATCH FOR:
- Students who notice the template structure -- celebrate
- Students wanting to redesign -- allow if they have a clear plan and time

[Literacy: I Do | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_PUBLISH_BLOCK = `SAY:
- Publish block. 30 minutes
- Write your FINAL neat copy in the writing area
- Decorate or design your title at the top
- Add ONE visual feature
- Best handwriting. Black or dark blue pen

DO:
- Distribute the publish template
- Start a 30-minute timer
- Circulate -- check handwriting and layout
- Halfway check (15 minutes): "How is your neat copy going? Anyone need more space?"

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use a printed font sample for the title (provided on the back of the template). Copy the writing zone-by-zone from your edited draft, taking your time
- Extra Notes: For students with handwriting difficulties, offer a typed version if a device is available

EXTENDING PROMPT:
- Task: Add a SECOND visual feature -- for example, a quote box that pulls out a sharp sentence from your piece for the wall display
- Extra Notes: This gets confident finishers building a richer display piece

TEACHER NOTES:
This is the longest writing block of the lesson. Students MUST end with a finished, neat, one-page publish piece. If a student is far behind, get them to copy the most important paragraph first so they finish SOMETHING.

WATCH FOR:
- Students starting over after 5 minutes -- redirect: "Keep going. Finished is better than perfect"
- Students with beautiful presentation -- celebrate publicly

[Literacy: Publish Block | VTLM 2.0: Supported Application]`;

const NOTES_GALLERY = `SAY:
- Time to share
- Two options today
- Option 1: Gallery Walk -- pieces laid out on desks, everyone walks and reads in silence, 5 minutes
- Option 2: Table Share -- each table takes turns reading one favourite paragraph aloud
- I will set up Option 1 first, then Option 2

DO:
- Display the gallery slide
- 5-minute Gallery Walk -- pieces on desks, students walk around in silence and read others' work
- Then 5-minute Table Share -- one student per table reads one favourite paragraph
- Capture reactions on the board (no names): one strong line per student

TEACHER NOTES:
The Gallery Walk is the celebration. Silent reading of classmates' work shows genuine respect. The Table Share gives students who want to read aloud the chance to do so without facing the whole class.

WATCH FOR:
- Students who hesitate to share -- give the partner-only option
- Students who hover at one piece, clearly enjoying it -- celebrate the work and the reader

[Literacy: Gallery | VTLM 2.0: Celebration]`;

const NOTES_CELEBRATE = `SAY:
- Final task -- the unit reflection
- On the celebrate card, write:
  - your BEST line from your publish piece
  - the biggest IMPROVEMENT you made this unit
  - what you would write NEXT
- 3 minutes

DO:
- 3 minutes silent
- Collect for portfolio / display

TEACHER NOTES:
The reflection card is both a self-assessment moment and an evidence-of-learning record. Keep these in student portfolios or as part of the display.

WATCH FOR:
- Students naming a specific improvement (e.g. "I cut a long sentence in the tension scene") -- evidence of metacognition
- Students who write generic answers -- prompt for specificity, but accept the response

[Literacy: Reflection | VTLM 2.0: Reflection]`;

const NOTES_CLOSING = `SAY:
- This is the LAST slide of the LAST lesson of the Holes unit
- Three weeks. Eleven sessions. One published piece on display
- You have written setting, character, tension, persuasive structure, evidence, emotive language, counter-arguments and a full publish piece
- That is REAL writing
- Partner share: what is ONE word you would use to describe the writer you have become over these three weeks?

DO:
- Run a brief fingers check on the SCs
- 60 seconds partner share
- Briefly: "Well done. Truly"
- Collect publish pieces for display
- Take photos for parents if school policy allows

TEACHER NOTES:
End on the reflection. Students choose what to claim as their own success. Collect the publish pieces for display and for portfolios. This is the last formal slide of the unit -- make the acknowledgement real.

WATCH FOR:
- Students naming a specific writer-word (sharper / braver / clearer / more confident) -- celebrate
- Students feeling uncertain -- offer a private check-in time later

[Literacy: Closing | VTLM 2.0: Review and Reflect / End of Unit]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Holes -- Lesson 11 -- Publish + Celebrate (FINAL)";

  // SLIDE 1 -- Title
  titleSlide(
    pres,
    "Publish + Celebrate",
    "The Final Session of Our Holes Unit",
    "Lesson 11  |  Week 3  |  Year 5/6 Literacy",
    NOTES_TITLE
  );

  // SLIDE 2 -- Resources
  addResourceSlide(
    pres,
    RESOURCE_ITEMS,
    T,
    FOOTER,
    NOTES_RESOURCES
  );

  // SLIDE 3 -- Hook
  contentSlide(
    pres,
    "Launch",
    C.PRIMARY,
    "Read Your Edited Draft One Last Time",
    [
      "60 seconds -- read your draft aloud quietly",
      "Whisper to yourself: \"This is the piece going on the wall\"",
      "Cold call: what is ONE thing you are looking forward to people seeing?",
    ],
    NOTES_HOOK,
    FOOTER
  );

  // SLIDE 4 -- LI / SC
  liSlide(
    pres,
    [
      "We are learning to publish our piece for classroom display and to celebrate the unit",
    ],
    [
      "I can write the final neat copy of my publish piece",
      "I can add one visual feature -- decorated title, illustration, or pull-out quote",
      "I can share one paragraph of my piece with a partner or my table",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // SLIDE 5 -- Visual Features
  contentSlide(
    pres,
    "Design",
    C.SECONDARY,
    "Pick ONE Visual Feature",
    [
      "DECORATED TITLE:    big, neat, coloured. Sets the tone of the page",
      "ILLUSTRATION:               a simple sketch related to your piece (a tens frame? a person at a door? a poster?)",
      "PULL-OUT QUOTE:        one sharp sentence from your piece, in a box, in larger font -- the line you want people to remember",
    ],
    NOTES_FEATURES,
    FOOTER
  );

  // SLIDE 6 -- I Do
  modellingSlide(
    pres,
    "I Do -- Layout",
    "Three Zones on the Publish Page",
    "TOP zone -- TITLE\n\nLarge. Neat. Decorated.\nAbout 1.5 inches tall.\n\n--------------\n\nMIDDLE zone -- WRITING\n\nYour final neat copy.\nLined area.\nBlack or dark blue pen.\n\n--------------\n\nBOTTOM (or SIDE) zone -- VISUAL FEATURE\n\nYour illustration, or your pull-out quote.\nAbout 3 inches tall.\n",
    "Quick layout tips:\n\n- LEAVE white space around your title\n- KEEP margins consistent -- the template already has them\n- BLACK or DARK BLUE pen for the writing -- no light colours\n- COLOUR is fine for the title and the visual feature\n- HANDWRITING that you can be proud of\n\nThe template has the zones marked. You just write inside them.",
    NOTES_IDO,
    FOOTER
  );

  // SLIDE 7 -- Publish Block
  contentSlide(
    pres,
    "You Do",
    C.PRIMARY,
    "Publish Block -- 30 Minutes",
    [
      "Write the FINAL neat copy in the writing zone",
      "Decorate your TITLE at the top",
      "Add ONE visual feature",
      "Best handwriting. Black or dark blue pen for the writing.",
    ],
    NOTES_PUBLISH_BLOCK,
    FOOTER
  );

  // SLIDE 8 -- Gallery Walk
  contentSlide(
    pres,
    "Gallery",
    C.SECONDARY,
    "Gallery Walk + Table Share",
    [
      "Gallery Walk (5 min): pieces on desks. Walk in silence. Read others' work",
      "Table Share (5 min): one student per table reads ONE favourite paragraph aloud",
      "Listen. Notice ONE strong line in someone else's piece",
    ],
    NOTES_GALLERY,
    FOOTER
  );

  // SLIDE 9 -- Celebrate Reflection
  contentSlide(
    pres,
    "Reflect",
    C.ACCENT,
    "Unit Reflection -- Three Questions",
    [
      "What is your BEST line from your publish piece?",
      "What is the biggest IMPROVEMENT you made in this unit?",
      "What would you write NEXT?",
      "Write your answers on the celebrate card -- 3 minutes",
    ],
    NOTES_CELEBRATE,
    FOOTER
  );

  // SLIDE 10 -- Closing (END OF UNIT)
  closingSlide(
    pres,
    {
      reflectionPrompt: "Partner share: what is ONE word you would use to describe the writer you have become over these three weeks?",
      scItems: [
        "I can write the final neat copy of my publish piece",
        "I can add one visual feature -- decorated title, illustration, or pull-out quote",
        "I can share one paragraph of my piece with a partner or my table",
      ],
      selfAssessment: {
        prompt: "Show on your fingers: 1 (need help) to 5 (could teach it)",
        options: ["1-2", "3", "4-5"],
      },
    },
    NOTES_CLOSING
  );

  // ---- PDF: Final Publish Template ------------------------------------------
  const pub = createPdf({ title: PUBLISH_RESOURCE.name });
  let pubY = addPdfHeader(pub, "Publish Page Template", {
    color: C.PRIMARY,
    subtitle: "Title -- Writing -- Visual Feature  |  Single-sided A4 for classroom display",
    lessonInfo: "Lesson 11 | Week 3 | Year 5/6 Literacy | Holes",
    showNameDate: true,
  });

  pubY = addTipBox(pub, "Use this template for your final neat copy. Title at the top. Writing in the middle. Visual feature at the bottom. Black or dark blue pen for the writing.", pubY, { color: C.PRIMARY });

  pubY = addSectionHeading(pub, "TITLE ZONE  -- write and decorate your title here", pubY, { color: C.PRIMARY });
  pubY = addLinedArea(pub, pubY, 3, { lineSpacing: 26 });
  pubY += 8;

  pubY = addSectionHeading(pub, "WRITING ZONE  -- final neat copy", pubY, { color: C.SECONDARY });
  pubY = addLinedArea(pub, pubY, 14, { lineSpacing: 22 });
  pubY += 6;

  pubY = addSectionHeading(pub, "VISUAL FEATURE ZONE  -- illustration OR pull-out quote", pubY, { color: C.ACCENT });
  pubY = addLinedArea(pub, pubY, 4, { lineSpacing: 20 });

  addPdfFooter(pub, "Lesson 11 | Final Publish Template");

  pub.addPage();
  let pubY2 = addPdfHeader(pub, "Title Font Samples (for tracing or copying)", {
    color: C.PRIMARY,
    subtitle: "Trace, copy or invent your own -- four styles to choose from",
    lessonInfo: "Lesson 11 | Week 3 | Year 5/6 Literacy",
    showNameDate: false,
  });

  pubY2 = addTipBox(pub, "If you want help with your title, use one of these styles as a guide. Trace it or copy it onto the title zone.", pubY2, { color: C.PRIMARY });

  pubY2 = addBodyText(pub, "Style 1 -- BLOCK CAPITALS: HOLES", pubY2, { fontSize: 24 });
  pubY2 += 12;
  pubY2 = addBodyText(pub, "Style 2 -- Underline: Camp Green Lake", pubY2, { fontSize: 22 });
  pubY2 += 12;
  pubY2 = addBodyText(pub, "Style 3 -- Italic: Should Camp Green Lake Close?", pubY2, { fontSize: 22, italic: true });
  pubY2 += 12;
  pubY2 = addBodyText(pub, "Style 4 -- Mixed weight: STANLEY DIGS", pubY2, { fontSize: 22 });
  pubY2 += 16;

  pubY2 = addSectionHeading(pub, "Pull-out quote tips", pubY2, { color: C.SECONDARY });
  pubY2 = addBodyText(pub, "1. Pick the SHARPEST sentence from your piece", pubY2);
  pubY2 = addBodyText(pub, "2. Write it in a box, in larger font than the main writing", pubY2);
  pubY2 = addBodyText(pub, "3. Use colour if you can", pubY2);
  pubY2 = addBodyText(pub, "4. Position it BESIDE your writing, not inside the paragraph", pubY2);

  addPdfFooter(pub, "Lesson 11 | Title + Quote Helpers -- Page 2");

  // ---- PDF: Celebration / Reflection ----------------------------------------
  const cb = createPdf({ title: CELEBRATE_RESOURCE.name });
  let cbY = addPdfHeader(cb, "End-of-Unit Reflection", {
    color: C.ACCENT,
    subtitle: "Three questions -- look back across the whole Holes unit",
    lessonInfo: "Lesson 11 | Week 3 | Year 5/6 Literacy | Holes",
    showNameDate: true,
  });

  cbY = addTipBox(cb, "Three weeks. Eleven sessions. One published piece on display. Take 3 minutes to look back.", cbY, { color: C.ACCENT });

  cbY = addSectionHeading(cb, "Question 1 -- What is your BEST line from your publish piece?", cbY, { color: C.PRIMARY });
  cbY = addBodyText(cb, "Quote it word for word.", cbY, { fontSize: 10, italic: true });
  cbY = addLinedArea(cb, cbY, 4, { lineSpacing: 22 });
  cbY += 6;

  cbY = addSectionHeading(cb, "Question 2 -- What is the biggest IMPROVEMENT you made this unit?", cbY, { color: C.SECONDARY });
  cbY = addBodyText(cb, "Be specific. What did you used to do? What do you do now?", cbY, { fontSize: 10, italic: true });
  cbY = addLinedArea(cb, cbY, 4, { lineSpacing: 22 });
  cbY += 6;

  cbY = addSectionHeading(cb, "Question 3 -- What would you write NEXT?", cbY, { color: C.ACCENT });
  cbY = addBodyText(cb, "If you had another writing unit tomorrow -- what would you most want to write?", cbY, { fontSize: 10, italic: true });
  cbY = addLinedArea(cb, cbY, 4, { lineSpacing: 22 });
  cbY += 8;

  cbY = addSectionHeading(cb, "One word for the writer you have become", cbY, { color: C.PRIMARY });
  cbY = addBodyText(cb, "Sharper? Braver? Clearer? More confident? Your word.", cbY, { fontSize: 10, italic: true });
  cbY = addLinedArea(cb, cbY, 2, { lineSpacing: 24 });

  addPdfFooter(cb, "Lesson 11 | Unit Reflection -- KEEP IN PORTFOLIO");

  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Holes_Lesson11.pptx` }),
    writePdf(pub, PUBLISH_PDF_PATH),
    writePdf(cb, CELEBRATE_PDF_PATH),
  ]);

  console.log("PPTX written to " + `${OUT_DIR}/Holes_Lesson11.pptx`);
  console.log("Done: " + PUBLISH_RESOURCE.name);
  console.log("Done: " + CELEBRATE_RESOURCE.name);
}

build().catch(console.error);
