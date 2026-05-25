"use strict";

// Holes Unit - Lesson 10 (Week 3, Session 3): Edit & Revise
// Year 5/6 Literacy
// Focus: revise the draft against success criteria + edit for spelling/punctuation/clarity

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

const SESSION_NUMBER = 10;
const FOOTER = "Edit + Revise | Lesson 10 | Week 3 | Year 5/6 Literacy | Holes";
const OUT_DIR = "output/Holes_Lesson10_Edit_Revise";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const EDIT_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Edit and Revise Checklist",
  "Two-pass edit + revise checklist with partner-conference questions."
);
const RESOURCE_ITEMS = [EDIT_RESOURCE];
const EDIT_PDF_PATH = path.join(OUT_DIR, EDIT_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Today: EDIT and REVISE day
- Two different jobs:
  - REVISE means changing CONTENT -- words, sentences, structure
  - EDIT means fixing SURFACE -- spelling, punctuation, capitals
- We do REVISE first, EDIT last

DO:
- Display title slide
- Have student drafts on desks
- Have the Edit + Revise checklist ready

TEACHER NOTES:
The lesson protects the order: revise first (big changes), edit last (surface fixes). Students who edit first often end up polishing sentences they later cut.

WATCH FOR:
- Students who jump straight to spelling -- redirect: "Spelling is the LAST thing. Today we improve the WRITING first"

[Literacy: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- One resource today: the Edit and Revise Checklist
- Two passes: revise pass, edit pass
- Plus partner-conference questions for the middle of the lesson

DO:
- Distribute the checklist
- Have student drafts + the SC checklist from Lesson 8 on desks

[Literacy: Resources | VTLM 2.0: Student Resources]`;

const NOTES_HOOK = `SAY:
- Quick read-aloud
- Read YOUR draft to yourself QUIETLY -- just your lips moving, not out loud
- Listen for: a sentence that does not sound right; a word that does not fit; a section that is too long or too short
- Mark anything you flag with a small circle

DO:
- 3 minutes silent self-read
- Brief: "How many circles did you mark? Three? Five? Ten? Those are your revision spots"

TEACHER NOTES:
Reading aloud is the single most effective self-revision technique. Most students will hear at least 3 things they did not see when writing. Capture the count on the board to make it concrete.

WATCH FOR:
- Students who circle nothing -- prompt: "Read it again, slower. There is always one"
- Students who circle too many -- reassure: "Pick three to fix today"

[Literacy: Hook | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_LI_SC = `SAY:
- Read the LI with me
- We are learning to REVISE for content and EDIT for surface, in that order
- Three "I can" statements

DO:
- Choral read LI and SCs
- Brief: "Revise BEFORE edit. Big things BEFORE small things"

TEACHER NOTES:
SC1 -- one specific revision. Achievable. SC2 (target) -- three specific revisions + one full edit pass. SC3 -- partner conference and act on one piece of partner feedback.

WATCH FOR:
- Students who think edit and revise are the same -- redirect: "Two different jobs. Two different colours of pen"

[Literacy: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_REVISE = `SAY:
- REVISE first. What are we looking for?
- Look at the slide -- 4 questions

DO:
- Display the revise slide
- Choral read each row
- Brief example for each:
  - SECTIONS: is each part there? (setting/character/tension OR intro/body/counter/conclusion)
  - SHARP: weak verb? soft adjective? upgrade
  - SHORT: is there one sentence that drags? cut it
  - SUCCESS: tick the SC checklist -- which boxes are not ticked yet?

TEACHER NOTES:
This is the structured revision pass. Each question gives a specific improvement target. Students should fix 3 specific things during the revision pass.

WATCH FOR:
- Students who circle ten things and try to fix all -- redirect: "Three is enough today"

[Literacy: Structure | VTLM 2.0: Explicit Teaching / Strategies]`;

const NOTES_IDO = `SAY:
- Watch me REVISE this paragraph
- Soft draft on the slide. Watch what I change
- I will use BLUE pen for revisions

DO:
- Display the I Do slide
- Read the soft draft
- Do 3 changes LIVE on the board (or on screen):
  - Replace one weak verb (was -> stood)
  - Cut one drag sentence
  - Replace one soft word (very tired -> exhausted)
- Re-read the revised paragraph

TEACHER NOTES:
The model shows the SAME paragraph BEFORE and AFTER 3 revisions. Students see the move. They will do the same to their own draft.

WATCH FOR:
- Students who notice you are not REWRITING -- celebrate; the goal is targeted change

[Literacy: I Do | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_WEDO = `SAY:
- Together. We revise a different paragraph
- Look at the slide -- a soft persuasive paragraph
- Where would YOU revise? Three places

DO:
- Display the We Do slide
- Cold call for 3 revision spots
- Make the changes on the board with class input
- Re-read the revised version

TEACHER NOTES:
Different content from I Do (persuasive instead of narrative). Same revision process. Keep to 5 minutes.

WATCH FOR:
- Students suggesting specific changes (verb -> stronger verb) -- celebrate
- Students saying "make it better" -- prompt: "WHICH word? WHICH sentence?"

[Literacy: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_CFU_BUILD = `SAY:
- Quick check. Two students describe what they will do first today. Which one understands REVISE vs EDIT?

DO:
- Display both
- Show Me Fingers
- Cold call 1-2 students

TEACHER NOTES:
A jumps to surface (spelling). B revises first (content). B has the right order.

WATCH FOR:
- Students who pick B and articulate revise-first -- ready
- Students who pick A because "spelling is important" -- redirect: "Yes. But save it for the LAST 10 minutes"

[Literacy: CFU | VTLM 2.0: Formative Assessment]`;

const NOTES_CFU_REVEAL = `SAY:
- Stronger order: B
- B revises first (content). Then edits (surface)
- A polishes spelling on sentences that might get cut

DO:
- Display the reveal banner

CFU CHECKPOINT:
Technique: Show Me Fingers
Script:
- "Which order is right?"
- Scan for: most students choose B (>=80%)
PROCEED (>=80%): Release to the revision block.
PIVOT (<80%): Use the optional re-teach slide -- shows the two-pass order with colours.

[Literacy: CFU Reveal | VTLM 2.0: Formative Assessment]`;

const NOTES_RETEACH = `SAY:
- The TWO-PASS order
- PASS 1 -- revise (blue pen): content, sentences, words
- PASS 2 -- edit (red pen): spelling, punctuation, capitals
- ALWAYS revise first

DO:
- Display the re-teach slide
- Walk through both colours
- Re-check: ask students which pen they will start with today

TEACHER NOTES:
OPTIONAL. Use only if CFU was below 80%. Different approach: an explicit colour-coded two-pass procedure.

[Literacy: Optional Re-teach | VTLM 2.0: Reteach with Different Approach]`;

const NOTES_REVISE_BLOCK = `SAY:
- REVISE block. 15 minutes
- Goal: 3 specific revisions in BLUE
- Use the checklist
- Read your draft aloud quietly. Mark. Change

DO:
- Display the revise block slide
- Start a 15-minute timer
- Circulate -- focus on students who marked nothing in the launch
- Quick conferences: "Read me your sharpest sentence. Now read me one that drags"

TEACHER NOTES:
The revision block is fast and focused. Three concrete changes is the target. Students who change one big thing (cutting a paragraph) are also fine -- one big change can equal three small ones.

WATCH FOR:
- Students who rewrite the whole draft -- redirect: "Three targeted changes. Not a full rewrite"
- Students stuck -- offer the checklist questions one at a time

[Literacy: Revise Block | VTLM 2.0: Supported Application]`;

const NOTES_PARTNER = `SAY:
- Partner conference. 5 minutes
- Swap drafts with your partner
- Each partner reads the other's draft aloud (your own ears do not catch as much)
- Use the partner-conference questions on the checklist
- Give ONE specific glow + ONE specific grow

DO:
- Display the partner slide
- 5 minutes -- 2.5 each way
- Circulate -- listen in on conferences
- Capture strong glows on the board (no names)

TEACHER NOTES:
The partner conference is the second-best self-revision tool (after reading aloud). Hearing your own writing in someone else's voice is revealing.

WATCH FOR:
- Partners who only say "good" -- redirect: "Specific. Which sentence?"
- Partners who criticise -- redirect: "ONE glow first, then ONE grow"

[Literacy: Partner Conference | VTLM 2.0: Feedback Loop]`;

const NOTES_EDIT_BLOCK = `SAY:
- EDIT block. 10 minutes
- Switch to RED pen
- Goal: spelling, punctuation, capitals, paragraph breaks
- Use the edit checklist on the back of your sheet

DO:
- Display the edit block slide
- Start a 10-minute timer
- Circulate -- focus on students who skipped revision
- Final 2 minutes: ask students to read aloud one last time

TEACHER NOTES:
This is surface only. Five common fixes: missing capitals, missing full stops, run-on sentences, wrong "their/there/they're", missing apostrophes.

WATCH FOR:
- Students who find their own errors -- celebrate
- Students who add NEW errors while editing -- gently note: "Read it slowly"

[Literacy: Edit Block | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket
- Read your final draft aloud quietly one more time
- On the checklist, tick: which 3 revisions did you make? Which 3 edits did you find?
- Hand it in with your draft

DO:
- 2 minutes
- Collect drafts + checklists
- Briefly: "Tomorrow we PUBLISH -- final neat copy + display"

TEACHER NOTES:
Use the tick counts to plan tomorrow. Students with no ticks need a teacher conference at the start of Lesson 11.

[Literacy: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Self-check
- Show on fingers 1 to 5
- Partner share: name ONE revision you made today that improved your writing

DO:
- Run fingers check
- 60 seconds partner share
- Briefly: "Tomorrow -- PUBLISH. Bring your draft. Bring colour. Bring your best handwriting"

TEACHER NOTES:
Sets up Lesson 11 (publish + celebrate). The piece is now revised and edited. Tomorrow students write the final neat copy and add their visual feature.

[Literacy: Closing | VTLM 2.0: Review and Reflect]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Holes -- Lesson 10 -- Edit + Revise";

  // SLIDE 1 -- Title
  titleSlide(
    pres,
    "Edit + Revise -- The Two-Pass Polish",
    "Revise First (Content). Then Edit (Surface).",
    "Lesson 10  |  Week 3  |  Year 5/6 Literacy",
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
    "Read Your Draft Aloud and Circle the Trouble Spots",
    [
      "3 minutes -- read YOUR draft to yourself quietly (lips moving, not out loud)",
      "Listen for: a sentence that does not sound right; a word that does not fit; a section that is too long or too short",
      "Mark each trouble spot with a small circle",
    ],
    NOTES_HOOK,
    FOOTER
  );

  // SLIDE 4 -- LI / SC
  liSlide(
    pres,
    [
      "We are learning to REVISE for content and EDIT for surface, in that order",
    ],
    [
      "I can make at least one specific revision to my draft",
      "I can make three specific revisions and complete a full edit pass",
      "I can use partner feedback to make one more improvement",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // SLIDE 5 -- Revise questions
  contentSlide(
    pres,
    "Revise",
    C.SECONDARY,
    "The Four REVISE Questions",
    [
      "SECTIONS:  is every required section there?  (narrative: setting/character/tension. persuasive: intro/body/counter/conclusion.)",
      "SHARP:           which verb or adjective is soft and could be sharper?  (was -> stood; very tired -> exhausted)",
      "SHORT:           which sentence drags? cut it. which paragraph is too long? trim it",
      "SUCCESS:    open the SC checklist -- which boxes are NOT ticked yet?",
    ],
    NOTES_REVISE,
    FOOTER
  );

  // SLIDE 6 -- I Do
  modellingSlide(
    pres,
    "I Do -- Watch Me Revise",
    "Same Paragraph, 3 Targeted Changes (Blue Pen)",
    "BEFORE (soft draft):\n\n\"Stanley was very tired. He walked into Tent D. The other boys looked at him and Stanley didn't know what to say so he just looked at his shoes for a long time.\"\n\n(I will make 3 changes with blue pen.)",
    "AFTER (revised):\n\n\"Stanley dragged himself into Tent D. The other boys turned to look. He stared at the toes of his shoes.\"\n\nMy three blue changes:\n\n- weak verb: \"was very tired\" -> \"dragged himself\"\n- soft watching: \"looked at him\" -> \"turned to look\"\n- drag cut: \"for a long time\" -> removed\n\nResult: tighter, sharper, same meaning.",
    NOTES_IDO,
    FOOTER
  );

  // SLIDE 7 -- We Do
  contentSlide(
    pres,
    "We Do",
    C.SECONDARY,
    "Together: Revise This Persuasive Paragraph",
    [
      "Soft: \"The camp is bad. The boys are very sad. It is not nice. It should close.\"",
      "Where would YOU revise? Three places",
      "Cold call -- we make the changes on the board together",
    ],
    NOTES_WEDO,
    FOOTER
  );

  // SLIDE 8 + 9 -- CFU
  function buildCfuBase() {
    const slide = pres.addSlide();
    addTopBar(slide, C.ALERT);
    addBadge(slide, "CFU", { color: C.ALERT });
    addTitle(slide, "Which Student Has the Right Order?", { color: C.ALERT });

    const stampW = 1.3;
    slide.addShape("roundRect", {
      x: 9.5 - stampW, y: 0.20, w: stampW, h: 0.32, rectRadius: 0.08,
      fill: { color: C.WHITE }, line: { color: C.ALERT, width: 1.5 },
    });
    slide.addText("CHECK", {
      x: 9.5 - stampW, y: 0.20, w: stampW, h: 0.32,
      fontSize: 11, fontFace: FONT_B, color: C.ALERT, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    slide.addShape("roundRect", {
      x: 0.5, y: CONTENT_TOP, w: 3.2, h: 0.40, rectRadius: 0.08,
      fill: { color: C.ALERT },
    });
    slide.addText("Show Me Fingers: 1 (A) or 2 (B)", {
      x: 0.5, y: CONTENT_TOP, w: 3.2, h: 0.40,
      fontSize: 12, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    const cardY = CONTENT_TOP + 0.55;
    const cardH = 1.4;
    addCard(slide, 0.5, cardY, 9, cardH, { strip: C.MUTED, fill: C.WHITE });
    slide.addText("A", {
      x: 0.7, y: cardY + 0.10, w: 0.5, h: 0.4,
      fontSize: 22, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0,
    });
    slide.addText("\"I'll start by fixing all my spelling mistakes. Then I'll check capitals. Then if I have time I might change a sentence or two.\"", {
      x: 1.2, y: cardY + 0.12, w: 8.0, h: cardH - 0.24,
      fontSize: 13, fontFace: FONT_H, color: C.CHARCOAL, italic: true, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    const cardBY = cardY + cardH + 0.18;
    const cardBH = SAFE_BOTTOM - cardBY - 0.55;
    addCard(slide, 0.5, cardBY, 9, cardBH, { strip: C.SECONDARY, fill: C.BG_CARD });
    slide.addText("B", {
      x: 0.7, y: cardBY + 0.10, w: 0.5, h: 0.4,
      fontSize: 22, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0,
    });
    slide.addText("\"First I'll revise: sharper verbs, cut drag sentences, check sections. Then at the end I'll edit spelling and capitals.\"", {
      x: 1.2, y: cardBY + 0.12, w: 8.0, h: cardBH - 0.24,
      fontSize: 13, fontFace: FONT_H, color: C.CHARCOAL, italic: true, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    addFooter(slide, FOOTER);
    slide.addNotes(NOTES_CFU_BUILD);
    return slide;
  }

  withReveal(
    buildCfuBase,
    (slide) => {
      const revealY = 4.90;
      addCard(slide, 0.5, revealY, 9, 0.42, { fill: C.SUCCESS });
      slide.addText("Right order: B  --  REVISE first (content), EDIT last (surface)", {
        x: 0.5, y: revealY, w: 9, h: 0.42,
        fontSize: 14, fontFace: FONT_B, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      slide.addNotes(NOTES_CFU_REVEAL);
    }
  );

  // SLIDE 10 -- Optional Re-teach
  modellingSlide(
    pres,
    "Optional Re-teach",
    "Two Passes, Two Colours",
    "PASS 1 -- REVISE (blue pen)\n\nContent. Sentences. Words.\n- weak verb -> sharper verb\n- drag sentence -> cut\n- soft adjective -> emotive\n- missing section -> add",
    "PASS 2 -- EDIT (red pen)\n\nSurface only.\n- missing capitals\n- missing full stops\n- spelling\n- their/there/they're\n- apostrophes\n\nALWAYS revise first.\n\nWhich pen are you starting with?",
    NOTES_RETEACH,
    FOOTER
  );

  // SLIDE 11 -- Revise Block
  contentSlide(
    pres,
    "You Do",
    C.PRIMARY,
    "REVISE Block -- 15 Minutes (Blue Pen)",
    [
      "Goal: 3 specific revisions in BLUE pen",
      "Read aloud quietly. Mark. Change",
      "Use the four revise questions on the checklist",
      "Big changes are fine -- one cut paragraph counts",
    ],
    NOTES_REVISE_BLOCK,
    FOOTER
  );

  // SLIDE 12 -- Partner Conference
  contentSlide(
    pres,
    "Partner",
    C.SECONDARY,
    "Partner Conference -- 5 Minutes",
    [
      "Swap drafts",
      "Each partner READS the other's draft aloud (your own ears miss things)",
      "Use the partner-conference questions on the checklist",
      "Give ONE specific glow + ONE specific grow",
    ],
    NOTES_PARTNER,
    FOOTER
  );

  // SLIDE 13 -- Edit Block
  contentSlide(
    pres,
    "You Do",
    C.ACCENT,
    "EDIT Block -- 10 Minutes (Red Pen)",
    [
      "Switch to RED pen",
      "Spelling, punctuation, capitals, paragraph breaks",
      "Five common fixes: missing capitals, missing full stops, run-ons, their/there/they're, missing apostrophes",
      "Final 2 minutes: read aloud one last time",
    ],
    NOTES_EDIT_BLOCK,
    FOOTER
  );

  // SLIDE 14 -- Exit Ticket
  contentSlide(
    pres,
    "Exit Ticket",
    C.ACCENT,
    "Read One More Time. Tick Your Changes.",
    [
      "Read the final draft aloud quietly one more time",
      "On the checklist: tick which 3 revisions you made + which 3 edits you found",
      "Hand in the checklist with your draft",
      "2 minutes",
    ],
    NOTES_EXIT,
    FOOTER
  );

  // SLIDE 15 -- Closing
  closingSlide(
    pres,
    {
      reflectionPrompt: "Partner share: name ONE revision you made today that improved your writing.",
      scItems: [
        "I can make at least one specific revision to my draft",
        "I can make three specific revisions and complete a full edit pass",
        "I can use partner feedback to make one more improvement",
      ],
      selfAssessment: {
        prompt: "Show on your fingers: 1 (need help) to 5 (could teach it)",
        options: ["1-2", "3", "4-5"],
      },
    },
    NOTES_CLOSING
  );

  // ---- PDF: Edit and Revise Checklist ---------------------------------------
  const ed = createPdf({ title: EDIT_RESOURCE.name });
  let edY = addPdfHeader(ed, "Edit + Revise Checklist", {
    color: C.PRIMARY,
    subtitle: "Two passes -- revise first (blue), edit last (red)",
    lessonInfo: "Lesson 10 | Week 3 | Year 5/6 Literacy | Holes",
    showNameDate: true,
  });

  edY = addTipBox(ed, "Revise CONTENT first (sentences, words, structure). Edit SURFACE last (spelling, punctuation, capitals). Two colours of pen.", edY, { color: C.PRIMARY });

  edY = addSectionHeading(ed, "PASS 1 -- REVISE (Blue Pen)", edY, { color: C.PRIMARY });
  edY = addBodyText(ed, "Tick after each revision. Target: 3 revisions.", edY, { fontSize: 10, italic: true });
  edY = addBodyText(ed, "__ I found ONE weak verb and replaced it with a sharper one", edY);
  edY = addBodyText(ed, "__ I found ONE soft adjective and replaced it with an emotive one", edY);
  edY = addBodyText(ed, "__ I cut at least ONE drag sentence", edY);
  edY = addBodyText(ed, "__ Every required section is present (narrative: setting/character/tension OR persuasive: intro/body/counter/conclusion)", edY);
  edY = addBodyText(ed, "__ I checked my SC checklist and ticked 2 more boxes", edY);
  edY += 8;

  edY = addSectionHeading(ed, "PARTNER CONFERENCE -- 5 Minutes", edY, { color: C.SECONDARY });
  edY = addBodyText(ed, "Swap drafts. Read the other's draft aloud. Use these questions:", edY, { fontSize: 10, italic: true });
  edY = addBodyText(ed, "GLOW (specific): What is the SHARPEST sentence in this draft?", edY, { fontSize: 11 });
  edY = addLinedArea(ed, edY, 2, { lineSpacing: 18 });
  edY = addBodyText(ed, "GROW (specific): Where does the writing slow down or drag?", edY, { fontSize: 11 });
  edY = addLinedArea(ed, edY, 2, { lineSpacing: 18 });
  edY += 4;

  edY = addSectionHeading(ed, "PASS 2 -- EDIT (Red Pen)", edY, { color: C.ACCENT });
  edY = addBodyText(ed, "Tick after each surface check. Target: 3 fixes.", edY, { fontSize: 10, italic: true });
  edY = addBodyText(ed, "__ Capital letter at the start of every sentence", edY);
  edY = addBodyText(ed, "__ Full stop or question mark or exclamation mark at the end of every sentence", edY);
  edY = addBodyText(ed, "__ No run-on sentences -- find the longest sentence; break it up if it does not flow", edY);
  edY = addBodyText(ed, "__ Spelling checked -- circle any word you are unsure about; check it", edY);
  edY = addBodyText(ed, "__ Their / there / they're used correctly", edY);
  edY = addBodyText(ed, "__ Apostrophes used correctly (don't, didn't, Stanley's)", edY);
  edY = addBodyText(ed, "__ Quotation marks if I used dialogue", edY);
  edY = addBodyText(ed, "__ Paragraph breaks where the topic changes", edY);

  addPdfFooter(ed, "Lesson 10 | Edit + Revise Checklist -- Page 1");

  ed.addPage();
  let edY2 = addPdfHeader(ed, "Notes -- What I Changed", {
    color: C.PRIMARY,
    subtitle: "Record your top 3 revisions and top 3 edits",
    lessonInfo: "Lesson 10 | Week 3 | Year 5/6 Literacy",
    showNameDate: false,
  });

  edY2 = addSectionHeading(ed, "My three REVISIONS (content changes)", edY2, { color: C.PRIMARY });
  edY2 = addBodyText(ed, "Revision 1:", edY2, { fontSize: 11 });
  edY2 = addLinedArea(ed, edY2, 2, { lineSpacing: 22 });
  edY2 = addBodyText(ed, "Revision 2:", edY2, { fontSize: 11 });
  edY2 = addLinedArea(ed, edY2, 2, { lineSpacing: 22 });
  edY2 = addBodyText(ed, "Revision 3:", edY2, { fontSize: 11 });
  edY2 = addLinedArea(ed, edY2, 2, { lineSpacing: 22 });
  edY2 += 4;

  edY2 = addSectionHeading(ed, "My three EDITS (surface fixes)", edY2, { color: C.ACCENT });
  edY2 = addBodyText(ed, "Edit 1:", edY2, { fontSize: 11 });
  edY2 = addLinedArea(ed, edY2, 1, { lineSpacing: 22 });
  edY2 = addBodyText(ed, "Edit 2:", edY2, { fontSize: 11 });
  edY2 = addLinedArea(ed, edY2, 1, { lineSpacing: 22 });
  edY2 = addBodyText(ed, "Edit 3:", edY2, { fontSize: 11 });
  edY2 = addLinedArea(ed, edY2, 1, { lineSpacing: 22 });

  addPdfFooter(ed, "Lesson 10 | Notes -- Page 2");

  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Holes_Lesson10.pptx` }),
    writePdf(ed, EDIT_PDF_PATH),
  ]);

  console.log("PPTX written to " + `${OUT_DIR}/Holes_Lesson10.pptx`);
  console.log("Done: " + EDIT_RESOURCE.name);
}

build().catch(console.error);
