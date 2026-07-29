"use strict";

// Holes Unit - Lesson 8 (Week 3, Session 1): Choose & Plan the Publish Piece
// Year 5/6 Literacy
// Focus: choose narrative or persuasive, plan the one-page piece for classroom display

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
  modellingSlide, vocabSlide,
} = T;

const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addLinedArea,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const SESSION_NUMBER = 8;
const FOOTER = "Plan Publish Piece | Lesson 8 | Week 3 | Year 5/6 Literacy | Holes";
const OUT_DIR = "output/Holes_Lesson8_Plan_Publish";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const PLAN_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Publish Piece Plan",
  "Student template: choose narrative or persuasive and plan a one-page publish piece for classroom display."
);
const SC_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Publish Success Criteria Checklist",
  "Single-page success criteria checklist for the narrative and persuasive publish options."
);
const RESOURCE_ITEMS = [PLAN_RESOURCE, SC_RESOURCE];
const PLAN_PDF_PATH = path.join(OUT_DIR, PLAN_RESOURCE.fileName);
const SC_PDF_PATH = path.join(OUT_DIR, SC_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Welcome to Week 3 -- PUBLISH week
- Across this week you will plan, draft, edit and publish ONE finished piece
- You CHOOSE: narrative or persuasive
- The finished piece goes on display in our classroom

DO:
- Display title slide
- Briefly show where finished pieces will go on display (a wall, a board, a corridor)
- Have ALL student work from Weeks 1 and 2 back on desks

TEACHER NOTES:
This is the start of the publish arc. Today: choose + plan. Lessons 9: draft. Lesson 10: edit + revise. Lesson 11: publish + celebrate.

Students must complete a SHORT one-page piece -- not an essay. The aim is a finished, display-worthy piece they can be proud of.

WATCH FOR:
- Students paralysed by the choice -- give them a 30-second decision rule (whichever piece you are PROUDEST of right now)
- Students who want to start something brand-new -- redirect: "Use one of your Week 1 or Week 2 pieces. We are FINISHING, not starting"

[Literacy: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- Two resources today
- The ${PLAN_RESOURCE.name} -- the publish plan
- The ${SC_RESOURCE.name} -- the success criteria checklist for both options

DO:
- Print plan (one per student)
- Print SC checklist (one per student)
- Have all student work from Weeks 1 + 2 ready

[Literacy: Resources | VTLM 2.0: Student Resources]`;

const NOTES_HOOK = `SAY:
- Two-minute decision warm-up
- Look at ALL your work from the last two weeks
- Ask yourself: which piece am I MOST PROUD of?
- That is your publish piece. We will polish it this week

DO:
- 2 minutes silent reading -- students read their own work
- 30 seconds partner share: "Which piece did you pick?"
- Cold call 3-5 students -- name the piece + one strong line

TEACHER NOTES:
Choice paralysis is the enemy this week. Force the decision quickly. If a student says "I'm not proud of any of them" -- prompt them to pick the LEAST broken one and explain that publish is about FINISHING, not starting fresh.

WATCH FOR:
- Students who say "I want to do something new" -- redirect firmly: "Use one of your existing pieces. We polish them"
- Students who pick a strong piece -- celebrate publicly

[Literacy: Hook | VTLM 2.0: Establishing Purpose]`;

const NOTES_LI_SC = `SAY:
- Read the LI with me
- We are learning to choose ONE piece, plan how we will polish it, and decide what makes it display-worthy
- Three "I can" statements

DO:
- Choral read LI and SCs
- Brief: "Today we PLAN. Tomorrow we DRAFT. Wednesday we EDIT. Thursday we PUBLISH"

TEACHER NOTES:
SC1 -- decide which piece. SC2 (target) -- decide + plan ONE specific improvement for each part of the piece. SC3 -- decide + plan + identify one writing technique to add or strengthen.

WATCH FOR:
- Students who plan to "make it longer" -- redirect: "Specific improvements. Not just more words"

[Literacy: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_CHOOSE = `SAY:
- The two options. Look at the slide
- Narrative: setting + character + tension scene (around 250-300 words)
- Persuasive: introduction + body + counter + conclusion (around 250-300 words)
- Both go on one page. Both go on display
- Pick the one you are MOST PROUD of, OR the one you most want to improve

DO:
- Display the choice slide
- Quick poll: hands up for narrative, hands up for persuasive
- Note the rough split
- If a student has not chosen, sit with them for 30 seconds

TEACHER NOTES:
Most students will follow their proudest piece. Some will pick the genre they want to GROW in. Both are valid. Do not force a specific split.

WATCH FOR:
- Students who pick narrative because it is "easier" -- accept; many will write strong narratives
- Students who pick persuasive because they have strong opinions -- celebrate

[Literacy: Structure | VTLM 2.0: Choice and Voice]`;

const NOTES_IDO_NARRATIVE = `SAY:
- I will show you what a planned NARRATIVE publish piece looks like
- It uses ALL THREE Week 1 skills together
- Setting + Character + Tension

DO:
- Display the narrative plan slide
- Read each section
- Show: setting first, then character, then tension scene
- Highlight: one short hook + sensory detail + show-don't-tell + short sentences
- Briefly: "If you choose narrative, your finished page has these three parts in order"

TEACHER NOTES:
This model shows how the three Week 1 lessons combine into ONE story. Some students will use their setting from Lesson 1, their character from Lesson 2, and their tension scene from Lesson 3 -- stitching them together with a thin thread.

WATCH FOR:
- Students who notice the three Week 1 skills -- celebrate
- Students intimidated by combining three pieces -- reassure: "You already have these. We just stitch them together"

[Literacy: I Do | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_IDO_PERSUASIVE = `SAY:
- And here is the planned PERSUASIVE publish piece
- Introduction + ONE body PEEL + Counter + Conclusion
- Same shape as Week 2, but a TIGHTER, one-page version

DO:
- Display the persuasive plan slide
- Read each section
- Highlight: hook + position + signpost / PEEL / acknowledge + rebut / restate + final line
- Briefly: "If you choose persuasive, your finished page has these four parts in order"

TEACHER NOTES:
This model shows the persuasive piece distilled to one page. Note that ONE PEEL body paragraph is enough for the publish piece -- more would push beyond a single display page.

WATCH FOR:
- Students who want to use ALL THREE body paragraphs -- redirect: "One strong PEEL. Not three half-strong ones"

[Literacy: I Do | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_WEDO = `SAY:
- Together. We are going to plan a publish piece for an IMAGINARY student
- Their proudest piece is the setting paragraph from Lesson 1 -- Camp Green Lake
- They have picked NARRATIVE
- Let's plan how their one-page piece could look

DO:
- Display the We Do slide
- Cold call for: what setting do they keep? what character do they add? what tension scene?
- Sketch the one-page on the board (top: title, middle: paragraph, bottom: illustration/quote)
- Aim: full one-page plan visible on the board

TEACHER NOTES:
This is rehearsal for the students' own planning. The imaginary-student exercise removes ego from the planning process and lets students see the moves before they make their own decisions.

WATCH FOR:
- Students who suggest specific Week 1 skills -- celebrate
- Students who plan TOO much -- redirect: "One page. Tight"

[Literacy: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_CFU_BUILD = `SAY:
- Quick check. Two publish plans for the persuasive piece. Which one is realistic for ONE PAGE?

DO:
- Display both
- Show Me Fingers
- Cold call 1-2 students

TEACHER NOTES:
A is overcrowded -- five PEEL paragraphs would not fit. B is realistic -- intro + one body + counter + conclusion. B will fit on one page.

WATCH FOR:
- Students who choose A because "more is better" -- redirect: "Quality, not quantity. One page"

[Literacy: CFU | VTLM 2.0: Formative Assessment]`;

const NOTES_CFU_REVEAL = `SAY:
- Stronger plan: B
- B fits on ONE page -- intro + one body + counter + conclusion
- A is an essay, not a publish piece

DO:
- Display the reveal banner

CFU CHECKPOINT:
Technique: Show Me Fingers
Script:
- "Which plan fits one page?"
- Scan for: most students choose B (>=80%)
PROCEED (>=80%): Release to You Do.
PIVOT (<80%): Use the optional re-teach slide that follows -- it shows what fits in 250 words.

[Literacy: CFU Reveal | VTLM 2.0: Formative Assessment]`;

const NOTES_RETEACH = `SAY:
- Quick re-teach -- what fits in 250 words?
- Look at the slide -- I have written 250 words exactly
- Notice how MUCH fits and how LITTLE fits
- About 4 paragraphs of writing maximum
- Plan for that, not for more

DO:
- Display the re-teach slide
- Read aloud the 250-word example
- Count paragraphs (4) and sentences (about 15-18)
- Re-check: ask students to circle which 4 sections of their plan they will write

TEACHER NOTES:
OPTIONAL. Use only if CFU < 80%. Different approach: show ACTUAL 250-word piece on screen so students see what a one-page publish piece really looks like.

[Literacy: Optional Re-teach | VTLM 2.0: Reteach with Different Approach]`;

const NOTES_YOUDO = `SAY:
- Your turn. Use the Publish Plan
- Decision: narrative or persuasive (circle on page 1)
- Choose your starting piece from Weeks 1 or 2
- Identify ONE specific improvement for each section
- Sketch the one-page layout
- 15 minutes. I will circulate

DO:
- Display the You Do slide
- Distribute the plan and SC checklist
- Circulate -- focus on decision making, not writing
- Quick conferences: "What is your ONE specific improvement?"

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the same setup as the We Do (Camp Green Lake setting). Use the sentence frames on the back of the SC checklist
- Extra Notes: These students still hit SC1 and SC2 -- the frames support their planning

EXTENDING PROMPT:
- Task: After planning the words, also plan ONE visual feature -- an illustration, a map of Camp Green Lake, a quote in a box, a title with calligraphy
- Extra Notes: This is the display feature for the classroom wall

TEACHER NOTES:
Today is PLANNING, not writing. Students should leave with:
1. A clear choice (narrative or persuasive)
2. A clear starting piece from Weeks 1-2
3. One specific improvement per section
4. A one-page layout sketch

Tomorrow they DRAFT.

WATCH FOR:
- Students drifting into writing -- redirect: "Plan today. Write tomorrow"
- Students stuck choosing -- give the 30-second rule

[Literacy: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket
- Write ONE sentence: which piece are you publishing, and what is the ONE thing you will improve?
- 2 minutes

DO:
- 2 minutes silent
- Collect

TEACHER NOTES:
Use the exit tickets to plan Lesson 9 circulation. Students with vague answers ("I will make it better") need a conference at the start of Lesson 9.

[Literacy: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Self-check
- Show on fingers 1 to 5
- Partner share: which piece did you choose, and why?

DO:
- Run fingers check
- 60 seconds partner share
- Briefly: "Tomorrow we DRAFT. Bring your plan and your original piece"

TEACHER NOTES:
End with momentum. Tomorrow is drafting day -- students need their planning and their original work.

[Literacy: Closing | VTLM 2.0: Review and Reflect]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Holes -- Lesson 8 -- Plan Publish Piece";

  // SLIDE 1 -- Title
  titleSlide(
    pres,
    "Publish Week -- Plan Your Piece",
    "Choose: Narrative or Persuasive",
    "Lesson 8  |  Week 3  |  Year 5/6 Literacy",
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
    "Find the Piece You Are Most Proud Of",
    [
      "2 minutes silent: read ALL your work from Weeks 1 + 2",
      "Pick the piece you are MOST PROUD of -- OR -- the one you most want to improve",
      "30 seconds partner share: which piece did you pick?",
    ],
    NOTES_HOOK,
    FOOTER
  );

  // SLIDE 4 -- LI / SC
  liSlide(
    pres,
    [
      "We are learning to choose, plan and prepare ONE piece of writing for classroom publication",
    ],
    [
      "I can decide which piece I will publish and explain why",
      "I can plan ONE specific improvement for each section of my piece",
      "I can identify one writing technique I will add or strengthen for display",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // SLIDE 5 -- The Choice
  contentSlide(
    pres,
    "Choose",
    C.SECONDARY,
    "Two Publish Options -- Pick One",
    [
      "NARRATIVE:    setting + character + tension scene  (around 250-300 words)",
      "PERSUASIVE:  introduction + ONE PEEL body + counter + conclusion  (around 250-300 words)",
      "Both fit on ONE page",
      "Both go on display in our classroom",
    ],
    NOTES_CHOOSE,
    FOOTER
  );

  // SLIDE 6 -- I Do (narrative plan)
  modellingSlide(
    pres,
    "I Do -- Watch Me",
    "Planned Narrative Publish Piece",
    "Title (top of page)\n\n[Story Title]\n\n--------------\n\nParagraph 1 -- SETTING\n(2 senses + short hook)\n\nParagraph 2 -- CHARACTER\n(show, don't tell + one feeling)\n\nParagraph 3 -- TENSION\n(short sentences + small detail + stop on the wait)\n\n--------------\n\n[Optional illustration at the bottom of the page]",
    "Example structure (the actual words come tomorrow):\n\nSETTING:\nThere is no lake at Camp Green Lake. The lake had dried up over a hundred years ago... [continues from Lesson 1]\n\nCHARACTER:\nStanley stopped in the doorway of Tent D. The other boys turned to look... [Lesson 2 skills applied to a specific scene]\n\nTENSION:\nThe lizard sat very still on the shovel handle. Stanley did not move. He could see one yellow spot... [Lesson 3 skills]\n\n(Notice: the three Week 1 skills come together in ONE story.)",
    NOTES_IDO_NARRATIVE,
    FOOTER
  );

  // SLIDE 7 -- I Do (persuasive plan)
  modellingSlide(
    pres,
    "I Do -- Watch Me",
    "Planned Persuasive Publish Piece",
    "Title (top of page)\n\n[Persuasive Title]\n\n--------------\n\nINTRODUCTION (Lesson 4)\n  - Hook\n  - Position\n  - Signpost (3 reasons)\n\nONE PEEL BODY (Lessons 5-6)\n  - Strongest reason\n  - Evidence from Holes\n  - Explanation\n  - Link\n\nCOUNTER (Lesson 7)\n  - Acknowledge\n  - Rebut\n  - Link\n\nCONCLUSION (Lesson 7)\n  - Restate\n  - Summarise\n  - Strong final line",
    "Why ONE body paragraph?\n\n- A publish piece fits ON ONE PAGE\n- The introduction is around 4 sentences\n- ONE PEEL body is around 4-6 sentences\n- The counter is around 3-4 sentences\n- The conclusion is around 3-5 sentences\n- Total: about 15-19 sentences = one page\n\nYour strongest body paragraph is enough.",
    NOTES_IDO_PERSUASIVE,
    FOOTER
  );

  // SLIDE 8 -- We Do
  {
    const s = pres.addSlide();
    addTopBar(s, C.SECONDARY);
    addBadge(s, "We Do", { color: C.SECONDARY });
    addTitle(s, "Together: Plan for an Imaginary Student");

    const cardY = CONTENT_TOP;
    const cardH = 2.0;
    addCard(s, 0.5, cardY, 9, cardH, { strip: C.SECONDARY, fill: C.WHITE });
    s.addText("Imaginary student profile:", {
      x: 0.75, y: cardY + 0.10, w: 8.5, h: 0.30,
      fontSize: 14, fontFace: FONT_B, color: C.SECONDARY, bold: true, margin: 0,
    });
    s.addText("- Proudest piece: setting paragraph from Lesson 1 (Camp Green Lake)\n- Choice: NARRATIVE\n- Strong vocabulary: barren, shimmer\n\nLet's plan their publish piece together:\n- Which character will they add? (Stanley? Zero? a new boy?)\n- Which tension scene? (the lizard? the warden? the heat?)\n- What ONE specific improvement to the setting?", {
      x: 0.75, y: cardY + 0.45, w: 8.5, h: cardH - 0.55,
      fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    const tipY = cardY + cardH + 0.18;
    const tipH = SAFE_BOTTOM - tipY;
    addCard(s, 0.5, tipY, 9, tipH, { strip: C.ACCENT, fill: C.BG_CARD });
    s.addText("On the board: sketch a one-page layout for the imaginary student", {
      x: 0.75, y: tipY + 0.10, w: 8.5, h: 0.30,
      fontSize: 13, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("- Top: title\n- Middle: 3 paragraphs (setting, character, tension)\n- Bottom: one illustration or pull-out quote", {
      x: 0.75, y: tipY + 0.42, w: 8.5, h: tipH - 0.55,
      fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_WEDO);
  }

  // SLIDE 9 + 10 -- CFU
  function buildCfuBase() {
    const slide = pres.addSlide();
    addTopBar(slide, C.ALERT);
    addBadge(slide, "CFU", { color: C.ALERT });
    addTitle(slide, "Which Plan Fits ONE Page?", { color: C.ALERT });

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
    const cardH = 1.40;
    addCard(slide, 0.5, cardY, 9, cardH, { strip: C.MUTED, fill: C.WHITE });
    slide.addText("A", {
      x: 0.7, y: cardY + 0.10, w: 0.5, h: 0.4,
      fontSize: 22, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0,
    });
    slide.addText("Persuasive plan: Hook + Position + Signpost + Body PEEL 1 + Body PEEL 2 + Body PEEL 3 + Counter A + Counter B + Conclusion + Call to Action + Author bio", {
      x: 1.2, y: cardY + 0.10, w: 8.0, h: cardH - 0.20,
      fontSize: 12, fontFace: FONT_H, color: C.CHARCOAL, italic: true, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    const cardBY = cardY + cardH + 0.18;
    const cardBH = SAFE_BOTTOM - cardBY - 0.55;
    addCard(slide, 0.5, cardBY, 9, cardBH, { strip: C.SECONDARY, fill: C.BG_CARD });
    slide.addText("B", {
      x: 0.7, y: cardBY + 0.10, w: 0.5, h: 0.4,
      fontSize: 22, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0,
    });
    slide.addText("Persuasive plan: Introduction (4 sentences) + ONE PEEL body (4-6 sentences) + Counter (3-4 sentences) + Conclusion (3-5 sentences)", {
      x: 1.2, y: cardBY + 0.14, w: 8.0, h: cardBH - 0.24,
      fontSize: 14, fontFace: FONT_H, color: C.CHARCOAL, italic: true, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    addFooter(slide, FOOTER);
    slide.addNotes(NOTES_CFU_BUILD);
    return slide;
  }

  withReveal(
    buildCfuBase,
    (slide) => {
      const revealY = 4.68;
      addCard(slide, 0.5, revealY, 9, 0.42, { fill: C.SUCCESS });
      slide.addText("Realistic for one page: B  --  intro + ONE body + counter + conclusion (~ 250 words)", {
        x: 0.5, y: revealY, w: 9, h: 0.42,
        fontSize: 13, fontFace: FONT_B, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      slide.addNotes(NOTES_CFU_REVEAL);
    }
  );

  // SLIDE 11 -- Optional Re-teach
  modellingSlide(
    pres,
    "Optional Re-teach",
    "What Fits in 250 Words?",
    "250 words = about one A4 page handwritten\n\nThat is about:\n- 4 paragraphs\n- 15-18 sentences\n- 4 main moves (intro / body / counter / conclusion)\n\nThat is the target.",
    "Example -- 250 words exactly:\n\n[An imaginary 250-word piece could go here -- count the paragraphs as the teacher reads]\n\nYour turn: circle the 4 main sections of your plan that you WILL write. Anything outside the circle stays out of the publish piece.",
    NOTES_RETEACH,
    FOOTER
  );

  // SLIDE 12 -- You Do
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "You Do", { color: C.PRIMARY, w: 1.5 });
    addTitle(s, "Plan YOUR Publish Piece");

    addCard(s, 0.5, CONTENT_TOP, 9, 1.85, { strip: C.PRIMARY, fill: C.WHITE });
    s.addText("Use the Publish Piece Plan", {
      x: 0.75, y: CONTENT_TOP + 0.10, w: 8.4, h: 0.30,
      fontSize: 14, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("Decide:\n- Narrative or Persuasive (circle on the plan)\n- Which piece from Weeks 1 + 2 you will polish\n- ONE specific improvement for each section\n- Sketch a one-page layout", {
      x: 0.75, y: CONTENT_TOP + 0.42, w: 8.4, h: 1.35,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    const tipY = CONTENT_TOP + 2.00;
    const tipH = SAFE_BOTTOM - tipY;
    addCard(s, 0.5, tipY, 9, tipH, { strip: C.ACCENT, fill: C.BG_CARD });
    s.addText("Planning Time: 15 minutes  |  No writing today -- just planning", {
      x: 0.75, y: tipY + 0.10, w: 8.5, h: 0.30,
      fontSize: 13, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("- Use the Success Criteria checklist to spot improvement targets\n- Sketch the layout on page 2 of the plan\n- Plan 1 visual feature (illustration, quote box, title style)", {
      x: 0.75, y: tipY + 0.42, w: 8.4, h: tipH - 0.55,
      fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  }

  // SLIDE 13 -- Exit Ticket
  contentSlide(
    pres,
    "Exit Ticket",
    C.ACCENT,
    "One Sentence",
    [
      "Which piece are you publishing?",
      "What is the ONE thing you will improve?",
      "Write your answer in ONE sentence at the bottom of your plan",
      "2 minutes -- drop on my desk",
    ],
    NOTES_EXIT,
    FOOTER
  );

  // SLIDE 14 -- Closing
  closingSlide(
    pres,
    {
      reflectionPrompt: "Partner share: which piece did you choose, and why?",
      scItems: [
        "I can decide which piece I will publish and explain why",
        "I can plan ONE specific improvement for each section of my piece",
        "I can identify one writing technique I will add or strengthen for display",
      ],
      selfAssessment: {
        prompt: "Show on your fingers: 1 (need help) to 5 (could teach it)",
        options: ["1-2", "3", "4-5"],
      },
    },
    NOTES_CLOSING
  );

  // ---- PDF: Publish Piece Plan ----------------------------------------------
  const plan = createPdf({ title: PLAN_RESOURCE.name });
  let planY = addPdfHeader(plan, "Publish Piece Plan", {
    color: C.PRIMARY,
    subtitle: "Choose narrative or persuasive. Plan the one-page piece.",
    lessonInfo: "Lesson 8 | Week 3 | Year 5/6 Literacy | Holes",
    showNameDate: true,
  });

  planY = addTipBox(plan, "Today is for PLANNING, not writing. Choose your piece, find ONE specific improvement for each section, sketch the layout. Tomorrow you draft.", planY, { color: C.PRIMARY });

  planY = addSectionHeading(plan, "Step 1 -- My choice", planY, { color: C.PRIMARY });
  planY = addBodyText(plan, "Circle ONE:  NARRATIVE  /  PERSUASIVE", planY, { fontSize: 11 });
  planY += 4;

  planY = addSectionHeading(plan, "Step 2 -- My starting piece (from Weeks 1 + 2)", planY, { color: C.PRIMARY });
  planY = addBodyText(plan, "Which piece am I most proud of, or most want to improve?", planY, { fontSize: 10, italic: true });
  planY = addLinedArea(plan, planY, 2, { lineSpacing: 20 });
  planY += 4;

  planY = addSectionHeading(plan, "Step 3 -- Specific improvements (one per section)", planY, { color: C.SECONDARY });
  planY = addBodyText(plan, "NARRATIVE -- setting: one specific improvement", planY, { fontSize: 11 });
  planY = addLinedArea(plan, planY, 1, { lineSpacing: 18 });
  planY = addBodyText(plan, "NARRATIVE -- character: one specific improvement", planY, { fontSize: 11 });
  planY = addLinedArea(plan, planY, 1, { lineSpacing: 18 });
  planY = addBodyText(plan, "NARRATIVE -- tension: one specific improvement", planY, { fontSize: 11 });
  planY = addLinedArea(plan, planY, 1, { lineSpacing: 18 });
  planY += 4;

  planY = addBodyText(plan, "PERSUASIVE -- intro: one specific improvement", planY, { fontSize: 11 });
  planY = addLinedArea(plan, planY, 1, { lineSpacing: 18 });
  planY = addBodyText(plan, "PERSUASIVE -- body PEEL: one specific improvement", planY, { fontSize: 11 });
  planY = addLinedArea(plan, planY, 1, { lineSpacing: 18 });
  planY = addBodyText(plan, "PERSUASIVE -- counter and conclusion: one specific improvement", planY, { fontSize: 11 });
  planY = addLinedArea(plan, planY, 1, { lineSpacing: 18 });

  addPdfFooter(plan, "Lesson 8 | Publish Plan -- Page 1");

  plan.addPage();
  let planY2 = addPdfHeader(plan, "Step 4 -- Sketch Your One-Page Layout", {
    color: C.PRIMARY,
    subtitle: "Draw boxes for: title, paragraphs, illustration, quote box",
    lessonInfo: "Lesson 8 | Week 3 | Year 5/6 Literacy",
    showNameDate: false,
  });

  planY2 = addTipBox(plan, "Draw a rough sketch of your one-page layout. Where will the title go? Where will the paragraphs go? Will you add an illustration, a map, a quote box?", planY2, { color: C.SECONDARY });

  planY2 = addLinedArea(plan, planY2, 18, { lineSpacing: 18 });

  addPdfFooter(plan, "Lesson 8 | Layout Sketch -- Page 2");

  plan.addPage();
  let planY3 = addPdfHeader(plan, "Exit Ticket", {
    color: C.ACCENT,
    subtitle: "One sentence -- piece + one improvement",
    lessonInfo: "Lesson 8 | Week 3",
    showNameDate: false,
  });

  planY3 = addTipBox(plan, "Write ONE sentence: which piece are you publishing, and what is the ONE thing you will improve?", planY3, { color: C.ACCENT });

  planY3 = addLinedArea(plan, planY3, 3, { lineSpacing: 24 });

  addPdfFooter(plan, "Lesson 8 | Exit Ticket");

  // ---- PDF: Success Criteria Checklist --------------------------------------
  const sc = createPdf({ title: SC_RESOURCE.name });
  let scY = addPdfHeader(sc, "Publish Success Criteria Checklist", {
    color: C.PRIMARY,
    subtitle: "Use this all week -- when planning, drafting, editing and publishing",
    lessonInfo: "Lesson 8 | Week 3 | Year 5/6 Literacy",
    showNameDate: true,
  });

  scY = addTipBox(sc, "Tick each item as you check it. Use the column that matches YOUR choice (narrative or persuasive).", scY, { color: C.PRIMARY });

  scY = addSectionHeading(sc, "NARRATIVE PUBLISH PIECE", scY, { color: C.PRIMARY });
  scY = addBodyText(sc, "__ My setting uses at least TWO senses", scY);
  scY = addBodyText(sc, "__ My setting has a short hook sentence (under 10 words)", scY);
  scY = addBodyText(sc, "__ My character is SHOWN through action (no feeling labels)", scY);
  scY = addBodyText(sc, "__ I use at least one body action and one small gesture", scY);
  scY = addBodyText(sc, "__ My tension scene uses AT LEAST 3 short sentences", scY);
  scY = addBodyText(sc, "__ I include one small specific detail in the tension scene", scY);
  scY = addBodyText(sc, "__ My piece is around 250-300 words", scY);
  scY = addBodyText(sc, "__ My piece fits on ONE page", scY);
  scY = addBodyText(sc, "__ I have included a TITLE", scY);
  scY = addBodyText(sc, "__ I have included ONE visual feature (illustration, quote, decorated title)", scY);
  scY += 8;

  scY = addSectionHeading(sc, "PERSUASIVE PUBLISH PIECE", scY, { color: C.SECONDARY });
  scY = addBodyText(sc, "__ My introduction has a HOOK, POSITION and SIGNPOST of three reasons", scY);
  scY = addBodyText(sc, "__ My body paragraph uses PEEL (Point, Evidence, Explanation, Link)", scY);
  scY = addBodyText(sc, "__ I include at least ONE specific detail from Holes as evidence", scY);
  scY = addBodyText(sc, "__ I use at least 3 emotive words", scY);
  scY = addBodyText(sc, "__ I use at least ONE rhetorical device (question, repetition or rule of three)", scY);
  scY = addBodyText(sc, "__ My counter-argument acknowledges, rebuts and links", scY);
  scY = addBodyText(sc, "__ My conclusion restates my position in NEW words", scY);
  scY = addBodyText(sc, "__ My conclusion ends on a strong final line", scY);
  scY = addBodyText(sc, "__ My piece is around 250-300 words", scY);
  scY = addBodyText(sc, "__ My piece fits on ONE page", scY);
  scY = addBodyText(sc, "__ I have included a TITLE", scY);
  scY = addBodyText(sc, "__ I have included ONE visual feature (illustration, quote, decorated title)", scY);
  scY += 8;

  scY = addSectionHeading(sc, "BOTH -- Editing Checklist", scY, { color: C.ACCENT });
  scY = addBodyText(sc, "__ Capital letters at the start of every sentence", scY);
  scY = addBodyText(sc, "__ Full stops, question marks or exclamation marks where needed", scY);
  scY = addBodyText(sc, "__ Quotation marks if I used dialogue", scY);
  scY = addBodyText(sc, "__ Spelling checked -- no obvious errors", scY);
  scY = addBodyText(sc, "__ Read aloud quietly -- it makes sense", scY);

  addPdfFooter(sc, "Lesson 8 | Success Criteria Checklist | USE ALL WEEK");

  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Holes_Lesson8.pptx` }),
    writePdf(plan, PLAN_PDF_PATH),
    writePdf(sc, SC_PDF_PATH),
  ]);

  console.log("PPTX written to " + `${OUT_DIR}/Holes_Lesson8.pptx`);
  console.log("Done: " + PLAN_RESOURCE.name);
  console.log("Done: " + SC_RESOURCE.name);
}

build().catch(console.error);
