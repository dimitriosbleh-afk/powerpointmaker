"use strict";

// Holes Unit - Lesson 5 (Week 2, Session 2): Persuasive Body -- Arguments with Evidence (PEEL)
// Year 5/6 Literacy
// Focus: turn signposted reasons into PEEL body paragraphs using textual evidence
// Anchor: Holes by Louis Sachar -- Chapters 5-7 (camp conditions), Ch 13 (Zero), Ch 28 (lizards) as quotable details

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
  modellingSlide, vocabSlide, quoteSlide,
} = T;

const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addLinedArea,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const SESSION_NUMBER = 5;
const FOOTER = "Persuasive Body | Lesson 5 | Week 2 | Year 5/6 Literacy | Holes";
const OUT_DIR = "output/Holes_Lesson5_Arguments_Evidence";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const PLAN_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "PEEL Body Paragraph Plan",
  "Student template: plan and write ONE PEEL body paragraph using evidence from Holes."
);
const MENTOR_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Mentor PEEL Paragraph",
  "Annotated model PEEL paragraph -- Point, Evidence, Explanation, Link."
);
const RESOURCE_ITEMS = [PLAN_RESOURCE, MENTOR_RESOURCE];
const PLAN_PDF_PATH = path.join(OUT_DIR, PLAN_RESOURCE.fileName);
const MENTOR_PDF_PATH = path.join(OUT_DIR, MENTOR_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Welcome back. Yesterday you wrote a persuasive introduction
- Today: turn your three reasons into BODY paragraphs
- One body paragraph per reason. Today you write ONE -- your strongest one

DO:
- Display title slide
- Read 2-3 strong position sentences from yesterday's exit ticket (no names) to celebrate
- Have student plan sheets from yesterday ready -- they will use them today

TEACHER NOTES:
The big move today is PEEL -- Point, Evidence, Explanation, Link. Students will write ONE PEEL paragraph from their strongest reason. Writing all three in one lesson is too much.

WATCH FOR:
- Students who lost yesterday's plan -- have spare plan sheets ready
- Students who changed sides overnight -- allow it; ask them to pick ONE clear position now

[Literacy: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- Two resources today
- The ${PLAN_RESOURCE.name} -- plan and write your strongest PEEL paragraph
- The ${MENTOR_RESOURCE.name} -- annotated model

DO:
- Print the plan (one per student)
- Print the mentor (one per student or pair)
- Have yesterday's plan sheets back on student desks
- Have the novel ready -- teachers may re-read 1-2 short passages to give students evidence to quote

[Literacy: Resources | VTLM 2.0: Student Resources]`;

const NOTES_HOOK = `SAY:
- Quick partner share
- Open yesterday's plan -- which side did you pick?
- 30 seconds: tell your partner the ONE reason you would pick as your strongest
- I will cold call 2-3 students to share

DO:
- 30 seconds partner share
- Cold call 2-3 students -- collect 2-3 strong reasons on the board
- Briefly: "These are POINTS. Today we turn each point into a paragraph with EVIDENCE"

TEACHER NOTES:
Activates yesterday's planning. Builds collective momentum. If a student says "I don't remember", let them pick a fresh reason now.

WATCH FOR:
- Students with strong reasons -- celebrate publicly
- Students stuck -- offer them a default ("unsafe" / "teaches responsibility")

[Literacy: Hook | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_LI_SC = `SAY:
- Read the LI with me
- We are learning to write a body paragraph using PEEL -- Point, Evidence, Explanation, Link
- Three "I can" statements

DO:
- Choral read LI and SCs
- Brief: "PEEL means every paragraph has FOUR sentences with FOUR different jobs"

TEACHER NOTES:
SC1 -- a clear topic sentence with the reason -- achievable for all. SC2 (target) -- full PEEL with one piece of evidence from Holes. SC3 -- evidence linked back to the position. Exit ticket targets SC2.

WATCH FOR:
- Students who think PEEL has to be exactly 4 sentences -- redirect: "Each LETTER is a sentence or two. PEEL is the SHAPE"

[Literacy: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_VOCAB_EVIDENCE = `SAY:
- One word today: EVIDENCE
- A noun. A fact or detail that PROVES your point
- In a story, evidence comes from the TEXT -- a description, an event, a character's behaviour
- It is your job to SHOW the reader where you got your point

DO:
- Choral say EVIDENCE
- Quick example: "The boys dig holes 'five feet deep and five feet across'. That is evidence the work is exhausting"
- Ask: "Where can you find evidence in a novel?" -- 30 seconds partner share

TEACHER NOTES:
EVIDENCE is the key idea of the lesson. Spend extra time here -- pair the word with concrete examples from the novel. The mentor PDF has 4-5 useable lines from Holes for students who cannot find their own.

WATCH FOR:
- Students saying "evidence is just my opinion" -- redirect: "Evidence is FROM the text. Your opinion comes after"

[Literacy: Vocabulary | VTLM 2.0: Vocabulary Instruction]`;

const NOTES_PEEL = `SAY:
- The PEEL structure
- P -- POINT. The first sentence. State the reason
- E -- EVIDENCE. Use a fact or short quote from Holes
- E -- EXPLANATION. WHY does this evidence prove your point?
- L -- LINK. Tie the paragraph back to your position

DO:
- Display the PEEL slide
- Choral read each letter
- Quick check: "Which letter goes FIRST?" (P)
- Quick check: "Which letter ties back to your position?" (L)

TEACHER NOTES:
This is the structural backbone of any persuasive (or analytical) body paragraph. Students may have met PEEL or TEEL or PEEL+ before -- this is the same shape.

WATCH FOR:
- Students confused by two Es -- write the two jobs clearly: EVIDENCE = fact; EXPLANATION = why it matters

[Literacy: Structure | VTLM 2.0: Explicit Teaching / Structure]`;

const NOTES_IDO = `SAY:
- Watch me write a PEEL body paragraph
- Position: Camp Green Lake should be shut down
- My point: The camp is unsafe
- Watch me find the EVIDENCE from Holes, then EXPLAIN it, then LINK back

DO:
- Display the I Do slide
- Read the model aloud
- After reading, point to each part:
  - P (point)
  - E (evidence -- quoted detail from camp life or Ch 28)
  - E (explanation -- why this proves unsafe)
  - L (link -- back to "shut down")
- Show how the evidence does NOT need to be a long quote -- a paraphrased detail is enough

TEACHER NOTES:
This model uses the "unsafe" reason. Students who picked this reason can borrow the shape but must change the words. Most will use other reasons (inhumane / does not work / responsibility / second chance).

WATCH FOR:
- Students who try to copy the words -- redirect: "Copy the SHAPE. Change the words"

[Literacy: I Do | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_WEDO = `SAY:
- Together. Same novel. Different point
- Point: The camp does NOT actually help the boys learn anything
- We need EVIDENCE from Holes
- I will collect ideas on the board

DO:
- Display the We Do slide
- Cold call for evidence ideas
  - Stanley is innocent but still sent to camp
  - The warden cares about gold, not the boys
  - Mr. Sir is cruel
  - The boys dig the same hole every day with no teaching
- Pick the strongest evidence as a class
- Build the EXPLANATION sentence together: "This shows that..."
- Aim: P + E sentences written on the board, with EXPLANATION drafted orally

TEACHER NOTES:
We Do uses a DIFFERENT point from the I Do (does not help vs unsafe). Same structure, different content. Keep to 6-7 minutes.

WATCH FOR:
- Students suggesting evidence that does not match the point -- gently redirect: "Does this prove the camp does not HELP the boys?"
- Students suggesting strong evidence -- write it on the board verbatim

[Literacy: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_CFU_BUILD = `SAY:
- Quick check. Two PEEL paragraphs. Both about the SAME point: Camp Green Lake is inhumane
- Which one uses EVIDENCE properly?

DO:
- Display both
- Show Me Fingers
- Scan: most students should choose B
- Cold call 1-2 students: "Where is the evidence in B?"

TEACHER NOTES:
A makes the point but never uses evidence from Holes -- it is opinion only. B uses a specific detail from the novel (the heat, the canteen) and EXPLAINS why it proves inhumane. B is the stronger paragraph.

WATCH FOR:
- Students who choose B and name the evidence -- ready
- Students who choose A because it has "stronger language" -- redirect: "Stronger language is not the same as evidence. Where is the FACT?"

[Literacy: CFU | VTLM 2.0: Formative Assessment]`;

const NOTES_CFU_REVEAL = `SAY:
- Stronger paragraph: B
- B has a specific detail from the novel (one canteen, 35-degree heat)
- B then EXPLAINS why that proves the point
- A is loud but empty -- it never uses evidence

DO:
- Display the reveal banner
- Read B aloud, pause on the evidence
- Briefly: "Loud language only works when there is evidence behind it"

CFU CHECKPOINT:
Technique: Show Me Fingers (1 for A, 2 for B)
Script:
- "Which paragraph uses evidence?"
- Scan for: most students choose B (>=80%)
PROCEED (>=80%): Release to You Do.
PIVOT (<80%): Use the optional re-teach slide that follows -- it isolates evidence and explanation in a sentence frame.

[Literacy: CFU Reveal | VTLM 2.0: Formative Assessment]`;

const NOTES_RETEACH = `SAY:
- Quick second look at evidence + explanation
- Watch me use a SENTENCE FRAME to lock the two together
- FRAME: "In Holes, [evidence from the text]. This shows that [the point]."
- Now watch me fill it in
- Evidence: "the boys dig a hole five feet deep and five feet across, every day, in 35-degree heat"
- Point: "the work is exhausting"
- Filled: "In Holes, the boys dig a hole five feet deep and five feet across every day in 35-degree heat. This shows that the work is exhausting."

DO:
- Display the re-teach slide
- Write the FRAME on the board
- Fill it together
- Re-check: ask students to fill the same frame on their boards with their OWN evidence and point

TEACHER NOTES:
OPTIONAL slide. Use only if CFU was below 80%. Different approach: a TWO-SENTENCE frame that locks evidence and explanation together. Most students who missed the CFU will get this when the structure is explicit.

WATCH FOR:
- Students who write "In Holes, [vague feeling]. This shows that [vague feeling]." -- prompt: "Use a SPECIFIC detail. Not a feeling"
- Students who fill the frame with real evidence -- ready

[Literacy: Optional Re-teach | VTLM 2.0: Reteach with Different Approach]`;

const NOTES_YOUDO = `SAY:
- Your turn. Take YOUR strongest reason from yesterday and write a PEEL body paragraph
- Use the PEEL template
- The mentor sheet has a list of useable evidence from Holes if you need it
- 18 minutes. I will circulate

DO:
- Display the You Do slide
- Distribute the PEEL plan and mentor
- Circulate -- prioritise students who needed the re-teach
- Quick conferences: "Read me your point. What is your evidence?"

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the sentence frames on the back of the mentor sheet. Pick ONE evidence detail from the mentor's evidence list. Write your paragraph by filling the frames
- Extra Notes: These students still hit SC1 and SC2 -- the frame supports them through evidence and explanation

EXTENDING PROMPT:
- Task: After your paragraph, write the TOPIC SENTENCE for a SECOND PEEL paragraph using your second strongest reason. (Just the topic sentence, not the whole paragraph.)
- Extra Notes: This sets up Lesson 6 and 7 for confident writers

TEACHER NOTES:
Students write ONE PEEL paragraph -- their strongest reason. Writing all three would dilute quality. Tomorrow we add emotive and rhetorical language; the day after we add counter-arguments and conclusions.

WATCH FOR:
- Students who write a point with no evidence -- prompt: "What did Sachar SAY about this?"
- Students who use a paraphrase rather than a direct quote -- accept; paraphrase is valid evidence
- Students who land all four PEEL letters -- celebrate publicly

[Literacy: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket
- On your plan, write ONE EVIDENCE sentence for this point: "Stanley does not deserve to be at Camp Green Lake"
- Use evidence from Holes -- one detail or short paraphrase
- 2 minutes

DO:
- 2 minutes silent
- Collect

TEACHER NOTES:
Exit ticket targets SC2 -- evidence use. Look for: students who can identify a specific detail (Stanley is innocent of stealing the shoes / Stanley was hit on the head by the shoes / the shoes fell from a bridge).

WATCH FOR:
- Students who write opinion only -- collect; flag
- Students who use specific detail -- evidence of SC2

[Literacy: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Self-check
- Show on fingers 1 to 5 for each SC
- Partner share: what is ONE specific detail from Holes you used today?

DO:
- Run fingers check
- 30 seconds partner share
- Briefly: "Tomorrow we add EMOTIVE and RHETORICAL language to make the writing punch"

TEACHER NOTES:
Sets up Lesson 6. Note any students who could not find a specific detail -- they may need extra evidence support tomorrow.

[Literacy: Closing | VTLM 2.0: Review and Reflect]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Holes -- Lesson 5 -- PEEL Body Paragraphs";

  // SLIDE 1 -- Title
  titleSlide(
    pres,
    "Arguments with Evidence",
    "Point, Evidence, Explanation, Link (PEEL)",
    "Lesson 5  |  Week 2  |  Year 5/6 Literacy",
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
    "Yesterday's Strongest Reason",
    [
      "Open yesterday's plan -- which SIDE did you pick?",
      "30 seconds: tell your partner the ONE reason you would pick as your strongest",
      "I will cold call 2-3 students to share",
    ],
    NOTES_HOOK,
    FOOTER
  );

  // SLIDE 4 -- LI/SC
  liSlide(
    pres,
    [
      "We are learning to write a body paragraph using PEEL -- Point, Evidence, Explanation, Link",
    ],
    [
      "I can write a topic sentence that states my reason (P)",
      "I can use one piece of evidence from Holes and explain why it proves my point (E + E)",
      "I can link the paragraph back to my overall position (L)",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // SLIDE 5 -- Vocab
  vocabSlide(
    pres,
    "evidence",
    "noun",
    "A fact or detail that PROVES your point. In a novel, evidence comes from the text -- a description, an event or a character's behaviour.",
    "The fact that the boys dig holes 'five feet deep' is evidence the work is exhausting.",
    NOTES_VOCAB_EVIDENCE,
    FOOTER
  );

  // SLIDE 6 -- PEEL structure
  contentSlide(
    pres,
    "Structure",
    C.SECONDARY,
    "PEEL -- the Body Paragraph",
    [
      "P -- POINT:                first sentence. State your reason clearly",
      "E -- EVIDENCE:        a fact, detail or short quote from Holes",
      "E -- EXPLANATION: why does this evidence PROVE your point?",
      "L -- LINK:                   tie the paragraph back to your position",
    ],
    NOTES_PEEL,
    FOOTER
  );

  // SLIDE 7 -- I Do
  modellingSlide(
    pres,
    "I Do -- Watch Me",
    "Model PEEL Paragraph: Camp Green Lake is Unsafe",
    "Position:\n\nCamp Green Lake should be shut down.\n\nMy point for this paragraph:\n\nThe camp is unsafe.\n\nWatch the four parts:\n- P (point)\n- E (evidence)\n- E (explanation)\n- L (link)",
    "P:  Firstly, Camp Green Lake puts the boys in real danger every single day.\n\nE:  In Holes, the boys dig holes five feet deep in 35-degree heat, surrounded by deadly yellow-spotted lizards.\n\nE:  This shows the camp is not only physically exhausting -- it puts the boys' lives at risk every time they pick up a shovel.\n\nL:  For this reason alone, Camp Green Lake should be closed immediately.\n\n(My choices:\n- Firstly -- signposts the first reason\n- one specific detail (35 degrees, yellow-spotted lizards)\n- 'This shows...' explanation\n- 'For this reason' loops back to position)",
    NOTES_IDO,
    FOOTER
  );

  // SLIDE 8 -- We Do
  {
    const s = pres.addSlide();
    addTopBar(s, C.SECONDARY);
    addBadge(s, "We Do", { color: C.SECONDARY });
    addTitle(s, "Together: Different Point -- the Camp Does Not HELP the Boys");

    const cardY = CONTENT_TOP;
    const cardH = 2.0;
    addCard(s, 0.5, cardY, 9, cardH, { strip: C.SECONDARY, fill: C.WHITE });
    s.addText("Where in Holes can we find EVIDENCE that the camp does not help?", {
      x: 0.75, y: cardY + 0.10, w: 8.5, h: 0.30,
      fontSize: 14, fontFace: FONT_B, color: C.SECONDARY, bold: true, margin: 0,
    });
    s.addText("- Stanley is INNOCENT but is still sent to camp\n- The boys dig the SAME hole every day -- no lessons, no teaching\n- The Warden cares about GOLD, not about helping the boys\n- Mr. Sir is cruel to the boys (not kind, not patient)\n- Zero cannot read -- nobody at camp teaches him", {
      x: 0.75, y: cardY + 0.45, w: 8.5, h: cardH - 0.55,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    const tipY = cardY + cardH + 0.18;
    const tipH = SAFE_BOTTOM - tipY;
    addCard(s, 0.5, tipY, 9, tipH, { strip: C.ACCENT, fill: C.BG_CARD });
    s.addText("Together on the board: P + E", {
      x: 0.75, y: tipY + 0.10, w: 8.5, h: 0.30,
      fontSize: 13, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("- Pick ONE piece of evidence above\n- Write the POINT sentence (\"Secondly, the camp does not help the boys learn anything.\")\n- Write the EVIDENCE sentence (\"In Holes, ...\")\n- Draft the EXPLANATION orally as a class", {
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
    addTitle(slide, "Which Paragraph Uses Evidence?", { color: C.ALERT });

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
    const cardH = 1.5;
    addCard(slide, 0.5, cardY, 9, cardH, { strip: C.MUTED, fill: C.WHITE });
    slide.addText("A", {
      x: 0.7, y: cardY + 0.10, w: 0.5, h: 0.4,
      fontSize: 22, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0,
    });
    slide.addText("\"Camp Green Lake is so cruel and inhumane. It is awful and unfair and the boys hate it. Nobody should have to live like that. It is the worst place ever.\"", {
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
    slide.addText("\"Camp Green Lake is inhumane. In Holes, the boys dig holes in 35-degree heat with only one canteen of water a day. This shows that the camp puts children's bodies in danger. No camp should treat children this way.\"", {
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
      slide.addText("Stronger paragraph: B  --  uses a SPECIFIC detail from the novel, then EXPLAINS why it proves the point", {
        x: 0.5, y: revealY, w: 9, h: 0.42,
        fontSize: 12, fontFace: FONT_B, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      slide.addNotes(NOTES_CFU_REVEAL);
    }
  );

  // SLIDE 11 -- Optional Re-teach
  modellingSlide(
    pres,
    "Optional Re-teach",
    "Lock Evidence and Explanation with a Frame",
    "FRAME:\n\n\"In Holes, [evidence from the text].\nThis shows that [point].\"\n\nTwo sentences. Two jobs.",
    "EXAMPLE:\n\nEvidence: \"the boys dig a hole five feet deep and five feet across, every day, in 35-degree heat\"\nPoint: \"the work is exhausting\"\n\nFilled:\n\n\"In Holes, the boys dig a hole five feet deep and five feet across, every day, in 35-degree heat. This shows that the work is exhausting.\"\n\nYour turn: fill the frame on your board with YOUR evidence and YOUR point.",
    NOTES_RETEACH,
    FOOTER
  );

  // SLIDE 12 -- You Do
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "You Do", { color: C.PRIMARY, w: 1.5 });
    addTitle(s, "Write YOUR PEEL Paragraph");

    addCard(s, 0.5, CONTENT_TOP, 9, 1.85, { strip: C.PRIMARY, fill: C.WHITE });
    s.addText("Use yesterday's plan", {
      x: 0.75, y: CONTENT_TOP + 0.10, w: 8.4, h: 0.30,
      fontSize: 14, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("Take your STRONGEST reason from yesterday\nTurn it into ONE PEEL body paragraph\nUse evidence from Holes (see the mentor sheet for a list)\nWrite 4 to 6 sentences total", {
      x: 0.75, y: CONTENT_TOP + 0.42, w: 8.4, h: 1.30,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    const tipY = CONTENT_TOP + 2.00;
    const tipH = SAFE_BOTTOM - tipY;
    addCard(s, 0.5, tipY, 9, tipH, { strip: C.ACCENT, fill: C.BG_CARD });
    s.addText("Writing Time: 18 minutes", {
      x: 0.75, y: tipY + 0.10, w: 5, h: 0.30,
      fontSize: 13, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("First:    Write your POINT (start with \"Firstly,\" or \"Secondly,\" or \"Most importantly,\")\nNext:    Write your EVIDENCE (\"In Holes,...\")\nThen:    Write the EXPLANATION (\"This shows that...\")\nFinally: Write the LINK back to your position", {
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
    "One Evidence Sentence -- Stanley Does Not Deserve Camp",
    [
      "Point: \"Stanley does not deserve to be at Camp Green Lake.\"",
      "Write ONE evidence sentence",
      "Use a specific detail from Holes -- start \"In Holes,...\"",
      "2 minutes -- drop on my desk",
    ],
    NOTES_EXIT,
    FOOTER
  );

  // SLIDE 14 -- Closing
  closingSlide(
    pres,
    {
      reflectionPrompt: "Partner share: what is ONE specific detail from Holes you used today?",
      scItems: [
        "I can write a topic sentence that states my reason (P)",
        "I can use one piece of evidence from Holes and explain why it proves my point (E + E)",
        "I can link the paragraph back to my overall position (L)",
      ],
      selfAssessment: {
        prompt: "Show on your fingers: 1 (need help) to 5 (could teach it)",
        options: ["1-2", "3", "4-5"],
      },
    },
    NOTES_CLOSING
  );

  // ---- PDF: PEEL Plan -------------------------------------------------------
  const plan = createPdf({ title: PLAN_RESOURCE.name });
  let planY = addPdfHeader(plan, "PEEL Body Paragraph Plan", {
    color: C.PRIMARY,
    subtitle: "Point + Evidence + Explanation + Link",
    lessonInfo: "Lesson 5 | Week 2 | Year 5/6 Literacy | Holes",
    showNameDate: true,
  });

  planY = addTipBox(plan, "Take your STRONGEST reason from yesterday. Plan it as a PEEL paragraph. Use evidence from Holes. Then write 4 to 6 sentences.", planY, { color: C.PRIMARY });

  planY = addSectionHeading(plan, "My position (from yesterday)", planY, { color: C.PRIMARY });
  planY = addLinedArea(plan, planY, 1, { lineSpacing: 20 });
  planY += 4;

  planY = addSectionHeading(plan, "P -- POINT (state the reason as a topic sentence)", planY, { color: C.PRIMARY });
  planY = addBodyText(plan, "Start with: \"Firstly,\" / \"Secondly,\" / \"Most importantly,\"", planY, { fontSize: 10, italic: true });
  planY = addLinedArea(plan, planY, 2, { lineSpacing: 20 });
  planY += 4;

  planY = addSectionHeading(plan, "E -- EVIDENCE (fact / detail / paraphrase from Holes)", planY, { color: C.SECONDARY });
  planY = addBodyText(plan, "Start with: \"In Holes, ...\"", planY, { fontSize: 10, italic: true });
  planY = addLinedArea(plan, planY, 2, { lineSpacing: 20 });
  planY += 4;

  planY = addSectionHeading(plan, "E -- EXPLANATION (why does this evidence prove your point?)", planY, { color: C.ACCENT });
  planY = addBodyText(plan, "Start with: \"This shows that...\" / \"This proves that...\"", planY, { fontSize: 10, italic: true });
  planY = addLinedArea(plan, planY, 2, { lineSpacing: 20 });
  planY += 4;

  planY = addSectionHeading(plan, "L -- LINK (tie back to your position)", planY, { color: C.PRIMARY });
  planY = addBodyText(plan, "Start with: \"For this reason,\" / \"This is why...\" / \"It is clear that...\"", planY, { fontSize: 10, italic: true });
  planY = addLinedArea(plan, planY, 2, { lineSpacing: 20 });

  addPdfFooter(plan, "Lesson 5 | PEEL Plan -- Page 1");

  plan.addPage();
  let planY2 = addPdfHeader(plan, "PEEL Body Paragraph -- Write Here", {
    color: C.PRIMARY,
    subtitle: "4 to 6 sentences. P -> E -> E -> L.",
    lessonInfo: "Lesson 5 | Week 2 | Year 5/6 Literacy",
    showNameDate: false,
  });

  planY2 = addTipBox(plan, "Use your plan. Write each part as one or two sentences. Read it back -- can your partner spot each PEEL letter?", planY2, { color: C.SECONDARY });

  planY2 = addLinedArea(plan, planY2, 14, { lineSpacing: 22 });

  addPdfFooter(plan, "Lesson 5 | PEEL Paragraph -- Page 2");

  plan.addPage();
  let planY3 = addPdfHeader(plan, "Exit Ticket", {
    color: C.ACCENT,
    subtitle: "One evidence sentence -- Stanley does not deserve Camp",
    lessonInfo: "Lesson 5 | Week 2",
    showNameDate: false,
  });

  planY3 = addTipBox(plan, "Point: \"Stanley does not deserve to be at Camp Green Lake.\" Write ONE evidence sentence from Holes that proves this. Start with \"In Holes,...\"", planY3, { color: C.ACCENT });

  planY3 = addLinedArea(plan, planY3, 4, { lineSpacing: 22 });

  addPdfFooter(plan, "Lesson 5 | Exit Ticket");

  // ---- PDF: Mentor PEEL Paragraph -------------------------------------------
  const mp = createPdf({ title: MENTOR_RESOURCE.name });
  let mpY = addPdfHeader(mp, "Mentor PEEL Paragraph -- Annotated", {
    color: C.PRIMARY,
    subtitle: "Camp Green Lake is unsafe -- annotated PEEL model + evidence bank from Holes",
    lessonInfo: "Lesson 5 | Week 2 | Year 5/6 Literacy",
    showNameDate: false,
  });

  mpY = addTipBox(mp, "This is a model PEEL paragraph for the point \"Camp Green Lake is unsafe\". Use the PATTERN. Your point, evidence and link will be different.", mpY, { color: C.PRIMARY });

  mpY = addSectionHeading(mp, "Model PEEL paragraph", mpY, { color: C.PRIMARY });
  mpY = addBodyText(mp, "Firstly, Camp Green Lake puts the boys in real danger every single day. In Holes, the boys dig holes five feet deep and five feet across in 35-degree heat, surrounded by deadly yellow-spotted lizards. This shows that the camp is not only physically exhausting -- it puts the boys' lives at risk every time they pick up a shovel. For this reason alone, Camp Green Lake should be shut down immediately.", mpY, { fontSize: 12 });
  mpY += 14;

  mpY = addSectionHeading(mp, "Annotations -- P (point)", mpY, { color: C.PRIMARY });
  mpY = addBodyText(mp, "\"Firstly, Camp Green Lake puts the boys in real danger every single day.\" -- one sentence. Names the reason. Signposts as the first reason.", mpY, { fontSize: 10 });
  mpY += 6;

  mpY = addSectionHeading(mp, "Annotations -- E (evidence)", mpY, { color: C.SECONDARY });
  mpY = addBodyText(mp, "\"In Holes, the boys dig holes five feet deep and five feet across in 35-degree heat, surrounded by deadly yellow-spotted lizards.\" -- specific details from the novel.", mpY, { fontSize: 10 });
  mpY += 6;

  mpY = addSectionHeading(mp, "Annotations -- E (explanation)", mpY, { color: C.ACCENT });
  mpY = addBodyText(mp, "\"This shows that the camp is not only physically exhausting -- it puts the boys' lives at risk every time they pick up a shovel.\" -- WHY the evidence matters. Goes beyond the literal.", mpY, { fontSize: 10 });
  mpY += 6;

  mpY = addSectionHeading(mp, "Annotations -- L (link)", mpY, { color: C.ALERT });
  mpY = addBodyText(mp, "\"For this reason alone, Camp Green Lake should be shut down immediately.\" -- loops back to the writer's overall position.", mpY, { fontSize: 10 });
  mpY += 10;

  mpY = addSectionHeading(mp, "Evidence bank from Holes (use these if you need)", mpY, { color: C.PRIMARY });
  mpY = addBodyText(mp, "- Every boy digs a hole five feet deep and five feet across, every day.", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "- Mr. Sir gives the boys only one canteen of water at a time.", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "- Stanley is innocent -- the shoes fell from a highway bridge onto his head.", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "- Zero (Hector Zeroni) cannot read; nobody at camp teaches him.", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "- Yellow-spotted lizards live in the holes and are deadly (Ch. 28).", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "- The Warden cares only about finding the gold, not the boys.", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "- Mr. Sir is cruel: he refills bottles slowly, in the dust.", mpY, { fontSize: 10 });
  mpY += 6;

  mpY = addSectionHeading(mp, "Sentence frames you can borrow", mpY, { color: C.PRIMARY });
  mpY = addBodyText(mp, "P: \"Firstly / Secondly / Most importantly, [the reason].\"", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "E: \"In Holes, [specific detail from the novel].\"", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "E: \"This shows that [why the detail matters].\"", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "L: \"For this reason, [position]. / This is why [position].\"", mpY, { fontSize: 10 });

  addPdfFooter(mp, "Lesson 5 | Mentor PEEL Paragraph -- REFERENCE + EVIDENCE BANK");

  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Holes_Lesson5.pptx` }),
    writePdf(plan, PLAN_PDF_PATH),
    writePdf(mp, MENTOR_PDF_PATH),
  ]);

  console.log("PPTX written to " + `${OUT_DIR}/Holes_Lesson5.pptx`);
  console.log("Done: " + PLAN_RESOURCE.name);
  console.log("Done: " + MENTOR_RESOURCE.name);
}

build().catch(console.error);
