"use strict";

// Holes Persuasive Unit (3 sessions) - Session 2: Persuasive Body -- Evidence (PEEL) + Emotive Punch
// Year 5/6 Literacy | Week 2
// Focus: turn a signposted reason into a PEEL body paragraph using textual evidence, then strengthen it
//        with emotive language + ONE rhetorical device. (Merges original Lessons 5 and 6.)
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

const SESSION_NUMBER = 2;
const FOOTER = "Persuasive Body | Session 2 | Year 5/6 Literacy | Holes";
const OUT_DIR = "output/Holes_Persuasive_S2_Body_Evidence_Emotive";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const PLAN_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "PEEL Paragraph Plan",
  "Student template: plan a PEEL body paragraph from Holes, then upgrade it with emotive words and one device."
);
const MENTOR_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Mentor Persuasive Body Paragraph",
  "Annotated model: PEEL paragraph + emotive/rhetorical upgrade, plus an evidence bank and word bank."
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
- Today: take your STRONGEST reason and turn it into a body paragraph
- Two moves: first prove it with EVIDENCE using PEEL, then make it PUNCH with emotive language

DO:
- Display title slide
- Read 2-3 strong position sentences from yesterday's exit ticket (no names) to celebrate
- Have student plan sheets from yesterday ready -- they will use them today

TEACHER NOTES:
This session merges two moves: PEEL (Point, Evidence, Explanation, Link) is the core; emotive + rhetorical language is the upgrade. Students write ONE strong body paragraph from their strongest reason. Get the PEEL solid first, then strengthen the words.

WATCH FOR:
- Students who lost yesterday's plan -- have spare plan sheets ready
- Students who changed sides overnight -- allow it; ask them to pick ONE clear position now

[Literacy: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- Two resources today
- The ${PLAN_RESOURCE.name} -- plan your PEEL paragraph, then upgrade the words
- The ${MENTOR_RESOURCE.name} -- an annotated model with an evidence bank and a word bank

DO:
- Print the plan (one per student)
- Print the mentor (one per student or pair)
- Have yesterday's plan sheets back on student desks
- Have the novel ready -- you may re-read 1-2 short passages to give students evidence to quote

[Literacy: Resources | VTLM 2.0: Student Resources]`;

const NOTES_HOOK = `SAY:
- Quick partner share
- Open yesterday's plan -- which side did you pick?
- 30 seconds: tell your partner the ONE reason you would pick as your strongest
- I will cold call 2-3 students to share

DO:
- 30 seconds partner share
- Cold call 2-3 students -- collect 2-3 strong reasons on the board
- Briefly: "These are POINTS. Today we turn one point into a paragraph with EVIDENCE"

TEACHER NOTES:
Activates yesterday's planning. Builds collective momentum. If a student says "I don't remember", let them pick a fresh reason now.

WATCH FOR:
- Students with strong reasons -- celebrate publicly
- Students stuck -- offer them a default ("unsafe" / "teaches responsibility")

[Literacy: Hook | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_LI_SC = `SAY:
- Read the LI with me
- We are learning to write a body paragraph that PROVES a reason with evidence -- then make it punch
- Three "I can" statements

DO:
- Choral read LI and SCs
- Brief: "PEEL is the SHAPE: Point, Evidence, Explanation, Link. Then we upgrade the words"

TEACHER NOTES:
SC1 -- a clear topic sentence with the reason -- achievable for all. SC2 (target) -- full PEEL with one piece of evidence from Holes, explained and linked. SC3 -- the paragraph is strengthened with emotive words and one rhetorical device. Exit ticket targets SC2 (evidence use).

WATCH FOR:
- Students who think PEEL has to be exactly 4 sentences -- redirect: "Each LETTER is a sentence or two. PEEL is the SHAPE"

[Literacy: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_VOCAB_EVIDENCE = `SAY:
- First word today: EVIDENCE
- A noun. A fact or detail that PROVES your point
- In a story, evidence comes from the TEXT -- a description, an event, a character's behaviour
- It is your job to SHOW the reader where you got your point

DO:
- Choral say EVIDENCE
- Quick example: "The boys dig holes 'five feet deep and five feet across'. That is evidence the work is exhausting"
- Ask: "Where can you find evidence in a novel?" -- 30 seconds partner share

TEACHER NOTES:
EVIDENCE is the key idea of the lesson. Spend extra time here -- pair the word with concrete examples from the novel. The mentor PDF has an evidence bank from Holes for students who cannot find their own.

WATCH FOR:
- Students saying "evidence is just my opinion" -- redirect: "Evidence is FROM the text. Your opinion comes after"

[Literacy: Vocabulary | VTLM 2.0: Vocabulary Instruction]`;

const NOTES_VOCAB_EMOTIVE = `SAY:
- Second word today: EMOTIVE
- An adjective. It means causing strong feeling
- Emotive language uses words that make the reader FEEL something
- "Tired" is mild. "Exhausted" is emotive. "Broken" is even more emotive

DO:
- Choral say EMOTIVE
- Quick ladder on the board:
  - mild:  tired
  - emotive:  exhausted
  - emotive +:  broken
- Ask: "Give me an emotive word for 'angry'" -- 30 seconds partner share (furious / livid / outraged)

TEACHER NOTES:
We meet EMOTIVE now so it is ready for the upgrade move later in the lesson. Keep the ladder on the board -- students will use it when they strengthen their paragraph.

WATCH FOR:
- Students using a thesaurus word they do not know -- prompt: "Read that out loud. Does it sound natural?"

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
This is the structural backbone of any persuasive (or analytical) body paragraph. Students may have met PEEL or TEEL before -- this is the same shape.

WATCH FOR:
- Students confused by two Es -- write the two jobs clearly: EVIDENCE = fact; EXPLANATION = why it matters

[Literacy: Structure | VTLM 2.0: Explicit Teaching / Structure]`;

const NOTES_IDO1 = `SAY:
- Watch me write a PEEL body paragraph
- Position: Camp Green Lake should be shut down
- My point: The camp is unsafe
- Watch me find the EVIDENCE from Holes, then EXPLAIN it, then LINK back

DO:
- Display the I Do slide
- Read the model aloud
- After reading, point to each part:
  - P (point)
  - E (evidence -- a quoted detail from camp life or Ch 28)
  - E (explanation -- why this proves unsafe)
  - L (link -- back to "shut down")
- Show how the evidence does NOT need to be a long quote -- a paraphrased detail is enough

TEACHER NOTES:
This model uses the "unsafe" reason. Students who picked this reason can borrow the shape but must change the words. Most will use other reasons (inhumane / does not work / responsibility / second chance). We will UPGRADE this same paragraph in the next I Do.

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
PROCEED (>=80%): Move on to the emotive upgrade, then release to You Do.
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
OPTIONAL slide. Use only if the CFU was below 80%. Different approach: a TWO-SENTENCE frame that locks evidence and explanation together. Most students who missed the CFU will get this when the structure is explicit.

WATCH FOR:
- Students who write "In Holes, [vague feeling]. This shows that [vague feeling]." -- prompt: "Use a SPECIFIC detail. Not a feeling"
- Students who fill the frame with real evidence -- ready

[Literacy: Optional Re-teach | VTLM 2.0: Reteach with Different Approach]`;

const NOTES_DEVICES = `SAY:
- Now we make the paragraph PUNCH
- Three rhetorical devices -- pick ONE, not all three
- Look at the slide

DO:
- Display the rhetorical devices slide
- Choral read each row
- Brief example for each:
  - Rhetorical question: "Would you send your worst enemy to such a place?"
  - Repetition: "Day after day. Hole after hole. Boy after boy."
  - Rule of three: "It is unsafe, inhumane and unjustified."
- Ask: "Which device sounds best to YOU?" -- 30 seconds partner share

TEACHER NOTES:
One device per paragraph keeps the cognitive load manageable. Students pick ONE for their own writing. Pair this with the emotive ladder from the vocabulary slide -- together they are the upgrade toolkit.

WATCH FOR:
- Students who confuse a rhetorical question with a real question -- redirect: "You are not waiting for an answer. You are making the reader think"

[Literacy: Devices | VTLM 2.0: Explicit Teaching / Strategies]`;

const NOTES_IDO2 = `SAY:
- Watch me UPGRADE the PEEL paragraph from before
- Same point: Camp Green Lake is unsafe
- Same PEEL shape -- I do NOT change the structure
- I add THREE emotive words and ONE rhetorical device

DO:
- Display the I Do slide
- Read the BEFORE paragraph (left) aloud
- Read the AFTER paragraph (right) aloud
- Point to each upgrade: emotive word 1, emotive word 2, emotive word 3, rhetorical device
- Quick informal check: thumbs up if you can name the rhetorical device I used

TEACHER NOTES:
Same paragraph, BEFORE and AFTER. This makes the upgrade explicit and shows the structure does not change. Students do the same to their own paragraph in You Do. This is the SC3 stretch.

WATCH FOR:
- Students who notice the structure did not change -- celebrate
- Students who notice the emotive words pulling more weight -- celebrate

[Literacy: I Do | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_YOUDO = `SAY:
- Your turn. Take YOUR strongest reason from yesterday and write a PEEL body paragraph
- First get the PEEL solid: point, evidence, explanation, link
- Then strengthen the words: emotive upgrades + ONE rhetorical device
- The mentor sheet has an evidence bank and a word bank if you need them
- 20 minutes. I will circulate

DO:
- Display the You Do slide
- Distribute the plan and mentor
- Circulate -- prioritise students who needed the re-teach
- Quick conferences: "Read me your point. What is your evidence?" then "Which word could be more emotive?"

CFU CHECKPOINT:
Technique: Roving conference (read-aloud check)
Script:
- "Read me your PEEL paragraph. Point to your evidence."
- Scan for: a topic sentence (P), a specific detail from Holes (E), an explanation (E), a link (L)
PROCEED: Students with solid PEEL move to the emotive upgrade.
PIVOT: Students missing evidence go back to the mentor evidence bank and the sentence frame.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the sentence frames and the evidence bank on the mentor sheet. Pick ONE evidence detail. Write the PEEL paragraph by filling the frames, then swap in 2 emotive words
- Extra Notes: These students still hit SC1 and SC2 -- the frames carry them through evidence and explanation

EXTENDING PROMPT:
- Task: After your paragraph, write the TOPIC SENTENCE for a SECOND PEEL paragraph using your second strongest reason
- Extra Notes: This sets up the counter-argument session for confident writers

TEACHER NOTES:
Students write ONE strong body paragraph -- their strongest reason -- with evidence first, then emotive punch. Writing all three reasons would dilute quality. Next session we add the counter-argument and conclusion.

WATCH FOR:
- Students who write a point with no evidence -- prompt: "What did Sachar SAY about this?"
- Students who use a paraphrase rather than a direct quote -- accept; paraphrase is valid evidence
- Students who upgrade EVERY word -- redirect: "Three emotive words is enough. One device is enough"

[Literacy: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket
- On your plan, write ONE EVIDENCE sentence for this point: "Stanley does not deserve to be at Camp Green Lake"
- Use evidence from Holes -- one detail or short paraphrase. Start "In Holes,..."
- 2 minutes

DO:
- 2 minutes silent
- Collect

TEACHER NOTES:
Exit ticket targets SC2 -- evidence use. Look for: students who can identify a specific detail (Stanley is innocent of stealing the shoes / the shoes fell from a highway bridge onto his head).

WATCH FOR:
- Students who write opinion only -- collect; flag
- Students who use a specific detail -- evidence of SC2

[Literacy: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Self-check
- Show on fingers 1 to 5 for each SC
- Partner share: what is ONE specific detail from Holes you used today?

DO:
- Run fingers check
- 30 seconds partner share
- Briefly: "Next session we acknowledge the OTHER side, rebut it, and write a strong conclusion"

TEACHER NOTES:
Sets up Session 3. Note any students who could not find a specific detail -- they may need extra evidence support next session.

[Literacy: Closing | VTLM 2.0: Review and Reflect]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Holes Persuasive -- Session 2 -- Body Paragraph (Evidence + Emotive)";

  // SLIDE 1 -- Title
  titleSlide(
    pres,
    "Arguments with Evidence -- and Punch",
    "Build with PEEL, then Strengthen with Emotive Language",
    "Session 2  |  Week 2  |  Year 5/6 Literacy",
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

  // SLIDE 3 -- Launch
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
      "We are learning to write a body paragraph that proves a reason with evidence, then strengthen it with emotive language",
    ],
    [
      "I can write a topic sentence that states my reason (P)",
      "I can use evidence from Holes, explain why it proves my point, and link back to my position (E + E + L)",
      "I can strengthen my paragraph with emotive words and one rhetorical device",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // SLIDE 5 -- Vocab: evidence
  vocabSlide(
    pres,
    "evidence",
    "noun",
    "A fact or detail that PROVES your point. In a novel, evidence comes from the text -- a description, an event or a character's behaviour.",
    "The fact that the boys dig holes 'five feet deep' is evidence the work is exhausting.",
    NOTES_VOCAB_EVIDENCE,
    FOOTER
  );

  // SLIDE 6 -- Vocab: emotive
  vocabSlide(
    pres,
    "emotive",
    "adjective",
    "Causing strong feeling. Emotive language uses words that make the reader FEEL something. 'Tired' is mild. 'Exhausted' is emotive. 'Broken' is even more emotive.",
    "The boys' faces were exhausted, blistered and broken.",
    NOTES_VOCAB_EMOTIVE,
    FOOTER
  );

  // SLIDE 7 -- PEEL structure
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

  // SLIDE 8 -- I Do #1 (PEEL model with evidence)
  modellingSlide(
    pres,
    "I Do -- Watch Me",
    "Model PEEL Paragraph: Camp Green Lake is Unsafe",
    "Position:\n\nCamp Green Lake should be shut down.\n\nMy point for this paragraph:\n\nThe camp is unsafe.\n\nWatch the four parts:\n- P (point)\n- E (evidence)\n- E (explanation)\n- L (link)",
    "P:  Firstly, Camp Green Lake puts the boys in real danger every single day.\n\nE:  In Holes, the boys dig holes five feet deep in 35-degree heat, surrounded by deadly yellow-spotted lizards.\n\nE:  This shows the camp is not only physically exhausting -- it puts the boys' lives at risk every time they pick up a shovel.\n\nL:  For this reason alone, Camp Green Lake should be closed immediately.\n\n(My choices:\n- Firstly -- signposts the first reason\n- one specific detail (35 degrees, yellow-spotted lizards)\n- 'This shows...' explanation\n- 'For this reason' loops back to position)",
    NOTES_IDO1,
    FOOTER
  );

  // SLIDE 9 -- We Do
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

  // SLIDE 10 + 11 -- CFU (evidence)
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
      const revealY = 4.70;
      const revealH = 0.38;
      addCard(slide, 0.5, revealY, 9, revealH, { fill: C.SUCCESS });
      slide.addText("Stronger paragraph: B  --  uses a SPECIFIC detail from the novel, then EXPLAINS why it proves the point", {
        x: 0.5, y: revealY, w: 9, h: revealH,
        fontSize: 12, fontFace: FONT_B, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      slide.addNotes(NOTES_CFU_REVEAL);
    }
  );

  // SLIDE 12 -- Optional Re-teach (evidence frame)
  modellingSlide(
    pres,
    "Optional Re-teach",
    "Lock Evidence and Explanation with a Frame",
    "FRAME:\n\n\"In Holes, [evidence from the text].\nThis shows that [point].\"\n\nTwo sentences. Two jobs.",
    "EXAMPLE:\n\nEvidence: \"the boys dig a hole five feet deep and five feet across, every day, in 35-degree heat\"\nPoint: \"the work is exhausting\"\n\nFilled:\n\n\"In Holes, the boys dig a hole five feet deep and five feet across, every day, in 35-degree heat. This shows that the work is exhausting.\"\n\nYour turn: fill the frame on your board with YOUR evidence and YOUR point.",
    NOTES_RETEACH,
    FOOTER
  );

  // SLIDE 13 -- Devices (rhetorical devices, pick one)
  contentSlide(
    pres,
    "Devices",
    C.SECONDARY,
    "Make it Punch -- Pick ONE Rhetorical Device",
    [
      "Rhetorical question:    a question you are not waiting to have answered  -  \"Would you send your worst enemy to such a place?\"",
      "Repetition:                       repeat a word or phrase to drive a point home  -  \"Day after day. Hole after hole. Boy after boy.\"",
      "Rule of three:                 a list of three for emphasis  -  \"It is unsafe, inhumane and unjustified.\"",
    ],
    NOTES_DEVICES,
    FOOTER
  );

  // SLIDE 14 -- I Do #2 (BEFORE / AFTER upgrade)
  modellingSlide(
    pres,
    "I Do -- Watch Me",
    "Upgrade in Action: BEFORE -> AFTER",
    "BEFORE (the PEEL paragraph):\n\nFirstly, Camp Green Lake puts the boys in real danger every single day. In Holes, the boys dig holes five feet deep in 35-degree heat with deadly yellow-spotted lizards. This shows the camp is exhausting and risky. For this reason, Camp Green Lake should close.\n\n(Solid -- but a bit flat.)",
    "AFTER (the upgrade):\n\nIs it acceptable to send children to a place that puts their lives at risk every single day? Firstly, Camp Green Lake exposes the boys to brutal, exhausting and deadly conditions. In Holes, the boys dig holes five feet deep in punishing 35-degree heat, surrounded by venomous yellow-spotted lizards. This proves the camp is not just hard -- it is dangerous. For this reason, Camp Green Lake must be shut down immediately.\n\nUpgrades:\n- emotive: brutal, exhausting, deadly, punishing, venomous, dangerous\n- device: rule of three (brutal, exhausting, deadly)\n- bonus: rhetorical question opening",
    NOTES_IDO2,
    FOOTER
  );

  // SLIDE 15 -- You Do
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "You Do", { color: C.PRIMARY, w: 1.5 });
    addTitle(s, "Write YOUR Body Paragraph -- Evidence then Punch");

    addCard(s, 0.5, CONTENT_TOP, 9, 1.70, { strip: C.PRIMARY, fill: C.WHITE });
    s.addText("Take your STRONGEST reason from yesterday", {
      x: 0.75, y: CONTENT_TOP + 0.10, w: 8.4, h: 0.30,
      fontSize: 14, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("Turn it into ONE PEEL body paragraph using evidence from Holes\nThen strengthen it: 3 emotive words + ONE rhetorical device\nKeep the same PEEL shape -- do not change the structure", {
      x: 0.75, y: CONTENT_TOP + 0.45, w: 8.4, h: 1.15,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    const tipY = CONTENT_TOP + 1.85;
    const tipH = SAFE_BOTTOM - tipY;
    addCard(s, 0.5, tipY, 9, tipH, { strip: C.ACCENT, fill: C.BG_CARD });
    s.addText("Writing Time: 20 minutes", {
      x: 0.75, y: tipY + 0.10, w: 5, h: 0.30,
      fontSize: 13, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("First:    Write your POINT (\"Firstly,\" / \"Secondly,\" / \"Most importantly,\")\nNext:    Add EVIDENCE (\"In Holes,...\") then EXPLAIN it (\"This shows that...\")\nThen:    Write the LINK back to your position\nFinally: Strengthen -- swap in emotive words + add ONE device", {
      x: 0.75, y: tipY + 0.42, w: 8.4, h: tipH - 0.55,
      fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  }

  // SLIDE 16 -- Exit Ticket
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

  // SLIDE 17 -- Closing
  closingSlide(
    pres,
    {
      reflectionPrompt: "Partner share: what is ONE specific detail from Holes you used today?",
      scItems: [
        "I can write a topic sentence that states my reason (P)",
        "I can use evidence from Holes, explain why it proves my point, and link back to my position (E + E + L)",
        "I can strengthen my paragraph with emotive words and one rhetorical device",
      ],
      selfAssessment: {
        prompt: "Show on your fingers: 1 (need help) to 5 (could teach it)",
        options: ["1-2", "3", "4-5"],
      },
    },
    NOTES_CLOSING
  );

  // ---- PDF: PEEL Paragraph Plan ---------------------------------------------
  const plan = createPdf({ title: PLAN_RESOURCE.name });
  let planY = addPdfHeader(plan, "PEEL Paragraph Plan", {
    color: C.PRIMARY,
    subtitle: "Point + Evidence + Explanation + Link, then strengthen the words",
    lessonInfo: "Session 2 | Week 2 | Year 5/6 Literacy | Holes",
    showNameDate: true,
  });

  planY = addTipBox(plan, "Take your STRONGEST reason from yesterday. Plan it as a PEEL paragraph using evidence from Holes. Get the shape solid first -- you will strengthen the words on page 2.", planY, { color: C.PRIMARY });

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

  addPdfFooter(plan, "Session 2 | PEEL Plan -- Page 1");

  plan.addPage();
  let planY2 = addPdfHeader(plan, "Strengthen It -- Emotive + One Device", {
    color: C.PRIMARY,
    subtitle: "Same paragraph. Same shape. Stronger words. Then write it out neatly.",
    lessonInfo: "Session 2 | Week 2 | Year 5/6 Literacy",
    showNameDate: false,
  });

  planY2 = addTipBox(plan, "Do NOT change the PEEL structure. Swap 3 soft words for emotive words. Add ONE rhetorical device.", planY2, { color: C.PRIMARY });

  planY2 = addSectionHeading(plan, "Step 1 -- Emotive word bank (steal from this list)", planY2, { color: C.ACCENT });
  planY2 = addBodyText(plan, "Tired -> exhausted / drained / broken", planY2, { fontSize: 10 });
  planY2 = addBodyText(plan, "Bad -> brutal / cruel / inhumane / unjust", planY2, { fontSize: 10 });
  planY2 = addBodyText(plan, "Hot -> burning / scorching / blistering", planY2, { fontSize: 10 });
  planY2 = addBodyText(plan, "Mean -> heartless / merciless / ruthless", planY2, { fontSize: 10 });
  planY2 = addBodyText(plan, "Hard -> punishing / gruelling / relentless", planY2, { fontSize: 10 });
  planY2 += 6;

  planY2 = addSectionHeading(plan, "Step 2 -- Tick ONE device", planY2, { color: C.SECONDARY });
  planY2 = addBodyText(plan, "rhetorical question   /   repetition   /   rule of three", planY2, { fontSize: 11 });
  planY2 += 4;

  planY2 = addSectionHeading(plan, "Step 3 -- Write the FULL upgraded paragraph", planY2, { color: C.PRIMARY });
  planY2 = addLinedArea(plan, planY2, 8, { lineSpacing: 22 });

  addPdfFooter(plan, "Session 2 | Strengthen It -- Page 2");

  plan.addPage();
  let planY3 = addPdfHeader(plan, "Exit Ticket", {
    color: C.ACCENT,
    subtitle: "One evidence sentence -- Stanley does not deserve Camp",
    lessonInfo: "Session 2 | Week 2",
    showNameDate: false,
  });

  planY3 = addTipBox(plan, "Point: \"Stanley does not deserve to be at Camp Green Lake.\" Write ONE evidence sentence from Holes that proves this. Start with \"In Holes,...\"", planY3, { color: C.ACCENT });

  planY3 = addLinedArea(plan, planY3, 4, { lineSpacing: 22 });

  addPdfFooter(plan, "Session 2 | Exit Ticket");

  // ---- PDF: Mentor Persuasive Body Paragraph --------------------------------
  const mp = createPdf({ title: MENTOR_RESOURCE.name });
  let mpY = addPdfHeader(mp, "Mentor Persuasive Body Paragraph -- Annotated", {
    color: C.PRIMARY,
    subtitle: "Camp Green Lake is unsafe -- PEEL model + evidence bank from Holes",
    lessonInfo: "Session 2 | Week 2 | Year 5/6 Literacy",
    showNameDate: false,
  });

  mpY = addTipBox(mp, "This is a model PEEL paragraph for the point \"Camp Green Lake is unsafe\". Use the PATTERN. Your point, evidence and link will be different.", mpY, { color: C.PRIMARY });

  mpY = addSectionHeading(mp, "Model PEEL paragraph", mpY, { color: C.PRIMARY });
  mpY = addBodyText(mp, "Firstly, Camp Green Lake puts the boys in real danger every single day. In Holes, the boys dig holes five feet deep and five feet across in 35-degree heat, surrounded by deadly yellow-spotted lizards. This shows that the camp is not only physically exhausting -- it puts the boys' lives at risk every time they pick up a shovel. For this reason alone, Camp Green Lake should be shut down immediately.", mpY, { fontSize: 12 });
  mpY += 14;

  mpY = addSectionHeading(mp, "Annotations -- P / E / E / L", mpY, { color: C.PRIMARY });
  mpY = addBodyText(mp, "P (point): \"Firstly, Camp Green Lake puts the boys in real danger every single day.\" -- one sentence, names the reason.", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "E (evidence): \"In Holes, the boys dig holes five feet deep and five feet across in 35-degree heat...\" -- specific details from the novel.", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "E (explanation): \"This shows that the camp... puts the boys' lives at risk...\" -- WHY the evidence matters.", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "L (link): \"For this reason alone, Camp Green Lake should be shut down immediately.\" -- loops back to the position.", mpY, { fontSize: 10 });
  mpY += 10;

  mpY = addSectionHeading(mp, "Evidence bank from Holes (use these if you need)", mpY, { color: C.PRIMARY });
  mpY = addBodyText(mp, "- Every boy digs a hole five feet deep and five feet across, every day.", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "- Mr. Sir gives the boys only one canteen of water at a time.", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "- Stanley is innocent -- the shoes fell from a highway bridge onto his head.", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "- Zero (Hector Zeroni) cannot read; nobody at camp teaches him.", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "- Yellow-spotted lizards live in the holes and are deadly (Ch. 28).", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "- The Warden cares only about finding the gold, not the boys.", mpY, { fontSize: 10 });

  addPdfFooter(mp, "Session 2 | Mentor Body Paragraph -- PEEL + EVIDENCE BANK");

  mp.addPage();
  let mpY2 = addPdfHeader(mp, "Mentor -- the Upgraded Paragraph", {
    color: C.PRIMARY,
    subtitle: "Same PEEL paragraph, strengthened with emotive language + one device",
    lessonInfo: "Session 2 | Week 2 | Year 5/6 Literacy",
    showNameDate: false,
  });

  mpY2 = addTipBox(mp, "This is the SAME paragraph after an emotive + rhetorical upgrade. The structure did not change -- only the words. Use the PATTERN.", mpY2, { color: C.PRIMARY });

  mpY2 = addSectionHeading(mp, "Model upgraded paragraph", mpY2, { color: C.PRIMARY });
  mpY2 = addBodyText(mp, "Is it acceptable to send children to a place that puts their lives at risk every single day? Firstly, Camp Green Lake exposes the boys to brutal, exhausting and deadly conditions. In Holes, the boys dig holes five feet deep in punishing 35-degree heat, surrounded by venomous yellow-spotted lizards. This proves the camp is not just hard -- it is dangerous. For this reason, Camp Green Lake must be shut down immediately.", mpY2, { fontSize: 12 });
  mpY2 += 14;

  mpY2 = addSectionHeading(mp, "Annotations -- the upgrade", mpY2, { color: C.SECONDARY });
  mpY2 = addBodyText(mp, "Emotive words: brutal, exhausting, deadly, punishing, venomous, dangerous -- each makes the reader FEEL the danger.", mpY2, { fontSize: 10 });
  mpY2 = addBodyText(mp, "Rule of three: \"brutal, exhausting and deadly\" -- three emotive words in a row for weight.", mpY2, { fontSize: 10 });
  mpY2 = addBodyText(mp, "Rhetorical question opener: \"Is it acceptable to send children to a place that puts their lives at risk...?\" -- the reader cannot say yes.", mpY2, { fontSize: 10 });
  mpY2 += 10;

  mpY2 = addSectionHeading(mp, "Emotive word bank", mpY2, { color: C.ACCENT });
  mpY2 = addBodyText(mp, "Tired -> exhausted, drained, broken    |    Bad -> brutal, cruel, inhumane, unjust", mpY2, { fontSize: 10 });
  mpY2 = addBodyText(mp, "Hot -> burning, scorching, blistering    |    Mean -> heartless, merciless, ruthless", mpY2, { fontSize: 10 });
  mpY2 = addBodyText(mp, "Hard -> punishing, gruelling, relentless   |   Risky -> dangerous, deadly, perilous", mpY2, { fontSize: 10 });
  mpY2 += 8;

  mpY2 = addSectionHeading(mp, "Device cheat sheet (pick ONE)", mpY2, { color: C.PRIMARY });
  mpY2 = addBodyText(mp, "Rhetorical question: \"Is it [fair/acceptable] to ___?\"", mpY2, { fontSize: 10 });
  mpY2 = addBodyText(mp, "Repetition: \"___ after ___. ___ after ___.\"", mpY2, { fontSize: 10 });
  mpY2 = addBodyText(mp, "Rule of three: \"It is ___, ___ and ___.\"", mpY2, { fontSize: 10 });
  mpY2 += 8;

  mpY2 = addSectionHeading(mp, "Sentence frames you can borrow", mpY2, { color: C.PRIMARY });
  mpY2 = addBodyText(mp, "P: \"Firstly / Secondly / Most importantly, [the reason].\"   E: \"In Holes, [detail].\"", mpY2, { fontSize: 10 });
  mpY2 = addBodyText(mp, "E: \"This shows that [why it matters].\"   L: \"For this reason, [position].\"", mpY2, { fontSize: 10 });

  addPdfFooter(mp, "Session 2 | Mentor Upgrade -- REFERENCE + WORD BANK");

  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Holes_Persuasive_S2.pptx` }),
    writePdf(plan, PLAN_PDF_PATH),
    writePdf(mp, MENTOR_PDF_PATH),
  ]);

  console.log("PPTX written to " + `${OUT_DIR}/Holes_Persuasive_S2.pptx`);
  console.log("Done: " + PLAN_RESOURCE.name);
  console.log("Done: " + MENTOR_RESOURCE.name);
}

build().catch(console.error);
