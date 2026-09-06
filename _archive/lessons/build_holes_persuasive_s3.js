"use strict";

// Holes Persuasive Unit (3 sessions) - Session 3: Counter-Argument + Conclusion
// Year 5/6 Literacy | Week 2
// Focus: acknowledge the other side, rebut it, and write a strong conclusion that links back to the position
// (Was original Lesson 7.)

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

const SESSION_NUMBER = 3;
const FOOTER = "Counter + Conclusion | Session 3 | Year 5/6 Literacy | Holes";
const OUT_DIR = "output/Holes_Persuasive_S3_Counter_Conclusion";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const PLAN_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Counter and Conclusion Plan",
  "Student template: write a counter-argument paragraph and a strong conclusion."
);
const MENTOR_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Mentor Counter and Conclusion",
  "Annotated model counter-argument + conclusion -- shows acknowledge, rebut, and link back."
);
const RESOURCE_ITEMS = [PLAN_RESOURCE, MENTOR_RESOURCE];
const PLAN_PDF_PATH = path.join(OUT_DIR, PLAN_RESOURCE.fileName);
const MENTOR_PDF_PATH = path.join(OUT_DIR, MENTOR_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Last session of the persuasive unit
- Today: COUNTER-ARGUMENT (acknowledging the other side) and CONCLUSION (the closer)
- After today, your persuasive piece is COMPLETE

DO:
- Display title slide
- Read 2-3 strong sentences from last session's work (no names) to celebrate
- Have student body paragraphs from Session 2 ready on desks

TEACHER NOTES:
This session finishes the unit. By the end, students will have:
- an introduction (Session 1)
- one strong body paragraph with evidence and emotive punch (Session 2)
- one counter-argument paragraph (today)
- a conclusion (today)
That is a complete short persuasive piece they can publish or develop further.

WATCH FOR:
- Students who think "counter-argument" means I'm changing sides -- reassure: "No. You acknowledge the other side, then explain why your side is still stronger"

[Literacy: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- Two resources today
- The ${PLAN_RESOURCE.name} -- plan and write your counter + conclusion
- The ${MENTOR_RESOURCE.name} -- annotated model showing both moves

DO:
- Print plan (one per student)
- Print mentor (one per student or pair)
- Have student paragraphs from Sessions 1-2 back on desks

[Literacy: Resources | VTLM 2.0: Student Resources]`;

const NOTES_HOOK = `SAY:
- Quick partner share
- 30 seconds: tell your partner the ONE strongest argument the OTHER side could make
- Yes -- the OTHER side. Not yours
- I will cold call 2-3 students

DO:
- 30 seconds partner share
- Cold call -- collect 2-3 strongest "other side" arguments on the board
- Briefly: "Today we ACKNOWLEDGE one of those. Then we REBUT it"

TEACHER NOTES:
This launch forces students to consider the other side. Many will resist -- gently remind them that good persuasive writers anticipate the opposite view to defeat it.

WATCH FOR:
- Students who refuse: "There IS no other side" -- redirect: "Pretend. What COULD someone say?"
- Students who name a strong other-side argument -- celebrate; it shows real thinking

[Literacy: Hook | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_LI_SC = `SAY:
- Read the LI with me
- We are learning to write a COUNTER-ARGUMENT and a strong CONCLUSION
- Three "I can" statements

DO:
- Choral read LI and SCs
- Brief: "A counter-argument is THREE moves: acknowledge, rebut, link"
- Brief: "A conclusion is the LAST word. Make it count"

TEACHER NOTES:
SC1 -- one sentence that acknowledges the other side. Achievable. SC2 (target) -- counter-argument (acknowledge + rebut + link) + a conclusion that ends on a strong line. SC3 -- a conclusion that uses one rhetorical device or a call to action. Exit ticket targets SC2.

WATCH FOR:
- Students confused by "rebut" -- give the simple version: "Rebut = explain why your side is still stronger"

[Literacy: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_VOCAB_REBUT = `SAY:
- One word today: REBUT
- A verb. To prove that something is wrong, or weaker than you thought
- In persuasive writing, you ACKNOWLEDGE the other side, then REBUT it

DO:
- Choral say REBUT
- Quick example: "Some say Camp Green Lake teaches responsibility. I rebut this: it teaches obedience, not real responsibility"
- Ask: "What is the difference between rebut and just disagreeing?" -- partner share, 30 seconds

TEACHER NOTES:
REBUT carries a specific meaning -- you prove the other side weaker. It is not "fighting back" or "dismissing". Students who use it accurately are doing real persuasive work.

WATCH FOR:
- Students who hear "rebut" as rude -- reassure: "It is respectful. You consider the other side, THEN show your side is stronger"

[Literacy: Vocabulary | VTLM 2.0: Vocabulary Instruction]`;

const NOTES_STRUCTURE = `SAY:
- The counter-argument has THREE moves
- Look at the slide

DO:
- Display the structure slide
- Choral read each row
- Quick example for each:
  - ACKNOWLEDGE:  "Some people argue that Camp Green Lake teaches responsibility."
  - REBUT:                 "However, digging holes for the Warden is not responsibility -- it is forced labour."
  - LINK:                       "This is why the camp must be shut down."

TEACHER NOTES:
The three moves are sequential. Acknowledge first (respect the other side). Rebut second (show why your view is stronger). Link third (return to your overall position).

WATCH FOR:
- Students who acknowledge but never rebut -- prompt: "You agreed with them. Where is the 'however'?"

[Literacy: Structure | VTLM 2.0: Explicit Teaching / Structure]`;

const NOTES_IDO = `SAY:
- Watch me write a counter-argument paragraph
- My position: Camp Green Lake should be shut down
- The OTHER side might say: "But the camp teaches responsibility"
- I will ACKNOWLEDGE this. Then REBUT. Then LINK back

DO:
- Display the I Do slide
- Read aloud
- Point to each part:
  - ACKNOWLEDGE (uses "Some argue that..." or "It might be said that...")
  - REBUT (uses "However..." or "But...")
  - LINK (ties back to position)

TEACHER NOTES:
The model uses "Some argue that... However... This is why..." -- a clean three-move counter. Students will use the same shape with their own choice of opposing argument.

WATCH FOR:
- Students who notice the respectful tone -- celebrate
- Students who try to mock the other side -- redirect: "Respect. Then defeat"

[Literacy: I Do | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_CONCLUSION = `SAY:
- Now the CONCLUSION
- A conclusion is the LAST thing the reader reads
- It has THREE jobs:
  - RESTATE your position (in NEW words)
  - SUMMARISE your three reasons (briefly)
  - LEAVE the reader with a strong line (a question, a call to action, or a one-line punch)

DO:
- Display the conclusion structure slide
- Read each row
- Brief example: "Camp Green Lake is unsafe, inhumane and ineffective. For Stanley -- and for every boy who deserves a real future -- it must close."

TEACHER NOTES:
The strong final line is the difference between a flat conclusion and a powerful one. Students should leave the reader with a feeling -- not just a summary.

WATCH FOR:
- Students who just repeat the introduction word-for-word -- redirect: "Same idea. New words"

[Literacy: Structure | VTLM 2.0: Explicit Teaching / Structure]`;

const NOTES_WEDO = `SAY:
- Together. We write a counter + conclusion for the OPPOSITE side
- Position: Camp Green Lake should STAY OPEN
- Counter-argument: someone might say "the boys are in danger"
- Let's write the three moves together

DO:
- Display the We Do slide
- Cold call for an ACKNOWLEDGE sentence ("Some argue that the camp is dangerous.")
- Cold call for a REBUT ("However, life teaches us that overcoming difficulty builds strength.")
- Cold call for a LINK ("This is why Camp Green Lake should stay open.")
- Read the full counter aloud
- Then draft a conclusion together orally

TEACHER NOTES:
We Do uses the OTHER side from the I Do. Same structure, different position. This proves the move works on any side.

WATCH FOR:
- Students arguing passionately for the original side -- gently redirect: "We are practising the STRUCTURE"
- Students with a sharp rebut -- celebrate

[Literacy: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_CFU_BUILD = `SAY:
- Quick check. Two counter-arguments. Which one uses ALL THREE moves (acknowledge, rebut, link)?
- Read carefully

DO:
- Display both
- Show Me Fingers
- Cold call 1-2 students: "Where is the rebut in B?"

TEACHER NOTES:
A acknowledges but never rebuts -- it agrees with the other side and stops. B uses all three moves: acknowledge, rebut, link.

WATCH FOR:
- Students who pick B and name the "however" -- ready
- Students who pick A because it "sounds polite" -- redirect: "Polite is good. But you also need to PROVE your side"

[Literacy: CFU | VTLM 2.0: Formative Assessment]`;

const NOTES_CFU_REVEAL = `SAY:
- Stronger counter: B
- B has acknowledge + REBUT (the "however" line) + link
- A only acknowledges -- it never explains why our side is stronger

DO:
- Display the reveal banner
- Read B aloud

CFU CHECKPOINT:
Technique: Show Me Fingers
Script:
- "Which counter has all three moves?"
- Scan for: most students choose B (>=80%)
PROCEED (>=80%): Release to You Do.
PIVOT (<80%): Use the optional re-teach slide that follows.

[Literacy: CFU Reveal | VTLM 2.0: Formative Assessment]`;

const NOTES_RETEACH = `SAY:
- Re-teach -- a SENTENCE FRAME for the counter
- Watch me fill it in
- FRAME:
  - "Some argue that ____.
   However, ____.
   This is why ____."
- Filled example:
  - "Some argue that homework helps students learn.
   However, research shows that primary students learn more from outdoor play and family time.
   This is why homework should be limited."

DO:
- Display the re-teach slide
- Build LIVE on the board
- Re-check: students fill the same frame on their boards using their topic

TEACHER NOTES:
OPTIONAL. Use only if the CFU was below 80%. Different approach: a THREE-SENTENCE FRAME that locks the three moves in place.

[Literacy: Optional Re-teach | VTLM 2.0: Reteach with Different Approach]`;

const NOTES_YOUDO = `SAY:
- Your turn. Two parts today
- Part 1: write a COUNTER-ARGUMENT paragraph (3 to 4 sentences)
- Part 2: write a CONCLUSION (3 to 5 sentences)
- Use the plan. Use the mentor
- 20 minutes. I will circulate

DO:
- Display the You Do slide
- Distribute plan and mentor
- Circulate -- focus first on counter-arguments, then conclusions
- Quick conferences: "What is the other side's best point? How do you rebut it?"

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the sentence frames on the mentor sheet for both the counter and the conclusion. Pick ONE acknowledged point from your partner's other-side discussion
- Extra Notes: These students still hit SC1 and SC2 with the frames

EXTENDING PROMPT:
- Task: End your conclusion with ONE rhetorical device (a question, repetition, or a one-line punch). Make it the last thing the reader sees
- Extra Notes: This is the SC3 lift -- the writer's signature move

TEACHER NOTES:
Today is the densest writing day of the unit. Students complete a short persuasive piece (intro + body + counter + conclusion). Some will not finish -- that is fine; they can finish or polish it next time.

WATCH FOR:
- Students who skip the counter -- redirect: "It is the move that lifts your writing. Spend 5 minutes on it"
- Students with a punchy final line -- celebrate publicly

[Literacy: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket
- Write ONE counter-argument paragraph for this position: "Homework should be banned in primary schools"
- Use all three moves: acknowledge, rebut, link
- 3 minutes

DO:
- 3 minutes silent
- Collect

TEACHER NOTES:
Exit ticket targets SC2. Three sentences = acknowledge + rebut + link.

WATCH FOR:
- Students who use "Some argue..." + "However..." -- evidence of SC2

[Literacy: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Last session of the unit -- you have completed a short persuasive piece
- Self-check: show on fingers 1 to 5
- Partner share: what is the BEST sentence in your persuasive piece?

DO:
- Run fingers check
- 60 seconds partner share -- let them enjoy what they wrote
- Briefly: "You now have a complete short persuasive piece -- intro, body, counter and conclusion"

TEACHER NOTES:
End the unit with genuine acknowledgement. Students have written a complete short persuasive piece in 3 sessions -- that is real work.

[Literacy: Closing | VTLM 2.0: Review and Reflect / End of Persuasive Unit]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Holes Persuasive -- Session 3 -- Counter + Conclusion";

  // SLIDE 1 -- Title
  titleSlide(
    pres,
    "Counter-Argument + Conclusion",
    "Acknowledge, Rebut, Land the Closer",
    "Session 3  |  Week 2  |  Year 5/6 Literacy",
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
    "What is the OTHER Side's Strongest Argument?",
    [
      "30 seconds with your partner -- not your side, THEIR side",
      "What is the ONE strongest argument someone could use AGAINST your position?",
      "I will cold call 2-3 students to share",
    ],
    NOTES_HOOK,
    FOOTER
  );

  // SLIDE 4 -- LI / SC
  liSlide(
    pres,
    [
      "We are learning to write a COUNTER-ARGUMENT paragraph and a strong CONCLUSION",
    ],
    [
      "I can acknowledge the other side in one clear sentence",
      "I can write a counter (acknowledge + rebut + link) and a conclusion that restates my position",
      "I can end my conclusion on a strong final line (question, repetition, or one-line punch)",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // SLIDE 5 -- Vocab
  vocabSlide(
    pres,
    "rebut",
    "verb",
    "To prove that something is wrong, or weaker than you thought. In persuasive writing, you acknowledge the other side, then REBUT it -- you explain why your view is still stronger.",
    "Some say Camp Green Lake teaches responsibility. I rebut this: it teaches obedience, not responsibility.",
    NOTES_VOCAB_REBUT,
    FOOTER
  );

  // SLIDE 6 -- Counter-argument structure
  contentSlide(
    pres,
    "Structure",
    C.SECONDARY,
    "Counter-Argument -- Three Moves",
    [
      "ACKNOWLEDGE:    name the other side's argument fairly  -  \"Some argue that...\"",
      "REBUT:                         explain why your side is stronger  -  \"However...\"",
      "LINK:                              tie back to YOUR position  -  \"This is why...\"",
    ],
    NOTES_STRUCTURE,
    FOOTER
  );

  // SLIDE 7 -- I Do (counter-argument model)
  modellingSlide(
    pres,
    "I Do -- Watch Me",
    "Model Counter-Argument: Shut Down Camp Green Lake",
    "Position:\n\nCamp Green Lake should be shut down.\n\nThe other side might say:\n\n\"The camp teaches the boys responsibility.\"\n\nWatch me ACKNOWLEDGE, REBUT, LINK.",
    "\"Some argue that Camp Green Lake teaches the boys responsibility through hard work and routine. However, digging meaningless holes for the Warden is not real responsibility -- it is forced labour, with no education and no second chance. True responsibility comes from being taught, not from being broken. This is why Camp Green Lake must be shut down.\"\n\n(My choices:\n- Acknowledge: \"Some argue that...\"\n- Rebut: \"However...\"\n- Link: \"This is why...\"\n- Bonus: \"taught, not broken\" -- a one-line punch)",
    NOTES_IDO,
    FOOTER
  );

  // SLIDE 8 -- Conclusion structure
  contentSlide(
    pres,
    "Structure",
    C.ACCENT,
    "Conclusion -- Three Jobs",
    [
      "RESTATE:        your position, in NEW words (not copied from your intro)",
      "SUMMARISE:  your three reasons -- briefly, in one sentence",
      "FINAL LINE:    leave the reader with a strong feeling (question, repetition, or one-line punch)",
    ],
    NOTES_CONCLUSION,
    FOOTER
  );

  // SLIDE 9 -- We Do
  {
    const s = pres.addSlide();
    addTopBar(s, C.SECONDARY);
    addBadge(s, "We Do", { color: C.SECONDARY });
    addTitle(s, "Together: Counter for the OPPOSITE Side");

    const cardY = CONTENT_TOP;
    const cardH = 2.0;
    addCard(s, 0.5, cardY, 9, cardH, { strip: C.SECONDARY, fill: C.WHITE });
    s.addText("Position (the other side):  Camp Green Lake should STAY OPEN", {
      x: 0.75, y: cardY + 0.10, w: 8.5, h: 0.30,
      fontSize: 14, fontFace: FONT_B, color: C.SECONDARY, bold: true, margin: 0,
    });
    s.addText("Build the counter together on the board:\n- ACKNOWLEDGE: \"Some argue that the camp is dangerous.\"\n- REBUT:                  \"However, life teaches us that overcoming difficulty builds resilience.\"\n- LINK:                       \"This is why Camp Green Lake should stay open.\"\n\nThen draft a conclusion ORALLY -- three jobs (restate, summarise, final line).", {
      x: 0.75, y: cardY + 0.45, w: 8.5, h: cardH - 0.55,
      fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    const tipY = cardY + cardH + 0.18;
    const tipH = SAFE_BOTTOM - tipY;
    addCard(s, 0.5, tipY, 9, tipH, { strip: C.ACCENT, fill: C.BG_CARD });
    s.addText("Note: we are practising the structure. You pick YOUR side for the You Do.", {
      x: 0.75, y: tipY + 0.10, w: 8.5, h: tipH - 0.20,
      fontSize: 13, fontFace: FONT_B, color: C.PRIMARY, bold: true, valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_WEDO);
  }

  // SLIDE 10 + 11 -- CFU
  function buildCfuBase() {
    const slide = pres.addSlide();
    addTopBar(slide, C.ALERT);
    addBadge(slide, "CFU", { color: C.ALERT });
    addTitle(slide, "Which Counter Has ALL THREE Moves?", { color: C.ALERT });

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
    slide.addText("\"Some people say that students should have homework. They are probably right. Homework can be useful for some students who need extra practice. We should all listen to both sides.\"", {
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
    slide.addText("\"Some argue that homework helps students practise key skills. However, research shows that primary students learn more from outdoor play, reading and family time than from worksheets. This is why homework should be limited in primary school.\"", {
      x: 1.2, y: cardBY + 0.12, w: 8.0, h: cardBH - 0.24,
      fontSize: 12, fontFace: FONT_H, color: C.CHARCOAL, italic: true, valign: "middle", margin: 0,
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
      slide.addText("Stronger counter: B  --  ACKNOWLEDGE + REBUT (\"However...\") + LINK (\"This is why...\")", {
        x: 0.5, y: revealY, w: 9, h: revealH,
        fontSize: 13, fontFace: FONT_B, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      slide.addNotes(NOTES_CFU_REVEAL);
    }
  );

  // SLIDE 12 -- Optional Re-teach
  modellingSlide(
    pres,
    "Optional Re-teach",
    "Three-Sentence Frame for the Counter",
    "FRAME:\n\n\"Some argue that ____.\nHowever, ____.\nThis is why ____.\"\n\nThree sentences. Three jobs.",
    "FILLED EXAMPLE:\n\n\"Some argue that homework helps students learn.\nHowever, research shows that primary students learn more from outdoor play and family time.\nThis is why homework should be limited.\"\n\nYour turn: fill the frame on your board for YOUR topic.",
    NOTES_RETEACH,
    FOOTER
  );

  // SLIDE 13 -- You Do
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "You Do", { color: C.PRIMARY, w: 1.5 });
    addTitle(s, "Write Your Counter-Argument + Conclusion");

    addCard(s, 0.5, CONTENT_TOP, 9, 1.85, { strip: C.PRIMARY, fill: C.WHITE });
    s.addText("Two parts. Take 20 minutes total.", {
      x: 0.75, y: CONTENT_TOP + 0.10, w: 8.4, h: 0.30,
      fontSize: 14, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("PART 1 -- Counter-argument paragraph (3 to 4 sentences):\n  - Acknowledge the other side\n  - Rebut with a \"However\" sentence\n  - Link back to your position\n\nPART 2 -- Conclusion (3 to 5 sentences):\n  - Restate your position in NEW words\n  - Summarise your three reasons in one sentence\n  - End on a strong final line", {
      x: 0.75, y: CONTENT_TOP + 0.42, w: 8.4, h: 1.35,
      fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    const tipY = CONTENT_TOP + 2.00;
    const tipH = SAFE_BOTTOM - tipY;
    addCard(s, 0.5, tipY, 9, tipH, { strip: C.ACCENT, fill: C.BG_CARD });
    s.addText("Writing Time: 20 minutes  |  Mentor sheet on your desk for frames", {
      x: 0.75, y: tipY + 0.10, w: 8.5, h: 0.30,
      fontSize: 13, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("- Use \"Some argue that...\" / \"However...\" / \"This is why...\" for the counter\n- Use NEW words for the conclusion -- do not copy your introduction\n- Read both aloud quietly when you finish", {
      x: 0.75, y: tipY + 0.42, w: 8.4, h: tipH - 0.55,
      fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  }

  // SLIDE 14 -- Exit Ticket
  contentSlide(
    pres,
    "Exit Ticket",
    C.ACCENT,
    "Counter-Argument for: Homework Should Be Banned in Primary",
    [
      "Write ONE counter-argument paragraph -- THREE sentences",
      "Use all three moves: ACKNOWLEDGE, REBUT, LINK",
      "Use the frame if you need: \"Some argue that... However... This is why...\"",
      "3 minutes -- drop on my desk",
    ],
    NOTES_EXIT,
    FOOTER
  );

  // SLIDE 15 -- Closing (end of unit)
  closingSlide(
    pres,
    {
      reflectionPrompt: "Partner share: what is the BEST sentence in your persuasive piece?",
      scItems: [
        "I can acknowledge the other side in one clear sentence",
        "I can write a counter (acknowledge + rebut + link) and a conclusion that restates my position",
        "I can end my conclusion on a strong final line (question, repetition, or one-line punch)",
      ],
      selfAssessment: {
        prompt: "Show on your fingers: 1 (need help) to 5 (could teach it)",
        options: ["1-2", "3", "4-5"],
      },
    },
    NOTES_CLOSING
  );

  // ---- PDF: Counter + Conclusion Plan ---------------------------------------
  const plan = createPdf({ title: PLAN_RESOURCE.name });
  let planY = addPdfHeader(plan, "Counter-Argument + Conclusion Plan", {
    color: C.PRIMARY,
    subtitle: "Two writing tasks: counter (3-4 sentences) + conclusion (3-5 sentences)",
    lessonInfo: "Session 3 | Week 2 | Year 5/6 Literacy | Holes",
    showNameDate: true,
  });

  planY = addTipBox(plan, "Counter = acknowledge the other side, then rebut. Conclusion = restate position + summarise + strong final line.", planY, { color: C.PRIMARY });

  planY = addSectionHeading(plan, "PART 1 -- Counter-Argument", planY, { color: C.PRIMARY });
  planY = addBodyText(plan, "What does the OTHER side argue?", planY, { fontSize: 10, italic: true });
  planY = addLinedArea(plan, planY, 2, { lineSpacing: 20 });
  planY += 4;

  planY = addSectionHeading(plan, "ACKNOWLEDGE sentence", planY, { color: C.PRIMARY });
  planY = addBodyText(plan, "Start with: \"Some argue that...\" or \"It might be said that...\"", planY, { fontSize: 10, italic: true });
  planY = addLinedArea(plan, planY, 2, { lineSpacing: 20 });
  planY += 4;

  planY = addSectionHeading(plan, "REBUT sentence", planY, { color: C.SECONDARY });
  planY = addBodyText(plan, "Start with: \"However,...\" -- explain why your side is STILL stronger", planY, { fontSize: 10, italic: true });
  planY = addLinedArea(plan, planY, 2, { lineSpacing: 20 });
  planY += 4;

  planY = addSectionHeading(plan, "LINK sentence", planY, { color: C.ACCENT });
  planY = addBodyText(plan, "Start with: \"This is why...\" -- tie back to your overall position", planY, { fontSize: 10, italic: true });
  planY = addLinedArea(plan, planY, 2, { lineSpacing: 20 });

  addPdfFooter(plan, "Session 3 | Counter Plan -- Page 1");

  plan.addPage();
  let planY2 = addPdfHeader(plan, "PART 2 -- Conclusion Plan", {
    color: C.PRIMARY,
    subtitle: "Three jobs: restate position, summarise reasons, strong final line",
    lessonInfo: "Session 3 | Week 2 | Year 5/6 Literacy",
    showNameDate: false,
  });

  planY2 = addSectionHeading(plan, "RESTATE -- your position in NEW words", planY2, { color: C.PRIMARY });
  planY2 = addBodyText(plan, "Start with: \"In conclusion,...\" / \"It is clear that...\" / \"For all these reasons,...\"", planY2, { fontSize: 10, italic: true });
  planY2 = addLinedArea(plan, planY2, 2, { lineSpacing: 20 });
  planY2 += 4;

  planY2 = addSectionHeading(plan, "SUMMARISE -- your three reasons (one sentence)", planY2, { color: C.SECONDARY });
  planY2 = addBodyText(plan, "Use the rule of three: \"...because it is ___, ___ and ___.\"", planY2, { fontSize: 10, italic: true });
  planY2 = addLinedArea(plan, planY2, 2, { lineSpacing: 20 });
  planY2 += 4;

  planY2 = addSectionHeading(plan, "STRONG FINAL LINE -- the last thing the reader sees", planY2, { color: C.ACCENT });
  planY2 = addBodyText(plan, "Try: a rhetorical question / a call to action / a one-line punch", planY2, { fontSize: 10, italic: true });
  planY2 = addLinedArea(plan, planY2, 2, { lineSpacing: 20 });

  addPdfFooter(plan, "Session 3 | Conclusion Plan -- Page 2");

  plan.addPage();
  let planY3 = addPdfHeader(plan, "Counter + Conclusion -- Write Here", {
    color: C.PRIMARY,
    subtitle: "Write your two paragraphs back-to-back",
    lessonInfo: "Session 3 | Week 2 | Year 5/6 Literacy",
    showNameDate: false,
  });

  planY3 = addTipBox(plan, "Counter first, then conclusion. Use frames from the plan. Read both aloud quietly when you finish.", planY3, { color: C.SECONDARY });

  planY3 = addLinedArea(plan, planY3, 16, { lineSpacing: 22 });

  addPdfFooter(plan, "Session 3 | Counter + Conclusion -- Page 3");

  plan.addPage();
  let planY4 = addPdfHeader(plan, "Exit Ticket", {
    color: C.ACCENT,
    subtitle: "Counter for: Homework should be banned in primary school",
    lessonInfo: "Session 3 | Week 2",
    showNameDate: false,
  });

  planY4 = addTipBox(plan, "Write ONE counter-argument paragraph -- THREE sentences. Acknowledge, rebut, link. You can use \"Some argue that... However... This is why...\"", planY4, { color: C.ACCENT });

  planY4 = addLinedArea(plan, planY4, 6, { lineSpacing: 22 });

  addPdfFooter(plan, "Session 3 | Exit Ticket");

  // ---- PDF: Mentor Counter + Conclusion -------------------------------------
  const mp = createPdf({ title: MENTOR_RESOURCE.name });
  let mpY = addPdfHeader(mp, "Mentor Counter + Conclusion -- Annotated", {
    color: C.PRIMARY,
    subtitle: "Camp Green Lake should be shut down -- model counter + conclusion",
    lessonInfo: "Session 3 | Week 2 | Year 5/6 Literacy",
    showNameDate: false,
  });

  mpY = addTipBox(mp, "This model shows the counter-argument and conclusion. Use the PATTERN. Your topic, position and final line are yours.", mpY, { color: C.PRIMARY });

  mpY = addSectionHeading(mp, "Model counter-argument", mpY, { color: C.PRIMARY });
  mpY = addBodyText(mp, "Some argue that Camp Green Lake teaches the boys responsibility through hard work and routine. However, digging meaningless holes for the Warden is not real responsibility -- it is forced labour, with no education and no second chance. True responsibility comes from being taught, not from being broken. This is why Camp Green Lake must be shut down.", mpY, { fontSize: 12 });
  mpY += 12;

  mpY = addSectionHeading(mp, "Annotations -- counter", mpY, { color: C.PRIMARY });
  mpY = addBodyText(mp, "ACKNOWLEDGE: \"Some argue that Camp Green Lake teaches the boys responsibility...\" -- respectful framing of the other side.", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "REBUT: \"However, digging meaningless holes for the Warden is not real responsibility -- it is forced labour, with no education and no second chance.\" -- the \"However\" word + explicit rebuttal.", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "PUNCH: \"True responsibility comes from being taught, not from being broken.\" -- a one-line phrase that stays in the reader's head.", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "LINK: \"This is why Camp Green Lake must be shut down.\" -- loops back to position.", mpY, { fontSize: 10 });
  mpY += 10;

  mpY = addSectionHeading(mp, "Model conclusion", mpY, { color: C.PRIMARY });
  mpY = addBodyText(mp, "In conclusion, Camp Green Lake is unsafe, inhumane and unfit to call itself a place of learning. The heat puts the boys' bodies at risk, the work breaks their spirit, and the camp itself teaches nothing about responsibility or recovery. For Stanley, for Zero, and for every boy who deserves a real future -- Camp Green Lake must close.", mpY, { fontSize: 12 });
  mpY += 12;

  mpY = addSectionHeading(mp, "Annotations -- conclusion", mpY, { color: C.SECONDARY });
  mpY = addBodyText(mp, "RESTATE: \"In conclusion, Camp Green Lake is unsafe, inhumane and unfit...\" -- restated position with rule of three.", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "SUMMARISE: \"The heat puts the boys' bodies at risk, the work breaks their spirit, and the camp itself teaches nothing...\" -- compressed summary of the three reasons.", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "FINAL LINE: \"For Stanley, for Zero, and for every boy who deserves a real future -- Camp Green Lake must close.\" -- names two characters from the novel + rule of three + call to action.", mpY, { fontSize: 10 });
  mpY += 10;

  mpY = addSectionHeading(mp, "Sentence frames you can borrow", mpY, { color: C.PRIMARY });
  mpY = addBodyText(mp, "COUNTER: \"Some argue that ___. However, ___. This is why ___.\"", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "RESTATE: \"In conclusion, ___ is ___, ___ and ___.\"", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "SUMMARISE: \"___ does ___, ___ does ___, and ___ does ___.\"", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "FINAL LINE: \"For ___, for ___, and for every ___ -- ___ must ___.\"", mpY, { fontSize: 10 });

  addPdfFooter(mp, "Session 3 | Mentor Counter + Conclusion -- REFERENCE");

  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Holes_Persuasive_S3.pptx` }),
    writePdf(plan, PLAN_PDF_PATH),
    writePdf(mp, MENTOR_PDF_PATH),
  ]);

  console.log("PPTX written to " + `${OUT_DIR}/Holes_Persuasive_S3.pptx`);
  console.log("Done: " + PLAN_RESOURCE.name);
  console.log("Done: " + MENTOR_RESOURCE.name);
}

build().catch(console.error);
