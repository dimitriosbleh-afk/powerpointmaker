"use strict";

// Holes Unit - Lesson 2 (Week 1, Session 2): Narrative Character -- Show, Don't Tell
// Year 5/6 Literacy
// Focus: revealing character through action, dialogue and small gesture
// Anchor: Holes by Louis Sachar, Chapters 5-7 (Stanley arrives at Camp Green Lake)

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
const FOOTER = "Narrative Character | Lesson 2 | Week 1 | Year 5/6 Literacy | Holes";
const OUT_DIR = "output/Holes_Lesson2_Character_Stanley";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const PLAN_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Character Plan Show Don't Tell",
  "Student template: plan a character through action and gesture, not labels."
);
const MENTOR_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Mentor Character Paragraph",
  "Annotated model character paragraph -- shows nervous, brave or kind without saying the word."
);
const RESOURCE_ITEMS = [PLAN_RESOURCE, MENTOR_RESOURCE];
const PLAN_PDF_PATH = path.join(OUT_DIR, PLAN_RESOURCE.fileName);
const MENTOR_PDF_PATH = path.join(OUT_DIR, MENTOR_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Welcome back to Holes
- Yesterday: setting -- the PLACE
- Today: CHARACTER -- the people inside the place

DO:
- Display title slide as students settle
- Read 2-3 of yesterday's exit ticket hooks aloud (no names) to celebrate
- Have copies of Holes ready for the Chapter 5 to 7 read-aloud

TEACHER NOTES:
This is Session 2 of the unit. The anchor today is Stanley arriving at Camp Green Lake and digging his first hole. Read a SHORT extract from Chapter 5 (Stanley meeting Mr. Sir) or Chapter 7 (Stanley digging his first hole) -- pick the one your class will engage with. You do not need to read all of both.

WATCH FOR:
- Students who remember "barren" and "shimmer" from yesterday -- briefly celebrate; this is unit-wide vocabulary

[Literacy: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- Two resources again today
- The ${PLAN_RESOURCE.name} -- your character plan and writing page
- The ${MENTOR_RESOURCE.name} -- the annotated model paragraph

DO:
- Print the plan (one per student)
- Print the mentor (one per student or one per pair)
- Have the novel ready for the Ch 5-7 read-aloud
- Optional: keep yesterday's setting drafts handy -- some students will combine setting + character later in Week 3

[Literacy: Resources | VTLM 2.0: Student Resources]`;

const NOTES_HOOK = `SAY:
- Eyes on me. I will read a SHORT bit from Chapter 5 or 7 -- when Stanley arrives at Camp Green Lake
- Your job: notice what STANLEY DOES with his body. His hands. His feet. His eyes
- After I read: turn to your partner. What did Stanley DO?

DO:
- Read Holes Chapter 5 (Stanley meeting Mr. Sir) OR Chapter 7 (digging his first hole) -- pick one
- A short extract -- 1 to 2 pages is plenty
- 60 seconds partner share: "What did Stanley DO with his body?"
- Cold call 2-3 students

TEACHER NOTES:
The point of this launch is to plant the show-don't-tell idea before naming it. Students will notice Stanley looks down, swallows, holds still, grips the shovel. Capture 2-3 observations on the board. Then transition to the LI/SC.

WATCH FOR:
- Students who say "Stanley was scared" -- gently redirect: "How do you KNOW? What did he DO?"
- Students who name a specific action -- celebrate: "Yes. The author SHOWED him being scared by what he did"

[Literacy: Text Launch | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_LI_SC = `SAY:
- Read the LI with me
- We are learning to SHOW a character through what they DO, not just by labelling them
- Three "I can" statements

DO:
- Choral read the LI and SCs
- Brief: "Show, don't tell. We do not say 'Stanley was scared'. We SHOW him being scared"

TEACHER NOTES:
SC1 -- one action -- is achievable for all. SC2 (target) is two actions + one small gesture. SC3 stretches confident writers to use action AND a short dialogue line. The exit ticket targets SC2.

WATCH FOR:
- Students who reach for adjectives ("scared", "brave") -- redirect: "What does scared LOOK like?"

[Literacy: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_VOCAB_FLINCH = `SAY:
- One key word today: FLINCH. A verb
- It means to move suddenly because you are scared, hurt or surprised
- A small, fast movement
- Try the gesture: pretend a wasp just flew past your face -- your shoulders jump up

DO:
- Choral say FLINCH
- Quick gesture: shoulders up, eyes shut briefly
- Ask: "When might you flinch?" -- 1 partner share, 30 seconds
- Take 2 examples

TEACHER NOTES:
FLINCH is exactly the kind of small action that SHOWS fear without naming it. Students will use it today in the You Do.

WATCH FOR:
- Students offering "when a balloon pops" or "when a teacher shouts" -- celebrate
- Students confused -- demonstrate again with the wasp gesture

[Literacy: Vocabulary | VTLM 2.0: Vocabulary Instruction]`;

const NOTES_IDO = `SAY:
- Watch me write a character paragraph that SHOWS a feeling without naming it
- I am going to write about a boy on his first day at a new school
- I will NOT use the word "nervous" -- but you will FEEL it
- Watch my CHOICES: actions, eyes, hands, feet

DO:
- Display the I Do slide
- Read the paragraph aloud
- After reading, point to each underlined / highlighted action
- Ask: "What FEELING did you get from this boy? Nervous? Scared? Lonely?"
- Confirm: "I never wrote the word nervous. But you felt it"

TEACHER NOTES:
The boy at the new school is deliberately NOT Stanley -- this keeps the I Do model close enough that the skill transfers, far enough that students cannot copy. The character method works for any feeling: nervous, brave, angry, kind. Show students how each action carries one piece of the feeling.

WATCH FOR:
- Students who name the feeling correctly without seeing the word -- celebrate; this proves the method works

[Literacy: I Do | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_WEDO = `SAY:
- Together now. We are going to plan a character paragraph about a CHARACTER who is BRAVE -- but we will not use the word brave
- Pick a scene: a girl walking up to the principal's office to report a problem
- I will collect your SHOWING ideas on the board

DO:
- Display the We Do slide
- Cold call students for actions, eyes, hands, voice
- Capture 4-6 strong showing details on the board
- Build the FIRST sentence together: a small concrete action, not a label
- Aim: full first sentence + 2-3 details written on the board

TEACHER NOTES:
The brave girl is different content from the I Do (nervous boy) and from the You Do (3 character options later). Keep the class build to 5-6 minutes. If students say "she was brave" or "she felt brave" -- redirect to action: "What did her HANDS do? Her FEET? Her VOICE?"

WATCH FOR:
- Students who name an action ("she squared her shoulders") -- celebrate
- Students who reach for "she was brave" -- redirect to a body word

[Literacy: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_CFU_BUILD = `SAY:
- Quick check. Two sentences trying to show that Stanley is anxious. Which one is SHOWING, and which is TELLING?
- A: "Stanley felt very nervous when Mr. Sir spoke to him."
- B: "Stanley swallowed and stared at the toes of his shoes."
- Show me A or B on your fingers

DO:
- Display both sentences
- Show Me Fingers (1 for A, 2 for B)
- Scan: most students should choose B
- Cold call 1-2 students: "Why B?"

TEACHER NOTES:
A is TELLING -- it labels the feeling. B is SHOWING -- it gives an action that means nervous. B is the stronger sentence. If many chose A, pivot to the re-teach slide.

WATCH FOR:
- Students who choose B and articulate why -- ready for You Do
- Students who choose A because it has "feel" in it -- redirect: "Showing means we see the BODY doing something"

[Literacy: CFU | VTLM 2.0: Formative Assessment]`;

const NOTES_CFU_REVEAL = `SAY:
- Stronger SHOWING sentence: B
- A told us Stanley was nervous. B made us SEE it
- Swallowing + staring at his shoes = nervous, without ever saying the word

DO:
- Display the reveal banner
- Read B aloud
- Briefly: "B trusts the reader. Show, don't tell"

CFU CHECKPOINT:
Technique: Show Me Fingers (1 for A, 2 for B)
Script:
- "Hold up 1 for A, 2 for B. Which is SHOWING?"
- Scan for: most students choose B (>=80%)
PROCEED (>=80%): Most chose B. Release to You Do.
PIVOT (<80%): Use the optional re-teach slide that follows -- it lifts ONE telling label into a showing line in front of the class.

WATCH FOR:
- Students who self-correct toward B -- celebrate

[Literacy: CFU Reveal | VTLM 2.0: Formative Assessment]`;

const NOTES_RETEACH = `SAY:
- Second look at SHOWING
- Watch -- I will take ONE telling label and turn it into a SHOWING line
- Telling: "She was angry."
- Now I think: what does angry LOOK like? In the HANDS? In the FACE? In the VOICE?
- Hands: "She clenched her fists."
- Face: "Her jaw locked."
- Voice: "She spoke softly. Too softly."
- Pick one. Or use all three. None of them say "angry"

DO:
- Display the re-teach slide
- Work the example LIVE -- write each option on the board
- Re-check: ask students to give ONE showing line for "She was sad" on their mini-whiteboards
- Take 3 board responses

TEACHER NOTES:
OPTIONAL slide. Use only if CFU showed fewer than about 80% understanding. Different approach: instead of comparing two sentences, take ONE telling label and TRANSFORM it. This makes the move from telling to showing more procedural and visible.

WATCH FOR:
- Students who write a body action for sad ("she looked at the floor", "she didn't lift her bag") -- ready
- Students who write "she was crying" -- accept; that IS showing

[Literacy: Optional Re-teach | VTLM 2.0: Reteach with Different Approach]`;

const NOTES_YOUDO = `SAY:
- Your turn. Write a short character paragraph
- Pick ONE feeling to SHOW (do not say the word):
  - excited
  - lonely
  - guilty
  - proud
  - frustrated
- Pick a place to set it (school, home, sports field, shop)
- Use the Character Plan
- 15 minutes. I will circulate

DO:
- Display the You Do slide
- Distribute the plan and mentor
- Circulate -- prioritise students who needed the re-teach
- Quick conferences: "Read me your first action. Does it show the feeling?"

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the mentor sheet. Find the highlighted actions and copy the PATTERN. Pick one feeling from the list. Use the sentence frames on the back of the mentor sheet
- Extra Notes: These students still hit SC1 and SC2 -- the frame supports sentence construction

EXTENDING PROMPT:
- Task: After your character paragraph, add ONE short line of dialogue that fits the feeling. Use a comma and quotation marks
- Extra Notes: Push for short, real dialogue: "I'm fine," she said. "Really. I'm fine"

TEACHER NOTES:
The five feelings keep the SKILL identical but the content varied. Active circulation is the formative assessment. Aim for 4-6 sentences. The dialogue extension targets SC3.

WATCH FOR:
- Students who write "He was excited" -- redirect: "Cross that out. What did he DO with his body?"
- Students who write a strong showing line -- celebrate publicly

[Literacy: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket time
- On your plan sheet, write TWO short SHOWING sentences for this feeling: SURPRISED
- The character has just opened a birthday present and it is the dog they have always wanted
- Do NOT use the word surprised, happy, or excited
- 2 minutes. Drop it on my desk

DO:
- Display the exit ticket slide
- 2 minutes silent
- Collect on the way out
- Read 3-5 aloud at the start of Lesson 3

TEACHER NOTES:
Exit ticket targets SC2 (two actions). Two sentences is enough.

WATCH FOR:
- Students who write "He was surprised" -- still collect; flag for next lesson
- Students who write strong body actions -- celebrate next session

[Literacy: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Self-check time
- Look at the three success criteria. Show on your fingers: 1 (need help) to 5 (could teach it)
- Partner share: what is ONE action you used in your paragraph that you are proud of?

DO:
- Run fingers check for each SC
- 30 seconds partner share
- Briefly: "Tomorrow we put setting AND character together with TENSION -- a tight, exciting short narrative"

TEACHER NOTES:
Tomorrow is the last narrative lesson before persuasive begins.

WATCH FOR:
- Students who name a precise body action -- celebrate
- Students stuck at 1-2 fingers -- private check-in

[Literacy: Closing | VTLM 2.0: Review and Reflect]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Holes -- Lesson 2 -- Character (Show, Don't Tell)";

  // SLIDE 1 -- Title
  titleSlide(
    pres,
    "Character: Show, Don't Tell",
    "Stanley Arrives at Camp Green Lake -- Holes",
    "Lesson 2  |  Week 1  |  Year 5/6 Literacy",
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

  // SLIDE 3 -- Hook / Text Launch (Ch 5-7 read-aloud)
  quoteSlide(
    pres,
    "Read Aloud",
    "Holes -- Chapter 5 or 7 (teacher choice)",
    "Listen for what STANLEY DOES with his body. His hands. His feet. His eyes.",
    "Stanley arriving at Camp Green Lake / first dig",
    "Partner talk: what did Stanley DO with his body? Name TWO actions you noticed.",
    NOTES_HOOK,
    FOOTER
  );

  // SLIDE 4 -- LI / SC
  liSlide(
    pres,
    [
      "We are learning to SHOW what a character is feeling through their actions, not by labelling them with a feeling word",
    ],
    [
      "I can write a character action that shows a feeling (no feeling word used)",
      "I can use two body actions plus one small gesture",
      "I can add a short line of dialogue that fits the character",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // SLIDE 5 -- Vocab
  vocabSlide(
    pres,
    "flinch",
    "verb",
    "To move suddenly because you are scared, hurt or surprised. A small, fast movement -- often the shoulders or eyes.",
    "When Mr. Sir slammed the cooler shut, Stanley flinched.",
    NOTES_VOCAB_FLINCH,
    FOOTER
  );

  // SLIDE 6 -- I Do
  modellingSlide(
    pres,
    "I Do -- Watch Me",
    "Model: Showing a Nervous Boy (the word 'nervous' is never used)",
    "Setup:\n\nA boy on his first day at a new school. Standing at the classroom door.\n\nMy goal:\n- Use 3 to 4 BODY actions\n- Add 1 small gesture (hands, eyes)\n- NEVER write the word 'nervous'\n- The reader should still feel it",
    "He stopped in the doorway. The class turned to look. He tightened his grip on the strap of his bag and stared at the toes of his shoes. His knee jumped, just once. \"Hi,\" he said. The word came out smaller than he wanted.\n\n(My choices:\n- action: stopped, tightened grip, stared\n- small gesture: knee jumped\n- dialogue: short, quiet\n- never said 'nervous')",
    NOTES_IDO,
    FOOTER
  );

  // SLIDE 7 -- We Do
  {
    const s = pres.addSlide();
    addTopBar(s, C.SECONDARY);
    addBadge(s, "We Do", { color: C.SECONDARY });
    addTitle(s, "Together: Brave Girl Walking to the Principal's Office");

    const cardY = CONTENT_TOP;
    const cardH = 2.0;
    addCard(s, 0.5, cardY, 9, cardH, { strip: C.SECONDARY, fill: C.WHITE });
    s.addText("Brainstorm SHOWING ideas -- never say the word 'brave'", {
      x: 0.75, y: cardY + 0.10, w: 8.5, h: 0.30,
      fontSize: 14, fontFace: FONT_B, color: C.SECONDARY, bold: true, margin: 0,
    });
    s.addText("- HANDS:    what do her hands do? (clench? open? hold something?)\n- FEET:      slow steps? steady steps? quick steps?\n- EYES:      where does she look? (straight ahead? at the floor? up?)\n- VOICE:    quiet? steady? cracking?\n- DETAIL:  one small specific thing (a folder under her arm? a deep breath?)", {
      x: 0.75, y: cardY + 0.45, w: 8.5, h: cardH - 0.55,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    const tipY = cardY + cardH + 0.18;
    const tipH = SAFE_BOTTOM - tipY;
    addCard(s, 0.5, tipY, 9, tipH, { strip: C.ACCENT, fill: C.BG_CARD });
    s.addText("On the board together: build the FIRST sentence", {
      x: 0.75, y: tipY + 0.10, w: 8.5, h: 0.30,
      fontSize: 13, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("- One body action only\n- No feeling labels\n- Example to try: \"She squared her shoulders and walked.\"", {
      x: 0.75, y: tipY + 0.42, w: 8.5, h: tipH - 0.55,
      fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_WEDO);
  }

  // SLIDE 8 + 9 -- CFU: Showing vs Telling (reveal pair)
  function buildCfuBase() {
    const slide = pres.addSlide();
    addTopBar(slide, C.ALERT);
    addBadge(slide, "CFU", { color: C.ALERT });
    addTitle(slide, "Showing or Telling?", { color: C.ALERT });

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
    const cardH = 1.20;
    addCard(slide, 0.5, cardY, 9, cardH, { strip: C.MUTED, fill: C.WHITE });
    slide.addText("A", {
      x: 0.7, y: cardY + 0.10, w: 0.5, h: 0.4,
      fontSize: 22, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0,
    });
    slide.addText("\"Stanley felt very nervous when Mr. Sir spoke to him.\"", {
      x: 1.2, y: cardY + 0.18, w: 8.0, h: cardH - 0.30,
      fontSize: 17, fontFace: FONT_H, color: C.CHARCOAL, italic: true, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    const cardBY = cardY + cardH + 0.18;
    const cardBH = SAFE_BOTTOM - cardBY - 0.55;
    addCard(slide, 0.5, cardBY, 9, cardBH, { strip: C.SECONDARY, fill: C.BG_CARD });
    slide.addText("B", {
      x: 0.7, y: cardBY + 0.10, w: 0.5, h: 0.4,
      fontSize: 22, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0,
    });
    slide.addText("\"Stanley swallowed and stared at the toes of his shoes.\"", {
      x: 1.2, y: cardBY + 0.14, w: 8.0, h: cardBH - 0.24,
      fontSize: 17, fontFace: FONT_H, color: C.CHARCOAL, italic: true, valign: "middle", margin: 0,
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
      slide.addText("Stronger sentence: B  --  it SHOWS nervous (swallow + stare at shoes), never names the feeling", {
        x: 0.5, y: revealY, w: 9, h: 0.42,
        fontSize: 13, fontFace: FONT_B, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      slide.addNotes(NOTES_CFU_REVEAL);
    }
  );

  // SLIDE 10 -- Optional Re-teach
  modellingSlide(
    pres,
    "Optional Re-teach",
    "Take ONE Telling Label and Turn It Into a SHOWING Line",
    "Telling label:\n\n\"She was angry.\"\n\nAsk yourself:\n- What do angry HANDS do?\n- What does an angry FACE do?\n- What does an angry VOICE do?",
    "Showing options:\n\nHANDS:   \"She clenched her fists.\"\nFACE:    \"Her jaw locked.\"\nVOICE:   \"She spoke softly. Too softly.\"\n\nPick one. Or use all three. None of them say angry.\n\nYour turn: write ONE showing line for \"She was sad\" on your board.",
    NOTES_RETEACH,
    FOOTER
  );

  // SLIDE 11 -- You Do
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "You Do", { color: C.PRIMARY, w: 1.5 });
    addTitle(s, "Write Your Own Character Paragraph");

    addCard(s, 0.5, CONTENT_TOP, 9, 1.85, { strip: C.PRIMARY, fill: C.WHITE });
    s.addText("Pick ONE feeling to SHOW (do not say the word):", {
      x: 0.75, y: CONTENT_TOP + 0.10, w: 8.4, h: 0.30,
      fontSize: 14, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("- excited\n- lonely\n- guilty\n- proud\n- frustrated", {
      x: 0.75, y: CONTENT_TOP + 0.45, w: 8.4, h: 1.30,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    const tipY = CONTENT_TOP + 2.00;
    const tipH = SAFE_BOTTOM - tipY;
    addCard(s, 0.5, tipY, 9, tipH, { strip: C.ACCENT, fill: C.BG_CARD });
    s.addText("Writing Time: 15 minutes", {
      x: 0.75, y: tipY + 0.10, w: 5, h: 0.30,
      fontSize: 13, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("First:    Use the Character Plan -- write your feeling at the top (only you see it)\nNext:    Brainstorm: hands, eyes, voice, one small gesture\nThen:    Write 4 to 6 sentences using those actions\nFinally: Read it to a partner -- can they guess the feeling?", {
      x: 0.75, y: tipY + 0.42, w: 8.4, h: tipH - 0.55,
      fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  }

  // SLIDE 12 -- Exit Ticket
  contentSlide(
    pres,
    "Exit Ticket",
    C.ACCENT,
    "Two Showing Sentences -- a Surprise Birthday Present",
    [
      "The character has just opened a birthday present -- it is the dog they always wanted",
      "Write TWO showing sentences",
      "Do NOT use 'surprised', 'happy' or 'excited'",
      "2 minutes -- drop on my desk",
    ],
    NOTES_EXIT,
    FOOTER
  );

  // SLIDE 13 -- Closing
  closingSlide(
    pres,
    {
      reflectionPrompt: "Partner share: what is ONE action you used that you are proud of?",
      scItems: [
        "I can write a character action that shows a feeling (no feeling word used)",
        "I can use two body actions plus one small gesture",
        "I can add a short line of dialogue that fits the character",
      ],
      selfAssessment: {
        prompt: "Show on your fingers: 1 (need help) to 5 (could teach it)",
        options: ["1-2", "3", "4-5"],
      },
    },
    NOTES_CLOSING
  );

  // ---- PDF: Character Plan ---------------------------------------------------
  const plan = createPdf({ title: PLAN_RESOURCE.name });
  let planY = addPdfHeader(plan, "Character Plan -- Show, Don't Tell", {
    color: C.PRIMARY,
    subtitle: "Plan a character through actions, gestures and a short line of dialogue",
    lessonInfo: "Lesson 2 | Week 1 | Year 5/6 Literacy | Holes",
    showNameDate: true,
  });

  planY = addTipBox(plan, "Today we SHOW a feeling without saying the word. Pick your feeling, then plan body actions, eyes, hands, voice and one small gesture. Write the feeling at the top -- but never inside your paragraph.", planY, { color: C.PRIMARY });

  planY = addSectionHeading(plan, "Step 1 -- Your secret feeling", planY, { color: C.PRIMARY });
  planY = addBodyText(plan, "Pick ONE -- circle: excited / lonely / guilty / proud / frustrated / your own: ____________", planY, { fontSize: 11 });
  planY += 4;

  planY = addSectionHeading(plan, "Step 2 -- Setting", planY, { color: C.PRIMARY });
  planY = addBodyText(plan, "Where is your character? (school, home, sports field, shop, somewhere else)", planY, { fontSize: 10, italic: true });
  planY = addLinedArea(plan, planY, 1, { lineSpacing: 20 });
  planY += 4;

  planY = addSectionHeading(plan, "Step 3 -- Brainstorm SHOWING ideas (no feeling words)", planY, { color: C.PRIMARY });
  planY = addBodyText(plan, "HANDS:", planY, { fontSize: 11 });
  planY = addLinedArea(plan, planY, 2, { lineSpacing: 18 });
  planY = addBodyText(plan, "EYES:", planY, { fontSize: 11 });
  planY = addLinedArea(plan, planY, 2, { lineSpacing: 18 });
  planY = addBodyText(plan, "VOICE / WORDS:", planY, { fontSize: 11 });
  planY = addLinedArea(plan, planY, 2, { lineSpacing: 18 });
  planY = addBodyText(plan, "ONE SMALL GESTURE (a knee jump, a half-smile, a tap):", planY, { fontSize: 11 });
  planY = addLinedArea(plan, planY, 1, { lineSpacing: 18 });

  addPdfFooter(plan, "Lesson 2 | Character Plan -- Page 1");

  plan.addPage();
  let planY2 = addPdfHeader(plan, "Character Paragraph -- Write Here", {
    color: C.PRIMARY,
    subtitle: "4 to 6 sentences. No feeling words.",
    lessonInfo: "Lesson 2 | Week 1 | Year 5/6 Literacy",
    showNameDate: false,
  });

  planY2 = addTipBox(plan, "Use your plan. Open with one body action. Build up details. Add one short line of dialogue if you can. Read it back -- can your partner guess the feeling?", planY2, { color: C.SECONDARY });

  planY2 = addLinedArea(plan, planY2, 14, { lineSpacing: 22 });

  addPdfFooter(plan, "Lesson 2 | Character Paragraph -- Page 2");

  plan.addPage();
  let planY3 = addPdfHeader(plan, "Exit Ticket", {
    color: C.ACCENT,
    subtitle: "Two SHOWING sentences -- surprised by a birthday dog",
    lessonInfo: "Lesson 2 | Week 1",
    showNameDate: false,
  });

  planY3 = addTipBox(plan, "Write TWO showing sentences. The character has just opened a birthday present and it is the dog they always wanted. Do NOT use surprised, happy or excited.", planY3, { color: C.ACCENT });

  planY3 = addLinedArea(plan, planY3, 4, { lineSpacing: 22 });

  addPdfFooter(plan, "Lesson 2 | Exit Ticket");

  // ---- PDF: Mentor Character Paragraph --------------------------------------
  const mp = createPdf({ title: MENTOR_RESOURCE.name });
  let mpY = addPdfHeader(mp, "Mentor Character Paragraph -- Annotated", {
    color: C.PRIMARY,
    subtitle: "First day at a new school -- nervous boy (the word 'nervous' is never used)",
    lessonInfo: "Lesson 2 | Week 1 | Year 5/6 Literacy",
    showNameDate: false,
  });

  mpY = addTipBox(mp, "This model SHOWS a nervous boy through action, gesture and a short line of dialogue. The word 'nervous' never appears. Use the PATTERN, not the words. Your character is different.", mpY, { color: C.PRIMARY });

  mpY = addSectionHeading(mp, "Model paragraph", mpY, { color: C.PRIMARY });
  mpY = addBodyText(mp, "He stopped in the doorway. The class turned to look. He tightened his grip on the strap of his bag and stared at the toes of his shoes. His knee jumped, just once. \"Hi,\" he said. The word came out smaller than he wanted.", mpY, { fontSize: 12 });
  mpY += 12;

  mpY = addSectionHeading(mp, "Annotations -- body actions", mpY, { color: C.PRIMARY });
  mpY = addBodyText(mp, "\"He stopped in the doorway.\" -- the action shows hesitation.", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "\"tightened his grip on the strap\" -- hands. A clench gives away nerves.", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "\"stared at the toes of his shoes\" -- eyes down. He cannot meet the class.", mpY, { fontSize: 10 });
  mpY += 8;

  mpY = addSectionHeading(mp, "Annotations -- small gesture", mpY, { color: C.SECONDARY });
  mpY = addBodyText(mp, "\"His knee jumped, just once.\" -- one tiny, involuntary movement. The body tells us what the boy will not.", mpY, { fontSize: 10 });
  mpY += 8;

  mpY = addSectionHeading(mp, "Annotations -- dialogue", mpY, { color: C.ACCENT });
  mpY = addBodyText(mp, "\"'Hi,' he said. The word came out smaller than he wanted.\" -- short. Quiet. Even the SIZE of the word carries the feeling.", mpY, { fontSize: 10 });
  mpY += 8;

  mpY = addSectionHeading(mp, "Sentence frames you can borrow", mpY, { color: C.PRIMARY });
  mpY = addBodyText(mp, "[Character] [body action]. [Setting detail].", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "[Character] [hand action] and [eye action].", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "[One small gesture], [adverb only -- 'just once' / 'for a second' / 'a single time'].", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "\"[Short word],\" [character] said. [Comment on the way it sounded].", mpY, { fontSize: 10 });

  addPdfFooter(mp, "Lesson 2 | Mentor Character Paragraph -- REFERENCE");

  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Holes_Lesson2.pptx` }),
    writePdf(plan, PLAN_PDF_PATH),
    writePdf(mp, MENTOR_PDF_PATH),
  ]);

  console.log("PPTX written to " + `${OUT_DIR}/Holes_Lesson2.pptx`);
  console.log("Done: " + PLAN_RESOURCE.name);
  console.log("Done: " + MENTOR_RESOURCE.name);
}

build().catch(console.error);
