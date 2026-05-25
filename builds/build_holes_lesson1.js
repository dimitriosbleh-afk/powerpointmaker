"use strict";

// Holes Unit - Lesson 1 (Week 1, Session 1): Narrative Setting -- Camp Green Lake
// Year 5/6 Literacy
// Focus: vivid setting description (sensory detail + place as character)
// Anchor: Holes by Louis Sachar, Chapter 1 read-aloud

const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");

const { createTheme, weekToVariant } = require("../themes/factory");
// Unit cohesion: all Holes lessons use variant 0 (week 1 of the 3-week unit).
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

const SESSION_NUMBER = 1;
const FOOTER = "Narrative Setting | Lesson 1 | Week 1 | Year 5/6 Literacy | Holes";
const OUT_DIR = "output/Holes_Lesson1_Setting_Camp_Green_Lake";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const PLAN_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Setting Plan Camp Green Lake",
  "Student template: plan a vivid setting paragraph using the five senses and one short anchor sentence."
);
const MENTOR_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Mentor Setting Paragraph",
  "Annotated model setting paragraph based on Camp Green Lake -- shows sensory detail and short hook sentence."
);
const RESOURCE_ITEMS = [PLAN_RESOURCE, MENTOR_RESOURCE];
const PLAN_PDF_PATH = path.join(OUT_DIR, PLAN_RESOURCE.fileName);
const MENTOR_PDF_PATH = path.join(OUT_DIR, MENTOR_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Welcome to our new unit: Holes by Louis Sachar
- Three weeks of writing: narrative this week, persuasive next week, publish in week 3
- Today: how authors build a strong SETTING -- the place feels alive

DO:
- Display title slide as students enter
- Hold up the novel
- Have copies of Holes ready -- you will be reading Chapter 1 aloud

TEACHER NOTES:
This is Session 1 of an 11-lesson unit. Students have learned narrative and persuasive writing earlier in the year (about 4 months ago) -- expect rusty recall, not fresh mastery. The unit is anchored in chapter read-alouds (you read, students listen and respond) -- students are not expected to read the whole novel themselves.

WATCH FOR:
- Students excited by the novel cover -- briefly preview the unit shape (narrative, persuasive, publish) so they know what is coming

[Literacy: Title | VTLM 2.0: Establishing Purpose and Relevance]`;

const NOTES_RESOURCES = `SAY:
- Two resources today
- The ${PLAN_RESOURCE.name} -- your setting plan and writing page
- The ${MENTOR_RESOURCE.name} -- the annotated model paragraph

DO:
- Print the plan template (one per student)
- Print the mentor (one per student or one per pair)
- Have a copy of Holes for the Chapter 1 read-aloud
- Optional: a world map or photo of a Texas desert landscape for the launch

TEACHER NOTES:
Resources are deliberately lean this week. The teaching is in the read-aloud and the modelling, not in a thick worksheet pack.

[Literacy: Resources | VTLM 2.0: Student Resources]`;

const NOTES_HOOK = `SAY:
- Eyes on me. I am going to read the FIRST page of our novel aloud
- Your job: close your eyes if you want to, and PICTURE the place in your head
- After I read, I will ask: what did you SEE?

DO:
- Read Holes, Chapter 1 aloud (pages 3-5 in most editions, the first 3-4 paragraphs is plenty)
- Keep your voice steady -- the place is hot, dry and still
- After reading, pause for 5 seconds
- Then: "Turn to your partner. What did you SEE in your head?"
- 60 seconds partner share
- Cold call 2-3 students

TEACHER NOTES:
The opening of Holes is famously vivid. Stanley has not even been introduced -- the place is the hook. Students will name heat, dust, a dry lake bed, nothing growing. Celebrate strong sensory language. If your school has a different edition of Holes, you can still read Chapter 1 -- the opening is short and unchanged.

WATCH FOR:
- Students who name a SENSE (saw, heard, felt) -- excellent
- Students who describe a feeling (scary, lonely) -- excellent
- Students who name a thing from the page (cracked earth, dry lake) -- excellent

[Literacy: Text Launch | VTLM 2.0: Activating Prior Knowledge / Establishing Purpose]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me
- We are learning to write a setting that the reader can SEE
- Three success criteria. Read them with me

DO:
- Choral read the LI
- Choral read each "I can" statement
- Brief: "Notice -- this is short. One paragraph. A snapshot of a place"

TEACHER NOTES:
SC1 is achievable for all students -- two senses and a place name. SC2 is the target for most students -- adding a short hook sentence. SC3 stretches confident writers -- linking the place to a feeling or mood. The exit ticket targets SC2.

WATCH FOR:
- Students who think they need to write a whole story -- redirect: "Just the SETTING. Not what happens. Where it is"

[Literacy: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_VOCAB_BARREN = `SAY:
- Two quick vocabulary words for the unit
- First: BARREN. An adjective. It means empty -- nothing grows there
- Camp Green Lake is barren
- Say it together: BARREN

DO:
- Choral say BARREN
- Quick gesture: open both hands flat -- "nothing here"
- Ask: "Think of a place that is barren in real life" -- 1 partner share, 20 seconds
- Take 2 quick examples

TEACHER NOTES:
Barren is high-value for setting work. Students will use it in their own writing today. Pair it with the gesture so weaker readers have a hook.

WATCH FOR:
- Students offering "the desert" or "the moon" -- accurate; celebrate
- Students saying "the playground" -- redirect gently: "Barren means NOTHING grows. Is your playground really empty?"

[Literacy: Vocabulary | VTLM 2.0: Vocabulary Instruction]`;

const NOTES_VOCAB_SHIMMER = `SAY:
- Second word: SHIMMER. A verb. It means to shine with a wavy, moving light
- Heat can shimmer above hot ground
- Say it together: SHIMMER

DO:
- Choral say SHIMMER
- Quick gesture: wave your hand slowly side to side, fingers wiggling
- Ask: "Think of something that shimmers" -- 1 partner share, 20 seconds
- Take 2 quick examples

TEACHER NOTES:
Shimmer is a precise verb. Better than 'shine' for desert heat. Encourage students to use it when describing Camp Green Lake.

WATCH FOR:
- Students offering "the sea", "a fish", "a star" -- celebrate
- Students saying "a torch" -- redirect: "Does a torch wave? Or does it just shine?"

[Literacy: Vocabulary | VTLM 2.0: Vocabulary Instruction]`;

const NOTES_IDO = `SAY:
- Watch me write a setting paragraph about Camp Green Lake
- Watch my CHOICES, not just my words
- I am going to start with a SHORT HOOK sentence -- something that grabs the reader
- Then I will use at least TWO senses: see, hear, smell, touch, taste
- I will NOT tell a story. Just the place

DO:
- Display the I Do slide
- Talk through each sentence as you read it
- Highlight the short hook: "There is no lake at Camp Green Lake."
- Highlight the senses used (sight, touch, hearing)
- Show the place name appearing twice

TEACHER NOTES:
The hook sentence "There is no lake at Camp Green Lake" is Sachar's actual opening line. Use it as your model hook. Then build the rest of the paragraph as a teacher think-aloud, using your own sensory detail. The model demonstrates SC1 (two senses, place name), SC2 (short hook sentence), and SC3 (mood linked to place).

WATCH FOR:
- Students who notice the hook sentence is SHORT -- celebrate
- Students who try to memorise the sentences -- redirect: "Your hook will be different. Use a different short sentence about a place YOU choose"

[Literacy: I Do | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_WEDO = `SAY:
- Together now. We are going to plan a setting paragraph about an EMPTY playground at dusk
- Different place from Camp Green Lake -- same skills
- I will collect your sensory ideas on the board
- We will write the hook sentence together

DO:
- Display the We Do slide
- Use the five-senses chart on the slide as a class prompt
- Cold call students for each sense
- Write the strongest 2 or 3 sensory details on the board
- Build a hook sentence together from a student suggestion
- Aim: full hook + first sentence written on the board by end of We Do

TEACHER NOTES:
The empty playground is deliberately different content from Camp Green Lake -- We Do and You Do must use different examples (CLAUDE.md: We Do and You Do must use different content/examples). Keep the class build to 5-6 minutes -- students need writing time. If students hesitate on a sense, give two options ("Do you hear the wind, or the chain on the swing?").

WATCH FOR:
- Students who only give SIGHT details -- prompt: "What can you HEAR? What can you SMELL?"
- Students who suggest "tasty playground" -- redirect: "Taste is hard for places. Save taste for food. Try smell instead"

[Literacy: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_CFU_BUILD = `SAY:
- Quick check. Two HOOK sentences about a winter forest. Which one grabs the reader better?
- A: "The forest had a lot of trees and it was cold and there was snow on the ground."
- B: "The forest had stopped breathing."
- Show me A or B on your fingers

DO:
- Display both sentences
- Show Me Fingers (1 finger for A, 2 for B)
- Scan: most students should choose B
- Cold call 1-2 students: "Why B?"

TEACHER NOTES:
A is a long list of facts. B is short and uses personification -- it grabs the reader. B is the stronger hook. A strong hook is short, surprising and shows mood. If many students chose A, pivot to the re-teach slide.

WATCH FOR:
- Students who choose B and articulate why -- ready for You Do
- Students who choose A because it has "more information" -- redirect: "A hook is the first taste. Less is more"

[Literacy: CFU | VTLM 2.0: Formative Assessment]`;

const NOTES_CFU_REVEAL = `SAY:
- Stronger hook: B
- B is SHORT, surprising, and tells you the place feels still and silent
- A is just a list of facts
- The HOOK is the first taste of the place. Less is more

DO:
- Display the reveal banner
- Read B aloud with feeling: "The forest had stopped breathing."
- If many picked A, slow down: "A tells me information. B makes me feel something."

CFU CHECKPOINT:
Technique: Show Me Fingers (1 for A, 2 for B)
Script:
- "Hold up 1 finger for A, 2 fingers for B. Which is the stronger hook?"
- Scan for: most students choose B (>=80%)
PROCEED (>=80%): Most chose B. Release to You Do.
PIVOT (<80%): Most likely issue -- students think more detail means better hook. Use the optional re-teach slide that follows for a different angle.

TEACHER NOTES:
After reveal, release students to You Do. Time is the limit.

WATCH FOR:
- Students who self-correct toward B -- celebrate

[Literacy: CFU Reveal | VTLM 2.0: Formative Assessment]`;

const NOTES_RETEACH = `SAY:
- Quick second look at the hook idea
- Watch -- I will turn ONE long sentence into a SHORT hook in front of you
- Long: "The old house was very big and had a lot of dust everywhere and the windows were dirty and broken."
- I am going to cut, cut, cut until only the surprise is left
- "The old house was big" -- still telling, not showing
- "Dust covered everything" -- closer. Three words and a strong verb
- "The house breathed dust." -- THAT is a hook

DO:
- Display the re-teach slide
- Work the example LIVE on a mini-whiteboard or the board
- Cross out words as you go
- Re-check: ask students for ONE short hook for the SAME old house
- Take 2 or 3 board responses

TEACHER NOTES:
OPTIONAL slide. Only use this if the CFU showed fewer than about 80% students choosing B. The first I Do showed the finished model. This re-teach shows the PROCESS of getting from long-and-listing to short-and-surprising. Different approach: cut, not write. Most students who missed the CFU will get this when they see the cutting in action.

WATCH FOR:
- Students who write a SHORT hook on their boards -- ready for You Do
- Students who write 3 long sentences -- prompt: "Cut everything except the surprise"

[Literacy: Optional Re-teach | VTLM 2.0: Reteach with Different Representation]`;

const NOTES_YOUDO = `SAY:
- Your turn. You will write a setting paragraph
- Choose ONE of these places: an old library, a busy train station, a beach at sunrise, a school at night, a forest in winter
- Use the Setting Plan to brainstorm two senses first
- Then write your hook sentence
- Then write 3 to 5 more sentences using your senses
- 15 minutes. I will circulate

DO:
- Display the You Do slide
- Distribute the plan template and the mentor sheet
- Circulate -- prioritise students who looked unsure during CFU
- Quick conferences: "Read me your hook. Does it grab the reader?"

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the mentor sheet. Find the highlighted hook sentence and the highlighted sensory details. Write your own version following the same pattern -- short hook first, then two senses
- Extra Notes: These students still hit SC1 and SC2 -- the scaffold gives them a clear pattern to follow

EXTENDING PROMPT:
- Task: After your paragraph, add one final sentence that links the place to a FEELING or a hint of what is about to happen
- Extra Notes: Push for mood: "The wind held its breath" / "Something was wrong here"

TEACHER NOTES:
The five place choices keep all students writing the same SKILL but with different content. Active circulation is the formative assessment. Most students should complete a hook plus 3-4 sentences inside the block.

WATCH FOR:
- Students who write a long opening sentence -- prompt: "Can you cut that to under 8 words?"
- Students who use only SIGHT -- prompt: "What can you HEAR?"
- Students who get stuck on the place -- pick one for them quickly: "Use the train station"

[Literacy: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Quick exit ticket
- On your plan sheet, write ONE more setting hook -- one sentence
- The place: a hospital waiting room at midnight
- Under 10 words. Make me feel something
- 2 minutes. Drop it on my desk on your way out

DO:
- Display the exit ticket slide
- 2 minutes silent
- Collect on the way out
- Read 3 to 5 aloud at the start of Lesson 2 (no names) to celebrate

TEACHER NOTES:
Exit ticket targets SC2 (short hook sentence). One sentence is enough -- this is a quick check, not a paragraph. Collect for evidence of learning and to read aloud at the start of Lesson 2.

WATCH FOR:
- Students who write more than 10 words -- still collect; note for next lesson
- Students who write a strong short hook -- celebrate next session

[Literacy: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Self-check time
- Look at the three success criteria
- Show on your fingers: 1 (need help), 3 (got it), 5 (could teach it)
- Then partner share: what is ONE word you used in your setting that you are proud of?

DO:
- Run fingers check for each SC
- 30 seconds partner share
- Briefly: "Tomorrow we move from setting to CHARACTER -- the people in the place"

TEACHER NOTES:
Use the fingers data to plan Lesson 2's launch. If SC1 or SC2 was weak, open Lesson 2 with a quick setting recall.

WATCH FOR:
- Students who name a precise word -- celebrate openly
- Students who hold up 1 finger across the board -- privately check in

[Literacy: Closing | VTLM 2.0: Review and Reflect]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Holes -- Lesson 1 -- Narrative Setting";

  // SLIDE 1 -- Title
  titleSlide(
    pres,
    "Setting: Bring the Place to Life",
    "Camp Green Lake -- Holes by Louis Sachar",
    "Lesson 1  |  Week 1  |  Year 5/6 Literacy",
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

  // SLIDE 3 -- Hook / Text Launch (Chapter 1 read-aloud)
  quoteSlide(
    pres,
    "Read Aloud",
    "Holes -- Chapter 1",
    "There is no lake at Camp Green Lake.",
    "Ch. 1, opening line",
    "Close your eyes. What do you SEE? What do you HEAR? What do you FEEL?",
    NOTES_HOOK,
    FOOTER
  );

  // SLIDE 4 -- LI / SC
  liSlide(
    pres,
    [
      "We are learning to write a vivid setting paragraph that helps the reader SEE, HEAR and FEEL a place",
    ],
    [
      "I can name a place and use at least two senses in my setting",
      "I can write a short hook sentence that grabs the reader",
      "I can link my setting to a feeling or mood",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // SLIDE 5 -- Vocab 1 (no image -- text only is allowed)
  vocabSlide(
    pres,
    "barren",
    "adjective",
    "Empty -- nothing grows there. A barren place is bare, dry and lifeless.",
    "Camp Green Lake is a barren stretch of cracked earth.",
    NOTES_VOCAB_BARREN,
    FOOTER
  );

  // SLIDE 6 -- Vocab 2
  vocabSlide(
    pres,
    "shimmer",
    "verb",
    "To shine with a wavy, moving light. Heat can make the air shimmer above hot ground.",
    "In the distance, the desert air began to shimmer.",
    NOTES_VOCAB_SHIMMER,
    FOOTER
  );

  // SLIDE 7 -- I Do (model setting paragraph)
  modellingSlide(
    pres,
    "I Do -- Watch Me",
    "Model Setting Paragraph: Camp Green Lake",
    "HOOK:\n\n\"There is no lake at Camp Green Lake.\"\n\n- Short\n- Surprising\n- Names the place\n\nThen 3 to 5 sentences using:\n- sight\n- hearing\n- touch (heat, wind, dust)",
    "The lake had dried up over a hundred years ago. Now the ground was cracked and pale, and the heat shimmered above it like a second sky. A faded sign creaked in the wind. Nothing moved. The land felt barren -- like the place itself had stopped breathing.\n\n(My choices:\n- short hook first\n- sight: cracked, pale\n- touch: heat, wind\n- sound: creaked\n- mood: stopped breathing)",
    NOTES_IDO,
    FOOTER
  );

  // SLIDE 8 -- We Do (guided practice -- different place from I Do)
  {
    const s = pres.addSlide();
    addTopBar(s, C.SECONDARY);
    addBadge(s, "We Do", { color: C.SECONDARY });
    addTitle(s, "Together: Empty Playground at Dusk");

    const cardY = CONTENT_TOP;
    const cardH = 1.8;
    addCard(s, 0.5, cardY, 9, cardH, { strip: C.SECONDARY, fill: C.WHITE });
    s.addText("Five-Senses Brainstorm", {
      x: 0.75, y: cardY + 0.10, w: 8.5, h: 0.30,
      fontSize: 14, fontFace: FONT_B, color: C.SECONDARY, bold: true, margin: 0,
    });
    s.addText("- SEE:      what do you see? (long shadows? a swing? cracked tarmac?)\n- HEAR:    what do you hear? (chain creaking? a distant car? silence?)\n- SMELL:  what do you smell? (rubber? rain? a barbeque next door?)\n- TOUCH:  what do you feel? (cold metal? warm wind? gravel underfoot?)", {
      x: 0.75, y: cardY + 0.45, w: 8.5, h: cardH - 0.55,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    const tipY = cardY + cardH + 0.18;
    const tipH = SAFE_BOTTOM - tipY;
    addCard(s, 0.5, tipY, 9, tipH, { strip: C.ACCENT, fill: C.BG_CARD });
    s.addText("Together we will write the HOOK on the board:", {
      x: 0.75, y: tipY + 0.10, w: 8.5, h: 0.30,
      fontSize: 13, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("- Short. Under 10 words.\n- Names the place OR surprises the reader\n- Try one as a class: \"The playground had emptied for the night.\"", {
      x: 0.75, y: tipY + 0.42, w: 8.5, h: tipH - 0.55,
      fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_WEDO);
  }

  // SLIDE 9 + 10 -- CFU: Which hook is stronger? (reveal pair)
  function buildCfuBase() {
    const slide = pres.addSlide();
    addTopBar(slide, C.ALERT);
    addBadge(slide, "CFU", { color: C.ALERT });
    addTitle(slide, "Which Hook Grabs the Reader Better?", { color: C.ALERT });

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
    slide.addText("\"The forest had a lot of trees and it was cold and there was snow on the ground.\"", {
      x: 1.2, y: cardY + 0.18, w: 8.0, h: cardH - 0.30,
      fontSize: 15, fontFace: FONT_H, color: C.CHARCOAL, italic: true, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    const cardBY = cardY + cardH + 0.18;
    const cardBH = SAFE_BOTTOM - cardBY - 0.55;
    addCard(slide, 0.5, cardBY, 9, cardBH, { strip: C.SECONDARY, fill: C.BG_CARD });
    slide.addText("B", {
      x: 0.7, y: cardBY + 0.10, w: 0.5, h: 0.4,
      fontSize: 22, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0,
    });
    slide.addText("\"The forest had stopped breathing.\"", {
      x: 1.2, y: cardBY + 0.14, w: 8.0, h: cardBH - 0.24,
      fontSize: 22, fontFace: FONT_H, color: C.CHARCOAL, italic: true, valign: "middle", margin: 0,
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
      slide.addText("Stronger hook: B  --  short, surprising, makes you FEEL something", {
        x: 0.5, y: revealY, w: 9, h: 0.42,
        fontSize: 14, fontFace: FONT_B, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      slide.addNotes(NOTES_CFU_REVEAL);
    }
  );

  // SLIDE 11 -- Optional Re-teach (CFU follow-up, different approach)
  modellingSlide(
    pres,
    "Optional Re-teach",
    "Cut, Cut, Cut Until Only the Surprise Is Left",
    "Long version:\n\n\"The old house was very big and had a lot of dust everywhere and the windows were dirty and broken.\"\n\nWatch me cut.",
    "Step 1: \"The old house was big.\"\n   (still telling, not showing)\n\nStep 2: \"Dust covered everything.\"\n   (closer -- strong verb)\n\nStep 3: \"The house breathed dust.\"\n   THAT is a hook.\n\nYour turn: write ONE short hook for the SAME old house on your board.",
    NOTES_RETEACH,
    FOOTER
  );

  // SLIDE 12 -- You Do
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "You Do", { color: C.PRIMARY, w: 1.5 });
    addTitle(s, "Write Your Own Setting Paragraph");

    addCard(s, 0.5, CONTENT_TOP, 9, 1.7, { strip: C.PRIMARY, fill: C.WHITE });
    s.addText("Pick ONE place:", {
      x: 0.75, y: CONTENT_TOP + 0.10, w: 8.4, h: 0.30,
      fontSize: 14, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("- an old library\n- a busy train station\n- a beach at sunrise\n- a school at night\n- a forest in winter", {
      x: 0.75, y: CONTENT_TOP + 0.42, w: 8.4, h: 1.20,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    const tipY = CONTENT_TOP + 1.85;
    const tipH = SAFE_BOTTOM - tipY;
    addCard(s, 0.5, tipY, 9, tipH, { strip: C.ACCENT, fill: C.BG_CARD });
    s.addText("Writing Time: 15 minutes", {
      x: 0.75, y: tipY + 0.10, w: 5, h: 0.30,
      fontSize: 13, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("First:    Use the Setting Plan -- brainstorm TWO senses\nNext:    Write a short hook sentence (under 10 words)\nThen:    Write 3 to 5 more sentences using your senses\nFinally: Read it back aloud quietly -- can you SEE the place?", {
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
    "One Setting Hook -- A Hospital Waiting Room at Midnight",
    [
      "Write ONE hook sentence about a hospital waiting room at midnight",
      "Under 10 words",
      "Make me FEEL something",
      "Drop it on my desk on your way out -- 2 minutes",
    ],
    NOTES_EXIT,
    FOOTER
  );

  // SLIDE 14 -- Closing Reflection
  closingSlide(
    pres,
    {
      reflectionPrompt: "Partner share: what is ONE word you used in your setting that you are proud of?",
      scItems: [
        "I can name a place and use at least two senses in my setting",
        "I can write a short hook sentence that grabs the reader",
        "I can link my setting to a feeling or mood",
      ],
      selfAssessment: {
        prompt: "Show on your fingers: 1 (need help) to 5 (could teach it)",
        options: ["1-2", "3", "4-5"],
      },
    },
    NOTES_CLOSING
  );

  // ---- PDF: Setting Plan -----------------------------------------------------
  const plan = createPdf({ title: PLAN_RESOURCE.name });
  let planY = addPdfHeader(plan, "Setting Plan -- Bring a Place to Life", {
    color: C.PRIMARY,
    subtitle: "Plan your setting paragraph using the five senses and one short hook",
    lessonInfo: "Lesson 1 | Week 1 | Year 5/6 Literacy | Holes",
    showNameDate: true,
  });

  planY = addTipBox(plan, "A strong setting has TWO things: a short hook (the first sentence), and details that use your SENSES. The reader should be able to SEE, HEAR or FEEL the place. Plan first, then write.", planY, { color: C.PRIMARY });

  planY = addSectionHeading(plan, "Step 1 -- Choose your place", planY, { color: C.PRIMARY });
  planY = addBodyText(plan, "Circle ONE: old library  /  busy train station  /  beach at sunrise  /  school at night  /  forest in winter  /  your own choice: __________", planY, { fontSize: 11 });
  planY += 6;

  planY = addSectionHeading(plan, "Step 2 -- Brainstorm two senses", planY, { color: C.PRIMARY });
  planY = addBodyText(plan, "Pick TWO senses you will use. Write 3 ideas under each.", planY, { fontSize: 10, italic: true });
  planY = addBodyText(plan, "SEE:", planY, { fontSize: 11 });
  planY = addLinedArea(plan, planY, 2, { lineSpacing: 18 });
  planY = addBodyText(plan, "HEAR:", planY, { fontSize: 11 });
  planY = addLinedArea(plan, planY, 2, { lineSpacing: 18 });
  planY = addBodyText(plan, "SMELL or TOUCH (one more):", planY, { fontSize: 11 });
  planY = addLinedArea(plan, planY, 2, { lineSpacing: 18 });
  planY += 4;

  planY = addSectionHeading(plan, "Step 3 -- Write your hook (one short sentence, under 10 words)", planY, { color: C.SECONDARY });
  planY = addLinedArea(plan, planY, 2, { lineSpacing: 20 });

  addPdfFooter(plan, "Lesson 1 | Setting Plan -- Page 1");

  // Page 2 -- write the paragraph
  plan.addPage();
  let planY2 = addPdfHeader(plan, "Setting Paragraph -- Write Here", {
    color: C.PRIMARY,
    subtitle: "Hook + 3 to 5 sensory sentences",
    lessonInfo: "Lesson 1 | Week 1 | Year 5/6 Literacy",
    showNameDate: false,
  });

  planY2 = addTipBox(plan, "Start with your hook. Then use your sensory notes from page 1. Aim for 4 to 6 sentences total. Read it back aloud quietly when you finish.", planY2, { color: C.SECONDARY });

  planY2 = addLinedArea(plan, planY2, 14, { lineSpacing: 22 });

  addPdfFooter(plan, "Lesson 1 | Setting Paragraph -- Page 2");

  // Page 3 -- Exit ticket space
  plan.addPage();
  let planY3 = addPdfHeader(plan, "Exit Ticket", {
    color: C.ACCENT,
    subtitle: "One setting hook -- a hospital waiting room at midnight",
    lessonInfo: "Lesson 1 | Week 1",
    showNameDate: false,
  });

  planY3 = addTipBox(plan, "Write ONE hook sentence about a hospital waiting room at midnight. Under 10 words. Make me feel something. Hand this in.", planY3, { color: C.ACCENT });

  planY3 = addLinedArea(plan, planY3, 3, { lineSpacing: 24 });

  addPdfFooter(plan, "Lesson 1 | Exit Ticket");

  // ---- PDF: Mentor Setting Paragraph ----------------------------------------
  const mp = createPdf({ title: MENTOR_RESOURCE.name });
  let mpY = addPdfHeader(mp, "Mentor Setting Paragraph -- Annotated", {
    color: C.PRIMARY,
    subtitle: "Camp Green Lake -- a model paragraph based on the opening of Holes",
    lessonInfo: "Lesson 1 | Week 1 | Year 5/6 Literacy",
    showNameDate: false,
  });

  mpY = addTipBox(mp, "This is a model setting paragraph for Camp Green Lake. Use it as a REFERENCE -- do not copy it. Your setting will use YOUR chosen place and YOUR senses.", mpY, { color: C.PRIMARY });

  mpY = addSectionHeading(mp, "Model paragraph", mpY, { color: C.PRIMARY });
  mpY = addBodyText(mp, "There is no lake at Camp Green Lake. The lake had dried up over a hundred years ago. Now the ground was cracked and pale, and the heat shimmered above it like a second sky. A faded sign creaked in the wind. Nothing moved. The land felt barren -- like the place itself had stopped breathing.", mpY, { fontSize: 12 });
  mpY += 12;

  mpY = addSectionHeading(mp, "Annotations -- the hook", mpY, { color: C.PRIMARY });
  mpY = addBodyText(mp, "\"There is no lake at Camp Green Lake.\" -- the actual opening line of the novel. Short (7 words). Surprising (a place named after something that does not exist). Names the place.", mpY, { fontSize: 10 });
  mpY += 6;

  mpY = addSectionHeading(mp, "Annotations -- the senses", mpY, { color: C.SECONDARY });
  mpY = addBodyText(mp, "SIGHT: \"cracked and pale\", \"faded sign\" -- precise visual detail.", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "TOUCH (heat): \"the heat shimmered above it like a second sky\" -- the reader can FEEL the heat.", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "SOUND: \"creaked in the wind\" -- one careful sound carries the silence.", mpY, { fontSize: 10 });
  mpY += 6;

  mpY = addSectionHeading(mp, "Annotations -- mood", mpY, { color: C.ACCENT });
  mpY = addBodyText(mp, "\"the place itself had stopped breathing\" -- personification. Connects the setting to a FEELING (stillness, lifelessness, dread).", mpY, { fontSize: 10 });
  mpY += 8;

  mpY = addSectionHeading(mp, "What to copy and what NOT to copy", mpY, { color: C.PRIMARY });
  mpY = addBodyText(mp, "COPY the pattern: short hook, then sensory detail, then a mood sentence.", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "DO NOT copy the words. Your place is different. Use YOUR senses, YOUR verbs, YOUR mood.", mpY, { fontSize: 10 });

  addPdfFooter(mp, "Lesson 1 | Mentor Setting Paragraph -- TEACHER AND STUDENT REFERENCE");

  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Holes_Lesson1.pptx` }),
    writePdf(plan, PLAN_PDF_PATH),
    writePdf(mp, MENTOR_PDF_PATH),
  ]);

  console.log("PPTX written to " + `${OUT_DIR}/Holes_Lesson1.pptx`);
  console.log("Done: " + PLAN_RESOURCE.name);
  console.log("Done: " + MENTOR_RESOURCE.name);
}

build().catch(console.error);
