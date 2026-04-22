"use strict";

// Inquiry - Chimney Sweep Boys in 1780s London
// Grade 5/6 | Source Analysis | 60-minute lesson
// Students use two contrasting sources to understand life for climbing boys
// Source 1: Blake, "The Chimney Sweeper" (1789) - primary source (public domain)
// Source 2: Historical Background card - factual context card (secondary source)

const path = require("path");
const fs = require("fs");
const PptxGenJS = require("pptxgenjs");
const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf,
  addPdfHeader, addSectionHeading, addBodyText, addTipBox,
  addPdfFooter, addLinedArea, addWriteLine, addTwoColumnOrganiser,
  addResourceSlide, makeSessionResource, getSessionResourceFolder,
  PAGE, hex,
} = require("../themes/pdf_helpers");

// -- Theme --
const T = createTheme("inquiry", "grade56", weekToVariant(2)); // Variant 1 - Detective palette
const {
  C, FONT_H, FONT_B,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  withReveal,
  addTopBar, addBadge, addTitle, addCard, addFooter, addTextOnShape,
  addInstructionCard,
  CONTENT_TOP, SAFE_BOTTOM, SLIDE_W,
  runSlideDiagnostics,
} = T;

// -- Output paths --
const UNIT = "Inquiry_Chimney_Sweeps_1780s";
const LESSON_FOLDER = path.join(__dirname, "..", "output", UNIT);
const PPTX_NAME = "Chimney Sweep Boys in 1780s London.pptx";
const FOOTER = "Inquiry | Grade 5/6 | Life in 1780s London";
const SESSION = 1;
const RES_DIR = path.join(LESSON_FOLDER, getSessionResourceFolder(SESSION));

// -- Resources --
const SOURCE_PACK = makeSessionResource(
  SESSION,
  "Source Card Pack",
  "Printable set of two source cards: Blake's poem and Historical Background."
);
const ORGANISER = makeSessionResource(
  SESSION,
  "Source Analysis Organiser",
  "Graphic organiser to record evidence from each source."
);
const EXIT_TICKET = makeSessionResource(
  SESSION,
  "Exit Ticket",
  "One-sentence claim with evidence about life for chimney sweep boys."
);
const RESOURCE_ITEMS = [SOURCE_PACK, ORGANISER, EXIT_TICKET];

fs.mkdirSync(RES_DIR, { recursive: true });

// ===============================================================
// Teacher Notes
// ===============================================================

const NOTES_TITLE = [
  "SAY:",
  "- Today we are stepping back to 1780s London to meet children who did a job most adults wouldn't do",
  "- We'll read two sources about chimney sweep boys and work out what life was really like for them",
  "",
  "DO:",
  "- Have the Source Card Pack printed and ready to distribute at the I Do stage",
  "- Have the Source Analysis Organiser ready for the You Do stage",
  "",
  "TEACHER NOTES:",
  "This is a source analysis lesson. Keep the focus on what the sources tell us, not on generalising beyond them.",
  "",
  "SENSITIVITY ADVISORY:",
  "- What it is: Child labour, mistreatment, and early death of children were real features of this period.",
  "- Framing language: 'This was a hard history. Historians study it so we understand what changed and why it changed.'",
  "- Watch for: students who appear unsettled or who have lived experience of harsh work.",
  "- Protocol: If a student needs a break, offer a quiet task at the desk and check in privately after class.",
  "",
  "WATCH FOR:",
  "- Students settling in; avoid rushing the topic setup",
  "",
  "[Inquiry: Title | VTLM 2.0: Establishing Purpose]",
].join("\n");

const NOTES_LI = [
  "SAY:",
  "- Read the Learning Intention aloud",
  "- Read each Success Criterion. We'll come back to these at the end to check how we went",
  "",
  "DO:",
  "- Point to each SC as you read it",
  "- Leave the slide visible for about 20 seconds",
  "",
  "TEACHER NOTES:",
  "SC1 is the foundation every student can reach - spotting one detail. SC2 is the core - using evidence to describe life. SC3 extends into comparing two sources.",
  "",
  "WATCH FOR:",
  "- Students unsure what a 'primary source' means - this is covered on the vocabulary slide next",
  "",
  "[Inquiry: LI/SC | VTLM 2.0: Learning Intentions]",
].join("\n");

const NOTES_HOOK = [
  "SAY:",
  "- Close your eyes for a moment. Picture a big London house in the year 1780",
  "- The house has tall chimneys. They are narrow - some only as wide as a dinner plate",
  "- Every room has a fireplace. Coal burns all winter. Soot builds up inside the chimneys and has to be cleaned out",
  "- Ask: Who do you think did this cleaning? [Accept: sweeps, workers] Why might grown adults not do it themselves? [Too big to fit]",
  "",
  "DO:",
  "- Read the scenario slowly. Leave pauses for picturing",
  "- Cold call 2-3 students for the follow-up question",
  "",
  "TEACHER NOTES:",
  "The hook establishes the practical problem: narrow chimneys needed small bodies to climb them. This sets up why young boys were used, without giving the answer yet.",
  "",
  "WATCH FOR:",
  "- Students who guess 'children' straight away - good, they are tracking with the scenario",
  "- Students who seem confused - recheck that they understand chimneys need cleaning",
  "",
  "[Inquiry: Hook | VTLM 2.0: Activating Prior Knowledge]",
].join("\n");

const NOTES_VOCAB = [
  "SAY:",
  "- Before we read our sources, let's get four words clear",
  "- Read each word and its meaning with me",
  "- 'Primary source' means evidence from the actual time - like a poem, a letter, or a photo from back then. That is what we will look at today",
  "",
  "DO:",
  "- Read each word aloud. Have students repeat 'climbing boy' and 'apprentice' after you",
  "- Point to the small icons as you go",
  "",
  "CFU CHECKPOINT:",
  "Technique: Finger Voting",
  "Script:",
  "- Say: Hold up 1 finger if the chimney sweep is the adult master, 2 if it is the boy who climbs up",
  "- Scan for: most students showing 1 - adults were the official sweeps; boys were the climbing boys",
  "PROCEED: If 80% show 1, move on to the source.",
  "PIVOT: If students are confused, re-read the two terms and give a physical example: 'The sweep is the adult running the business. The climbing boy is the child doing the dangerous work inside the chimney.'",
  "",
  "TEACHER NOTES:",
  "Students need climbing boy, apprentice, and soot before the Blake poem makes sense. Keep this brief.",
  "",
  "WATCH FOR:",
  "- Students confusing 'sweep' (the trade/adult) with 'climbing boy' (the child) - correct straight away",
  "",
  "[Inquiry: Vocabulary | VTLM 2.0: Establishing Knowledge]",
].join("\n");

const NOTES_SOURCE1_IDO = [
  "SAY:",
  "- This is Source 1. A poem from the year 1789, written by William Blake, a London poet",
  "- Watch me read it first. Pay attention to WHO is speaking",
  "- [Read stanza aloud slowly with expression]",
  "- Ask: Who is telling this poem? [The chimney sweep boy himself]",
  "- Think-aloud: I notice the speaker says 'my mother died I was very young'. That tells me he lost his mother early",
  "- Think-aloud: Then 'my father sold me'. That word 'sold' is shocking. Fathers were paid to apprentice their sons to masters",
  "- Think-aloud: 'weep weep weep' - that was the actual street cry of climbing boys. They were trying to shout 'sweep' but they were so small it came out as 'weep'",
  "- Think-aloud: 'in soot I sleep' - he doesn't even have a clean bed. He sleeps in the soot he cleans",
  "",
  "DO:",
  "- Distribute the Session 1 Source Card Pack",
  "- Read the stanza aloud, pointing to each line",
  "- Circle or underline the four key details on your own copy as you model",
  "",
  "TEACHER NOTES:",
  "This is the I Do modelling slide. Demonstrate how a historian notices details, not just reads the words. Students watch; they respond in the next slide.",
  "",
  "SOURCES:",
  "- William Blake, 'The Chimney Sweeper' from Songs of Innocence (1789). Public domain. First stanza used.",
  "",
  "WATCH FOR:",
  "- Students following along on their source card",
  "- Students reacting to 'sold' - acknowledge the reaction and name it: that is the moment that tells us a lot",
  "",
  "[Inquiry: I Do | VTLM 2.0: Explicit Modelling]",
].join("\n");

const NOTES_CFU_Q = [
  "SAY:",
  "- On your mini-whiteboard, write the ONE word from Blake's poem that best shows the boy was treated as a thing, not a person",
  "- You have 30 seconds. Use your source card",
  "",
  "DO:",
  "- Set a 30-second timer",
  "- Scan whiteboards as students hold them up",
  "",
  "CFU CHECKPOINT:",
  "Technique: Show Me Boards",
  "Script:",
  "- Say: Boards up on three. One, two, three",
  "- Scan for: 'sold' as the target answer. 'weep' or 'sleep' show partial understanding",
  "PROCEED: If 80% write 'sold', reveal and move on.",
  "PIVOT: If many write 'weep' or 'sleep', re-read line 2 and ask: 'Which word shows the boy was treated like property?' Most should now land on 'sold'.",
  "",
  "TEACHER NOTES:",
  "This is the hinge. Wrong answers are still useful - weep shows they noticed sound; sleep shows they noticed conditions; sold shows they noticed the boy's status.",
  "",
  "WATCH FOR:",
  "- Students staring at the slide rather than the source card - redirect to their printed source",
  "- Students answering in full sentences - praise but remind them: one word only for this check",
  "",
  "[Inquiry: CFU | VTLM 2.0: Checking for Understanding]",
].join("\n");

const NOTES_CFU_A = [
  "SAY:",
  "- The target word is 'sold'",
  "- 'Sold' shows the boy was treated as property - something that could be bought and handed over, not a child with rights",
  "- If you wrote 'weep' or 'sleep' - good noticing. Those words show the hard conditions. Blake packed a lot into just four lines",
  "",
  "DO:",
  "- Click to reveal the answer",
  "- Briefly acknowledge 'weep' and 'sleep' answers as close-but-not-target",
  "",
  "TEACHER NOTES:",
  "Validate partial answers so students stay in the game for the We Do. Do not dwell - move on promptly.",
  "",
  "WATCH FOR:",
  "- Students who want to argue their alternative answer - acknowledge and say we'll use those words in the next source",
  "",
  "[Inquiry: CFU Reveal | VTLM 2.0: Feedback]",
].join("\n");

const NOTES_SOURCE2_QWE = [
  "SAY:",
  "- Here is Source 2. This is a Historical Background card. It is not a poem - it is a summary of what historians have found out about climbing boys",
  "- I'll read it once. Then we'll find details together",
  "- [Read the four background points aloud]",
  "- Ask: What detail tells us HOW YOUNG some climbing boys were? [As young as four years old]",
  "- Ask: What detail tells us about HEALTH danger? [Soot caused a cancer; burns; suffocation]",
  "- Ask: What detail tells us things changed? [An Act of Parliament in 1788 set a minimum age of 8]",
  "",
  "DO:",
  "- Leave this slide up",
  "- Cold call for each question above; write student answers on the board",
  "- After questions, tell students: on the next slide we'll check the four key details",
  "",
  "CFU CHECKPOINT:",
  "Technique: Turn and Tell",
  "Script:",
  "- Say: Turn to your partner. In 20 seconds, tell them ONE detail from Source 2 that matches a detail in Blake's poem",
  "- Scan for: students pointing to their source cards; pairs agreeing on a match (e.g. 'sold' in poem -> apprenticed to a master)",
  "PROCEED: If pairs are producing matches, reveal the class findings on the next click.",
  "PIVOT: If pairs are stuck, model one match: 'Blake says the boy was sold. Source 2 says boys were apprenticed. Those are two different words for the same thing.'",
  "",
  "TEACHER NOTES:",
  "Students are now applying the noticing skill they saw modelled. The Turn and Tell builds confidence before the You Do task.",
  "",
  "SOURCES:",
  "- Historical Background adapted from widely documented history of 18th century London child labour. Key facts: climbing boys apprenticed as young as four; risks included burns, falls, suffocation, and the 'chimney sweeps' carcinoma' identified by Percivall Pott in 1775; the Chimney Sweepers Act 1788 set a minimum apprenticeship age of eight. Teachers may substitute a school-approved extract.",
  "",
  "WATCH FOR:",
  "- Students reading from the poem instead of from Source 2 - prompt them to use the Historical Background card",
  "",
  "[Inquiry: We Do | VTLM 2.0: Scaffold Practice]",
].join("\n");

const NOTES_SOURCE2_AWE = [
  "SAY:",
  "- Here are the key details historians want us to notice from Source 2",
  "- Age - some boys were as young as four",
  "- Danger - burns, falls, and a cancer caused by soot were common",
  "- Living conditions - boys slept near the soot they collected",
  "- Change over time - laws began to protect them in 1788",
  "- Ask: Which detail surprised you most? [Accept answers; this is for processing, not grading]",
  "",
  "DO:",
  "- Click to reveal the four highlighted details",
  "- Give 30 seconds for students to annotate on their source cards",
  "- Take 1-2 hands for the surprise question",
  "",
  "TEACHER NOTES:",
  "This reveal consolidates the We Do. Students now have four clear details from Source 2 plus four from Blake. They are ready for independent analysis.",
  "",
  "WATCH FOR:",
  "- Students writing the four details on their source card - this is the scaffold for the You Do",
  "",
  "[Inquiry: We Do Reveal | VTLM 2.0: Consolidation]",
].join("\n");

const NOTES_YOUDO = [
  "SAY:",
  "- Your turn. You are going to use BOTH sources to answer one big question",
  "- Question: What was life like for chimney sweep boys in 1780s London?",
  "- First: Write one detail from Blake's poem in the left column",
  "- Next: Write one detail from the Historical Background in the right column",
  "- Then: In the bottom box, write one sentence that uses evidence from BOTH sources",
  "",
  "DO:",
  "- Distribute the Session 1 Source Analysis Organiser",
  "- Project the slide while students work",
  "- Circulate. Prompt quiet students with: 'Look at line 2 of Blake. What word jumps out?'",
  "- Set a 10-minute timer",
  "",
  "ENABLING & EXTENDING:",
  "ENABLING PROMPT:",
  "- Task: Students highlight one detail in each source with a coloured pencil, then copy those two details into the organiser. Use a sentence starter: 'Life was hard for chimney sweep boys because ___ and ___.'",
  "- Extra Notes: Pair with a partner who can read the sources aloud if needed.",
  "EXTENDING PROMPT:",
  "- Task: Students find one detail that appears in BOTH sources (e.g. hard conditions) and one detail that appears in only ONE source (e.g. the 1788 law). They write a sentence explaining why each source tells us something different.",
  "",
  "TEACHER NOTES:",
  "The You Do uses different cognitive demand from the We Do. Students must combine evidence across sources rather than just spot details in one.",
  "",
  "WATCH FOR:",
  "- Students copying out whole lines rather than choosing one detail - redirect to the strongest single detail",
  "- Students finishing early - prompt them to try the extending task",
  "",
  "[Inquiry: You Do | VTLM 2.0: Supported Application]",
].join("\n");

const NOTES_EXIT = [
  "SAY:",
  "- One sentence on your exit ticket to finish",
  "- Complete this frame: 'Life for chimney sweep boys in 1780s London was ___ because ___.'",
  "- Your 'because' must use evidence from a source - name the source",
  "",
  "DO:",
  "- Distribute the Session 1 Exit Ticket",
  "- Allow 4 minutes",
  "- Collect as students leave",
  "",
  "TEACHER NOTES:",
  "This exit ticket directly assesses SC2 (using evidence) and touches SC3 (naming the source). Quickly sort into three piles after class: evidence named clearly / evidence implied / no evidence - plan reteach for the third pile.",
  "",
  "WATCH FOR:",
  "- Students writing feelings without evidence - prompt: 'Which word from which source tells you that?'",
  "",
  "[Inquiry: Exit Ticket | VTLM 2.0: Evidence of Learning]",
].join("\n");

const NOTES_CLOSING = [
  "SAY:",
  "- Let's come back to our success criteria. Thumbs up, sideways, or down for each one",
  "- SC1: I can identify details from a primary source",
  "- SC2: I can describe what life was like for chimney sweep boys using evidence",
  "- SC3: I can compare what two different sources tell us",
  "- Ask: Why do historians use more than one source when they study a topic? [Different sources show different parts; one source alone can mislead]",
  "",
  "DO:",
  "- Read each SC aloud and pause for thumbs",
  "- Note which SC students feel least confident about",
  "- Acknowledge the weight of the topic. These were real children",
  "",
  "TEACHER NOTES:",
  "The closing question builds the key historical thinking move - using multiple sources. This is the bridge to any next lesson on Industrial Revolution child labour, reform movements, or primary/secondary source work.",
  "",
  "WATCH FOR:",
  "- Students showing thumbs down on SC3 - flag for a short revisit next session",
  "- Students who seem reflective - this is a sign of engagement, not distress; affirm their thinking",
  "",
  "[Inquiry: Closing | VTLM 2.0: Review and Reflect]",
].join("\n");

const NOTES_RESOURCES = [
  "SAY:",
  "- These three resources support today's lesson",
  "",
  "DO:",
  "- Print the Source Card Pack - one per student, double-sided if possible",
  "- Print the Source Analysis Organiser - one per student",
  "- Print the Exit Ticket - one per student",
  "",
  "TEACHER NOTES:",
  "All three resources are in the lesson folder. The Source Card Pack doubles as the reference handout; students can annotate it throughout the lesson.",
  "",
  "WATCH FOR:",
  "- Print the Source Card Pack before the lesson starts so it is ready for the I Do",
  "",
  "[Inquiry: Resources | VTLM 2.0: Preparation]",
].join("\n");

// ===============================================================
// Build function
// ===============================================================

async function build() {
  const pres = new PptxGenJS();
  pres.layout = "LAYOUT_16x9";

  // -- Slide 1: Title --
  titleSlide(
    pres,
    "Chimney Sweep Boys\nin 1780s London",
    "Inquiry - Using Primary Sources",
    "Grade 5/6 | Session 1",
    NOTES_TITLE
  );

  // -- Slide 2: LI / SC --
  liSlide(
    pres,
    ["We are learning how to use primary sources to understand what life was like for chimney sweep boys in 1780s London"],
    [
      "I can identify details from a primary source",
      "I can describe what life was like for chimney sweep boys using evidence from a source",
      "I can compare what two different sources tell us about the same topic",
    ],
    NOTES_LI,
    FOOTER
  );

  // -- Slide 3: Hook / Imagine --
  contentSlide(
    pres,
    "Hook",
    C.SECONDARY,
    "Imagine London, 1780",
    [
      "Tall brick houses with tall, narrow chimneys",
      "Coal fires burn in every room all winter",
      "Soot builds up inside the chimneys",
      "Someone has to climb inside and clean them out",
    ],
    NOTES_HOOK,
    FOOTER,
    (s, lg) => {
      // Right column: atmospheric scene card (simple silhouette mockup)
      const rX = lg.rightX;
      const rW = lg.rightW;
      const topY = lg.panelTopPadded;
      const cardH = 3.35;

      // Scene card - dark background suggesting soot-stained London
      addCard(s, rX, topY, rW, cardH, { strip: C.SECONDARY, fill: C.BG_DARK });

      // "Sky" - pale top band
      s.addShape("rect", {
        x: rX + 0.15, y: topY + 0.5, w: rW - 0.3, h: 0.75,
        fill: { color: "3E3528" },
      });

      // Chimney silhouettes - stacks of roundRects
      const chimneyCfg = [
        { x: rX + 0.4, h: 1.5, w: 0.25 },
        { x: rX + 0.85, h: 1.9, w: 0.28 },
        { x: rX + 1.35, h: 1.35, w: 0.22 },
        { x: rX + 1.75, h: 2.05, w: 0.3 },
        { x: rX + 2.25, h: 1.65, w: 0.25 },
        { x: rX + 2.7, h: 1.85, w: 0.28 },
        { x: rX + 3.15, h: 1.45, w: 0.22 },
        { x: rX + 3.55, h: 1.75, w: 0.26 },
      ];
      chimneyCfg.forEach((ch) => {
        if (ch.x + ch.w > rX + rW - 0.1) return;
        const chY = topY + cardH - 0.5 - ch.h;
        // House block
        s.addShape("rect", {
          x: ch.x - 0.02, y: chY + 0.3, w: ch.w + 0.04, h: ch.h - 0.3,
          fill: { color: "1A1410" },
        });
        // Chimney stack
        s.addShape("rect", {
          x: ch.x + ch.w * 0.3, y: chY, w: ch.w * 0.4, h: 0.35,
          fill: { color: "0F0B08" },
        });
        // Smoke wisp
        s.addShape("ellipse", {
          x: ch.x + ch.w * 0.2, y: chY - 0.25, w: ch.w * 0.6, h: 0.2,
          fill: { color: "6B5E4E", transparency: 60 },
        });
      });

      // Ground strip
      s.addShape("rect", {
        x: rX + 0.15, y: topY + cardH - 0.5, w: rW - 0.3, h: 0.35,
        fill: { color: "0A0806" },
      });

      // Caption band
      s.addShape("rect", {
        x: rX + 0.15, y: topY + cardH - 0.35, w: rW - 0.3, h: 0.25,
        fill: { color: C.SECONDARY },
      });
      s.addText("London rooftops at dawn", {
        x: rX + 0.15, y: topY + cardH - 0.35, w: rW - 0.3, h: 0.25,
        fontSize: 10, fontFace: FONT_B, color: C.WHITE, italic: true,
        align: "center", valign: "middle", margin: 0,
      });
    }
  );

  // -- Slide 4: Key Vocabulary --
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "Vocabulary", { color: C.PRIMARY, w: 1.65 });
    addTitle(s, "Four Words Before We Read");

    const vocab = [
      { word: "climbing boy", meaning: "a young boy who climbed up inside chimneys to clean out the soot", color: C.SECONDARY },
      { word: "apprentice", meaning: "a child bound by contract to work for a master to learn a trade", color: C.PRIMARY },
      { word: "soot", meaning: "the black powder left behind when coal or wood burns", color: C.ACCENT },
      { word: "primary source", meaning: "evidence from the time itself - a poem, letter, or photo made back then", color: C.SUCCESS },
    ];

    const cardW = 4.45;
    const cardH = 1.8;
    const gapX = 0.1;
    const gapY = 0.1;

    vocab.forEach((v, i) => {
      const col = i % 2;
      const row = Math.floor(i / 2);
      const x = 0.5 + col * (cardW + gapX);
      const y = CONTENT_TOP + row * (cardH + gapY);

      addCard(s, x, y, cardW, cardH, { strip: v.color, fill: C.WHITE });

      // Word
      s.addText(v.word, {
        x: x + 0.2, y: y + 0.15, w: cardW - 0.4, h: 0.5,
        fontSize: 22, fontFace: FONT_H, color: v.color, bold: true, margin: 0,
        fit: "shrink",
      });

      // Meaning
      s.addText(v.meaning, {
        x: x + 0.2, y: y + 0.7, w: cardW - 0.4, h: cardH - 0.85,
        fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
        valign: "top",
      });
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_VOCAB);
  }

  // ---- Shared helper: source card for the poem / background extract ----
  function drawPoemSourceCard(s, x, y, w, h, opts) {
    const o = opts || {};
    // Aged-paper background
    addCard(s, x, y, w, h, { strip: o.strip || C.PRIMARY, fill: "FAF5E8" });

    // Source label
    s.addText(o.sourceType || "PRIMARY SOURCE", {
      x: x + 0.25, y: y + 0.1, w: w - 0.5, h: 0.22,
      fontSize: 9.5, fontFace: FONT_B, color: C.MUTED, bold: true, margin: 0,
    });
    // Source title / citation
    s.addText(o.title || "", {
      x: x + 0.25, y: y + 0.3, w: w - 0.5, h: 0.3,
      fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0,
      fit: "shrink",
    });
    if (o.citation) {
      s.addText(o.citation, {
        x: x + 0.25, y: y + 0.58, w: w - 0.5, h: 0.22,
        fontSize: 10, fontFace: FONT_B, color: C.CHARCOAL, italic: true, margin: 0,
      });
    }

    // Divider
    s.addShape("line", {
      x: x + 0.25, y: y + 0.82, w: w - 0.5, h: 0,
      line: { color: C.MUTED, width: 0.8 },
    });

    // Poem body - single text box, let PptxGenJS handle line breaks so wraps don't collide
    const lines = o.lines || [];
    const bodyY = y + 0.9;
    const bodyH = y + h - bodyY - 0.15;
    const fontSize = o.fontSize || 14;
    s.addText(
      lines.map((ln, i) => ({
        text: ln,
        options: {
          breakLine: i < lines.length - 1,
          fontSize,
          color: "2C2416",
        },
      })),
      {
        x: x + 0.35, y: bodyY, w: w - 0.7, h: bodyH,
        fontFace: "Georgia", margin: 0, valign: "top",
        paraSpaceAfter: 4,
      }
    );
  }

  function drawBackgroundSourceCard(s, x, y, w, h, items, opts) {
    const o = opts || {};
    addCard(s, x, y, w, h, { strip: o.strip || C.SECONDARY, fill: "FAF5E8" });

    s.addText("HISTORICAL BACKGROUND", {
      x: x + 0.25, y: y + 0.1, w: w - 0.5, h: 0.22,
      fontSize: 9.5, fontFace: FONT_B, color: C.MUTED, bold: true, margin: 0,
    });
    s.addText(o.title || "What Historians Have Found", {
      x: x + 0.25, y: y + 0.3, w: w - 0.5, h: 0.3,
      fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0,
      fit: "shrink",
    });
    if (o.citation) {
      s.addText(o.citation, {
        x: x + 0.25, y: y + 0.58, w: w - 0.5, h: 0.22,
        fontSize: 10, fontFace: FONT_B, color: C.CHARCOAL, italic: true, margin: 0,
      });
    }

    s.addShape("line", {
      x: x + 0.25, y: y + 0.82, w: w - 0.5, h: 0,
      line: { color: C.MUTED, width: 0.8 },
    });

    // Body bullets - each a labelled row
    const bodyY = y + 0.95;
    const rowH = 0.48;
    items.forEach((it, i) => {
      const ry = bodyY + i * rowH;
      if (ry + rowH - 0.05 > y + h - 0.1) return;
      // Icon circle
      s.addShape("roundRect", {
        x: x + 0.3, y: ry + 0.05, w: 0.32, h: 0.32,
        rectRadius: 0.16, fill: { color: it.color || C.SECONDARY },
      });
      s.addText(it.label || "", {
        x: x + 0.3, y: ry + 0.05, w: 0.32, h: 0.32,
        fontSize: 11, fontFace: FONT_H, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      // Detail
      s.addText(it.text || "", {
        x: x + 0.72, y: ry, w: w - 0.92, h: rowH,
        fontSize: 12.5, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
        valign: "middle",
      });
    });
  }

  // -- Slide 5: I Do - Source 1 (Blake poem) --
  contentSlide(
    pres,
    "I Do",
    C.PRIMARY,
    "Source 1: A Poem from 1789",
    [
      "Who is speaking in this poem?",
      "What has happened to his family?",
      "What is the word 'weep' telling us?",
      "Where does he sleep?",
    ],
    NOTES_SOURCE1_IDO,
    FOOTER,
    (s, lg) => {
      const rX = lg.rightX;
      const rW = lg.rightW;
      const topY = lg.panelTopPadded;

      drawPoemSourceCard(s, rX, topY, rW, 3.3, {
        sourceType: "PRIMARY SOURCE - POEM",
        title: "The Chimney Sweeper",
        citation: "William Blake, 1789 (first stanza)",
        strip: C.PRIMARY,
        lines: [
          "When my mother died I was very young,",
          "And my father sold me while yet my tongue",
          "Could scarcely cry 'weep! weep! weep! weep!'",
          "So your chimneys I sweep, and in soot I sleep.",
        ],
      });
    }
  );

  // -- Slide 6 / 6a: CFU hinge with reveal --
  const cfuQText = "Write the ONE word from Blake's poem that best shows the boy was treated as a thing, not a person.";

  function buildCfuBase() {
    const s = pres.addSlide();
    addTopBar(s, C.ALERT);
    addBadge(s, "CFU", { color: C.ALERT });
    addTitle(s, "Which Word Tells Us Most?", { color: C.ALERT });

    // Technique pill
    s.addShape("roundRect", {
      x: 0.5, y: CONTENT_TOP, w: 2.8, h: 0.40, rectRadius: 0.08,
      fill: { color: C.ALERT },
    });
    s.addText("Show Me Boards", {
      x: 0.5, y: CONTENT_TOP, w: 2.8, h: 0.40,
      fontSize: 12, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    // Answer bar anchored near SAFE_BOTTOM; work cards upward from there
    const aH = 0.55;
    const aY = SAFE_BOTTOM - aH - 0.05;
    const rH = 0.85;
    const rY = aY - 0.2 - rH;
    const qY = CONTENT_TOP + 0.56;
    const qH = rY - 0.25 - qY;

    // Question card
    addCard(s, 0.5, qY, 9, qH, { strip: C.ALERT, fill: C.WHITE });
    s.addText(cfuQText, {
      x: 0.75, y: qY + 0.18, w: 8.5, h: qH - 0.36,
      fontSize: 20, fontFace: FONT_B, color: C.CHARCOAL, valign: "top", margin: 0,
    });

    // Compact source reminder card - just the key line, no wrap risk
    addCard(s, 0.5, rY, 9, rH, { strip: C.PRIMARY, fill: "FAF5E8" });
    s.addText("Source 1 reminder  -  use your printed source card for the full text", {
      x: 0.75, y: rY + 0.08, w: 8.25, h: 0.24,
      fontSize: 10.5, fontFace: FONT_B, color: C.MUTED, bold: true, margin: 0,
    });
    s.addText("'my father sold me ... So your chimneys I sweep, and in soot I sleep.'", {
      x: 0.75, y: rY + 0.34, w: 8.25, h: rH - 0.4,
      fontSize: 14.5, fontFace: "Georgia", color: "2C2416", italic: true, margin: 0,
      valign: "middle",
    });

    // Answer bar placeholder
    s.addShape("roundRect", {
      x: 0.5, y: aY, w: 9, h: aH, rectRadius: 0.1,
      fill: { color: C.BG_LIGHT },
      line: { color: C.MUTED, width: 0.6, dashType: "dash" },
    });
    s.addText("Answer revealed next click", {
      x: 0.5, y: aY, w: 9, h: aH,
      fontSize: 11, fontFace: FONT_B, color: C.MUTED, italic: true,
      align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_CFU_Q);
    return s;
  }

  withReveal(
    buildCfuBase,
    (s) => {
      // Overlay answer band - same position as the dashed placeholder
      const aH = 0.55;
      const aY = SAFE_BOTTOM - aH - 0.05;
      s.addShape("roundRect", {
        x: 0.5, y: aY, w: 9, h: aH, rectRadius: 0.1,
        fill: { color: C.SUCCESS },
      });
      s.addText("sold  -  treated as property to be bought and handed over", {
        x: 0.5, y: aY, w: 9, h: aH,
        fontSize: 16, fontFace: FONT_B, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      s.addNotes(NOTES_CFU_A);
    }
  );

  // -- Slide 7 / 7a: We Do - Source 2 with reveal --
  const s2Items = [
    { label: "1", text: "Some climbing boys were apprenticed as young as four years old", color: C.SECONDARY },
    { label: "2", text: "Chimneys were narrow - sometimes only 9 inches wide. Boys were pushed up them", color: C.SECONDARY },
    { label: "3", text: "Common dangers: burns, suffocation, and a cancer caused by soot", color: C.SECONDARY },
    { label: "4", text: "The Chimney Sweepers Act 1788 set a minimum apprentice age of eight", color: C.SECONDARY },
  ];

  function buildWeDoBase() {
    const s = pres.addSlide();
    addTopBar(s, C.SECONDARY);
    addBadge(s, "We Do", { color: C.SECONDARY });
    addTitle(s, "Source 2: What Historians Found");

    // Left prompt card
    addInstructionCard(s, [
      { role: "header", text: "Our Task" },
      { role: "body", text: "Read Source 2 with me" },
      { role: "body", text: "Find ONE detail that matches something in Blake's poem" },
      { role: "body", text: "Be ready to share with a partner" },
    ], {
      x: 0.5, y: CONTENT_TOP, w: 3.25, h: SAFE_BOTTOM - CONTENT_TOP,
      strip: C.SECONDARY, fill: C.WHITE,
      headerColor: C.SECONDARY,
    });

    // Right: historical background source card
    drawBackgroundSourceCard(s, 3.95, CONTENT_TOP, 5.55, SAFE_BOTTOM - CONTENT_TOP, s2Items, {
      title: "Life for London's Climbing Boys",
      citation: "Historical Background card",
      strip: C.SECONDARY,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_SOURCE2_QWE);
    return s;
  }

  withReveal(
    buildWeDoBase,
    (s) => {
      // Overlay: highlight the numbered rows with coloured outlines
      const x = 3.95;
      const cardTop = CONTENT_TOP;
      const bodyY = cardTop + 0.95;
      const rowH = 0.48;
      const highlight = C.ALERT;

      s2Items.forEach((_, i) => {
        const ry = bodyY + i * rowH;
        s.addShape("roundRect", {
          x: x + 0.25, y: ry + 0.02, w: 5.05, h: rowH - 0.04,
          rectRadius: 0.06,
          fill: { type: "none" },
          line: { color: highlight, width: 1.8 },
        });
      });

      // Full-size cover over the left "Our Task" card so original text is hidden
      const leftH = SAFE_BOTTOM - CONTENT_TOP;
      s.addShape("roundRect", {
        x: 0.5, y: CONTENT_TOP, w: 3.25, h: leftH,
        rectRadius: 0.1, fill: { color: C.SUCCESS },
      });
      s.addText([
        { text: "Four key details found", options: { fontSize: 18, bold: true, breakLine: true } },
        { text: "", options: { fontSize: 8, breakLine: true } },
        { text: "Each outlined detail tells us something specific about life for climbing boys", options: { fontSize: 13, italic: true } },
      ], {
        x: 0.65, y: CONTENT_TOP + 0.35, w: 2.95, h: leftH - 0.7,
        fontFace: FONT_B, color: C.WHITE,
        align: "center", valign: "middle", margin: 0,
      });
      s.addNotes(NOTES_SOURCE2_AWE);
    }
  );

  // -- Slide 8: You Do --
  contentSlide(
    pres,
    "You Do",
    C.ACCENT,
    "Use Both Sources to Answer",
    [
      "Question: What was life like for chimney sweep boys in 1780s London?",
      "First: Write ONE detail from Blake in the left column",
      "Next: Write ONE detail from Source 2 in the right column",
      "Then: Write one sentence using evidence from BOTH",
    ],
    NOTES_YOUDO,
    FOOTER,
    (s, lg) => {
      const rX = lg.rightX;
      const rW = lg.rightW;
      const topY = lg.panelTopPadded;

      // Mini organiser preview
      addCard(s, rX, topY, rW, 3.35, { strip: C.ACCENT, fill: C.WHITE });

      s.addText("Organiser Preview", {
        x: rX + 0.15, y: topY + 0.08, w: rW - 0.3, h: 0.28,
        fontSize: 12, fontFace: FONT_H, color: C.ACCENT, bold: true,
        align: "center", margin: 0,
      });

      // Two-column boxes
      const colY = topY + 0.45;
      const colH = 1.5;
      const colW = (rW - 0.45) / 2;

      // Column 1 - Blake
      s.addShape("roundRect", {
        x: rX + 0.15, y: colY, w: colW, h: 0.3, rectRadius: 0.05,
        fill: { color: C.PRIMARY },
      });
      s.addText("From Blake", {
        x: rX + 0.15, y: colY, w: colW, h: 0.3,
        fontSize: 11, fontFace: FONT_B, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      s.addShape("rect", {
        x: rX + 0.15, y: colY + 0.3, w: colW, h: colH - 0.3,
        fill: { color: "FAF5E8" },
        line: { color: C.MUTED, width: 0.5 },
      });

      // Column 2 - Source 2
      s.addShape("roundRect", {
        x: rX + 0.3 + colW, y: colY, w: colW, h: 0.3, rectRadius: 0.05,
        fill: { color: C.SECONDARY },
      });
      s.addText("From Source 2", {
        x: rX + 0.3 + colW, y: colY, w: colW, h: 0.3,
        fontSize: 11, fontFace: FONT_B, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      s.addShape("rect", {
        x: rX + 0.3 + colW, y: colY + 0.3, w: colW, h: colH - 0.3,
        fill: { color: "FAF5E8" },
        line: { color: C.MUTED, width: 0.5 },
      });

      // Evidence sentence box
      const eY = colY + colH + 0.15;
      const eH = 0.95;
      s.addShape("roundRect", {
        x: rX + 0.15, y: eY, w: rW - 0.3, h: 0.3, rectRadius: 0.05,
        fill: { color: C.ACCENT },
      });
      s.addText("One Sentence Using Both", {
        x: rX + 0.15, y: eY, w: rW - 0.3, h: 0.3,
        fontSize: 11, fontFace: FONT_B, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      s.addShape("rect", {
        x: rX + 0.15, y: eY + 0.3, w: rW - 0.3, h: eH - 0.3,
        fill: { color: "FAF5E8" },
        line: { color: C.MUTED, width: 0.5 },
      });
    }
  );

  // -- Slide 9: Exit Ticket --
  cfuSlide(
    pres,
    "Exit",
    "One-Sentence Claim with Evidence",
    "Exit Ticket",
    "Finish this sentence on your exit ticket:\n\n'Life for chimney sweep boys in 1780s London was ___ because ___.'\n\nYour 'because' must name a detail from Source 1 or Source 2.",
    NOTES_EXIT,
    FOOTER
  );

  // -- Slide 10: Closing --
  closingSlide(
    pres,
    "Why do historians use more than one source when they study a topic?",
    [
      "Primary sources give us voices from the time itself",
      "Climbing boys were apprenticed as young as four and did dangerous work",
      "Comparing two sources gives us a fuller picture than one alone",
    ],
    NOTES_CLOSING
  );

  // -- Slide 11: Resources --
  addResourceSlide(pres, RESOURCE_ITEMS, T, FOOTER, NOTES_RESOURCES);

  // -- Write PPTX --
  fs.mkdirSync(LESSON_FOLDER, { recursive: true });
  await pres.writeFile({ fileName: path.join(LESSON_FOLDER, PPTX_NAME) });
  console.log("PPTX written to", path.join(LESSON_FOLDER, PPTX_NAME));

  // ===============================================================
  // PDF 1: Source Card Pack
  // ===============================================================
  {
    const doc = createPdf({ title: "Source Card Pack" });
    let y = addPdfHeader(doc, "Session 1 Source Card Pack", {
      subtitle: "Two sources on chimney sweep boys in 1780s London",
      color: C.PRIMARY,
      lessonInfo: "Grade 5/6 Inquiry - Life in 1780s London",
    });

    // Source 1 heading
    y = addSectionHeading(doc, "Source 1: Primary Source - Poem", y, { color: C.PRIMARY });
    y = addBodyText(doc, "The Chimney Sweeper - William Blake, 1789 (first stanza)", y, {
      fontSize: 10.5, italic: true, color: "555555",
    });

    // Aged-paper box for the poem
    const boxX = PAGE.MARGIN + 20;
    const boxW = PAGE.CONTENT_W - 40;
    const boxH = 130;
    doc.save();
    doc.roundedRect(boxX, y, boxW, boxH, 6).fill("#FAF5E8");
    doc.roundedRect(boxX, y, boxW, boxH, 6).lineWidth(0.8).strokeColor(hex(C.MUTED)).stroke();
    doc.restore();

    const poem = [
      "When my mother died I was very young,",
      "And my father sold me while yet my tongue",
      "Could scarcely cry 'weep! weep! weep! weep!'",
      "So your chimneys I sweep, and in soot I sleep.",
    ];
    doc.fontSize(13).font("Sans").fillColor("#2C2416");
    poem.forEach((ln, i) => {
      doc.text(ln, boxX + 20, y + 20 + i * 22, { width: boxW - 40 });
    });
    y += boxH + 12;

    y = addTipBox(doc, "Notice: who is speaking, what has happened, the street cry 'weep', and where he sleeps. Circle one word that tells us the most.", y, { color: C.PRIMARY });

    // Source 2 heading
    y = addSectionHeading(doc, "Source 2: Historical Background Card", y, { color: C.SECONDARY });
    y = addBodyText(doc, "Adapted from documented history of 18th century London.", y, {
      fontSize: 10.5, italic: true, color: "555555",
    });

    // Source 2 box with four numbered details
    const bg2X = PAGE.MARGIN + 20;
    const bg2W = PAGE.CONTENT_W - 40;
    const s2Facts = [
      "Some climbing boys were apprenticed as young as four years old.",
      "Chimneys were narrow - sometimes only 9 inches wide. Boys were pushed up them to scrape soot.",
      "Common dangers included burns from hot flues, suffocation, falls, and a cancer caused by soot (first documented by Percivall Pott in 1775).",
      "The Chimney Sweepers Act of 1788 set a minimum apprentice age of eight. Enforcement was weak.",
    ];

    // Measure box height
    doc.fontSize(10.5).font("Sans");
    let bg2H = 14;
    const lineH = 14;
    const rowPad = 10;
    const rowHeights = s2Facts.map((f) => {
      const h = doc.heightOfString(f, { width: bg2W - 70 });
      return Math.max(h + rowPad, 34);
    });
    bg2H += rowHeights.reduce((a, b) => a + b, 0) + 12;

    doc.save();
    doc.roundedRect(bg2X, y, bg2W, bg2H, 6).fill("#FAF5E8");
    doc.roundedRect(bg2X, y, bg2W, bg2H, 6).lineWidth(0.8).strokeColor(hex(C.MUTED)).stroke();
    doc.restore();

    let rowY = y + 10;
    s2Facts.forEach((fact, i) => {
      // Number badge
      doc.save();
      doc.circle(bg2X + 24, rowY + 12, 11).fill(hex(C.SECONDARY));
      doc.fontSize(11).font("Sans-Bold").fillColor("#FFFFFF");
      doc.text(String(i + 1), bg2X + 14, rowY + 6, { width: 20, align: "center" });
      doc.restore();
      // Text
      doc.fontSize(10.5).font("Sans").fillColor("#2C2416");
      doc.text(fact, bg2X + 44, rowY + 6, { width: bg2W - 60 });
      rowY += rowHeights[i];
    });
    y = y + bg2H + 10;

    y = addTipBox(doc, "Highlight one detail from each source that gives the strongest evidence for what life was like.", y, { color: C.SECONDARY });

    addPdfFooter(doc, "Inquiry | Grade 5/6 | Session 1 - Source Cards");
    const outPath = path.join(RES_DIR, "Session 1 Source Card Pack.pdf");
    await writePdf(doc, outPath);
    console.log("PDF written:", outPath);
  }

  // ===============================================================
  // PDF 2: Source Analysis Organiser
  // ===============================================================
  {
    const doc = createPdf({ title: "Source Analysis Organiser" });
    let y = addPdfHeader(doc, "Session 1 Source Analysis Organiser", {
      subtitle: "Use both sources to answer one big question",
      color: C.ACCENT,
      lessonInfo: "Grade 5/6 Inquiry - Life in 1780s London",
    });

    // Big question box
    y = addTipBox(doc, "Big question: What was life like for chimney sweep boys in 1780s London?", y, { color: C.PRIMARY });

    y = addSectionHeading(doc, "Step 1: Evidence From Each Source", y, { color: C.ACCENT });
    y = addBodyText(doc, "Write ONE clear detail from each source.", y, { fontSize: 10.5, italic: true });

    y = addTwoColumnOrganiser(doc, "From Blake's poem (1789)", "From the Historical Background", y, {
      color: C.ACCENT,
      rows: 3,
      rowH: 48,
    });

    y = addSectionHeading(doc, "Step 2: One Sentence Using Both Sources", y, { color: C.ACCENT });
    y = addBodyText(doc, "Write one sentence that uses evidence from BOTH sources to answer the big question. Name the sources.", y, { fontSize: 10.5, italic: true });

    y = addLinedArea(doc, y, 4, { lineSpacing: 24 });

    y += 6;
    y = addSectionHeading(doc, "Step 3: What Does Each Source Add?", y, { color: C.ACCENT });
    y = addBodyText(doc, "Some details appear in both sources. Some appear in only one. Write one of each.", y, { fontSize: 10.5, italic: true });

    y = addTwoColumnOrganiser(doc, "Detail in BOTH sources", "Detail in ONLY ONE source (name it)", y, {
      color: C.ACCENT,
      rows: 2,
      rowH: 42,
    });

    addPdfFooter(doc, "Inquiry | Grade 5/6 | Session 1 - Source Analysis");
    const outPath = path.join(RES_DIR, "Session 1 Source Analysis Organiser.pdf");
    await writePdf(doc, outPath);
    console.log("PDF written:", outPath);
  }

  // ===============================================================
  // PDF 3: Exit Ticket
  // ===============================================================
  {
    const doc = createPdf({ title: "Exit Ticket" });
    let y = addPdfHeader(doc, "Session 1 Exit Ticket", {
      subtitle: "One-sentence claim with evidence",
      color: C.ALERT,
      lessonInfo: "Grade 5/6 Inquiry - Life in 1780s London",
    });

    y = addTipBox(doc, "You are showing me you can use a source to back up a claim. Keep it to one clear sentence.", y, { color: C.ALERT });

    y = addSectionHeading(doc, "Complete This Sentence", y, { color: C.ALERT });

    // Frame box
    const frameX = PAGE.MARGIN;
    const frameW = PAGE.CONTENT_W;
    const frameH = 130;
    doc.save();
    doc.roundedRect(frameX, y, frameW, frameH, 6).fill("#FAF5E8");
    doc.roundedRect(frameX, y, frameW, frameH, 6).lineWidth(0.8).strokeColor(hex(C.MUTED)).stroke();
    doc.restore();

    doc.fontSize(14).font("Sans-Bold").fillColor(hex(C.ALERT));
    doc.text("Life for chimney sweep boys in 1780s London was", frameX + 18, y + 16, {
      width: frameW - 36,
    });
    // First blank
    doc.save();
    doc.moveTo(frameX + 18, y + 50).lineTo(frameX + frameW - 18, y + 50)
      .strokeColor("#000000").lineWidth(0.9).stroke();
    doc.restore();

    doc.fontSize(14).font("Sans-Bold").fillColor(hex(C.ALERT));
    doc.text("because", frameX + 18, y + 58);
    // Second and third lines
    doc.save();
    doc.moveTo(frameX + 18, y + 92).lineTo(frameX + frameW - 18, y + 92)
      .strokeColor("#000000").lineWidth(0.9).stroke();
    doc.moveTo(frameX + 18, y + 116).lineTo(frameX + frameW - 18, y + 116)
      .strokeColor("#000000").lineWidth(0.9).stroke();
    doc.restore();

    y += frameH + 16;

    y = addSectionHeading(doc, "Name the Source You Used", y, { color: C.ALERT });
    doc.fontSize(11).font("Sans").fillColor("#000000");
    doc.text("Tick one:", PAGE.MARGIN, y);
    y = doc.y + 8;

    const opts = [
      "Source 1 - Blake's poem (1789)",
      "Source 2 - Historical Background card",
      "Both sources",
    ];
    opts.forEach((o) => {
      doc.save();
      doc.rect(PAGE.MARGIN + 6, y + 2, 12, 12).lineWidth(0.9).strokeColor("#000000").stroke();
      doc.restore();
      doc.fontSize(11).font("Sans").fillColor("#000000");
      doc.text(o, PAGE.MARGIN + 26, y);
      y = doc.y + 6;
    });

    addPdfFooter(doc, "Inquiry | Grade 5/6 | Session 1 - Exit Ticket");
    const outPath = path.join(RES_DIR, "Session 1 Exit Ticket.pdf");
    await writePdf(doc, outPath);
    console.log("PDF written:", outPath);
  }

  console.log("\nBuild complete!");
  console.log("Output folder:", LESSON_FOLDER);
}

build().catch((err) => { console.error(err); process.exit(1); });
