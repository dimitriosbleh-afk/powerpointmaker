"use strict";

// Respectful Relationships - What are Character Strengths?
// Grade 5/6 Wellbeing - 20-minute session
// Activity 1: Identifying character strengths in self and others

const path = require("path");
const fs = require("fs");
const PptxGenJS = require("pptxgenjs");
const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf,
  addPdfHeader, addSectionHeading, addBodyText, addTipBox,
  addPdfFooter, addLinedArea,
  addResourceSlide, makeSessionResource,
  getSessionResourceFolder,
  PAGE, hex,
} = require("../themes/pdf_helpers");

// Theme
const T = createTheme("wellbeing", "grade56", weekToVariant(1));
const {
  C, FONT_H, FONT_B,
  titleSlide, liSlide, contentSlide, closingSlide,
  pairShareSlide,
  addTopBar, addBadge, addTitle, addCard, addFooter,
  CONTENT_TOP, SAFE_BOTTOM,
} = T;

// Output paths
const UNIT = "RR_Character_Strengths";
const LESSON_FOLDER = path.join(__dirname, "..", "output", UNIT);
const PPTX_NAME = "What are Character Strengths.pptx";
const FOOTER = "Respectful Relationships | Grade 5/6 Wellbeing";
const SESSION = 1;
const RES_DIR = path.join(LESSON_FOLDER, getSessionResourceFolder(SESSION));

// Resources
const HANDOUT_RESOURCE = makeSessionResource(
  SESSION,
  "Character Strengths Handout",
  "Reference list of 16 character strengths with kid-friendly definitions and tick boxes."
);
const SKETCH_RESOURCE = makeSessionResource(
  SESSION,
  "My Kind Action Sketch",
  "Sketch planner for drawing and labelling a kind or caring action."
);
const RESOURCE_ITEMS = [HANDOUT_RESOURCE, SKETCH_RESOURCE];

fs.mkdirSync(RES_DIR, { recursive: true });

// ============================================================
// Teacher Notes
// ============================================================

const NOTES_TITLE = [
  "SAY:",
  "- Today we're thinking about the positive qualities we value in ourselves and others.",
  "- We call these qualities character strengths -- things like kindness, courage, and humour.",
  "",
  "DO:",
  "- Display slide as students settle.",
  "- Have the Character Strengths Handout and My Kind Action Sketch ready to distribute.",
  "",
  "TEACHER NOTES:",
  "This is a short 20-minute session introducing character strengths. Students draw a sketch of a good action, share, and then meet the concept of character strengths as distinct from talents or skills.",
  "",
  "WATCH FOR:",
  "- Students transitioning calmly from the previous subject.",
  "",
  "[General: Title | VTLM 2.0: Establishing Purpose]",
].join("\n");

const NOTES_LI = [
  "SAY:",
  "- Read the Learning Intention from the slide.",
  "- Read each Success Criterion aloud -- we'll come back to these at the end.",
  "",
  "DO:",
  "- Point to each SC as you read it.",
  "- Leave the slide visible for 15-20 seconds.",
  "",
  "TEACHER NOTES:",
  "SC1 is the ultra-achievable floor (everyone can name a kind action). SC2 is the core target (distinguishing talent from character). SC3 extends into noticing and naming.",
  "",
  "WATCH FOR:",
  "- Students unfamiliar with the phrase 'character strength' -- this will be unpacked on slide 5.",
  "",
  "[General: LI/SC | VTLM 2.0: Learning Intentions]",
].join("\n");

const NOTES_LAUNCH = [
  "SAY:",
  "- Think of a time someone did something kind, fair, or caring -- for you or for someone else.",
  "- It doesn't have to be big. A small act of helping counts.",
  "- In a moment you'll draw a quick sketch of that good action and label what's happening.",
  "- This isn't an art lesson -- stick figures are perfect.",
  "",
  "DO:",
  "- Distribute the My Kind Action Sketch planner.",
  "- Give students about 3 minutes to sketch and label.",
  "- Circulate and prompt hesitant students with examples (holding a door, sharing lunch, standing up for someone).",
  "",
  "TEACHER NOTES:",
  "The sketch is the hook. Concrete examples anchor the abstract idea of character strengths that follows. Keep pacing brisk -- this is a 3-minute draw, not a polished piece of work.",
  "",
  "WATCH FOR:",
  "- Students who can't think of an example -- prompt them: 'Has someone ever helped you when you were upset?'",
  "- Students drawing themselves doing the kind action -- that's fine; self-recognition matters too.",
  "",
  "[General: I Do / You Do launch | VTLM 2.0: Explicit Explanation]",
].join("\n");

const NOTES_SHARE = [
  "SAY:",
  "- Who'd like to share their sketch? Tell us what's happening in the picture.",
  "- I'm going to write up the key action words on the board as we go -- things like sharing, helping, listening.",
  "- Ask: What does it take for someone to do this? What kind of person does this? [kindness, fairness, courage, caring, honesty, patience...]",
  "",
  "DO:",
  "- Invite 4-6 students to share -- keep each share to about 20 seconds.",
  "- Record the key action and the quality behind it on the board in two columns: ACTION | QUALITY.",
  "- Validate every contribution: 'That's a great example of...'",
  "",
  "CFU CHECKPOINT:",
  "Technique: Cold Call with Think-Pair-Share.",
  "Script:",
  "- Say: Turn to your partner. Share one word that describes the quality behind your sketch. 20 seconds.",
  "- Cold call 3 pairs. Listen for quality words (kind, fair, brave, honest) rather than action words (sharing, helping).",
  "- Scan for: students distinguishing the QUALITY from the ACTION.",
  "PROCEED: If most pairs offer a quality word (kind, fair, brave), move on.",
  "PIVOT: If students only restate the action ('sharing', 'helping'), model the shift: 'Sharing is the action -- what's the quality inside the person who shares? [kindness, generosity].' Re-ask one pair.",
  "",
  "TEACHER NOTES:",
  "This slide captures the raw language students will formalise as 'character strengths' on the next slide. The board notes become the bridge to the definition.",
  "",
  "ENABLING & EXTENDING:",
  "ENABLING PROMPT:",
  "- Task: Students who struggle to name a quality choose from a 3-word menu on the board: KIND, BRAVE, FAIR. They point to the one that matches their sketch.",
  "- Extra Notes: Offer a sentence stem: 'The person in my sketch was being ___.'",
  "EXTENDING PROMPT:",
  "- Task: Students identify TWO strengths the person in their sketch might have shown at once (e.g. kindness + courage when standing up for a friend).",
  "",
  "WATCH FOR:",
  "- Students saying 'nice' for everything -- push for more precise words.",
  "- Sensitivity: a student may share a moment involving someone close to them -- acknowledge warmly without probing.",
  "",
  "[General: We Do | VTLM 2.0: Scaffold Practice]",
].join("\n");

const NOTES_DEFINITION = [
  "SAY:",
  "- The people in your sketches were showing character strengths.",
  "- Read from slide: 'Your character is what makes you who you are.'",
  "- Character strengths are different from talents and skills.",
  "- A talent is something you're good at -- like football, or drawing, or maths.",
  "- A character strength is part of who you are -- like kindness, courage, humour, or persistence.",
  "- Talents help you perform. Character strengths help you make the most of your life and handle tough moments.",
  "",
  "DO:",
  "- Point to the 'Talent vs Character' visual as you explain.",
  "- Give a quick personal example: 'I might have the talent of running fast, but my character strength is pushing through when I'm tired -- that's persistence.'",
  "",
  "TEACHER NOTES:",
  "Coaching point from the source activity: distinguishing skills/talents from character strengths is genuinely hard for this age group. Expect follow-up confusion and revisit with examples.",
  "",
  "MISCONCEPTIONS:",
  "- Misconception: A talent and a character strength are the same thing.",
  "  Why: Students hear both words used as 'something good about a person.'",
  "  Impact: Students will praise talents (running, singing) as if they were character, and may overlook quieter strengths like kindness or honesty.",
  "  Quick correction: Use the contrast: 'Being good at football is a talent. Being a good teammate is a character strength.'",
  "",
  "WATCH FOR:",
  "- Students offering hobbies or subjects (gaming, drawing) as character strengths -- redirect gently.",
  "- Students connecting the definition to someone they know -- this is a positive sign of transfer.",
  "",
  "[General: I Do | VTLM 2.0: Explicit Explanation]",
].join("\n");

const NOTES_HANDOUT_REVIEW = [
  "SAY:",
  "- On your Character Strengths Handout you'll see 16 strengths with short descriptions.",
  "- Look back at the sketches we shared. Which strengths did we see? Tick them on your handout.",
  "- Ask: Which strengths were in our pictures? [Expect: kindness, fairness, courage, caring, honesty.]",
  "- Were any strengths missed? Which one would you add?",
  "- Ask 1-2 students to quickly sketch an action that shows a missing strength on the board.",
  "",
  "DO:",
  "- Distribute the Character Strengths Handout if not already handed out.",
  "- Give students 2 minutes to tick the strengths that appeared in the class sketches.",
  "- Call on students to name the strengths they ticked; record them on the board.",
  "- Invite one or two volunteers to add a quick sketch for a strength that was missing.",
  "",
  "CFU CHECKPOINT:",
  "Technique: Thumbs Up / Sideways / Down.",
  "Script:",
  "- Say: I'm going to read two examples. Thumbs up if it's a character strength. Thumbs down if it's a talent.",
  "- Example 1: 'Sam is really good at drawing.' [Talent - thumbs down.]",
  "- Example 2: 'Sam stayed with a friend who was upset at lunch.' [Character strength - thumbs up.]",
  "- Scan thumbs. Look for at least 80% of the class getting both correct.",
  "PROCEED: If most students separate talent from character, move to the closing.",
  "PIVOT: If students confuse the two, use a class example: 'Think about the sketches we shared -- none of them were about being the best at a skill. They were about how a person chose to treat someone else. That's the difference.' Re-ask with one more pair.",
  "",
  "TEACHER NOTES:",
  "The handout is a shared reference. Ticking strengths makes the abstract list feel concrete and linked to the students' own examples. Expect students to keep confusing some strengths with talents -- this is normal and will need revisiting across the unit.",
  "",
  "ENABLING & EXTENDING:",
  "ENABLING PROMPT:",
  "- Task: Students work with a partner and focus on ticking just 5 strengths from the handout -- courage, kindness, fairness, honesty, and humour -- using the class sketches.",
  "- Extra Notes: Offer a sentence stem: 'I saw ___ in the sketch about ___.'",
  "EXTENDING PROMPT:",
  "- Task: Students identify one character strength NOT in the class sketches and explain on the back of their handout why it matters and when someone might show it.",
  "",
  "WATCH FOR:",
  "- Students ticking every strength -- prompt them to justify with a specific sketch.",
  "- Students getting stuck on definitions -- pair them with a confident partner.",
  "",
  "[General: We Do | VTLM 2.0: Scaffold Practice]",
].join("\n");

const NOTES_CLOSING = [
  "SAY:",
  "- Let's look back at our success criteria. Give me a thumbs up, sideways, or down for each one.",
  "- SC1: I can name a kind action someone has done.",
  "- SC2: I can explain the difference between a talent and a character strength.",
  "- SC3: I can name character strengths I admire in others.",
  "- Turn & Talk: Which character strength do you admire most in a family member or friend? Tell your partner one example.",
  "- Ask a volunteer to summarise: what is a character strength?",
  "",
  "DO:",
  "- Read each SC aloud; pause for thumbs.",
  "- Allow 60 seconds for Turn & Talk.",
  "- Challenge students to notice and compliment a strength they see in someone this week.",
  "",
  "TEACHER NOTES:",
  "The closing connects back to the lesson's core purpose: noticing and valuing character in others. The 'spot it and say it' challenge extends the learning beyond the session.",
  "",
  "WATCH FOR:",
  "- Students showing sideways or down on SC2 -- note for a quick revisit in the next session.",
  "- Students who struggle to name someone they admire -- offer the example of a teacher or book character.",
  "",
  "[General: Closing | VTLM 2.0: Review and Reflect]",
].join("\n");

const NOTES_RESOURCES = [
  "SAY:",
  "- These are the two handouts for this session.",
  "",
  "DO:",
  "- Print the Character Strengths Handout and My Kind Action Sketch before the session.",
  "- Distribute the Sketch planner at the launch; the Character Strengths Handout after the share.",
  "",
  "TEACHER NOTES:",
  "Both resources are in the resources folder linked from this slide. Keep the Character Strengths Handout -- students will revisit it in later Respectful Relationships sessions.",
  "",
  "WATCH FOR:",
  "- Ensure every student has a sketch planner before the launch.",
  "",
  "[General: Resources | VTLM 2.0: Preparation]",
].join("\n");

// ============================================================
// Character strengths list (used on slides AND in PDF)
// ============================================================

const STRENGTHS = [
  { name: "Kindness",       desc: "Caring about others and helping when it matters." },
  { name: "Courage",        desc: "Doing the right thing even when it's scary or hard." },
  { name: "Honesty",        desc: "Telling the truth and being real about who you are." },
  { name: "Fairness",       desc: "Treating everyone evenly and giving people a fair go." },
  { name: "Humour",         desc: "Bringing lightness, smiles, and laughter to others." },
  { name: "Persistence",    desc: "Keeping going when something is difficult." },
  { name: "Gratitude",      desc: "Noticing good things and saying thank you." },
  { name: "Curiosity",      desc: "Wanting to learn, explore, and ask questions." },
  { name: "Teamwork",       desc: "Working well with others to get things done." },
  { name: "Leadership",     desc: "Helping a group move in a good direction." },
  { name: "Forgiveness",    desc: "Letting go of hurt instead of holding a grudge." },
  { name: "Self-Control",   desc: "Managing your feelings and actions carefully." },
  { name: "Creativity",     desc: "Coming up with new ideas or different ways of doing things." },
  { name: "Love of Learning", desc: "Enjoying finding out new things." },
  { name: "Hope",           desc: "Expecting good things and working toward them." },
  { name: "Caring",         desc: "Looking out for how others feel and what they need." },
];

// ============================================================
// Build function
// ============================================================

async function build() {
  const pres = new PptxGenJS();
  pres.layout = "LAYOUT_16x9";

  // -- Slide 1: Title --
  titleSlide(
    pres,
    "What Are\nCharacter Strengths?",
    "Respectful Relationships",
    "Grade 5/6 Wellbeing  |  Session 1",
    NOTES_TITLE
  );

  addResourceSlide(pres, RESOURCE_ITEMS, T, FOOTER, NOTES_RESOURCES);

  // -- Slide 2: LI / SC --
  liSlide(
    pres,
    ["We are learning to identify character strengths in ourselves and others, and how they are different from talents and skills"],
    [
      "I can name a kind action someone has done",
      "I can explain the difference between a talent and a character strength",
      "I can name character strengths I admire in others",
    ],
    NOTES_LI,
    FOOTER
  );

  // -- Slide 3: Launch / Sketch Activity (I Do instructions for the task) --
  contentSlide(
    pres,
    "I Do",
    C.PRIMARY,
    "Sketch a Good Action",
    [
      "First: Think of a kind, fair, or caring thing someone did",
      "Next: Draw a quick sketch -- stick figures are perfect",
      "Then: Label what's happening in your picture",
      "You have about 3 minutes",
    ],
    NOTES_LAUNCH,
    FOOTER,
    (s, lg) => {
      // Right column: a simple stick-figure example mock
      const rX = lg.rightX;
      const rW = lg.rightW;
      const topY = lg.panelTopPadded;
      const cardH = 3.3;

      addCard(s, rX, topY, rW, cardH, { strip: C.ACCENT, fill: C.WHITE });

      s.addText("Example Sketch", {
        x: rX + 0.15, y: topY + 0.1, w: rW - 0.3, h: 0.3,
        fontSize: 13, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0, align: "center",
      });

      // Simple stick-figure scene (two figures, one handing something to the other)
      const sceneY = topY + 0.55;
      const sceneH = 1.7;

      // Ground line
      s.addShape("line", {
        x: rX + 0.3, y: sceneY + sceneH - 0.15,
        w: rW - 0.6, h: 0,
        line: { color: C.MUTED, width: 1.2 },
      });

      // Figure A (left): head + body + arm extended
      const aX = rX + 0.7;
      const headR = 0.18;
      // Head
      s.addShape("roundRect", {
        x: aX, y: sceneY + 0.25, w: headR * 2, h: headR * 2,
        rectRadius: headR,
        line: { color: C.PRIMARY, width: 1.5 },
        fill: { color: C.WHITE },
      });
      // Body line
      s.addShape("line", {
        x: aX + headR, y: sceneY + 0.25 + headR * 2,
        w: 0, h: 0.55,
        line: { color: C.PRIMARY, width: 1.5 },
      });
      // Arm extended to right
      s.addShape("line", {
        x: aX + headR, y: sceneY + 0.82,
        w: 0.5, h: -0.05,
        line: { color: C.PRIMARY, width: 1.5 },
      });
      // Legs
      s.addShape("line", {
        x: aX + headR, y: sceneY + 1.15,
        w: -0.2, h: 0.35,
        line: { color: C.PRIMARY, width: 1.5 },
      });
      s.addShape("line", {
        x: aX + headR, y: sceneY + 1.15,
        w: 0.2, h: 0.35,
        line: { color: C.PRIMARY, width: 1.5 },
      });

      // Object being passed (a small heart / gift)
      s.addShape("roundRect", {
        x: aX + headR + 0.55, y: sceneY + 0.72, w: 0.22, h: 0.22,
        rectRadius: 0.04,
        fill: { color: C.ALERT },
        line: { color: C.ALERT, width: 1 },
      });

      // Figure B (right): head + body + arm reaching
      const bX = rX + rW - 1.0;
      s.addShape("roundRect", {
        x: bX, y: sceneY + 0.25, w: headR * 2, h: headR * 2,
        rectRadius: headR,
        line: { color: C.SECONDARY, width: 1.5 },
        fill: { color: C.WHITE },
      });
      s.addShape("line", {
        x: bX + headR, y: sceneY + 0.25 + headR * 2,
        w: 0, h: 0.55,
        line: { color: C.SECONDARY, width: 1.5 },
      });
      // Arm reaching left
      s.addShape("line", {
        x: bX + headR, y: sceneY + 0.82,
        w: -0.3, h: -0.05,
        line: { color: C.SECONDARY, width: 1.5 },
      });
      s.addShape("line", {
        x: bX + headR, y: sceneY + 1.15,
        w: -0.2, h: 0.35,
        line: { color: C.SECONDARY, width: 1.5 },
      });
      s.addShape("line", {
        x: bX + headR, y: sceneY + 1.15,
        w: 0.2, h: 0.35,
        line: { color: C.SECONDARY, width: 1.5 },
      });

      // Label
      s.addText("Sharing lunch with a new student", {
        x: rX + 0.15, y: topY + cardH - 0.55, w: rW - 0.3, h: 0.45,
        fontSize: 11, fontFace: FONT_B, color: C.CHARCOAL,
        italic: true, align: "center", valign: "middle", margin: 0,
      });
    }
  );

  // -- Slide 4: Share & Capture --
  pairShareSlide(
    pres,
    "Share & Spot the Quality",
    [
      "Share your sketch: What is happening in the picture?",
      "What kind of person does this? (e.g. kind, fair, brave)",
      "Listen: What words come up again and again?",
      "Turn & Talk: One word to describe the quality behind your sketch.",
    ],
    NOTES_SHARE,
    FOOTER
  );

  // -- Slide 5: Definition: Character vs Talent (I Do) --
  contentSlide(
    pres,
    "I Do",
    C.PRIMARY,
    "Character vs Talent",
    [
      "Your character is what makes you who you are",
      "Character strengths are attributes like:",
      "  courage, kindness, humour, persistence",
      "They help you make the most of your experiences",
      "They help you deal with challenges",
    ],
    NOTES_DEFINITION,
    FOOTER,
    (s, lg) => {
      // Right column: Talent vs Character comparison
      const rX = lg.rightX;
      const rW = lg.rightW;
      const topY = lg.panelTopPadded;
      const cardH = 3.4;
      const halfW = (rW - 0.1) / 2;

      // TALENT card (left side of right column)
      const tX = rX;
      addCard(s, tX, topY, halfW, cardH, { strip: C.SECONDARY, fill: C.WHITE });
      s.addText("TALENT", {
        x: tX + 0.1, y: topY + 0.1, w: halfW - 0.2, h: 0.35,
        fontSize: 13, fontFace: FONT_H, color: C.SECONDARY, bold: true, align: "center", margin: 0,
      });
      s.addText("Something you are good at", {
        x: tX + 0.1, y: topY + 0.52, w: halfW - 0.2, h: 0.45,
        fontSize: 11, fontFace: FONT_B, color: C.CHARCOAL, italic: true, align: "center", margin: 0,
      });
      s.addText([
        { text: "Football", options: { bullet: true, breakLine: true, fontSize: 12 } },
        { text: "Drawing",  options: { bullet: true, breakLine: true, fontSize: 12 } },
        { text: "Maths",    options: { bullet: true, breakLine: true, fontSize: 12 } },
        { text: "Singing",  options: { bullet: true, breakLine: false, fontSize: 12 } },
      ], {
        x: tX + 0.2, y: topY + 1.1, w: halfW - 0.3, h: cardH - 1.25,
        fontFace: FONT_B, color: C.CHARCOAL, valign: "top", margin: 0,
      });

      // CHARACTER STRENGTH card
      const cX = rX + halfW + 0.1;
      addCard(s, cX, topY, halfW, cardH, { strip: C.ACCENT, fill: C.WHITE });
      s.addText("CHARACTER", {
        x: cX + 0.1, y: topY + 0.1, w: halfW - 0.2, h: 0.35,
        fontSize: 13, fontFace: FONT_H, color: C.ACCENT, bold: true, align: "center", margin: 0,
      });
      s.addText("Who you are on the inside", {
        x: cX + 0.1, y: topY + 0.52, w: halfW - 0.2, h: 0.45,
        fontSize: 11, fontFace: FONT_B, color: C.CHARCOAL, italic: true, align: "center", margin: 0,
      });
      s.addText([
        { text: "Kindness",    options: { bullet: true, breakLine: true, fontSize: 12 } },
        { text: "Courage",     options: { bullet: true, breakLine: true, fontSize: 12 } },
        { text: "Humour",      options: { bullet: true, breakLine: true, fontSize: 12 } },
        { text: "Persistence", options: { bullet: true, breakLine: false, fontSize: 12 } },
      ], {
        x: cX + 0.2, y: topY + 1.1, w: halfW - 0.3, h: cardH - 1.25,
        fontFace: FONT_B, color: C.CHARCOAL, valign: "top", margin: 0,
      });
    }
  );

  // -- Slide 6: Handout Review (We Do) --
  {
    const s = pres.addSlide();
    addTopBar(s, C.SECONDARY);
    addBadge(s, "We Do", { color: C.SECONDARY });
    addTitle(s, "Which Strengths Did We See?");

    // Instruction strip
    s.addText("Tick the strengths on your handout that appeared in our class sketches. Which were missing?", {
      x: 0.5, y: CONTENT_TOP, w: 9, h: 0.35,
      fontSize: 12.5, fontFace: FONT_B, color: C.PRIMARY, italic: true, margin: 0,
    });

    // Grid of 16 strengths: 4 columns x 4 rows
    const gridY = CONTENT_TOP + 0.45;
    const gridH = SAFE_BOTTOM - gridY - 0.05;
    const cols = 4;
    const rows = 4;
    const gap = 0.12;
    const cellW = (9 - gap * (cols - 1)) / cols;
    const cellH = (gridH - gap * (rows - 1)) / rows;

    STRENGTHS.slice(0, 16).forEach((st, i) => {
      const r = Math.floor(i / cols);
      const c = i % cols;
      const cx = 0.5 + c * (cellW + gap);
      const cy = gridY + r * (cellH + gap);

      // Card
      s.addShape("roundRect", {
        x: cx, y: cy, w: cellW, h: cellH, rectRadius: 0.08,
        fill: { color: C.WHITE },
        line: { color: C.ACCENT, width: 1 },
      });
      // Top strip
      s.addShape("rect", {
        x: cx, y: cy, w: cellW, h: 0.08,
        fill: { color: C.ACCENT },
      });
      // Name
      s.addText(st.name, {
        x: cx + 0.1, y: cy + 0.14, w: cellW - 0.2, h: 0.30,
        fontSize: 12, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0,
      });
      // Description
      s.addText(st.desc, {
        x: cx + 0.1, y: cy + 0.46, w: cellW - 0.2, h: cellH - 0.56,
        fontSize: 9, fontFace: FONT_B, color: C.CHARCOAL, valign: "top", margin: 0,
      });
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_HANDOUT_REVIEW);
  }

  // -- Slide 7: Closing --
  closingSlide(
    pres,
    "Which character strength do you admire most in a family member or friend? Tell your partner one example of when you saw it.",
    [
      "A character strength is part of WHO you are -- not what you can DO",
      "Kindness, courage, honesty, fairness, humour are all character strengths",
      "This week: notice a strength in someone -- and tell them",
    ],
    NOTES_CLOSING
  );

  // -- Slide 8: Resources --


  // Write PPTX
  fs.mkdirSync(LESSON_FOLDER, { recursive: true });
  await pres.writeFile({ fileName: path.join(LESSON_FOLDER, PPTX_NAME) });
  console.log("PPTX written to", path.join(LESSON_FOLDER, PPTX_NAME));

  // ============================================================
  // Companion PDF 1: Character Strengths Handout
  // ============================================================
  {
    const doc = createPdf({ title: "Character Strengths Handout" });
    let y = addPdfHeader(doc, "Session 1 Character Strengths Handout", {
      subtitle: "Reference: 16 character strengths with descriptions",
      color: C.PRIMARY,
      lessonInfo: "Grade 5/6 Wellbeing - Respectful Relationships",
    });

    y = addTipBox(doc, "Your character is what makes you who you are. Character strengths are attributes like courage, kindness, humour, and persistence. They are different from talents and skills.", y, { color: C.ACCENT });

    y = addSectionHeading(doc, "The 16 Strengths - Tick the ones you saw in our class sketches", y, { color: C.PRIMARY });

    // Grid of strengths: 2 columns
    const colCount = 2;
    const colGap = 14;
    const colW = (PAGE.CONTENT_W - colGap * (colCount - 1)) / colCount;
    const rowH = 44;
    const boxSize = 12;

    STRENGTHS.forEach((st, i) => {
      const col = i % colCount;
      const row = Math.floor(i / colCount);
      const cx = PAGE.MARGIN + col * (colW + colGap);
      const cy = y + row * rowH;

      if (cy + rowH > PAGE.H - PAGE.MARGIN - 40) return;

      // Card border
      doc.save();
      doc.roundedRect(cx, cy, colW, rowH - 6, 4)
        .lineWidth(0.8).strokeColor(hex(C.MUTED)).stroke();
      doc.restore();

      // Tick box
      doc.save();
      doc.rect(cx + 8, cy + 10, boxSize, boxSize)
        .lineWidth(1).strokeColor(hex(C.PRIMARY)).stroke();
      doc.restore();

      // Name
      doc.fontSize(11).font("Sans-Bold").fillColor(hex(C.PRIMARY));
      doc.text(st.name, cx + 26, cy + 8, { width: colW - 34 });

      // Description
      doc.fontSize(9).font("Sans").fillColor(hex(C.CHARCOAL));
      doc.text(st.desc, cx + 26, cy + 22, { width: colW - 34 });
    });

    // Move y past the grid
    const gridRows = Math.ceil(STRENGTHS.length / colCount);
    y = y + gridRows * rowH + 10;

    // Reflection prompts
    if (y + 80 > PAGE.H - PAGE.MARGIN - 40) {
      doc.addPage();
      y = PAGE.MARGIN;
    }

    y = addSectionHeading(doc, "My Reflection", y, { color: C.ACCENT });
    y = addBodyText(doc, "A character strength I admire in a family member or friend:", y, { fontSize: 10 });
    y = addLinedArea(doc, y, 2, { lineSpacing: 22 });

    y = addBodyText(doc, "A character strength I want to grow in myself:", y, { fontSize: 10 });
    y = addLinedArea(doc, y, 2, { lineSpacing: 22 });

    addPdfFooter(doc, "Respectful Relationships | Grade 5/6 | Session 1");
    const handoutPath = path.join(RES_DIR, "Session 1 Character Strengths Handout.pdf");
    await writePdf(doc, handoutPath);
    console.log("PDF written:", handoutPath);
  }

  // ============================================================
  // Companion PDF 2: My Kind Action Sketch
  // ============================================================
  {
    const doc = createPdf({ title: "My Kind Action Sketch" });
    let y = addPdfHeader(doc, "Session 1 My Kind Action Sketch", {
      subtitle: "Sketch and label a kind, fair, or caring action",
      color: C.SECONDARY,
      lessonInfo: "Grade 5/6 Wellbeing - Respectful Relationships",
    });

    y = addTipBox(doc, "Think of a time someone did something kind, fair, or caring. It could be something you saw, something done for you, or something you did for someone else.", y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Step 1: Draw your sketch", y, { color: C.SECONDARY });
    y = addBodyText(doc, "Stick figures are perfect. You have about 3 minutes.", y, { fontSize: 10, italic: true });

    // Drawing space
    const drawH = 260;
    doc.save();
    doc.roundedRect(PAGE.MARGIN, y, PAGE.CONTENT_W, drawH, 4)
      .lineWidth(1).strokeColor(hex(C.MUTED)).stroke();
    doc.restore();
    y += drawH + 12;

    y = addSectionHeading(doc, "Step 2: Label your sketch", y, { color: C.SECONDARY });
    y = addBodyText(doc, "What is happening in the picture?", y, { fontSize: 10 });
    y = addLinedArea(doc, y, 2, { lineSpacing: 22 });

    y = addSectionHeading(doc, "Step 3: Name the character strength", y, { color: C.SECONDARY });
    y = addBodyText(doc, "What kind of person does this? (kind, fair, brave, honest, caring...)", y, { fontSize: 10 });
    y = addLinedArea(doc, y, 1, { lineSpacing: 22 });

    addPdfFooter(doc, "Respectful Relationships | Grade 5/6 | Session 1");
    const sketchPath = path.join(RES_DIR, "Session 1 My Kind Action Sketch.pdf");
    await writePdf(doc, sketchPath);
    console.log("PDF written:", sketchPath);
  }

  console.log("\nBuild complete!");
  console.log("Output folder:", LESSON_FOLDER);
}

build().catch((err) => { console.error(err); process.exit(1); });
