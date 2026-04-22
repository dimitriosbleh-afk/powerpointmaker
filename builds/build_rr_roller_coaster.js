"use strict";

// Respectful Relationships — A Roller Coaster of Emotions
// Grade 5/6 Wellbeing — Combined Activities 4 & 5
// 30-minute session: emotional intensity concept + Casey's story + create own roller coaster

const path = require("path");
const fs = require("fs");
const PptxGenJS = require("pptxgenjs");
const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf,
  addPdfHeader, addSectionHeading, addBodyText, addTipBox,
  addPdfFooter, addLinedArea, addWriteLine, addTwoColumnOrganiser,
  addResourceSlide, makeSessionResource, formatSessionResourceFileName,
  getSessionResourceFolder,
  PAGE, hex,
} = require("../themes/pdf_helpers");

// ── Theme ──
const T = createTheme("wellbeing", "grade56", weekToVariant(4));
const {
  C, FONT_H, FONT_B,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  pairShareSlide, scenarioSlide, reflectionSlide,
  withReveal,
  addTopBar, addBadge, addTitle, addCard, addFooter, addTextOnShape,
  addInstructionCard,
  CONTENT_TOP, SAFE_BOTTOM, SLIDE_W,
  runSlideDiagnostics,
} = T;

// ── Output paths ──
const UNIT = "RR_Roller_Coaster_Emotions";
const LESSON_FOLDER = path.join(__dirname, "..", "output", UNIT);
const PPTX_NAME = "A Roller Coaster of Emotions.pptx";
const FOOTER = "Respectful Relationships | Grade 5/6 Wellbeing";
const SESSION = 1;
const RES_DIR = path.join(LESSON_FOLDER, getSessionResourceFolder(SESSION));

// ── Resources ──
const CASEYS_DAY_RESOURCE = makeSessionResource(
  SESSION,
  "Caseys Day Handout",
  "Casey's morning story with roller coaster diagram for labelling emotions."
);
const WORD_LIST_RESOURCE = makeSessionResource(
  SESSION,
  "Emotional Intensity Word List",
  "Reference list showing emotions at increasing levels of intensity."
);
const ROLLER_COASTER_RESOURCE = makeSessionResource(
  SESSION,
  "My Roller Coaster Story Planner",
  "Planning sheet for creating your own roller coaster day story."
);
const RESOURCE_ITEMS = [CASEYS_DAY_RESOURCE, WORD_LIST_RESOURCE, ROLLER_COASTER_RESOURCE];

fs.mkdirSync(RES_DIR, { recursive: true });

// ═══════════════════════════════════════════════════════════════
// Teacher Notes
// ═══════════════════════════════════════════════════════════════

const NOTES_TITLE = [
  "SAY:",
  "- Today we are going to explore how our emotions can change throughout a single day",
  "- We'll look at how some emotions feel mild and others feel really strong -- and what makes that happen",
  "",
  "DO:",
  "- Display slide as students settle",
  "- Have Casey's Day handouts and Emotional Intensity Word Lists ready to distribute",
  "",
  "TEACHER NOTES:",
  "This session combines two Respectful Relationships activities into one 30-minute lesson. Students first explore the concept of emotional intensity through Casey's story, then apply this understanding by creating their own roller coaster story. Ensure handouts are pre-printed.",
  "",
  "WATCH FOR:",
  "- Students settling into wellbeing mode after a subject transition",
  "- Any students who may find emotional content sensitive -- check in privately",
  "",
  "[General: Title | VTLM 2.0: Establishing Purpose]",
].join("\n");

const NOTES_LI = [
  "SAY:",
  "- Read the Learning Intention from the slide",
  "- Read each Success Criterion aloud -- we will come back to these at the end to check how we went",
  "",
  "DO:",
  "- Point to each SC as you read it",
  "- Leave the slide visible for 15-20 seconds so students can read along",
  "",
  "TEACHER NOTES:",
  "The LI focuses on emotional intensity as the key concept. SC1 is the foundation -- all students should be able to describe what intensity means. SC2 connects events to emotions. SC3 extends into creative application.",
  "",
  "WATCH FOR:",
  "- Students who seem unfamiliar with the word 'intensity' -- this will be defined on the next slide",
  "- Students who are engaged and ready to learn",
  "",
  "[General: LI/SC | VTLM 2.0: Learning Intentions]",
].join("\n");

const NOTES_INTENSITY_CONCEPT = [
  "SAY:",
  "- We've looked at positive and negative emotions, and even mixed emotions. Today we are adding another layer -- intensity",
  "- Intensity means how strong an emotion feels. Think of a volume dial on a speaker",
  "- You might feel a little bit annoyed -- that's low intensity. Or you might feel absolutely enraged -- that's high intensity. Same type of emotion, different strength",
  "- We sometimes call these emotional 'highs' and 'lows'",
  "",
  "DO:",
  "- Point to the intensity scale visual on slide as you explain",
  "- Use hand gestures (low to high) to demonstrate the concept physically",
  "- Write 'intensity' on the board if a physical board is available",
  "",
  "TEACHER NOTES:",
  "This builds on prior Respectful Relationships work on positive/negative emotions and mixed emotions (Activity 2). The volume dial analogy is concrete and accessible for this age group.",
  "",
  "MISCONCEPTIONS:",
  "- Misconception: Intense emotions are always negative",
  "  Why: Students often associate 'intense' with 'bad' from everyday language",
  "  Impact: May not recognise that joy, excitement, and enthusiasm can also be intense",
  "  Quick correction: Give a quick example -- 'Think about how you feel when your team wins a grand final. That excitement is intense too.'",
  "",
  "WATCH FOR:",
  "- Students who confuse intensity with whether an emotion is positive or negative",
  "- Students nodding along -- the concept is intuitive once explained",
  "",
  "[General: I Do | VTLM 2.0: Explicit Explanation]",
].join("\n");

const NOTES_CASEYS_STORY = [
  "SAY:",
  "- I'm going to read you a story about Casey's morning. Listen carefully for the moments where Casey's emotions change",
  "- Pay attention to what EVENTS or SITUATIONS cause those shifts -- we call those triggers",
  "- Notice when emotions feel mild and when they feel really strong",
  "",
  "DO:",
  "- Read the story aloud with expression -- pause briefly at each emotional shift point",
  "- Distribute Casey's Day handout before or during reading",
  "- Point to the roller coaster diagram on the handout as you read",
  "",
  "TEACHER NOTES:",
  "The story is the anchor text for both the We Do labelling activity and the subsequent You Do creative task. Reading with expression helps students hear the emotional shifts naturally.",
  "",
  "WATCH FOR:",
  "- Students following along on their handout",
  "- Facial expressions that show students are connecting with the emotional moments in the story",
  "",
  "[General: I Do | VTLM 2.0: Modelling]",
].join("\n");

const NOTES_ROLLER_COASTER_LABEL_Q = [
  "SAY:",
  "- Now let's look at Casey's roller coaster. Each point on this ride matches a moment in the story",
  "- Look at the first point on the roller coaster -- stepping on mushy biscuits. What emotion might Casey have felt?",
  "- Ask: How intense was that emotion? Was it mild, medium, or strong? [Mild to medium -- disgusted, annoyed]",
  "- Let's use the Emotional Intensity Word List to find the right word. Is Casey just a bit irritated, or really angry?",
  "",
  "DO:",
  "- Distribute the Emotional Intensity Word List handout if not already distributed",
  "- Work through the first 2-3 roller coaster points together as a class",
  "- Cold call students to suggest emotions and justify their intensity choice",
  "- Model referencing the word list to select precise emotion words",
  "",
  "CFU CHECKPOINT:",
  "Technique: Cold Call with Word List Reference",
  "Script:",
  "- Say: Look at the point where Casey discovers the maths test. What emotion is Casey feeling? Use your word list to pick the exact word.",
  "- Cold call 3-4 students. Listen for: students distinguishing between mild and strong versions (e.g. 'nervous' vs 'fearful')",
  "- Scan for: students referencing the word list rather than guessing generic words",
  "PROCEED: If students are selecting intensity-appropriate words and can justify their choices, move to the next slide.",
  "PIVOT: If students are stuck on basic emotion names without considering intensity, model one more example explicitly: 'Casey could feel surprised or astounded. Surprised is mild -- astounded is really strong. The maths test is probably closer to nervous or fearful.'",
  "",
  "TEACHER NOTES:",
  "This is the guided practice phase. Working through the roller coaster together teaches students to match events to emotions AND calibrate intensity using the word list. The word list is the key scaffold.",
  "",
  "ENABLING & EXTENDING:",
  "ENABLING PROMPT:",
  "- Task: Students identify whether each emotion is positive or negative before considering intensity. Use a simpler two-choice frame: 'Is this a happy feeling or an unhappy feeling? Is it a little bit or a lot?'",
  "- Extra Notes: Pair enabling students with a supportive partner for this activity.",
  "EXTENDING PROMPT:",
  "- Task: Students identify moments where Casey might have felt mixed emotions (e.g. waiting for sister -- annoyed but also maybe anxious about being late) and explain how two emotions at different intensities can coexist.",
  "",
  "WATCH FOR:",
  "- Students who name emotions but cannot calibrate intensity -- prompt them to look at the word list columns",
  "- Students who are ready to label independently -- this signals readiness for You Do",
  "",
  "[General: We Do | VTLM 2.0: Scaffold Practice]",
].join("\n");

const NOTES_ROLLER_COASTER_LABEL_A = [
  "SAY:",
  "- Here are the emotions we identified together. Check your handout -- did you get similar words?",
  "- Notice how the roller coaster goes up for positive emotions and down for negative ones",
  "- The HEIGHT of each point shows the intensity -- the higher or lower the point, the stronger the emotion",
  "",
  "DO:",
  "- Click to reveal the labelled roller coaster",
  "- Allow 30 seconds for students to compare with their own annotations",
  "- Ask 1-2 students to share if they chose a different word and why",
  "",
  "TEACHER NOTES:",
  "The reveal validates student thinking and models how intensity maps to the roller coaster shape. Brief comparison discussion reinforces that there can be multiple valid emotion words at similar intensity levels.",
  "",
  "WATCH FOR:",
  "- Students who chose very different intensity levels -- use this as a teaching moment about how the same event can affect people differently",
  "- Students who are ready to move to creating their own story",
  "",
  "[General: We Do Reveal | VTLM 2.0: Scaffold Practice]",
].join("\n");

const NOTES_YOUR_ROLLER_COASTER = [
  "SAY:",
  "- Now it's your turn. You are going to create your own roller coaster day story",
  "- First: Choose a topic with your partner. You could do a day for a new prep student, a new student joining your class, or a babysitter's first evening",
  "- Next: Plan your roller coaster together -- list at least six emotions including highs, lows, and in-betweens. Use your word list",
  "- Then: Draw your roller coaster and label the events and emotions",
  "- Finally: Write up your story individually",
  "",
  "DO:",
  "- Distribute the My Roller Coaster Story Planner handout",
  "- Organise students into pairs",
  "- Set a timer -- 3 minutes for planning, 5 minutes for drawing the roller coaster, remaining time for writing",
  "- Circulate and check that pairs have chosen a topic and are using the word list",
  "",
  "TEACHER NOTES:",
  "Students plan in pairs for social skill practice and confidence building, then write individually. The planner handout provides structure. Ensure students use at least one high, one low, and one in-between emotion as per the activity requirements.",
  "",
  "ENABLING & EXTENDING:",
  "ENABLING PROMPT:",
  "- Task: Students use the Session 1 Caseys Day Handout as a model and create a shorter roller coaster with only 4 emotions (2 positive, 2 negative) using the simpler emotion words from the left side of the intensity word list.",
  "- Extra Notes: Provide sentence starters on the planner for the story write-up.",
  "EXTENDING PROMPT:",
  "- Task: Students write their story from two different characters' perspectives, showing how the same event can trigger different emotions at different intensities for different people. Include a reflection paragraph on empathy.",
  "",
  "WATCH FOR:",
  "- Pairs who are stuck choosing a topic -- suggest the new student scenario as it is most relatable",
  "- Students who list only one type of emotion (all positive or all negative) -- prompt them to include variety",
  "- Students finishing the roller coaster drawing quickly -- redirect to writing up the story",
  "",
  "[General: You Do | VTLM 2.0: Supported Application]",
].join("\n");

const NOTES_CLOSING = [
  "SAY:",
  "- Let's come back to our success criteria. Give me a thumbs up, sideways, or down for each one",
  "- SC1: I can describe what emotional intensity means",
  "- SC2: I can identify events that trigger emotions at different intensity levels",
  "- SC3: I can create a roller coaster story showing a range of emotional highs and lows",
  "- Ask: Why do you think it matters to understand that emotions come in different strengths? [Helps us understand ourselves and others better, helps us respond to big emotions, connects to empathy]",
  "",
  "DO:",
  "- Read each SC aloud and pause for thumbs",
  "- Note which SC students feel least confident about -- this informs next session's review",
  "- If time allows, invite 1-2 pairs to briefly share their roller coaster topic and one emotion they chose",
  "",
  "TEACHER NOTES:",
  "The closing connects directly to the concept of empathy -- understanding that others experience emotions at different intensities helps us imagine how they feel. This is the bridge to Activity 5's discussion about empathy and responding to others' experiences.",
  "",
  "WATCH FOR:",
  "- Students showing thumbs down on SC1 -- may need the intensity concept revisited next session",
  "- Students who seem reflective or thoughtful -- this is a positive sign of engagement with emotional content",
  "",
  "[General: Closing | VTLM 2.0: Review and Reflect]",
].join("\n");

const NOTES_RESOURCES = [
  "SAY:",
  "- These are the handouts for today's session. Make sure you have all three before we start",
  "",
  "DO:",
  "- Print Casey's Day Handout and Emotional Intensity Word List before the session",
  "- Print My Roller Coaster Story Planner -- one per student",
  "- Distribute Casey's Day and Word List at the start; Planner at the You Do stage",
  "",
  "TEACHER NOTES:",
  "All three resources are in the resources folder linked from this slide. Print double-sided where possible to save paper. The Word List is a reference tool students keep on their desk throughout the session.",
  "",
  "WATCH FOR:",
  "- Ensure all students have their handouts before beginning the read-aloud",
  "",
  "[General: Resources | VTLM 2.0: Preparation]",
].join("\n");

// ═══════════════════════════════════════════════════════════════
// Build function
// ═══════════════════════════════════════════════════════════════

async function build() {
  const pres = new PptxGenJS();
  pres.layout = "LAYOUT_16x9";

  // ── Slide 1: Title ──
  titleSlide(
    pres,
    "A Roller Coaster\nof Emotions",
    "Respectful Relationships",
    "Grade 5/6 Wellbeing",
    NOTES_TITLE
  );

  // ── Slide 2: LI / SC ──
  liSlide(
    pres,
    ["We are learning that emotions vary in intensity and that different events can trigger stronger or milder feelings"],
    [
      "I can describe what emotional intensity means",
      "I can identify events that trigger emotions at different intensity levels",
      "I can create a roller coaster story showing a range of emotional highs and lows",
    ],
    NOTES_LI,
    FOOTER
  );

  // ── Slide 3: What is Emotional Intensity? (I Do) ──
  contentSlide(
    pres,
    "I Do",
    C.PRIMARY,
    "What is Emotional Intensity?",
    [
      "Intensity means how strong an emotion feels",
      "Same emotion, different strength:",
      "  irritated ... annoyed ... angry ... enraged",
      "We call the strongest moments emotional 'highs' and 'lows'",
    ],
    NOTES_INTENSITY_CONCEPT,
    FOOTER,
    (s, lg) => {
      // Right column: intensity "volume dial" visual
      const rX = lg.rightX;
      const rW = lg.rightW;
      const topY = lg.panelTopPadded;

      // Intensity scale card
      addCard(s, rX, topY, rW, 3.4, { strip: C.ACCENT, fill: C.WHITE });

      // Title
      s.addText("Intensity Scale", {
        x: rX + 0.15, y: topY + 0.1, w: rW - 0.3, h: 0.3,
        fontSize: 13, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0, align: "center",
      });

      // Scale bars (low to high intensity)
      const levels = [
        { label: "Mild", color: C.SUCCESS, w: 0.6 },
        { label: "Moderate", color: C.ACCENT, w: 1.4 },
        { label: "Strong", color: C.ALERT, w: 2.2 },
        { label: "Intense", color: C.PRIMARY, w: 3.0 },
      ];
      const barStartY = topY + 0.55;
      const barH = 0.45;
      const barGap = 0.15;

      levels.forEach((lev, i) => {
        const by = barStartY + i * (barH + barGap);
        // Bar
        s.addShape("roundRect", {
          x: rX + 0.2, y: by, w: lev.w, h: barH, rectRadius: 0.06,
          fill: { color: lev.color },
        });
        // Label on bar
        s.addText(lev.label, {
          x: rX + 0.25, y: by, w: lev.w - 0.1, h: barH,
          fontSize: 11, fontFace: FONT_B, color: C.WHITE, bold: true,
          valign: "middle", margin: 0,
        });
      });

      // Arrow label
      s.addText("Getting stronger ->", {
        x: rX + 0.2, y: barStartY + 4 * (barH + barGap) - 0.1, w: rW - 0.4, h: 0.25,
        fontSize: 10, fontFace: FONT_B, color: C.MUTED, italic: true, margin: 0, align: "center",
      });
    }
  );

  // ── Slide 4: Casey's Morning (I Do — Story Read-Aloud) ──
  scenarioSlide(
    pres,
    "I Do",
    "Casey's Morning",
    "The day started like any other. Casey woke up and stepped right on something slimy -- baby sister's half-chewed biscuits. 'Dadddddd!!' But Dad had made Casey's favourite breakfast: egg and bacon muffin. Then waiting ages for big sister to get ready... three minutes to nine! Sneaking back from the office sign-in, Casey's best friend had saved a seat. But then -- a maths test on decimals. At recess, picked last for basketball. 'Don't pick him. He'll make us lose.' But then Rosie called out, 'Why don't we just sit and talk?'",
    [
      "Listen for: What events TRIGGERED Casey's emotions to shift?",
      "Notice: When did emotions feel MILD and when did they feel STRONG?",
    ],
    NOTES_CASEYS_STORY,
    FOOTER
  );

  // ── Roller coaster helper: draw track + event points ──
  function drawRollerCoaster(s, events, showAnswers) {
    const pointXPositions = events.map((_, i) => 0.6 + i * 1.25);

    // Draw connecting track lines
    for (let i = 0; i < events.length - 1; i++) {
      s.addShape("line", {
        x: pointXPositions[i] + 0.35, y: events[i].y + 0.15,
        w: pointXPositions[i + 1] - pointXPositions[i] - 0.35,
        h: events[i + 1].y - events[i].y,
        line: { color: C.MUTED, width: 2.0, dashType: "dash" },
      });
    }

    // Draw event points
    events.forEach((evt, i) => {
      const ex = pointXPositions[i];
      if (showAnswers) {
        // Coloured emotion pill
        s.addShape("roundRect", {
          x: ex, y: evt.y - 0.05, w: 0.7, h: 0.35,
          rectRadius: 0.08, fill: { color: evt.color || C.ACCENT },
        });
        s.addText(evt.emotion, {
          x: ex, y: evt.y - 0.05, w: 0.7, h: 0.35,
          fontSize: 9, fontFace: FONT_B, color: C.WHITE, bold: true,
          align: "center", valign: "middle", margin: 0,
        });
      } else {
        // Question mark dot
        s.addShape("roundRect", {
          x: ex + 0.15, y: evt.y, w: 0.4, h: 0.3,
          rectRadius: 0.15, fill: { color: C.ACCENT },
        });
        s.addText("?", {
          x: ex + 0.15, y: evt.y, w: 0.4, h: 0.3,
          fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true,
          align: "center", valign: "middle", margin: 0,
        });
      }
      // Event label
      s.addText(evt.label, {
        x: ex - 0.15, y: evt.y + 0.35, w: 1.0, h: 0.4,
        fontSize: 8.5, fontFace: FONT_B, color: C.CHARCOAL,
        align: "center", valign: "top", margin: 0,
      });
    });
  }

  const rollerCoasterEvents = [
    { label: "Slimy biscuits", y: 3.2, emotion: "Disgusted", color: C.ALERT },
    { label: "Favourite breakfast", y: 1.8, emotion: "Pleased", color: C.SUCCESS },
    { label: "Sister makes us late", y: 3.0, emotion: "Annoyed", color: C.ALERT },
    { label: "Best friend saved a seat", y: 1.6, emotion: "Happy", color: C.SUCCESS },
    { label: "Maths test on decimals", y: 3.4, emotion: "Fearful", color: C.ALERT },
    { label: "Picked last", y: 3.8, emotion: "Miserable", color: C.ALERT },
    { label: "Rosie: Let's just talk", y: 2.0, emotion: "Pleased", color: C.SUCCESS },
  ];

  // ── Slide 5-6: Casey's Roller Coaster — Label Together (We Do, with reveal) ──
  function buildRollerCoasterBase() {
    const s = pres.addSlide();
    addTopBar(s, C.SECONDARY);
    addBadge(s, "We Do", { color: C.SECONDARY });
    addTitle(s, "Casey's Roller Coaster");

    // Instruction text
    s.addText("Use your Emotional Intensity Word List to label each emotion", {
      x: 0.5, y: CONTENT_TOP, w: 9, h: 0.3,
      fontSize: 12, fontFace: FONT_B, color: C.PRIMARY, italic: true, margin: 0,
    });

    drawRollerCoaster(s, rollerCoasterEvents, false);
    addFooter(s, FOOTER);
    s.addNotes(NOTES_ROLLER_COASTER_LABEL_Q);
    return s;
  }

  withReveal(
    buildRollerCoasterBase,
    (s) => {
      // Add answer overlay: emotion pills on top of the ? marks
      // The base slide already has the ? marks; we overlay the answers
      rollerCoasterEvents.forEach((evt, i) => {
        const ex = 0.6 + i * 1.25;
        // Coloured emotion pill (covers the ? dot)
        s.addShape("roundRect", {
          x: ex, y: evt.y - 0.05, w: 0.7, h: 0.35,
          rectRadius: 0.08, fill: { color: evt.color },
        });
        s.addText(evt.emotion, {
          x: ex, y: evt.y - 0.05, w: 0.7, h: 0.35,
          fontSize: 9, fontFace: FONT_B, color: C.WHITE, bold: true,
          align: "center", valign: "middle", margin: 0,
        });
      });

      // Override instruction text
      s.addShape("rect", {
        x: 0.5, y: CONTENT_TOP, w: 9, h: 0.3,
        fill: { color: C.BG_CARD },
      });
      s.addText("Check your labels -- did you choose similar intensity words?", {
        x: 0.5, y: CONTENT_TOP, w: 9, h: 0.3,
        fontSize: 12, fontFace: FONT_B, color: C.PRIMARY, italic: true, margin: 0,
      });

      s.addNotes(NOTES_ROLLER_COASTER_LABEL_A);
    }
  );

  // ── Slide 7: Your Roller Coaster Story (You Do) ──
  contentSlide(
    pres,
    "You Do",
    C.ACCENT,
    "Your Roller Coaster Story",
    [
      "First: Choose a topic with your partner",
      "Next: Plan at least 6 emotions (highs, lows, in-betweens) using your word list",
      "Then: Draw your roller coaster and label events and emotions",
      "Finally: Write up your story individually",
    ],
    NOTES_YOUR_ROLLER_COASTER,
    FOOTER,
    (s, lg) => {
      // Right column: topic choices
      const rX = lg.rightX;
      const rW = lg.rightW;
      const topY = lg.panelTopPadded;

      addCard(s, rX, topY, rW, 3.2, { strip: C.SECONDARY, fill: C.WHITE });

      s.addText("Choose a Topic", {
        x: rX + 0.15, y: topY + 0.1, w: rW - 0.3, h: 0.3,
        fontSize: 13, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0, align: "center",
      });

      const topics = [
        "A day in the life of a new Prep student",
        "A day in the life of a new student joining YOUR class",
        "The evening experience of a new babysitter",
      ];

      topics.forEach((topic, i) => {
        const ty = topY + 0.55 + i * 0.75;
        // Number circle
        s.addShape("roundRect", {
          x: rX + 0.2, y: ty, w: 0.35, h: 0.35,
          rectRadius: 0.175, fill: { color: C.SECONDARY },
        });
        s.addText(String(i + 1), {
          x: rX + 0.2, y: ty, w: 0.35, h: 0.35,
          fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true,
          align: "center", valign: "middle", margin: 0,
        });
        // Topic text
        s.addText(topic, {
          x: rX + 0.65, y: ty, w: rW - 0.9, h: 0.5,
          fontSize: 12.5, fontFace: FONT_B, color: C.CHARCOAL, valign: "middle", margin: 0,
        });
      });

      // Reminder
      s.addText("Include at least 1 high, 1 low, and 1 in-between emotion", {
        x: rX + 0.15, y: topY + 2.85, w: rW - 0.3, h: 0.25,
        fontSize: 10, fontFace: FONT_B, color: C.MUTED, italic: true, margin: 0, align: "center",
      });
    }
  );

  // ── Slide 8: Closing ──
  closingSlide(
    pres,
    "Why does it matter that emotions come in different strengths? How does understanding intensity help us understand others?",
    [
      "Emotions vary in intensity -- from mild to strong",
      "Events and situations trigger emotional shifts throughout any day",
      "Understanding intensity helps us name our feelings precisely and empathise with others",
    ],
    NOTES_CLOSING
  );

  // ── Slide 9: Resources ──
  addResourceSlide(pres, RESOURCE_ITEMS, T, FOOTER, NOTES_RESOURCES);

  // ── Write PPTX ──
  fs.mkdirSync(LESSON_FOLDER, { recursive: true });
  await pres.writeFile({ fileName: path.join(LESSON_FOLDER, PPTX_NAME) });
  console.log("PPTX written to", path.join(LESSON_FOLDER, PPTX_NAME));

  // ═══════════════════════════════════════════════════════════════
  // Companion PDF 1: Casey's Day Handout
  // ═══════════════════════════════════════════════════════════════
  {
    const doc = createPdf({ title: "Casey's Day Handout" });
    let y = addPdfHeader(doc, "Session 1 Caseys Day Handout", {
      subtitle: "Respectful Relationships - A Roller Coaster of Emotions",
      color: C.PRIMARY,
      lessonInfo: "Grade 5/6 Wellbeing",
    });

    y = addSectionHeading(doc, "Casey's Morning", y, { color: C.PRIMARY });
    y = addBodyText(doc, [
      "The day started like any other, I woke up and climbed out of bed. Eww my foot landed right on something slimy and mushy -- my baby sister's half chewed biscuits... 'Dadddddd!!'",
      "",
      "'Just get ready for school,' was all he said.",
      "",
      "Walking out into the kitchen I noticed he had made my favourite breakfast: egg and bacon muffin. 'Thanks Da' I said. He smiled at me and ruffled my hair.",
      "",
      "I ate and went quickly to get dressed for school to have time to play with my friends before the bell. But I had to wait ages until my big sister was finally ready to leave for school. It was three minutes to nine! We were going to be late. My teacher gets so annoyed when we are late. She sends us to the office to sign in.",
      "",
      "When I arrived in class, after sneaking back from the office sign-in, I saw that my best friend had saved me a seat. Then I discovered we were doing a Maths test on decimals; and last week I got 7 out of 10 wrong!",
      "",
      "When the recess bell rang, Jordan suggested, 'Let's play basketball'. As they picked teams I stood and waited and waited and waited. I was the last one chosen. I could hear Lou saying, 'Don't pick him. He'll make us lose'. I turned around and walked off.",
      "",
      "As I trudged over to the taps, Rosie called, 'Why don't we just sit and talk? I don't even like basketball.'"
    ].join("\n"), y, { fontSize: 10.5 });

    y += 10;
    y = addSectionHeading(doc, "Casey's Roller Coaster", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Label each point on the roller coaster with the emotion Casey felt and how intense it was. Use your Emotional Intensity Word List.", y, { fontSize: 10, italic: true });

    // Roller coaster events as a labelling table
    const rcEvents = [
      "1. Slimy biscuits on the floor",
      "2. Dad made favourite breakfast",
      "3. Sister makes us late for school",
      "4. Best friend saved a seat",
      "5. Maths test on decimals",
      "6. Picked last for basketball",
      "7. Rosie says 'Let's just sit and talk'",
    ];

    y = addTwoColumnOrganiser(doc, "Event in Casey's Day", "Emotion + Intensity Level", y, {
      color: C.PRIMARY,
      rows: 7,
      rowH: 38,
      leftContent: rcEvents,
    });

    y += 5;
    y = addTipBox(doc, "Hint: Use the Emotional Intensity Word List to choose the precise word. Is the emotion mild (e.g. irritated), moderate (e.g. annoyed), strong (e.g. angry), or intense (e.g. enraged)?", y, { color: C.ACCENT });

    addPdfFooter(doc, "Respectful Relationships | Grade 5/6 | Session 1");
    const caseyPath = path.join(RES_DIR, "Session 1 Caseys Day Handout.pdf");
    await writePdf(doc, caseyPath);
    console.log("PDF written:", caseyPath);
  }

  // ═══════════════════════════════════════════════════════════════
  // Companion PDF 2: Emotional Intensity Word List
  // ═══════════════════════════════════════════════════════════════
  {
    const doc = createPdf({ title: "Emotional Intensity Word List" });
    let y = addPdfHeader(doc, "Session 1 Emotional Intensity Word List", {
      subtitle: "Reference: Emotions at increasing levels of intensity",
      color: C.PRIMARY,
      showNameDate: false,
      lessonInfo: "Grade 5/6 Wellbeing - Respectful Relationships",
    });

    y = addTipBox(doc, "Use this word list to find the precise word for how someone feels. Move from left (mild) to right (intense) to match the strength of the emotion.", y, { color: C.ACCENT });

    // Intensity table
    const categories = [
      { label: "Fear",        words: ["apprehension", "nervous", "fearful", "terrified"] },
      { label: "Surprise",    words: ["surprised", "amazed", "astounded", "stunned"] },
      { label: "Disgust",     words: ["dislike", "horror", "disgusted", "loathing"] },
      { label: "Anger",       words: ["irritated", "annoyed", "angry", "enraged"] },
      { label: "Happiness",   words: ["pleased", "happy", "joyful", "ecstatic"] },
      { label: "Interest",    words: ["interested", "keen", "eager", "enthusiastic"] },
      { label: "Sadness",     words: ["unhappy", "sad", "miserable", "despairing"] },
    ];

    const colW = (PAGE.CONTENT_W - 80) / 4;
    const hdrH = 26;
    const rowH = 32;
    const labelColW = 80;
    const tableX = PAGE.MARGIN;

    // Column headers
    const intensityHeaders = ["Mild", "Moderate", "Strong", "Intense"];
    doc.save();
    // Label column header
    doc.rect(tableX, y, labelColW, hdrH).fill(hex(C.PRIMARY));
    doc.fontSize(9).font("Sans-Bold").fillColor("#FFFFFF");
    doc.text("Emotion", tableX, y + 7, { width: labelColW, align: "center" });
    // Intensity headers
    intensityHeaders.forEach((h, i) => {
      const cx = tableX + labelColW + i * colW;
      doc.rect(cx, y, colW, hdrH).fill(hex(C.ACCENT));
      doc.fontSize(9).font("Sans-Bold").fillColor("#FFFFFF");
      doc.text(h, cx, y + 7, { width: colW, align: "center" });
    });
    doc.restore();
    y += hdrH;

    // Data rows
    categories.forEach((cat, ri) => {
      const bgColor = ri % 2 === 0 ? "#FFFFFF" : hex("F3F4F6");
      doc.save();
      // Label cell
      doc.rect(tableX, y, labelColW, rowH).fill(bgColor);
      doc.rect(tableX, y, labelColW, rowH).lineWidth(0.5).strokeColor(hex("D1D5DB")).stroke();
      doc.fontSize(9).font("Sans-Bold").fillColor(hex(C.PRIMARY));
      doc.text(cat.label, tableX + 4, y + 9, { width: labelColW - 8, align: "center" });

      // Word cells
      cat.words.forEach((word, wi) => {
        const cx = tableX + labelColW + wi * colW;
        doc.rect(cx, y, colW, rowH).fill(bgColor);
        doc.rect(cx, y, colW, rowH).lineWidth(0.5).strokeColor(hex("D1D5DB")).stroke();
        doc.fontSize(9.5).font("Sans").fillColor(hex(C.CHARCOAL));
        doc.text(word, cx + 4, y + 9, { width: colW - 8, align: "center" });
      });
      doc.restore();
      y += rowH;
    });

    y += 15;
    y = addTipBox(doc, "The arrow direction shows emotions becoming MORE INTENSE from left to right. The same type of emotion (e.g. anger) can be felt at different strengths depending on the situation.", y, { color: C.PRIMARY });

    y = addSectionHeading(doc, "How to Use This Word List", y, { color: C.PRIMARY });
    y = addBodyText(doc, "1. Identify the TYPE of emotion (e.g. anger, sadness, happiness)\n2. Think about how STRONG the feeling is\n3. Choose the word that best matches the intensity\n4. Use that word to describe the emotion precisely", y, { fontSize: 10 });

    addPdfFooter(doc, "Respectful Relationships | Grade 5/6 | Session 1");
    const wordListPath = path.join(RES_DIR, "Session 1 Emotional Intensity Word List.pdf");
    await writePdf(doc, wordListPath);
    console.log("PDF written:", wordListPath);
  }

  // ═══════════════════════════════════════════════════════════════
  // Companion PDF 3: My Roller Coaster Story Planner
  // ═══════════════════════════════════════════════════════════════
  {
    const doc = createPdf({ title: "My Roller Coaster Story Planner" });
    let y = addPdfHeader(doc, "Session 1 My Roller Coaster Story Planner", {
      subtitle: "Plan and create your own roller coaster day story",
      color: C.SECONDARY,
      lessonInfo: "Grade 5/6 Wellbeing - Respectful Relationships",
    });

    y = addSectionHeading(doc, "Step 1: Choose Your Topic", y, { color: C.SECONDARY });
    y = addBodyText(doc, "Circle one:", y, { fontSize: 10 });

    const topicChoices = [
      "A day in the life of a new Prep student",
      "A day in the life of a new student joining your class",
      "The evening experience of a new babysitter",
    ];
    topicChoices.forEach((topic, i) => {
      doc.fontSize(10.5).font("Sans").fillColor(hex(C.CHARCOAL));
      doc.text(`   ${i + 1}.  ${topic}`, PAGE.MARGIN, y);
      y = doc.y + 4;
    });
    y += 8;

    y = addSectionHeading(doc, "Step 2: Plan Your Emotions", y, { color: C.SECONDARY });
    y = addTipBox(doc, "List at least 6 emotions. Include at least 1 high (positive/intense), 1 low (negative/intense), and 1 in-between. Use your Emotional Intensity Word List.", y, { color: C.ACCENT });

    // Planning table
    y = addTwoColumnOrganiser(doc, "Event / Situation", "Emotion + Intensity", y, {
      color: C.SECONDARY,
      rows: 7,
      rowH: 36,
    });

    y += 5;
    y = addSectionHeading(doc, "Step 3: Draw Your Roller Coaster", y, { color: C.SECONDARY });
    y = addBodyText(doc, "Draw your roller coaster below. Place events at high points for positive emotions, low points for negative emotions, and middle for in-between. Label each point with the event and emotion.", y, { fontSize: 9.5, italic: true });

    // Drawing space (large box)
    const drawH = 180;
    if (y + drawH > PAGE.H - PAGE.MARGIN - 40) {
      doc.addPage();
      y = PAGE.MARGIN;
      y = addSectionHeading(doc, "Step 3: Draw Your Roller Coaster", y, { color: C.SECONDARY });
    }
    doc.save();
    doc.roundedRect(PAGE.MARGIN, y, PAGE.CONTENT_W, drawH, 4)
      .lineWidth(1).strokeColor(hex(C.MUTED)).stroke();
    // Axis labels in the drawing area
    doc.fontSize(7.5).font("Sans-Italic").fillColor(hex("9CA3AF"));
    doc.text("POSITIVE (highs)", PAGE.MARGIN + 4, y + 6);
    doc.text("NEGATIVE (lows)", PAGE.MARGIN + 4, y + drawH - 14);
    doc.restore();
    y += drawH + 15;

    // Page 2: Story writing space
    doc.addPage();
    y = PAGE.MARGIN;
    y = addSectionHeading(doc, "Step 4: Write Your Story", y, { color: C.SECONDARY });
    y = addBodyText(doc, "Write your roller coaster day story. Include the events, the emotions your character felt, and how intense those emotions were. Use your emotion words.", y, { fontSize: 10, italic: true });
    y += 5;
    y = addLinedArea(doc, y, 22, { lineSpacing: 26 });

    addPdfFooter(doc, "Respectful Relationships | Grade 5/6 | Session 1");
    const plannerPath = path.join(RES_DIR, "Session 1 My Roller Coaster Story Planner.pdf");
    await writePdf(doc, plannerPath);
    console.log("PDF written:", plannerPath);
  }

  console.log("\nBuild complete!");
  console.log("Output folder:", LESSON_FOLDER);
}

build().catch((err) => { console.error(err); process.exit(1); });
