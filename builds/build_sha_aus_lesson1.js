"use strict";

// Inquiry - Shaping Australia, Lesson 1 (Term 2, Week 1)
// Year 5/6 | First session of "Change - Shaping Australia" unit.
// Activates prior knowledge, watches two BTN clips (whole-class), builds a
// class KWL chart, then groups of 3 fill in their own copy and share their
// wonderings. Sets up the 10-week unit toward a Celebration Day in Week 9.
//
// User-supplied video links (preserved exactly):
//   BTN - Australia Since Cook
//     https://www.abc.net.au/btn/classroom/australia-since-cook/12176924?jwsource=cl
//   BTN - Indigenous Perspectives of Cook's Visit
//     https://www.abc.net.au/btn/classroom/indigenous-perspectives-of-cook-visit/12176674?jwsource=cl

const path = require("path");
const fs = require("fs");
const PptxGenJS = require("pptxgenjs");
const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addBodyText, addTipBox,
  addPdfFooter, addResourceSlide, makeSessionResource,
  getSessionResourceFolder, PAGE, hex,
} = require("../themes/pdf_helpers");

// -- Theme --
// Term 2 Week 1 -> variant 0 (Explorer - bold olive expedition palette)
const T = createTheme("inquiry", "grade56", weekToVariant(1));
const {
  C, FONT_H, FONT_B,
  titleSlide, liSlide, contentSlide, closingSlide,
  exitTicketSlide, pairShareSlide, withReveal,
  addTopBar, addBadge, addTitle, addCard, addFooter,
  addInstructionCard,
  CONTENT_TOP, SAFE_BOTTOM,
} = T;

// -- Output paths --
const UNIT = "Shaping_Australia";
const LESSON_FOLDER = path.join(__dirname, "..", "output", UNIT);
const PPTX_NAME = "Lesson 1 - Shaping Australia Launch.pptx";
const FOOTER = "Inquiry | Year 5/6 | Shaping Australia | Lesson 1";
const SESSION = 1;
const RES_DIR = path.join(LESSON_FOLDER, getSessionResourceFolder(SESSION));

// -- Video URLs (user-supplied, preserve exactly) --
const URL_BTN_COOK = "https://www.abc.net.au/btn/classroom/australia-since-cook/12176924?jwsource=cl";
const URL_BTN_INDIG = "https://www.abc.net.au/btn/classroom/indigenous-perspectives-of-cook-visit/12176674?jwsource=cl";

// -- Resources --
const KWL_CHART = makeSessionResource(
  SESSION,
  "Group KWL Chart",
  "A3 (or A4) chart for groups of 3: What we already know - What we learnt - What we wonder."
);
const RESOURCE_ITEMS = [KWL_CHART];

fs.mkdirSync(RES_DIR, { recursive: true });

// ===============================================================
// Teacher Notes
// ===============================================================

const NOTES_TITLE = [
  "SAY:",
  "- Welcome to our new inquiry unit. Over the next ten weeks we are going to look at how Australia has changed",
  "- Today is the launch lesson. We collect what we already think we know, watch two short clips, and start a class wonderings list",
  "- We are working towards a Celebration Day in Week 9 - a chance to share what we have learnt about the people and events that shaped Australia",
  "",
  "DO:",
  "- Have the class KWL chart drawn up on the whiteboard before students arrive (three columns: What we already know - What we learnt - What we wonder)",
  "- Have the Group KWL Chart printed and ready for the You Do (one per group of three)",
  "- Have the two BTN clip links queued so they are ready to play",
  "",
  "TEACHER NOTES:",
  "This is the first lesson in a ten-week unit. Treat it as activation, not assessment. Many students will hold partial or inaccurate ideas - that is the point of the K column.",
  "",
  "WATCH FOR:",
  "- Students settling in - keep the topic open and curious, not corrective",
  "",
  "[Inquiry: Title | VTLM 2.0: Establishing Purpose]",
].join("\n");

const NOTES_RESOURCES = [
  "SAY:",
  "- We have one printed resource and a class chart for today",
  "",
  "DO:",
  "- Print the Group KWL Chart - one per group of three (A3 if you have it, A4 works)",
  "- Draw up the matching three-column class chart on the whiteboard",
  "- Open both BTN clip links on the teacher computer ready to play",
  "- Have mini-whiteboards available for the CFU check",
  "",
  "TEACHER NOTES:",
  "This is a low-resource lesson by design - the talking, watching, and group thinking are the lesson. The chart structures the thinking; it should not become the focus.",
  "",
  "WATCH FOR:",
  "- Print the Group KWL Chart before the lesson - it is needed at the You Do stage",
  "",
  "[Inquiry: Resources | VTLM 2.0: Preparation]",
].join("\n");

const NOTES_HOOK = [
  "SAY:",
  "- Picture Australia about 250 years ago. Long before phones, cars, or shopping centres",
  "- Now picture Australia today. The streets, the people, the buildings, the schools",
  "- A LOT has changed. Some of those changes were good for some people and hard for others",
  "- Today we start asking: how did we get from then to now?",
  "",
  "DO:",
  "- Read the slide aloud, slowly. Allow students 10 seconds of thinking time",
  "- Quick partner share: 'Tell your partner ONE way you think Australia is different to 250 years ago'",
  "- Cold call 2-3 students to share what their partner said",
  "",
  "TEACHER NOTES:",
  "The hook bridges from familiar (today) to unfamiliar (250 years ago). Keep it open - any reasonable answer counts. This becomes the K column of the class chart.",
  "",
  "WATCH FOR:",
  "- Students who say 'I don't know anything about it' - reassure them: this is the K column - we collect what we think we already know, even if we are unsure",
  "- Students who jump to specific events - capture those for the brainstorm next slide",
  "",
  "[Inquiry: Launch | VTLM 2.0: Activating Prior Knowledge]",
].join("\n");

const NOTES_VOCAB = [
  "SAY:",
  "- Four inquiry words to keep in mind today",
  "- Read each word and meaning with me",
  "- 'Perspective' is the big one - it means the view someone has based on who they are. Two people can see the same event very differently",
  "",
  "DO:",
  "- Read each word aloud. Have students say 'perspective' and 'wondering' after you",
  "- Point to the small icons as you go",
  "",
  "CFU CHECKPOINT:",
  "Technique: Finger Voting",
  "Script:",
  "- Say: Hold up 1 finger if a 'wondering' is the same as a fact, 2 if a wondering is a question we still want to find out",
  "- Scan for: most students showing 2 - a wondering is a question we have not yet answered",
  "PROCEED: If 80% show 2, move on to the LI / SC.",
  "PIVOT: If students are confused, say: 'A wondering is a question. It starts with what, why, how, who, when. We collect them today and answer them across the unit.'",
  "",
  "TEACHER NOTES:",
  "These four words anchor the unit. 'Perspective' will return when we watch the second clip. 'Wondering' is the third column of the chart.",
  "",
  "WATCH FOR:",
  "- Students confusing 'wondering' with 'guessing' - clarify: a wondering is a real question, not a guess at an answer",
  "",
  "[Inquiry: Vocabulary | VTLM 2.0: Establishing Knowledge]",
].join("\n");

const NOTES_LI = [
  "SAY:",
  "- Read the Learning Intention aloud",
  "- Read each Success Criterion. We will come back to these at the end to check how we went",
  "",
  "DO:",
  "- Point to each I can statement as you read it",
  "- Leave the slide visible for about 20 seconds",
  "",
  "TEACHER NOTES:",
  "Internal tier mapping (do not share with students). SC1 - everyone can recall something. SC2 - most students will form questions. SC3 - groups think about cause-and-effect across events.",
  "",
  "WATCH FOR:",
  "- Students who feel they 'know nothing' - reassure them that recalling even one thing meets SC1",
  "",
  "[Inquiry: LI/SC | VTLM 2.0: Learning Intentions]",
].join("\n");

const NOTES_IDO = [
  "SAY:",
  "- Watch how I add to our class chart",
  "- I'm going to add ONE thing I think I already know in the K column",
  "- Think aloud: 'I think the British arrived a long time ago, around 1788. I'm not 100 percent sure of the year, but I'll write it - we'll check it later'",
  "- Notice I said 'I'm not 100 percent sure' - that is fine. The K column is what we THINK we know",
  "",
  "DO:",
  "- Stand at the whiteboard with your marker",
  "- Write your example in the first (left) column out loud",
  "- Underline 'around' to model uncertainty - we are not pretending to be certain",
  "- Then invite the class: 'Brainstorm with your partner - one thing you think you already know about how Australia has changed'",
  "- Take 6-8 student contributions and add them to the K column. Use student wording. Do not correct unless safety or clearly factual error",
  "",
  "TEACHER NOTES:",
  "This is the I Do for an inquiry brainstorm - it is short. The teacher models that uncertain knowledge belongs in the K column. Avoid turning this into a fact-check session.",
  "",
  "WATCH FOR:",
  "- Students who name First Nations history (e.g. 'Aboriginal people were here first') - capture this exactly. It connects to the second clip",
  "- Students who only mention modern things (cars, phones) - prompt: 'What about further back?'",
  "",
  "[Inquiry: I Do | VTLM 2.0: Explicit Modelling]",
].join("\n");

const NOTES_VIDEO1 = [
  "SAY:",
  "- Now we watch our first clip - BTN's 'Australia Since Cook'",
  "- Your job while you watch: notice ONE thing that you didn't know before",
  "- That one thing goes in the middle column of our chart - 'What we learnt'",
  "",
  "DO:",
  "- Open the BTN link",
  "- Play the full clip (about 4 minutes)",
  "- Pause once after the clip - do NOT discuss yet",
  "- Quick partner share: 'Tell your partner ONE thing you learnt'",
  "- Take 5-6 contributions. Add them to the L column with student wording",
  "",
  "SOURCES:",
  "- BTN, 'Australia Since Cook'. ABC Education classroom resource. Used as supplied: " + URL_BTN_COOK,
  "",
  "TEACHER NOTES:",
  "We watch first, then we share. Avoid pausing the clip mid-way - first viewing is for getting the shape of the story. Save pauses for later lessons in the unit when students dig into specific events.",
  "",
  "SENSITIVITY ADVISORY:",
  "- What it is: The clip discusses British colonisation, including impacts on First Nations people.",
  "- Framing language: 'This is real Australian history. It includes parts that are hard. Historians study it so we understand what changed and why.'",
  "- Watch for: students with strong personal or family connection to this history.",
  "- Protocol: If a student needs a moment, offer them quiet space at the desk and check in privately afterwards.",
  "",
  "WATCH FOR:",
  "- Students wanting to comment during the clip - hold them with: 'Save it - we'll share when it ends'",
  "- Students whose first reaction is shock at a specific date or fact - acknowledge it: 'That surprised me too the first time I learnt it'",
  "",
  "[Inquiry: We Do (Watch) | VTLM 2.0: Shared Experience]",
].join("\n");

const NOTES_VIDEO2 = [
  "SAY:",
  "- Now our second clip - BTN's 'Indigenous Perspectives of Cook's Visit'",
  "- This clip looks at the SAME event we just watched - Cook arriving - but from a DIFFERENT perspective",
  "- Remember our vocabulary word? Perspective means the view someone has based on who they are",
  "- Your job: notice anything that is DIFFERENT from the first clip. Different details, different feelings, different words used",
  "",
  "DO:",
  "- Open the second BTN link",
  "- Play the full clip",
  "- Pause - do NOT discuss yet",
  "- Quick partner share: 'Tell your partner ONE thing that was different in this clip'",
  "- Take 5-6 contributions. Add them to the L column - use student wording",
  "- Briefly point to entries from the first clip and ask: 'Did this clip change anything we wrote?'",
  "",
  "SOURCES:",
  "- BTN, 'Indigenous Perspectives of Cook's Visit'. ABC Education classroom resource. Used as supplied: " + URL_BTN_INDIG,
  "",
  "TEACHER NOTES:",
  "The point of pairing the two clips is exactly perspective. Students should notice the same event told differently. This sets up the CFU on the next slide.",
  "",
  "SENSITIVITY ADVISORY:",
  "- What it is: First Nations perspectives include accounts of dispossession and harm.",
  "- Framing language: 'This is part of our shared history. We listen carefully to all perspectives so we can understand the whole picture.'",
  "- Watch for: First Nations students or students with personal connection - check in quietly during or after.",
  "- Protocol: Offer a private chat after class if needed. Avoid singling out any student to speak for a community.",
  "",
  "WATCH FOR:",
  "- Students saying 'one of these clips is wrong' - prompt: 'Or are they two different perspectives on the same event?'",
  "- Students sitting quietly - this content is heavy. Quiet thinking is a fine response",
  "",
  "[Inquiry: We Do (Watch) | VTLM 2.0: Multiple Perspectives]",
].join("\n");

const NOTES_CFU_Q = [
  "SAY:",
  "- Mini-whiteboards out",
  "- Write ONE word that finishes this sentence: 'When two people tell the story of the same event differently, we call those different ___.'",
  "- You have 30 seconds. You can use the vocabulary slide if you need it",
  "",
  "DO:",
  "- Set a 30-second timer",
  "- Scan whiteboards as students hold them up",
  "",
  "CFU CHECKPOINT:",
  "Technique: Show Me Boards",
  "Script:",
  "- Say: Boards up on three. One, two, three",
  "- Scan for: 'perspectives' as the target answer. 'opinions' or 'sides' show partial understanding",
  "PROCEED: If 80% write perspectives, reveal and move on.",
  "PIVOT: If many write opinions or sides, say: 'Close. Opinions are what people THINK. Perspectives are the VIEW someone has - based on who they are, where they are, what has happened to them.' Then re-check with the same prompt.",
  "",
  "TEACHER NOTES:",
  "This hinge checks the key vocabulary word and the big idea behind today's two clips. Wrong answers are still useful - opinions and sides show students noticed difference, just need the precise word.",
  "",
  "WATCH FOR:",
  "- Students writing 'stories' - acknowledge but redirect: 'Stories is good - what is the inquiry word for the way each story is shaped?'",
  "",
  "[Inquiry: CFU | VTLM 2.0: Checking for Understanding]",
].join("\n");

const NOTES_CFU_A = [
  "SAY:",
  "- The target word is perspectives",
  "- Two perspectives on Cook's visit: the British perspective in clip 1, and the Indigenous perspective in clip 2",
  "- Across this unit we are going to keep asking - whose perspective is this? Whose perspective is missing?",
  "",
  "DO:",
  "- Click to reveal the answer",
  "- Briefly acknowledge close answers ('opinions', 'sides') - validate the noticing",
  "- Move on promptly to the You Do",
  "",
  "TEACHER NOTES:",
  "Validate partial answers so students stay engaged for the You Do. This is also seeding a habit - whose perspective is this? - that will return throughout the unit.",
  "",
  "WATCH FOR:",
  "- Students who are still rewriting after the reveal - that is fine, that is learning",
  "",
  "[Inquiry: CFU Reveal | VTLM 2.0: Feedback]",
].join("\n");

const NOTES_YOUDO = [
  "SAY:",
  "- Your turn. Move into your groups of three",
  "- Each group gets ONE Group KWL Chart",
  "- First: as a group, agree on TWO things to write in the K column - what your group already thought you knew",
  "- Next: write THREE things in the L column - things from the two clips that you learnt",
  "- Then: write at least TWO wonderings in the W column - questions you now want to find out",
  "- Use full sentences for your wonderings - they should start with what, why, how, who, or when",
  "",
  "DO:",
  "- Distribute the Group KWL Chart - one per group of three",
  "- Project this slide while groups work",
  "- Set a 12-minute timer",
  "- Circulate. Prompt quiet groups: 'What was the one thing in the second clip that surprised you?'",
  "- Photograph each group's chart for the unit folder before students leave",
  "",
  "ENABLING & EXTENDING:",
  "ENABLING PROMPT:",
  "- Task: Pair the group with another group nearby. Share aloud first, then write. Use a sentence starter for wonderings: 'I wonder why...' or 'What happened when...'",
  "- Extra Notes: It is fine for an enabling group to write fewer items if they are well thought through.",
  "EXTENDING PROMPT:",
  "- Task: After the group fills the chart, the group writes ONE 'cause and effect' wondering - a question that links two events. For example: 'Why did the British settlement change the lives of First Nations people?'",
  "",
  "TEACHER NOTES:",
  "The You Do is collaborative - groups of three work better than pairs for this thinking work. Encourage groups to use exact wording from the clips when they can.",
  "",
  "WATCH FOR:",
  "- Groups who write yes/no questions in the W column - prompt: 'Open it up. Start with why or how'",
  "- Groups who finish early - direct them to the extending task",
  "- Groups who get stuck on the W column - prompt: 'What was the most surprising thing in either clip? Turn that into a question'",
  "",
  "[Inquiry: You Do | VTLM 2.0: Supported Application]",
].join("\n");

const NOTES_DISCUSS = [
  "SAY:",
  "- Each group, choose ONE wondering you would most like the class to investigate",
  "- We are going to hear them now and add the strongest ones to our class chart",
  "- These wonderings shape what we study across the next nine weeks - so choose carefully",
  "",
  "DO:",
  "- Go round each group. One person from each group reads their chosen wondering aloud",
  "- Add each wondering to the W column of the class chart on the whiteboard",
  "- After all groups have shared, ask the class: 'Which two or three of these wonderings should we start with next lesson?'",
  "- Star the chosen wonderings on the class chart",
  "",
  "TEACHER NOTES:",
  "This share-out is the bridge into Lesson 2. The chosen wonderings become the unit's research questions. Photograph the class chart before erasing.",
  "",
  "WATCH FOR:",
  "- Quiet groups - have them nominate a different speaker if the usual one is reluctant",
  "- Wonderings that overlap - group them together rather than treating each as separate",
  "",
  "[Inquiry: Share | VTLM 2.0: Collaborative Sense-Making]",
].join("\n");

const NOTES_EXIT = [
  "SAY:",
  "- One thing on a sticky note before you leave",
  "- Finish this sentence: 'A wondering I want to find an answer to this term is...'",
  "- Stick it onto the class chart on your way out",
  "",
  "DO:",
  "- Hand each student a sticky note (or small slip of paper)",
  "- Allow 3 minutes",
  "- Collect by sticking on the W column of the class chart",
  "",
  "TEACHER NOTES:",
  "This exit ticket assesses SC2 - forming questions. The sticky note format is fast and the wonderings stay visible across the term.",
  "",
  "WATCH FOR:",
  "- Students who write a yes/no question - prompt: 'Open it up - start with why, how, what'",
  "- Students who write a fact, not a question - prompt: 'What is something you still don't know?'",
  "",
  "[Inquiry: Exit Ticket | VTLM 2.0: Evidence of Learning]",
].join("\n");

const NOTES_CLOSING = [
  "SAY:",
  "- Let's come back to our success criteria. Thumbs up, sideways, or down for each one",
  "- I can recall something I knew about how Australia has changed",
  "- I can ask a wondering question about an event in Australia's history",
  "- I can think with my group about why one event might lead to another",
  "- Reflection: which clip will stay with you the most this week, and why?",
  "",
  "DO:",
  "- Read each I can statement and pause for thumbs",
  "- Acknowledge the weight of the second clip. These are real perspectives on real events",
  "- Tell students: next lesson we start investigating the wonderings we starred together",
  "",
  "TEACHER NOTES:",
  "This launch lesson is intentionally activation-heavy. Across the next eight lessons, students dig into specific events, people, and perspectives. The Celebration Day in Week 9 is the public outcome.",
  "",
  "WATCH FOR:",
  "- Students who show thumbs down on SC2 - flag for support next lesson with sentence starters",
  "- Students who seem reflective - this is engagement, not distress; affirm their thinking",
  "",
  "[Inquiry: Closing | VTLM 2.0: Review and Reflect]",
].join("\n");

// ===============================================================
// Build function
// ===============================================================

async function build() {
  const pres = new PptxGenJS();
  pres.layout = "LAYOUT_16x9";
  pres.title = "Shaping Australia - Lesson 1 - Launch";
  pres.author = "Year 5/6 Inquiry";

  // -- Slide 1: Title --
  titleSlide(
    pres,
    "Shaping Australia",
    "How has Australia changed over time?",
    "Year 5/6 Inquiry  |  Term 2  |  Lesson 1",
    NOTES_TITLE
  );

  // -- Slide 2: Teacher Resources (per §44 - immediately after title) --
  addResourceSlide(pres, {
    resources: RESOURCE_ITEMS,
    studentTools: [
      "Mini-whiteboards (one per student)",
      "Sticky notes for exit ticket",
      "Pencils / markers for groups of three",
    ],
    boardSetup: [
      "Three-column class chart drawn on whiteboard: 'What we already know' - 'What we learnt' - 'What we wonder'",
      "Headings written before students arrive; columns roughly equal width",
    ],
    videos: [
      "BTN - Australia Since Cook (about 4 minutes)",
      "BTN - Indigenous Perspectives of Cook's Visit (about 3 minutes)",
    ],
    urls: [
      URL_BTN_COOK,
      URL_BTN_INDIG,
    ],
  }, T, FOOTER, NOTES_RESOURCES);

  // -- Slide 3: Hook / Launch --
  contentSlide(
    pres,
    "Launch",
    C.SECONDARY,
    "Then ... and now",
    [
      "Picture Australia 250 years ago",
      "Picture Australia today",
      "What's changed?",
      "How did we get from then to now?",
    ],
    NOTES_HOOK,
    FOOTER,
    (s, lg) => {
      const rX = lg.rightX;
      const rW = lg.rightW;
      const topY = lg.panelTopPadded;
      const cardH = 3.35;

      // Side-by-side "Then" and "Now" comparison cards
      const halfW = (rW - 0.18) / 2;

      // THEN card (left half, dark)
      addCard(s, rX, topY, halfW, cardH, { strip: C.PRIMARY, fill: C.BG_DARK });
      s.addText("THEN", {
        x: rX + 0.1, y: topY + 0.15, w: halfW - 0.2, h: 0.45,
        fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true,
        align: "center", margin: 0,
      });
      s.addText("about 1775", {
        x: rX + 0.1, y: topY + 0.62, w: halfW - 0.2, h: 0.30,
        fontSize: 13, fontFace: FONT_B, color: C.SUBTITLE, italic: true,
        align: "center", margin: 0,
      });
      // Stylised silhouette - sail shape
      s.addShape("triangle", {
        x: rX + halfW * 0.30, y: topY + 1.05, w: halfW * 0.40, h: 1.20,
        fill: { color: C.SUBTITLE },
      });
      s.addShape("rect", {
        x: rX + halfW * 0.46, y: topY + 1.05, w: halfW * 0.08, h: 1.45,
        fill: { color: C.ACCENT },
      });
      // Water line
      s.addShape("rect", {
        x: rX + 0.1, y: topY + 2.45, w: halfW - 0.2, h: 0.40,
        fill: { color: C.PRIMARY },
      });
      s.addText("ships  -  no roads  -  no cities", {
        x: rX + 0.1, y: topY + 2.92, w: halfW - 0.2, h: 0.30,
        fontSize: 10, fontFace: FONT_B, color: C.SUBTITLE, italic: true,
        align: "center", margin: 0,
      });

      // NOW card (right half, light)
      const nowX = rX + halfW + 0.18;
      addCard(s, nowX, topY, halfW, cardH, { strip: C.ACCENT, fill: C.BG_LIGHT });
      s.addText("NOW", {
        x: nowX + 0.1, y: topY + 0.15, w: halfW - 0.2, h: 0.45,
        fontSize: 22, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", margin: 0,
      });
      s.addText("today", {
        x: nowX + 0.1, y: topY + 0.62, w: halfW - 0.2, h: 0.30,
        fontSize: 13, fontFace: FONT_B, color: C.MUTED, italic: true,
        align: "center", margin: 0,
      });
      // Stylised city skyline - varying rectangles
      const buildings = [
        { wRatio: 0.10, hRatio: 0.65, color: C.PRIMARY },
        { wRatio: 0.13, hRatio: 0.85, color: C.SECONDARY },
        { wRatio: 0.10, hRatio: 0.50, color: C.ACCENT },
        { wRatio: 0.14, hRatio: 0.95, color: C.PRIMARY },
        { wRatio: 0.09, hRatio: 0.60, color: C.SECONDARY },
        { wRatio: 0.12, hRatio: 0.75, color: C.ACCENT },
        { wRatio: 0.11, hRatio: 0.55, color: C.PRIMARY },
      ];
      let bx = nowX + 0.15;
      const bgRowY = topY + 1.05;
      const bgRowH = 1.45;
      buildings.forEach((b) => {
        const bw = halfW * b.wRatio;
        const bh = bgRowH * b.hRatio;
        const by = bgRowY + (bgRowH - bh);
        s.addShape("rect", { x: bx, y: by, w: bw, h: bh, fill: { color: b.color } });
        bx += bw + 0.02;
      });
      // Ground line
      s.addShape("rect", {
        x: nowX + 0.1, y: topY + 2.45, w: halfW - 0.2, h: 0.40,
        fill: { color: C.MUTED },
      });
      s.addText("cities  -  millions of people  -  many cultures", {
        x: nowX + 0.1, y: topY + 2.92, w: halfW - 0.2, h: 0.30,
        fontSize: 10, fontFace: FONT_B, color: C.MUTED, italic: true,
        align: "center", margin: 0,
      });
    }
  );

  // -- Slide 4: Vocabulary --
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "Vocabulary", { color: C.PRIMARY, w: 1.7 });
    addTitle(s, "Four Inquiry Words");

    const vocab = [
      { word: "history", meaning: "the story of what happened in the past, told from evidence", color: C.PRIMARY },
      { word: "change over time", meaning: "how things become different across years and generations", color: C.SECONDARY },
      { word: "perspective", meaning: "the view someone has, based on who they are and what they've lived through", color: C.ACCENT },
      { word: "wondering", meaning: "a question we still want to find out the answer to", color: C.SUCCESS },
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
      s.addText(v.word, {
        x: x + 0.2, y: y + 0.15, w: cardW - 0.4, h: 0.50,
        fontSize: 22, fontFace: FONT_H, color: v.color, bold: true, margin: 0,
        fit: "shrink",
      });
      s.addText(v.meaning, {
        x: x + 0.2, y: y + 0.70, w: cardW - 0.4, h: cardH - 0.85,
        fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
        valign: "top",
      });
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_VOCAB);
  }

  // -- Slide 5: LI / SC --
  liSlide(
    pres,
    ["We are learning the ways in which Australia has changed over time"],
    [
      "I can recall something I knew about how Australia has changed",
      "I can ask a wondering question about an event in Australia's history",
      "I can think with my group about why one event might lead to another",
    ],
    NOTES_LI,
    FOOTER
  );

  // -- Slide 6: I Do - teacher models the KWL chart --
  contentSlide(
    pres,
    "I Do",
    C.PRIMARY,
    "Watch me add to the K column",
    [
      "Pick ONE thing I think I already know",
      "Write it in the K column out loud",
      "Use 'I think...' or 'around...' if I'm not sure",
      "Then it's your turn to brainstorm with a partner",
    ],
    NOTES_IDO,
    FOOTER,
    (s, lg) => {
      const rX = lg.rightX;
      const rW = lg.rightW;
      const topY = lg.panelTopPadded;
      const cardH = 3.35;

      // Mini KWL chart preview (right column)
      addCard(s, rX, topY, rW, cardH, { strip: C.PRIMARY, fill: C.WHITE });

      s.addText("How Australia Has Changed Over Time", {
        x: rX + 0.15, y: topY + 0.10, w: rW - 0.3, h: 0.30,
        fontSize: 11, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", margin: 0,
      });

      // 3 columns
      const tableX = rX + 0.15;
      const tableY = topY + 0.50;
      const tableW = rW - 0.3;
      const tableH = cardH - 0.65;
      const colW = tableW / 3;
      const hdrH = 0.45;

      // Header row
      const headers = [
        { label: "What we already know", color: C.SECONDARY },
        { label: "What we learnt", color: C.PRIMARY },
        { label: "What we wonder", color: C.ACCENT },
      ];
      headers.forEach((h, i) => {
        s.addShape("rect", {
          x: tableX + i * colW, y: tableY, w: colW, h: hdrH,
          fill: { color: h.color },
        });
        s.addText(h.label, {
          x: tableX + i * colW + 0.04, y: tableY, w: colW - 0.08, h: hdrH,
          fontSize: 9.5, fontFace: FONT_B, color: C.WHITE, bold: true,
          align: "center", valign: "middle", margin: 0,
        });
      });

      // Body rows
      const rowY = tableY + hdrH;
      const rowH = tableH - hdrH;
      headers.forEach((_, i) => {
        s.addShape("rect", {
          x: tableX + i * colW, y: rowY, w: colW, h: rowH,
          fill: { color: i === 0 ? "FAF5E8" : C.WHITE },
          line: { color: C.MUTED, width: 0.6 },
        });
      });

      // Sample teacher entry in K column - models uncertainty.
      // (Teacher does the live underlining on the actual whiteboard.)
      s.addText("The British arrived around 1788", {
        x: tableX + 0.06, y: rowY + 0.08, w: colW - 0.12, h: 0.50,
        fontSize: 9.5, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
        valign: "top", margin: 0,
      });
    }
  );

  // -- Slide 7: Watch Clip 1 --
  {
    const s = pres.addSlide();
    addTopBar(s, C.SECONDARY);
    addBadge(s, "Watch", { color: C.SECONDARY, w: 1.4 });
    addTitle(s, "Clip 1: Australia Since Cook");

    // Left: instruction card
    addInstructionCard(s, [
      { role: "header", text: "Watch for" },
      { role: "body", text: "ONE thing you didn't know before" },
      { role: "body", text: "We add it to the L column" },
      { role: "body", text: "Save your comments till the end" },
    ], {
      x: 0.5, y: CONTENT_TOP, w: 3.5, h: SAFE_BOTTOM - CONTENT_TOP,
      strip: C.SECONDARY, fill: C.WHITE,
      headerColor: C.SECONDARY,
    });

    // Right: video player mockup
    const rX = 4.2;
    const rW = 5.3;
    const rH = SAFE_BOTTOM - CONTENT_TOP;

    // Outer frame
    addCard(s, rX, CONTENT_TOP, rW, rH, { strip: C.PRIMARY, fill: C.BG_DARK });

    // Video viewport
    const vpX = rX + 0.25;
    const vpY = CONTENT_TOP + 0.30;
    const vpW = rW - 0.5;
    const vpH = rH - 1.20;
    s.addShape("rect", {
      x: vpX, y: vpY, w: vpW, h: vpH,
      fill: { color: "0A0A0A" },
      line: { color: C.MUTED, width: 0.8 },
    });

    // Play triangle
    const playSize = 0.7;
    s.addShape("triangle", {
      x: vpX + vpW / 2 - playSize / 2,
      y: vpY + vpH / 2 - playSize / 2,
      w: playSize, h: playSize,
      fill: { color: C.WHITE },
      rotate: 90,
    });

    // BTN label
    s.addShape("roundRect", {
      x: vpX + 0.15, y: vpY + 0.15, w: 0.7, h: 0.30, rectRadius: 0.04,
      fill: { color: C.ALERT },
    });
    s.addText("BTN", {
      x: vpX + 0.15, y: vpY + 0.15, w: 0.7, h: 0.30,
      fontSize: 11, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    // Title overlay
    s.addText("Australia Since Cook", {
      x: vpX, y: vpY + vpH - 0.55, w: vpW, h: 0.40,
      fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
    s.addText("about 4 minutes  |  ABC Education", {
      x: vpX, y: vpY + vpH - 0.20, w: vpW, h: 0.20,
      fontSize: 9, fontFace: FONT_B, color: C.SUBTITLE, italic: true,
      align: "center", margin: 0,
    });

    // Link strip (clickable)
    const linkY = vpY + vpH + 0.15;
    s.addShape("roundRect", {
      x: vpX, y: linkY, w: vpW, h: 0.42, rectRadius: 0.06,
      fill: { color: C.ACCENT },
    });
    s.addText("Open BTN clip", {
      x: vpX, y: linkY, w: vpW, h: 0.42,
      fontSize: 12, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
      hyperlink: { url: URL_BTN_COOK, tooltip: "Open BTN - Australia Since Cook" },
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_VIDEO1);
  }

  // -- Slide 8: Watch Clip 2 --
  {
    const s = pres.addSlide();
    addTopBar(s, C.SECONDARY);
    addBadge(s, "Watch", { color: C.SECONDARY, w: 1.4 });
    addTitle(s, "Clip 2: Indigenous Perspectives");

    // Left: instruction card
    addInstructionCard(s, [
      { role: "header", text: "Watch for" },
      { role: "body", text: "What is DIFFERENT to Clip 1?" },
      { role: "body", text: "Same event - different perspective" },
      { role: "body", text: "Notice the words and feelings" },
    ], {
      x: 0.5, y: CONTENT_TOP, w: 3.5, h: SAFE_BOTTOM - CONTENT_TOP,
      strip: C.SECONDARY, fill: C.WHITE,
      headerColor: C.SECONDARY,
    });

    // Right: video player mockup
    const rX = 4.2;
    const rW = 5.3;
    const rH = SAFE_BOTTOM - CONTENT_TOP;

    addCard(s, rX, CONTENT_TOP, rW, rH, { strip: C.ACCENT, fill: C.BG_DARK });

    const vpX = rX + 0.25;
    const vpY = CONTENT_TOP + 0.30;
    const vpW = rW - 0.5;
    const vpH = rH - 1.20;
    s.addShape("rect", {
      x: vpX, y: vpY, w: vpW, h: vpH,
      fill: { color: "0A0A0A" },
      line: { color: C.MUTED, width: 0.8 },
    });

    const playSize = 0.7;
    s.addShape("triangle", {
      x: vpX + vpW / 2 - playSize / 2,
      y: vpY + vpH / 2 - playSize / 2,
      w: playSize, h: playSize,
      fill: { color: C.WHITE },
      rotate: 90,
    });

    s.addShape("roundRect", {
      x: vpX + 0.15, y: vpY + 0.15, w: 0.7, h: 0.30, rectRadius: 0.04,
      fill: { color: C.ALERT },
    });
    s.addText("BTN", {
      x: vpX + 0.15, y: vpY + 0.15, w: 0.7, h: 0.30,
      fontSize: 11, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    s.addText("Indigenous Perspectives of Cook's Visit", {
      x: vpX, y: vpY + vpH - 0.55, w: vpW, h: 0.40,
      fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
    s.addText("about 3 minutes  |  ABC Education", {
      x: vpX, y: vpY + vpH - 0.20, w: vpW, h: 0.20,
      fontSize: 9, fontFace: FONT_B, color: C.SUBTITLE, italic: true,
      align: "center", margin: 0,
    });

    const linkY = vpY + vpH + 0.15;
    s.addShape("roundRect", {
      x: vpX, y: linkY, w: vpW, h: 0.42, rectRadius: 0.06,
      fill: { color: C.ACCENT },
    });
    s.addText("Open BTN clip", {
      x: vpX, y: linkY, w: vpW, h: 0.42,
      fontSize: 12, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
      hyperlink: { url: URL_BTN_INDIG, tooltip: "Open BTN - Indigenous Perspectives of Cook's Visit" },
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_VIDEO2);
  }

  // -- Slide 9 / 9a: CFU hinge with reveal --
  const cfuQText = "When two people tell the story of the same event differently, we call those different ___.";

  function buildCfuBase() {
    const s = pres.addSlide();
    addTopBar(s, C.ALERT);
    addBadge(s, "CFU", { color: C.ALERT });
    addTitle(s, "What's the inquiry word?", { color: C.ALERT });

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

    // Reserve answer-bar zone at the bottom
    const aH = 0.55;
    const aY = SAFE_BOTTOM - aH - 0.05;
    const qY = CONTENT_TOP + 0.56;
    const qH = aY - 0.20 - qY;

    addCard(s, 0.5, qY, 9, qH, { strip: C.ALERT, fill: C.WHITE });
    s.addText(cfuQText, {
      x: 0.75, y: qY + 0.30, w: 8.5, h: qH - 0.60,
      fontSize: 22, fontFace: FONT_B, color: C.CHARCOAL, valign: "middle",
      align: "center", margin: 0,
      fit: "shrink", shrinkText: true,
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
      const aH = 0.55;
      const aY = SAFE_BOTTOM - aH - 0.05;
      s.addShape("roundRect", {
        x: 0.5, y: aY, w: 9, h: aH, rectRadius: 0.1,
        fill: { color: C.SUCCESS },
      });
      s.addText("perspectives  -  the view someone has, shaped by who they are", {
        x: 0.5, y: aY, w: 9, h: aH,
        fontSize: 16, fontFace: FONT_B, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      s.addNotes(NOTES_CFU_A);
    }
  );

  // -- Slide 10: You Do - groups of 3 fill the chart --
  contentSlide(
    pres,
    "You Do",
    C.ACCENT,
    "Your group's KWL chart",
    [
      "Groups of 3 - one chart per group",
      "TWO things in the K column (what you thought you knew)",
      "THREE things in the L column (from the clips)",
      "TWO wonderings in the W column - start with why, how, what",
    ],
    NOTES_YOUDO,
    FOOTER,
    (s, lg) => {
      const rX = lg.rightX;
      const rW = lg.rightW;
      const topY = lg.panelTopPadded;
      const cardH = 3.35;

      // Mini chart preview
      addCard(s, rX, topY, rW, cardH, { strip: C.ACCENT, fill: C.WHITE });

      s.addText("How Australia Has Changed Over Time", {
        x: rX + 0.15, y: topY + 0.10, w: rW - 0.3, h: 0.30,
        fontSize: 11, fontFace: FONT_H, color: C.ACCENT, bold: true,
        align: "center", margin: 0,
      });

      const tableX = rX + 0.15;
      const tableY = topY + 0.50;
      const tableW = rW - 0.3;
      const tableH = cardH - 0.65;
      const colW = tableW / 3;
      const hdrH = 0.45;

      const headers = [
        { label: "K", sub: "We know", color: C.SECONDARY, target: "x 2" },
        { label: "L", sub: "We learnt", color: C.PRIMARY, target: "x 3" },
        { label: "W", sub: "We wonder", color: C.ACCENT, target: "x 2" },
      ];
      headers.forEach((h, i) => {
        s.addShape("rect", {
          x: tableX + i * colW, y: tableY, w: colW, h: hdrH,
          fill: { color: h.color },
        });
        s.addText(h.label, {
          x: tableX + i * colW, y: tableY, w: colW, h: 0.26,
          fontSize: 13, fontFace: FONT_H, color: C.WHITE, bold: true,
          align: "center", valign: "middle", margin: 0,
        });
        s.addText(h.sub + "  -  " + h.target, {
          x: tableX + i * colW, y: tableY + 0.24, w: colW, h: 0.20,
          fontSize: 8, fontFace: FONT_B, color: C.WHITE,
          align: "center", valign: "middle", margin: 0,
        });
      });

      const rowY = tableY + hdrH;
      const rowH = tableH - hdrH;
      headers.forEach((_, i) => {
        s.addShape("rect", {
          x: tableX + i * colW, y: rowY, w: colW, h: rowH,
          fill: { color: "FAF5E8" },
          line: { color: C.MUTED, width: 0.6 },
        });
        // Lined writing area
        const lineCount = 4;
        const lineGap = (rowH - 0.20) / lineCount;
        for (let li = 0; li < lineCount; li++) {
          const ly = rowY + 0.30 + li * lineGap;
          s.addShape("line", {
            x: tableX + i * colW + 0.10, y: ly,
            w: colW - 0.20, h: 0,
            line: { color: C.MUTED, width: 0.5 },
          });
        }
      });
    }
  );

  // -- Slide 11: Discuss / Share wonderings --
  pairShareSlide(
    pres,
    "Share your group's strongest wondering",
    [
      "Each group: which ONE wondering would you most like the class to investigate?",
      "Why is this wondering important to your group?",
      "Which wonderings should we star and start with next lesson?",
    ],
    NOTES_DISCUSS,
    FOOTER
  );

  // -- Slide 12: Exit Ticket --
  exitTicketSlide(
    pres,
    [
      "Sticky note - finish this sentence and stick it on the class chart on your way out:\n\n'A wondering I want to find an answer to this term is ...'",
    ],
    NOTES_EXIT,
    FOOTER,
    { assessesSc: 2, title: "One wondering before you leave" }
  );

  // -- Slide 13: Closing --
  closingSlide(
    pres,
    {
      reflectionPrompt: "Which clip will stay with you the most this week, and why?",
      scItems: [
        "I can recall something I knew about how Australia has changed",
        "I can ask a wondering question about an event in Australia's history",
        "I can think with my group about why one event might lead to another",
      ],
      selfAssessment: {
        prompt: "Thumbs up, sideways, or down for each one",
        options: ["Got it", "Getting there", "Need more"],
      },
    },
    NOTES_CLOSING
  );

  // -- Write PPTX --
  fs.mkdirSync(LESSON_FOLDER, { recursive: true });
  await pres.writeFile({ fileName: path.join(LESSON_FOLDER, PPTX_NAME) });
  console.log("PPTX written to", path.join(LESSON_FOLDER, PPTX_NAME));

  // ===============================================================
  // PDF: Group KWL Chart
  // ===============================================================
  {
    // Portrait A4 - createPdf does not currently expose the layout option,
    // so we use the default. Three columns at ~6 cm each works for a
    // group-of-three handwritten chart with 9 writing lines per column.
    const doc = createPdf({ title: "Group KWL Chart" });
    let y = addPdfHeader(doc, "Shaping Australia - KWL Chart", {
      subtitle: "How Australia has changed over time. One chart per group of three.",
      color: C.PRIMARY,
      lessonInfo: "Year 5/6 Inquiry  |  Shaping Australia  |  Lesson 1",
    });

    // Group / names line
    doc.fontSize(11).font("Sans-Bold").fillColor(hex(C.CHARCOAL));
    doc.text("Group members:", PAGE.MARGIN, y);
    doc.save();
    doc.moveTo(PAGE.MARGIN + 110, y + 14)
      .lineTo(PAGE.MARGIN + PAGE.CONTENT_W, y + 14)
      .strokeColor("#000000").lineWidth(0.9).stroke();
    doc.restore();
    y += 28;

    // Three column table
    const tableX = PAGE.MARGIN;
    const tableW = PAGE.CONTENT_W;
    const colW = tableW / 3;
    const hdrH = 56;
    const tableYStart = y;
    const tableBottom = PAGE.H - PAGE.MARGIN - 30;
    const bodyH = tableBottom - (tableYStart + hdrH);

    // Headers - keep labels short so they don't wrap inside the narrow column
    const hdrs = [
      { letter: "K", label: "What we know", color: C.SECONDARY, target: "TWO things" },
      { letter: "L", label: "What we learnt", color: C.PRIMARY, target: "THREE things" },
      { letter: "W", label: "What we wonder", color: C.ACCENT, target: "TWO wonderings" },
    ];

    hdrs.forEach((h, i) => {
      const cx = tableX + i * colW;
      doc.save();
      doc.rect(cx, tableYStart, colW, hdrH).fill(hex(h.color));
      doc.restore();
      // Big letter on the left
      doc.fontSize(22).font("Sans-Bold").fillColor("#FFFFFF");
      doc.text(h.letter, cx + 10, tableYStart + 14, { width: 24, align: "left" });
      // Label and target right of the letter
      doc.fontSize(12).font("Sans-Bold").fillColor("#FFFFFF");
      doc.text(h.label, cx + 38, tableYStart + 12, { width: colW - 46, align: "left" });
      doc.fontSize(9.5).font("Sans").fillColor("#FFFFFF");
      doc.text(h.target, cx + 38, tableYStart + 32, { width: colW - 46, align: "left" });
    });

    // Body cells
    hdrs.forEach((_, i) => {
      const cx = tableX + i * colW;
      doc.save();
      doc.rect(cx, tableYStart + hdrH, colW, bodyH)
        .lineWidth(0.9).strokeColor("#000000").stroke();
      doc.restore();

      // Writing lines
      const lineCount = 9;
      const lineGap = bodyH / (lineCount + 1);
      doc.save();
      doc.strokeColor("#B0B0B0").lineWidth(0.5);
      for (let li = 1; li <= lineCount; li++) {
        const ly = tableYStart + hdrH + li * lineGap;
        doc.moveTo(cx + 12, ly).lineTo(cx + colW - 12, ly).stroke();
      }
      doc.restore();
    });

    // Footer tip
    addPdfFooter(doc, "Year 5/6 Inquiry  |  Shaping Australia  |  Lesson 1 - Group KWL Chart");
    const outPath = path.join(RES_DIR, "Session 1 Group KWL Chart.pdf");
    await writePdf(doc, outPath);
    console.log("PDF written:", outPath);
  }

  console.log("\nBuild complete!");
  console.log("Output folder:", LESSON_FOLDER);
}

build().catch((err) => { console.error(err); process.exit(1); });
