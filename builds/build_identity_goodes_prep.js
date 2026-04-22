"use strict";

// Identity, Labels & Standing Up — Grade 5/6 Inquiry
// Single 60-minute session preparing students for the Adam Goodes documentary
// Combines identity/labels/bystander-upstander/media influence themes

const path = require("path");
const fs = require("fs");
const PptxGenJS = require("pptxgenjs");
const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf,
  addPdfHeader, addSectionHeading, addBodyText, addTipBox,
  addPdfFooter, addLinedArea, addWriteLine,
  addResourceSlide, makeSessionResource, formatSessionResourceFileName,
  getSessionResourceFolder, addTwoColumnOrganiser,
  PAGE, hex,
} = require("../themes/pdf_helpers");

// ── Theme ──
const T = createTheme("inquiry", "grade56", weekToVariant(5));
const {
  C, FONT_H, FONT_B,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  investigationSlide, pairShareSlide,
  withReveal,
  addTopBar, addBadge, addTitle, addCard, addFooter, addTextOnShape,
  addInstructionCard,
  CONTENT_TOP, SAFE_BOTTOM, SLIDE_W,
  runSlideDiagnostics,
} = T;

// ── Output paths ──
const UNIT = "Identity_Goodes_Documentary_Prep";
const LESSON_FOLDER = path.join(__dirname, "..", "output", UNIT);
const PPTX_NAME = "Identity Labels and Standing Up.pptx";
const FOOTER = "Identity, Labels & Standing Up | Grade 5/6 Inquiry";
const SESSION = 1;
const RES_DIR = path.join(LESSON_FOLDER, getSessionResourceFolder(SESSION));

// ── Resources ──
const REFLECTION_RESOURCE = makeSessionResource(
  SESSION,
  "Labels and Identity Reflection",
  "Graphic organiser for exploring labels, perceptions, and identity."
);
const VIEWING_GUIDE_RESOURCE = makeSessionResource(
  SESSION,
  "Documentary Viewing Guide",
  "Structured viewing guide for the Adam Goodes documentary."
);
const RESOURCE_ITEMS = [REFLECTION_RESOURCE, VIEWING_GUIDE_RESOURCE];

fs.mkdirSync(RES_DIR, { recursive: true });

// ═══════════════════════════════════════════════════════════════
// Teacher Notes
// ═══════════════════════════════════════════════════════════════

const NOTES_TITLE = [
  "SAY:",
  "- Today we are exploring identity, labels, and what it means to stand up for others",
  "- Everything we do today is building toward something important -- next session we will be watching a documentary about Adam Goodes",
  "- To get the most out of that film, we need to understand some big ideas first",
  "",
  "DO:",
  "- Display slide as students settle",
  "- Have inquiry books on desks",
  "",
  "TEACHER NOTES:",
  "This session combines identity, labels/perceptions, bystander-upstander, and media influence as preparation for the Adam Goodes documentary (The Final Quarter or The Australian Dream). The session moves through connected activities that build conceptual readiness for the documentary's themes.",
  "",
  "WATCH FOR:",
  "- Students who may already know about Adam Goodes -- acknowledge briefly but do not spoil the documentary",
  "- Students settling into inquiry mode after a subject transition",
  "",
  "[General: Title | VTLM 2.0: Establishing Purpose]",
].join("\n");

const NOTES_LI = [
  "SAY:",
  "- Read the learning intention together from the slide",
  "- Read through each success criterion",
  "- SC1 is about understanding how labels work. SC2 is about knowing what bystanders and upstanders are. SC3 is about thinking carefully about what media tells us",
  "",
  "DO:",
  "- Choral read the LI, then each SC",
  "- Brief check: Thumbs up if you have ever heard someone described using a label they did not choose for themselves",
  "",
  "TEACHER NOTES:",
  "The three SC map directly to the documentary's major themes: identity/labels (Goodes' experience of racial labelling), bystander/upstander (crowd and media response), and media influence (how media framed the story). SC1 is the floor -- every student can engage with labels.",
  "",
  "WATCH FOR:",
  "- Students unsure about 'labels' in this context -- clarify: words or categories people use to describe others, sometimes without their permission",
  "- Students who associate labels only with physical objects -- redirect to social meaning",
  "",
  "[General: LI/SC | VTLM 2.0: Clear Learning Intention]",
].join("\n");

const NOTES_HOOK = [
  "SAY:",
  "- I am going to show you a photo of someone famous. When you see them, think about what words or labels you would use to describe this person",
  "- Ask: What words come to mind? [Accept all responses -- athlete, champion, role model, Indigenous, tall, etc.]",
  "- Now think about yourself. In your inquiry books, write three quick responses: What words would you use to describe yourself? What words might others use to describe you? Are there words others might use that you would NOT choose for yourself?",
  "",
  "DO:",
  "- Display the image on slide (teacher to select a famous person students will recognise)",
  "- Cold Call 4-5 students for their labels about the famous person",
  "- Allow 2-3 minutes for the personal writing in inquiry books",
  "- Do not force sharing of personal labels -- this is private reflection",
  "",
  "TEACHER NOTES:",
  "This hook activates prior knowledge about labels and primes the distinction between self-chosen and externally imposed identity. The famous person activity is lower-stakes; the personal reflection builds vulnerability gradually. Choose a figure students will have strong, varied perceptions of.",
  "",
  "WATCH FOR:",
  "- Students who only write positive self-labels -- prompt: 'Are there any words others might use that surprise you or that you would not pick?'",
  "- Students who seem uncomfortable with the personal reflection -- reassure this stays in their books",
  "",
  "[General: Hook | VTLM 2.0: Prior Knowledge Activation]",
].join("\n");

const NOTES_IDENTITY_IDO = [
  "SAY:",
  "- Our identity -- who we are -- is shaped by two forces. There is how we see ourselves, and there is how others see us",
  "- Sometimes those match up. You might describe yourself as funny, and your friends agree",
  "- But sometimes they do not match. You might see yourself as shy, but others label you as rude or stuck-up -- just because you are quiet",
  "- Labels that others put on us can be powerful. They can change how we feel about ourselves, especially if we hear them enough",
  "- Ask: Can you think of a time when someone described you in a way that did not match how you see yourself? [Allow a few volunteers, do not force]",
  "",
  "DO:",
  "- Point to the two-part visual on the slide: 'How I see myself' vs 'How others see me'",
  "- Model one example: 'A person might see themselves as hardworking, but if they come from a different background, others might label them as different or an outsider'",
  "",
  "TEACHER NOTES:",
  "This I Do introduces the core concept of identity as a negotiation between self-perception and external perception. Keep examples relatable but prepare students for the documentary's heavier examples of imposed identity. The modelled example deliberately hints at the Goodes themes without naming him.",
  "",
  "WATCH FOR:",
  "- Students who share deeply personal examples -- acknowledge warmly and briefly, do not dwell or probe",
  "- Students who seem to equate labels only with name-calling -- broaden: labels can be assumptions, stereotypes, or categories too",
  "",
  "[General: I Do | VTLM 2.0: Explicit Explanation]",
].join("\n");

const NOTES_LABELS_WEDO = [
  "SAY:",
  "- Let's practise thinking about labels together",
  "- I am going to show you some scenarios. With your partner, discuss: What labels are being used? Who is putting the label on? How might the person being labelled feel?",
  "- Scenario 1: A student who is new to Australia gets called 'the foreign kid' even though they have lived here for three years",
  "- Scenario 2: A girl who loves maths is told by classmates 'Girls are not good at maths'",
  "",
  "DO:",
  "- Display scenarios on slide",
  "- Allow 60 seconds per scenario for Turn and Talk",
  "- Cold Call 2-3 pairs after each scenario to share their thinking",
  "",
  "CFU CHECKPOINT:",
  "Technique: Turn and Talk with Cold Call follow-up",
  "Script:",
  "- Say: Turn to your partner. For Scenario 1, who is putting the label on, and how might the person feel? You have 60 seconds.",
  "- After 60 seconds, Cold Call 2-3 pairs: What did you discuss?",
  "- Scan for: students identifying the labeller (classmates), the impact (feeling excluded, not belonging, frustrated), and the mismatch between the label and reality (lived here 3 years).",
  "PROCEED: If 80%+ can identify the labeller, the impact, and the mismatch, move to bystander/upstander.",
  "PIVOT: If students struggle to articulate impact, reteach: 'Imagine hearing that label every day for a year. What starts to happen? You might start believing it. That is the danger of labels -- they can reshape how someone sees themselves.' Re-check with Scenario 2.",
  "",
  "ENABLING & EXTENDING:",
  "ENABLING PROMPT:",
  "- Task: Students work with a simpler scenario -- 'Someone calls a quiet student bossy because they reminded the class about the rules.' Identify: Who labelled? What was the label? Was it fair?",
  "EXTENDING PROMPT:",
  "- Task: Students think of a time a positive label became limiting (e.g., 'the smart kid' feeling pressure to never make mistakes). Write 2-3 sentences in their inquiry book.",
  "",
  "TEACHER NOTES:",
  "This is the We Do for identity/labels. Scenarios are deliberately school-based and relatable. The cultural label scenario primes students for the documentary's racial identity themes without being heavy-handed.",
  "",
  "WATCH FOR:",
  "- Partners who only discuss the 'obvious' label without considering the impact -- prompt: 'But how does it feel?'",
  "- Students who dismiss the scenarios as not serious -- redirect: 'Even small labels repeated over time can shape identity'",
  "- Readiness signal: students articulating both the label AND its emotional/identity impact",
  "",
  "[General: We Do | VTLM 2.0: Scaffold Practice]",
].join("\n");

const NOTES_BYSTANDER = [
  "SAY:",
  "- Now let's connect labels to action. When someone is being labelled unfairly, other people have a choice",
  "- A bystander is someone who sees something wrong happening and does nothing. They stand by",
  "- An upstander is someone who sees something wrong and takes action. They stand up",
  "- Ask: Why do you think people sometimes choose to be bystanders even when they know something is wrong? [They want to belong, they are scared of being targeted too, they do not know what to do]",
  "- That is a really important insight. Sometimes wanting to belong to a group makes us stay silent when we should speak up",
  "",
  "DO:",
  "- Point to the bystander/upstander definitions on the slide",
  "- Allow 30 seconds of think time before Cold Calling for the question",
  "- Accept multiple answers -- there is no single reason",
  "",
  "TEACHER NOTES:",
  "This I Do introduces bystander/upstander as a direct extension of the labels discussion. The question about WHY people stay silent is critical preparation for the documentary -- Goodes experienced widespread bystander behaviour from fans, media, and the AFL itself.",
  "",
  "WATCH FOR:",
  "- Students who say 'I would always stand up' -- gently challenge: 'That is a great goal. But it is harder than it sounds, especially when a whole group is involved'",
  "- Students connecting this to real school experiences -- affirm and note for sensitivity",
  "",
  "[General: I Do | VTLM 2.0: Explicit Explanation]",
].join("\n");

const NOTES_CFU_BYSTANDER = [
  "SAY:",
  "- Quick check. I am going to read a scenario. On your fingers, show me: 1 finger for bystander, 2 fingers for upstander",
  "- Scenario: During a footy game, the crowd starts booing one player every time he touches the ball. A person in the crowd feels uncomfortable but keeps clapping along with everyone else",
  "- Show me your fingers now",
  "",
  "DO:",
  "- Read the scenario clearly",
  "- Students show fingers (1 = bystander, 2 = upstander)",
  "- Scan the room quickly",
  "",
  "CFU CHECKPOINT:",
  "Technique: Finger Voting",
  "Script:",
  "- Say: Listen to this scenario. Show me 1 finger for bystander, 2 for upstander. Ready?",
  "- Read the scenario. Say: Show me now.",
  "- Scan for: all students showing 1 finger (bystander -- the person does nothing despite discomfort).",
  "PROCEED: If 80%+ correctly identify this as bystander behaviour, move to media influence.",
  "PIVOT: If students are split, clarify: 'The key question is: did they take action? Feeling uncomfortable is not the same as standing up. A bystander can know something is wrong but still do nothing. That is what makes bystander behaviour so common -- and so important to recognise.' Re-check: 'Same scenario, but now the person says to friends: That booing is not right. We should stop. Bystander or upstander?' [Upstander]",
  "",
  "TEACHER NOTES:",
  "The scenario is deliberately close to what happened with Goodes without naming him. This builds the conceptual bridge so when students see the documentary, they can apply this framework. The distinction between feeling uncomfortable and taking action is the key insight.",
  "",
  "MISCONCEPTIONS:",
  "- Misconception: A bystander is someone who does not notice what is happening.",
  "  Why: Students conflate ignorance with inaction. In reality, bystanders often notice and feel uncomfortable but choose not to act.",
  "  Impact: If students think bystanders are unaware, they cannot recognise bystander behaviour in themselves or the documentary.",
  "  Quick correction: 'A bystander knows something is wrong -- that is what makes it a choice. If you do not notice, you are just not involved. If you notice and do nothing, that is bystander behaviour.'",
  "",
  "WATCH FOR:",
  "- Students who show 2 fingers because 'they felt uncomfortable' -- key teaching moment about action vs feeling",
  "- Students connecting the footy scenario to real events -- acknowledge without confirming the Goodes link yet",
  "",
  "[General: CFU | VTLM 2.0: Monitor Progress]",
].join("\n");

const NOTES_MEDIA = [
  "SAY:",
  "- Now let's talk about one of the biggest label-makers in our world: the media",
  "- Media includes news, social media, blogs, YouTube, and more. Media does not just report what happens -- it shapes how we see people and events",
  "- Ask: If a news channel shows the same person in a negative way every single night, what starts to happen? [People start believing the negative story, they form opinions without knowing the full picture]",
  "- Not all media is equally reliable. Some sources check facts carefully. Others want clicks, views, or reactions more than truth",
  "",
  "DO:",
  "- Point to the key points on the slide",
  "- Allow 20 seconds of think time for the question before Cold Calling",
  "- Brief whole-class response",
  "",
  "TEACHER NOTES:",
  "This connects labels to a systemic level. The media discussion is essential preparation for the documentary, which shows how media framing contributed to the vilification of Goodes. Keep it general here -- the documentary will provide the specific case study.",
  "",
  "WATCH FOR:",
  "- Students who trust all media equally -- note for the next activity",
  "- Students who distrust all media -- both extremes need nuancing",
  "",
  "[General: I Do | VTLM 2.0: Explicit Explanation]",
].join("\n");

const NOTES_MEDIA_WEDO = [
  "SAY:",
  "- Let's think critically about media together. I am going to show you some types of media. With your partner, discuss: Is this a reliable source of information? Why or why not?",
  "- Remember: no source is 100% reliable all the time. The question is how much checking and evidence a source uses",
  "",
  "DO:",
  "- Display the media types on the slide",
  "- Point to each one and allow 30 seconds of partner discussion",
  "- Cold Call pairs for their reasoning after every 2-3 items",
  "- Land the key message: We must always read and view critically, no matter the source",
  "",
  "ENABLING & EXTENDING:",
  "ENABLING PROMPT:",
  "- Task: Students sort media types into two columns in their inquiry book: 'More likely to check facts' and 'Less likely to check facts'. Use the list on the slide as a guide.",
  "EXTENDING PROMPT:",
  "- Task: Students write a short paragraph explaining how media could create a false perception of a person. Use a hypothetical example -- e.g., 'Imagine a news channel only ever showed one student making mistakes...'",
  "",
  "TEACHER NOTES:",
  "This We Do builds media literacy as documentary preparation. Students need this critical lens to understand how media coverage shaped public perception of Goodes. The no-source-is-perfect message prevents black-and-white thinking.",
  "",
  "WATCH FOR:",
  "- Students who fixate on 'fake news' as the only problem -- broaden: even real news can be biased in what it chooses to show",
  "- Students engaging in genuine debate about specific sources -- encourage this",
  "- Readiness signal: students reasoning about WHY a source might be more or less reliable, not just guessing",
  "",
  "[General: We Do | VTLM 2.0: Scaffold Practice]",
].join("\n");

const NOTES_YOUDO = [
  "SAY:",
  "- We have explored labels, bystanders, upstanders, and media today. Next session, we are watching a documentary about Adam Goodes -- an Indigenous Australian AFL champion",
  "- His story connects to everything we have discussed today: identity, the labels others put on him, how bystanders and upstanders responded, and the role media played",
  "- Before we watch, I want you to write a reflection in your inquiry book",
  "- First: Write one example of a label that others use that does not match who the person really is",
  "- Next: Write what an upstander could do if they saw someone being labelled unfairly",
  "- Then: Write one question you want to keep in mind while watching the documentary",
  "",
  "DO:",
  "- Display the three reflection prompts on the slide",
  "- Distribute the Session 1 Labels and Identity Reflection if using the handout",
  "- Allow 8-10 minutes for individual writing",
  "- Circulate and support. Check that students are engaging with all three prompts",
  "- In the final 2 minutes, invite 2-3 volunteers to share one of their responses",
  "",
  "ENABLING & EXTENDING:",
  "ENABLING PROMPT:",
  "- Task: Students answer just the first two prompts using sentence starters: 'A label that does not match is when someone is called ___ but they are really ___' and 'An upstander could ___'.",
  "- Extra Notes: The sentence starters scaffold the thinking without reducing the cognitive demand of the concepts.",
  "EXTENDING PROMPT:",
  "- Task: Students write a fourth response: 'How might media make bystander behaviour more likely?' connecting all three themes (labels, bystander/upstander, media influence) in a short paragraph.",
  "",
  "TEACHER NOTES:",
  "This You Do asks students to synthesise all three session themes into personal reflection. The documentary preparation framing gives purpose -- students are not just reflecting for reflection's sake, they are building a lens for viewing. Different content from We Do activities (personal reflection vs scenario analysis).",
  "",
  "SENSITIVITY ADVISORY:",
  "- What it is: Students may connect labels to their own experiences of discrimination, exclusion, or bullying.",
  "- Framing language: 'This is private reflection. You can write as much or as little as you feel comfortable with. If anything today has brought up difficult feelings, you can talk to me or to [wellbeing staff member] after class.'",
  "- Watch for: Students who become visibly upset, withdraw, or stop writing suddenly.",
  "- Protocol: Do not require them to continue. Offer a quiet check-in after class. Follow the school's wellbeing referral process if needed.",
  "",
  "WATCH FOR:",
  "- Students writing very surface-level responses ('bullying is bad') -- prompt: 'Can you give a specific example? What would an upstander actually say or do?'",
  "- Students who finish quickly -- direct them to the extending prompt about media and bystander behaviour",
  "- Readiness signal: students writing reflections that connect labels to impact and identify specific upstander actions",
  "",
  "[General: You Do | VTLM 2.0: Supported Application]",
].join("\n");

const NOTES_CLOSING = [
  "SAY:",
  "- Let's check in on our success criteria",
  "- SC1: I can explain how labels and perceptions can shape a person's identity. Thumbs up, sideways, or down",
  "- SC2: I can describe the difference between a bystander and an upstander. Thumbs?",
  "- SC3: I can think critically about how media influences our perception of others. Thumbs?",
  "- Next session, we will watch the Adam Goodes documentary. As you watch, look for labels, look for bystanders and upstanders, and think about the role media played. The ideas from today are your viewing lens",
  "",
  "DO:",
  "- Display SC on slide. Read each one. Students show thumbs for each",
  "- Note students showing thumbs-down on any SC for documentary viewing support",
  "- Final Turn and Talk prompt on slide",
  "",
  "TEACHER NOTES:",
  "The closing frames the documentary as the next step in the learning journey. Students who are solid on all three SC will get more from the viewing. Students who are shaky on SC2 or SC3 may benefit from the Session 1 Documentary Viewing Guide as a scaffold during the film.",
  "",
  "WATCH FOR:",
  "- Students consistently at thumbs-sideways on SC2 -- they may need the bystander/upstander definitions visible during the documentary",
  "- Students showing genuine curiosity about Adam Goodes -- this is the ideal hook for next session",
  "",
  "[General: Closing | VTLM 2.0: Review and Reflection]",
].join("\n");

const NOTES_RESOURCES = [
  "SAY:",
  "- Here are the resources for today's session and for the documentary viewing",
  "",
  "DO:",
  "- Print Session 1 Labels and Identity Reflection -- one per student (optional: can use inquiry books instead)",
  "- Print Session 1 Documentary Viewing Guide -- one per student for the documentary viewing session",
  "",
  "TEACHER NOTES:",
  "The reflection handout supports the You Do activity. The documentary viewing guide is for NEXT session -- print in advance. It provides structured prompts students complete while watching, connecting what they see to today's concepts.",
  "",
  "WATCH FOR:",
  "- Ensure the viewing guide is printed before the documentary session",
  "",
  "[General: Resources | VTLM 2.0: Preparation]",
].join("\n");

// ═══════════════════════════════════════════════════════════════
// Build
// ═══════════════════════════════════════════════════════════════

async function build() {
  const pres = new PptxGenJS();
  pres.layout = "LAYOUT_16x9";

  // ================================================================
  // SLIDE 1: Title
  // ================================================================
  titleSlide(
    pres,
    "Identity, Labels\n& Standing Up",
    "Preparing for the Adam Goodes Documentary",
    "Grade 5/6 Inquiry",
    NOTES_TITLE
  );

  // ================================================================
  // SLIDE 2: LI / SC
  // ================================================================
  liSlide(
    pres,
    ["We are learning how identity, labels, and the actions of bystanders and upstanders shape people's experiences"],
    [
      "I can explain how labels and perceptions can shape a person's identity",
      "I can describe the difference between a bystander and an upstander",
      "I can think critically about how media influences our perception of others",
    ],
    NOTES_LI,
    FOOTER
  );

  // ================================================================
  // SLIDE 3: Hook — Labels Activity
  // ================================================================
  contentSlide(
    pres,
    "Think",
    C.PRIMARY,
    "What Labels Do We Use?",
    [
      "Look at this person. What words would you use to describe them?",
      "Now think about yourself:",
      "What words would YOU use to describe yourself?",
      "What words might OTHERS use to describe you?",
      "Are there words others use that you would NOT choose?",
    ],
    NOTES_HOOK,
    FOOTER,
    (slide, layoutGuide) => {
      // Right panel: reflection prompts card
      const rx = layoutGuide.rightX;
      const rw = layoutGuide.rightW;
      const topY = layoutGuide.panelTopPadded;

      addCard(slide, rx, topY, rw, 3.0, { fill: C.BG_CARD, strip: C.ACCENT });
      slide.addText("In Your Inquiry Book", {
        x: rx + 0.15, y: topY + 0.1, w: rw - 0.3, h: 0.35,
        fontSize: 13, fontFace: FONT_H, color: C.ACCENT, bold: true, margin: 0,
      });

      const prompts = [
        "1.  Words I use to describe myself",
        "2.  Words others use to describe me",
        "3.  Words others use that I would NOT choose",
      ];
      prompts.forEach((p, i) => {
        slide.addText(p, {
          x: rx + 0.2, y: topY + 0.55 + i * 0.7, w: rw - 0.4, h: 0.6,
          fontSize: 12.5, fontFace: FONT_B, color: C.CHARCOAL, valign: "middle", margin: 0,
        });
      });
    }
  );

  // ================================================================
  // SLIDE 4: I Do — Identity & Perception
  // ================================================================
  contentSlide(
    pres,
    "I Do",
    C.PRIMARY,
    "Identity: Two Forces",
    [
      "How I see myself",
      "How others see me",
      "Sometimes these match -- sometimes they do not",
      "Labels others put on us can change how we see ourselves",
    ],
    NOTES_IDENTITY_IDO,
    FOOTER,
    (slide, layoutGuide) => {
      const rx = layoutGuide.rightX;
      const rw = layoutGuide.rightW;
      const topY = layoutGuide.panelTopPadded;

      // "How I see myself" circle
      const circleW = 1.8;
      const circleH = 1.2;
      const cx = rx + (rw - circleW) / 2;

      addTextOnShape(slide, "How I see myself", {
        x: cx, y: topY + 0.1, w: circleW, h: circleH, rectRadius: 0.15,
        fill: { color: C.SUCCESS },
      }, { fontSize: 13, fontFace: FONT_B, color: C.WHITE, bold: true });

      // Overlap indicator
      slide.addText("vs", {
        x: cx, y: topY + 1.35, w: circleW, h: 0.35,
        fontSize: 14, fontFace: FONT_H, color: C.MUTED, align: "center", valign: "middle", margin: 0,
      });

      // "How others see me" circle
      addTextOnShape(slide, "How others see me", {
        x: cx, y: topY + 1.7, w: circleW, h: circleH, rectRadius: 0.15,
        fill: { color: C.ALERT },
      }, { fontSize: 13, fontFace: FONT_B, color: C.WHITE, bold: true });
    }
  );

  // ================================================================
  // SLIDE 5: We Do — Label Scenarios
  // ================================================================
  contentSlide(
    pres,
    "We Do",
    C.SECONDARY,
    "Thinking About Labels",
    [
      "With your partner, discuss each scenario:",
      "What label is being used?",
      "Who is putting the label on?",
      "How might the person being labelled feel?",
    ],
    NOTES_LABELS_WEDO,
    FOOTER,
    (slide, layoutGuide) => {
      const rx = layoutGuide.rightX;
      const rw = layoutGuide.rightW;
      const topY = layoutGuide.panelTopPadded;

      // Scenario cards
      const scenarios = [
        {
          num: "1",
          text: "A student who has lived in Australia for three years is still called 'the foreign kid'",
          color: C.PRIMARY,
        },
        {
          num: "2",
          text: "A girl who loves maths is told by classmates: 'Girls are not good at maths'",
          color: C.SECONDARY,
        },
      ];

      scenarios.forEach((sc, i) => {
        const sy = topY + i * 1.6;
        addCard(slide, rx, sy, rw, 1.4, { fill: C.WHITE, strip: sc.color });
        slide.addText("Scenario " + sc.num, {
          x: rx + 0.15, y: sy + 0.08, w: rw - 0.3, h: 0.3,
          fontSize: 11, fontFace: FONT_B, color: sc.color, bold: true, margin: 0,
        });
        slide.addText(sc.text, {
          x: rx + 0.15, y: sy + 0.38, w: rw - 0.3, h: 0.9,
          fontSize: 12.5, fontFace: FONT_B, color: C.CHARCOAL, valign: "top", margin: 0,
        });
      });
    }
  );

  // ================================================================
  // SLIDE 6: I Do — Bystander vs Upstander
  // ================================================================
  contentSlide(
    pres,
    "I Do",
    C.PRIMARY,
    "Bystander vs Upstander",
    [
      "Bystander: sees something wrong and does nothing",
      "Upstander: sees something wrong and takes action",
      "Why do people sometimes stay silent?",
      "They want to belong  |  They are scared  |  They do not know what to do",
    ],
    NOTES_BYSTANDER,
    FOOTER,
    (slide, layoutGuide) => {
      const rx = layoutGuide.rightX;
      const rw = layoutGuide.rightW;
      const topY = layoutGuide.panelTopPadded;

      // Bystander box
      addTextOnShape(slide, "BYSTANDER\nStands BY", {
        x: rx + 0.1, y: topY + 0.1, w: rw - 0.2, h: 1.2, rectRadius: 0.1,
        fill: { color: C.ALERT },
      }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });

      // Upstander box
      addTextOnShape(slide, "UPSTANDER\nStands UP", {
        x: rx + 0.1, y: topY + 1.5, w: rw - 0.2, h: 1.2, rectRadius: 0.1,
        fill: { color: C.SUCCESS },
      }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // ================================================================
  // SLIDE 7: CFU — Bystander/Upstander Check (with reveal)
  // ================================================================
  withReveal(
    () => cfuSlide(
      pres,
      "Check",
      "Bystander or Upstander?",
      "Finger Voting",
      "During a footy game, the crowd starts booing one player every time he touches the ball.\n\nA person in the crowd feels uncomfortable but keeps clapping along with everyone else.\n\n1 finger = Bystander    2 fingers = Upstander",
      NOTES_CFU_BYSTANDER,
      FOOTER
    ),
    (slide) => {
      addTextOnShape(slide, "BYSTANDER -- Feeling uncomfortable is not the same as taking action", {
        x: 0.5, y: 4.25, w: 9, h: 0.55, rectRadius: 0.08,
        fill: { color: C.ALERT },
      }, { fontSize: 14, fontFace: FONT_B, color: C.WHITE, bold: true });
    }
  );

  // ================================================================
  // SLIDE 8: I Do — Media & Perception
  // ================================================================
  contentSlide(
    pres,
    "I Do",
    C.PRIMARY,
    "Media Shapes Perception",
    [
      "Media does not just report -- it shapes how we see people",
      "If media shows someone negatively over and over, people start believing it",
      "Not all sources are equally reliable",
      "We must always read and view critically",
    ],
    NOTES_MEDIA,
    FOOTER,
    (slide, layoutGuide) => {
      const rx = layoutGuide.rightX;
      const rw = layoutGuide.rightW;
      const topY = layoutGuide.panelTopPadded;

      addCard(slide, rx, topY, rw, 3.0, { fill: C.BG_CARD, strip: C.PRIMARY });
      slide.addText("Key Question", {
        x: rx + 0.15, y: topY + 0.1, w: rw - 0.3, h: 0.35,
        fontSize: 13, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0,
      });
      slide.addText("If a news channel shows the same person negatively every night, what starts to happen?", {
        x: rx + 0.2, y: topY + 0.5, w: rw - 0.4, h: 1.2,
        fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, italic: true, valign: "top", margin: 0,
      });

      // Reliability spectrum
      slide.addText("Reliability Spectrum", {
        x: rx + 0.15, y: topY + 1.8, w: rw - 0.3, h: 0.3,
        fontSize: 11, fontFace: FONT_B, color: C.MUTED, bold: true, margin: 0,
      });

      const spectrumItems = [
        { label: "Peer-reviewed research", color: C.SUCCESS },
        { label: "Quality journalism", color: C.ACCENT },
        { label: "Social media posts", color: C.ALERT },
        { label: "Click-bait headlines", color: C.ALERT },
      ];
      spectrumItems.forEach((item, i) => {
        const iy = topY + 2.15 + i * 0.28;
        slide.addShape("roundRect", {
          x: rx + 0.2, y: iy, w: 0.15, h: 0.15, rectRadius: 0.075,
          fill: { color: item.color },
        });
        slide.addText(item.label, {
          x: rx + 0.45, y: iy - 0.03, w: rw - 0.65, h: 0.22,
          fontSize: 10, fontFace: FONT_B, color: C.CHARCOAL, valign: "middle", margin: 0,
        });
      });
    }
  );

  // ================================================================
  // SLIDE 9: We Do — Media Reliability Discussion
  // ================================================================
  contentSlide(
    pres,
    "We Do",
    C.SECONDARY,
    "Reliable or Unreliable?",
    [
      "With your partner, discuss each media type:",
      "Is this a reliable source? Why or why not?",
      "Remember: no source is 100% reliable",
      "The question is how much checking and evidence it uses",
    ],
    NOTES_MEDIA_WEDO,
    FOOTER,
    (slide, layoutGuide) => {
      const rx = layoutGuide.rightX;
      const rw = layoutGuide.rightW;
      const topY = layoutGuide.panelTopPadded;

      addCard(slide, rx, topY, rw, 3.2, { fill: C.BG_CARD, strip: C.SECONDARY });
      slide.addText("Discuss These Sources", {
        x: rx + 0.15, y: topY + 0.1, w: rw - 0.3, h: 0.35,
        fontSize: 13, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0,
      });

      const sources = [
        "ABC News article",
        "A friend's Instagram story",
        "Wikipedia",
        "A YouTube vlog",
        "A government health website",
        "A blog post by an unknown author",
        "A newspaper editorial",
      ];
      sources.forEach((src, i) => {
        slide.addText(src, {
          x: rx + 0.25, y: topY + 0.55 + i * 0.37, w: rw - 0.5, h: 0.32,
          fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, valign: "middle", margin: 0,
          bullet: true,
        });
      });
    }
  );

  // ================================================================
  // SLIDE 10: You Do — Documentary Preparation Reflection
  // ================================================================
  contentSlide(
    pres,
    "You Do",
    C.SUCCESS,
    "Preparing for the Documentary",
    [
      "Next session we will watch a documentary about Adam Goodes -- an Indigenous Australian AFL champion",
      "In your inquiry book or on the handout:",
      "First: Write one example of a label that does not match who the person really is",
      "Next: Write what an upstander could do if they saw someone being labelled unfairly",
      "Then: Write one question you want to keep in mind while watching",
    ],
    NOTES_YOUDO,
    FOOTER,
    (slide, layoutGuide) => {
      const rx = layoutGuide.rightX;
      const rw = layoutGuide.rightW;
      const topY = layoutGuide.panelTopPadded;

      addCard(slide, rx, topY, rw, 3.0, { fill: C.BG_CARD, strip: C.SUCCESS });
      slide.addText("Your Viewing Lens", {
        x: rx + 0.15, y: topY + 0.1, w: rw - 0.3, h: 0.35,
        fontSize: 13, fontFace: FONT_H, color: C.SUCCESS, bold: true, margin: 0,
      });

      const lenses = [
        { icon: "Labels", desc: "Who is being labelled? By whom?" },
        { icon: "Bystanders", desc: "Who stays silent? Why?" },
        { icon: "Upstanders", desc: "Who takes action? How?" },
        { icon: "Media", desc: "How does media shape the story?" },
      ];
      lenses.forEach((lens, i) => {
        const ly = topY + 0.55 + i * 0.6;
        addTextOnShape(slide, lens.icon, {
          x: rx + 0.15, y: ly, w: 1.3, h: 0.45, rectRadius: 0.06,
          fill: { color: C.PRIMARY },
        }, { fontSize: 10, fontFace: FONT_B, color: C.WHITE, bold: true });
        slide.addText(lens.desc, {
          x: rx + 1.55, y: ly, w: rw - 1.7, h: 0.45,
          fontSize: 11.5, fontFace: FONT_B, color: C.CHARCOAL, valign: "middle", margin: 0,
        });
      });
    }
  );

  // ================================================================
  // SLIDE 11: Closing
  // ================================================================
  closingSlide(
    pres,
    "Turn and Talk: Which of today's ideas -- labels, bystander/upstander, or media influence -- do you think will be most important when watching the documentary? Why?",
    [
      "I can explain how labels and perceptions can shape identity",
      "I can describe the difference between a bystander and an upstander",
      "I can think critically about how media influences perception",
    ],
    NOTES_CLOSING
  );

  // ================================================================
  // SLIDE 12: Resources
  // ================================================================
  addResourceSlide(
    pres,
    RESOURCE_ITEMS,
    { C, FONT_H, FONT_B },
    FOOTER,
    NOTES_RESOURCES
  );

  // ── Write PPTX ──
  const pptxPath = path.join(LESSON_FOLDER, PPTX_NAME);
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to", pptxPath);

  // ═══════════════════════════════════════════════════════════════
  // Companion PDFs
  // ═══════════════════════════════════════════════════════════════

  // ── PDF 1: Labels and Identity Reflection ──
  await generateReflectionHandout();

  // ── PDF 2: Documentary Viewing Guide ──
  await generateViewingGuide();

  console.log("All resources written to:", RES_DIR);
}

// ═══════════════════════════════════════════════════════════════
// PDF Generation
// ═══════════════════════════════════════════════════════════════

async function generateReflectionHandout() {
  const doc = createPdf({ title: "Session 1 Labels and Identity Reflection" });

  let y = addPdfHeader(doc, "Labels and Identity Reflection", {
    subtitle: "Identity, Labels & Standing Up",
    color: hex(C.PRIMARY),
    lessonInfo: "Grade 5/6 Inquiry | Preparing for the Adam Goodes Documentary",
  });

  // Name/Date
  y = addBodyText(doc, "Name: ___________________________    Date: _______________", y);
  y += 5;

  // Section 1: Labels
  y = addSectionHeading(doc, "Part 1: Labels and Identity", y, { color: hex(C.PRIMARY) });

  y = addBodyText(doc, "Write one example of a label that does not match who the person really is. This can be about someone you know, a character from a story, or a famous person.", y);
  y = addLinedArea(doc, y, 4);

  y = addBodyText(doc, "How might this label affect the person? How could it change how they see themselves?", y);
  y = addLinedArea(doc, y, 3);

  // Section 2: Bystander / Upstander
  y = addSectionHeading(doc, "Part 2: Bystander vs Upstander", y, { color: hex(C.SECONDARY) });

  y = addBodyText(doc, "What could an upstander do if they saw someone being labelled unfairly? Write at least two specific actions.", y);
  y = addLinedArea(doc, y, 4);

  y = addBodyText(doc, "Why is it sometimes hard to be an upstander? What makes people stay silent?", y);
  y = addLinedArea(doc, y, 3);

  // Section 3: Documentary preparation
  y = addSectionHeading(doc, "Part 3: Preparing to Watch", y, { color: hex(C.ACCENT) });

  y = addBodyText(doc, "Write one question you want to keep in mind while watching the Adam Goodes documentary.", y);
  y = addLinedArea(doc, y, 3);

  y = addTipBox(doc, "As you watch the documentary, look for: Labels people use, who acts as a bystander or upstander, and how media shapes the story.", y, { color: hex(C.SUCCESS) });

  addPdfFooter(doc, "Identity, Labels & Standing Up | Grade 5/6 Inquiry");

  const filePath = path.join(LESSON_FOLDER, REFLECTION_RESOURCE.fileName);
  await writePdf(doc, filePath);
  console.log("PDF written:", filePath);
}

async function generateViewingGuide() {
  const doc = createPdf({ title: "Session 1 Documentary Viewing Guide" });

  let y = addPdfHeader(doc, "Documentary Viewing Guide", {
    subtitle: "Adam Goodes Documentary",
    color: hex(C.PRIMARY),
    lessonInfo: "Grade 5/6 Inquiry | Identity, Labels & Standing Up",
  });

  // Name/Date
  y = addBodyText(doc, "Name: ___________________________    Date: _______________", y);
  y += 5;

  y = addTipBox(doc, "Use this guide while watching the documentary. Pause to write notes when you notice something important. You do not need to fill every box -- focus on what stands out to you.", y, { color: hex(C.PRIMARY) });

  // Section 1: Labels
  y = addSectionHeading(doc, "Labels and Identity", y, { color: hex(C.PRIMARY) });

  y = addBodyText(doc, "What labels or words are used to describe Adam Goodes? Who uses them?", y);
  y = addLinedArea(doc, y, 4);

  y = addBodyText(doc, "How do these labels affect him? Does he agree with them?", y);
  y = addLinedArea(doc, y, 3);

  // Section 2: Bystanders and Upstanders
  y = addSectionHeading(doc, "Bystanders and Upstanders", y, { color: hex(C.SECONDARY) });

  y = addBodyText(doc, "Who acts as a bystander in the documentary? What do they do (or not do)?", y);
  y = addLinedArea(doc, y, 3);

  y = addBodyText(doc, "Who acts as an upstander? What actions do they take?", y);
  y = addLinedArea(doc, y, 3);

  // Page break for clean layout
  doc.addPage();
  y = 50;

  // Section 3: Media
  y = addSectionHeading(doc, "Media Influence", y, { color: hex(C.ACCENT) });

  y = addBodyText(doc, "How does the media present Adam Goodes? Is the coverage fair and balanced?", y);
  y = addLinedArea(doc, y, 4);

  y = addBodyText(doc, "How might media coverage have influenced how the public saw him?", y);
  y = addLinedArea(doc, y, 3);

  // Section 4: Reflection
  y = addSectionHeading(doc, "Your Reflection", y, { color: hex(C.SUCCESS) });

  y = addBodyText(doc, "What moment in the documentary stood out to you the most? Why?", y);
  y = addLinedArea(doc, y, 4);

  y = addBodyText(doc, "How does Adam Goodes' experience connect to what we learned about labels, bystanders/upstanders, and media?", y);
  y = addLinedArea(doc, y, 4);

  y = addBodyText(doc, "What is one thing you will take away from this documentary?", y);
  y = addLinedArea(doc, y, 3);

  addPdfFooter(doc, "Identity, Labels & Standing Up | Grade 5/6 Inquiry");

  const filePath = path.join(LESSON_FOLDER, VIEWING_GUIDE_RESOURCE.fileName);
  await writePdf(doc, filePath);
  console.log("PDF written:", filePath);
}

// ── Run ──
build().catch((err) => {
  console.error(err);
  process.exit(1);
});
