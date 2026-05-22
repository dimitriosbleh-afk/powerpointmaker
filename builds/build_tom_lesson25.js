"use strict";

// Tom Unit -- Lesson 25 (FINAL): Plan & Write the Conclusion + Publish & Celebrate
// Week 6, Lesson 25, Grade 5/6 Literacy
// Writing: Plan + write the concluding paragraph of the information report using the T-S-G structure
// Final session of the entire unit -- includes a publish, share and celebration block

const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");

const { createTheme, weekToVariant } = require("../themes/factory");
// Unit cohesion: all Tom lessons use the same variant.
const T = createTheme("literacy", "grade56", weekToVariant(2));
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  addTopBar, addBadge, addTitle, addCard, addFooter,
  withReveal,
  titleSlide, liSlide, contentSlide,
  cfuSlide, closingSlide,
  modellingSlide,
} = T;

const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addLinedArea,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const SESSION_NUMBER = 25;
const FOOTER = "Conclusion + Publish & Celebrate | Lesson 25 | Week 6 | Year 5/6 Literacy";
const OUT_DIR = "output/Tom_Lesson25_Conclusion_Publish_Celebrate";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const TEMPLATE_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Conclusion Plan and Writing Template",
  "Student template: plan the T-S-G conclusion and write it on the same page."
);
const MENTOR_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Mentor Conclusion Annotated",
  "Annotated model conclusion paragraph showing T, S and G sentences and language features."
);
const PUBLISH_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Publish and Share Checklist",
  "Student checklist for the publish, present and celebrate stage of the information report."
);
const RESOURCE_ITEMS = [TEMPLATE_RESOURCE, MENTOR_RESOURCE, PUBLISH_RESOURCE];
const TEMPLATE_PDF_PATH = path.join(OUT_DIR, TEMPLATE_RESOURCE.fileName);
const MENTOR_PDF_PATH = path.join(OUT_DIR, MENTOR_RESOURCE.fileName);
const PUBLISH_PDF_PATH = path.join(OUT_DIR, PUBLISH_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Lesson 25. The FINAL session of our Tom unit
- We finish our information report by writing the conclusion today
- Then we publish, share and celebrate the work you have done across the whole unit

DO:
- Display title slide as students settle
- Have student information report drafts ready (introduction + 4 body paragraphs)
- Have template, mentor and publish checklist printed (do not distribute yet)
- Optional: have illustration paper, maps, devices ready for the publish block

TEACHER NOTES:
This is the final lesson of the unit. It has two distinct halves: write the conclusion (focused, time-protected), then publish and celebrate (open, joyful). Plan to spend about 30 minutes on the conclusion writing and about 15 to 20 minutes on the publish and share block. Make the celebration genuine.

WATCH FOR:
- Students who think the unit is over and switch off -- this lesson still has real writing in it
- Students who are anxious about sharing -- give a partner option, not a class option, for those students

[Literacy: Title | VTLM 2.0: Establishing Purpose and Relevance]`;

const NOTES_RESOURCES = `SAY:
- Three resources today
- The ${TEMPLATE_RESOURCE.name} is your conclusion plan and writing page
- The ${MENTOR_RESOURCE.name} is the model -- annotated
- The ${PUBLISH_RESOURCE.name} is your checklist for sharing your finished report

DO:
- Print the template (one per student)
- Print the mentor (one per student or per pair)
- Print the publish checklist (one per student)
- Have student drafts of intro + 4 body paragraphs ready on desks

TEACHER NOTES:
The publish checklist invites choice -- students can choose how to share. Have illustration paper, maps from the article, and (if available) one or two devices for those choosing a digital format.

[Literacy: Resources | VTLM 2.0: Student Resources]`;

const NOTES_HOOK = `SAY:
- Before we write, a quick connection back to what we have built across the unit
- Mini-whiteboards out
- 30 seconds: list the topics of our four body paragraphs
- We will reveal together

DO:
- 30 seconds whiteboard work
- Boards up
- Cold call 1-2 students for each topic
- Confirm:
  - Body paragraph 1: Life in 18th century England (the world Tom came from)
  - Body paragraph 2: The First Fleet (Tom's journey)
  - Body paragraph 3: Arrival in Sydney Cove and early settlement
  - Body paragraph 4: The impact of British settlement on First Nations Australians

TEACHER NOTES:
This launch activates the whole information report as one piece of work. Students need to see the full scope before they conclude it. If your unit had slightly different paragraph topics, adjust the names accordingly on this slide.

WATCH FOR:
- Students who name all four -- excellent unit-wide recall; celebrate
- Students who miss one -- supply gently before the write

[Literacy: Hook | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention
- We are using a new structure today for the conclusion: T - S - G
- Read the success criteria

DO:
- Choral read the LI, then the SCs
- Brief: "The conclusion is the LAST paragraph of your report. It wraps everything together"

TEACHER NOTES:
SC1 targets the T-S-G structure. SC2 targets the language features (must keep matching the body paragraphs -- past tense, third person, key terms). SC3 targets editing for cohesion across the WHOLE report, not just this paragraph.

WATCH FOR:
- Students who think "I just need to summarise" -- redirect: "The conclusion is more than a summary. It addresses the topic and closes the report"

[Literacy: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_REVISE_TSG = `SAY:
- The conclusion has THREE parts. Look at the slide
- T: Thesis statement -- makes clear what your body paragraphs discussed. Big picture
- S: Specific statement -- addresses the topic directly using key terms (Sydney Cove, Eora, convicts, the First Fleet)
- G: General statement -- summarises the topic and CLOSES the report. The last sentence
- T, then S, then G. In that order

DO:
- Display the T-S-G card
- Choral read each row
- Quick oral check: "Which one comes first? Which one closes the report?"

TEACHER NOTES:
Use the simple meanings on the slide. 'Thesis statement', 'specific statement' and 'general statement' are the formal terms from the school planner. The slide pairs each formal term with a simple meaning so students can see both.

WATCH FOR:
- Students who try to write the G first -- redirect: "Start with T. T is the big picture. G is the closer"
- Students who confuse T and S -- T is the big picture, S is the specific facts

[Literacy: Revise | VTLM 2.0: Explicit Teaching / New Structure]`;

const NOTES_IDO = `SAY:
- Watch me think through a conclusion using T - S - G
- T (big picture): "This report has explored what life was like for both the British settlers and the First Nations peoples of Australia during the first years of the colony at Sydney Cove."
- S (specific): "From the convict gaols of London, through the long First Fleet voyage, to the early years of settlement and the devastating impact on the Eora and other First Nations peoples, the events of 1788 changed many lives."
- G (closer): "Although the early colony brought hardship for the British settlers and profound change for First Nations peoples, the legacy of those years is still part of Australia today."
- Three sentences. T, then S, then G

DO:
- Display the I Do slide
- Talk through each sentence
- Highlight key terms in the S sentence (Sydney Cove, First Fleet, Eora)
- Highlight the closer feeling in the G sentence
- Note the respectful framing throughout

TEACHER NOTES:
This is the most important model in the lesson. Read it twice if needed. Students should leave this slide with a clear sense of what each of T, S, G sounds like in practice.

WATCH FOR:
- Students who try to memorise the model -- redirect: "Your topic is the same. Your sentences will use YOUR notes and YOUR words"
- Students who notice the respectful framing -- celebrate

[Literacy: I Do | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_CFU_BUILD = `SAY:
- Quick check. Two possible General statements to CLOSE the report. Which one closes better?
- A: "And that is the end of the report."
- B: "Although the early colony brought hardship for the British settlers and profound change for First Nations peoples, the legacy of those years is still part of Australia today."
- Hold up A or B. Three, two, one -- show

DO:
- Display both
- Show Me Boards
- Scan: most students should choose B
- Cold Call 1-2 students: "Why B?"

TEACHER NOTES:
A is a flat ending -- it tells the reader the report is over but does not give them anything to take away. B closes by acknowledging both groups, naming the legacy, and connecting to today. B is the stronger closer.

WATCH FOR:
- Students who pick A because it is simpler -- redirect: "A just tells me the report has stopped. Does B leave you with a bigger idea?"
- Students who pick B and articulate WHY -- ready to write

[Literacy: CFU | VTLM 2.0: Formative Assessment]`;

const NOTES_CFU_REVEAL = `SAY:
- Stronger closer: B
- B names both groups, names the legacy, and connects to today
- A just signals the report has stopped
- Aim for B-style depth in your G sentence

DO:
- Display the reveal banner
- Read B aloud
- Pivot if many picked A: "What does B give the reader that A does not?"

CFU CHECKPOINT:
Technique: Show Me Boards (A or B)
Script:
- "Hold up A or B. Which closes the report better?"
- Scan for: most students choose B (>=80%)
PROCEED (>=80%): Most chose B. Release to plan and write.
PIVOT (<80%): Most likely issue -- students think a conclusion just stops. Reteach: "The G sentence is the LAST thing the reader sees. Give them something to take away."

TEACHER NOTES:
After reveal, release students to plan and write. Time is now the limiting factor.

WATCH FOR:
- Students who self-correct toward a stronger G -- celebrate

[Literacy: CFU Reveal | VTLM 2.0: Formative Assessment]`;

const NOTES_YOUDO = `SAY:
- Your turn. Use the Conclusion template
- First: plan your T, S and G on page 1
- Next: write your conclusion paragraph on page 2 -- 3 sentences, in T, S, G order
- Then: read it back aloud quietly. Does it CLOSE the whole report?
- 18 minutes. I will circulate

DO:
- Distribute the template, mentor and publish checklist
- Place student drafts on each desk
- Circulate -- prioritise students who looked unsure during the CFU
- Quick conferences: "Read me your T sentence. Does it set up the conclusion?"

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use a sentence-frame version of the mentor on the back of the mentor sheet. Students fill in the blanks using their notes and the body paragraph topics
- Suggested frame: "This report has explored ____. From ____, through ____, to ____, the events of 1788 ____. Although the early colony brought ____, the legacy of those years is ____."
- Extra Notes: These students still hit all three SCs but with reduced cognitive load on sentence construction

EXTENDING PROMPT:
- Task: After writing your three sentences, add a fourth optional sentence that names what this teaches us today
- Extra Notes: Push for connection -- "Today we acknowledge..." or "These stories remind us that..."

TEACHER NOTES:
The 18-minute writing block protects the focused work. Active circulation is the formative assessment. Most students should complete the three sentences inside the block.

WATCH FOR:
- Students who write the G first -- redirect: "Start with T. Then S. Then G"
- Students who only write a summary (no big-picture T) -- prompt: "What does this report explore? That is your T"
- Students who use 'we' or 'you' -- redirect to third person

[Literacy: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EDIT = `SAY:
- Time to edit your WHOLE report
- Read your introduction, then your four body paragraphs, then your conclusion -- aloud and quietly
- Use the publish checklist for the language features
- Look for: does each paragraph link to the next? Are key terms consistent? Are there any small fixes?
- Then make any small changes neatly

DO:
- 6 minutes silent reading and editing
- Circulate -- prompt students who finish quickly to swap with a partner
- Encourage editing in pen or in a different colour so the changes are visible

TEACHER NOTES:
This edit pass is the cohesion edit -- not just the conclusion. Students need to see the WHOLE report as one piece. If students cannot complete this in the time available, mark it for next available pocket of time.

WATCH FOR:
- Students who tick the checklist without checking -- prompt: "Show me where in your report you used an appositive"
- Students who do not change anything -- prompt: "Read it aloud. Pick one sentence that does not flow"

[Literacy: Edit | VTLM 2.0: Reflection and Refinement]`;

const NOTES_PUBLISH = `SAY:
- Time to publish and share
- You have CHOICE. Look at the publish checklist
- Option 1: add features -- illustrations, a map of Sydney Cove, a timeline, a diagram of the First Fleet ship
- Option 2: present your report -- read your favourite paragraph aloud to your partner or your table
- Option 3: digital version -- if a device is available, type your report or add it to a slide
- 15 minutes. Make it yours

DO:
- Display the publish slide
- Walk students through the three options
- Distribute illustration paper or open devices as needed
- Circulate -- celebrate specific work; take photos for the class wall if appropriate

TEACHER NOTES:
This is the celebration block. Make it generous. Students have written FIVE paragraphs (intro + 4 body) plus a conclusion across many sessions. Honour the achievement. Photograph good work for the class wall.

WATCH FOR:
- Students who finish quickly -- offer the next option
- Students who are nervous about presenting -- give the partner option, never force a whole-class reading

[Literacy: Publish | VTLM 2.0: Application and Celebration]`;

const NOTES_CLOSING = `SAY:
- This is the LAST slide of the LAST lesson in our unit
- Quick self-check against the success criteria for this lesson
- Then a unit reflection: turn to your partner. What are you most proud of from the WHOLE unit? It can be a sentence, a paragraph, a moment from the novel, a vocabulary word you learned, or anything else
- 90 seconds

DO:
- Run thumbs / fingers check for each SC
- Listen in on partner shares -- this is the moment to honour the unit
- Briefly: "Well done. You have read a whole novel, learned new sentence-level skills, and written a five-paragraph information report. That is real work"
- Collect finished reports for display, marking or filing

TEACHER NOTES:
End on the reflection -- not on a teacher speech. Students choose what to claim as their own success. Collect the reports for whatever next step the school has planned (display, marking, portfolio).

WATCH FOR:
- Students who name a specific paragraph they are proud of -- celebrate openly
- Students who name a moment from the novel -- evidence of full engagement
- Students who feel uncertain -- offer a private check-in time

[Literacy: Closing | VTLM 2.0: Review and Reflect / End of Unit]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Tom Appleby Conclusion + Publish -- Lesson 25 (Final)";

  // SLIDE 1 -- Title
  titleSlide(
    pres,
    "Conclusion + Publish & Celebrate",
    "The Final Session of Our Unit",
    "Lesson 25  |  Week 6  |  Year 5/6 Literacy",
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

  // SLIDE 3 -- Hook / Launch
  contentSlide(
    pres,
    "Launch",
    C.PRIMARY,
    "Our Four Body Paragraphs So Far",
    [
      "Mini-whiteboards out",
      "30 seconds: list the topics of our FOUR body paragraphs",
      "Hint: each one covered a different part of the bigger story",
      "Boards up when called",
    ],
    NOTES_HOOK,
    FOOTER
  );

  // SLIDE 4 -- LI / SC
  liSlide(
    pres,
    [
      "We are learning to plan and write a conclusion paragraph for our information report using the T-S-G structure, and to publish and share our finished work",
    ],
    [
      "I can write a conclusion paragraph with a T sentence, an S sentence and a G sentence",
      "I can use language features of an information report including key terms from my body paragraphs",
      "I can edit my full report and share my finished work with my class",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // SLIDE 5 -- Revise: T-S-G structure
  contentSlide(
    pres,
    "Revise",
    C.SECONDARY,
    "The Conclusion -- T then S then G",
    [
      "T: Thesis statement -- what your body paragraphs discussed (big picture)",
      "S: Specific statement -- addresses the topic directly using key terms",
      "G: General statement -- summarises and CLOSES the report (last sentence)",
      "Order: T, then S, then G. Three sentences.",
      "Past tense, third person, respectful language -- same as the body paragraphs",
    ],
    NOTES_REVISE_TSG,
    FOOTER
  );

  // SLIDE 6 -- I Do: model conclusion
  modellingSlide(
    pres,
    "I Do -- Watch Me",
    "Model Conclusion: T - S - G",
    "T (big picture):\n\n\"This report has explored what life was like for both the British settlers and the First Nations peoples of Australia during the first years of the colony at Sydney Cove.\"\n\nNotice:\n- big picture\n- names what the body paragraphs discussed\n- third person",
    "S (specific):\n\n\"From the convict gaols of London, through the long First Fleet voyage, to the early years of settlement and the devastating impact on the Eora and other First Nations peoples, the events of 1788 changed many lives.\"\n\nG (closer):\n\n\"Although the early colony brought hardship for the British settlers and profound change for First Nations peoples, the legacy of those years is still part of Australia today.\"",
    NOTES_IDO,
    FOOTER
  );

  // SLIDE 7 + 8 -- CFU: which G sentence closes better? (reveal pair)
  function buildCfuBase() {
    const slide = pres.addSlide();
    addTopBar(slide, C.ALERT);
    addBadge(slide, "CFU", { color: C.ALERT });
    addTitle(slide, "Which G Sentence Closes the Report Better?", { color: C.ALERT });

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
      x: 0.5, y: CONTENT_TOP, w: 2.8, h: 0.40, rectRadius: 0.08,
      fill: { color: C.ALERT },
    });
    slide.addText("Show Me Boards: A or B", {
      x: 0.5, y: CONTENT_TOP, w: 2.8, h: 0.40,
      fontSize: 12, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    const cardY = CONTENT_TOP + 0.55;
    const cardH = 1.00;
    addCard(slide, 0.5, cardY, 9, cardH, { strip: C.MUTED, fill: C.WHITE });
    slide.addText("A", {
      x: 0.7, y: cardY + 0.10, w: 0.5, h: 0.4,
      fontSize: 22, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0,
    });
    slide.addText("\"And that is the end of the report.\"", {
      x: 1.2, y: cardY + 0.18, w: 8.0, h: cardH - 0.30,
      fontSize: 18, fontFace: FONT_H, color: C.CHARCOAL, italic: true, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    const cardBY = cardY + cardH + 0.18;
    const cardBH = SAFE_BOTTOM - cardBY - 0.55;
    addCard(slide, 0.5, cardBY, 9, cardBH, { strip: C.SECONDARY, fill: C.BG_CARD });
    slide.addText("B", {
      x: 0.7, y: cardBY + 0.10, w: 0.5, h: 0.4,
      fontSize: 22, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0,
    });
    slide.addText("\"Although the early colony brought hardship for the British settlers and profound change for First Nations peoples, the legacy of those years is still part of Australia today.\"", {
      x: 1.2, y: cardBY + 0.14, w: 8.0, h: cardBH - 0.24,
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
      slide.addText("Stronger closer: B  --  names both groups, names the legacy, connects to today", {
        x: 0.5, y: revealY, w: 9, h: 0.42,
        fontSize: 14, fontFace: FONT_B, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      slide.addNotes(NOTES_CFU_REVEAL);
    }
  );

  // SLIDE 9 -- You Do: Plan + Write conclusion
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "You Do", { color: C.PRIMARY, w: 1.5 });
    addTitle(s, "Plan & Write Your Conclusion");

    addCard(s, 0.5, CONTENT_TOP, 9, 1.95, { strip: C.PRIMARY, fill: C.WHITE });
    s.addText("Three sentences:  T  ->  S  ->  G", {
      x: 0.75, y: CONTENT_TOP + 0.10, w: 8.4, h: 0.30,
      fontSize: 14, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("First:    Plan your T -- big picture sentence about what the report explored\nNext:     Plan your S -- specific sentence with key terms (Sydney Cove, First Fleet, Eora)\nThen:     Plan your G -- closer sentence that connects to today\nFinally:  Write the three sentences in T, S, G order on page 2", {
      x: 0.75, y: CONTENT_TOP + 0.46, w: 8.4, h: 1.40,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    const tipY = CONTENT_TOP + 2.10;
    addCard(s, 0.5, tipY, 9, SAFE_BOTTOM - tipY, { strip: C.ACCENT, fill: C.BG_CARD });
    s.addText("Writing Time: 18 minutes", {
      x: 0.75, y: tipY + 0.10, w: 5, h: 0.30,
      fontSize: 13, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("- Use third person -- no I, we or you\n- The mentor paragraph is on your desk if you need a model\n- Read it back aloud quietly -- does it CLOSE the whole report?", {
      x: 0.75, y: tipY + 0.42, w: 8.4, h: 0.85,
      fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  }

  // SLIDE 10 -- Edit the whole report
  contentSlide(
    pres,
    "Edit",
    C.ACCENT,
    "Edit Your Whole Report",
    [
      "Read your intro, body 1, body 2, body 3, body 4, then conclusion -- aloud and quietly",
      "Use the publish checklist for the language features",
      "Check: does each paragraph link to the next? Are key terms consistent?",
      "Make any small changes neatly in pen or a different colour",
      "6 minutes",
    ],
    NOTES_EDIT,
    FOOTER
  );

  // SLIDE 11 -- Publish & Share
  contentSlide(
    pres,
    "Publish",
    C.SECONDARY,
    "Publish, Share & Celebrate",
    [
      "Option 1:  Add features -- illustrations, a map of Sydney Cove, a timeline, a diagram",
      "Option 2:  Present -- read your favourite paragraph aloud to your partner or your table",
      "Option 3:  Digital -- if a device is available, type your report or add it to a slide",
      "Use the publish checklist on your desk",
      "15 minutes -- make it yours",
    ],
    NOTES_PUBLISH,
    FOOTER
  );

  // SLIDE 12 -- Closing (unit-wide reflection)
  closingSlide(
    pres,
    {
      reflectionPrompt: "Turn to your partner. What are you MOST proud of from the whole unit? A sentence, a paragraph, a moment from the novel, a word -- anything.",
      scItems: [
        "I can write a conclusion paragraph with a T sentence, an S sentence and a G sentence",
        "I can use language features of an information report including key terms from my body paragraphs",
        "I can edit my full report and share my finished work with my class",
      ],
      selfAssessment: {
        prompt: "Show on your fingers: 1 (need help) to 5 (got it)",
        options: ["1-2", "3", "4-5"],
      },
    },
    NOTES_CLOSING
  );

  // ---- PDF: Conclusion Plan + Writing Template ----------------------------
  const tp = createPdf({ title: TEMPLATE_RESOURCE.name });
  let tpY = addPdfHeader(tp, "Conclusion Plan -- T, S, G", {
    color: C.PRIMARY,
    subtitle: "Plan your three sentences before you write",
    lessonInfo: "Lesson 25 | Week 6 | Year 5/6 Literacy",
    showNameDate: true,
  });

  tpY = addTipBox(tp, "T (Thesis) = big picture sentence about what the report explored. S (Specific) = sentence with key terms from your body paragraphs. G (General) = closing sentence that connects to today.", tpY, { color: C.PRIMARY });

  tpY = addSectionHeading(tp, "T -- Thesis Statement (big picture)", tpY, { color: C.PRIMARY });
  tpY = addBodyText(tp, "Hint: what did your body paragraphs discuss? Think across all four (England, First Fleet, Sydney Cove, First Nations).", tpY, { fontSize: 10, italic: true });
  tpY = addLinedArea(tp, tpY, 3, { lineSpacing: 22 });
  tpY += 4;

  tpY = addSectionHeading(tp, "S -- Specific Statement (key terms)", tpY, { color: C.PRIMARY });
  tpY = addBodyText(tp, "Hint: include key terms -- Sydney Cove, the First Fleet, the Eora, convicts, 1788.", tpY, { fontSize: 10, italic: true });
  tpY = addLinedArea(tp, tpY, 3, { lineSpacing: 22 });
  tpY += 4;

  tpY = addSectionHeading(tp, "G -- General Statement (closer)", tpY, { color: C.PRIMARY });
  tpY = addBodyText(tp, "Hint: leave the reader with a bigger idea. What does this teach us? What is the legacy today?", tpY, { fontSize: 10, italic: true });
  tpY = addLinedArea(tp, tpY, 3, { lineSpacing: 22 });

  addPdfFooter(tp, "Lesson 25 | Conclusion Plan -- Page 1");

  // Page 2 - write here
  tp.addPage();
  let tpY2 = addPdfHeader(tp, "Conclusion Paragraph -- Write Here", {
    color: C.PRIMARY,
    subtitle: "Three sentences in T, S, G order",
    lessonInfo: "Lesson 25 | Week 6 | Year 5/6 Literacy",
    showNameDate: false,
  });

  tpY2 = addTipBox(tp, "Write your three sentences as one paragraph -- T first, then S, then G. Use past tense and third person. Read it back aloud quietly when you finish.", tpY2, { color: C.SECONDARY });

  tpY2 = addLinedArea(tp, tpY2, 14, { lineSpacing: 22 });

  addPdfFooter(tp, "Lesson 25 | Conclusion Paragraph -- Page 2");

  // ---- PDF: Mentor Conclusion --------------------------------------------
  const mp = createPdf({ title: MENTOR_RESOURCE.name });
  let mpY = addPdfHeader(mp, "Mentor Conclusion -- Annotated", {
    color: C.PRIMARY,
    subtitle: "T - S - G structure with language features",
    lessonInfo: "Lesson 25 | Week 6 | Year 5/6 Literacy",
    showNameDate: false,
  });

  mpY = addTipBox(mp, "This is a model conclusion showing the T-S-G structure, language features and respectful framing. Use it as a reference -- do not copy it. Your conclusion will use YOUR notes and YOUR words.", mpY, { color: C.PRIMARY });

  mpY = addSectionHeading(mp, "Model Conclusion", mpY, { color: C.PRIMARY });
  mpY = addBodyText(mp, "This report has explored what life was like for both the British settlers and the First Nations peoples of Australia during the first years of the colony at Sydney Cove. From the convict gaols of London, through the long First Fleet voyage, to the early years of settlement and the devastating impact on the Eora and other First Nations peoples, the events of 1788 changed many lives. Although the early colony brought hardship for the British settlers and profound change for First Nations peoples, the legacy of those years is still part of Australia today.", mpY, { fontSize: 12 });
  mpY += 14;

  mpY = addSectionHeading(mp, "Annotations -- Structure", mpY, { color: C.PRIMARY });
  mpY = addBodyText(mp, "T (Thesis): \"This report has explored what life was like for both the British settlers and the First Nations peoples of Australia during the first years of the colony at Sydney Cove.\" -- big picture, names what was discussed.", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "S (Specific): \"From the convict gaols of London, through the long First Fleet voyage, to the early years of settlement and the devastating impact on the Eora and other First Nations peoples, the events of 1788 changed many lives.\" -- includes key terms from each body paragraph.", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "G (General): \"Although the early colony brought hardship for the British settlers and profound change for First Nations peoples, the legacy of those years is still part of Australia today.\" -- closes by acknowledging both groups and connecting to today.", mpY, { fontSize: 10 });
  mpY += 10;

  mpY = addSectionHeading(mp, "Annotations -- Language Features", mpY, { color: C.SECONDARY });
  mpY = addBodyText(mp, "Past tense: \"has explored\", \"changed\", \"brought\".", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "Third person: \"the British settlers\", \"the First Nations peoples\", \"the events of 1788\".", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "Adverbial of place / journey: \"From the convict gaols of London, through the long First Fleet voyage, to the early years of settlement...\".", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "Key terms threaded through: \"Sydney Cove\", \"First Fleet\", \"the Eora\", \"First Nations peoples\", \"1788\".", mpY, { fontSize: 10 });
  mpY += 10;

  mpY = addSectionHeading(mp, "Annotations -- Respectful Framing", mpY, { color: C.ALERT });
  mpY = addBodyText(mp, "Names both groups in the T sentence: \"both the British settlers and the First Nations peoples of Australia\".", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "Uses accurate impact language: \"devastating\", \"profound change\" -- avoids 'wiped out' or similar.", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "Closes with continuity: \"the legacy of those years is still part of Australia today\" -- First Nations peoples are still here.", mpY, { fontSize: 10 });

  addPdfFooter(mp, "Lesson 25 | Mentor Conclusion -- TEACHER AND STUDENT REFERENCE");

  // ---- PDF: Publish & Share Checklist ------------------------------------
  const pb = createPdf({ title: PUBLISH_RESOURCE.name });
  let pbY = addPdfHeader(pb, "Publish & Share -- Checklist", {
    color: C.PRIMARY,
    subtitle: "Final edit, publish and celebrate your finished report",
    lessonInfo: "Lesson 25 | Week 6 | Year 5/6 Literacy",
    showNameDate: true,
  });

  pbY = addTipBox(pb, "Use this checklist when you finish writing. Tick each item as you check it. Then choose how you want to publish and share your work.", pbY, { color: C.PRIMARY });

  pbY = addSectionHeading(pb, "Final Edit", pbY, { color: C.PRIMARY });
  pbY = addBodyText(pb, "__ My report has an introduction, four body paragraphs and a conclusion (six paragraphs in total)", pbY);
  pbY = addBodyText(pb, "__ I have used past tense throughout (arrived, lived, sailed)", pbY);
  pbY = addBodyText(pb, "__ I have used third person throughout (no I, we or you)", pbY);
  pbY = addBodyText(pb, "__ I have used key terms consistently (Sydney Cove, First Fleet, Eora, convicts)", pbY);
  pbY = addBodyText(pb, "__ I have used respectful language about First Nations peoples", pbY);
  pbY = addBodyText(pb, "__ I have at least one appositive or relative clause somewhere in my report", pbY);
  pbY = addBodyText(pb, "__ My conclusion has a T, an S and a G sentence in that order", pbY);
  pbY = addBodyText(pb, "__ My spelling and punctuation are checked", pbY);
  pbY += 8;

  pbY = addSectionHeading(pb, "Choose Your Publish Option", pbY, { color: C.SECONDARY });
  pbY = addBodyText(pb, "Option 1 -- Add features:", pbY, { fontSize: 11 });
  pbY = addBodyText(pb, "__ Draw an illustration to accompany your report", pbY);
  pbY = addBodyText(pb, "__ Add a map of Sydney Cove", pbY);
  pbY = addBodyText(pb, "__ Add a timeline of the events from 1788", pbY);
  pbY = addBodyText(pb, "__ Add a diagram (e.g. a First Fleet ship, a layout of the early colony)", pbY);
  pbY += 4;

  pbY = addBodyText(pb, "Option 2 -- Present:", pbY, { fontSize: 11 });
  pbY = addBodyText(pb, "__ Read your favourite paragraph aloud to your partner", pbY);
  pbY = addBodyText(pb, "__ Read your favourite paragraph aloud to your table", pbY);
  pbY = addBodyText(pb, "__ (Optional) Present your introduction to the whole class", pbY);
  pbY += 4;

  pbY = addBodyText(pb, "Option 3 -- Digital:", pbY, { fontSize: 11 });
  pbY = addBodyText(pb, "__ Type your report on a device, if one is available", pbY);
  pbY = addBodyText(pb, "__ Create a slide for one body paragraph and include images", pbY);
  pbY += 8;

  pbY = addSectionHeading(pb, "Unit Reflection", pbY, { color: C.ACCENT });
  pbY = addBodyText(pb, "What I am most proud of from this unit:", pbY, { fontSize: 10, italic: true });
  pbY = addLinedArea(pb, pbY, 4, { lineSpacing: 22 });
  pbY = addBodyText(pb, "Something I learned that surprised me:", pbY, { fontSize: 10, italic: true });
  pbY = addLinedArea(pb, pbY, 4, { lineSpacing: 22 });

  addPdfFooter(pb, "Lesson 25 | Publish & Share Checklist");

  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Tom_Lesson25.pptx` }),
    writePdf(tp, TEMPLATE_PDF_PATH),
    writePdf(mp, MENTOR_PDF_PATH),
    writePdf(pb, PUBLISH_PDF_PATH),
  ]);

  console.log("PPTX written to " + `${OUT_DIR}/Tom_Lesson25.pptx`);
  console.log("Done: " + TEMPLATE_RESOURCE.name);
  console.log("Done: " + MENTOR_RESOURCE.name);
  console.log("Done: " + PUBLISH_RESOURCE.name);
}

build().catch(console.error);
