"use strict";

// Information Reports - Term 3 Week 10, Session 4 (Year 5/6 Literacy)
// "Publish it so a reader can use it"
//
// Victorian Curriculum 2.0: English, Literacy, Levels 5-6 - VC2E5LY10 and
// VC2E6LY09, creating texts with multimodal elements as appropriate, and
// VC2E6LA07, the way still images are placed and captioned in a text.
//
// UNIT ANCHOR (locked): "Classify it. Describe it, one aspect at a time.
// Wrap it up. Facts all the way through." Publishing changes the look, never
// the structure - the anchor is what the layout has to make visible.
//
// Lesson shape: compare-two-models. Students judge finished pages against five
// rules, then publish their own and read each other's.
//
// The teacher's plan lists publishing and sharing as two sessions; Week 10 has
// four teaching days, so they are combined here. Stated on the overview slide
// in Session 1.
//
// Sources: teacher's unit plan - the three publishing formats are theirs.

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme } = require("../themes/factory");
const { composeGlanceNotes } = require("../themes/core/composeNotes");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
  PAGE, hex,
} = require("../themes/pdf_helpers");

const P = require("./inforep_lib");

const UNIT_VARIANT = 2;
const T = createTheme("literacy", "grade56", UNIT_VARIANT);
const {
  C, FONT_H, FONT_B, SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, closingSlide,
  annotatedModelSlide, compareVisualSlide,
  addFooter, addTextOnShape, addRevealAnswerBar,
  clickBuild, runSlideDiagnostics,
} = T;

const SESSION = 4;
const FOOTER = "Information Reports | Week 10 Session 4 | Year 5/6 Literacy";
const OUT_DIR = path.join(__dirname, "..", "output", "InfoReport_W10_S4");
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const GUIDE_RES = makeSessionResource(SESSION,
  "Publishing Guide",
  "The five layout rules, the three formats, and cut-out gallery walk slips.");
const RESOURCE_ITEMS = [GUIDE_RES];
fs.mkdirSync(RES_DIR, { recursive: true });

/* --- Teacher notes (Glance Format, megaprompt sections 45-47) ------------- */

const NOTES_TITLE =
  "Session 4, the last of the unit. Today the reports are published and read by an audience.";

const NOTES_RESOURCES =
  "Prep slide. Print the Publishing Guide, one per student, and cut the gallery walk slips before the lesson.\n" +
  "Devices, good paper or a slide template ready, whichever format your class is using. Clear a wall or desks for the gallery walk.\n" +
  "CATCH-UP: draft not finished? Publish what exists. One classified opening and one aspect paragraph, laid out properly, still meets today's intention.";

const NOTES_LAUNCH = composeGlanceNotes({
  answer: "open - listen for a reason tied to the topic, not to what is easiest.",
  beats: [
    [
      "SAY: Three ways to publish. Same facts, three different looks.",
      "Read the three with me.",
    ],
    [
      "ASK: Which one suits YOUR topic, and why?",
      "30 sec. Turn and tell, partner A first.",
      "EXPECT: a reason about the topic, not about the tech.",
    ],
    [
      "SCAN the room. Listen for because.",
      "80%+ -> cold call two different choices, then move on.",
      "Less -> name one topic yourself, choose aloud, re-ask.",
    ],
  ],
  trap: [
    "choosing the format that looks quickest.",
    "Fix: ask what the reader gets, student chooses again.",
  ],
  prep: "Low-coupling launch: three formats on screen, nothing earlier assumed. The three are the teacher's own. Under 5 minutes.",
  sources: P.SOURCE_LINE_UNIT,
  tag: "[Launch | Attention, focus and regulation | HITS 6]",
});

const NOTES_LI = composeGlanceNotes({
  beats: [
    [
      "POINT to the learning intention.",
      "SAY: Today we lay it out so a reader can actually use it.",
    ],
    "SAY: Read the I can statements with me. Everyone, together, on three.",
    "SAY: The first one everyone leaves with. A title and subheadings.",
  ],
  prep: "SC1 is reachable by every student; SC2 is the core target the exit ticket assesses; SC3 stretches into peer feedback. Tier labels stay off the slide.",
  tag: "[LI/SC | Planning made visible | HITS 1]",
});

const NOTES_IDO = composeGlanceNotes({
  answer: "title, subheadings, one font, visual beside its part.",
  beats: [
    [
      "SAY: A published page has five rules.",
      "Title at the top. Subheadings from your plan.",
    ],
    [
      "POINT to each rule card. The visual goes beside its paragraph.",
      "Caption directly underneath.",
    ],
    [
      "ASK: Which rule is broken most often?",
      "20 sec. Boards up on cue.",
      "EXPECT: the visual in the wrong place.",
    ],
    [
      "SCAN boards.",
      "80%+ -> cold call: why is that harder to read?",
      "Less -> name each rule on the page, re-ask.",
    ],
  ],
  trap: [
    "thinking publishing means decorating.",
    "Fix: ask which rule a border follows, student re-reads.",
  ],
  stretch: "Name a sixth rule a magazine follows.",
  help: "Name a rule, student finds it on the page.",
  prep: "Say the anchor once here. The layout has to make the structure visible: subheadings ARE the aspects.",
  sources: P.SOURCE_LINE_MODEL,
  tag: "[I Do | Explicit teaching | SC1 | HITS 3, 4]",
});

const NOTES_CFU = composeGlanceNotes({
  answer: "B. Beside the appearance paragraph, caption underneath.",
  beats: [
    [
      "SAY: You have a diagram of a frog's toe pad. Where does it go?",
      "Do not call out. This one goes on boards.",
    ],
    [
      "ASK: Where does the diagram belong?",
      "30 sec. Boards up on cue.",
      "EXPECT: B.",
    ],
    [
      "SCAN boards, back row first.",
      "80%+ -> cold call one B: why not the last page?",
      "Less -> read A, ask when the reader would see it, re-ask.",
    ],
    [
      "REVEAL after boards are scanned.",
      "SAY: A picture belongs beside the words it explains.",
    ],
  ],
  trap: [
    "putting every picture together at the end.",
    "Fix: ask which paragraph it explains, student moves it.",
  ],
  prep: "The hinge of the lesson. A is the all-pictures-at-the-end habit and C is decoration before content. Both are common.",
  tag: "[CFU hinge | Supported application | SC2 | HITS 7, 8]",
});

const NOTES_WEDO = composeGlanceNotes({
  answer: "page 1. Page 2 has no title, no subheadings and no caption.",
  beats: [
    [
      "SAY: Two pages, same report. One is ready. One is not.",
      "Check both against the five rules.",
    ],
    [
      "ASK: Name two rules page 2 breaks.",
      "60 sec. Write two... chin it... show me.",
      "EXPECT: no subheadings, no caption.",
    ],
    [
      "SCAN boards for two named rules.",
      "80%+ -> cold call one: which one would annoy a reader most?",
      "Less -> cover page 1, ask what is missing, re-ask.",
    ],
  ],
  trap: [
    "saying page 2 looks boring rather than naming a rule.",
    "Fix: ask which of the five, student names it.",
  ],
  stretch: "Name a third rule page 2 breaks.",
  help: "Point at one missing feature. Student names its rule.",
  prep: "Judging a finished page against named rules is what makes their own layout choices deliberate rather than decorative.",
  sources: P.SOURCE_LINE_MODEL,
  tag: "[We Do | Supported application | SC2 | HITS 4, 5]",
});

const NOTES_YOUDO = composeGlanceNotes({
  answer: "a published report, laid out to the five rules, and two slips written.",
  beats: [
    [
      "SAY: Publish yours. Five rules, and your visual beside its paragraph.",
      "Then we read each other's work.",
    ],
    [
      "POINT to the steps. TIME: twenty-five minutes to publish.",
      "Then reports out on desks, and everyone walks.",
    ],
    [
      "CIRCULATE while they publish. Check subheadings match their aspects.",
      "COLLECT two slips to read at the closing.",
    ],
  ],
  trap: [
    "spending the time on borders and colours.",
    "Fix: ask which of the five rules that follows, student redirects.",
  ],
  stretch: "Write a slip for a report on a topic nothing like yours.",
  help: "Publish one paragraph and its visual to the five rules.",
  prep: "The gallery walk is the audience the whole unit has been for. Protect the last ten minutes for it, whatever the publishing state.",
  sources: P.SOURCE_LINE_UNIT,
  tag: "[You Do | Mastery and application | SC3 | HITS 5, 10]",
});

const NOTES_EXIT = composeGlanceNotes({
  answer: "one author's name plus one specific thing they did.",
  beats: [
    [
      "SAY: One report you read today. Name the author.",
      "Then name one thing they did that you will use next time.",
    ],
    [
      "TIME: three minutes. Write it on the back of your guide.",
      "COLLECT by walking the rows, not by calling names.",
    ],
    [
      "SORT into two piles as you collect: specific, or vague.",
      "A thick vague pile -> the class needs more practice naming craft, not more writing.",
    ],
  ],
  prep: "Assesses the core target from the reader's side. Specific answers name a technique; vague answers say it was good.",
  tag: "[Exit Ticket | Mastery and application | SC3 | HITS 8]",
});

const NOTES_CLOSING = composeGlanceNotes({
  answer: "open - most students should show two or three.",
  beats: [
    "READ the two collected slips aloud. SAY: Listen for the specific thing.",
    [
      "ASK: Rate yourself on the three statements.",
      "15 sec. Fingers at your chest, one, two or three. Show me.",
      "EXPECT: most on two or three.",
    ],
    "SAY: You read reports like writers, then you wrote one. That is the whole unit.",
  ],
  prep: "End of the unit. Keep the published reports up: they become the mentor texts for the next information report unit.",
  tag: "[Closing | Retention and recall | HITS 1, 9]",
});

/* --- Slide content -------------------------------------------------------- */

const CFU_OPTIONS = [
  { key: "A", text: "On the last page, with all the other pictures." },
  { key: "B", text: "Beside the appearance paragraph, with the caption under it." },
  { key: "C", text: "At the very top, before the title, so it looks good." },
];

const PUBLISH_RULE_CARDS = [
  { label: "Title", detail: "One clear title at the top." },
  { label: "Subheadings", detail: "They match the aspects on your plan." },
  { label: "One font", detail: "Same font and size for all body text." },
  { label: "Visual in place", detail: "Beside its paragraph, caption under it." },
];

/* --- Build ---------------------------------------------------------------- */

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // 1. Title
  titleSlide(pres, "Publish and Share",
    "Lay it out so a reader can use it.",
    "Week 10 Session 4 | Year 5/6 Literacy", NOTES_TITLE);

  // 2. Teacher Resources
  addResourceSlide(pres, {
    resources: RESOURCE_ITEMS,
    studentTools: [
      "Their edited draft",
      "The Publishing Guide, with the slips already cut",
      "Good paper, devices or a slide template",
      "Mini-whiteboards for the two checks",
    ],
    boardSetup: [
      "Clear a wall or a run of desks for the gallery walk",
      "Protect the last ten minutes for the walk",
    ],
  }, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // 3. Launch - three ways to publish
  (() => {
    const s = P.customSlide(pres, T, "Launch", C.SECONDARY, "Three ways to publish. Which suits yours?");
    const y0 = CONTENT_TOP;
    const colGap = 0.22;
    const colW = (9 - colGap * 2) / 3;
    const colors = [C.PRIMARY, C.SECONDARY, C.ASSESS];

    P.PUBLISH_FORMATS.forEach((fmt, i) => {
      const cx = 0.5 + i * (colW + colGap);
      s.addShape("roundRect", {
        x: cx, y: y0, w: colW, h: 2.85, rectRadius: 0.08,
        fill: { color: C.WHITE }, line: { color: colors[i], width: 1.5 },
      });
      addTextOnShape(s, fmt.label, {
        x: cx, y: y0, w: colW, h: 0.62, rectRadius: 0.08,
        fill: { color: colors[i] },
      }, {
        fontSize: 19, fontFace: FONT_H, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      s.addText(fmt.detail, {
        x: cx + 0.20, y: y0 + 0.72, w: colW - 0.40, h: 2.00,
        fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL,
        align: "center", valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
    });

    addTextOnShape(s, "Turn and tell: which one suits your topic, and why?", {
      x: 0.5, y: y0 + 2.98, w: 9, h: 0.62, rectRadius: 0.08,
      fill: { color: C.SECONDARY },
    }, {
      fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_LAUNCH);
    runSlideDiagnostics(s, pres);
  })();

  // 4. Learning Intention and Success Criteria
  liSlide(pres,
    "We are learning to publish an information report so it is clear and easy to read.",
    [
      "I can lay out my report with a title and subheadings.",
      "I can place my visual and caption where they help the reader.",
      "I can give another author specific feedback about their report.",
    ],
    NOTES_LI, FOOTER);

  // 5. I Do - the rules of a published page
  annotatedModelSlide(pres, "I Do", "What a published page looks like",
    [
      { text: "Five rules", role: "header" },
      { text: "One title at the top." },
      { text: "Subheadings that match your aspects." },
      { text: "One font, one size, for the body." },
      { text: "The visual beside its paragraph.", role: "emphasis" },
      { text: "The caption directly underneath." },
    ],
    "A published report page",
    PUBLISH_RULE_CARDS,
    NOTES_IDO, FOOTER,
    {
      badgeColor: C.PRIMARY,
      previewSpec: P.reportPageSpec(C, P.FROG),
      sourceType: "Model report",
    });

  // 6. CFU hinge - where does the diagram belong?
  (() => {
    const s = P.customSlide(pres, T, "CFU", C.ALERT, "Where does the diagram belong?");
    const y0 = CONTENT_TOP;

    addTextOnShape(s, "You have a labelled diagram of a frog's toe pad.", {
      x: 0.5, y: y0, w: 9, h: 0.48, rectRadius: 0.08,
      fill: { color: C.ALERT },
    }, {
      fontSize: 16, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    P.drawOptionStack(s, T, CFU_OPTIONS, {
      y: y0 + 0.60, optionH: 0.66, gap: 0.12, color: C.PRIMARY, fontSize: 17,
    });

    clickBuild(s, [
      () => {
        addRevealAnswerBar(s, ["B - beside the words it explains"], {
          y: P.REVEAL_Y, h: P.REVEAL_H, fontSize: 22, color: C.SUCCESS,
        });
      },
    ]);

    addFooter(s, FOOTER);
    s.addNotes(NOTES_CFU);
    runSlideDiagnostics(s, pres);
  })();

  // 7. We Do - which page is ready?
  compareVisualSlide(pres, "We Do", "Which page is ready to publish?",
    "Write two rules that page 2 breaks. Chin it, show me.",
    {
      panelTitle: "Page 1",
      title: "Title, subheadings, caption",
      strip: C.SUCCESS,
      previewSpec: P.reportPageSpec(C, P.FROG),
      previewH: 2.10,
    },
    {
      panelTitle: "Page 2",
      title: "Same words, different layout",
      strip: C.ALERT,
      previewSpec: P.messyPageSpec(C, P.FROG),
      previewH: 2.10,
    },
    NOTES_WEDO, FOOTER,
    // Enlarged from the 2.15" default: students have to judge these pages
    // against five named rules, so they have to be legible.
    { badgeColor: C.SUCCESS, badgeFill: C.SUCCESS, badgeW: 1.5,
      promptFill: C.SUCCESS, cardH: 3.00 });

  // 8. You Do - publish, then read each other's
  (() => {
    const s = P.customSlide(pres, T, "You Do", C.ASSESS, "Publish it, then read each other's");
    const y0 = CONTENT_TOP;
    const h = SAFE_BOTTOM - y0;

    T.addInstructionCard(s, [
      { text: "Publish, then walk", role: "header" },
      { text: "1. Lay yours out using the five rules." },
      { text: "2. Report on your desk, open at page one." },
      { text: "3. Walk. Read three. Write two slips." },
      { text: "Twenty-five minutes, then the walk." },
    ], {
      x: 0.5, y: y0, w: 4.7, h,
      strip: C.ASSESS, fill: C.WHITE,
      headerColor: C.ASSESS, emphasisColor: C.ALERT,
    });

    const rx = 5.5;
    const rw = 4.0;
    // Header 0.50 at 1.30, five rules from 1.92 at 0.56 with 0.08 gaps
    // finish at 5.04, clear of the 5.10 floor.
    addTextOnShape(s, "The five rules", {
      x: rx, y: y0, w: rw, h: 0.50, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, {
      fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
    P.PUBLISH_RULES.forEach((rule, i) => {
      const ry = y0 + 0.62 + i * 0.64;
      s.addShape("roundRect", {
        x: rx, y: ry, w: rw, h: 0.56, rectRadius: 0.07,
        fill: { color: C.WHITE }, line: { color: C.ACCENT, width: 1.3 },
      });
      s.addShape("roundRect", {
        x: rx + 0.14, y: ry + 0.14, w: 0.28, h: 0.28, rectRadius: 0.14,
        fill: { color: C.ACCENT },
      });
      s.addText(String(i + 1), {
        x: rx + 0.14, y: ry + 0.14, w: 0.28, h: 0.28,
        fontSize: 12, fontFace: FONT_H, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      s.addText(rule, {
        x: rx + 0.52, y: ry, w: rw - 0.68, h: 0.56,
        fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL,
        valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
    runSlideDiagnostics(s, pres);
  })();

  // 9. Exit ticket
  P.exitTicketPanel(pres, T, {
    topic: "One report you read on the gallery walk",
    task: "Name the author, and one thing they did that you will use next time.",
    cue: "One name. One specific thing.",
    taskSize: 25,
    footer: FOOTER,
    notes: NOTES_EXIT,
  });

  // 10. Closing
  closingSlide(pres, {
    reflectionPrompt: "Your report is finished. What would you do differently on the next one?",
    scItems: [
      "I can lay out my report with a title and subheadings.",
      "I can place my visual and caption where they help the reader.",
      "I can give another author specific feedback about their report.",
    ],
    selfAssessment: "Fingers at your chest. One, two or three. Show me.",
    takeaways: [P.ANCHOR_PHRASE],
  }, NOTES_CLOSING);

  const pptxPath = path.join(OUT_DIR, "Information Reports W10 S4.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);
}

/* --- Resource: the Publishing Guide --------------------------------------- */

async function buildGuide() {
  const doc = createPdf({ title: GUIDE_RES.name });
  let y = addPdfHeader(doc, "Publishing Guide", {
    subtitle: "Five rules, three formats, and the slips for the gallery walk.",
    color: C.PRIMARY,
    lessonInfo: "Year 5/6 Literacy | Information Reports | Week 10",
  });

  y = addTipBox(doc, P.ANCHOR_PHRASE, y, { color: C.ACCENT });

  y = addSectionHeading(doc, "The five rules of a published page", y, { color: C.PRIMARY });
  y = P.addChecklistPdf(doc, y, P.PUBLISH_RULES, { color: hex(C.PRIMARY) });
  y += 4;

  y = addSectionHeading(doc, "Choose your format", y, { color: C.SECONDARY });
  P.PUBLISH_FORMATS.forEach((fmt, i) => {
    const accent = hex([C.PRIMARY, C.SECONDARY, C.ASSESS][i]);
    doc.save();
    doc.roundedRect(PAGE.MARGIN, y, 110, 20, 3).fill(accent);
    doc.fontSize(9.5).font("Sans-Bold").fillColor("#FFFFFF");
    doc.text(fmt.label, PAGE.MARGIN, y + 6, { width: 110, align: "center", lineBreak: false });
    doc.restore();
    doc.fontSize(10.5).font("Sans").fillColor("#000000");
    doc.text(fmt.detail, PAGE.MARGIN + 120, y + 4, { width: PAGE.CONTENT_W - 128 });
    y = Math.max(doc.y, y + 20) + 8;
  });
  y += 4;

  y = addBodyText(doc,
    "Whichever format you choose, the five rules are the same. A report that looks busy is harder to read, not better.",
    y, { italic: true });

  y = addSectionHeading(doc, "Before you print or present", y, { color: C.SUCCESS });
  y = P.addChecklistPdf(doc, y, [
    "My subheadings are the same aspects as my plan.",
    "Every visual sits beside the paragraph it explains.",
    "Every visual has a numbered caption underneath.",
    "One font, one size, for all the body text.",
    "I have read it out loud once, to catch what my eyes missed.",
  ], { color: hex(C.SUCCESS) });

  doc.addPage();
  y = PAGE.MARGIN;

  y = addSectionHeading(doc, "Gallery walk slips", y, { color: C.ASSESS });
  y = addBodyText(doc,
    "Cut along the lines. Read three reports. Leave a slip on two of them. Be specific: name the thing you liked, not just that you liked it.", y);
  y += 6;

  for (let i = 0; i < 4; i += 1) {
    const boxH = 130;
    doc.save();
    doc.roundedRect(PAGE.MARGIN, y, PAGE.CONTENT_W, boxH, 4)
      .lineWidth(1).strokeColor(hex(C.ASSESS)).dash(4, { space: 3 }).stroke();
    doc.undash();
    doc.restore();

    doc.fontSize(10).font("Sans-Bold").fillColor(hex(C.ASSESS));
    doc.text("To the author of:", PAGE.MARGIN + 12, y + 10, { width: 110, lineBreak: false });
    doc.moveTo(PAGE.MARGIN + 128, y + 22).lineTo(PAGE.MARGIN + PAGE.CONTENT_W - 12, y + 22)
      .lineWidth(0.9).strokeColor("#000000").stroke();

    doc.fontSize(10).font("Sans-Bold").fillColor(hex(C.SUCCESS));
    doc.text("One thing you did well, and why it worked:", PAGE.MARGIN + 12, y + 34,
      { width: PAGE.CONTENT_W - 24, lineBreak: false });
    [52, 74].forEach((off) => {
      doc.moveTo(PAGE.MARGIN + 12, y + off).lineTo(PAGE.MARGIN + PAGE.CONTENT_W - 12, y + off)
        .lineWidth(0.9).strokeColor("#000000").stroke();
    });

    doc.fontSize(10).font("Sans-Bold").fillColor(hex(C.ACCENT));
    doc.text("One thing I am taking for my own next report:", PAGE.MARGIN + 12, y + 86,
      { width: PAGE.CONTENT_W - 24, lineBreak: false });
    [104, 122].forEach((off) => {
      doc.moveTo(PAGE.MARGIN + 12, y + off).lineTo(PAGE.MARGIN + PAGE.CONTENT_W - 12, y + off)
        .lineWidth(0.9).strokeColor("#000000").stroke();
    });

    y += boxH + 14;
  }

  addPdfFooter(doc, "Information Reports | Year 5/6 Literacy | Week 10");
  await writePdf(doc, path.join(OUT_DIR, GUIDE_RES.fileName));
  console.log("Wrote " + GUIDE_RES.name);
}

(async () => {
  await build();
  await buildGuide();
})().catch((err) => {
  console.error(err);
  process.exit(1);
});
