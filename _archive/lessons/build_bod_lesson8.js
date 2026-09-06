"use strict";

// BODMAS Unit - Session 8: Calendar Project - 2-Operation Equations
// Year 5/6 Numeracy, Week 9 Term 2 (variant 2)

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addProblem, addWriteLine,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const T = createTheme("numeracy", "grade56", weekToVariant(9));
const {
  C, FONT_H, FONT_B,
  SLIDE_W, SLIDE_H, SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  workedExSlide, exitTicketSlide, addStageBadge,
  addTextOnShape, addCard, addFooter, addTopBar, addBadge, addTitle,
  withReveal,
  STAGE_COLORS,
} = T;

const SESSION = 8;
const FOOTER = "BODMAS | Session 8 of 10 - Calendar Project | Year 5/6 Numeracy";
const OUT_DIR = "output/BOD_Session8_Calendar_Two_Ops";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const HARD_DATES_RES = makeSessionResource(SESSION, "Tricky Dates Hint Sheet", "Strategies for the harder target numbers - 1, 11, 13, 17, 19, 22, 23.");
const RESOURCE_ITEMS = [HARD_DATES_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// --- Teacher Notes -----------------------------------------------------------

const NOTES_TITLE = `SAY:
- Day 2 of the calendar project
- Today: focus on getting all 14 dates that need 2 operations done
- We'll learn strategies for the tricky dates

DO:
- Display title slide
- Have draft sheets and calendars ready - students collect from yesterday's pile
- Have hint sheet ready

TEACHER NOTES:
Session 8 of 10. Today is the bulk-progress day on the project. By the end of today, 14-21 of the 28 dates should be done. We focus on 2-operation equations because they're easier and there are 14 of them required.

WATCH FOR:
- Students who haven't picked a month yet - help them now
- Students who finished week 1 quickly - give them tricky dates to focus on

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_DR_Q = `SAY:
- Quick warm-up - 2 BODMAS equations
- Show me when done

DO:
- Display 2 equations
- 3 minutes

TEACHER NOTES:
Short Daily Review - we want maximum project time.

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Tick or fix - quick

DO:
- Reveal answers

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_LI_SC = `SAY:
- Today's intention: complete the 14 dates with 2 operations
- Read SC together

DO:
- Choral read
- Brief

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_HARD_DATES = `SAY:
- Some dates are TRICKY with just 2 operations and digits 2,3,4,5
- Date 1: 4 - 3 + 2 - 2 nope reused 2... 5 - 4 = 1 (just 1 op)... (3 + 2) - 4 = 1 YES
- Date 11: 5 + 4 + 2 = 11 YES (uses 3 digits, 2 ops)
- Date 13: 5 + 4 + 4 nope... 5 x 3 - 2 = 13 YES (3 digits, 2 ops)
- Date 17: 5 x 4 - 3 = 17 YES (3 digits, 2 ops)
- Date 19: 5 x 4 - 3 + 2 = 19 (uses 4 digits, 3 ops - keep for 3-op slot)
- Date 22: 5 x 4 + 2 = 22 YES (3 digits, 2 ops)
- Date 23: 5 x 4 + 3 = 23 YES (3 digits, 2 ops)

DO:
- Display the hint slide
- Read each tricky date and one possible equation
- Don't give EVERY answer - leave some discovery

TEACHER NOTES:
The hint sheet does NOT give every answer - it shows STRATEGIES for the trickier dates. Students still need to verify and personalise.

WATCH FOR:
- Students who copy verbatim - prompt them to find a different equation
- Students who use the hints to break through and find their own - excellent

[Stage 2: I Do | VTLM 2.0: Strategy Modelling]`;

const NOTES_QUALITY_CHECK = `SAY:
- Quality check - look at one of your draft equations
- Run the rubric checks
- Does it equal the date? Does it use the right digits? Right number of operations?
- Tick if all yes - then write neatly on calendar

DO:
- Walk through the rubric again
- Ask students to run a check on one of their drafts

TEACHER NOTES:
Quality check is a habit they need. By session 10 they should run this check automatically. Today we model it.

WATCH FOR:
- Students who skip the check - prompt them
- Students who catch their own errors during the check - this is the win

[Stage 2: I Do | VTLM 2.0: Quality Modelling]`;

const NOTES_CFU1 = `SAY:
- Quick check. On your whiteboards
- Date 11 - build an equation
- Need: 3 digits from 2,3,4,5 (each once), 2 different operations
- 60 seconds

DO:
- Students build and show
- Examples: 5 + 4 + 2 = 11 (uses + only - might be 2 ops though - + and +, only 1 different op!)
- Better: 5 + 2 x 3 = 11 (uses + and x - 2 different ops)
- Or: 4 x 3 - 2 + 1 nope no 1
- Or: 3 x 4 - 2 + ? nope

TEACHER NOTES:
This CFU shows the trickiness of date 11 with the rule "2 DIFFERENT operations". 5+4+2 has 2 plus signs but only 1 DIFFERENT operation. Students need to think carefully.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Build equation for 11 using 3 digits (each once) and 2 DIFFERENT operations
- Show me
PROCEED:
- Most students show valid equation - go to project
PIVOT:
- Many show 5+4+2 (only 1 different op): clarify "different operations" means different SYMBOLS

WATCH FOR:
- Students with 5+4+2 - prompt: "How many different operations? + and + is only ONE different op"
- Students with 5 + 2 x 3 - excellent, two different operations

[Stage 2: CFU | VTLM 2.0: Active Checking]`;

const NOTES_CFU1_A = `SAY:
- Two different operations means + and x, or + and -, or x and / etc
- 5 + 4 + 2 has only ONE different operation (just +)
- 5 + 2 x 3 has TWO different operations (+ and x) - BETTER for the rule
- (5 - 2) x 4 - nope, that's 12 not 11
- 4 x 3 - 2 + 1 no 1 available
- Many right answers - keep trying

DO:
- Reveal sample answers
- Discuss the "different operations" rule

[Stage 2: CFU Answer | VTLM 2.0: Active Checking]`;

const NOTES_PROJECT_TIME = `SAY:
- Project time
- Today: focus on the 14 dates that need 2 operations
- That's about half your calendar
- Use the hint sheet for the trickier dates
- Use the rubric to check your work

DO:
- Hand out hint sheet to anyone who needs it
- Allow 35-40 minutes
- Circulate
- Pull a small group for any students struggling with the project

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Pull small group. Use the hint sheet. Start with the easiest dates - 5, 6, 7. Build confidence
- Extra Notes: Give them a "this is a 2-op equation" example for each tricky date
EXTENDING PROMPT:
- Task: Strong students - try to get the same date with TWO different equations. Or attempt some 3-op equations early
- Extra Notes: Give them creative freedom

WATCH FOR:
- Students close to finishing 2-op section - move them to 3-op
- Students stuck on tricky dates - small group support
- Quality - some students rush. Pull them back to the rubric

[Stage 4: Project | VTLM 2.0: Independent Practice]`;

const NOTES_EXIT = `SAY:
- Exit ticket
- Three quick questions

DO:
- Hand out exit ticket
- Collect

TEACHER NOTES:
Exit ticket tracks progress. Use the data to plan tomorrow.

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Pack up draft sheets and calendars
- Read SC together
- Self-check
- Tomorrow we tackle the 3-operation dates

DO:
- Closing slide
- Pack up
- Stack by table

TEACHER NOTES:
Tomorrow is the harder section - 3 operations. Students who finished 2-op today are ahead. Use the data.

[General: Closing | VTLM 2.0: Reflection]`;

const NOTES_RESOURCES = `SAY:
- Hint sheet for the tricky dates - take one if you need

DO:
- Hand out

[General: Resources | VTLM 2.0: Resource Awareness]`;

// --- Build -------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres,
    "Calendar - Two-Operation Day",
    "Complete your 14 two-operation dates",
    "Session 8 of 10  |  Year 5/6 Numeracy  |  Project Day 2",
    NOTES_TITLE);

  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slide 2-3: Daily Review (short)
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Quick BODMAS Warm-Up", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });

      addCard(s, 0.5, CONTENT_TOP, 9, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });

      addTextOnShape(s, "1.   2 squared + (8 - 5) x 4", {
        x: 1.0, y: CONTENT_TOP + 0.4, w: 8.0, h: 1.0, rectRadius: 0.1,
        fill: { color: STAGE_COLORS["1"] },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });

      addTextOnShape(s, "2.   30 - 6 x 4 + 12 / 3", {
        x: 1.0, y: CONTENT_TOP + 1.7, w: 8.0, h: 1.0, rectRadius: 0.1,
        fill: { color: STAGE_COLORS["1"] },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });

      s.addText("Solve in your books  |  3 minutes", {
        x: 0.5, y: CONTENT_TOP + 2.95, w: 9, h: 0.4,
        fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, align: "center", italic: true, margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "1)  4 + 3 x 4 = 16    2)  30 - 24 + 4 = 10", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slide 4: LI/SC
  liSlide(pres,
    ["I am completing the 14 two-operation dates on my BODMAS Calendar"],
    [
      "I can use 'two different operations' correctly (not the same operation twice)",
      "I can build an equation for a tricky date using a strategy",
      "I can check my equation against the rubric before writing it on the calendar",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 5: The "two different operations" trap
  contentSlide(pres, "I Do", C.PRIMARY, "Two DIFFERENT Operations",
    [
      "Read the rule again carefully",
      "'Two different operations' - not just two operation symbols",
      "",
      "5 + 4 + 2 has TWO + signs - but only ONE different operation",
      "5 + 2 x 3 has + AND x - that's TWO different operations",
      "",
      "Different = different SYMBOL",
    ],
    NOTES_QUALITY_CHECK, FOOTER,
    (slide, lg) => {
      // Comparison panel
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 1.6, { strip: C.MUTED });
      slide.addText("Doesn't meet rule", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 13, fontFace: FONT_B, color: C.MUTED, bold: true, margin: 0,
      });
      slide.addText("5 + 4 + 2", {
        x: lg.rightX, y: lg.panelTopPadded + 0.42, w: lg.rightW, h: 0.5,
        fontSize: 22, fontFace: FONT_H, color: C.CHARCOAL, bold: true, align: "center", margin: 0,
      });
      addTextOnShape(slide, "Only 1 different op", {
        x: lg.rightX + 0.6, y: lg.panelTopPadded + 0.95, w: lg.rightW - 1.2, h: 0.45, rectRadius: 0.06,
        fill: { color: C.MUTED },
      }, { fontSize: 12, fontFace: FONT_B, color: C.WHITE, bold: true });

      addCard(slide, lg.rightX, lg.panelTopPadded + 1.8, lg.rightW, 1.6, { strip: C.SUCCESS });
      slide.addText("Meets rule", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 1.88, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 13, fontFace: FONT_B, color: C.SUCCESS, bold: true, margin: 0,
      });
      slide.addText("5 + 2 x 3", {
        x: lg.rightX, y: lg.panelTopPadded + 2.22, w: lg.rightW, h: 0.5,
        fontSize: 22, fontFace: FONT_H, color: C.CHARCOAL, bold: true, align: "center", margin: 0,
      });
      addTextOnShape(slide, "+ and x = 2 diff ops", {
        x: lg.rightX + 0.6, y: lg.panelTopPadded + 2.75, w: lg.rightW - 1.2, h: 0.45, rectRadius: 0.06,
        fill: { color: C.SUCCESS },
      }, { fontSize: 12, fontFace: FONT_B, color: C.WHITE, bold: true });
    }
  );

  // Slide 6: Tricky dates strategies
  contentSlide(pres, "I Do", C.SECONDARY, "Tricky Dates - Strategies",
    [
      "Some dates are tricky with only 2,3,4,5",
      "Strategy 1: Start with multiplication (5x4=20, 5x3=15, 4x3=12)",
      "Strategy 2: Use brackets to create a useful intermediate",
      "Strategy 3: Use subtraction to reach a smaller target",
      "",
      "Examples on the hint sheet for dates: 1, 11, 13, 17, 22, 23",
    ],
    NOTES_HARD_DATES, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.4, { strip: C.SECONDARY });
      slide.addText("Sample tricky dates", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 13, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0, align: "center",
      });

      const tricky = [
        { d: "1",  eq: "(3 + 2) - 4" },
        { d: "11", eq: "5 + 2 x 3" },
        { d: "17", eq: "5 x 4 - 3" },
        { d: "22", eq: "5 x 4 + 2" },
      ];
      tricky.forEach((t, i) => {
        const y = lg.panelTopPadded + 0.45 + i * 0.55;
        addTextOnShape(slide, t.d, {
          x: lg.rightX + 0.2, y, w: 0.55, h: 0.45, rectRadius: 0.06,
          fill: { color: C.SECONDARY },
        }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
        slide.addText(t.eq, {
          x: lg.rightX + 0.85, y, w: lg.rightW - 1.05, h: 0.45,
          fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, valign: "middle", margin: 0,
        });
      });

      slide.addText("These are SAMPLES - find your own!", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 2.95, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 11, fontFace: FONT_B, color: C.MUTED, italic: true, align: "center", margin: 0,
      });
    }
  );

  // Slide 7-8: CFU 1
  withReveal(
    () => cfuSlide(pres, "CFU", "Build for Date 11", "Show Me Boards",
      "Build an equation for date 11.\n\nUse 3 digits from 2, 3, 4, 5 (each once).\nUse 2 DIFFERENT operations.",
      NOTES_CFU1, FOOTER),
    (slide) => {
      addTextOnShape(slide, "Examples: 5 + 2 x 3 = 11   |   2 x 4 + 3 = 11   |   5 x 3 - 4 = 11", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 13, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_CFU1_A);
    }
  );

  // Slide 9: Project Time
  workedExSlide(pres, 4, "Project Time", "Two-Operation Day",
    [
      "Goal: 14 of your dates need 3 digits + 2 different operations",
      "Use the hint sheet for tricky dates",
      "Use the rubric to check before writing on the calendar",
      "",
      "Tip: 'Two different operations' = two different SYMBOLS",
      "  + and + is only ONE different operation",
      "  + and x is TWO different operations",
      "",
      "If you finish your 14 two-op dates, start three-op tomorrow's work",
      "",
      "35 minutes",
    ],
    NOTES_PROJECT_TIME, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.4, { strip: C.PRIMARY });
      slide.addText("Today's targets", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0, align: "center",
      });

      // Summary panel
      slide.addText("By the end of today:", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.48, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, bold: true, margin: 0,
      });

      slide.addText([
        { text: "14 of 28 dates DONE", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Each with 3 digits", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Each with 2 different ops", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Each rubric-checked", options: { bullet: true, fontSize: 13, color: C.SUCCESS, bold: true } },
      ], {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.85, w: lg.rightW - 0.5, h: 1.7,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addTextOnShape(slide, "35 minutes", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 2.7, w: lg.rightW - 0.4, h: 0.5, rectRadius: 0.08,
        fill: { color: C.ALERT },
      }, { fontSize: 16, fontFace: FONT_B, color: C.WHITE, bold: true });
    }
  );

  // Slide 10: Exit Ticket
  exitTicketSlide(pres,
    [
      "How many dates have you completed in TOTAL?  __ / 28",
      "Which date was your favourite equation today? Write it below.",
      "Which date is still missing? What's your plan?",
    ],
    NOTES_EXIT, FOOTER);

  // Slide 11: Closing
  closingSlide(pres, {
    reflectionPrompt: "Tell your partner: how is your calendar coming along? What's left?",
    scItems: [
      "I can use 'two different operations' correctly",
      "I can build an equation for a tricky date using a strategy",
      "I can check my equation against the rubric before writing it on the calendar",
    ],
    selfAssessment: { prompt: "Self-check", options: ["Got it", "Getting there", "Need more practice"] },
  }, NOTES_CLOSING);

  // Slide 12: Resources


  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "BOD_Session8_Calendar_Two_Ops.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // --- PDFs ------------------------------------------------------------------

  await (async () => {
    const doc = createPdf({ title: HARD_DATES_RES.name });
    let y = addPdfHeader(doc, HARD_DATES_RES.name, {
      subtitle: "Strategies for the Tricky Dates",
      color: C.NAVY,
      lessonInfo: "Session 8 of 10 | Year 5/6 Numeracy",
    });
    y = addTipBox(doc, "These are STRATEGY HINTS, not the only answers. Try the strategy, then find YOUR own equation. Each equation must use 3 digits from 2, 3, 4, 5 (each once) and 2 different operations.", y, { color: C.TEAL });

    y = addSectionHeading(doc, "Easy dates (lots of options)", y, { color: C.NAVY });
    y = addBodyText(doc, "Date 5:    4 + 3 - 2 = 5      OR     5 - 3 + 4 - 1 nope", y);
    y = addBodyText(doc, "Date 6:    5 + 3 - 2 = 6      OR     2 x 4 - 3 nope only one 4", y);
    y = addBodyText(doc, "Date 7:    (5 + 4) - 2 = 7    OR     5 + 4 - 2 = 7", y);
    y = addBodyText(doc, "Date 8:    (5 - 3) x 4 = 8    OR     2 x 5 - 3 + 1 nope no 1", y);
    y = addBodyText(doc, "Date 9:    5 + 2 + 4 - 2 nope reused 2     2 + 3 + 4 = 9", y);

    y = addSectionHeading(doc, "Tricky dates - strategies", y, { color: C.NAVY });
    y = addBodyText(doc, "Date 1:    Start with brackets to make a small target", y);
    y = addBodyText(doc, "           Try: (3 + 2) - 4 = 5 - 4 = 1   YES (uses brackets and -)", y);
    y = addBodyText(doc, "Date 11:   Strategy: use multiplication for the bulk", y);
    y = addBodyText(doc, "           Try: 5 + 2 x 3 = 5 + 6 = 11   YES", y);
    y = addBodyText(doc, "Date 13:   Use 5x3 = 15, then subtract", y);
    y = addBodyText(doc, "           Try: 5 x 3 - 2 = 13   YES", y);
    y = addBodyText(doc, "Date 17:   Use 5x4 = 20, subtract a digit", y);
    y = addBodyText(doc, "           Try: 5 x 4 - 3 = 17   YES", y);
    y = addBodyText(doc, "Date 22:   Use 5x4 = 20, add a digit", y);
    y = addBodyText(doc, "           Try: 5 x 4 + 2 = 22   YES", y);
    y = addBodyText(doc, "Date 23:   Use 5x4 = 20, add a digit", y);
    y = addBodyText(doc, "           Try: 5 x 4 + 3 = 23   YES", y);

    y = addSectionHeading(doc, "Other dates - try yourself", y, { color: C.NAVY });
    y = addBodyText(doc, "Date 2, 3, 4, 10, 12, 14, 15, 16, 18, 20, 21, 24, 25, 26, 27, 28", y);
    y = addBodyText(doc, "Use the strategies above. Many right answers exist!", y);

    y = addSectionHeading(doc, "Reminder: Rule for two-op equations", y, { color: C.NAVY });
    y = addBodyText(doc, "  3 digits from 2, 3, 4, 5 (each used only once)", y);
    y = addBodyText(doc, "  2 DIFFERENT operations (not the same op twice)", y);
    y = addBodyText(doc, "  Brackets are a tool - don't count as an operation", y);

    addPdfFooter(doc, "Session 8 | Tricky Dates Hint Sheet | Year 5/6");
    await writePdf(doc, path.join(OUT_DIR, HARD_DATES_RES.fileName));
    console.log("PDF written: " + HARD_DATES_RES.fileName);
  })();

  console.log("Session 8 build complete.");
}

build().catch(err => { console.error(err); process.exit(1); });
