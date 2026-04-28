"use strict";

// BODMAS Unit - Session 7: Calendar Project - Launch and Plan
// Year 5/6 Numeracy, Week 9 Term 2 (variant 2)
// First of 4 calendar sessions

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

const SESSION = 7;
const FOOTER = "BODMAS | Session 7 of 10 - Calendar Project | Year 5/6 Numeracy";
const OUT_DIR = "output/BOD_Session7_Calendar_Launch";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const CALENDAR_RES = makeSessionResource(SESSION, "Calendar Template", "A3 calendar template - print one per student.");
const DRAFT_RES = makeSessionResource(SESSION, "Equation Draft Sheet", "Workbook-style draft sheet for trying equations before transferring.");
const RUBRIC_RES = makeSessionResource(SESSION, "Project Rubric", "Teacher rubric and student-friendly success criteria for the calendar.");
const RESOURCE_ITEMS = [CALENDAR_RES, DRAFT_RES, RUBRIC_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// --- Teacher Notes -----------------------------------------------------------

const NOTES_TITLE = `SAY:
- Welcome back to BODMAS
- This week is the project we have been building towards
- We're creating a BODMAS Calendar for our birth months
- Today is launch day - you'll plan your month and start your first week of equations

DO:
- Display title slide
- Have calendar template, draft sheet, and rubric ready to hand out
- Create energy - this IS a special project

TEACHER NOTES:
Session 7 of 10. The launch is critical - students must understand the project AND start producing equations today. Don't get bogged down in admin - get them building equations within 20 minutes of starting.

WATCH FOR:
- Students who already know their birth month - good
- Students unsure - they can pick any month, doesn't have to be birth

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_DR_Q = `SAY:
- Quick warm-up - solve these two equations
- Show me when done

DO:
- Display 2 quick BODMAS equations
- Allow 3 minutes
- Brief - we want time for the project today

TEACHER NOTES:
Shorter Daily Review today - we need maximum time on the project. 2 equations is enough to keep BODMAS warm.

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Tick or fix - quick

DO:
- Reveal answers

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_LI_SC = `SAY:
- Today we launch the calendar
- Read SC together

DO:
- Choral read
- Brief - we want to launch

TEACHER NOTES:
SC are project-focused. SC1 = understand the rules. SC2 = produce a first week of equations. SC3 = write equations using both 2-operation and 3-operation forms.

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_PROJECT_BIG = `SAY:
- The big picture
- Pick a month - your birth month
- 28 days - 28 equations
- Each equation EQUALS the date
- Day 7 equation = 7. Day 14 equation = 14. And so on
- We'll display these in the classroom!

DO:
- Display the project overview
- Show example for day 7 and day 14
- Build excitement

TEACHER NOTES:
The big picture lands today. Tell students the calendars will be displayed - this raises stakes and quality.

WATCH FOR:
- Students who understand immediately - good
- Students who look confused - walk through one example slowly

[Stage 2: I Do | VTLM 2.0: Establishing Purpose]`;

const NOTES_RULES = `SAY:
- The rules
- One: only use the digits 2, 3, 4, 5
- Two: each equation can only use each digit ONCE
- Three: 14 equations need 3 digits AND 2 different operations
- Four: 14 equations need 3 digits AND 3 different operations
- Brackets are a TOOL - they don't count as an operation
- Operations: +, -, x, /, powers (orders)

DO:
- Display rules
- Read each one out loud
- Hand out the rubric for reference

TEACHER NOTES:
Rules are critical - students need to know them so they don't waste effort. The 14/14 split (2 ops vs 3 ops) is the structure of the project. Most days will have 2-op equations, some need 3-op.

WATCH FOR:
- Students confused about the 14/14 - clarify: half the dates use 2 ops, half use 3 ops

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_WORKED = `SAY:
- Let's plan for date 5
- Target: 5
- Need 3 digits and 2 different operations
- Try: 2 + 4 - 3 + ? wait need exactly 3 digits, 2 different operations
- 2 + 4 - 3 = 3 - that's 3, not 5
- Let me try: 4 + 3 - 2 = 5 - YES! 3 digits (4, 3, 2), 2 different operations (+ and -)
- That works for date 5

DO:
- Think aloud at the board
- Show the trial-and-error process
- Mark on the draft sheet how you would record it

TEACHER NOTES:
This is the routine for the whole project: pick a target, try, check, adjust. Model the messiness honestly. Students will hit dead ends - that's normal.

WATCH FOR:
- Students who watch the process and start their own thinking - they're ready to go

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU1 = `SAY:
- Quick check. On your whiteboards
- Build an equation for date 6
- 3 digits from 2, 3, 4, 5 (each once)
- 2 different operations
- 60 seconds

DO:
- Students build and show
- Walk around - check each whiteboard

TEACHER NOTES:
This CFU lets you scan readiness for the project. Most students should produce a valid equation. Common: 5 + 3 - 2 = 6, 4 + 5 - 3 = 6, 2 x 4 - 3 nope only 2 ops still works (x and -), 2 + 4 + ? nope can't reuse.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Build equation for 6 with the rules
- Show me
PROCEED:
- Most students show valid equation - go to project
PIVOT:
- Several students stuck - reteach with another example. Show: 5 + 4 - 3 = 6. Then ask for a different one

WATCH FOR:
- Students with valid equations - ready
- Students stuck - flag for small group during the project

[Stage 2: CFU | VTLM 2.0: Active Checking]`;

const NOTES_CFU1_A = `SAY:
- Many right answers
- 5 + 3 - 2 = 6
- 4 + 5 - 3 = 6
- 2 x 5 - 4 = 6 (this uses x and -, two different operations)

DO:
- Reveal sample answers

[Stage 2: CFU Answer | VTLM 2.0: Active Checking]`;

const NOTES_PROJECT_TIME = `SAY:
- Project time
- You have your calendar template, draft sheet, and rubric
- Today: focus on dates 1 to 7 - the first week
- Use the draft sheet to TRY equations
- Once you've checked it works, write it neatly on the calendar
- Remember the rules

DO:
- Hand out templates if not already
- Allow 25-30 minutes
- Circulate - support stuck students
- Pull a small group for any students who failed the CFU

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Pull small group. Start with date 1 - simplest target. Use 2 digits to start. Build up
- Extra Notes: Build confidence with simple wins before complex equations
EXTENDING PROMPT:
- Task: For students moving fast - try date 7 with brackets, then attempt a 3-operation equation for date 5
- Extra Notes: Stretch the strong workers without making it feel like punishment

WATCH FOR:
- Students who copy the same pattern across dates - encourage variety
- Students stuck on a single date - skip and come back

[Stage 4: Project | VTLM 2.0: Independent Practice]`;

const NOTES_EXIT = `SAY:
- Exit ticket
- Two questions about your calendar today
- One reflection question

DO:
- Hand out exit ticket
- Collect

TEACHER NOTES:
Exit ticket assesses progress on the project, not BODMAS skill. Use it to plan tomorrow's session.

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Pack up your draft sheets and calendars - we'll keep going tomorrow
- Read SC together
- Self-check
- Tell your partner: which date was hardest today?

DO:
- Closing slide
- Self-check
- Pack up

TEACHER NOTES:
Closing previews tomorrow. Self-check identifies who needs help. Stack the draft sheets and calendars by table.

[General: Closing | VTLM 2.0: Reflection]`;

const NOTES_RESOURCES = `SAY:
- Three resources for the project
- Calendar template - your final product
- Draft sheet - try equations here first
- Rubric - shows what makes a strong calendar

DO:
- Confirm everyone has all three

[General: Resources | VTLM 2.0: Resource Awareness]`;

// --- Build -------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres,
    "BODMAS Calendar - Launch",
    "Pick your month. Build 28 equations. One per date.",
    "Session 7 of 10  |  Year 5/6 Numeracy  |  Project Day 1",
    NOTES_TITLE);

  // Slide 2-3: Daily Review (shorter today)
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Quick BODMAS Warm-Up", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });

      addCard(s, 0.5, CONTENT_TOP, 9, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });

      // 2 large equations side by side
      addTextOnShape(s, "1.   3 squared + (10 - 4) - 5", {
        x: 1.0, y: CONTENT_TOP + 0.4, w: 8.0, h: 1.0, rectRadius: 0.1,
        fill: { color: STAGE_COLORS["1"] },
      }, { fontSize: 26, fontFace: FONT_H, color: C.WHITE, bold: true });

      addTextOnShape(s, "2.   20 - 4 x 3 + 6 / 2", {
        x: 1.0, y: CONTENT_TOP + 1.7, w: 8.0, h: 1.0, rectRadius: 0.1,
        fill: { color: STAGE_COLORS["1"] },
      }, { fontSize: 26, fontFace: FONT_H, color: C.WHITE, bold: true });

      s.addText("Solve in your books - show every step  |  3 minutes", {
        x: 0.5, y: CONTENT_TOP + 2.95, w: 9, h: 0.4,
        fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, align: "center", italic: true, margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "1) 9 + 6 - 5 = 10    2) 20 - 12 + 3 = 11", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slide 4: LI/SC
  liSlide(pres,
    ["I am launching my BODMAS Calendar - my month, my equations"],
    [
      "I can name the rules of the calendar project",
      "I can build equations for my first week of dates",
      "I can use 3 digits and 2 different operations in my equations",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 5: Project big picture
  contentSlide(pres, "I Do", C.PRIMARY, "The BODMAS Calendar Project",
    [
      "Pick a month (your birth month)",
      "28 dates - 28 equations",
      "Each equation must EQUAL the date",
      "",
      "Day 7 equation = 7",
      "Day 14 equation = 14",
      "Day 28 equation = 28",
      "",
      "Calendars will be displayed - bring your A-game!",
    ],
    NOTES_PROJECT_BIG, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.4, { strip: C.SECONDARY });
      slide.addText("Mini calendar mock-up", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 13, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0, align: "center",
      });

      // Mini calendar with sample equations
      const startY = lg.panelTopPadded + 0.45;
      const cellW = 0.55;
      const cellH = 0.42;
      // First row days 1-7
      for (let c = 0; c < 7; c++) {
        const day = c + 1;
        const x = lg.rightX + 0.15 + c * cellW;
        slide.addShape("roundRect", {
          x, y: startY, w: cellW - 0.05, h: cellH - 0.04, rectRadius: 0.04,
          fill: { color: C.WHITE }, line: { color: C.MUTED, width: 0.4 },
        });
        slide.addText(String(day), {
          x, y: startY, w: cellW - 0.05, h: cellH - 0.04,
          fontSize: 9, fontFace: FONT_B, color: C.CHARCOAL, align: "center", valign: "top", margin: 0,
        });
      }

      // Sample equation under day 7
      slide.addText("Sample for date 7:", {
        x: lg.rightX + 0.15, y: startY + 0.55, w: lg.rightW - 0.3, h: 0.3,
        fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, bold: true, margin: 0,
      });
      addTextOnShape(slide, "(5 + 4) - 2 = 7", {
        x: lg.rightX + 0.5, y: startY + 0.9, w: lg.rightW - 1.0, h: 0.5, rectRadius: 0.08,
        fill: { color: C.PRIMARY },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });

      slide.addText("Sample for date 14:", {
        x: lg.rightX + 0.15, y: startY + 1.55, w: lg.rightW - 0.3, h: 0.3,
        fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, bold: true, margin: 0,
      });
      addTextOnShape(slide, "5 + 4 + 3 + 2 = 14", {
        x: lg.rightX + 0.5, y: startY + 1.9, w: lg.rightW - 1.0, h: 0.5, rectRadius: 0.08,
        fill: { color: C.PRIMARY },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 6: The rules
  contentSlide(pres, "I Do", C.ALERT, "The Rules",
    [
      "Only use digits: 2, 3, 4, 5",
      "Each digit can only be used ONCE per equation",
      "Operations allowed: +, -, x, /, powers",
      "Brackets are a TOOL - they don't count as an operation",
      "",
      "14 equations: 3 digits and 2 different operations",
      "14 equations: 3 digits and 3 different operations",
    ],
    NOTES_RULES, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.4, { strip: C.ALERT });
      slide.addText("Quick reminder", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 14, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0, align: "center",
      });

      // Seed digits panel
      slide.addText("Seed digits:", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.45, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, bold: true, margin: 0,
      });
      const digits = ["2", "3", "4", "5"];
      digits.forEach((d, i) => {
        addTextOnShape(slide, d, {
          x: lg.rightX + 0.25 + i * 0.85, y: lg.panelTopPadded + 0.85, w: 0.7, h: 0.7, rectRadius: 0.08,
          fill: { color: C.PRIMARY },
        }, { fontSize: 28, fontFace: FONT_H, color: C.WHITE, bold: true });
      });

      // 14 / 14 split
      addTextOnShape(slide, "14 dates: 2 ops", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 1.85, w: lg.rightW - 0.4, h: 0.45, rectRadius: 0.06,
        fill: { color: C.SECONDARY },
      }, { fontSize: 13, fontFace: FONT_B, color: C.WHITE, bold: true });
      addTextOnShape(slide, "14 dates: 3 ops", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 2.4, w: lg.rightW - 0.4, h: 0.45, rectRadius: 0.06,
        fill: { color: C.SUCCESS },
      }, { fontSize: 13, fontFace: FONT_B, color: C.WHITE, bold: true });
      slide.addText("That's 28 dates total - the whole month!", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 2.95, w: lg.rightW - 0.4, h: 0.4,
        fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, italic: true, align: "center", margin: 0,
      });
    }
  );

  // Slide 7: Worked example
  workedExSlide(pres, 2, "I Do", "Build for Date 5",
    [
      "Target: 5",
      "Need: 3 digits, 2 different operations",
      "",
      "Try 1: 2 + 4 - 3 = 3   No - target is 5",
      "Try 2: 4 + 3 - 2 = 5   YES! 3 digits, 2 ops",
      "Try 3 (bonus): 5 + 4 - 4 = 5   No - reused 4",
      "Try 4 (bonus): (5 - 3) + 2 + 1 = 5   No - no 1 available, also 3 ops",
      "",
      "Final for date 5:  4 + 3 - 2 = 5",
    ],
    NOTES_WORKED, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.4, { strip: C.SUCCESS });
      slide.addText("Date 5 worked", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 14, fontFace: FONT_H, color: C.SUCCESS, bold: true, margin: 0, align: "center",
      });

      addTextOnShape(slide, "4 + 3 - 2", {
        x: lg.rightX + 0.4, y: lg.panelTopPadded + 0.55, w: lg.rightW - 0.8, h: 0.7, rectRadius: 0.1,
        fill: { color: C.PRIMARY },
      }, { fontSize: 28, fontFace: FONT_H, color: C.WHITE, bold: true });

      slide.addText("Check:", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 1.4, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, bold: true, margin: 0,
      });
      slide.addText([
        { text: "3 digits: 4, 3, 2", options: { bullet: true, fontSize: 12, color: C.CHARCOAL, breakLine: true } },
        { text: "2 different ops: + and -", options: { bullet: true, fontSize: 12, color: C.CHARCOAL, breakLine: true } },
        { text: "Equals 5", options: { bullet: true, fontSize: 12, color: C.SUCCESS, bold: true } },
      ], {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 1.75, w: lg.rightW - 0.5, h: 1.2,
        fontFace: FONT_B, margin: 0, valign: "top",
      });
    }
  );

  // Slide 8-9: CFU 1
  withReveal(
    () => cfuSlide(pres, "CFU", "Build for Date 6", "Show Me Boards",
      "Build an equation for date 6.\n\nUse 3 digits from 2, 3, 4, 5 (each once).\nUse 2 different operations.",
      NOTES_CFU1, FOOTER),
    (slide) => {
      addTextOnShape(slide, "Examples: 5 + 3 - 2 = 6,  4 + 5 - 3 = 6,  2 x 5 - 4 = 6", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_CFU1_A);
    }
  );

  // Slide 10: Project Time - Big launch slide
  workedExSlide(pres, 4, "Project Time", "Build Your First Week (Dates 1-7)",
    [
      "Use the DRAFT SHEET to try equations",
      "Once one works, write it on your CALENDAR",
      "",
      "Today: focus on dates 1 to 7",
      "Most should use 2 different operations",
      "",
      "If you're stuck on one date - skip it",
      "Come back later",
      "",
      "30 minutes",
    ],
    NOTES_PROJECT_TIME, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.4, { strip: C.PRIMARY });
      slide.addText("Today's targets", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0, align: "center",
      });

      // Show dates 1-7 visual
      const startY = lg.panelTopPadded + 0.45;
      const cellW = 0.55;
      const cellH = 0.55;
      for (let c = 0; c < 7; c++) {
        const day = c + 1;
        const x = lg.rightX + 0.15 + c * cellW;
        addTextOnShape(slide, String(day), {
          x, y: startY, w: cellW - 0.05, h: cellH - 0.04, rectRadius: 0.06,
          fill: { color: C.PRIMARY },
        }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
      }

      slide.addText("Build one equation for each", {
        x: lg.rightX + 0.2, y: startY + 0.7, w: lg.rightW - 0.4, h: 0.4,
        fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, italic: true, align: "center", margin: 0,
      });

      addTextOnShape(slide, "Use draft sheet first!", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 1.85, w: lg.rightW - 0.4, h: 0.45, rectRadius: 0.06,
        fill: { color: C.ALERT },
      }, { fontSize: 13, fontFace: FONT_B, color: C.WHITE, bold: true });

      slide.addText([
        { text: "Try, check, adjust", options: { bullet: true, fontSize: 12, color: C.CHARCOAL, breakLine: true } },
        { text: "Use the rubric for guidance", options: { bullet: true, fontSize: 12, color: C.CHARCOAL, breakLine: true } },
        { text: "Skip and come back if stuck", options: { bullet: true, fontSize: 12, color: C.CHARCOAL } },
      ], {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 2.4, w: lg.rightW - 0.5, h: 0.9,
        fontFace: FONT_B, margin: 0, valign: "top",
      });
    }
  );

  // Slide 11: Exit Ticket
  exitTicketSlide(pres,
    [
      "How many of dates 1-7 did you complete today?  __ / 7",
      "Write your favourite equation from today below",
      "Tell your teacher: which date was hardest? Why?",
    ],
    NOTES_EXIT, FOOTER);

  // Slide 12: Closing
  closingSlide(pres, {
    reflectionPrompt: "Tell your partner: which date was hardest today? Why?",
    scItems: [
      "I can name the rules of the calendar project",
      "I can build equations for my first week of dates",
      "I can use 3 digits and 2 different operations in my equations",
    ],
    selfAssessment: { prompt: "Self-check", options: ["Got it", "Getting there", "Need more practice"] },
  }, NOTES_CLOSING);

  // Slide 13: Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "BOD_Session7_Calendar_Launch.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // --- PDFs ------------------------------------------------------------------

  // Calendar Template (A3 - landscape recommended)
  await (async () => {
    const doc = createPdf({ title: CALENDAR_RES.name });
    let y = addPdfHeader(doc, CALENDAR_RES.name, {
      subtitle: "BODMAS Calendar - One Equation Per Date",
      color: C.NAVY,
      lessonInfo: "Session 7-10 of 10 | Year 5/6 Numeracy",
    });
    y = addTipBox(doc, "Print on A3 if possible. Top half: decorate with month name, your name, and pictures. Bottom half: the calendar with one equation per date.", y, { color: C.TEAL });

    y = addSectionHeading(doc, "My Month: ___________________________   My Name: ___________________________", y, { color: C.NAVY });
    y += 6;

    // Draw a 7x4 calendar grid
    const startX = 50;
    const cellW = 70;
    const cellH = 70;
    const gridY = y + 5;

    // Day headers
    const dayNames = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"];
    dayNames.forEach((d, i) => {
      doc.rect(startX + i * cellW, gridY, cellW, 18).stroke("#" + C.NAVY);
      doc.fontSize(10).fillColor("#" + C.NAVY)
        .text(d, startX + i * cellW, gridY + 4, { width: cellW, align: "center" });
    });

    // Calendar cells - 28 days in 4 rows
    for (let row = 0; row < 4; row++) {
      for (let col = 0; col < 7; col++) {
        const day = row * 7 + col + 1;
        if (day > 28) continue;
        const cellY = gridY + 18 + row * cellH;
        doc.rect(startX + col * cellW, cellY, cellW, cellH).stroke("#" + C.NAVY);
        doc.fontSize(11).fillColor("#" + C.NAVY).text(String(day), startX + col * cellW + 4, cellY + 4);
      }
    }

    y = gridY + 18 + 4 * cellH + 25;
    y = addBodyText(doc, "Write your equation NEATLY in each box. Make sure you've checked the working before you write here.", y);

    addPdfFooter(doc, "BODMAS Calendar Template | Year 5/6");
    await writePdf(doc, path.join(OUT_DIR, CALENDAR_RES.fileName));
    console.log("PDF written: " + CALENDAR_RES.fileName);
  })();

  // Draft Sheet
  await (async () => {
    const doc = createPdf({ title: DRAFT_RES.name });
    let y = addPdfHeader(doc, DRAFT_RES.name, {
      subtitle: "Draft Your Equations Here First",
      color: C.NAVY,
      lessonInfo: "Session 7-10 of 10 | Year 5/6 Numeracy",
    });
    y = addTipBox(doc, "Use this sheet to try equations. Once one works, write it on your calendar. Working is messy - that's normal.", y, { color: C.TEAL });

    // Two columns of date drafts
    const startX = 50;
    const colW = 250;
    const rowH = 38;

    for (let i = 0; i < 14; i++) {
      const day = i + 1;
      const col = i < 7 ? 0 : 1;
      const row = col === 0 ? i : i - 7;
      const x = startX + col * (colW + 10);
      const cellY = y + row * rowH;

      doc.rect(x, cellY, colW, rowH - 4).stroke("#" + C.NAVY);
      doc.fontSize(10).fillColor("#" + C.NAVY).text("Date " + day + ":", x + 5, cellY + 4);
      doc.fontSize(8).fillColor("#666666").text("Try: __________________________", x + 50, cellY + 4);
      doc.text("Working: __________________________", x + 50, cellY + 16);
    }
    y += 7 * rowH + 10;

    y = addSectionHeading(doc, "Sample workings", y, { color: C.NAVY });
    y = addBodyText(doc, "Date 7:   (5 + 4) - 2 = 9 - 2 = 7   YES   |   Uses 5, 4, 2 - 3 digits, 2 ops", y);
    y = addBodyText(doc, "Date 12:  4 x 3 = 12   YES   |   Uses 4, 3 - 2 digits, 1 op (might not meet rule!)", y);
    y = addBodyText(doc, "Date 12:  5 + 4 + 3 = 12   YES   |   Uses 5, 4, 3 - 3 digits, 2 ops (this works!)", y);

    addPdfFooter(doc, "Draft Sheet | Year 5/6");
    await writePdf(doc, path.join(OUT_DIR, DRAFT_RES.fileName));
    console.log("PDF written: " + DRAFT_RES.fileName);
  })();

  // Rubric
  await (async () => {
    const doc = createPdf({ title: RUBRIC_RES.name });
    let y = addPdfHeader(doc, RUBRIC_RES.name, {
      subtitle: "Calendar Project Rubric",
      color: C.NAVY,
      lessonInfo: "Session 7-10 of 10 | Year 5/6 Numeracy",
    });

    y = addTipBox(doc, "Read this before you start. Use it as a checklist as you go. Use it again at the end to check your calendar.", y, { color: C.TEAL });

    y = addSectionHeading(doc, "I have...", y, { color: C.NAVY });
    y = addBodyText(doc, "[ ] An equation for every date 1 to 28", y);
    y = addBodyText(doc, "[ ] Each equation equals the date", y);
    y = addBodyText(doc, "[ ] Used only the digits 2, 3, 4 and 5", y);
    y = addBodyText(doc, "[ ] Each digit used only once per equation", y);
    y = addBodyText(doc, "[ ] 14 dates with 3 digits and 2 different operations", y);
    y = addBodyText(doc, "[ ] 14 dates with 3 digits and 3 different operations", y);
    y = addBodyText(doc, "[ ] Used brackets in some equations", y);
    y = addBodyText(doc, "[ ] Used at least one power (order)", y);
    y = addBodyText(doc, "[ ] Decorated the top half with month, name, and pictures", y);
    y = addBodyText(doc, "[ ] Written equations neatly on the calendar", y);

    y = addSectionHeading(doc, "Checking your equations", y, { color: C.NAVY });
    y = addBodyText(doc, "For each equation, ask:", y);
    y = addBodyText(doc, "  1. Does it use only 2, 3, 4, 5? (each once)", y);
    y = addBodyText(doc, "  2. Does it equal the date when I solve with BODMAS?", y);
    y = addBodyText(doc, "  3. Does it have the right number of operations?", y);

    y = addSectionHeading(doc, "Going further", y, { color: C.NAVY });
    y = addBodyText(doc, "True Order of Operations: if you solve the equation left-to-right WITHOUT BODMAS, you get a DIFFERENT answer.", y);
    y = addBodyText(doc, "Try to write equations where this is true - your calendar will be more interesting!", y);

    addPdfFooter(doc, "Calendar Project Rubric | Year 5/6");
    await writePdf(doc, path.join(OUT_DIR, RUBRIC_RES.fileName));
    console.log("PDF written: " + RUBRIC_RES.fileName);
  })();

  console.log("Session 7 build complete.");
}

build().catch(err => { console.error(err); process.exit(1); });
