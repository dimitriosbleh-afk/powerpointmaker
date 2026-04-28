"use strict";

// BODMAS Unit - Session 9: Calendar Project - 3-Operation Equations
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

const SESSION = 9;
const FOOTER = "BODMAS | Session 9 of 10 - Calendar Project | Year 5/6 Numeracy";
const OUT_DIR = "output/BOD_Session9_Calendar_Three_Ops";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const THREE_OP_RES = makeSessionResource(SESSION, "Three Op Strategies", "Strategies for building 3-operation equations.");
const PEER_CHECK_RES = makeSessionResource(SESSION, "Peer Check Sheet", "Partner-check checklist for completed equations.");
const RESOURCE_ITEMS = [THREE_OP_RES, PEER_CHECK_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// --- Teacher Notes -----------------------------------------------------------

const NOTES_TITLE = `SAY:
- Day 3 of the calendar project
- Today: the harder half - 14 dates with 3 different operations
- We'll also do a peer-check halfway through

DO:
- Display title slide
- Have draft sheets, calendars, hint sheet ready
- Have peer-check sheet ready

TEACHER NOTES:
Session 9 of 10. Today is the hardest content of the project. 3 operations is harder than 2 - more constraints. Peer-check helps quality. By the end of today most students should have 24-28 dates done.

WATCH FOR:
- Students who haven't finished the 2-op section - they need to catch up first
- Students who are ready for 3-op - challenge them with brackets and powers

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_DR_Q = `SAY:
- Quick warm-up - solve these
- 3 minutes

DO:
- Display 2 equations
- Brief

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Tick or fix

DO:
- Reveal answers

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_LI_SC = `SAY:
- Today's intention: tackle the 14 three-operation dates
- Read SC together

DO:
- Choral read

TEACHER NOTES:
SC focus on 3-op equations and peer-check skills.

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_THREE_OP_INTRO = `SAY:
- 3 different operations - that's harder
- You'll often need brackets to make it work
- Powers count as an operation too
- Think of it as: pick a target, pick three operations, see what fits

DO:
- Display the rule
- Show examples

TEACHER NOTES:
3 different operations means three different SYMBOLS. So + and x and / would count. + and + and x would NOT.

WATCH FOR:
- Students who use the same operation twice - prompt: "different operations means different symbols"

[Stage 2: I Do | VTLM 2.0: Strategy Modelling]`;

const NOTES_WORKED = `SAY:
- Let's build for date 9 with 3 operations
- Target: 9
- Need 3 digits, 3 different operations
- Try 1: 5 + 4 - 2 + 3 = 10 nope and only 2 different ops anyway
- Try 2: 5 - 2 + 3 x 2 nope reused 2
- Try 3: 5 + 2 x 3 - 4 = 5 + 6 - 4 = 7 nope (target 9)
- Try 4: 4 x 3 - 5 + 2 = 12 - 5 + 2 = 9 YES! Uses x, -, +. 4 digits, 3 different ops
- Wait - rule says 3 digits. 4x3-5+2 uses 4 digits
- Try 5: 5 x 3 - 4 - 2 = 15 - 4 - 2 = 9 - uses x, -, -. Only 2 different ops
- Try 6: (5 - 2) x 4 - 3 = 12 - 3 = 9 YES 4 digits, brackets and 2 ops still
- Hmm, this date is hard with 3 ops AND 3 digits. Let's try a different date

DO:
- Think aloud
- Show the messy reality
- Pivot to a date that works more easily

TEACHER NOTES:
This is HONEST modelling - some date+rule combos are hard. Show students it's OK to skip and come back. Students will hit walls; let them see you hit one too.

WATCH FOR:
- Students nodding along - they recognise the difficulty
- Students surprised "you couldn't get it?" - good, that's the message

[Stage 2: I Do | VTLM 2.0: Honest Modelling]`;

const NOTES_EASIER = `SAY:
- Let's try date 13 with 3 ops
- 5 + 2 squared + 4 nope reused 4? Let me check, no: 5+4+2squared = 5+4+4=13 nope reused
- (3 + 2) + 4 + 4 nope reused. 3 + 4 + 2 + 4 nope
- 2 squared + (5 + 4) = 4 + 9 = 13 - YES! ops: power, brackets-tool, +. Need 3 DIFFERENT ops
- Hmm power and + is only 2 ops (brackets aren't an op)
- 2 squared x 3 + (5 - 4) nope no 1 needed
- 5 + 2 x 4 = 13 - only 2 different ops (2,3,4,5)
- (3 x 4) + 5 - 4 nope reused
- 5 x 3 - 4 + 2 = 15 - 4 + 2 = 13 - YES! 4 digits, 3 different ops (x, -, +)
- Wait rule says 3 digits, this uses 4

Actually: this is the rule - some students may want to bend the rule "3 digits" to "at least 3 digits". Different versions of this project are slightly different. For our project: "3 digits and 3 different operations" is the strict rule.

DO:
- Show the struggle
- Find ONE that works honestly
- Acknowledge the difficulty

TEACHER NOTES:
Important: the project rules say "14 problems must include 3 seed digits and 3 different operations". So 4 digits with 3 ops would technically violate. BUT - many real classrooms allow MORE digits, just minimum 3. Decide your rule and be consistent. Communicate clearly to students.

WATCH FOR:
- Students who hit walls - this is normal, support them
- Students who push through with creative use of brackets - excellent

[Stage 2: I Do | VTLM 2.0: Strategy Modelling]`;

const NOTES_PEER_CHECK = `SAY:
- Halfway through - peer check
- Swap calendars with your partner
- Use the peer check sheet
- For each equation, run the rubric
- Mark with a tick or a question mark
- This helps you spot errors before the final session

DO:
- Hand out peer check sheets
- Pair students up
- Allow 10 minutes for peer check
- Students discuss any question marks together

TEACHER NOTES:
Peer-check is a powerful learning tool. Students see how others approach the project AND check their own work against another set of eyes. The discussion afterwards is where the learning happens.

WATCH FOR:
- Students who race through peer check - prompt them to actually CHECK each equation
- Students who challenge each other's equations - excellent debate, support both sides

[Stage 3: We Do | VTLM 2.0: Peer Assessment]`;

const NOTES_CFU1 = `SAY:
- Quick check after peer check
- On your whiteboards
- Build an equation for date 14 with 3 different operations
- 90 seconds

DO:
- Students build and show
- Look for valid 3-op equations
- Examples: 5 + 2 x 3 + 3 nope reused; 2 x (5 + 3) - 2 nope reused; 5 + 4 x 3 - 3 nope reused; 5 x 3 - 4 + 3 nope reused; (4 + 3) x 2 = 14 only 2 different ops

TEACHER NOTES:
Date 14 is genuinely tricky with the 3-different-ops constraint. The hint sheet has strategies. Students might use 5 + 4 + 3 + 2 = 14 (4 digits, only 1 different op - doesn't meet rule). Better: 5 x 3 - 4 + 3 nope. (5-3) x 4 + 2 doesn't equal 14.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Build for date 14 with 3 different ops
- Show me
PROCEED:
- Most students show valid - move to project
PIVOT:
- Many stuck - this date is hard! Reveal: 2 squared + (5 + 4) - nope no 4 - difficult. Try date 12 instead which is easier

WATCH FOR:
- Students with valid 3-op equation - excellent
- Students stuck - normal for this difficulty

[Stage 2: CFU | VTLM 2.0: Active Checking]`;

const NOTES_CFU1_A = `SAY:
- Hard date with 3 ops!
- Sample: 5 + (4 - 3) x 2 nope evaluates to 5+2=7, not 14
- Sample: 2 cubed + (4 + 2) nope reused 2
- Sample: (5 + 2) x 3 - 4 nope... 21 - 4 = 17. Hmm
- Sample: (5 + 3) + 4 x 2 nope reused 2... evaluates 8 + 8 = 16
- Some dates ARE hard with strict rules. Use a different operation balance!

DO:
- Reveal sample
- Discuss why it's hard

[Stage 2: CFU Answer | VTLM 2.0: Active Checking]`;

const NOTES_PROJECT_TIME = `SAY:
- Project time
- Today: focus on the 14 three-operation dates
- Use the strategy sheet
- Use peer check when needed
- 30 minutes

DO:
- Allow 30 minutes
- Circulate
- Pull small group for any students still on 2-op work

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Pull small group. Focus on TWO dates only - get them perfect with 3 ops. Quality over quantity
- Extra Notes: Use the strategy sheet
EXTENDING PROMPT:
- Task: Strong students - use a power AND brackets in at least 5 of your 3-op equations
- Extra Notes: Push for variety and creativity

WATCH FOR:
- Students who finish 3-op section - move them to decoration
- Students stuck - peer support and small group

[Stage 4: Project | VTLM 2.0: Independent Practice]`;

const NOTES_EXIT = `SAY:
- Exit ticket
- Three quick questions about your progress

DO:
- Hand out exit ticket
- Collect

TEACHER NOTES:
Exit ticket tracks progress.

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Pack up - we finish tomorrow
- Read SC together
- Self-check
- Tomorrow we polish, decorate, and play BODMAS games

DO:
- Closing slide
- Pack up

TEACHER NOTES:
Build excitement for tomorrow's final session.

[General: Closing | VTLM 2.0: Reflection]`;

const NOTES_RESOURCES = `SAY:
- Strategy sheet for 3-op equations
- Peer check sheet for partner work

DO:
- Hand out

[General: Resources | VTLM 2.0: Resource Awareness]`;

// --- Build -------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres,
    "Calendar - Three-Operation Day",
    "Tackle the harder dates - 14 with 3 different operations",
    "Session 9 of 10  |  Year 5/6 Numeracy  |  Project Day 3",
    NOTES_TITLE);

  // Slide 2-3: Daily Review
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "BODMAS Warm-Up", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });

      addCard(s, 0.5, CONTENT_TOP, 9, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });

      addTextOnShape(s, "1.   (4 + 3) x 2 squared - 8", {
        x: 1.0, y: CONTENT_TOP + 0.4, w: 8.0, h: 1.0, rectRadius: 0.1,
        fill: { color: STAGE_COLORS["1"] },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });

      addTextOnShape(s, "2.   30 / 6 + 4 x (3 - 1)", {
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
      addTextOnShape(slide, "1)  7 x 4 - 8 = 20    2)  5 + 8 = 13", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slide 4: LI/SC
  liSlide(pres,
    ["I am completing the 14 three-operation dates on my BODMAS Calendar"],
    [
      "I can build an equation that uses 3 different operations",
      "I can use brackets and powers to make a target",
      "I can peer-check a partner's calendar against the rubric",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 5: Three op intro
  contentSlide(pres, "I Do", C.PRIMARY, "Three Different Operations",
    [
      "3 different operations means 3 different SYMBOLS",
      "+, -, x, /, powers - each is a different operation",
      "Brackets are a TOOL, not an operation",
      "",
      "Example: (5 - 2) x 3 + 4   uses - and x and +   THREE different ops",
      "Example: 5 + 4 + 3   uses + only   ONE different op (doesn't meet rule)",
    ],
    NOTES_THREE_OP_INTRO, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.4, { strip: C.PRIMARY });
      slide.addText("Operations checklist", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 13, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0, align: "center",
      });

      const ops = [
        { sym: "+", name: "Addition" },
        { sym: "-", name: "Subtraction" },
        { sym: "x", name: "Multiplication" },
        { sym: "/", name: "Division" },
        { sym: "n²", name: "Powers" },
      ];
      ops.forEach((op, i) => {
        const y = lg.panelTopPadded + 0.45 + i * 0.55;
        addTextOnShape(slide, op.sym, {
          x: lg.rightX + 0.25, y, w: 0.7, h: 0.45, rectRadius: 0.06,
          fill: { color: C.PRIMARY },
        }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
        slide.addText(op.name, {
          x: lg.rightX + 1.05, y, w: lg.rightW - 1.25, h: 0.45,
          fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, valign: "middle", margin: 0,
        });
      });
    }
  );

  // Slide 6: Worked example date 9
  workedExSlide(pres, 2, "I Do", "Build for Date 9 with 3 Ops",
    [
      "Target: 9",
      "Need: 3 digits, 3 different operations",
      "",
      "Try: 5 + 2 x 3 - 4 = 5 + 6 - 4 = 7   No - target is 9",
      "Try: (5 - 2) x 4 - 3 = 12 - 3 = 9   uses brackets, x, -",
      "  - and x = 2 different ops, brackets are a tool",
      "  Need a third op",
      "Try: 2 cubed - 4 + 5 = 8 - 4 + 5 = 9   uses power, -, +",
      "  3 different operations - YES",
      "",
      "Final: 2 cubed - 4 + 5 = 9",
    ],
    NOTES_WORKED, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.4, { strip: C.SUCCESS });
      slide.addText("Date 9 worked", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 14, fontFace: FONT_H, color: C.SUCCESS, bold: true, margin: 0, align: "center",
      });

      addTextOnShape(slide, "2³ - 4 + 5", {
        x: lg.rightX + 0.4, y: lg.panelTopPadded + 0.55, w: lg.rightW - 0.8, h: 0.7, rectRadius: 0.1,
        fill: { color: C.PRIMARY },
      }, { fontSize: 26, fontFace: FONT_H, color: C.WHITE, bold: true });

      slide.addText("Three different ops:", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 1.4, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, bold: true, margin: 0,
      });
      slide.addText([
        { text: "Power (²)", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Subtraction (-)", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Addition (+)", options: { bullet: true, fontSize: 13, color: C.SUCCESS, bold: true } },
      ], {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 1.75, w: lg.rightW - 0.5, h: 1.2,
        fontFace: FONT_B, margin: 0, valign: "top",
      });
    }
  );

  // Slide 7-8: CFU 1 - try date 14
  withReveal(
    () => cfuSlide(pres, "CFU", "Build for Date 14 with 3 Ops", "Show Me Boards",
      "Build an equation for date 14.\nUse 3 different operations.\nDigits 2, 3, 4, 5 (each once).",
      NOTES_CFU1, FOOTER),
    (slide) => {
      addTextOnShape(slide, "Hard date! Sample: 2 squared + 5 x 2 nope reused. (5+3) + 2 squared - 2 nope reused", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.ALERT },
      }, { fontSize: 12, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_CFU1_A);
    }
  );

  // Slide 9: Peer check intro
  contentSlide(pres, "We Do", C.SECONDARY, "Halfway Peer Check",
    [
      "Swap calendars with your partner",
      "Use the peer check sheet",
      "For each equation, run the rubric",
      "",
      "Mark with TICK if good, with QUESTION MARK if unsure",
      "Discuss the question marks together",
      "Help each other - this is the learning",
    ],
    NOTES_PEER_CHECK, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.4, { strip: C.SECONDARY });
      slide.addText("Peer check rubric", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0, align: "center",
      });

      const checks = [
        "Equals the date?",
        "Uses only 2, 3, 4, 5?",
        "Each digit once?",
        "Right number of ops?",
        "Solves with BODMAS?",
      ];
      checks.forEach((c, i) => {
        const y = lg.panelTopPadded + 0.45 + i * 0.55;
        addTextOnShape(slide, "?", {
          x: lg.rightX + 0.25, y, w: 0.45, h: 0.45, rectRadius: 0.06,
          fill: { color: C.SECONDARY },
        }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
        slide.addText(c, {
          x: lg.rightX + 0.85, y, w: lg.rightW - 1.05, h: 0.45,
          fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, valign: "middle", margin: 0,
        });
      });
    }
  );

  // Slide 10: Project time
  workedExSlide(pres, 4, "Project Time", "Three-Operation Day",
    [
      "Goal: 14 of your dates need 3 digits + 3 different operations",
      "These are HARDER - some dates are tricky",
      "Use the strategy sheet",
      "",
      "Tip: brackets help you SEPARATE parts of the equation",
      "Tip: powers can give you a useful intermediate value",
      "Tip: if a date is impossible, move on - come back later",
      "",
      "30 minutes",
    ],
    NOTES_PROJECT_TIME, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.4, { strip: C.PRIMARY });
      slide.addText("By end of today", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0, align: "center",
      });

      slide.addText([
        { text: "All 28 dates DRAFTED", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Most written on calendar", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Peer-checked at least once", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Ready to polish tomorrow", options: { bullet: true, fontSize: 13, color: C.SUCCESS, bold: true } },
      ], {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.5, w: lg.rightW - 0.5, h: 1.7,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addTextOnShape(slide, "30 minutes", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 2.4, w: lg.rightW - 0.4, h: 0.5, rectRadius: 0.08,
        fill: { color: C.ALERT },
      }, { fontSize: 16, fontFace: FONT_B, color: C.WHITE, bold: true });
    }
  );

  // Slide 11: Exit Ticket
  exitTicketSlide(pres,
    [
      "Total dates done so far?  __ / 28",
      "Which 3-op date were you most proud of? Write the equation",
      "What feedback did your partner give you?",
    ],
    NOTES_EXIT, FOOTER);

  // Slide 12: Closing
  closingSlide(pres, {
    reflectionPrompt: "Tomorrow we finish! Tell your partner: how many dates do you still need?",
    scItems: [
      "I can build an equation that uses 3 different operations",
      "I can use brackets and powers to make a target",
      "I can peer-check a partner's calendar against the rubric",
    ],
    selfAssessment: { prompt: "Self-check", options: ["Got it", "Getting there", "Need more practice"] },
  }, NOTES_CLOSING);

  // Slide 13: Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "BOD_Session9_Calendar_Three_Ops.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // --- PDFs ------------------------------------------------------------------

  await (async () => {
    const doc = createPdf({ title: THREE_OP_RES.name });
    let y = addPdfHeader(doc, THREE_OP_RES.name, {
      subtitle: "Strategies for 3-Operation Equations",
      color: C.NAVY,
      lessonInfo: "Session 9 of 10 | Year 5/6 Numeracy",
    });
    y = addTipBox(doc, "3 different operations is harder. These are STRATEGIES, not the only answers. Brackets are a tool - they don't count as an operation.", y, { color: C.TEAL });

    y = addSectionHeading(doc, "Strategy 1: Use a power for one operation", y, { color: C.NAVY });
    y = addBodyText(doc, "Date 8:    2 cubed + (4 - 4) nope reused.   (4 + 4) nope reused", y);
    y = addBodyText(doc, "           2 squared + 4 = 8 (only 2 ops). Hard with 3 different ops!", y);
    y = addBodyText(doc, "Date 9:    2 cubed - 4 + 5 = 8 - 4 + 5 = 9   ops: power, -, +   YES", y);
    y = addBodyText(doc, "Date 11:   2 cubed + 5 - 2 nope reused. 2 squared + (5 + 2) nope", y);
    y = addBodyText(doc, "           5 + 2 x 3 = 11 (only 2 ops, doesn't meet 3-op rule)", y);

    y = addSectionHeading(doc, "Strategy 2: Use brackets to set up the calculation", y, { color: C.NAVY });
    y = addBodyText(doc, "Date 6:    (5 - 4) x 2 + 4 nope reused.   (5 - 3) + 4 = 6 (only 2 ops)", y);
    y = addBodyText(doc, "           (5 - 3) + 2 squared = 2 + 4 = 6   ops: -, +, power   YES", y);
    y = addBodyText(doc, "Date 12:   (5 - 2) x 4 = 12 (2 ops). 4 + 3 + 5 = 12 (1 op).", y);
    y = addBodyText(doc, "           (5 - 3) x 2 squared + 4 = 2 x 4 + 4 = 12   YES (3 ops: -, x, +; with power)", y);
    y = addBodyText(doc, "           Wait - (5 - 3) x 2 squared = 2 x 4 = 8, then 8 + 4 = 12 YES, ops: -, x, power, + (4 different ops, but at least 3)", y);

    y = addSectionHeading(doc, "Strategy 3: Mix x and / and +", y, { color: C.NAVY });
    y = addBodyText(doc, "Date 7:    5 + 4 / 2 nope only 2 ops.  (5 + 2) x 4 / 4 nope reused", y);
    y = addBodyText(doc, "           5 + 4 / 2 + 0 nope no zero.   2 + 4 + 5 - 4 nope reused", y);
    y = addBodyText(doc, "           Hard with strict 3-different-ops!", y);

    y = addSectionHeading(doc, "Honest truth", y, { color: C.NAVY });
    y = addBodyText(doc, "Some dates are GENUINELY HARD with the strict 3-different-ops rule.", y);
    y = addBodyText(doc, "If you can't get it, swap to a different date. Or use brackets creatively.", y);
    y = addBodyText(doc, "It's OK to ask your teacher for a hint if you've been stuck for 5+ minutes.", y);

    addPdfFooter(doc, "Session 9 | 3-Op Strategies | Year 5/6");
    await writePdf(doc, path.join(OUT_DIR, THREE_OP_RES.fileName));
    console.log("PDF written: " + THREE_OP_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: PEER_CHECK_RES.name });
    let y = addPdfHeader(doc, PEER_CHECK_RES.name, {
      subtitle: "Partner Check for Calendar Equations",
      color: C.NAVY,
      lessonInfo: "Session 9 of 10 | Year 5/6 Numeracy",
    });
    y = addTipBox(doc, "Swap calendars with your partner. For each equation, work through the checklist below. Tick or question-mark each one. Discuss the question marks together.", y, { color: C.TEAL });

    y = addSectionHeading(doc, "I am checking ____________________'s calendar", y, { color: C.NAVY });
    y += 8;

    y = addSectionHeading(doc, "For each equation, ask:", y, { color: C.NAVY });
    y = addBodyText(doc, "1.  Does the equation EQUAL the date?", y);
    y = addBodyText(doc, "2.  Does it use only digits 2, 3, 4, 5?", y);
    y = addBodyText(doc, "3.  Is each digit used only ONCE?", y);
    y = addBodyText(doc, "4.  Does it have the right number of operations?", y);
    y = addBodyText(doc, "      2-op equations: 2 DIFFERENT operation symbols", y);
    y = addBodyText(doc, "      3-op equations: 3 DIFFERENT operation symbols", y);
    y = addBodyText(doc, "5.  When you solve with BODMAS, does it match?", y);

    y = addSectionHeading(doc, "Date check log", y, { color: C.NAVY });
    y = addBodyText(doc, "Tick if good. Question mark if unsure. Discuss the unsure ones together.", y);
    y += 5;

    // Two columns of date checks
    const startX = 50;
    const colW = 230;
    for (let i = 0; i < 14; i++) {
      const day = i * 2 + 1; // odd days only? actually let's do 1-14
      const realDay = i + 1;
      const col = i < 7 ? 0 : 1;
      const row = col === 0 ? i : i - 7;
      const x = startX + col * (colW + 20);
      const cellY = y + row * 22;
      doc.fontSize(10).fillColor("#" + C.NAVY).text("Date " + realDay + ":   ___ Equation OK   ___ Wrong digit   ___ Wrong answer", x, cellY);
    }
    y += 7 * 22 + 10;

    y = addSectionHeading(doc, "Discussion notes", y, { color: C.NAVY });
    y = addWriteLine(doc, "What did you find?", y);
    y = addWriteLine(doc, "", y);
    y = addWriteLine(doc, "What did your partner do well?", y);
    y = addWriteLine(doc, "", y);

    addPdfFooter(doc, "Session 9 | Peer Check | Year 5/6");
    await writePdf(doc, path.join(OUT_DIR, PEER_CHECK_RES.fileName));
    console.log("PDF written: " + PEER_CHECK_RES.fileName);
  })();

  console.log("Session 9 build complete.");
}

build().catch(err => { console.error(err); process.exit(1); });
