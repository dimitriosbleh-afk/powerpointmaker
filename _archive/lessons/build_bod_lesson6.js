"use strict";

// BODMAS Unit - Session 6: From Solving to Creating
// Year 5/6 Numeracy, Week 9 Term 2 (variant 2)
// Bridge lesson - prepares for the BODMAS Calendar project

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

const SESSION = 6;
const FOOTER = "BODMAS | Session 6 of 10 | Year 5/6 Numeracy";
const OUT_DIR = "output/BOD_Session6_Creating_Equations";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const PRACTICE_RES = makeSessionResource(SESSION, "Equation Building Sheet", "Build equations to hit a target number using seed digits.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Answer Key", "Sample equations that work for each target.");
const RESOURCE_ITEMS = [PRACTICE_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// --- Teacher Notes -----------------------------------------------------------

const NOTES_TITLE = `SAY:
- Day 6 of BODMAS
- Last week we SOLVED equations
- This week we CREATE them
- Today is the bridge - tomorrow we start our calendar project

DO:
- Display title slide
- Have BODMAS reference visible
- Mention the calendar project briefly to spark interest

TEACHER NOTES:
Session 6 of 10. The shift from solving to creating is significant. Students need to think backwards - given a target, what equation gets there? This is the foundation for the BODMAS calendar project starting tomorrow.

WATCH FOR:
- Students who think "creating is easier than solving" - they will hit walls when constraints are added
- Students who immediately see how to play - they have transferred their solving skills

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_DR_Q = `SAY:
- Daily Review. A full BODMAS warm-up
- Solve - show every step

DO:
- Display 4 mixed equations
- Allow 5 minutes

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Tick or fix
- Quick check

DO:
- Reveal answers

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY = `SAY:
- Fluency. Quick mixed facts
- Whiteboards

DO:
- Display 8 facts
- Brisk pace

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_FLUENCY_A = `SAY:
- Tick or fix

DO:
- Reveal

[Stage 1: Fluency Answers | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Today's intention: we are CREATING equations
- Read SC together
- Ask: how could we make an equation that equals 10? Many ways!

DO:
- Choral read
- Quick brainstorm: ways to make 10

TEACHER NOTES:
Today's threshold concept: the equals sign works in both directions. We've used = to mean "this calculates to that". Today we use it to mean "build something on the left that equals this on the right". The two are mathematically the same but cognitively different.

WATCH FOR:
- Students who only suggest one way to make a target - this is fine, encourage variety

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_BRACKET_PLACE = `SAY:
- Look at this equation: 3 + 4 x 8 - 6 / 2
- Without brackets, BODMAS gives one answer
- But what if I add ONE pair of brackets?
- (3 + 4) x 8 - 6 / 2 - different answer
- 3 + 4 x (8 - 6) / 2 - another different answer
- 3 + 4 x 8 - (6 / 2) - same answer (the brackets don't change anything because BODMAS already does division first)
- Brackets are POWERFUL - the writer chooses what gets done first

DO:
- Display the original equation
- Add brackets in different places one at a time
- Calculate each variant on the board

TEACHER NOTES:
This shows brackets as a TOOL. The writer of the equation has a choice. Students will need this when they create their calendar - sometimes brackets are needed to hit the target, sometimes they aren't.

WATCH FOR:
- Students who notice that some bracket placements DON'T change the answer - excellent, that's important

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_BUILD_INTRO = `SAY:
- Now let's build an equation
- Target: I want my equation to equal 7
- I have these digits to use: 2, 3, 4, 5
- I can use +, -, x, /, brackets, and powers
- Let me think... 5 - 3 + 2 + 3 = 7? Wait, I can only use each digit once
- 4 + 3 = 7. Done. Easiest.
- Or: 5 + 4 - 2 = 7. Uses three digits.
- Or trickier: (5 - 3) x 4 - 2 + 3 = 9. No, that's 9. Let me adjust...
- (5 + 2) - 3 + 3 = 7. Or just 4 + 3 = 7

DO:
- Think out loud at the board
- Show the search-and-test process
- Don't pretend it's easy - model the trial and error

TEACHER NOTES:
The think-aloud is the MAIN teaching here. Students need to see that creating equations is messy - you try, you check, you adjust. This models the productive struggle they will face in the calendar project.

WATCH FOR:
- Students who watch the trial and error and realise this IS the work - good
- Students who think there's a clever trick - explain there's no shortcut, just trying ideas

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_RULES_PEEK = `SAY:
- Tomorrow's calendar has rules
- You'll only use the digits 2, 3, 4, 5
- Each equation must equal the date - the equation for the 7th equals 7
- Some equations need 2 different operations, some need 3
- Today we practise this small-scale - no calendar yet

DO:
- Display the rules preview
- Don't get bogged down in the project details
- Keep it as a peek

TEACHER NOTES:
Brief preview of the calendar rules. Students don't need to memorise them today - we just want them to know what they're working towards. The full project launches in session 7.

[Stage 2: I Do | VTLM 2.0: Establishing Purpose]`;

const NOTES_CFU1 = `SAY:
- Quick check. Whiteboards
- Use ANY operations and the digits 2, 3, 4
- Build an equation that equals 5
- 30 seconds

DO:
- Students write on whiteboards
- Scan for valid equations: 4 + 3 - 2 = 5, 2 + 3 = 5 (uses two digits), 4 - (3 - 2) x 2 nope...
- Lots of correct possibilities

TEACHER NOTES:
This CFU has many right answers. Students need to find ONE that works. Common right answers: 4 + 3 - 2, 2 + 3, 4 - 2 + 3, (4 - 3) x 5 nope, etc.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Build an equation that equals 5 using digits 2, 3, 4
- Check it works
- Show me
PROCEED:
- Most students show a valid equation
PIVOT:
- If many show invalid (use the wrong digits or wrong answer): "Try the easiest one - 2 + 3 = 5"

WATCH FOR:
- Students who use digits not in the list - prompt
- Students who use operations correctly - their equation works

[Stage 2: CFU | VTLM 2.0: Active Checking]`;

const NOTES_CFU1_A = `SAY:
- Lots of right answers!
- 4 + 3 - 2 = 5
- 2 + 3 = 5
- 4 - 2 + 3 = 5
- Many more

DO:
- Share 2 or 3 student answers
- Celebrate variety

[Stage 2: CFU Answer | VTLM 2.0: Active Checking]`;

const NOTES_WEDO_Q = `SAY:
- Together. Build an equation
- Target: 12
- Digits available: 2, 3, 4, 5
- Use AT LEAST 2 different operations
- Turn to your partner - work together

DO:
- Display the challenge
- Allow 3 minutes for partner work
- Circulate
- Then take 3 student answers and check them at the board

TEACHER NOTES:
Partner work builds collaboration. Many right answers - students should produce a few different ones. Examples: 5 x 3 - 4 + 2 - 1 (no, only 4 digits available), 4 x 3 = 12, 2 + (5 x 2) nope only one 2... etc.

WATCH FOR:
- Students who only find ONE answer - prompt them to find another
- Students who find an answer that uses brackets cleverly - share with class

[Stage 3: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_A = `SAY:
- Lots of right answers
- 4 x 3 = 12 (uses two digits, one operation - meets target but only ONE operation)
- 5 + 4 + 3 = 12 (three digits, two operations - good!)
- 2 + 4 + 3 + 3 - no, can only use each digit once
- (5 - 2) x 4 = 12 (clever!)

DO:
- Share 3 student examples
- Celebrate variety

[Stage 3: We Do Answer | VTLM 2.0: Guided Practice]`;

const NOTES_CFU2 = `SAY:
- Hinge question. Whiteboards
- Build an equation that equals 8
- Use digits 2, 3, 4, 5 (you don't have to use all of them)
- Use at least one BRACKETS

DO:
- 60 seconds
- Students show whiteboards
- Look for valid equations using brackets

TEACHER NOTES:
Hinge tests creative use of brackets - the key skill for the calendar. Examples: (5 + 3), (4 - 2) x 4 = 8, 2 x (3 + 1) nope no 1, (3 + 5), etc.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Build an equation that equals 8 using brackets and the digits 2, 3, 4, 5
- Show me
PROCEED:
- Most students produce a valid equation
PIVOT:
- Most likely difficulty: students forget that brackets are required
- Reteach: "Brackets are needed. Try: (5 + 3) = 8. Or (4 + 4) - oh, can't use 4 twice. (3 + 5)"

WATCH FOR:
- Students with creative answers - share with class
- Students stuck - prompt with one example to start

[Stage 3: Hinge CFU | VTLM 2.0: Active Checking]`;

const NOTES_CFU2_A = `SAY:
- Many right answers!
- (5 + 3) = 8
- (4 + 2) + 3 - 1 - no, no 1 available
- (5 - 3) x 4 = 8
- 4 x (5 - 3) = 8

DO:
- Share student answers
- Celebrate creative use of brackets

[Stage 3: Hinge Answer | VTLM 2.0: Active Checking]`;

const NOTES_YOUDO = `SAY:
- Independent practice
- Building equations on the sheet
- Each row gives you a target and rules
- Find AT LEAST one equation that works

DO:
- Hand out equation building sheet
- Allow 12 minutes
- Circulate

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Pull a small group. Start with the simplest target. Use just two digits. Build up gradually
- Extra Notes: Remove the bracket requirement at first. Add it back when they have one valid equation
EXTENDING PROMPT:
- Task: Challenge: build an equation for 14 using ALL of 2, 3, 4, 5 and at least one bracket and one power
- Extra Notes: Hint: 2 squared + (5 + 3) + 2 nope, can only use each digit once... (3 + 4) x 2 = 14. Many answers

WATCH FOR:
- Students stuck - prompt with the simplest possible answer
- Students finding multiple answers - excellent for the calendar project

[Stage 4: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_EXIT = `SAY:
- Three exit ticket questions
- The third one previews tomorrow's calendar project

DO:
- Hand out exit ticket
- Collect

TEACHER NOTES:
Q1 builds equation for 6, Q2 with brackets for 10, Q3 explanation. Tests SC2.

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Tomorrow we start the BODMAS Calendar project
- You'll build a calendar for your birth month
- Each date will have its own equation that equals that date
- Read SC together, self-check, and tell your partner one thing you noticed today

DO:
- Closing slide
- Self-check
- Turn and talk
- Brief excitement-building about tomorrow

TEACHER NOTES:
Set up the calendar project. Students should leave excited about the next session. Use self-check to gauge readiness - if many "Need more practice", a 10-minute warm-up at the start of session 7 is wise.

[General: Closing | VTLM 2.0: Reflection]`;

const NOTES_RESOURCES = `SAY:
- Equation building sheet for tonight or extra practice

DO:
- Hand out sheets

[General: Resources | VTLM 2.0: Resource Awareness]`;

// --- Build -------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres,
    "From Solving to Creating",
    "Build your own BODMAS equations",
    "Session 6 of 10  |  Year 5/6 Numeracy",
    NOTES_TITLE);

  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slide 2-3: Daily Review
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Full BODMAS Warm-Up", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });

      addCard(s, 0.5, CONTENT_TOP, 4.5, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
      const dr = [
        "1.  4 + 3 x 5 - 6",
        "2.  (12 - 4) / 2 + 3",
        "3.  3 squared - (4 + 5)",
        "4.  20 / 4 + 6 - 2",
      ];
      s.addText(dr.map((p, i) => ({
        text: p,
        options: { fontSize: 20, color: C.CHARCOAL, breakLine: i < dr.length - 1, paraSpaceAfter: 12 },
      })), {
        x: 0.75, y: CONTENT_TOP + 0.2, w: 4.1, h: SAFE_BOTTOM - CONTENT_TOP - 0.4,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addCard(s, 5.2, CONTENT_TOP, 4.3, 2.0, { strip: C.SECONDARY });
      s.addText([
        { text: "All four BODMAS levels", options: { fontSize: 16, bold: true, color: C.SECONDARY, breakLine: true } },
        { text: "", options: { fontSize: 6, breakLine: true } },
        { text: "Show every step", options: { bullet: true, fontSize: 14, color: C.CHARCOAL, breakLine: true } },
        { text: "5 minutes", options: { bullet: true, fontSize: 14, color: C.ALERT, bold: true } },
      ], {
        x: 5.45, y: CONTENT_TOP + 0.15, w: 3.8, h: 1.7,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "1) 13    2) 7    3) 0    4) 9", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slide 4-5: Fluency
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Fluency");
      addTitle(s, "Mixed Operations", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });

      addCard(s, 0.5, CONTENT_TOP, 9, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
      const fluency = [
        "7 x 8", "63 / 9", "57 + 38",
        "100 - 67", "12 x 6", "84 / 7",
        "126 + 89", "203 - 76",
      ];
      fluency.forEach((q, i) => {
        const col = i % 4;
        const row = Math.floor(i / 4);
        const x = 0.7 + col * 2.16;
        const y = CONTENT_TOP + 0.4 + row * 1.55;
        addTextOnShape(s, q, {
          x, y, w: 1.95, h: 1.15, rectRadius: 0.1,
          fill: { color: STAGE_COLORS["1"] },
        }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_FLUENCY);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "56  |  7  |  95  |  33  |  72  |  12  |  215  |  127", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // Slide 6: LI/SC
  liSlide(pres,
    ["I am learning to create my own BODMAS equations to hit a target number"],
    [
      "I can build an equation that equals a given target",
      "I can use brackets to change my answer",
      "I can build an equation that uses 2 or 3 different operations",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 7: Bracket placement power
  contentSlide(pres, "I Do", C.PRIMARY, "Brackets - the Writer's Tool",
    [
      "Same numbers - different brackets - different answer",
      "The writer chooses what gets done first",
      "",
      "Try: 3 + 4 x 8 - 6 / 2",
    ],
    NOTES_BRACKET_PLACE, FOOTER,
    (slide, lg) => {
      const variants = [
        { eq: "No brackets", val: "32" },
        { eq: "(3 + 4) x 8", val: "53" },
        { eq: "3 + 4 x (8 - 6) / 2", val: "7" },
        { eq: "(3 + 4) x (8 - 6) / 2", val: "7" },
      ];
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.4, { strip: C.PRIMARY });
      slide.addText("Different brackets, different answers", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 13, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0, align: "center",
      });
      variants.forEach((v, i) => {
        const y = lg.panelTopPadded + 0.5 + i * 0.65;
        slide.addText(v.eq, {
          x: lg.rightX + 0.2, y, w: lg.rightW - 1.2, h: 0.55,
          fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, valign: "middle", margin: 0,
        });
        addTextOnShape(slide, "= " + v.val, {
          x: lg.rightX + lg.rightW - 1.0, y: y + 0.05, w: 0.8, h: 0.45, rectRadius: 0.06,
          fill: { color: C.PRIMARY },
        }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
      });
    }
  );

  // Slide 8: Build an equation - think aloud
  workedExSlide(pres, 2, "I Do", "Build to a Target",
    [
      "Target: 7",
      "Available digits: 2, 3, 4, 5",
      "Each digit can only be used ONCE",
      "",
      "Try 1: 4 + 3 = 7 -- works! Two digits, one operation",
      "Try 2: 5 + 4 - 2 = 7 -- works! Three digits, two operations",
      "Try 3: (5 - 3) + 2 + 3 -- can't reuse 3",
      "Try 4: (5 + 4) - 2 = 7 -- works!",
      "",
      "Many right answers exist",
      "Try, check, adjust",
    ],
    NOTES_BUILD_INTRO, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.4, { strip: C.SUCCESS });
      slide.addText("Examples", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.1, w: lg.rightW - 0.4, h: 0.36,
        fontSize: 16, fontFace: FONT_H, color: C.SUCCESS, bold: true, margin: 0, align: "center",
      });

      const examples = [
        "4 + 3 = 7",
        "5 + 4 - 2 = 7",
        "(5 + 4) - 2 = 7",
        "5 - 3 + 4 + ? - ?",
      ];
      examples.forEach((eq, i) => {
        const y = lg.panelTopPadded + 0.55 + i * 0.55;
        slide.addText(eq, {
          x: lg.rightX + 0.3, y, w: lg.rightW - 0.6, h: 0.45,
          fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL, bold: true, valign: "middle", margin: 0,
        });
      });

      addTextOnShape(slide, "Try, check, adjust", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 2.95, w: lg.rightW - 0.4, h: 0.4, rectRadius: 0.06,
        fill: { color: C.PRIMARY },
      }, { fontSize: 14, fontFace: FONT_B, color: C.WHITE, bold: true });
    }
  );

  // Slide 9: Calendar rules preview
  contentSlide(pres, "Preview", C.SECONDARY, "Tomorrow's Project: BODMAS Calendar",
    [
      "Pick your birth month - 28 days to fill",
      "Each day's equation must EQUAL that date",
      "Day 7 equation = 7. Day 14 equation = 14. And so on.",
      "",
      "You will use the seed digits: 2, 3, 4, 5",
      "Some equations need 2 different operations",
      "Some equations need 3 different operations",
    ],
    NOTES_RULES_PEEK, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.4, { strip: C.SECONDARY });
      slide.addText("Sneak peek", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0, align: "center",
      });

      // Mini calendar mock
      const startY = lg.panelTopPadded + 0.45;
      const cellW = 0.55;
      const cellH = 0.42;
      // 7 columns x 4 rows
      for (let r = 0; r < 4; r++) {
        for (let c = 0; c < 7; c++) {
          const day = r * 7 + c + 1;
          if (day > 28) break;
          const x = lg.rightX + 0.15 + c * cellW;
          const y = startY + r * cellH;
          slide.addShape("roundRect", {
            x, y, w: cellW - 0.05, h: cellH - 0.04, rectRadius: 0.04,
            fill: { color: C.WHITE }, line: { color: C.MUTED, width: 0.4 },
          });
          slide.addText(String(day), {
            x, y, w: cellW - 0.05, h: cellH - 0.04,
            fontSize: 9, fontFace: FONT_B, color: C.CHARCOAL, align: "center", valign: "middle", margin: 0,
          });
        }
      }

      addTextOnShape(slide, "Coming tomorrow!", {
        x: lg.rightX + 0.4, y: lg.panelTopPadded + 2.4, w: lg.rightW - 0.8, h: 0.45, rectRadius: 0.08,
        fill: { color: C.ACCENT },
      }, { fontSize: 14, fontFace: FONT_B, color: C.WHITE, bold: true });
    }
  );

  // Slide 10-11: CFU 1
  withReveal(
    () => cfuSlide(pres, "CFU", "Build an Equation", "Show Me Boards",
      "Use digits 2, 3, 4 (each only once).\n\nBuild an equation that equals 5.",
      NOTES_CFU1, FOOTER),
    (slide) => {
      addTextOnShape(slide, "Many right answers: 2 + 3, 4 + 3 - 2, 4 - 2 + 3 - and more", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_CFU1_A);
    }
  );

  // Slide 12-13: We Do
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Build for Target 12",
      [
        "Target: 12",
        "Digits available: 2, 3, 4, 5 (each once)",
        "Use AT LEAST 2 different operations",
        "",
        "Work with your partner",
        "Find ONE equation that works",
        "Check by solving with BODMAS",
        "",
        "Bonus: find a SECOND different equation",
      ],
      NOTES_WEDO_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.4, { strip: C.SECONDARY });
        slide.addText("Strategy", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.1, w: lg.rightW - 0.4, h: 0.36,
          fontSize: 16, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0, align: "center",
        });
        slide.addText([
          { text: "Pick an operation to start", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
          { text: "Pick digits to combine", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
          { text: "Solve to check it equals 12", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
          { text: "Adjust digits or operations", options: { bullet: true, fontSize: 13, color: C.CHARCOAL } },
        ], {
          x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.55, w: lg.rightW - 0.5, h: 1.6,
          fontFace: FONT_B, margin: 0, valign: "top",
        });

        addTextOnShape(slide, "Many answers exist!", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 2.6, w: lg.rightW - 0.4, h: 0.5, rectRadius: 0.08,
          fill: { color: C.ALERT },
        }, { fontSize: 13, fontFace: FONT_B, color: C.WHITE, bold: true });
      }
    ),
    (slide) => {
      addTextOnShape(slide, "Examples: 5 + 4 + 3 = 12,  (5 - 2) x 4 = 12,  4 x 3 = 12,  2 x (3 + 4) - 2 nope", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 11, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Slide 14-15: Hinge CFU
  withReveal(
    () => cfuSlide(pres, "CFU", "Use Brackets", "Show Me Boards",
      "Build an equation that equals 8.\nDigits: 2, 3, 4, 5 (each once).\nUse at least one PAIR of brackets.",
      NOTES_CFU2, FOOTER),
    (slide) => {
      addTextOnShape(slide, "Examples: (5 + 3) = 8,  (5 - 3) x 4 = 8,  4 x (5 - 3) = 8", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_CFU2_A);
    }
  );

  // Slide 16: You Do
  workedExSlide(pres, 4, "You Do", "Equation Builder",
    [
      "Use the equation builder sheet",
      "Each row gives you a target and rules",
      "Find at least ONE equation per row",
      "Bonus: find TWO different equations",
      "",
      "Remember:",
      "  Each digit used only once per equation",
      "  Solve with BODMAS to check",
      "  Brackets count as a tool, not an operation",
      "",
      "Challenge: build for 14 with brackets AND a power",
    ],
    NOTES_YOUDO, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.0, { strip: C.ALERT });
      slide.addText("Builder rules", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.1, w: lg.rightW - 0.4, h: 0.36,
        fontSize: 15, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0,
      });
      slide.addText([
        { text: "Try, check, adjust", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Each digit once per equation", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Solve with BODMAS to verify", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Multiple answers welcome", options: { bullet: true, fontSize: 13, color: C.SUCCESS, bold: true } },
      ], {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.55, w: lg.rightW - 0.5, h: 1.8,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addTextOnShape(slide, "12 minutes", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 2.45, w: lg.rightW - 0.4, h: 0.4, rectRadius: 0.06,
        fill: { color: C.PRIMARY },
      }, { fontSize: 14, fontFace: FONT_B, color: C.WHITE, bold: true });
    }
  );

  // Slide 17: Exit Ticket
  exitTicketSlide(pres,
    [
      "Build an equation that equals 6 using digits 2, 3, 4 (each once)",
      "Build an equation that equals 10 using brackets and digits 2, 3, 4, 5",
      "Tell your teacher: which is harder - solving or creating? Why?",
    ],
    NOTES_EXIT, FOOTER);

  // Slide 18: Closing
  closingSlide(pres, {
    reflectionPrompt: "Tomorrow we start the BODMAS Calendar! Tell your partner one thing you noticed today.",
    scItems: [
      "I can build an equation that equals a given target",
      "I can use brackets to change my answer",
      "I can build an equation that uses 2 or 3 different operations",
    ],
    selfAssessment: { prompt: "Self-check", options: ["Got it", "Getting there", "Need more practice"] },
  }, NOTES_CLOSING);

  // Slide 19: Resources


  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "BOD_Session6_Creating_Equations.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // --- PDFs ------------------------------------------------------------------

  await (async () => {
    const doc = createPdf({ title: PRACTICE_RES.name });
    let y = addPdfHeader(doc, PRACTICE_RES.name, {
      subtitle: "Build Equations to Hit a Target",
      color: C.NAVY,
      lessonInfo: "Session 6 of 10 | Year 5/6 Numeracy",
    });
    y = addTipBox(doc, "Use ONLY the seed digits 2, 3, 4, 5. Each digit can be used only once per equation. Show your equation, then prove it with BODMAS.", y, { color: C.TEAL });

    y = addSectionHeading(doc, "Section A - One operation only", y, { color: C.NAVY });
    y = addBodyText(doc, "1.  Target = 5    Equation: __________________________   Working: __________", y);
    y = addBodyText(doc, "2.  Target = 7    Equation: __________________________   Working: __________", y);
    y = addBodyText(doc, "3.  Target = 12   Equation: __________________________   Working: __________", y);

    y = addSectionHeading(doc, "Section B - Two different operations", y, { color: C.NAVY });
    y = addBodyText(doc, "4.  Target = 6    Equation: __________________________   Working: __________", y);
    y = addBodyText(doc, "5.  Target = 9    Equation: __________________________   Working: __________", y);
    y = addBodyText(doc, "6.  Target = 14   Equation: __________________________   Working: __________", y);

    y = addSectionHeading(doc, "Section C - Use brackets", y, { color: C.NAVY });
    y = addBodyText(doc, "7.  Target = 8    Equation: __________________________   Working: __________", y);
    y = addBodyText(doc, "8.  Target = 11   Equation: __________________________   Working: __________", y);
    y = addBodyText(doc, "9.  Target = 18   Equation: __________________________   Working: __________", y);

    y = addSectionHeading(doc, "Challenge - 3 operations and a power", y, { color: C.NAVY });
    y = addBodyText(doc, "10. Target = 14   Use brackets AND a power.   Equation: __________________________", y);
    y += 8;
    y = addWriteLine(doc, "Working: ", y);
    y = addWriteLine(doc, "", y);

    addPdfFooter(doc, "Session 6 | Equation Building | Year 5/6");
    await writePdf(doc, path.join(OUT_DIR, PRACTICE_RES.fileName));
    console.log("PDF written: " + PRACTICE_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Sample Equations - Many Right Answers",
      color: C.NAVY,
      lessonInfo: "Session 6 of 10 | Year 5/6 Numeracy",
    });

    y = addBodyText(doc, "These are SAMPLE answers - many other valid equations exist for each target.", y);
    y = addBodyText(doc, "Accept any equation that uses only the digits 2, 3, 4, 5 (each once) and equals the target.", y);

    y = addSectionHeading(doc, "Section A - One operation", y, { color: C.NAVY });
    y = addBodyText(doc, "1.  Target = 5     2 + 3 = 5", y);
    y = addBodyText(doc, "2.  Target = 7     3 + 4 = 7   OR   2 x 4 - 3 nope two operations", y);
    y = addBodyText(doc, "3.  Target = 12    4 x 3 = 12", y);

    y = addSectionHeading(doc, "Section B - Two operations", y, { color: C.NAVY });
    y = addBodyText(doc, "4.  Target = 6     5 + 3 - 2 = 6", y);
    y = addBodyText(doc, "5.  Target = 9     2 + 4 + 3 = 9", y);
    y = addBodyText(doc, "6.  Target = 14    5 + 4 + 3 + 2 = 14", y);

    y = addSectionHeading(doc, "Section C - Brackets", y, { color: C.NAVY });
    y = addBodyText(doc, "7.  Target = 8     (5 + 3) = 8   OR   (5 - 3) x 4 = 8", y);
    y = addBodyText(doc, "8.  Target = 11    (5 + 4) + 2 = 11", y);
    y = addBodyText(doc, "9.  Target = 18    (5 + 4) x 2 = 18   OR   (4 + 2) x 3 = 18", y);

    y = addSectionHeading(doc, "Challenge", y, { color: C.NAVY });
    y = addBodyText(doc, "10.  Target = 14, brackets AND a power", y);
    y = addBodyText(doc, "      2 squared + (5 + 3) + 2 nope (only one 2)", y);
    y = addBodyText(doc, "      (3 + 4) x 2 = 14, but no power yet... add: 2 squared + (5 - 1) nope no 1", y);
    y = addBodyText(doc, "      Sample: 3 squared + (5 - 4) x 2 + 3 nope reused 3", y);
    y = addBodyText(doc, "      Best fit: 3 squared + 5 = 14   (uses one digit, one power - and one operation)", y);
    y = addBodyText(doc, "      Or: 2 cubed + (4 + 2) nope reused 2 - hard! Many right answers exist; accept any that work.", y);

    addPdfFooter(doc, "Session 6 | Answer Key | Year 5/6");
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Session 6 build complete.");
}

build().catch(err => { console.error(err); process.exit(1); });
