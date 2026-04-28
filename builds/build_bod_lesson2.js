"use strict";

// BODMAS Unit - Session 2: Brackets and Orders
// Year 5/6 Numeracy, Week 9 Term 2 (variant 2)
// VC2M6A02
// Daily Review: Solving Equations with Multiplication, Division, and Operations
// Number Fluency: Brackets warm-up

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

const SESSION = 2;
const FOOTER = "BODMAS | Session 2 of 10 | Year 5/6 Numeracy";
const OUT_DIR = "output/BOD_Session2_Brackets_And_Orders";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const PRACTICE_RES = makeSessionResource(SESSION, "Practice Sheet", "Brackets and orders practice problems.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Answer Key", "Teacher reference.");
const RESOURCE_ITEMS = [PRACTICE_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// --- Teacher Notes -----------------------------------------------------------

const NOTES_TITLE = `SAY:
- Welcome back to BODMAS
- Yesterday we learned what BODMAS stands for and used it on simple equations
- Today we focus on the first two letters - B and O - Brackets and Orders

DO:
- Display title slide
- Have whiteboards ready
- Have BODMAS reference sheet from yesterday visible

TEACHER NOTES:
Session 2 of 10. Brackets are first in BODMAS for a reason - they let the writer of the equation tell us what to do first. Orders (powers) come second. Both are foundational for the rest of the unit.

WATCH FOR:
- Students who can't remember what each letter stands for - quick choral recap before starting

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_DR_Q = `SAY:
- Daily Review. Solve each one in your book
- These are mixed +, -, x, / from yesterday
- Use BODMAS - which letters appear in each?

DO:
- Display 4 review equations
- Allow 4 minutes
- Circulate

TEACHER NOTES:
Daily Review revisits BODMAS from yesterday with simple two-step equations. This is spaced retrieval - it consolidates yesterday's learning before adding brackets and orders today.

WATCH FOR:
- Students reverting to left-to-right thinking - this is the persistent error
- Students who show clean BODMAS working - they are ready for today's content

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Tick or fix
- Notice: in each one, the multiplication or division had to come first

DO:
- Reveal answers
- Quick tally

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY = `SAY:
- Fluency. Quick mental maths inside the brackets
- Just give me the value of what's INSIDE the brackets
- Choral response

DO:
- Display 6 brackets-only mental maths
- Use choral response - the whole class shouts the answer
- Move briskly

TEACHER NOTES:
Fluency primes the brackets-first habit. Students just compute the inside of each brackets group - this rehearses the mental routine they will use today and tomorrow.

WATCH FOR:
- Students who hesitate - they may not yet have automatic facts
- Strong choral response - the whole class is engaged

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_FLUENCY_A = `SAY:
- Check yourself - all should be confident with these
- These are the kind of small calculations you'll do INSIDE brackets all the time

DO:
- Reveal answers

[Stage 1: Fluency Answers | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Today we are learning about Brackets and Orders
- Read the LI with me
- Read each success criterion together
- Ask: What do you remember about brackets from earlier maths? [Hands down]

DO:
- Choral read
- Brief turn-and-talk on what brackets do

TEACHER NOTES:
Brackets are the writer's tool to control order. Orders (powers) are a compact way to write repeated multiplication. Both come early in BODMAS because they "package" parts of the equation that must be evaluated as a unit.

WATCH FOR:
- Students who think brackets are decorative - they are NOT
- Students who don't yet know what 5 squared means - they will need extra modelling

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_BRACKETS_INTRO = `SAY:
- Brackets are the writer's TOOL
- They tell us: solve THIS part first, no matter what's around it
- Look at this pair: 2 + 3 x 4 versus (2 + 3) x 4
- Without brackets, BODMAS says multiply first: 2 + 12 = 14
- With brackets around 2 + 3, the brackets say "do me first": (2 + 3) = 5, then 5 x 4 = 20
- Same numbers, different brackets, different answer

DO:
- Write both equations on the board side by side
- Box the brackets visibly
- Solve each step by step
- Use a different colour for the bracket part

TEACHER NOTES:
This is the core teaching for brackets. Students need to SEE that brackets change the answer. The two-equation comparison makes the role of brackets concrete - they are not optional decoration, they shift what gets solved first.

MISCONCEPTIONS:
- Misconception: Brackets are optional and don't change the answer
  Why: In simple equations like (5) + 3, brackets do nothing
  Impact: Students ignore brackets in multi-step problems and get the wrong answer
  Quick correction: "Look at (2+3) x 4 versus 2+3 x 4. Different answers. Brackets ALWAYS go first."

WATCH FOR:
- Students who say "but they look the same!" - this is the misconception worth teaching to

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_ORDERS_INTRO = `SAY:
- Now Orders. Orders means powers - little numbers up high
- 5 squared, written 5 with a small 2, means 5 x 5 = 25
- 2 cubed, written 2 with a small 3, means 2 x 2 x 2 = 8
- The little number tells you HOW MANY of the big number to multiply together
- In BODMAS, we evaluate the power BEFORE we do any multiplication or division

DO:
- Write each example on the board
- Say each one out loud: "5 squared is 5 times 5"
- Use a different colour for the small power number

TEACHER NOTES:
Orders is often unfamiliar. Students need to see the expanded form (e.g., 5x5) before they accept the shorthand. After today, they should recognise a power and replace it with its value as the second BODMAS step.

WATCH FOR:
- Students who multiply the base by the power instead of repeating the base (e.g., 5 squared = 10) - common error
- Students who ignore the power and just use the base - they are skipping the O step

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO_COMBINED = `SAY:
- Now let's use both. Watch this equation
- 2 squared x 3 + (4 + 1)
- Step 1 - Brackets: (4 + 1) = 5. Equation now: 2 squared x 3 + 5
- Step 2 - Orders: 2 squared = 4. Equation now: 4 x 3 + 5
- Step 3 - D/M: 4 x 3 = 12. Equation now: 12 + 5
- Step 4 - A/S: 12 + 5 = 17
- Final answer: 17

DO:
- Write each step on the board
- Cross out what you've solved and replace it
- Use the BODMAS reference sheet visibly

TEACHER NOTES:
This worked example combines brackets and orders with a multiplication and addition. Students see all four BODMAS levels in one equation. Walk through each one, even when there's nothing to do at that step (it builds the habit).

WATCH FOR:
- Students who do orders before brackets - re-emphasise the order
- Students who think brackets are "just for show" once they see the answer is bigger - they are still surface-learning

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU1 = `SAY:
- Quick check. On your whiteboards
- (8 - 3) x 2
- Final answer
- 30 seconds. Show me

DO:
- Students compute and show
- Scan for "10"
- 10 means correct (8-3=5, 5x2=10)
- 2 means they did 3x2 first, then 8-6 = 2 (left to right or ignoring brackets)

TEACHER NOTES:
This CFU tests whether students do brackets first. Wrong answers usually mean they ignored the brackets and did the multiplication first.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Final answer only
- Show me
- Scan for: 10
PROCEED:
- 80% or more show 10 - move on
PIVOT:
- Most likely misconception: students ignored brackets and did 3 x 2 first
- Reteach: "Brackets first, ALWAYS. Look at the brackets - they make 8-3 a unit"
- Re-check with: (10 - 4) x 3 [answer: 18]

WATCH FOR:
- Students showing 2 - ignored brackets entirely
- Students showing 10 confidently - bracket habit is forming

[Stage 2: CFU | VTLM 2.0: Active Checking]`;

const NOTES_CFU1_A = `SAY:
- Answer: 10
- 8 minus 3 inside brackets is 5
- Then 5 x 2 = 10

DO:
- Reveal with full working

[Stage 2: CFU Answer | VTLM 2.0: Active Checking]`;

const NOTES_WEDO_Q = `SAY:
- Together now
- 3 squared + (10 - 6) x 2
- What comes first?
- Brackets - do (10 - 6). What is it? [4]
- Now: 3 squared + 4 x 2. What's next?
- Orders - 3 squared = 9. Now: 9 + 4 x 2
- Then D/M: 4 x 2 = 8. Now: 9 + 8
- Then A/S: 9 + 8 = 17

DO:
- Walk through with class input
- Pause for choral response at each step
- Write each step on the board

TEACHER NOTES:
We Do shows all four BODMAS levels firing in one equation. The order matters - brackets, then orders, then x and /, then + and -. Students should be able to call out each step.

WATCH FOR:
- Students who try to do 3 squared before the brackets - reinforce: brackets ALWAYS first
- Students confidently calling "next step is..." - they are getting it

[Stage 3: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_A = `SAY:
- Final answer: 17
- Tick or fix in your books

DO:
- Reveal answer

[Stage 3: We Do Answer | VTLM 2.0: Guided Practice]`;

const NOTES_CFU2 = `SAY:
- Hinge question. Whiteboards
- 4 squared - (5 + 3)
- Final answer
- Use BODMAS

DO:
- Students compute and show
- Scan for "8"
- 8 means correct (4 squared = 16, brackets = 8, 16 - 8 = 8)

TEACHER NOTES:
Hinge tests both Orders and Brackets. The key step is recognising 4 squared = 16 (not 8). Wrong answer of 16 means they ignored the brackets. Wrong answer of 0 means they squared the wrong way.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Final answer
- Show me
- Scan for: 8
PROCEED:
- 80% or more show 8 - go to You Do
PIVOT:
- Most likely misconception: 4 squared = 8 (multiplied 4 by 2 instead of 4 by 4)
- Reteach: "5 squared means 5 multiplied by ITSELF. 4 squared means 4 x 4 = 16. Not 4 x 2."
- Re-check with: 3 squared + (5 - 2) [answer: 12]

WATCH FOR:
- Students showing 0 - they did 4x2=8 then 8-8=0. Both errors
- Students showing 8 confidently - solid

[Stage 3: Hinge CFU | VTLM 2.0: Active Checking]`;

const NOTES_CFU2_A = `SAY:
- Answer: 8
- (5 + 3) = 8 from brackets
- 4 squared = 4 x 4 = 16 from orders
- 16 - 8 = 8

DO:
- Reveal full working

[Stage 3: Hinge Answer | VTLM 2.0: Active Checking]`;

const NOTES_YOUDO = `SAY:
- Independent practice
- Open books
- Solve each one using BODMAS - show every step
- Challenge at the bottom if you finish

DO:
- Hand out practice sheet
- Allow 10 minutes
- Pull a small group for any students who failed the hinge

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Pull small group. Use the BODMAS reference. For each equation, FIRST identify and circle the brackets and any powers
- Extra Notes: Once they can identify B and O, ask them to solve those parts and rewrite the equation
EXTENDING PROMPT:
- Task: Challenge: 3 squared + (10 - 4) x 2 - 5
- Extra Notes: All four BODMAS levels. Answer: 16

WATCH FOR:
- Students who skip the bracket step - reinforce
- Students who confuse 5 squared with 5 x 2

[Stage 4: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_EXIT = `SAY:
- Exit ticket. Three questions
- Show your steps

DO:
- Hand out exit ticket
- Collect

TEACHER NOTES:
Q1 brackets only, Q2 orders only, Q3 both. Tests SC2 - using both B and O steps.

WATCH FOR:
- Students who get Q3 wrong - they need more brackets-AND-orders practice

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Review SC together
- Self-check
- Tell your partner: which letter in BODMAS is hardest for you so far?

DO:
- Display closing slide
- Self-check then turn and talk

TEACHER NOTES:
Closing identifies which step students find hardest - useful for tomorrow's planning.

[General: Closing | VTLM 2.0: Reflection]`;

const NOTES_RESOURCES = `SAY:
- Practice sheet for tonight or for early finishers
- Use your BODMAS reference

DO:
- Hand out practice sheets

[General: Resources | VTLM 2.0: Resource Awareness]`;

// --- Build -------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres,
    "Brackets and Orders",
    "The first two letters of BODMAS",
    "Session 2 of 10  |  Year 5/6 Numeracy",
    NOTES_TITLE);

  // Slide 2-3: Daily Review
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "BODMAS Recap from Yesterday", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });

      addCard(s, 0.5, CONTENT_TOP, 4.5, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
      const dr = [
        "1.  12 + 3 x 4",
        "2.  20 - 8 / 2",
        "3.  6 x 5 - 10",
        "4.  18 / 3 + 7",
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
        { text: "Use BODMAS", options: { fontSize: 16, bold: true, color: C.SECONDARY, breakLine: true } },
        { text: "", options: { fontSize: 6, breakLine: true } },
        { text: "Find the FIRST step", options: { bullet: true, fontSize: 14, color: C.CHARCOAL, breakLine: true } },
        { text: "Show every step", options: { bullet: true, fontSize: 14, color: C.CHARCOAL, breakLine: true } },
        { text: "4 minutes", options: { bullet: true, fontSize: 14, color: C.ALERT, bold: true } },
      ], {
        x: 5.45, y: CONTENT_TOP + 0.15, w: 3.8, h: 1.7,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "1) 24    2) 16    3) 20    4) 13", {
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
      addTitle(s, "What's Inside the Brackets?", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });

      addCard(s, 0.5, CONTENT_TOP, 9, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
      const fluency = [
        "(8 + 4) = ?",
        "(15 - 7) = ?",
        "(6 x 5) = ?",
        "(20 / 4) = ?",
        "(3 + 9) = ?",
        "(10 - 2) = ?",
      ];
      // 3x2 grid
      fluency.forEach((q, i) => {
        const col = i % 3;
        const row = Math.floor(i / 3);
        const x = 0.85 + col * 2.85;
        const y = CONTENT_TOP + 0.4 + row * 1.4;
        addTextOnShape(s, q, {
          x, y, w: 2.6, h: 1.0, rectRadius: 0.1,
          fill: { color: STAGE_COLORS["1"] },
        }, { fontSize: 26, fontFace: FONT_H, color: C.WHITE, bold: true });
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_FLUENCY);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "12  |  8  |  30  |  5  |  12  |  8", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // Slide 6: LI/SC
  liSlide(pres,
    ["I am learning to use Brackets and Orders in BODMAS"],
    [
      "I can solve what is inside brackets first",
      "I can find the value of a power, like 5 squared",
      "I can use Brackets and Orders together in a BODMAS equation",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 7: Brackets intro
  contentSlide(pres, "I Do", C.PRIMARY, "Brackets - the Writer's Tool",
    [
      "Brackets tell us: solve THIS first",
      "Without brackets, BODMAS chooses",
      "With brackets, the writer chooses",
      "",
      "Look at how brackets change the answer:",
    ],
    NOTES_BRACKETS_INTRO, FOOTER,
    (slide, lg) => {
      // Side by side comparison
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 1.6, { strip: C.MUTED });
      slide.addText("No brackets", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 13, fontFace: FONT_B, color: C.MUTED, bold: true, margin: 0,
      });
      slide.addText("2 + 3 x 4", {
        x: lg.rightX, y: lg.panelTopPadded + 0.42, w: lg.rightW, h: 0.5,
        fontSize: 22, fontFace: FONT_H, color: C.CHARCOAL, bold: true, align: "center", margin: 0,
      });
      addTextOnShape(slide, "= 14", {
        x: lg.rightX + 1.0, y: lg.panelTopPadded + 0.95, w: 2.2, h: 0.45, rectRadius: 0.06,
        fill: { color: C.MUTED },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });

      addCard(slide, lg.rightX, lg.panelTopPadded + 1.8, lg.rightW, 1.6, { strip: C.PRIMARY });
      slide.addText("With brackets", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 1.88, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 13, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
      });
      slide.addText("(2 + 3) x 4", {
        x: lg.rightX, y: lg.panelTopPadded + 2.22, w: lg.rightW, h: 0.5,
        fontSize: 22, fontFace: FONT_H, color: C.CHARCOAL, bold: true, align: "center", margin: 0,
      });
      addTextOnShape(slide, "= 20", {
        x: lg.rightX + 1.0, y: lg.panelTopPadded + 2.75, w: 2.2, h: 0.45, rectRadius: 0.06,
        fill: { color: C.PRIMARY },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 8: Orders intro
  contentSlide(pres, "I Do", C.PRIMARY, "Orders - Powers Like 5 Squared",
    [
      "An 'order' means a power",
      "5 squared means 5 x 5 = 25",
      "2 cubed means 2 x 2 x 2 = 8",
      "The little number tells you HOW MANY of the big number",
    ],
    NOTES_ORDERS_INTRO, FOOTER,
    (slide, lg) => {
      // Examples panel
      const examples = [
        { power: "5 squared", expanded: "5 x 5", value: "25" },
        { power: "2 cubed", expanded: "2 x 2 x 2", value: "8" },
        { power: "3 squared", expanded: "3 x 3", value: "9" },
        { power: "10 squared", expanded: "10 x 10", value: "100" },
      ];
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.4, { strip: C.SECONDARY });
      slide.addText("Examples", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0, align: "center",
      });
      examples.forEach((ex, i) => {
        const y = lg.panelTopPadded + 0.46 + i * 0.7;
        slide.addText(ex.power, {
          x: lg.rightX + 0.2, y, w: 1.5, h: 0.4,
          fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, bold: true, valign: "middle", margin: 0,
        });
        slide.addText("=  " + ex.expanded, {
          x: lg.rightX + 1.7, y, w: 1.6, h: 0.4,
          fontSize: 13, fontFace: FONT_B, color: C.MUTED, italic: true, valign: "middle", margin: 0,
        });
        addTextOnShape(slide, ex.value, {
          x: lg.rightX + 3.4, y: y + 0.02, w: 0.7, h: 0.36, rectRadius: 0.05,
          fill: { color: C.SUCCESS },
        }, { fontSize: 13, fontFace: FONT_H, color: C.WHITE, bold: true });
      });
    }
  );

  // Slide 9: I Do - combined
  workedExSlide(pres, 2, "I Do", "Brackets AND Orders Together",
    [
      "Equation: 2 squared x 3 + (4 + 1)",
      "",
      "Step 1 - B (brackets):",
      "  (4 + 1) = 5",
      "  Now: 2 squared x 3 + 5",
      "",
      "Step 2 - O (orders):",
      "  2 squared = 4",
      "  Now: 4 x 3 + 5",
      "",
      "Step 3 - D/M:",
      "  4 x 3 = 12",
      "  Now: 12 + 5",
      "",
      "Step 4 - A/S: 12 + 5 = 17",
    ],
    NOTES_IDO_COMBINED, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.85, { strip: C.PRIMARY });
      slide.addText("BODMAS Order", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 15, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0, align: "center",
      });
      const checks = [
        { L: "B", text: "(4 + 1) = 5",   used: true },
        { L: "O", text: "2 sq = 4",      used: true },
        { L: "D/M", text: "4 x 3 = 12",  used: true },
        { L: "A/S", text: "12 + 5 = 17", used: true },
      ];
      checks.forEach((st, i) => {
        const y = lg.panelTopPadded + 0.45 + i * 0.55;
        addTextOnShape(slide, st.L, {
          x: lg.rightX + 0.25, y, w: 0.7, h: 0.42, rectRadius: 0.06,
          fill: { color: C.SUCCESS },
        }, { fontSize: 13, fontFace: FONT_H, color: C.WHITE, bold: true });
        slide.addText(st.text, {
          x: lg.rightX + 1.05, y, w: lg.rightW - 1.25, h: 0.42,
          fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, valign: "middle", margin: 0,
        });
      });

      addTextOnShape(slide, "ANSWER: 17", {
        x: lg.rightX, y: lg.panelTopPadded + 2.95, w: lg.rightW, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 10-11: CFU 1
  withReveal(
    () => cfuSlide(pres, "CFU", "Brackets First", "Show Me Boards",
      "(8 - 3) x 2\n\nFinal answer on your whiteboard.",
      NOTES_CFU1, FOOTER),
    (slide) => {
      addTextOnShape(slide, "Answer: 10   (8 - 3 = 5, then 5 x 2 = 10)", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_CFU1_A);
    }
  );

  // Slide 12-13: We Do
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Together - All Four Steps",
      [
        "Equation: 3 squared + (10 - 6) x 2",
        "",
        "Call out the next BODMAS step",
        "Brackets first - what's (10 - 6)?",
        "Orders next - what's 3 squared?",
        "Then D/M, then A/S",
        "",
        "Show every step in your books",
      ],
      NOTES_WEDO_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.4, { strip: C.SECONDARY });
        slide.addText("With your partner:", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.1, w: lg.rightW - 0.4, h: 0.36,
          fontSize: 16, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0,
        });
        slide.addText([
          { text: "Brackets first", options: { bullet: true, fontSize: 14, color: C.CHARCOAL, breakLine: true } },
          { text: "Then orders (powers)", options: { bullet: true, fontSize: 14, color: C.CHARCOAL, breakLine: true } },
          { text: "Then division/multiplication", options: { bullet: true, fontSize: 14, color: C.CHARCOAL, breakLine: true } },
          { text: "Then addition/subtraction", options: { bullet: true, fontSize: 14, color: C.CHARCOAL } },
        ], {
          x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.55, w: lg.rightW - 0.5, h: 1.6,
          fontFace: FONT_B, margin: 0, valign: "top",
        });

        addTextOnShape(slide, "Show every step!", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 2.5, w: lg.rightW - 0.4, h: 0.5, rectRadius: 0.08,
          fill: { color: C.ALERT },
        }, { fontSize: 14, fontFace: FONT_B, color: C.WHITE, bold: true });
      }
    ),
    (slide) => {
      addTextOnShape(slide, "(10-6)=4   |   3 squared=9   |   4x2=8   |   9+8=17", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Slide 14-15: Hinge CFU
  withReveal(
    () => cfuSlide(pres, "CFU", "Hinge Question", "Show Me Boards",
      "4 squared - (5 + 3)\n\nFinal answer.",
      NOTES_CFU2, FOOTER),
    (slide) => {
      addTextOnShape(slide, "Answer: 8   (4 squared = 16, brackets = 8, 16 - 8 = 8)", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_CFU2_A);
    }
  );

  // Slide 16: You Do
  workedExSlide(pres, 4, "You Do", "Independent Practice",
    [
      "Solve each using BODMAS - show every step",
      "",
      "1.  (12 - 4) x 3",
      "2.  20 - (3 + 5)",
      "3.  2 squared + 6",
      "4.  3 squared - (4 + 2)",
      "5.  (10 - 6) x 2 squared",
      "",
      "Challenge: 3 squared + (10 - 4) x 2 - 5",
    ],
    NOTES_YOUDO, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.0, { strip: C.ALERT });
      slide.addText("Reminder", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.1, w: lg.rightW - 0.4, h: 0.36,
        fontSize: 15, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0,
      });
      slide.addText([
        { text: "B - brackets first", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "O - orders (powers)", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "D/M - left to right", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "A/S - left to right", options: { bullet: true, fontSize: 13, color: C.CHARCOAL } },
      ], {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.55, w: lg.rightW - 0.5, h: 1.8,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addTextOnShape(slide, "10 minutes", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 2.45, w: lg.rightW - 0.4, h: 0.4, rectRadius: 0.06,
        fill: { color: C.PRIMARY },
      }, { fontSize: 14, fontFace: FONT_B, color: C.WHITE, bold: true });
    }
  );

  // Slide 17: Exit Ticket
  exitTicketSlide(pres,
    [
      "Solve:  (8 + 4) - 3   (show steps)",
      "Solve:  4 squared + 5   (show steps)",
      "Solve:  (10 - 2) - 2 squared   (show steps)",
    ],
    NOTES_EXIT, FOOTER);

  // Slide 18: Closing
  closingSlide(pres, {
    reflectionPrompt: "Tell your partner: which BODMAS letter feels hardest for you so far?",
    scItems: [
      "I can solve what is inside brackets first",
      "I can find the value of a power, like 5 squared",
      "I can use Brackets and Orders together in a BODMAS equation",
    ],
    selfAssessment: { prompt: "Self-check", options: ["Got it", "Getting there", "Need more practice"] },
  }, NOTES_CLOSING);

  // Slide 19: Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "BOD_Session2_Brackets_And_Orders.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // --- PDFs ------------------------------------------------------------------

  await (async () => {
    const doc = createPdf({ title: PRACTICE_RES.name });
    let y = addPdfHeader(doc, PRACTICE_RES.name, {
      subtitle: "Brackets and Orders Practice",
      color: C.NAVY,
      lessonInfo: "Session 2 of 10 | Year 5/6 Numeracy",
    });
    y = addTipBox(doc, "BODMAS reminder: Brackets first. Then Orders (powers). Then D and M (left to right). Then A and S (left to right).", y, { color: C.TEAL });

    y = addSectionHeading(doc, "Section A - Brackets only", y, { color: C.NAVY });
    y = addProblem(doc, 1, "(7 + 3) x 2 =                         Show every step", y, { color: C.NAVY });
    y = addProblem(doc, 2, "20 - (4 + 6) =                        Show every step", y, { color: C.NAVY });
    y = addProblem(doc, 3, "(15 - 5) / 2 =                        Show every step", y, { color: C.NAVY });
    y = addProblem(doc, 4, "(8 + 4) / (3 + 1) =                   Two brackets!", y, { color: C.NAVY });

    y = addSectionHeading(doc, "Section B - Orders (powers) only", y, { color: C.NAVY });
    y = addProblem(doc, 5, "5 squared + 6 =                       Show every step", y, { color: C.NAVY });
    y = addProblem(doc, 6, "20 - 3 squared =                      Show every step", y, { color: C.NAVY });
    y = addProblem(doc, 7, "2 cubed + 4 =                         Show every step", y, { color: C.NAVY });

    y = addSectionHeading(doc, "Section C - Brackets AND Orders", y, { color: C.NAVY });
    y = addProblem(doc, 8, "4 squared + (5 - 2) =                 Show every step", y, { color: C.NAVY });
    y = addProblem(doc, 9, "(10 - 6) x 3 squared =                Show every step", y, { color: C.NAVY });

    y = addSectionHeading(doc, "Challenge", y, { color: C.NAVY });
    y = addProblem(doc, 10, "3 squared + (10 - 4) x 2 - 5 =       Show every step", y, { color: C.NAVY });

    addPdfFooter(doc, "Session 2 | Practice Sheet | Year 5/6");
    await writePdf(doc, path.join(OUT_DIR, PRACTICE_RES.fileName));
    console.log("PDF written: " + PRACTICE_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Brackets and Orders Answer Key",
      color: C.NAVY,
      lessonInfo: "Session 2 of 10 | Year 5/6 Numeracy",
    });

    y = addSectionHeading(doc, "Section A - Brackets only", y, { color: C.NAVY });
    y = addBodyText(doc, "1.  (7 + 3) x 2  -->  10 x 2 = 20", y);
    y = addBodyText(doc, "2.  20 - (4 + 6)  -->  20 - 10 = 10", y);
    y = addBodyText(doc, "3.  (15 - 5) / 2  -->  10 / 2 = 5", y);
    y = addBodyText(doc, "4.  (8 + 4) / (3 + 1)  -->  12 / 4 = 3", y);

    y = addSectionHeading(doc, "Section B - Orders only", y, { color: C.NAVY });
    y = addBodyText(doc, "5.  5 squared + 6  -->  25 + 6 = 31", y);
    y = addBodyText(doc, "6.  20 - 3 squared  -->  20 - 9 = 11", y);
    y = addBodyText(doc, "7.  2 cubed + 4  -->  8 + 4 = 12", y);

    y = addSectionHeading(doc, "Section C - Brackets AND Orders", y, { color: C.NAVY });
    y = addBodyText(doc, "8.  4 squared + (5 - 2)  -->  16 + 3 = 19", y);
    y = addBodyText(doc, "9.  (10 - 6) x 3 squared  -->  4 x 9 = 36", y);

    y = addSectionHeading(doc, "Challenge", y, { color: C.NAVY });
    y = addBodyText(doc, "10.  3 squared + (10 - 4) x 2 - 5", y);
    y = addBodyText(doc, "     B: (10 - 4) = 6   -->  3 squared + 6 x 2 - 5", y);
    y = addBodyText(doc, "     O: 3 squared = 9  -->  9 + 6 x 2 - 5", y);
    y = addBodyText(doc, "     M: 6 x 2 = 12     -->  9 + 12 - 5", y);
    y = addBodyText(doc, "     A/S: 9 + 12 - 5 = 16   ANSWER: 16", y);

    addPdfFooter(doc, "Session 2 | Answer Key | Year 5/6");
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Session 2 build complete.");
}

build().catch(err => { console.error(err); process.exit(1); });
