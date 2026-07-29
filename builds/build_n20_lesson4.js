"use strict";

// Numbers to 20 Unit — Session 4: Ordering Numbers to 20
// Term 2 Week 1, Foundation Numeracy, Variant 0
// DR: Data Collection — I can sort and classify objects into categories
// Fluency: 1-1 correspondence (touch and count)
// VC2MFN01 — name, represent and order numbers including zero to at least 20

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addProblem, addWriteLine,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
  formatSessionResourceFileName,
} = require("../themes/pdf_helpers");

const T = createTheme("numeracy", "foundation", 0);
const {
  C, FONT_H, FONT_B,
  SLIDE_W, SLIDE_H, SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  workedExSlide, exitTicketSlide, addStageBadge,
  addTextOnShape, addCard, addFooter, addTopBar, addBadge, addTitle,
  withReveal, runSlideDiagnostics, addNumberLine,
  STAGE_COLORS,
} = T;

const SESSION = 4;
const UNIT_TITLE = "Numbers to 20";
const FOOTER = "Numbers to 20 | Session 4 of 5 | Foundation Numeracy";
const OUT_DIR = "output/N20_Session4_Ordering_Numbers";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Worksheet", "Independent practice - ordering numbers to 20.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Answer Key", "Teacher reference with answers.");
const ENABLING_RES = makeSessionResource(SESSION, "Enabling Scaffold", "Number line with visual supports for ordering to 10.");
const EXTENDING_RES = makeSessionResource(SESSION, "Extension", "Self-contained investigation into ordering and finding missing numbers beyond 20.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, ENABLING_RES, EXTENDING_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Teacher Notes ───────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Today we are going to learn how to put numbers in order
- When we put numbers in order, we go from smallest to biggest
- By the end you will be able to look at some numbers and put them in the right order

DO:
- Display title slide
- Have number cards 0-20, whiteboards, and a large floor number line ready

TEACHER NOTES:
Session 4 of 5. This lesson introduces ordering using number lines and comparison language (bigger, smaller, before, after). Students have built and represented numbers to 20 in Sessions 1-3. Now they learn to sequence and compare.

WATCH FOR:
- Students who are still building confidence with teen numbers -- the ordering activities provide natural revision of number recognition

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_DR_Q = `SAY:
- Warm up! We are sorting today
- Look at the pictures on the screen. Sort them into two groups
- Group 1: Things that are round. Group 2: Things that are NOT round
- On your whiteboard, write the letter of each thing in the right group

DO:
- Display 8 simple objects labelled A-H: ball (A), book (B), coin (C), box (D), apple (E), pencil (F), plate (G), ruler (H)
- Students sort into Round / Not Round on whiteboards
- Allow 90 seconds

TEACHER NOTES:
Daily Review targets data collection -- sorting and classifying. This is a different strand from the main lesson (number) but retrieves skills from the data curriculum. Students practise applying a sorting rule to classify objects.

WATCH FOR:
- Students who disagree about borderline items (apple is roughly round) -- this is good mathematical thinking about classification criteria
- Students who sort quickly and correctly -- classification skills are strong

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Check your sorting
- Round: A (ball), C (coin), E (apple), G (plate)
- Not round: B (book), D (box), F (pencil), H (ruler)
- Ask: Was the apple tricky? It is sort of round but not perfectly round. Good thinking!

DO:
- Reveal the correct sorting
- Discuss the apple as a borderline case -- good for mathematical reasoning

TEACHER NOTES:
Tick-and-fix. The borderline case (apple) teaches that classification sometimes requires judgement -- an important data skill.

WATCH FOR:
- Students who sorted the apple differently and can justify it -- strong reasoning
- Students who made random errors -- they may need more sorting practice

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY = `SAY:
- Fluency time! Touch and count
- I will show you groups of objects on the screen
- Touch each one on the screen with your eyes and count in your head
- Write how many on your whiteboard

DO:
- Display 6 collections (scattered arrangements, not lined up): 8, 5, 12, 7, 10, 15
- Allow 15 seconds per collection
- Students write the count on their whiteboard
- Self-check against revealed answers

TEACHER NOTES:
1-1 correspondence fluency with scattered arrangements (not lined up) is harder than counting objects in a row. It requires students to track which objects they have counted. This revisits Session 1 skills at a slightly higher level.

WATCH FOR:
- Students who struggle with scattered arrangements -- they may benefit from the crossing-out strategy (mark each one as counted)
- Students who count larger collections (12, 15) accurately -- strong counting skills

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Read from slide: "We are learning to put numbers in order from smallest to biggest"
- Our success criteria:
- Read from slide: "I can say which number is bigger"
- Read from slide: "I can put three numbers in order"
- Read from slide: "I can find a missing number in a sequence"

DO:
- Choral read LI and each SC
- Hold up number cards: "We are going to use these to practise ordering"

TEACHER NOTES:
SC1 is comparison (ultra-achievable -- students can compare two numbers). SC2 is three-number ordering (core target). SC3 is finding gaps in a sequence (depth). The number line is the key visual tool for ordering.

WATCH FOR:
- Students who already know bigger/smaller language from everyday life -- build on this

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO = `SAY:
- Watch me put numbers in order
- I have three number cards: 7, 3, 9
- Which is the smallest? I think about counting. When I count, 3 comes first. So 3 is the smallest
- Next is 7. It comes after 3 when I count
- Last is 9. It comes after 7
- In order from smallest to biggest: 3, 7, 9
- I can check on the number line: 3 is on the left, 7 is in the middle, 9 is on the right
- Numbers get BIGGER as we go to the right on the number line

DO:
- Use physical number cards under the document camera
- Place them out of order, then rearrange
- Point to a number line on the wall to verify
- Emphasise the left-to-right = small-to-big rule

TEACHER NOTES:
The number line is the key tool. "Left = smaller, right = bigger" is the ordering principle. The think-aloud models the decision process: which comes first when I count? Students who understand the counting sequence can use it to order.

MISCONCEPTIONS:
- Misconception: Bigger number = physically bigger card or object
  Why: Young children may associate "bigger number" with physical size rather than quantity
  Impact: Students may order by digit appearance (9 looks smaller than 12 because 9 is one digit)
  Quick correction: "Bigger number means MORE. 12 is a bigger number than 9 because 12 is more than 9. We can check by counting."

WATCH FOR:
- Students who can identify the smallest number quickly -- they understand the counting sequence well
- Students who seem unsure about "bigger" vs "smaller" -- use manipulatives: show 3 cubes vs 9 cubes

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO2 = `SAY:
- Now with teen numbers. I have: 15, 11, 18
- Which is smallest? When I count, 11 comes before 15 and 18. So 11 is first
- Next is 15 -- it comes after 11 but before 18
- Last is 18 -- the biggest
- In order: 11, 15, 18
- Check the number line -- yes! 11 is left, 15 is middle, 18 is right

DO:
- Rearrange physical number cards
- Trace along the number line to verify order

TEACHER NOTES:
Second worked example uses teen numbers. The same strategy applies (which comes first when counting?) but students need to know the teen number sequence.

WATCH FOR:
- Students who hesitate with teen number ordering -- they may still be building the teen sequence
- Students who say "11 is first because it starts with 1" -- this is a useful strategy that works for teens

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU1 = `SAY:
- Quick check! Which number is bigger: 8 or 14?
- Think about the number line. Which one is further to the right?
- Show me: hold up your LEFT hand for 8, RIGHT hand for 14

DO:
- Students respond with hand signals
- Scan for right hands (14)

CFU CHECKPOINT:
Technique: Hand Signals
Script:
- Say: "Which is bigger, 8 or 14? Left hand for 8, right hand for 14. Think... show me!"
- Scan for: right hands (14 is bigger). Students who respond quickly understand comparison.
PROCEED: If 80%+ show right hand, move to We Do.
PIVOT: Most likely misconception -- students choose 8 because it is a single digit and looks "normal" while 14 has two digits and looks confusing. Reteach: "Let's count: 1, 2, 3... 8. Keep going: 9, 10, 11, 12, 13, 14. 14 comes AFTER 8, so 14 is bigger. More counting = bigger number."

TEACHER NOTES:
This hinge crosses the single-digit to teen-number boundary, which is the most common source of ordering errors for Foundation students.

WATCH FOR:
- Students who respond with right hand instantly -- comparison across the decade boundary is secure
- Students who hesitate -- they may need the counting strategy to compare

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn! With your partner, put these numbers in order from smallest to biggest
- Use your number cards if you need them
- Set 1: 5, 2, 8
- Set 2: 12, 7, 16
- Write the order on your whiteboard

DO:
- Partners work together with number cards
- Circulate: look for students using the counting sequence or number line to decide order
- Allow 2 minutes for both sets
- Cold call pairs to share their order

TEACHER NOTES:
Set 1 uses single digits only (accessible). Set 2 crosses into teens (harder). Partners can use physical number cards to place in order before writing on the whiteboard.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the Session 4 Enabling Scaffold with a number line 0-10 printed on the page. Students order sets of two numbers only (instead of three) by placing them on the number line. Numbers stay within 0-10.
- Extra Notes: Distribute the Session 4 Enabling Scaffold PDF.
EXTENDING PROMPT:
- Task: Order sets of FOUR numbers including teens, and find missing numbers in sequences (e.g. 3, ___, 5 or 14, ___, 16). The extension PDF teaches the "counting up" and "counting back" strategies.
- Extra Notes: Distribute the Session 4 Extension PDF. Self-contained.

WATCH FOR:
- Students who order Set 1 correctly but struggle with Set 2 -- the teen number sequence is the barrier, not the ordering skill
- Students who use the number line to verify -- strong strategy use
- Readiness signal: pairs completing both sets correctly within 90 seconds

[Stage 3: We Do | VTLM 2.0: Scaffold Practice]`;

const NOTES_WEDO_A = `SAY:
- Let's check. Set 1: smallest to biggest?
- 2, 5, 8. Did everyone get that?
- Set 2: 7, 12, 16. The 7 is smallest because it comes first when we count
- Ask: How did you know 12 comes before 16? [Because 12 comes first when counting]

DO:
- Reveal answers
- Trace the number line to show the order
- Celebrate strategies students used

TEACHER NOTES:
Draw out the reasoning: "I know because when I count..." This is the key strategy for ordering.

WATCH FOR:
- Students who got both sets correct -- ordering skills are developing well
- Students who swapped 7 and 12 -- the single-digit to teen boundary is still tricky

[Stage 3: We Do Answers | VTLM 2.0: Scaffold Practice]`;

const NOTES_CFU2 = `SAY:
- Hinge check! What number is missing?
- 10, 11, ___, 13, 14
- Write the missing number on your whiteboard. Show me!

DO:
- Students write on whiteboards
- Scan for 12

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: "What number is missing? 10, 11, blank, 13, 14. Write it. Ready... show me!"
- Scan for: 12. Students who write 12 quickly can identify gaps in the counting sequence.
PROCEED: If 80%+ show 12, move to You Do.
PIVOT: Most likely misconception -- students write a random teen number because they cannot hold the sequence in working memory. Reteach: "Let's count together slowly: 10, 11... what comes next? 12! Then 13, 14. The missing number is 12."

TEACHER NOTES:
Finding missing numbers is the hardest SC (SC3). It requires holding the sequence in memory and identifying the gap.

WATCH FOR:
- Students who write 12 instantly -- sequence knowledge is automated
- Students who need to count from 10 to find the gap -- this is valid and shows the strategy is working

[Stage 3: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Independent practice time. Open your worksheet
- First: Put each set of numbers in order (smallest to biggest)
- Next: Find the missing numbers in each sequence
- Then: Try the challenge at the bottom!
- You have 6 minutes

DO:
- Distribute Session 4 Worksheet
- Students work independently
- Circulate: prioritise students who struggled with teen ordering
- For enabling students, distribute the Session 4 Enabling Scaffold
- For extending students, distribute the Session 4 Extension

TEACHER NOTES:
Worksheet has three sections: ordering sets of 3, finding missing numbers in sequences, and a challenge (ordering sets of 4). Content is different from the We Do (different numbers).

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the Session 4 Enabling Scaffold with number line support and pairs of numbers to compare (not sets of 3). Numbers are 0-10 only.
- Extra Notes: Distribute the Session 4 Enabling Scaffold PDF.
EXTENDING PROMPT:
- Task: Investigate ordering beyond 20 and finding multiple missing numbers. The extension includes counting sequences with two gaps and numbers up to 30.
- Extra Notes: Distribute the Session 4 Extension PDF. Self-contained.

WATCH FOR:
- Students who order single-digit sets easily but slow down on teen sets -- this is normal progression
- Students who use a number line from the wall to help -- affirm this strategy
- Readiness signal: 80%+ completing Section A correctly within 3 minutes

[Stage 4: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket time. Three quick questions
- Work on your own. You have 2 minutes

DO:
- Display exit ticket on screen
- Students write answers on whiteboards

TEACHER NOTES:
Q1 assesses SC1 (comparison), Q2 assesses SC2 (ordering 3 numbers), Q3 assesses SC3 (missing number).

WATCH FOR:
- Students who get Q1-Q2 correct but struggle with Q3 -- comparison and ordering are secure, sequence gaps need more work
- Students who complete all three correctly -- ready for the consolidation lesson in Session 5

[Stage 5: Exit Ticket | VTLM 2.0: Monitor Progress]`;

const NOTES_CLOSING = `SAY:
- Let's check our success criteria
- SC1: I can say which number is bigger. Thumbs?
- SC2: I can put three numbers in order. Thumbs?
- SC3: I can find a missing number in a sequence. Thumbs?
- Turn and talk: How do you decide which number is bigger? Tell your partner your strategy

DO:
- Display success criteria
- Run thumbs check for each SC
- Allow 30 seconds for Turn and Talk
- Cold call 2 students: "What is your strategy?"

TEACHER NOTES:
Key strategies to listen for: "I think about counting -- which one comes first" or "I look at the number line." Both are valid. The number line strategy is more visual and may be more accessible for some Foundation students.

WATCH FOR:
- Students who articulate a strategy -- metacognition is developing
- Students who show thumbs down on SC3 -- they will get more practice in Session 5

[General: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- These are the resources for today's lesson

DO:
- Point out each resource

TEACHER NOTES:
Resource slide provides clickable links to all companion PDFs for this session.

WATCH FOR:
- N/A -- resource slide is teacher reference only

[General: Resources | VTLM 2.0: Planning]`;

// ─── Build ───────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Session 4: Ordering Numbers to 20", "Foundation Numeracy | Session 4 of 5 | Term 2 Week 1", NOTES_TITLE);

  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slide 2-3: Daily Review (withReveal) — Sorting
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Sort It! Round or Not Round?", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });

      // Objects to sort
      const items = [
        { label: "A. Ball", round: true },
        { label: "B. Book", round: false },
        { label: "C. Coin", round: true },
        { label: "D. Box", round: false },
        { label: "E. Apple", round: true },
        { label: "F. Pencil", round: false },
        { label: "G. Plate", round: true },
        { label: "H. Ruler", round: false },
      ];

      addCard(s, 0.5, CONTENT_TOP, 5.0, 2.2, { strip: STAGE_COLORS["1"] });
      s.addText("Objects:", {
        x: 0.75, y: CONTENT_TOP + 0.08, w: 4.5, h: 0.3,
        fontSize: 15, fontFace: FONT_H, color: STAGE_COLORS["1"], bold: true, margin: 0,
      });
      items.forEach((item, i) => {
        const col = Math.floor(i / 4);
        const row = i % 4;
        s.addText(item.label, {
          x: 0.85 + col * 2.3, y: CONTENT_TOP + 0.45 + row * 0.4, w: 2.0, h: 0.35,
          fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, margin: 0, valign: "middle",
        });
      });

      // Sort table
      addCard(s, 5.8, CONTENT_TOP, 3.7, 2.2, { strip: C.SECONDARY });
      s.addText("Round", {
        x: 5.95, y: CONTENT_TOP + 0.08, w: 1.6, h: 0.3,
        fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0, align: "center",
      });
      s.addText("Not Round", {
        x: 7.7, y: CONTENT_TOP + 0.08, w: 1.6, h: 0.3,
        fontSize: 14, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0, align: "center",
      });
      // Divider line
      s.addShape("line", {
        x: 7.6, y: CONTENT_TOP + 0.4, w: 0, h: 1.6,
        line: { color: C.MUTED, width: 1 },
      });
      // Write lines
      for (let i = 0; i < 4; i++) {
        s.addShape("line", { x: 6.0, y: CONTENT_TOP + 0.55 + i * 0.4, w: 1.4, h: 0, line: { color: C.MUTED, width: 0.8 } });
        s.addShape("line", { x: 7.8, y: CONTENT_TOP + 0.55 + i * 0.4, w: 1.4, h: 0, line: { color: C.MUTED, width: 0.8 } });
      }

      addCard(s, 0.5, CONTENT_TOP + 2.5, 9, 0.8, { strip: C.ACCENT });
      s.addText("Write the letter of each object in the right column on your whiteboard.", {
        x: 0.75, y: CONTENT_TOP + 2.62, w: 8.5, h: 0.55,
        fontSize: 15, fontFace: FONT_B, color: C.CHARCOAL, margin: 0, valign: "middle",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "Round: A, C, E, G     Not Round: B, D, F, H", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slide 4: Fluency — 1-1 correspondence (scattered collections)
  const sFluency = pres.addSlide();
  addTopBar(sFluency, STAGE_COLORS["1"]);
  addStageBadge(sFluency, 1, "Fluency");
  addTitle(sFluency, "Touch and Count", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });

  addCard(sFluency, 0.5, CONTENT_TOP, 9, 2.5, { strip: STAGE_COLORS["1"] });
  // Six scattered collections shown as numbers
  const collections = [
    { count: 8, label: "1" }, { count: 5, label: "2" }, { count: 12, label: "3" },
    { count: 7, label: "4" }, { count: 10, label: "5" }, { count: 15, label: "6" },
  ];
  collections.forEach((col, i) => {
    const cx = 0.8 + (i % 3) * 3.0;
    const cy = CONTENT_TOP + 0.2 + Math.floor(i / 3) * 1.2;
    // Small card per collection
    sFluency.addShape("roundRect", {
      x: cx, y: cy, w: 2.5, h: 0.95, rectRadius: 0.06,
      fill: { color: C.WHITE },
      line: { color: C.PRIMARY, width: 1 },
    });
    sFluency.addText(col.label + ".", {
      x: cx + 0.05, y: cy + 0.02, w: 0.3, h: 0.25,
      fontSize: 10, fontFace: FONT_B, color: C.MUTED, margin: 0,
    });
    // Draw scattered dots
    const dotPositions = [];
    for (let d = 0; d < Math.min(col.count, 10); d++) {
      const dx = 0.25 + (d % 5) * 0.38;
      const dy = 0.2 + Math.floor(d / 5) * 0.38;
      dotPositions.push([dx, dy]);
    }
    dotPositions.forEach(([dx, dy]) => {
      sFluency.addShape("roundRect", {
        x: cx + dx, y: cy + dy, w: 0.2, h: 0.2, rectRadius: 0.1,
        fill: { color: C.SECONDARY },
      });
    });
    if (col.count > 10) {
      sFluency.addText("+" + (col.count - 10) + " more", {
        x: cx + 1.6, y: cy + 0.55, w: 0.8, h: 0.25,
        fontSize: 9, fontFace: FONT_B, color: C.ALERT, margin: 0,
      });
    }
  });

  addCard(sFluency, 0.5, CONTENT_TOP + 2.7, 9, 0.7, { strip: C.ACCENT });
  sFluency.addText("Count each group carefully. Write how many on your whiteboard.", {
    x: 0.75, y: CONTENT_TOP + 2.82, w: 8.5, h: 0.45,
    fontSize: 15, fontFace: FONT_B, color: C.CHARCOAL, margin: 0, valign: "middle",
  });

  addFooter(sFluency, FOOTER);
  sFluency.addNotes(NOTES_FLUENCY);

  // Slide 5: LI/SC
  liSlide(pres,
    ["We are learning to put numbers in order from smallest to biggest"],
    [
      "I can say which number is bigger",
      "I can put three numbers in order",
      "I can find a missing number in a sequence",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 6: I Do — Ordering 3, 7, 9
  workedExSlide(pres, 2, "I Do", "Putting Numbers in Order",
    [
      "Numbers: 7, 3, 9",
      "",
      "Step 1: Find the smallest",
      "  When I count, 3 comes first. So 3.",
      "",
      "Step 2: What comes next?",
      "  7 comes before 9 when I count.",
      "",
      "Step 3: What is left?",
      "  9 is the biggest.",
      "",
      "Order: 3, 7, 9",
    ],
    NOTES_IDO, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.0, { strip: C.PRIMARY });
      slide.addText("Number Line Check", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.3, h: 0.3,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0, align: "center",
      });

      // Ordered number cards
      const orderNums = ["3", "7", "9"];
      orderNums.forEach((n, i) => {
        addTextOnShape(slide, n, {
          x: lg.rightX + 0.15 + i * 1.35, y: lg.panelTopPadded + 0.65, w: 0.75, h: 0.55, rectRadius: 0.08,
          fill: { color: C.PRIMARY },
        }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
        if (i < orderNums.length - 1) {
          slide.addText("->", {
            x: lg.rightX + 0.95 + i * 1.35, y: lg.panelTopPadded + 0.72, w: 0.35, h: 0.4,
            fontSize: 16, fontFace: FONT_H, color: C.CHARCOAL, align: "center", valign: "middle", margin: 0,
          });
        }
      });

      // Ordered result
      addTextOnShape(slide, "3  ->  7  ->  9", {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 1.6, w: lg.rightW - 0.6, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });

      slide.addText("Smallest                          Biggest", {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 2.15, w: lg.rightW - 0.6, h: 0.25,
        fontSize: 11, fontFace: FONT_B, color: C.MUTED, margin: 0,
      });

      // Key rule
      addTextOnShape(slide, "Numbers get bigger going right!", {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 2.5, w: lg.rightW - 0.6, h: 0.35, rectRadius: 0.06,
        fill: { color: C.PRIMARY },
      }, { fontSize: 12, fontFace: FONT_B, color: C.WHITE, bold: true });
    }
  );

  // Slide 7: I Do — Ordering 15, 11, 18
  workedExSlide(pres, 2, "I Do", "Ordering Teen Numbers",
    [
      "Numbers: 15, 11, 18",
      "",
      "Step 1: Find the smallest",
      "  11 comes first when counting. So 11.",
      "",
      "Step 2: What comes next?",
      "  15 comes before 18.",
      "",
      "Step 3: What is left?",
      "  18 is the biggest.",
      "",
      "Order: 11, 15, 18",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.6, { strip: C.PRIMARY });
      slide.addText("Number Line Check", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.3, h: 0.3,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0, align: "center",
      });

      // Ordered number cards
      const teenNums = ["11", "15", "18"];
      teenNums.forEach((n, i) => {
        addTextOnShape(slide, n, {
          x: lg.rightX + 0.15 + i * 1.35, y: lg.panelTopPadded + 0.65, w: 0.75, h: 0.55, rectRadius: 0.08,
          fill: { color: C.SECONDARY },
        }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
        if (i < teenNums.length - 1) {
          slide.addText("->", {
            x: lg.rightX + 0.95 + i * 1.35, y: lg.panelTopPadded + 0.72, w: 0.35, h: 0.4,
            fontSize: 16, fontFace: FONT_H, color: C.CHARCOAL, align: "center", valign: "middle", margin: 0,
          });
        }
      });

      addTextOnShape(slide, "11  ->  15  ->  18", {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 1.6, w: lg.rightW - 0.6, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });

      slide.addText("Smallest                          Biggest", {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 2.15, w: lg.rightW - 0.6, h: 0.25,
        fontSize: 11, fontFace: FONT_B, color: C.MUTED, margin: 0,
      });
    }
  );

  // Slide 8-9: CFU 1 (withReveal)
  withReveal(
    () => cfuSlide(pres, "CFU", "Quick Check", "Hand Signals",
      "Which is BIGGER?\n\n8  or  14\n\nLeft hand = 8     Right hand = 14",
      NOTES_CFU1, FOOTER),
    (slide) => {
      addTextOnShape(slide, "14 is bigger!  (14 comes after 8 when counting)", {
        x: 0.5, y: 4.0, w: 9, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 10-11: We Do (withReveal)
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Put These in Order!",
      [
        "With your partner, order each set",
        "from smallest to biggest.",
        "Use your number cards to help!",
        "",
        "Set 1:  5,  2,  8",
        "",
        "Set 2:  12,  7,  16",
        "",
        "Write the order on your whiteboard.",
      ],
      NOTES_WEDO_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 1.8, { strip: C.SECONDARY });
        slide.addText("Strategy:", {
          x: lg.rightX + 0.15, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.3, h: 0.3,
          fontSize: 15, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0,
        });
        slide.addText([
          { text: "1. Find the SMALLEST first", options: { bullet: true, fontSize: 14, color: C.CHARCOAL, breakLine: true } },
          { text: "2. Then the next one", options: { bullet: true, fontSize: 14, color: C.CHARCOAL, breakLine: true } },
          { text: "3. The BIGGEST goes last", options: { bullet: true, fontSize: 14, color: C.CHARCOAL } },
        ], {
          x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.45, w: lg.rightW - 0.5, h: 1.0,
          fontFace: FONT_B, margin: 0, valign: "top",
        });
      }
    ),
    (slide) => {
      addTextOnShape(slide, "Set 1: 2, 5, 8     Set 2: 7, 12, 16", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Slide 12-13: CFU 2 Hinge (withReveal) — Missing number
  withReveal(
    () => cfuSlide(pres, "CFU", "Hinge Check", "Show Me Boards",
      "What number is MISSING?\n\n10,  11,  ___,  13,  14\n\nWrite it on your whiteboard!",
      NOTES_CFU2, FOOTER),
    (slide) => {
      addTextOnShape(slide, "12!  (10, 11, 12, 13, 14)", {
        x: 1.5, y: 4.0, w: 7.0, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 14: You Do
  workedExSlide(pres, 4, "You Do", "Ordering Practice",
    [
      "First: Put each set in order (smallest",
      "         to biggest).",
      "Next: Find the missing numbers in",
      "         each sequence.",
      "Then: Try the challenge!",
      "",
      "You have 6 minutes.",
    ],
    NOTES_YOUDO, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 1.4, { strip: C.ALERT });
      slide.addText("Remember:", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.3, h: 0.3,
        fontSize: 14, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0,
      });
      slide.addText([
        { text: "Smallest first, biggest last", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Think about counting order", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Use the number line to check!", options: { bullet: true, fontSize: 13, color: C.PRIMARY, bold: true } },
      ], {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.42, w: lg.rightW - 0.5, h: 0.8,
        fontFace: FONT_B, margin: 0, valign: "top",
      });
    }
  );

  // Slide 15: Exit Ticket
  exitTicketSlide(pres,
    [
      "Which is bigger:  6  or  13?",
      "Put in order (smallest to biggest):  9,  4,  17",
      "What number is missing?  15, 16, ___, 18, 19",
    ],
    NOTES_EXIT, FOOTER);

  // Slide 16: Closing
  closingSlide(pres,
    "How do you decide which number is bigger? Tell your partner your strategy.",
    [
      "I can say which number is bigger",
      "I can put three numbers in order",
      "I can find a missing number in a sequence",
    ],
    NOTES_CLOSING);

  // Slide 17: Resources


  // Write PPTX
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "N20_Session4_Ordering_Numbers.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── Generate PDFs ─────────────────────────────────────────────────────────

  // Worksheet
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Ordering Numbers to 20",
      color: C.NAVY,
      lessonInfo: "Session 4 of 5 | Foundation Numeracy",
    });
    y = addTipBox(doc, "Put numbers in order from smallest to biggest. Think about counting order!", y, { color: C.TEAL });

    y = addSectionHeading(doc, "Section A: Order These Numbers", y, { color: C.NAVY });
    y = addProblem(doc, 1, "Order:  6,  2,  9        ___,  ___,  ___", y, { color: C.NAVY });
    y = addProblem(doc, 2, "Order:  4,  8,  1        ___,  ___,  ___", y, { color: C.NAVY });
    y = addProblem(doc, 3, "Order:  13,  5,  18      ___,  ___,  ___", y, { color: C.NAVY });
    y = addProblem(doc, 4, "Order:  11,  19,  14     ___,  ___,  ___", y, { color: C.NAVY });
    y = addProblem(doc, 5, "Order:  20,  7,  15      ___,  ___,  ___", y, { color: C.NAVY });

    y = addSectionHeading(doc, "Section B: Find the Missing Number", y, { color: C.NAVY });
    y = addProblem(doc, 6, "4,  5,  ___,  7,  8", y, { color: C.NAVY });
    y = addProblem(doc, 7, "11,  12,  ___,  14,  15", y, { color: C.NAVY });
    y = addProblem(doc, 8, "16,  ___,  18,  19,  20", y, { color: C.NAVY });
    y = addProblem(doc, 9, "___,  8,  9,  10,  11", y, { color: C.NAVY });

    y = addSectionHeading(doc, "Challenge: Order Four Numbers!", y, { color: C.NAVY });
    y = addProblem(doc, 10, "Order:  14,  3,  19,  8      ___,  ___,  ___,  ___", y, { color: C.NAVY });
    y = addProblem(doc, 11, "Order:  17,  6,  11,  20     ___,  ___,  ___,  ___", y, { color: C.NAVY });

    addPdfFooter(doc, "Session 4 | Numbers to 20 | Foundation");
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Answer Key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Ordering Numbers to 20",
      color: C.NAVY,
      lessonInfo: "Session 4 of 5 | Foundation Numeracy",
    });
    y = addSectionHeading(doc, "Section A", y, { color: C.NAVY });
    y = addBodyText(doc, "1. 2, 6, 9", y);
    y = addBodyText(doc, "2. 1, 4, 8", y);
    y = addBodyText(doc, "3. 5, 13, 18", y);
    y = addBodyText(doc, "4. 11, 14, 19", y);
    y = addBodyText(doc, "5. 7, 15, 20", y);
    y = addSectionHeading(doc, "Section B", y, { color: C.NAVY });
    y = addBodyText(doc, "6. Missing: 6", y);
    y = addBodyText(doc, "7. Missing: 13", y);
    y = addBodyText(doc, "8. Missing: 17", y);
    y = addBodyText(doc, "9. Missing: 7", y);
    y = addSectionHeading(doc, "Challenge", y, { color: C.NAVY });
    y = addBodyText(doc, "10. 3, 8, 14, 19", y);
    y = addBodyText(doc, "11. 6, 11, 17, 20", y);
    addPdfFooter(doc, "Session 4 | Answer Key | Foundation");
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  // Enabling Scaffold — Number line support, pairs only, 0-10
  await (async () => {
    const doc = createPdf({ title: ENABLING_RES.name });
    let y = addPdfHeader(doc, ENABLING_RES.name, {
      subtitle: "Ordering Numbers with a Number Line",
      color: C.TEAL,
      lessonInfo: "Session 4 of 5 | Foundation Numeracy",
    });
    y = addTipBox(doc, "Use the number line to help! Find each number on the line. The one on the LEFT is smaller.", y, { color: C.TEAL });

    // Draw a number line 0-10
    y += 5;
    const lineX = 60;
    const lineW = 450;
    doc.moveTo(lineX, y + 10).lineTo(lineX + lineW, y + 10).lineWidth(2).strokeColor("#" + C.NAVY).stroke();
    for (let i = 0; i <= 10; i++) {
      const tx = lineX + (i / 10) * lineW;
      doc.moveTo(tx, y + 4).lineTo(tx, y + 16).lineWidth(1.5).strokeColor("#" + C.NAVY).stroke();
      doc.fontSize(12).font("Sans-Bold").fillColor("#333333").text(String(i), tx - 6, y + 20);
    }
    y += 45;

    y = addSectionHeading(doc, "Which is bigger? Circle the bigger number.", y, { color: C.NAVY });
    y = addProblem(doc, 1, "3  or  7        bigger: _____", y, { color: C.NAVY });
    y = addProblem(doc, 2, "9  or  5        bigger: _____", y, { color: C.NAVY });
    y = addProblem(doc, 3, "2  or  8        bigger: _____", y, { color: C.NAVY });
    y = addProblem(doc, 4, "6  or  4        bigger: _____", y, { color: C.NAVY });

    y = addSectionHeading(doc, "Put in order (smallest to biggest)", y, { color: C.NAVY });
    y = addProblem(doc, 5, "3,  1,  5        ___,  ___,  ___", y, { color: C.NAVY });
    y = addProblem(doc, 6, "8,  4,  6        ___,  ___,  ___", y, { color: C.NAVY });
    y = addProblem(doc, 7, "10,  2,  7       ___,  ___,  ___", y, { color: C.NAVY });

    y = addSectionHeading(doc, "Find the missing number", y, { color: C.NAVY });
    y = addProblem(doc, 8, "2,  3,  ___,  5,  6", y, { color: C.NAVY });
    y = addProblem(doc, 9, "7,  ___,  9,  10", y, { color: C.NAVY });

    addPdfFooter(doc, "Session 4 | Enabling Scaffold | Foundation");
    await writePdf(doc, path.join(OUT_DIR, ENABLING_RES.fileName));
    console.log("PDF written: " + ENABLING_RES.fileName);
  })();

  // Extension — Ordering beyond 20 and double gaps
  await (async () => {
    const doc = createPdf({ title: EXTENDING_RES.name });
    let y = addPdfHeader(doc, EXTENDING_RES.name, {
      subtitle: "Ordering and Missing Numbers Challenge",
      color: C.TEAL,
      lessonInfo: "Session 4 of 5 | Foundation Numeracy",
    });

    y = addSectionHeading(doc, "Counting Beyond 20", y, { color: C.NAVY });
    y = addBodyText(doc, "Did you know numbers keep going after 20? Here is how:", y);
    y = addBodyText(doc, "20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30!", y);
    y = addBodyText(doc, "The pattern is the same as 0-9, but with a 2 at the front (twenties).", y);

    y = addSectionHeading(doc, "Order These Numbers (include numbers past 20!)", y, { color: C.NAVY });
    y = addProblem(doc, 1, "Order:  18,  25,  12      ___,  ___,  ___", y, { color: C.NAVY });
    y = addProblem(doc, 2, "Order:  22,  9,  30       ___,  ___,  ___", y, { color: C.NAVY });
    y = addProblem(doc, 3, "Order:  27,  14,  20,  6  ___,  ___,  ___,  ___", y, { color: C.NAVY });

    y = addSectionHeading(doc, "Find the Missing Numbers (two gaps!)", y, { color: C.NAVY });
    y = addProblem(doc, 4, "15,  ___,  17,  ___,  19,  20", y, { color: C.NAVY });
    y = addProblem(doc, 5, "___,  22,  23,  ___,  25,  26", y, { color: C.NAVY });
    y = addProblem(doc, 6, "8,  ___,  ___,  11,  12,  13", y, { color: C.NAVY });

    y = addSectionHeading(doc, "Count Backwards!", y, { color: C.NAVY });
    y = addBodyText(doc, "Fill in the counting backwards sequence:", y);
    y = addProblem(doc, 7, "20,  19,  18,  ___,  ___,  15,  ___,  13", y, { color: C.NAVY });

    y = addSectionHeading(doc, "Did You Know?", y, { color: C.NAVY });
    y = addBodyText(doc, "After 30 comes the thirties (31, 32, 33...), then the forties, fifties, and so on. Numbers go on FOREVER!", y);

    addPdfFooter(doc, "Session 4 | Extension | Foundation");
    await writePdf(doc, path.join(OUT_DIR, EXTENDING_RES.fileName));
    console.log("PDF written: " + EXTENDING_RES.fileName);
  })();

  console.log("Session 4 build complete.");
}

build().catch(err => { console.error(err); process.exit(1); });
