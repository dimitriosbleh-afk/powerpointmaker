"use strict";

// BODMAS Unit - Session 10: Calendar Finalisation and BODMAS Darts
// Year 5/6 Numeracy, Week 9 Term 2 (variant 2)
// Final session of unit - showcase + game + reflection

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

const SESSION = 10;
const FOOTER = "BODMAS | Session 10 of 10 - Final Session | Year 5/6 Numeracy";
const OUT_DIR = "output/BOD_Session10_Final_And_Games";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const DARTS_RES = makeSessionResource(SESSION, "BODMAS Darts Boards", "Easier and harder dartboard targets for the game.");
const DARTS_RULES = makeSessionResource(SESSION, "BODMAS Darts Rules", "How to play - rules and scoring.");
const RESOURCE_ITEMS = [DARTS_RES, DARTS_RULES];

fs.mkdirSync(RES_DIR, { recursive: true });

// --- Teacher Notes -----------------------------------------------------------

const NOTES_TITLE = `SAY:
- Final session of the BODMAS unit
- Today: finish your calendar, play BODMAS Darts, reflect on what you learned

DO:
- Display title slide
- Energy day - this is the celebration session
- Have BODMAS Darts boards ready

TEACHER NOTES:
Session 10 of 10. Last session. Mix of work and play. Calendar finalisation in the first half, then BODMAS Darts as a game-based application of skills. End with reflection on the unit.

WATCH FOR:
- Students who haven't finished their calendar - prioritise the finish
- Students who finished early - they can start Darts sooner

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_DR_Q = `SAY:
- Final daily review - one big BODMAS challenge
- Show me when done

DO:
- Display 1 challenging equation
- Allow 4 minutes
- Brief - we want maximum project + game time

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Tick or fix
- Look how far you've come

DO:
- Reveal answer
- Brief celebration

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_LI_SC = `SAY:
- Final session - read SC together

DO:
- Choral read

TEACHER NOTES:
SC focus on finalisation and applied use of BODMAS in the game.

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_FINAL_CHECKLIST = `SAY:
- Final calendar checklist
- Equation for every date 1-28? Tick
- Each one checked with BODMAS? Tick
- Right number of operations on each? Tick
- Top half decorated with month, name, pictures? Tick
- Written neatly? Tick
- Going on the wall? Yes!

DO:
- Display the final checklist
- Walk through each item
- Hand back any peer-check sheets

TEACHER NOTES:
The final checklist is a quality gate. Students who can tick every item are done. Students with missing items focus on those.

WATCH FOR:
- Students rushing to "done" - prompt them to actually run the check
- Students who notice their own missing items - this is the reflection happening

[Stage 2: I Do | VTLM 2.0: Final Quality]`;

const NOTES_FINISH_TIME = `SAY:
- Finish time
- 25 minutes
- Aim: every date filled, calendar decorated, ready to display
- Use the checklist
- If your calendar is done, hand it to me and pick up the BODMAS Darts board

DO:
- Allow 25 minutes
- Circulate
- Collect finished calendars
- Hand out Darts boards to early finishers

TEACHER NOTES:
This is the final push. Most students finish in this session. Some may need a final small group session OR home-time work to finish.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Pull small group of students who are behind. Focus on JUST the equations - skip decoration. Get the maths done first
- Extra Notes: Decoration can happen later
EXTENDING PROMPT:
- Task: Strong students - finished calendar - start Darts AND attempt to write the True Order of Operations bonus (equations where solving left-to-right gives a different answer)
- Extra Notes: This raises the calendar's quality

WATCH FOR:
- Students rushing to game - quality check first
- Students stuck on last equations - one-on-one support

[Stage 4: Finish | VTLM 2.0: Independent Practice]`;

const NOTES_DARTS_INTRO = `SAY:
- BODMAS Darts!
- A game where you USE your BODMAS skills
- Each dartboard has target numbers - some easy, some hard
- Use the digits provided to write equations that hit each target
- Score points for each target you hit

DO:
- Show the dartboard
- Read the rules
- Play one example as a class

TEACHER NOTES:
Darts is the celebration application. Students apply BODMAS to a fun puzzle. The game format raises engagement.

WATCH FOR:
- Students who instantly engage - celebrate
- Students who hesitate - pair them up

[Stage 3: We Do | VTLM 2.0: Application]`;

const NOTES_DARTS_PLAY = `SAY:
- Time to play
- Pair up - or play solo
- Use the dartboard sheet
- Write your equations in your books or on the sheet
- Compare with your partner - both got the same answer? Good
- Different answers? Check each other's work

DO:
- Hand out dartboard sheets
- Allow 15-20 minutes
- Circulate
- Award points or stickers for creative answers

TEACHER NOTES:
Game time. Students apply BODMAS at speed and with purpose. Build energy and excitement.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Small group at a separate dartboard. Use the easier targets only
- Extra Notes: Build confidence with small wins
EXTENDING PROMPT:
- Task: Strong students - try the harder dartboard. Or invent their own dartboard for a partner
- Extra Notes: Encourage creativity

WATCH FOR:
- Engagement - is the room buzzing?
- Quality - are students checking with BODMAS, not just guessing?

[Stage 4: Play | VTLM 2.0: Active Application]`;

const NOTES_REFLECT = `SAY:
- Two weeks of BODMAS - what did you learn?
- Self-check on the unit success criteria
- Tell your partner: what was your favourite moment?
- What was hardest?
- Pack up

DO:
- Display reflection slide
- Allow 5 minutes for partner talk
- Collect any final calendars and dart sheets

TEACHER NOTES:
End of unit reflection. Use this to gather feedback for next time.

[Stage 5: Reflect | VTLM 2.0: Reflection]`;

const NOTES_CLOSING = `SAY:
- Two weeks of BODMAS done
- Read the unit success criteria together
- Self-check
- Celebrate - this was a big project!

DO:
- Closing slide
- Self-check
- Brief celebration

TEACHER NOTES:
End-of-unit closing. Acknowledge growth. Calendars can go on the wall.

[General: Closing | VTLM 2.0: Reflection]`;

const NOTES_RESOURCES = `SAY:
- BODMAS Darts boards and rules - take one to play

DO:
- Hand out

[General: Resources | VTLM 2.0: Resource Awareness]`;

// --- Build -------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres,
    "Final Session - Finish and Play",
    "Complete your calendar. Play BODMAS Darts. Celebrate.",
    "Session 10 of 10  |  Year 5/6 Numeracy",
    NOTES_TITLE);

  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slide 2-3: Daily Review (one big challenge)
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Final BODMAS Challenge", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });

      addCard(s, 0.5, CONTENT_TOP, 9, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });

      addTextOnShape(s, "3 squared + (12 - 4) x 2 / 4 - 5", {
        x: 0.7, y: CONTENT_TOP + 0.7, w: 8.6, h: 1.4, rectRadius: 0.12,
        fill: { color: STAGE_COLORS["1"] },
      }, { fontSize: 28, fontFace: FONT_H, color: C.WHITE, bold: true });

      s.addText("All four BODMAS levels  |  4 minutes  |  Show every step", {
        x: 0.5, y: CONTENT_TOP + 2.3, w: 9, h: 0.4,
        fontSize: 15, fontFace: FONT_B, color: C.CHARCOAL, align: "center", italic: true, margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "B: (12-4)=8  O: 3 squared=9  D/M: 8x2=16, 16/4=4  A/S: 9+4-5 = 8", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 12, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slide 4: LI/SC
  liSlide(pres,
    ["I am finishing my BODMAS Calendar and using BODMAS to play and reflect"],
    [
      "I can complete my calendar to the rubric standard",
      "I can use BODMAS to write equations that hit a dartboard target",
      "I can reflect on what I learned about BODMAS this unit",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 5: Final checklist
  contentSlide(pres, "I Do", C.PRIMARY, "Final Calendar Checklist",
    [
      "Equation for every date 1 to 28",
      "Each one checked with BODMAS",
      "14 dates with 2 different operations",
      "14 dates with 3 different operations",
      "Top half decorated with month, name, pictures",
      "Written neatly",
    ],
    NOTES_FINAL_CHECKLIST, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.4, { strip: C.PRIMARY });
      slide.addText("Quick checklist", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0, align: "center",
      });

      const checks = [
        "All 28 dates filled",
        "Each equals the date",
        "Right digits used",
        "Right number of ops",
        "Decorated and neat",
      ];
      checks.forEach((c, i) => {
        const y = lg.panelTopPadded + 0.45 + i * 0.55;
        addTextOnShape(slide, "[ ]", {
          x: lg.rightX + 0.25, y, w: 0.55, h: 0.45, rectRadius: 0.06,
          fill: { color: C.WHITE }, line: { color: C.PRIMARY, width: 1 },
        }, { fontSize: 16, fontFace: FONT_H, color: C.PRIMARY, bold: true });
        slide.addText(c, {
          x: lg.rightX + 0.95, y, w: lg.rightW - 1.15, h: 0.45,
          fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, valign: "middle", margin: 0,
        });
      });
    }
  );

  // Slide 6: Finish time
  workedExSlide(pres, 4, "Finish Time", "Final Calendar Push",
    [
      "25 minutes",
      "Goal: every date filled, calendar decorated, ready to display",
      "Use the checklist - tick off as you go",
      "",
      "When done:",
      "  Hand your calendar in",
      "  Pick up a BODMAS Darts board",
      "  Start playing while others finish",
      "",
      "Stuck? Ask a partner first, then me",
    ],
    NOTES_FINISH_TIME, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.4, { strip: C.SECONDARY });
      slide.addText("Final push priorities", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0, align: "center",
      });

      const priorities = [
        { num: "1", text: "Equations FIRST" },
        { num: "2", text: "Decoration second" },
        { num: "3", text: "Quality check" },
        { num: "4", text: "Hand in" },
      ];
      priorities.forEach((p, i) => {
        const y = lg.panelTopPadded + 0.5 + i * 0.6;
        addTextOnShape(slide, p.num, {
          x: lg.rightX + 0.25, y, w: 0.55, h: 0.5, rectRadius: 0.06,
          fill: { color: C.SECONDARY },
        }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
        slide.addText(p.text, {
          x: lg.rightX + 0.95, y, w: lg.rightW - 1.15, h: 0.5,
          fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, bold: true, valign: "middle", margin: 0,
        });
      });

      addTextOnShape(slide, "25 minutes", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 2.95, w: lg.rightW - 0.4, h: 0.45, rectRadius: 0.06,
        fill: { color: C.ALERT },
      }, { fontSize: 14, fontFace: FONT_B, color: C.WHITE, bold: true });
    }
  );

  // Slide 7: Darts intro
  contentSlide(pres, "We Do", C.ACCENT, "BODMAS Darts!",
    [
      "Each dartboard has target numbers",
      "Use the digits provided to build equations",
      "Each equation must equal a target",
      "",
      "Score points for each target you hit",
      "Easier targets: 1 point. Harder targets: 3 points",
    ],
    NOTES_DARTS_INTRO, FOOTER,
    (slide, lg) => {
      // Mini dartboard mock
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.4, { strip: C.ACCENT });
      slide.addText("Sample dartboard", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 13, fontFace: FONT_H, color: C.ACCENT, bold: true, margin: 0, align: "center",
      });

      // Concentric circles for the dartboard look (smaller, lower)
      const cx = lg.rightX + lg.rightW / 2;
      const cy = lg.panelTopPadded + 1.95;
      slide.addShape("ellipse", {
        x: cx - 0.85, y: cy - 0.7, w: 1.7, h: 1.4,
        fill: { color: C.ACCENT },
      });
      slide.addShape("ellipse", {
        x: cx - 0.55, y: cy - 0.45, w: 1.1, h: 0.9,
        fill: { color: C.SECONDARY },
      });
      slide.addShape("ellipse", {
        x: cx - 0.27, y: cy - 0.22, w: 0.54, h: 0.44,
        fill: { color: C.SUCCESS },
      });

      // Inner number on the bullseye
      slide.addText("24", {
        x: cx - 0.27, y: cy - 0.18, w: 0.54, h: 0.36,
        fontSize: 12, fontFace: FONT_H, color: C.WHITE, bold: true, align: "center", valign: "middle", margin: 0,
      });

      // Sample target labels in a row below
      slide.addText("Targets: 10, 15, 24, 30 ...", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 2.85, w: lg.rightW - 0.4, h: 0.4,
        fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, italic: true, align: "center", margin: 0,
      });
    }
  );

  // Slide 8: Play time
  workedExSlide(pres, 4, "Play Time", "BODMAS Darts",
    [
      "Pair up or play solo",
      "Use the dartboard sheet provided",
      "For each target, write an equation that hits it",
      "Use the digits the dartboard gives you",
      "",
      "Score yourself - tally points",
      "Compare with your partner",
      "Different equations? Both right? Excellent!",
      "",
      "20 minutes - have fun!",
    ],
    NOTES_DARTS_PLAY, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.4, { strip: C.SUCCESS });
      slide.addText("Game time!", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 14, fontFace: FONT_H, color: C.SUCCESS, bold: true, margin: 0, align: "center",
      });

      slide.addText([
        { text: "Have fun", options: { bullet: true, fontSize: 14, color: C.CHARCOAL, breakLine: true } },
        { text: "Use BODMAS to check", options: { bullet: true, fontSize: 14, color: C.CHARCOAL, breakLine: true } },
        { text: "Try harder targets", options: { bullet: true, fontSize: 14, color: C.CHARCOAL, breakLine: true } },
        { text: "Be creative", options: { bullet: true, fontSize: 14, color: C.SUCCESS, bold: true } },
      ], {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.55, w: lg.rightW - 0.5, h: 1.8,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addTextOnShape(slide, "20 minutes", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 2.5, w: lg.rightW - 0.4, h: 0.5, rectRadius: 0.08,
        fill: { color: C.ALERT },
      }, { fontSize: 16, fontFace: FONT_B, color: C.WHITE, bold: true });
    }
  );

  // Slide 9: Reflection
  contentSlide(pres, "Reflect", C.PRIMARY, "Two Weeks of BODMAS",
    [
      "What's the most important thing you learned?",
      "What was hardest for you?",
      "What was your favourite moment?",
      "How will you use BODMAS in future maths?",
      "",
      "Tell your partner. Then tell the class.",
    ],
    NOTES_REFLECT, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.4, { strip: C.PRIMARY });
      slide.addText("This unit's journey", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0, align: "center",
      });

      const journey = [
        "1. What is BODMAS?",
        "2. Brackets and Orders",
        "3. M and D - same level",
        "4. A and S - same level",
        "5. Full BODMAS practice",
        "6-10. Calendar project",
      ];
      journey.forEach((j, i) => {
        const y = lg.panelTopPadded + 0.45 + i * 0.45;
        slide.addText(j, {
          x: lg.rightX + 0.3, y, w: lg.rightW - 0.6, h: 0.4,
          fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, valign: "middle", margin: 0,
        });
      });
    }
  );

  // Slide 10: Closing
  closingSlide(pres, {
    reflectionPrompt: "Two weeks of BODMAS done. What's the one thing you'll remember from this unit?",
    scItems: [
      "I can complete my calendar to the rubric standard",
      "I can use BODMAS to write equations that hit a dartboard target",
      "I can reflect on what I learned about BODMAS this unit",
    ],
    selfAssessment: { prompt: "Self-check", options: ["Got it", "Getting there", "Need more practice"] },
  }, NOTES_CLOSING);

  // Slide 11: Resources


  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "BOD_Session10_Final_And_Games.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // --- PDFs ------------------------------------------------------------------

  // Darts boards
  await (async () => {
    const doc = createPdf({ title: DARTS_RES.name });
    let y = addPdfHeader(doc, DARTS_RES.name, {
      subtitle: "Easier and Harder Dartboards",
      color: C.NAVY,
      lessonInfo: "Session 10 of 10 | Year 5/6 Numeracy",
    });
    y = addTipBox(doc, "For each dartboard, use the digits given to build equations that equal each target. Score 1 point per easier target, 3 per harder target.", y, { color: C.TEAL });

    // Easier dartboard
    y = addSectionHeading(doc, "Easier Dartboard - digits 2, 3, 4, 5", y, { color: C.NAVY });
    y = addBodyText(doc, "Targets:  6, 8, 10, 12, 14, 17, 20", y);
    y += 5;
    // Draw 7 target circles in a row
    const startX = 60;
    const circleR = 28;
    const spacing = 70;
    for (let i = 0; i < 7; i++) {
      const targets = [6, 8, 10, 12, 14, 17, 20];
      const cx = startX + i * spacing;
      doc.circle(cx, y + circleR, circleR).stroke("#" + C.NAVY);
      doc.fontSize(14).fillColor("#" + C.NAVY).text(String(targets[i]), cx - 12, y + circleR - 8, { width: 24, align: "center" });
    }
    y += circleR * 2 + 10;
    y = addBodyText(doc, "Your equation for 6:    ___________________________", y);
    y = addBodyText(doc, "Your equation for 8:    ___________________________", y);
    y = addBodyText(doc, "Your equation for 10:   ___________________________", y);
    y = addBodyText(doc, "Your equation for 12:   ___________________________", y);
    y = addBodyText(doc, "Your equation for 14:   ___________________________", y);
    y = addBodyText(doc, "Your equation for 17:   ___________________________", y);
    y = addBodyText(doc, "Your equation for 20:   ___________________________", y);

    doc.addPage();
    y = 50;

    // Harder dartboard
    y = addSectionHeading(doc, "Harder Dartboard - digits 2, 3, 4, 5, 6", y, { color: C.NAVY });
    y = addBodyText(doc, "Targets:  19, 21, 24, 27, 31, 34, 40", y);
    y += 5;
    for (let i = 0; i < 7; i++) {
      const targets = [19, 21, 24, 27, 31, 34, 40];
      const cx = startX + i * spacing;
      doc.circle(cx, y + circleR, circleR).stroke("#" + C.NAVY);
      doc.fontSize(14).fillColor("#" + C.NAVY).text(String(targets[i]), cx - 12, y + circleR - 8, { width: 24, align: "center" });
    }
    y += circleR * 2 + 10;
    y = addBodyText(doc, "Your equation for 19:   ___________________________", y);
    y = addBodyText(doc, "Your equation for 21:   ___________________________", y);
    y = addBodyText(doc, "Your equation for 24:   ___________________________", y);
    y = addBodyText(doc, "Your equation for 27:   ___________________________", y);
    y = addBodyText(doc, "Your equation for 31:   ___________________________", y);
    y = addBodyText(doc, "Your equation for 34:   ___________________________", y);
    y = addBodyText(doc, "Your equation for 40:   ___________________________", y);

    y += 10;
    y = addSectionHeading(doc, "My total score:  ___________ points", y, { color: C.NAVY });
    y = addBodyText(doc, "Easier targets: 1 point each.   Harder targets: 3 points each.", y);

    addPdfFooter(doc, "Session 10 | BODMAS Darts | Year 5/6");
    await writePdf(doc, path.join(OUT_DIR, DARTS_RES.fileName));
    console.log("PDF written: " + DARTS_RES.fileName);
  })();

  // Darts rules
  await (async () => {
    const doc = createPdf({ title: DARTS_RULES.name });
    let y = addPdfHeader(doc, DARTS_RULES.name, {
      subtitle: "How to Play",
      color: C.NAVY,
      lessonInfo: "Session 10 of 10 | Year 5/6 Numeracy",
    });
    y = addTipBox(doc, "BODMAS Darts is a game where you use BODMAS to hit target numbers. Score points for each target you hit with a valid equation.", y, { color: C.TEAL });

    y = addSectionHeading(doc, "Setup", y, { color: C.NAVY });
    y = addBodyText(doc, "1.  Pick a dartboard (easier or harder).", y);
    y = addBodyText(doc, "2.  Look at the digits available for that board.", y);
    y = addBodyText(doc, "3.  Look at the target numbers.", y);

    y = addSectionHeading(doc, "How to play", y, { color: C.NAVY });
    y = addBodyText(doc, "1.  Pick a target on the dartboard.", y);
    y = addBodyText(doc, "2.  Build an equation using the digits that equals that target.", y);
    y = addBodyText(doc, "3.  Use BODMAS to solve and check.", y);
    y = addBodyText(doc, "4.  Write your equation in the space provided.", y);
    y = addBodyText(doc, "5.  Move to the next target.", y);

    y = addSectionHeading(doc, "Rules", y, { color: C.NAVY });
    y = addBodyText(doc, "Each equation can use each digit only ONCE.", y);
    y = addBodyText(doc, "You can use +, -, x, /, brackets, and powers.", y);
    y = addBodyText(doc, "Brackets are a TOOL - they don't count as an operation.", y);
    y = addBodyText(doc, "Equations must SOLVE to the target with BODMAS.", y);

    y = addSectionHeading(doc, "Scoring", y, { color: C.NAVY });
    y = addBodyText(doc, "Easier dartboard: 1 point per target hit.", y);
    y = addBodyText(doc, "Harder dartboard: 3 points per target hit.", y);
    y = addBodyText(doc, "Bonus: 2 extra points if your equation uses brackets AND a power.", y);

    y = addSectionHeading(doc, "Partner play", y, { color: C.NAVY });
    y = addBodyText(doc, "Take turns picking targets.", y);
    y = addBodyText(doc, "Compare your equations - both got the same target with different equations? Excellent!", y);
    y = addBodyText(doc, "Check each other's BODMAS - do you both get the target?", y);

    y = addSectionHeading(doc, "Variations", y, { color: C.NAVY });
    y = addBodyText(doc, "Speed round: 5 minutes - hit as many targets as you can.", y);
    y = addBodyText(doc, "Brackets only: every equation must use at least one pair of brackets.", y);
    y = addBodyText(doc, "Inventor: design your own dartboard for a partner to play.", y);

    addPdfFooter(doc, "Session 10 | Darts Rules | Year 5/6");
    await writePdf(doc, path.join(OUT_DIR, DARTS_RULES.fileName));
    console.log("PDF written: " + DARTS_RULES.fileName);
  })();

  console.log("Session 10 build complete.");
}

build().catch(err => { console.error(err); process.exit(1); });
