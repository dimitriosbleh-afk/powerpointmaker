"use strict";

// Teen Numbers Unit — Session 4: Writing Teen Numbers
// Term 2 Week 1, Foundation Numeracy, Variant 0
// DR: Teen Numbers — I can name the numbers from 0 to at least 20
// Fluency: Writing teen numbers
// VC2MFN01 — name, represent and order numbers including zero to at least 20

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addProblem, addWriteLine,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const T = createTheme("numeracy", "foundation", 0);
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  workedExSlide, exitTicketSlide, addStageBadge,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle,
  withReveal, runSlideDiagnostics,
  STAGE_COLORS,
} = T;

const SESSION = 4;
const UNIT_TITLE = "Teen Numbers";
const FOOTER = "Teen Numbers | Session 4 of 5 | Foundation Numeracy";
const OUT_DIR = "output/Teen_Session4_Writing_Teen_Numbers";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Worksheet", "Writing practice for teen numbers 11-20.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Answer Key", "Teacher reference with correct number formations.");
const ENABLING_RES = makeSessionResource(SESSION, "Enabling Scaffold", "Traced number formations with visual guides.");
const EXTENDING_RES = makeSessionResource(SESSION, "Extension", "Number words and numeral matching investigation.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, ENABLING_RES, EXTENDING_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Teacher Notes ──────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- We have been learning to say and recognise teen numbers
- Today we are going to practise WRITING them
- Writing numbers is an important skill -- you need it for maths every day

DO:
- Display title slide
- Have whiteboards and markers ready

TEACHER NOTES:
Session 4 of 5. Shifts from recognition to production. Writing teen numbers requires both knowing what the number looks like (recognition) and being able to form the digits (fine motor). This is cognitively harder than identification.

WATCH FOR:
- Students who are confident with naming but anxious about writing -- reassure that practice makes it easier

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_DR_Q = `SAY:
- Let's see how well you know your numbers
- I will show you a ten frame. Tell me the number.
- Then I will say a number. You show me with your fingers how many extras.

DO:
- Flash 4 ten frame images (quick recognition drill)
- Students call out the number chorally
- Then 3 oral prompts: "Show me the extras for 15, 18, 12"

TEACHER NOTES:
Daily Review checks teen number naming from Sessions 1-2. The oral format keeps it fast and maintains the identification skill while today focuses on writing.

WATCH FOR:
- Students who name all numbers correctly and quickly -- naming is automated
- Students who hesitate on 16-19 -- they may need more naming practice alongside today's writing

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Check: ten frame A was 13, B was 17, C was 11, D was 20
- Extras: 15 has 5 extras, 18 has 8 extras, 12 has 2 extras
- Well done if you got them all!

DO:
- Click to reveal answers
- Quick choral confirmation
- Note any students still struggling with naming

TEACHER NOTES:
Quick tick-and-fix. Naming fluency underpins today's writing work -- students who cannot name the number will struggle to write it.

WATCH FOR:
- Students who got all correct -- ready for the writing focus
- Students who missed 17 or 20 -- may need a quick recap before I Do

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY = `SAY:
- Fluency time. Today we are going to air-write teen numbers
- Watch me first. I will write 14 in the air with my finger. 1... 4.
- Now you do it. Air-write 14. Big movements!
- Next: air-write 17. 1... 7.
- Now air-write 12. 1... 2.

DO:
- Model air-writing each number with large arm movements
- Students copy (standing at desks)
- Do 5 numbers: 14, 17, 12, 19, 16
- Emphasise: always start with the 1

TEACHER NOTES:
Air-writing engages gross motor before fine motor (whiteboard writing in We Do). The large movements build muscle memory for digit sequences. Emphasising "always start with 1" reinforces the tens-and-ones structure.

WATCH FOR:
- Students who write the second digit first (e.g., 7 then 1 for 17) -- they need the "1 first" rule
- Students who form digits correctly in the air -- gross motor is ready for fine motor

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me
- Read from slide: "We are learning to write the teen numbers correctly"
- Now our success criteria.
- Ask: What digit do ALL teen numbers start with? [1!]

DO:
- Choral read LI and SC
- Quick check: "What digit starts every teen number?" [1]

TEACHER NOTES:
The "1 first" insight is the key structural understanding for writing teen numbers. SC1 is the floor (write 11-20). SC2 is the core (match word to numeral). SC3 extends to dictation (hearing and writing).

WATCH FOR:
- Students who say "1" confidently -- they understand the structure
- Students who are unsure -- revisit the tens-and-ones idea before I Do

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO1 = `SAY:
- Watch me write teen numbers. The secret: they ALL start with 1.
- 11: I write a 1, then another 1. One-one. Eleven.
- 12: I write a 1, then a 2. One-two. Twelve.
- 13: I write a 1, then a 3. One-three. Thirteen.
- See the pattern? The first digit is always 1 (that is the ten). The second digit tells you the extras.

DO:
- Write each number large and slowly on the board or use the slide
- Say the digit sequence as you write: "one, then three"
- Circle the 1 in each number: "This 1 means TEN"
- Circle the second digit: "This tells us the extras"

TEACHER NOTES:
Connecting the written form to the ten-and-ones structure deepens understanding. Students are not just copying shapes -- they are writing a number that MEANS something. The I Do makes this meaning visible by linking each digit to the ten frame concept.

WATCH FOR:
- Students who mouth the digits as you write -- they are internalising the sequence
- Students who look at their fingers or count -- they are connecting digits to quantity

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO2 = `SAY:
- Now the bigger teen numbers. Same rule: 1 first, then the extras digit.
- 16: 1 then 6. The 6 tells us 6 extras after the ten.
- 18: 1 then 8. Eight extras.
- 20: This one is different! It starts with 2, then 0.
- 20 means two tens and zero extras. So we write 2-0.

DO:
- Write 16, 18, 20 clearly
- For 20, explain why it starts with 2 not 1
- Show the contrast: 12 starts with 1 (one ten), 20 starts with 2 (two tens)

MISCONCEPTIONS:
- Misconception: Students write teen numbers with digits reversed (e.g., 61 instead of 16)
  Why: In English, we say "sixteen" which starts with "six" -- students write what they hear first
  Impact: Reversed digits create a completely different number (61 vs 16), which will cause confusion in all future maths
  Quick correction: "I heard you write the 6 first because you hear 'six' first in 'sixteen.' But in teen numbers, the 1 always goes first because it is the TEN. The ten comes first: 1, then 6."

WATCH FOR:
- Students who correctly write 20 with a 2 first -- they understand the two-tens concept
- Students who write 02 for 20 -- they are overapplying the "1 first" rule

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU1 = `SAY:
- Quick check. I am going to say a number. You write it on your whiteboard.
- The number is... fourteen. Write fourteen.
- Ready... show me!

DO:
- Students write on whiteboards
- Scan for 14
- Check for reversals (41)

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: "Write the number fourteen on your whiteboard. Remember: the 1 comes first! You have 5 seconds. Show me."
- Scan for: 14. Check for: 41 (reversal), 4 (only second digit), 114 (extra digit).
PROCEED: If 80%+ write 14 correctly, move to We Do.
PIVOT: Most likely misconception -- students write 41 because they hear "four" first in "fourteen." Reteach: "Fourteen is a TEEN number. All teen numbers start with 1 because they have one TEN. The 1 goes first, then the 4. Say it with me: one-four, fourteen."

TEACHER NOTES:
Writing from dictation is harder than copying because students must recall the digit sequence. The "1 first" rule is the scaffold.

WATCH FOR:
- Students who write 14 correctly without hesitation -- writing is secure
- Students who write 41 -- the reversal needs addressing now before it becomes habitual

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Partner practice. I will say a number. Both of you write it.
- Then compare. Do they match?
- First number: seventeen.

DO:
- Call out 4 numbers: 17, 11, 19, 15
- Partners write on whiteboards and compare
- Allow 10 seconds per number
- Circulate to check formations

TEACHER NOTES:
We Do uses different numbers from the I Do examples. Partner comparison adds a checking mechanism. Students catch each other's reversals.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the Session 4 Enabling Scaffold with traced number outlines. Students trace over the numbers first, then write them independently beside the trace.
- Extra Notes: Distribute the Session 4 Enabling Scaffold PDF.
EXTENDING PROMPT:
- Task: Use the Session 4 Extension to match number words (eleven, twelve...) to numerals and write numbers from word prompts.
- Extra Notes: Self-contained investigation with word-to-numeral matching.

WATCH FOR:
- Students who write all four correctly -- digit formation is strong
- Partners who catch each other's reversals -- peer checking is working

[Stage 3: We Do | VTLM 2.0: Scaffold Practice]`;

const NOTES_WEDO_A = `SAY:
- Let's check. Seventeen: 1-7. Did you both write 17?
- Eleven: 1-1. Two ones! Did you get 11?
- Nineteen: 1-9. One-nine. 19.
- Fifteen: 1-5. One-five. 15.

DO:
- Reveal each number
- Students check and fix
- For each, say the digit sequence aloud

TEACHER NOTES:
Saying the digit sequence ("one-seven") reinforces the connection between the spoken name and the written form.

WATCH FOR:
- Students who self-corrected reversals -- they are learning from comparison
- All four correct -- ready for independent practice

[Stage 3: We Do Answers | VTLM 2.0: Scaffold Practice]`;

const NOTES_CFU2 = `SAY:
- Hinge check. I will say a number word. You write the number.
- Ready? Sixteen.
- Think about which digit comes first... write it... show me!

DO:
- Students write on whiteboards
- Scan for 16 (not 61)
- Cold call: "What digit did you write first? Why?"

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: "Write sixteen. Remember the rule! Show me in 5 seconds."
- Scan for: 16. Critical check: how many wrote 61?
- Follow up cold call: "What did you write first? [1] Why? [Because it is the ten / because teen numbers start with 1]"
PROCEED: If 80%+ write 16 correctly, move to You Do.
PIVOT: If significant reversals (61), reteach: "Sixteen sounds like six-teen, but we write the TEEN part first. Teen means ten. Ten is 1. So 1 goes first, then 6. One-six. Sixteen."

TEACHER NOTES:
16 is a strong reversal candidate because "six" is very prominent at the start of "sixteen." If students get this right, the rule is internalised.

WATCH FOR:
- Students who write 16 and can explain why the 1 comes first -- deep understanding
- Readiness signal: no reversals in the class

[Stage 3: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Independent writing practice
- On your worksheet, write each number I call out
- First: Listen to the number. Next: Think about the digits. Then: Write it carefully.
- I will also ask you to match number words to numbers.
- You have 6 minutes.

DO:
- Distribute Session 4 Worksheet
- Call out numbers for Section A (or students read the word prompts)
- Students work independently
- Circulate: check for reversals, especially 16, 17, 18

TEACHER NOTES:
You Do shifts from dictation (hearing) to reading (number words). This is a different input channel from the We Do, ensuring genuine transfer. The worksheet includes both sections.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Complete the enabling scaffold with traced numbers. Students trace 11-15 first, then write 11-15 independently. Simpler range only.
- Extra Notes: Distribute the Session 4 Enabling Scaffold PDF.
EXTENDING PROMPT:
- Task: Use the Session 4 Extension for number word matching (eleven to 11, twelve to 12, etc.) and writing numbers from word prompts.
- Extra Notes: Self-contained investigation.

WATCH FOR:
- Students who write all numbers without reversals -- formation is secure
- Students who match words to numerals correctly -- reading number words is developing
- Readiness signal: completing 8+ items correctly in 5 minutes

[Stage 4: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket. Write three numbers.
- 13, 16, 19. Work on your own. 2 minutes.

DO:
- Students write in workbooks or on whiteboards
- Check for correct formation and no reversals
- Sort: secure, developing, beginning

TEACHER NOTES:
Exit ticket targets SC1 (write the numbers) and checks for the common reversal error. Q2 (16) is the strongest reversal candidate.

WATCH FOR:
- Students who write all three correctly -- writing is secure
- Students who reverse 16 -- they need continued practice with the "1 first" rule

[Stage 5: Exit Ticket | VTLM 2.0: Monitor Progress]`;

const NOTES_CLOSING = `SAY:
- Let's check our success criteria.
- SC1: I can write the numbers 11 to 20. Thumbs?
- SC2: I can match a number word to the correct numeral. Thumbs?
- SC3: I can write a teen number when I hear it said aloud. Thumbs?
- Turn and talk: What is the rule for writing teen numbers? Tell your partner.

DO:
- Display SC
- Thumbs check
- Turn and Talk (answer: the 1 comes first / teen numbers start with 1)
- Cold call

TEACHER NOTES:
The Turn and Talk targets the "1 first" rule. Students who can articulate this have a reliable strategy for all teen number writing.

WATCH FOR:
- Students who say "1 first" or "start with 1" -- the rule is internalised
- Students who are unsure -- they need more practice in Session 5

[General: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- Today's resources are linked here

DO:
- Point out resources

TEACHER NOTES:
Resource slide for Session 4.

WATCH FOR:
- N/A

[General: Resources | VTLM 2.0: Planning]`;

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  titleSlide(pres, UNIT_TITLE, "Session 4: Writing Teen Numbers",
    "Foundation Numeracy | Session 4 of 5 | Term 2 Week 1", NOTES_TITLE);

  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // DR (withReveal) — naming numbers 0-20
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Name That Number!", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });

      addCard(s, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
      s.addText([
        { text: "What number is each ten frame?", options: { fontSize: 15, bold: true, color: C.PRIMARY, breakLine: true } },
        { text: "", options: { fontSize: 8, breakLine: true } },
        { text: "A.  Full ten + 3 extras = ___", options: { fontSize: 15, color: C.CHARCOAL, breakLine: true, paraSpaceAfter: 10 } },
        { text: "B.  Full ten + 7 extras = ___", options: { fontSize: 15, color: C.CHARCOAL, breakLine: true, paraSpaceAfter: 10 } },
        { text: "C.  Full ten + 1 extra = ___", options: { fontSize: 15, color: C.CHARCOAL, breakLine: true, paraSpaceAfter: 10 } },
        { text: "D.  Two full tens = ___", options: { fontSize: 15, color: C.CHARCOAL } },
      ], {
        x: 0.75, y: CONTENT_TOP + 0.1, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.2,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addCard(s, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.SECONDARY });
      s.addText([
        { text: "Call out each number!", options: { fontSize: 16, bold: true, color: C.SECONDARY, breakLine: true } },
        { text: "", options: { fontSize: 8, breakLine: true } },
        { text: "Then show me:", options: { fontSize: 14, color: C.CHARCOAL, breakLine: true } },
        { text: "How many extras for 15?", options: { bullet: true, fontSize: 14, color: C.CHARCOAL, breakLine: true } },
        { text: "How many extras for 18?", options: { bullet: true, fontSize: 14, color: C.CHARCOAL, breakLine: true } },
        { text: "How many extras for 12?", options: { bullet: true, fontSize: 14, color: C.CHARCOAL } },
      ], {
        x: 5.45, y: CONTENT_TOP + 0.1, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.2,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "A) 13   B) 17   C) 11   D) 20   |   Extras: 15=5, 18=8, 12=2", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 12, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Fluency — Air-writing teen numbers
  contentSlide(pres, "Fluency", STAGE_COLORS["1"], "Air-Write Teen Numbers",
    [
      "Stand up! Use your finger to write in the air.",
      "",
      "1.  Write 14 in the air: 1 ... 4",
      "2.  Write 17 in the air: 1 ... 7",
      "3.  Write 12 in the air: 1 ... 2",
      "4.  Write 19 in the air: 1 ... 9",
      "5.  Write 16 in the air: 1 ... 6",
    ],
    NOTES_FLUENCY, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.2, { strip: C.ACCENT });
      slide.addText("The Rule:", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 15, fontFace: FONT_H, color: C.ACCENT, bold: true, margin: 0,
      });

      addTextOnShape(slide, "1 always comes FIRST", {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.5, w: lg.rightW - 0.6, h: 0.45, rectRadius: 0.08,
        fill: { color: C.PRIMARY },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });

      slide.addText([
        { text: "The 1 means TEN", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "The second digit is the extras", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Big arm movements!", options: { bullet: true, fontSize: 13, color: C.CHARCOAL } },
      ], {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 1.1, w: lg.rightW - 0.5, h: 0.8,
        fontFace: FONT_B, margin: 0, valign: "top",
      });
    }
  );

  // LI/SC
  liSlide(pres,
    ["We are learning to write the teen numbers correctly"],
    [
      "I can write the numbers 11 to 20",
      "I can match a number word to the correct numeral",
      "I can write a teen number when I hear it said aloud",
    ],
    NOTES_LI_SC, FOOTER);

  // I Do — Writing 11-15
  workedExSlide(pres, 2, "I Do", "Writing Teen Numbers: 11-15",
    [
      "The SECRET: all teen numbers start with 1",
      "",
      "11 = write 1, then 1   (ten and one)",
      "12 = write 1, then 2   (ten and two)",
      "13 = write 1, then 3   (ten and three)",
      "14 = write 1, then 4   (ten and four)",
      "15 = write 1, then 5   (ten and five)",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.3, { strip: C.PRIMARY });
      slide.addText("Digit Sequence", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.3, h: 0.3,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0, align: "center",
      });

      // Large number display
      const nums = [11, 12, 13, 14, 15];
      nums.forEach((n, i) => {
        const ny = lg.panelTopPadded + 0.45 + i * 0.52;
        // Digit 1 in a circle
        slide.addShape("roundRect", {
          x: lg.rightX + 0.4, y: ny, w: 0.5, h: 0.42, rectRadius: 0.08,
          fill: { color: C.PRIMARY },
        });
        slide.addText("1", {
          x: lg.rightX + 0.4, y: ny, w: 0.5, h: 0.42,
          fontSize: 22, fontFace: FONT_H, color: C.WHITE, align: "center", valign: "middle", bold: true, margin: 0,
        });
        // Second digit
        slide.addShape("roundRect", {
          x: lg.rightX + 1.0, y: ny, w: 0.5, h: 0.42, rectRadius: 0.08,
          fill: { color: C.ACCENT },
        });
        slide.addText(String(n - 10), {
          x: lg.rightX + 1.0, y: ny, w: 0.5, h: 0.42,
          fontSize: 22, fontFace: FONT_H, color: C.WHITE, align: "center", valign: "middle", bold: true, margin: 0,
        });
        // = number
        slide.addText("= " + n, {
          x: lg.rightX + 1.65, y: ny, w: 1.2, h: 0.42,
          fontSize: 18, fontFace: FONT_H, color: C.CHARCOAL, valign: "middle", bold: true, margin: 0,
        });
      });
    }
  );

  // I Do — Writing 16-20
  workedExSlide(pres, 2, "I Do", "Writing Teen Numbers: 16-20",
    [
      "Same rule: 1 first, then the extras digit",
      "",
      "16 = write 1, then 6",
      "17 = write 1, then 7",
      "18 = write 1, then 8",
      "19 = write 1, then 9",
      "",
      "20 is different: write 2, then 0",
      "  (two tens, zero extras)",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.3, { strip: C.SECONDARY });
      slide.addText("Watch for the Switch!", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.3, h: 0.3,
        fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0, align: "center",
      });

      const nums2 = [16, 17, 18, 19, 20];
      nums2.forEach((n, i) => {
        const ny = lg.panelTopPadded + 0.45 + i * 0.52;
        const d1 = n === 20 ? "2" : "1";
        const d2 = n === 20 ? "0" : String(n - 10);
        const is20 = n === 20;

        slide.addShape("roundRect", {
          x: lg.rightX + 0.4, y: ny, w: 0.5, h: 0.42, rectRadius: 0.08,
          fill: { color: is20 ? C.ALERT : C.PRIMARY },
        });
        slide.addText(d1, {
          x: lg.rightX + 0.4, y: ny, w: 0.5, h: 0.42,
          fontSize: 22, fontFace: FONT_H, color: C.WHITE, align: "center", valign: "middle", bold: true, margin: 0,
        });
        slide.addShape("roundRect", {
          x: lg.rightX + 1.0, y: ny, w: 0.5, h: 0.42, rectRadius: 0.08,
          fill: { color: is20 ? C.ALERT : C.ACCENT },
        });
        slide.addText(d2, {
          x: lg.rightX + 1.0, y: ny, w: 0.5, h: 0.42,
          fontSize: 22, fontFace: FONT_H, color: C.WHITE, align: "center", valign: "middle", bold: true, margin: 0,
        });
        slide.addText("= " + n, {
          x: lg.rightX + 1.65, y: ny, w: 1.2, h: 0.42,
          fontSize: 18, fontFace: FONT_H, color: is20 ? C.ALERT : C.CHARCOAL, valign: "middle", bold: true, margin: 0,
        });
      });
    }
  );

  // CFU 1 (withReveal)
  withReveal(
    () => cfuSlide(pres, "CFU", "Write: Fourteen", "Show Me Boards",
      "Write the number fourteen on your whiteboard.\n\nRemember: which digit comes first?\n\nReady... show me!",
      NOTES_CFU1, FOOTER),
    (slide) => {
      addTextOnShape(slide, "14  (1 first, then 4)", {
        x: 1.0, y: 4.0, w: 8, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // We Do (withReveal)
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "Write These Numbers", { y: 0.65, fontSize: 22, color: STAGE_COLORS["3"] });

      addCard(s, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["3"] });
      s.addText([
        { text: "With your partner:", options: { fontSize: 16, bold: true, color: C.SECONDARY, breakLine: true } },
        { text: "Write each number on your whiteboard.", options: { fontSize: 14, color: C.CHARCOAL, breakLine: true } },
        { text: "Then compare with your partner.", options: { fontSize: 14, color: C.CHARCOAL, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "1.  seventeen", options: { fontSize: 18, bold: true, color: C.PRIMARY, breakLine: true, paraSpaceAfter: 12 } },
        { text: "2.  eleven", options: { fontSize: 18, bold: true, color: C.PRIMARY, breakLine: true, paraSpaceAfter: 12 } },
        { text: "3.  nineteen", options: { fontSize: 18, bold: true, color: C.PRIMARY, breakLine: true, paraSpaceAfter: 12 } },
        { text: "4.  fifteen", options: { fontSize: 18, bold: true, color: C.PRIMARY } },
      ], {
        x: 0.75, y: CONTENT_TOP + 0.1, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.2,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addCard(s, 5.2, CONTENT_TOP, 4.3, 2.0, { strip: C.ACCENT });
      addTextOnShape(s, "1 comes FIRST!", {
        x: 5.4, y: CONTENT_TOP + 0.2, w: 3.9, h: 0.5, rectRadius: 0.1,
        fill: { color: C.PRIMARY },
      }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });
      s.addText([
        { text: "Write the 1 (the ten)", options: { bullet: true, fontSize: 14, color: C.CHARCOAL, breakLine: true } },
        { text: "Then write the extras digit", options: { bullet: true, fontSize: 14, color: C.CHARCOAL, breakLine: true } },
        { text: "Compare with your partner", options: { bullet: true, fontSize: 14, color: C.CHARCOAL } },
      ], {
        x: 5.45, y: CONTENT_TOP + 0.85, w: 3.8, h: 1.0,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "1) 17    2) 11    3) 19    4) 15", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // CFU 2 Hinge (withReveal)
  withReveal(
    () => cfuSlide(pres, "CFU", "Write: Sixteen", "Show Me Boards",
      "Write the number sixteen.\n\nCareful! Which digit comes first?\n\nShow me!",
      NOTES_CFU2, FOOTER),
    (slide) => {
      addTextOnShape(slide, "16  (NOT 61!  The 1 comes first because it is the TEN)", {
        x: 0.5, y: 4.0, w: 9, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // You Do
  workedExSlide(pres, 4, "You Do", "Your Turn: Write Teen Numbers",
    [
      "First: Listen to (or read) each number.",
      "Next: Think about the digits (1 first!).",
      "Then: Write the number carefully.",
      "",
      "Section B: Match the number word to",
      "the correct numeral.",
      "",
      "You have 6 minutes.",
    ],
    NOTES_YOUDO, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 1.6, { strip: C.ALERT });
      slide.addText("Remember:", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 14, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0,
      });
      slide.addText([
        { text: "Teen numbers: 1 comes first", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "20 is different: 2 then 0", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Check for reversals!", options: { bullet: true, fontSize: 13, color: C.CHARCOAL } },
      ], {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.42, w: lg.rightW - 0.5, h: 0.9,
        fontFace: FONT_B, margin: 0, valign: "top",
      });
    }
  );

  // Exit Ticket
  exitTicketSlide(pres,
    [
      "Write the number: thirteen",
      "Write the number: sixteen",
      "Write the number: nineteen",
    ],
    NOTES_EXIT, FOOTER);

  // Closing
  closingSlide(pres,
    "Tell your partner: What is the rule for writing teen numbers?",
    [
      "I can write the numbers 11 to 20",
      "I can match a number word to the correct numeral",
      "I can write a teen number when I hear it said aloud",
    ],
    NOTES_CLOSING);

  // Resources


  // Write PPTX
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "Teen_Session4_Writing_Teen_Numbers.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── PDFs ──────────────────────────────────────────────────────────────────

  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, { subtitle: "Writing Teen Numbers 11-20", color: C.PRIMARY, lessonInfo: "Session 4 of 5 | Foundation Numeracy" });
    y = addTipBox(doc, "Remember: teen numbers start with 1! Write the 1 first, then the extras digit.", y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Section A: Write the number", y, { color: C.PRIMARY });
    const words = ["twelve", "sixteen", "fourteen", "eighteen", "eleven", "twenty", "thirteen", "nineteen"];
    words.forEach((w, i) => {
      y = addProblem(doc, i + 1, w + "  =  ___", y, { color: C.PRIMARY });
    });

    y = addSectionHeading(doc, "Section B: Match the word to the number", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Draw a line from each word to its number:", y);
    y = addBodyText(doc, "fifteen          18", y);
    y = addBodyText(doc, "eighteen         15", y);
    y = addBodyText(doc, "seventeen        20", y);
    y = addBodyText(doc, "twenty           17", y);

    addPdfFooter(doc, "Session 4 | Teen Numbers | Foundation Numeracy");
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF: " + WORKSHEET_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, { subtitle: "Writing Teen Numbers", color: C.PRIMARY, lessonInfo: "Session 4 of 5 | Foundation Numeracy" });
    y = addSectionHeading(doc, "Section A", y, { color: C.PRIMARY });
    y = addBodyText(doc, "1. twelve = 12    2. sixteen = 16    3. fourteen = 14    4. eighteen = 18", y);
    y = addBodyText(doc, "5. eleven = 11    6. twenty = 20    7. thirteen = 13    8. nineteen = 19", y);
    y = addSectionHeading(doc, "Section B", y, { color: C.PRIMARY });
    y = addBodyText(doc, "fifteen = 15, eighteen = 18, seventeen = 17, twenty = 20", y);
    addPdfFooter(doc, "Session 4 | Answer Key | Foundation Numeracy");
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF: " + ANSWER_KEY_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ENABLING_RES.name });
    let y = addPdfHeader(doc, ENABLING_RES.name, { subtitle: "Traced Number Practice", color: C.ACCENT, lessonInfo: "Session 4 of 5 | Foundation Numeracy" });
    y = addTipBox(doc, "Trace each number, then write it on your own next to it.", y, { color: C.ACCENT });
    y = addSectionHeading(doc, "Trace and write", y, { color: C.PRIMARY });
    [11, 12, 13, 14, 15].forEach((n, i) => {
      y = addProblem(doc, i + 1, "Trace: " + n + "    Write your own: ___", y, { color: C.PRIMARY });
    });
    addPdfFooter(doc, "Session 4 | Enabling Scaffold | Foundation Numeracy");
    await writePdf(doc, path.join(OUT_DIR, ENABLING_RES.fileName));
    console.log("PDF: " + ENABLING_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: EXTENDING_RES.name });
    let y = addPdfHeader(doc, EXTENDING_RES.name, { subtitle: "Number Words and Numerals", color: C.PRIMARY, lessonInfo: "Session 4 of 5 | Foundation Numeracy" });
    y = addSectionHeading(doc, "Number Words", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Every number has a NAME (word) and a NUMERAL (digits).", y);
    y = addBodyText(doc, "eleven = 11    twelve = 12    thirteen = 13    fourteen = 14    fifteen = 15", y);
    y = addBodyText(doc, "sixteen = 16    seventeen = 17    eighteen = 18    nineteen = 19    twenty = 20", y);
    y = addSectionHeading(doc, "Write the numeral for each word", y, { color: C.PRIMARY });
    const extWords = ["seventeen", "twelve", "twenty", "fifteen", "nineteen", "eleven"];
    extWords.forEach((w, i) => {
      y = addProblem(doc, i + 1, w + "  =  ___", y, { color: C.PRIMARY });
    });
    y = addSectionHeading(doc, "Write the word for each numeral", y, { color: C.PRIMARY });
    y = addProblem(doc, 7, "14  =  f___", y, { color: C.PRIMARY });
    y = addProblem(doc, 8, "18  =  e___", y, { color: C.PRIMARY });
    y = addProblem(doc, 9, "16  =  s___", y, { color: C.PRIMARY });
    y = addTipBox(doc, "Challenge: Can you write all the numbers from 10 to 20 in order AND write their names?", y, { color: C.ACCENT });
    addPdfFooter(doc, "Session 4 | Extension | Foundation Numeracy");
    await writePdf(doc, path.join(OUT_DIR, EXTENDING_RES.fileName));
    console.log("PDF: " + EXTENDING_RES.fileName);
  })();

  console.log("Session 4 build complete.");
}

build().catch(err => { console.error(err); process.exit(1); });
