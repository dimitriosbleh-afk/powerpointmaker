"use strict";

// Inquiry - Year 6 - Events That Shaped Australia (Timeline)
// One 60-minute lesson.
// Students learn to sequence events chronologically and plan a 3-event
// Google Slideshow for their grade's designated era.
// Worked example uses verifiable Australian history dates (1851 / 1854 / 1901)
// from the post-1800 British colonies / Federation period. Teachers can
// substitute events for their grade's era using the Event Planning Sheet.

const path = require("path");
const fs = require("fs");
const PptxGenJS = require("pptxgenjs");
const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf,
  addPdfHeader, addSectionHeading, addBodyText, addTipBox,
  addPdfFooter, addLinedArea, addWriteLine, addTwoColumnOrganiser,
  addResourceSlide, makeSessionResource, getSessionResourceFolder,
  PAGE, hex,
} = require("../themes/pdf_helpers");

// -- Theme --
const T = createTheme("inquiry", "grade56", weekToVariant(4));
const {
  C, FONT_H, FONT_B,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  withReveal,
  addTopBar, addBadge, addTitle, addCard, addFooter,
  addInstructionCard,
  CONTENT_TOP, SAFE_BOTTOM, SLIDE_W,
} = T;

// -- Output paths --
const UNIT = "Inquiry_Y6_Australia_Timeline";
const LESSON_FOLDER = path.join(__dirname, "..", "output", UNIT);
const PPTX_NAME = "Events That Shaped Australia - Timeline.pptx";
const FOOTER = "Inquiry | Year 6 | Events that shaped Australia";
const SESSION = 1;
const RES_DIR = path.join(LESSON_FOLDER, getSessionResourceFolder(SESSION));

// -- Resources --
const PLANNER = makeSessionResource(
  SESSION,
  "Event Planning Sheet",
  "Plan three events from your grade's era. Includes date, what happened, and one reason it shaped Australia."
);
const EXIT_TICKET = makeSessionResource(
  SESSION,
  "Exit Ticket",
  "One event chosen, its date, and one reason it shaped Australia."
);
const RESOURCE_ITEMS = [PLANNER, EXIT_TICKET];

fs.mkdirSync(RES_DIR, { recursive: true });

// ===============================================================
// Teacher Notes
// ===============================================================

const NOTES_TITLE = [
  "SAY:",
  "- Today we are going to start building the timelines that we'll display for our Enrichment showcase",
  "- Your job by the end of the lesson is to know which three events from your era you want to include and to have them in the right order",
  "",
  "DO:",
  "- Have the National Museum 'Defining Moments' page ready on the board for the launch",
  "- Have the Event Planning Sheet printed ready for the You Do",
  "",
  "TEACHER NOTES:",
  "Year 6's era for this unit is set by the teacher. The worked example below uses three well-known events from the post-1800 colonies and Federation period. Substitute events from your grade's era if helpful, but keep the chronology-and-significance skill the same.",
  "",
  "WATCH FOR:",
  "- Students settling in and ready to focus on history work",
  "",
  "[Inquiry: Title | VTLM 2.0: Establishing Purpose]",
].join("\n");

const NOTES_RESOURCES = [
  "SAY:",
  "- These are the two resources for today's lesson",
  "",
  "DO:",
  "- Print the Event Planning Sheet, one per student",
  "- Print the Exit Ticket, one per student",
  "- Open the National Museum 'Defining Moments' page in a browser tab",
  "- Open a blank Google Slideshow template ready to demonstrate slide structure later",
  "",
  "TEACHER NOTES:",
  "Materials needed: Chromebooks or iPads for students, mini-whiteboards and markers, sticky notes, board space for a class timeline. Have the National Museum and Deadly Story tabs open as research starting points students can use after this lesson.",
  "",
  "SOURCES:",
  "- National Museum of Australia, Defining Moments Timeline (https://www.nma.gov.au/defining-moments/defining-moments-timeline)",
  "- Deadly Story, Culture and History (https://www.deadlystory.com/page/culture/history)",
  "",
  "WATCH FOR:",
  "- Confirm Chromebook trolley is charged and Google account access is working before You Do",
  "",
  "[Inquiry: Resources | VTLM 2.0: Preparation]",
].join("\n");

const NOTES_LAUNCH = [
  "SAY:",
  "- Quick partner talk: name one event from Australia's history you already know about",
  "- Then tell your partner roughly when it happened. A year if you know it, or 'a long time ago' or 'fairly recent' if you don't",
  "- Ask: Why might it matter to know WHEN things happened, not just THAT they happened?",
  "",
  "DO:",
  "- Set a 60-second partner-talk timer",
  "- Cold call 2-3 pairs after the timer",
  "- Drop in: today we'll learn how historians build timelines",
  "",
  "TEACHER NOTES:",
  "The launch activates what students already know and bridges into the lesson focus, which is putting events in order and explaining why each one mattered.",
  "",
  "WATCH FOR:",
  "- Students who can name events but not order them - this is the gap today's lesson fills",
  "",
  "[Inquiry: Launch | VTLM 2.0: Activating Prior Knowledge]",
].join("\n");

const NOTES_VOCAB = [
  "SAY:",
  "- Four words before we start. Read each word and its meaning with me",
  "- 'Chronological order' just means earliest first, latest last",
  "- 'Era' is a chunk of history that has its own feel - the Gold Rush era, the post-Federation era, and so on",
  "- A 'defining moment' is an event that changed Australia in a way you can still see or feel today",
  "",
  "DO:",
  "- Read each word aloud, point to its picture or model",
  "- Have students say 'chronological' and 'era' back to you",
  "",
  "CFU CHECKPOINT:",
  "Technique: Finger Voting",
  "Script:",
  "- Say: I'll say a phrase. Show 1 if it is chronological order, 2 if it is not",
  "- Say: 1788, 1851, 1901 - [1]",
  "- Say: 1901, 1788, 1854 - [2]",
  "- Scan for: most students showing the correct finger count",
  "PROCEED: If 80% show correct, move on",
  "PIVOT: If many show 2 for the first one, remind students that chronological means earliest first",
  "",
  "TEACHER NOTES:",
  "Students need 'chronological order' and 'defining moment' to understand the success criteria. Keep this brisk.",
  "",
  "WATCH FOR:",
  "- Students mixing up 'era' with 'year' - clarify that an era is a longer chunk of time",
  "",
  "[Inquiry: Vocabulary | VTLM 2.0: Establishing Knowledge]",
].join("\n");

const NOTES_LI = [
  "SAY:",
  "- Read the Learning Intention aloud",
  "- Read each Success Criterion. We'll come back to these at the end",
  "",
  "DO:",
  "- Point to each SC as you read it",
  "- Leave the slide visible for about 20 seconds",
  "",
  "TEACHER NOTES:",
  "SC1 is the foundation any student can reach with the planning sheet. SC2 is the core target and is what the exit ticket assesses. SC3 stretches into selecting three events and explaining the significance of each one.",
  "",
  "WATCH FOR:",
  "- Students unsure what 'chronological order' means - we taught this on the vocabulary slide; reuse the language",
  "",
  "[Inquiry: LI/SC | VTLM 2.0: Learning Intentions]",
].join("\n");

const NOTES_IDO1 = [
  "SAY:",
  "- Watch how I put three events in chronological order",
  "- Think-aloud: I read each year first - 1851, 1854, 1901",
  "- Think-aloud: 1851 is the smallest number, so it goes on the left of my timeline",
  "- Think-aloud: 1854 is next - that's only three years later than the Gold Rush starting",
  "- Think-aloud: 1901 goes on the right - that's the most recent event",
  "- Ask: What did I look at first when I started ordering? [The year]",
  "",
  "DO:",
  "- Point to each year on the timeline as you say it",
  "- Trace your finger left to right across the timeline so students see the direction",
  "",
  "TEACHER NOTES:",
  "This I Do models the SC1 move: read year, find smallest, place left to right. Keep the focus on the sequencing skill, not on the historical content yet.",
  "",
  "WATCH FOR:",
  "- Students tracking with you on the timeline visual rather than reading the bullet steps - the visual is the teaching",
  "",
  "[Inquiry: I Do | VTLM 2.0: Explicit Modelling]",
].join("\n");

const NOTES_CFU_Q = [
  "SAY:",
  "- On your whiteboard, write which event happened FIRST",
  "- Your choices: the Eureka Stockade or Federation",
  "- You have 20 seconds. Use the dates if you need them",
  "",
  "DO:",
  "- Set a 20-second timer",
  "- Scan whiteboards as students hold them up",
  "",
  "CFU CHECKPOINT:",
  "Technique: Show Me Boards",
  "Script:",
  "- Say: Boards up on three. One, two, three",
  "- Scan for: 'Eureka Stockade' or '1854' - both are correct",
  "PROCEED: If 80% show Eureka or 1854, reveal and move on",
  "PIVOT: If many show Federation, model the comparison: '1854 versus 1901. Which number is smaller? That tells me which event came first.'",
  "",
  "TEACHER NOTES:",
  "This is the hinge check on SC1. The most likely error is choosing the more familiar event, not the earlier one. Use the year comparison to redirect.",
  "",
  "WATCH FOR:",
  "- Students looking at the years rather than guessing - praise the strategy out loud when you see it",
  "",
  "[Inquiry: CFU | VTLM 2.0: Checking for Understanding]",
].join("\n");

const NOTES_CFU_A = [
  "SAY:",
  "- The Eureka Stockade came first - 1854",
  "- Federation came later - 1901",
  "- The smaller the year number, the earlier the event",
  "",
  "DO:",
  "- Click to reveal the answer",
  "- Briefly affirm students who used the year strategy",
  "",
  "TEACHER NOTES:",
  "Validate the strategy of using the year number to decide. This is what we want students doing on their planning sheet next.",
  "",
  "WATCH FOR:",
  "- Students who answered Federation - check in privately before the We Do to confirm they have the year strategy now",
  "",
  "[Inquiry: CFU Reveal | VTLM 2.0: Feedback]",
].join("\n");

const NOTES_WEDO = [
  "SAY:",
  "- Now we'll order four events together, then choose three to keep",
  "- I'll read each event card. You write the years in order on your whiteboard",
  "- Then we'll decide together which three of the four to feature on a slideshow and why",
  "",
  "DO:",
  "- Read each event card aloud",
  "- Have one student come up and stick the cards on the board in order if you have physical cards ready",
  "- Open the floor: which three of the four are the strongest 'defining moments' for our era? Why?",
  "",
  "CFU CHECKPOINT:",
  "Technique: Turn and Tell",
  "Script:",
  "- Say: With your partner, agree on the three events you would put on a slideshow and one reason for each",
  "- Scan for: pairs naming three events and a reason - not just listing four",
  "PROCEED: If pairs are choosing three with a reason, reveal the model decision",
  "PIVOT: If pairs are stuck on choosing, model: 'I'd keep the Gold Rush because it brought new people. I'd keep Federation because it made one country.'",
  "",
  "TEACHER NOTES:",
  "We Do practises both SC1 (ordering) and SC3 (choosing three with significance). Students see that historians make choices about which events to feature - it isn't random.",
  "",
  "WATCH FOR:",
  "- Students wanting all four - good thinking; remind them the slideshow is only three slides, so they have to choose",
  "",
  "[Inquiry: We Do | VTLM 2.0: Guided Practice]",
].join("\n");

const NOTES_WEDO_REVEAL = [
  "SAY:",
  "- Here is one possible choice and the order",
  "- 1788, 1851, 1901 - First Fleet, Gold Rush, Federation",
  "- This is one valid answer. Yours may be different and still be a good choice",
  "",
  "DO:",
  "- Click to reveal the model answer",
  "- Validate alternative choices that came with a reason",
  "",
  "TEACHER NOTES:",
  "The reveal gives a model order. The choice of which three to keep is a thinking move - students should be able to defend their choice, not match the model exactly.",
  "",
  "WATCH FOR:",
  "- Students who match the model exactly without a reason - prompt: 'And WHY those three?'",
  "",
  "[Inquiry: We Do Reveal | VTLM 2.0: Consolidation]",
].join("\n");

const NOTES_IDO2 = [
  "SAY:",
  "- Once you have your three events and they are in order, you'll set up three slides in Google Slides",
  "- Watch how I set up ONE slide",
  "- Think-aloud: At the top I put the event name - that's my title",
  "- Think-aloud: Below the title I put the year - that's so the timeline order is clear",
  "- Think-aloud: I leave a space for an image - we'll add that after researching",
  "- Think-aloud: At the bottom I write one or two sentences on why this event shaped Australia",
  "",
  "DO:",
  "- Open a blank Google Slide on the board and build the four parts live as you talk",
  "- Show the same parts on this slide's right-hand mockup",
  "",
  "TEACHER NOTES:",
  "This second I Do covers SC2: organising on Google Slides with added information. Keep the demo short and concrete - the planning sheet does the heavy thinking.",
  "",
  "WATCH FOR:",
  "- Students copying the layout into the planning sheet - that's the bridge into the You Do",
  "",
  "[Inquiry: I Do | VTLM 2.0: Explicit Modelling]",
].join("\n");

const NOTES_YOUDO = [
  "SAY:",
  "- Your turn. Open your Event Planning Sheet",
  "- First: Choose three events from our era. Use the National Museum page if you need ideas",
  "- Next: Write the year for each event",
  "- Then: Write one sentence on why each event shaped Australia",
  "- We'll move to Google Slides once your planning sheet is complete",
  "",
  "DO:",
  "- Distribute the Event Planning Sheet",
  "- Project the National Museum Defining Moments page on the board",
  "- Circulate. Prompt quiet students with: 'Which event grabs your attention the most?'",
  "- Set a 12-minute timer for the planning step",
  "",
  "ENABLING & EXTENDING:",
  "ENABLING PROMPT:",
  "- Task: Pair the student with a partner. Use a shorter pre-selected list of 6 events with dates already attached. Student chooses three, copies the date and event name, and uses a sentence starter: 'This event shaped Australia because ___.'",
  "- Extra Notes: Sit close to support if needed.",
  "EXTENDING PROMPT:",
  "- Task: Student chooses three events that connect to each other - for example, one cause and two effects - and explains the link on a sticky note next to the planning sheet.",
  "",
  "TEACHER NOTES:",
  "The You Do uses different content than the We Do because students choose their own events from their era. Planning is the focus here; the Google Slides build follows once the sheet is complete.",
  "",
  "WATCH FOR:",
  "- Students picking three events of the same type (all wars, all gold rush events) - prompt for a mix",
  "- Students rushing to Google Slides without the planning sheet finished - hold the line on the order",
  "",
  "[Inquiry: You Do | VTLM 2.0: Supported Application]",
].join("\n");

const NOTES_EXIT = [
  "SAY:",
  "- One quick check before we finish",
  "- On your exit ticket, write ONE event you've chosen, its year, and ONE reason it shaped Australia",
  "- Keep it short - one event only",
  "",
  "DO:",
  "- Distribute the Exit Ticket",
  "- Allow 3 minutes",
  "- Collect as students leave",
  "",
  "TEACHER NOTES:",
  "This exit ticket assesses SC2 directly - event with a date and added information. Quickly sort into three piles after class: clear event with year and reason / partial / missing - plan next session's support around the third pile.",
  "",
  "WATCH FOR:",
  "- Students writing a reason without a year, or a year without a reason - prompt to add the missing piece before they leave",
  "",
  "[Inquiry: Exit Ticket | VTLM 2.0: Evidence of Learning]",
].join("\n");

const NOTES_CLOSING = [
  "SAY:",
  "- Back to our success criteria. Thumbs up, sideways, or down for each",
  "- I can put events in chronological order",
  "- I can plan an event on a slide with a date and a short explanation",
  "- I can choose three different events that shaped Australia for my slideshow",
  "- Ask: Why do historians put events in chronological order before they tell a story?",
  "",
  "DO:",
  "- Read each SC aloud and pause for thumbs",
  "- Note which SC students feel least confident about",
  "- Acknowledge that choosing three from many events is genuinely hard work",
  "",
  "TEACHER NOTES:",
  "The closing question links chronology to historical thinking - cause-and-effect needs an order to make sense. This is the bridge into next session's slide-building work.",
  "",
  "WATCH FOR:",
  "- Students showing thumbs down on SC3 - flag for a small-group revisit before the Google Slides build",
  "",
  "[Inquiry: Closing | VTLM 2.0: Review and Reflect]",
].join("\n");

// ===============================================================
// Drawing helpers
// ===============================================================

// Horizontal timeline with N event markers
function drawTimeline(s, x, y, w, h, events, opts) {
  const o = opts || {};
  const stripColor = o.strip || C.PRIMARY;

  // Card backing
  addCard(s, x, y, w, h, { strip: stripColor, fill: C.WHITE });

  // Caption (optional)
  if (o.caption) {
    s.addText(o.caption, {
      x: x + 0.2, y: y + 0.1, w: w - 0.4, h: 0.28,
      fontSize: 11, fontFace: FONT_H, color: stripColor, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
  }

  // Timeline axis
  const axisY = y + h * 0.55;
  const axisStart = x + 0.4;
  const axisEnd = x + w - 0.4;
  s.addShape("line", {
    x: axisStart, y: axisY, w: axisEnd - axisStart, h: 0,
    line: { color: stripColor, width: 2.5 },
  });

  // Arrow at end
  s.addShape("rtTriangle", {
    x: axisEnd - 0.1, y: axisY - 0.09, w: 0.18, h: 0.18,
    fill: { color: stripColor }, rotate: 90, line: { type: "none" },
  });

  // Event markers (inset from axis ends so labels stay inside the card)
  const n = events.length;
  const edgePad = 0.4;
  events.forEach((ev, i) => {
    const pos = n === 1
      ? (axisStart + axisEnd) / 2
      : axisStart + edgePad + (axisEnd - axisStart - 2 * edgePad) * (i / (n - 1));

    // Vertical tick
    s.addShape("line", {
      x: pos, y: axisY - 0.12, w: 0, h: 0.24,
      line: { color: stripColor, width: 2 },
    });

    // Year badge above
    s.addShape("roundRect", {
      x: pos - 0.45, y: axisY - 0.7, w: 0.9, h: 0.42,
      rectRadius: 0.08, fill: { color: stripColor },
    });
    s.addText(String(ev.year), {
      x: pos - 0.45, y: axisY - 0.7, w: 0.9, h: 0.42,
      fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    // Label below
    s.addText(ev.label, {
      x: pos - 0.6, y: axisY + 0.2, w: 1.2, h: 0.7,
      fontSize: 10.5, fontFace: FONT_B, color: C.CHARCOAL, bold: true,
      align: "center", valign: "top", margin: 0,
    });
  });
}

// Loose event card (used in We Do)
function drawEventCard(s, x, y, w, h, year, label, color) {
  s.addShape("roundRect", {
    x, y, w, h, rectRadius: 0.08,
    fill: { color: C.WHITE },
    line: { color, width: 1.3 },
  });
  s.addShape("roundRect", {
    x, y, w: w, h: 0.32, rectRadius: 0.08,
    fill: { color },
  });
  s.addText(String(year), {
    x, y, w, h: 0.32,
    fontSize: 13, fontFace: FONT_H, color: C.WHITE, bold: true,
    align: "center", valign: "middle", margin: 0,
  });
  s.addText(label, {
    x: x + 0.1, y: y + 0.35, w: w - 0.2, h: h - 0.4,
    fontSize: 11, fontFace: FONT_B, color: C.CHARCOAL, bold: true,
    align: "center", valign: "middle", margin: 0,
  });
}

// Google Slide mockup
function drawGoogleSlideMockup(s, x, y, w, h, opts) {
  const o = opts || {};
  const stripColor = o.strip || C.PRIMARY;

  // Frame to look like a slide preview
  s.addShape("rect", {
    x, y, w, h,
    fill: { color: C.WHITE },
    line: { color: C.MUTED, width: 0.6 },
  });

  // Title bar
  const titleH = 0.55;
  s.addShape("rect", {
    x, y, w, h: titleH,
    fill: { color: stripColor },
  });
  s.addText(o.title || "Federation", {
    x: x + 0.15, y, w: w - 1.6, h: titleH,
    fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true,
    align: "left", valign: "middle", margin: 0,
  });
  // Year pill on title bar
  s.addShape("roundRect", {
    x: x + w - 1.4, y: y + 0.1, w: 1.25, h: titleH - 0.2,
    rectRadius: 0.08, fill: { color: C.WHITE },
  });
  s.addText(o.year || "1901", {
    x: x + w - 1.4, y: y + 0.1, w: 1.25, h: titleH - 0.2,
    fontSize: 13, fontFace: FONT_H, color: stripColor, bold: true,
    align: "center", valign: "middle", margin: 0,
  });

  // Image placeholder area
  const imgY = y + titleH + 0.12;
  const imgH = (h - titleH - 0.3) * 0.55;
  s.addShape("rect", {
    x: x + 0.15, y: imgY, w: w - 0.3, h: imgH,
    fill: { color: "EDEAE2" },
    line: { color: C.MUTED, width: 0.5, dashType: "dash" },
  });
  s.addText("Image goes here", {
    x: x + 0.15, y: imgY, w: w - 0.3, h: imgH,
    fontSize: 11, fontFace: FONT_B, color: C.MUTED, italic: true,
    align: "center", valign: "middle", margin: 0,
  });

  // Caption box
  const capY = imgY + imgH + 0.1;
  const capH = h - (capY - y) - 0.15;
  s.addShape("roundRect", {
    x: x + 0.15, y: capY, w: w - 0.3, h: capH,
    rectRadius: 0.06,
    fill: { color: "FAF5E8" },
    line: { color: C.MUTED, width: 0.5 },
  });
  s.addText("Why it matters: short sentence here", {
    x: x + 0.25, y: capY, w: w - 0.5, h: capH,
    fontSize: 11, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
    align: "left", valign: "middle", margin: 0,
  });
}

// Planner preview (miniature)
function drawPlannerPreview(s, x, y, w, h) {
  addCard(s, x, y, w, h, { strip: C.ACCENT, fill: C.WHITE });
  s.addText("Event Planning Sheet", {
    x: x + 0.15, y: y + 0.08, w: w - 0.3, h: 0.3,
    fontSize: 12, fontFace: FONT_H, color: C.ACCENT, bold: true,
    align: "center", margin: 0,
  });

  const rowsY = y + 0.5;
  const rowH = (h - 0.65) / 3;
  const labels = ["Event 1", "Event 2", "Event 3"];
  labels.forEach((lbl, i) => {
    const ry = rowsY + i * rowH;
    s.addShape("roundRect", {
      x: x + 0.18, y: ry + 0.05, w: 1.0, h: rowH - 0.18,
      rectRadius: 0.05,
      fill: { color: C.ACCENT },
    });
    s.addText(lbl, {
      x: x + 0.18, y: ry + 0.05, w: 1.0, h: rowH - 0.18,
      fontSize: 10, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
    // Date stub
    s.addShape("rect", {
      x: x + 1.3, y: ry + 0.05, w: 0.7, h: rowH - 0.18,
      fill: { color: "FAF5E8" }, line: { color: C.MUTED, width: 0.4 },
    });
    s.addText("Year", {
      x: x + 1.3, y: ry + 0.05, w: 0.7, h: rowH - 0.18,
      fontSize: 9, fontFace: FONT_B, color: C.MUTED, italic: true,
      align: "center", valign: "middle", margin: 0,
    });
    // Reason stub
    s.addShape("rect", {
      x: x + 2.1, y: ry + 0.05, w: w - 2.3, h: rowH - 0.18,
      fill: { color: "FAF5E8" }, line: { color: C.MUTED, width: 0.4 },
    });
    s.addText("Why it shaped Australia", {
      x: x + 2.1, y: ry + 0.05, w: w - 2.3, h: rowH - 0.18,
      fontSize: 9, fontFace: FONT_B, color: C.MUTED, italic: true,
      align: "center", valign: "middle", margin: 0,
    });
  });
}

// ===============================================================
// Build
// ===============================================================

async function build() {
  const pres = new PptxGenJS();
  pres.layout = "LAYOUT_16x9";

  // -- Slide 1: Title --
  titleSlide(
    pres,
    "Events That\nShaped Australia",
    "Building a Timeline",
    "Year 6 | Inquiry | Session 1",
    NOTES_TITLE
  );

  // -- Slide 2: Teacher Resources --
  addResourceSlide(pres, RESOURCE_ITEMS, T, FOOTER, NOTES_RESOURCES);

  // -- Slide 3: Launch --
  contentSlide(
    pres,
    "Launch",
    C.SECONDARY,
    "When Did It Happen?",
    [
      "Name one event from Australia's history you already know",
      "Roughly when did it happen?",
      "Why might it matter to know when, not just what?",
    ],
    NOTES_LAUNCH,
    FOOTER,
    (s, lg) => {
      const rX = lg.rightX;
      const rW = lg.rightW;
      const topY = lg.panelTopPadded;

      // Three event cards "scattered" - illustrating the unordered state
      addCard(s, rX, topY, rW, 3.35, { strip: C.SECONDARY, fill: C.BG_LIGHT });
      s.addText("Without an order...", {
        x: rX + 0.2, y: topY + 0.1, w: rW - 0.4, h: 0.3,
        fontSize: 12, fontFace: FONT_H, color: C.SECONDARY, bold: true,
        align: "center", italic: true, margin: 0,
      });

      // Three angled-looking cards
      drawEventCard(s, rX + 0.3, topY + 0.55, 1.7, 0.85, "1851", "Gold Rush", C.PRIMARY);
      drawEventCard(s, rX + 2.25, topY + 0.55, 1.7, 0.85, "1901", "Federation", C.SECONDARY);
      drawEventCard(s, rX + 1.25, topY + 1.55, 1.7, 0.85, "1854", "Eureka Stockade", C.ACCENT);

      // Arrow + caption
      s.addShape("downArrow", {
        x: rX + rW / 2 - 0.15, y: topY + 2.5, w: 0.3, h: 0.35,
        fill: { color: C.SECONDARY }, line: { type: "none" },
      });
      s.addText("A timeline puts them in order", {
        x: rX + 0.2, y: topY + 2.92, w: rW - 0.4, h: 0.35,
        fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
    }
  );

  // -- Slide 4: Vocabulary --
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "Vocabulary", { color: C.PRIMARY, w: 1.65 });
    addTitle(s, "Four Words Before We Start");

    const vocab = [
      { word: "timeline", meaning: "a line that shows events in the order they happened, earliest on the left", color: C.PRIMARY },
      { word: "chronological order", meaning: "earliest first, latest last", color: C.SECONDARY },
      { word: "era", meaning: "a chunk of history with its own feel - for example, the Gold Rush era", color: C.ACCENT },
      { word: "defining moment", meaning: "an event that changed Australia in a way you can still see or feel today", color: C.SUCCESS },
    ];

    const cardW = 4.45;
    const cardH = 1.8;
    const gapX = 0.1;
    const gapY = 0.1;

    vocab.forEach((v, i) => {
      const col = i % 2;
      const row = Math.floor(i / 2);
      const x = 0.5 + col * (cardW + gapX);
      const y = CONTENT_TOP + row * (cardH + gapY);

      addCard(s, x, y, cardW, cardH, { strip: v.color, fill: C.WHITE });

      s.addText(v.word, {
        x: x + 0.2, y: y + 0.15, w: cardW - 0.4, h: 0.5,
        fontSize: 22, fontFace: FONT_H, color: v.color, bold: true, margin: 0,
        fit: "shrink",
      });

      s.addText(v.meaning, {
        x: x + 0.2, y: y + 0.7, w: cardW - 0.4, h: cardH - 0.85,
        fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
        valign: "top",
      });
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_VOCAB);
  }

  // -- Slide 5: LI / SC --
  liSlide(
    pres,
    ["I am learning to make a timeline of specific events that shaped Australia"],
    [
      "I can put events in chronological order",
      "I can plan an event on a slide with a date and a short explanation",
      "I can choose three different events that shaped Australia for my slideshow",
    ],
    NOTES_LI,
    FOOTER
  );

  // -- Slide 6: I Do - Sequencing --
  contentSlide(
    pres,
    "I Do",
    C.PRIMARY,
    "Putting Events in Chronological Order",
    [
      "Read the year for each event",
      "Find the smallest year - that one goes on the left",
      "Place each event left to right by year",
    ],
    NOTES_IDO1,
    FOOTER,
    (s, lg) => {
      const rX = lg.rightX;
      const rW = lg.rightW;
      const topY = lg.panelTopPadded;

      drawTimeline(s, rX, topY, rW, 3.0, [
        { year: 1851, label: "Gold Rush begins" },
        { year: 1854, label: "Eureka Stockade" },
        { year: 1901, label: "Federation" },
      ], {
        caption: "Three events from 1851 - 1901",
        strip: C.PRIMARY,
      });
    }
  );

  // -- Slide 7 / 7a: CFU hinge with reveal --
  function buildCfuBase() {
    const s = pres.addSlide();
    addTopBar(s, C.ALERT);
    addBadge(s, "CFU", { color: C.ALERT });
    addTitle(s, "Which Event Came First?", { color: C.ALERT });

    // Technique pill
    s.addShape("roundRect", {
      x: 0.5, y: CONTENT_TOP, w: 2.8, h: 0.40, rectRadius: 0.08,
      fill: { color: C.ALERT },
    });
    s.addText("Show Me Boards", {
      x: 0.5, y: CONTENT_TOP, w: 2.8, h: 0.40,
      fontSize: 12, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    // Layout positions
    const aH = 0.55;
    const aY = SAFE_BOTTOM - aH - 0.05;
    const qY = CONTENT_TOP + 0.56;
    const optsY = qY + 0.7;
    const optsH = aY - 0.2 - optsY;

    // Question card
    addCard(s, 0.5, qY, 9, 0.7, { strip: C.ALERT, fill: C.WHITE });
    s.addText("Which event came FIRST?", {
      x: 0.75, y: qY + 0.15, w: 8.5, h: 0.4,
      fontSize: 22, fontFace: FONT_B, color: C.CHARCOAL, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    // Two option cards
    const optW = 4.2;
    const optGap = 0.6;
    const opt1X = (10 - optW * 2 - optGap) / 2;
    const opt2X = opt1X + optW + optGap;

    // Option A - Eureka
    addCard(s, opt1X, optsY, optW, optsH, { strip: C.PRIMARY, fill: C.WHITE });
    s.addShape("roundRect", {
      x: opt1X + 0.3, y: optsY + 0.2, w: optW - 0.6, h: 0.55,
      rectRadius: 0.08, fill: { color: C.PRIMARY },
    });
    s.addText("1854", {
      x: opt1X + 0.3, y: optsY + 0.2, w: optW - 0.6, h: 0.55,
      fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
    s.addText("Eureka Stockade", {
      x: opt1X + 0.2, y: optsY + 0.95, w: optW - 0.4, h: 0.6,
      fontSize: 18, fontFace: FONT_B, color: C.CHARCOAL, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    // Option B - Federation
    addCard(s, opt2X, optsY, optW, optsH, { strip: C.SECONDARY, fill: C.WHITE });
    s.addShape("roundRect", {
      x: opt2X + 0.3, y: optsY + 0.2, w: optW - 0.6, h: 0.55,
      rectRadius: 0.08, fill: { color: C.SECONDARY },
    });
    s.addText("1901", {
      x: opt2X + 0.3, y: optsY + 0.2, w: optW - 0.6, h: 0.55,
      fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
    s.addText("Federation", {
      x: opt2X + 0.2, y: optsY + 0.95, w: optW - 0.4, h: 0.6,
      fontSize: 18, fontFace: FONT_B, color: C.CHARCOAL, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    // Answer placeholder
    s.addShape("roundRect", {
      x: 0.5, y: aY, w: 9, h: aH, rectRadius: 0.1,
      fill: { color: C.BG_LIGHT },
      line: { color: C.MUTED, width: 0.6, dashType: "dash" },
    });
    s.addText("Answer revealed next click", {
      x: 0.5, y: aY, w: 9, h: aH,
      fontSize: 11, fontFace: FONT_B, color: C.MUTED, italic: true,
      align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_CFU_Q);
    return s;
  }

  withReveal(
    buildCfuBase,
    (s) => {
      const aH = 0.55;
      const aY = SAFE_BOTTOM - aH - 0.05;
      s.addShape("roundRect", {
        x: 0.5, y: aY, w: 9, h: aH, rectRadius: 0.1,
        fill: { color: C.SUCCESS },
      });
      s.addText("Eureka Stockade - 1854 came before 1901", {
        x: 0.5, y: aY, w: 9, h: aH,
        fontSize: 16, fontFace: FONT_B, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      s.addNotes(NOTES_CFU_A);
    }
  );

  // -- Slide 8 / 8a: We Do - Pick three and order --
  function buildWeDoBase() {
    const s = pres.addSlide();
    addTopBar(s, C.SECONDARY);
    addBadge(s, "We Do", { color: C.SECONDARY });
    addTitle(s, "Pick Three and Order Them");

    // Left instruction card
    addInstructionCard(s, [
      { role: "header", text: "Our Task" },
      { role: "body", text: "Read the four events" },
      { role: "body", text: "Order them earliest to latest" },
      { role: "body", text: "Choose three to keep for a slideshow" },
    ], {
      x: 0.5, y: CONTENT_TOP, w: 3.25, h: SAFE_BOTTOM - CONTENT_TOP,
      strip: C.SECONDARY, fill: C.WHITE,
      headerColor: C.SECONDARY,
    });

    // Right: four event cards in a 2x2 grid
    const rX = 4.05;
    const rW = 5.45;
    const cardW = (rW - 0.25) / 2;
    const cardH = 1.5;
    const gap = 0.2;

    const events = [
      { year: 1901, label: "Federation - one Australia", color: C.SECONDARY },
      { year: 1788, label: "First Fleet arrives", color: C.PRIMARY },
      { year: 1915, label: "Gallipoli landing", color: C.MUTED },
      { year: 1851, label: "Gold Rush begins", color: C.ACCENT },
    ];

    events.forEach((ev, i) => {
      const col = i % 2;
      const row = Math.floor(i / 2);
      const x = rX + col * (cardW + 0.25);
      const y = CONTENT_TOP + row * (cardH + gap);
      drawEventCard(s, x, y, cardW, cardH, ev.year, ev.label, ev.color);
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_WEDO);
    return s;
  }

  withReveal(
    buildWeDoBase,
    (s) => {
      // Cover the left card with the model order
      const leftH = SAFE_BOTTOM - CONTENT_TOP;
      s.addShape("roundRect", {
        x: 0.5, y: CONTENT_TOP, w: 3.25, h: leftH,
        rectRadius: 0.1, fill: { color: C.SUCCESS },
      });
      s.addText([
        { text: "One valid choice", options: { fontSize: 18, bold: true, breakLine: true } },
        { text: "", options: { fontSize: 8, breakLine: true } },
        { text: "1788 - First Fleet", options: { fontSize: 14, breakLine: true } },
        { text: "1851 - Gold Rush", options: { fontSize: 14, breakLine: true } },
        { text: "1901 - Federation", options: { fontSize: 14, breakLine: true } },
        { text: "", options: { fontSize: 8, breakLine: true } },
        { text: "Yours may be different and still be a good choice", options: { fontSize: 12, italic: true } },
      ], {
        x: 0.65, y: CONTENT_TOP + 0.3, w: 2.95, h: leftH - 0.6,
        fontFace: FONT_B, color: C.WHITE,
        align: "center", valign: "middle", margin: 0,
      });

      // Outline the three chosen events on the right with a green "chosen" ring
      const rX = 4.05;
      const rW = 5.45;
      const cardW = (rW - 0.25) / 2;
      const cardH = 1.5;
      const gap = 0.2;
      // Order chosen: 1788 (i=1, col=1 row=0), 1851 (i=3, col=1 row=1), 1901 (i=0, col=0 row=0)
      const chosenIdx = [1, 3, 0];
      chosenIdx.forEach((idx) => {
        const col = idx % 2;
        const row = Math.floor(idx / 2);
        const x = rX + col * (cardW + 0.25);
        const y = CONTENT_TOP + row * (cardH + gap);
        // Thick green selection ring
        s.addShape("roundRect", {
          x: x - 0.06, y: y - 0.06, w: cardW + 0.12, h: cardH + 0.12,
          rectRadius: 0.12,
          fill: { type: "none" },
          line: { color: C.SUCCESS, width: 3.5 },
        });
        // Tick badge in the top-right corner of the card
        s.addShape("roundRect", {
          x: x + cardW - 0.35, y: y - 0.15, w: 0.4, h: 0.4,
          rectRadius: 0.2, fill: { color: C.SUCCESS },
        });
        s.addText("✓", {
          x: x + cardW - 0.35, y: y - 0.15, w: 0.4, h: 0.4,
          fontSize: 18, fontFace: FONT_B, color: C.WHITE, bold: true,
          align: "center", valign: "middle", margin: 0,
        });
      });

      s.addNotes(NOTES_WEDO_REVEAL);
    }
  );

  // -- Slide 9: I Do - Google Slide structure --
  contentSlide(
    pres,
    "I Do",
    C.PRIMARY,
    "What Goes On Each Slide",
    [
      "Title - the event name",
      "Year - when it happened",
      "Image - what we see",
      "Why it matters - one or two sentences",
    ],
    NOTES_IDO2,
    FOOTER,
    (s, lg) => {
      const rX = lg.rightX;
      const rW = lg.rightW;
      const topY = lg.panelTopPadded;

      drawGoogleSlideMockup(s, rX, topY, rW, 3.3, {
        title: "Federation",
        year: "1901",
        strip: C.PRIMARY,
      });
    }
  );

  // -- Slide 10: You Do --
  contentSlide(
    pres,
    "You Do",
    C.ACCENT,
    "Plan Your Three Events",
    [
      "First: Choose three events from your era",
      "Next: Write the year for each event",
      "Then: Write one reason each shaped Australia",
    ],
    NOTES_YOUDO,
    FOOTER,
    (s, lg) => {
      const rX = lg.rightX;
      const rW = lg.rightW;
      const topY = lg.panelTopPadded;

      drawPlannerPreview(s, rX, topY, rW, 3.3);
    }
  );

  // -- Slide 11: Exit Ticket --
  cfuSlide(
    pres,
    "Exit",
    "One Event with Date and Reason",
    "Exit Ticket",
    "Write ONE event you have chosen, its year, and ONE reason it shaped Australia.\n\nExample frame: '___ (year ___) shaped Australia because ___.'",
    NOTES_EXIT,
    FOOTER
  );

  // -- Slide 12: Closing --
  closingSlide(
    pres,
    "Why do historians put events in chronological order before they tell a story?",
    [
      "I can put events in chronological order",
      "I can plan an event on a slide with a date and a short explanation",
      "I can choose three different events that shaped Australia for my slideshow",
    ],
    NOTES_CLOSING
  );

  // -- Write PPTX --
  fs.mkdirSync(LESSON_FOLDER, { recursive: true });
  await pres.writeFile({ fileName: path.join(LESSON_FOLDER, PPTX_NAME) });
  console.log("PPTX written to", path.join(LESSON_FOLDER, PPTX_NAME));

  // ===============================================================
  // PDF 1: Event Planning Sheet
  // ===============================================================
  {
    const doc = createPdf({ title: "Event Planning Sheet" });
    let y = addPdfHeader(doc, "Session 1 Event Planning Sheet", {
      subtitle: "Plan three events from your era for your Google Slideshow",
      color: C.ACCENT,
      lessonInfo: "Year 6 Inquiry - Events that shaped Australia",
    });

    y = addTipBox(doc, "Choose three events from your grade's era. Put them in chronological order. Write one reason each event shaped Australia.", y, { color: C.PRIMARY });

    // Three event rows. Each row: number badge, year box, event name lines, reason lines.
    const rowY = (yStart, num) => {
      let yy = yStart;
      // Section heading: "Event N"
      yy = addSectionHeading(doc, `Event ${num}`, yy, { color: C.ACCENT });

      const boxX = PAGE.MARGIN;
      const boxW = PAGE.CONTENT_W;
      const boxH = 130;

      doc.save();
      doc.roundedRect(boxX, yy, boxW, boxH, 6).fill("#FAF5E8");
      doc.roundedRect(boxX, yy, boxW, boxH, 6).lineWidth(0.8).strokeColor(hex(C.MUTED)).stroke();
      doc.restore();

      // Year label and box
      doc.fontSize(11).font("Sans-Bold").fillColor(hex(C.ACCENT));
      doc.text("Year:", boxX + 14, yy + 14);
      doc.save();
      doc.roundedRect(boxX + 60, yy + 10, 110, 26, 3).lineWidth(0.7).strokeColor("#888888").stroke();
      doc.restore();

      // Event name label and lines
      doc.fontSize(11).font("Sans-Bold").fillColor(hex(C.ACCENT));
      doc.text("Event name:", boxX + 14, yy + 48);
      // Single long line
      doc.save();
      doc.moveTo(boxX + 14, yy + 70).lineTo(boxX + boxW - 14, yy + 70)
        .lineWidth(0.9).strokeColor("#000000").stroke();
      doc.restore();

      // Why it shaped Australia
      doc.fontSize(11).font("Sans-Bold").fillColor(hex(C.ACCENT));
      doc.text("Why it shaped Australia:", boxX + 14, yy + 80);
      // Two lines
      doc.save();
      doc.moveTo(boxX + 14, yy + 102).lineTo(boxX + boxW - 14, yy + 102)
        .lineWidth(0.9).strokeColor("#000000").stroke();
      doc.moveTo(boxX + 14, yy + 122).lineTo(boxX + boxW - 14, yy + 122)
        .lineWidth(0.9).strokeColor("#000000").stroke();
      doc.restore();

      return yy + boxH + 8;
    };

    y = rowY(y, 1);
    y = rowY(y, 2);
    y = rowY(y, 3);

    y += 6;
    y = addTipBox(doc, "Once your three events are planned, you are ready to build them in Google Slides. Use one slide per event.", y, { color: C.SUCCESS });

    addPdfFooter(doc, "Inquiry | Year 6 | Session 1 - Event Planning");
    const outPath = path.join(RES_DIR, "Session 1 Event Planning Sheet.pdf");
    await writePdf(doc, outPath);
    console.log("PDF written:", outPath);
  }

  // ===============================================================
  // PDF 2: Exit Ticket
  // ===============================================================
  {
    const doc = createPdf({ title: "Exit Ticket" });
    let y = addPdfHeader(doc, "Session 1 Exit Ticket", {
      subtitle: "One event, one date, one reason",
      color: C.ALERT,
      lessonInfo: "Year 6 Inquiry - Events that shaped Australia",
    });

    y = addTipBox(doc, "Show me you can name one event from your era, when it happened, and why it shaped Australia.", y, { color: C.ALERT });

    y = addSectionHeading(doc, "Event Name", y, { color: C.ALERT });
    y = addLinedArea(doc, y, 1, { lineSpacing: 24 });

    y += 4;
    y = addSectionHeading(doc, "Year It Happened", y, { color: C.ALERT });
    // Single short box for the year
    const yrX = PAGE.MARGIN;
    doc.save();
    doc.roundedRect(yrX, y, 160, 28, 4).lineWidth(0.9).strokeColor("#000000").stroke();
    doc.restore();
    y += 28 + 14;

    y = addSectionHeading(doc, "One Reason It Shaped Australia", y, { color: C.ALERT });
    y = addLinedArea(doc, y, 3, { lineSpacing: 24 });

    y += 10;
    y = addTipBox(doc, "Frame to help you start: '___ (year ___) shaped Australia because ___.'", y, { color: C.SUCCESS });

    addPdfFooter(doc, "Inquiry | Year 6 | Session 1 - Exit Ticket");
    const outPath = path.join(RES_DIR, "Session 1 Exit Ticket.pdf");
    await writePdf(doc, outPath);
    console.log("PDF written:", outPath);
  }

  console.log("\nBuild complete!");
  console.log("Output folder:", LESSON_FOLDER);
}

build().catch((err) => { console.error(err); process.exit(1); });
