"use strict";

/**
 * Shark Tank Day | Monday 7 December 2026 | Term 4 Week 10
 *
 * The finale of the Shark Tank unit (Sessions 1 to 8, builds/t4w*_shark_*.json).
 * The Grade 5 teams pitch and run stalls; the Grade 6s are the sharks and
 * invest tokens. Written as a script, not a lesson spec, because the day is an
 * event, not a lesson: there is no LI, practice round or exit ticket, so the
 * spec validator's lesson rules do not apply. Every slide uses theme builders.
 */

const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");
const { createTheme } = require("../themes/factory");
const { addResourceSlide, getSessionResourceFolder } = require("../themes/pdf_helpers");
const { buildResources } = require("../themes/lesson/resources");

const T = createTheme("inquiry", "grade56", 3);
const SESSION = 9;
const FOOTER = "Shark Tank | Grade 5 and 6 | Shark Tank Day, Monday 7 December";
const OUT_DIR = path.join(__dirname, "..", "output", "T4W10_SharkTank_Day");

const token = { text: "Shark Token", visual: { type: "pictogram", name: "coins" } };

const SPEC = {
  lesson: { title: "Shark Tank Day", meta: "Grade 5 and 6 | Shark Tank", footer: FOOTER },
  resources: [
    {
      kind: "page",
      label: "Shark Notes",
      description: "One per Grade 6 shark: tick the checks for each team before investing.",
      title: "Shark Notes",
      subtitle: "Tick each check you see. Then invest your tokens.",
      blocks: [
        { visual: { type: "pictograms", items: [{ name: "magnifier", label: "Real problem?" }, { name: "recycle", label: "Upcycled?" }, { name: "build", label: "Can they make it work?" }, { name: "people", label: "Will people want it?" }] } },
        { organiser: { left: "Team and product", right: "Checks I saw (tick up to 4)", rows: 12, rowH: 34 } },
        { tip: "Sharks invest. They don't buy. Five tokens: spread them or stack them, and keep your choices secret." },
      ],
    },
    {
      kind: "cards",
      label: "Shark Tokens",
      description: "Print and cut: five tokens per Grade 6 shark, 12 a page, so about 25 copies for 60 sharks.",
      cols: 3,
      showNameDate: false,
      cards: Array.from({ length: 12 }, () => ({ ...token, visual: { ...token.visual } })),
    },
  ],
};

const notes = (beats, prep, tag, answer) => T.composeGlanceNotes(answer ? { answer, beats, prep, tag } : { beats, prep, tag });

(async () => {
  fs.mkdirSync(path.join(OUT_DIR, getSessionResourceFolder(SESSION)), { recursive: true });
  const resourceItems = await buildResources(SPEC, T, OUT_DIR, SESSION);

  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  // No bullet points on any slide (megaprompt 85), as buildLesson does for specs.
  const addSlide = pres.addSlide.bind(pres);
  pres.addSlide = (...args) => { const sl = addSlide(...args); sl.__noBullets = true; return sl; };

  T.titleSlide(pres, "Shark Tank Day", "Monday 7 December | Term 4 Week 10", "Grade 5 and 6 | Shark Tank",
    "The finale. Grade 5 teams pitch; Grade 6 sharks invest their tokens.", {
      visual: { type: "pictograms", items: [
        { name: "magnifier", label: "Find a problem" },
        { name: "recycle", label: "Upcycle a fix" },
        { name: "megaphone", label: "Pitch it" },
      ] },
    });

  addResourceSlide(pres, {
    resources: resourceItems,
    manipulatives: [
      "One jar or cup per team, labelled with the team's product name",
      "Five Shark Tokens per Grade 6, cut before the day",
      "A timer the audience can see, for 90-second pitches",
    ],
    studentTools: [
      "Each team's prototype, poster with the price, cue cards and business plan",
      "One Shark Notes sheet and a pencil per Grade 6",
    ],
    boardSetup: [
      "Book the hall or gym: about 60 Grade 5s at stalls plus every Grade 6 will not fit in a classroom",
      "Stalls around the room: prototype front and centre, poster behind, jar at the front",
      "The pitch order written up, with every team's product name",
      "A camera or iPad for showcase photos, and the media consent list: photograph only students with consent",
    ],
  }, T, FOOTER, "Set up before the bell: stalls, jars, tokens, Shark Notes and the pitch order. Collect every business plan at the end; it is the assessment piece. Photograph only students with media consent; check the list before the day.");

  T.heroVisualSlide(pres, "Teacher run sheet", "How the day runs", {
    type: "table",
    fontSize: 17,
    rows: [
      ["When", "What happens"],
      ["Before", "Hall booked, stalls set up, jars out, tokens cut, consent list checked"],
      ["10 minutes", "Shark briefing for the Grade 6s"],
      ["About 45 minutes", "Pitches: 90 seconds each, then one shark question"],
      ["15 minutes", "Stall walk: sharks invest their tokens"],
      ["10 minutes", "Count the tokens, announce the winner"],
    ],
  }, notes([
    ["SAY to staff: one teacher keeps time, one runs the pitch order.", "A third counts tokens during the stall walk if you have one."],
    ["TIME: pitches take about 2 minutes per team with changeover.", "Twenty teams is about 45 minutes."],
    "Teacher-facing only. Skip it when the audience arrives.",
  ], "Run sheet for about 85 minutes. Adjust the pitch block to the number of teams.", "[Event | Planning made visible]"), FOOTER, { label: "Showcase photos: students with media consent only." });

  T.heroVisualSlide(pres, "Welcome", "Welcome to the Shark Tank", {
    type: "pictograms",
    items: [
      { name: "magnifier", label: "Find a problem" },
      { name: "recycle", label: "Upcycle a fix" },
      { name: "megaphone", label: "Pitch it" },
    ],
  }, notes([
    ["SAY: Welcome to the Shark Tank. Every product today fixes a real problem.", "And every one is made from something already used."],
    ["SAY: Grade 5s, you've worked on this since Week 2.", "Today you find out what the sharks think."],
    ["ASK: Grade 5s, our anchor.", "5 sec. Cue: Everyone, together, on three.", "EXPECT: Find a problem. Upcycle a fix. Pitch it."],
  ], "Whole audience. Grade 5 teams sit together with their prototypes.", "[Event | Welcome]", "Find a problem. Upcycle a fix. Pitch it."), FOOTER, { label: "Every product fixes a real problem, made from something already used." });

  T.heroVisualSlide(pres, "Shark briefing", "Sharks invest. They don't buy.", {
    type: "table",
    fontSize: 20,
    rows: [
      ["", "Buying", "Investing"],
      ["You give", "Money for one product", "Tokens to help the business grow"],
      ["You get", "The product to keep", "A share of what it earns"],
      ["Today", "Nobody buys", "Every shark invests"],
    ],
  }, notes([
    ["SAY: Grade 6s, you learned this in Session 2.", "You're not shopping today. You're investing."],
    ["ASK: Do you take a product home today?", "5 sec. Cue: Everyone, together, on three.", "EXPECT: no"],
    ["SAY: Invest in the business you believe can grow,", "not the thing you'd like to own."],
  ], "Shark briefing, part 1. The table is the same one the Grade 6s met in Session 2.", "[Event | Shark briefing]", "no - sharks invest, they don't buy"), FOOTER, { label: "Invest in the business you believe can grow." });

  T.heroVisualSlide(pres, "Shark briefing", "The four shark checks", {
    type: "pictograms",
    items: [
      { name: "magnifier", label: "Real problem?" },
      { name: "recycle", label: "Upcycled?" },
      { name: "build", label: "Can they make it work?" },
      { name: "people", label: "Will people want it?" },
    ],
  }, notes([
    ["SAY: Same four checks you used as consultant sharks.", "Tick what you see on your Shark Notes for every stall you visit."],
    ["ASK: Grade 6s, say the four checks.", "5 sec. Cue: Everyone, together, on three.", "EXPECT: problem, upcycled, make it work, people want it"],
    ["SAY: A great pitch helps, but the checks come first.", "Ask about the business, kind and fair. No 'I'm out'."],
  ], "Shark briefing, part 2. Hand out Shark Notes and pencils now.", "[Event | Shark briefing]", "Real problem? Upcycled? Can they make it work? Will people want it?"), FOOTER, { label: "Tick each check on your Shark Notes." });

  T.heroVisualSlide(pres, "Shark briefing", "Your five tokens", {
    type: "pictograms",
    items: [
      { name: "coins", label: "Five tokens each" },
      { name: "sort", label: "Spread or stack them" },
      { name: "lock", label: "Keep it secret" },
    ],
  }, notes([
    ["SAY: Five tokens each. Put them all on one team, or spread them.", "Drop them in the team's jar during the stall walk."],
    ["SAY: Keep your choices secret.", "No telling teams where your tokens went."],
    ["ASK: Thumbs: ready to invest?", "5 sec. Cue: thumbs at your chest... show me.", "EXPECT: thumbs up"],
  ], "Hand out five tokens per Grade 6. Spares stay with the token counter.", "[Event | Shark briefing]", "thumbs up - five tokens, spread or stack, kept secret"), FOOTER, { label: "Drop your tokens in the team's jar during the stall walk." });

  T.heroVisualSlide(pres, "Pitches", "Pitch time", {
    type: "pictograms",
    items: [
      { name: "microphone", label: "Pitch: 90 seconds" },
      { name: "question", label: "One shark question" },
      { name: "clap", label: "Applause, next team" },
    ],
  }, notes([
    ["SAY: Each team has 90 seconds, then one question from a shark.", "Clap every team on and off."],
    ["TIME each pitch on the visible timer.", "At 90 seconds, a gentle wave means finish your sentence."],
    "POINT to the next team on the pitch order. Keep changeovers under 30 seconds.",
  ], "Leave this slide up for the whole pitch block. Photos of every team, consent list only.", "[Event | Pitches]"), FOOTER, { label: "Teams pitch in the order on the board." });

  T.heroVisualSlide(pres, "Stall walk", "Investing time", {
    type: "pictograms",
    items: [
      { name: "walk", label: "Visit at least five stalls" },
      { name: "magnifier", label: "Run the checks" },
      { name: "coins", label: "Invest your tokens" },
    ],
  }, notes([
    ["SAY: Sharks, visit at least five stalls. Try the prototype. Ask one question.", "Then drop your tokens in the jars you choose."],
    ["SAY: Grade 5s, stay at your stall.", "Answer questions calmly, with your evidence."],
    "TIME 15 minutes. Collect the jars at the end and count away from the teams.",
  ], "Token count happens out of sight. Note the top three in case of a tie.", "[Event | Stall walk]"), FOOTER, { label: "Grade 5s stay at their stalls and answer questions." });

  T.heroVisualSlide(pres, "The results", "And the winner is...", {
    type: "pictograms",
    items: [{ name: "trophy", label: "The winning business" }],
  }, notes([
    ["SAY: The sharks have invested. Every team here built a real business.", "But one team earned the most tokens."],
    "SAY: Announce the winner, then the runners-up if you're naming them.",
    "SAY: Every team gets a round of applause. Then photos at the stalls.",
  ], "Read the winner from the token count. Photo of the team with its prototype, consent list only.", "[Event | Results]"), FOOTER, { label: "Most tokens invested by the Grade 6 sharks." });

  T.heroVisualSlide(pres, "Thank you", "Thank you, entrepreneurs", {
    type: "pictograms",
    items: [
      { name: "magnifier", label: "You found a problem" },
      { name: "recycle", label: "You upcycled a fix" },
      { name: "megaphone", label: "You pitched it" },
    ],
  }, notes([
    ["SAY: You found a problem, upcycled a fix and pitched it.", "That's what real entrepreneurs do."],
    "SAY: Grade 6s, thank you for being our sharks.",
    "COLLECT every business plan before the teams pack up. It's the assessment piece.",
  ], "Business plans collected for assessment. Photos to the showcase folder.", "[Event | Close]"), FOOTER, { label: "Business plans in before you pack up." });

  const fileName = path.join(OUT_DIR, "Shark Tank Day.pptx");
  await pres.writeFile({ fileName });
  console.log("PPTX written to " + fileName);
})().catch((err) => { console.error(err); process.exit(1); });
