"use strict";

// Soccer Basics — Team Talk
// One-session coaching deck explaining the basics of soccer for a mixed team:
//   - includes complete beginners and players with several years of experience
//   - explains the pitch, aim, kick-off, throw-ins, corners, handballs, free kicks
//   - explains 4-4-2 positions and a simple "spread out, keep your shape" tactic
// Treated as a Wellbeing/PE introductory session so the theme system fits.

const path = require("path");
const fs = require("fs");
const PptxGenJS = require("pptxgenjs");
const { createTheme, weekToVariant } = require("../themes/factory");

// ── Theme ──
const T = createTheme("wellbeing", "grade34", weekToVariant(1)); // Forest Haven
const {
  C, FONT_H, FONT_B,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  addTopBar, addBadge, addTitle, addCard, addFooter, addTextOnShape,
  CONTENT_TOP, SAFE_BOTTOM, SLIDE_W, SLIDE_H,
  runSlideDiagnostics,
} = T;

// ── Output paths ──
const FOLDER_NAME = "Soccer_Basics_Team_Talk";
const LESSON_FOLDER = path.join(__dirname, "..", "output", FOLDER_NAME);
const PPTX_NAME = "Soccer Basics - Team Talk.pptx";
const FOOTER = "Soccer Basics | Team Talk";

fs.mkdirSync(LESSON_FOLDER, { recursive: true });

// ── Pitch colours (literal — independent of theme palette) ──
const GRASS = "5BAE6A";
const LINE = "FFFFFF";
const OWN_TEAM = "1B3F94";
const OPP_TEAM = "C94030";
const GK_GOLD = "FFB81C";
const BALL = "FFFFFF";

// Row x-fractions for the 4-4-2 mini pitch on slides 8-10. (Slide 7 uses a wider pitch with its own values.)
const ROW_X_FRACTIONS = [0.18, 0.40, 0.60, 0.82];
const ATT_X_FRACTIONS = [0.38, 0.62];
const FADED = "BBBBBB";

// ═══════════════════════════════════════════════════════════════
// Drawing helpers — kept local to this build
// ═══════════════════════════════════════════════════════════════

// Draw a top-down soccer pitch (length horizontal). Returns useful coords.
function drawPitchHorizontal(slide, x, y, w, h) {
  slide.addShape("rect", {
    x, y, w, h,
    fill: { color: GRASS },
    line: { color: LINE, width: 1.5 },
  });
  slide.addShape("line", {
    x: x + w / 2, y: y, w: 0, h,
    line: { color: LINE, width: 1.5 },
  });
  // Centre circle — roundRect with grass fill renders as a hollow outline
  // (LibreOffice oval rendering bug, per MEMORY.md)
  const cr = Math.min(w, h) * 0.13;
  slide.addShape("roundRect", {
    x: x + w / 2 - cr, y: y + h / 2 - cr, w: cr * 2, h: cr * 2, rectRadius: cr,
    fill: { color: GRASS },
    line: { color: LINE, width: 1.2 },
  });
  slide.addShape("roundRect", {
    x: x + w / 2 - 0.04, y: y + h / 2 - 0.04, w: 0.08, h: 0.08, rectRadius: 0.04,
    fill: { color: LINE },
  });
  const paW = w * 0.16;
  const paH = h * 0.55;
  slide.addShape("rect", {
    x, y: y + (h - paH) / 2, w: paW, h: paH,
    fill: { color: GRASS },
    line: { color: LINE, width: 1.2 },
  });
  slide.addShape("rect", {
    x: x + w - paW, y: y + (h - paH) / 2, w: paW, h: paH,
    fill: { color: GRASS },
    line: { color: LINE, width: 1.2 },
  });
  const sixW = paW * 0.4;
  const sixH = paH * 0.55;
  slide.addShape("rect", {
    x, y: y + (h - sixH) / 2, w: sixW, h: sixH,
    fill: { color: GRASS },
    line: { color: LINE, width: 1.0 },
  });
  slide.addShape("rect", {
    x: x + w - sixW, y: y + (h - sixH) / 2, w: sixW, h: sixH,
    fill: { color: GRASS },
    line: { color: LINE, width: 1.0 },
  });
  return { x, y, w, h, cx: x + w / 2, cy: y + h / 2 };
}

// Draw a top-down soccer pitch (length vertical, own goal at bottom).
function drawPitchVertical(slide, x, y, w, h) {
  slide.addShape("rect", {
    x, y, w, h,
    fill: { color: GRASS },
    line: { color: LINE, width: 1.5 },
  });
  // Halfway line (horizontal)
  slide.addShape("line", {
    x, y: y + h / 2, w, h: 0,
    line: { color: LINE, width: 1.5 },
  });
  const cr = Math.min(w, h) * 0.10;
  slide.addShape("roundRect", {
    x: x + w / 2 - cr, y: y + h / 2 - cr, w: cr * 2, h: cr * 2, rectRadius: cr,
    fill: { color: GRASS },
    line: { color: LINE, width: 1.2 },
  });
  slide.addShape("roundRect", {
    x: x + w / 2 - 0.04, y: y + h / 2 - 0.04, w: 0.08, h: 0.08, rectRadius: 0.04,
    fill: { color: LINE },
  });
  const paW = w * 0.55;
  const paH = h * 0.16;
  // Top penalty area (opponent goal)
  slide.addShape("rect", {
    x: x + (w - paW) / 2, y, w: paW, h: paH,
    fill: { color: GRASS },
    line: { color: LINE, width: 1.2 },
  });
  // Bottom penalty area (own goal)
  slide.addShape("rect", {
    x: x + (w - paW) / 2, y: y + h - paH, w: paW, h: paH,
    fill: { color: GRASS },
    line: { color: LINE, width: 1.2 },
  });
  return { x, y, w, h, cx: x + w / 2, cy: y + h / 2 };
}

function drawPlayer(slide, cx, cy, color, label, size = 0.34) {
  const r = size / 2;
  slide.addShape("roundRect", {
    x: cx - r, y: cy - r, w: size, h: size, rectRadius: r,
    fill: { color },
    line: { color: LINE, width: 1.2 },
  });
  if (label != null) {
    slide.addText(String(label), {
      x: cx - r, y: cy - r, w: size, h: size,
      fontSize: 11, fontFace: FONT_B, color: LINE,
      bold: true, align: "center", valign: "middle", margin: 0,
    });
  }
}

function drawBall(slide, cx, cy, size = 0.22) {
  const r = size / 2;
  slide.addShape("roundRect", {
    x: cx - r, y: cy - r, w: size, h: size, rectRadius: r,
    fill: { color: BALL },
    line: { color: "2C2C2C", width: 0.75 },
  });
}

// Right-column 4-4-2 mini pitch used on the defenders / midfielders / attackers slides.
// `highlight` ∈ {"def","mid","att"} draws that row in `highlightColor`, fading the others.
function drawPositionPitch(slide, layoutGuide, opts) {
  const { highlight, highlightColor, caption } = opts;
  const rx = layoutGuide.rightX;
  const rw = layoutGuide.rightW;
  const topY = layoutGuide.panelTopPadded;
  const pitchH = 2.8;
  const py = topY + 0.05;
  const pw = rw * 0.85;
  const px = rx + (rw - pw) / 2;

  drawPitchVertical(slide, px, py, pw, pitchH);
  drawPlayer(slide, px + pw / 2, py + pitchH - 0.24, GK_GOLD, "GK", 0.28);

  const defY = py + pitchH * 0.72;
  const midY = py + pitchH * 0.48;
  const attY = py + pitchH * 0.22;
  const defColor = highlight === "def" ? highlightColor : FADED;
  const midColor = highlight === "mid" ? highlightColor : FADED;
  const attColor = highlight === "att" ? highlightColor : FADED;
  const defSize = highlight === "def" ? 0.30 : 0.26;
  const midSize = highlight === "mid" ? 0.30 : 0.26;
  const attSize = highlight === "att" ? 0.32 : 0.26;

  ROW_X_FRACTIONS.forEach((f) => drawPlayer(slide, px + pw * f, defY, defColor, "D", defSize));
  ROW_X_FRACTIONS.forEach((f) => drawPlayer(slide, px + pw * f, midY, midColor, "M", midSize));
  ATT_X_FRACTIONS.forEach((f) => drawPlayer(slide, px + pw * f, attY, attColor, "A", attSize));

  slide.addText(caption, {
    x: rx, y: py + pitchH + 0.10, w: rw, h: 0.30,
    fontSize: 12, fontFace: FONT_B, color: highlightColor,
    bold: true, align: "center", margin: 0,
  });
}

// ═══════════════════════════════════════════════════════════════
// Teacher Notes
// ═══════════════════════════════════════════════════════════════

const NOTES_TITLE = [
  "SAY:",
  "- Welcome team. Today we are going to talk through how the game works so we are all on the same page",
  "- Some of you have played for years, some of you are brand new. Today is about the basics so we all start together",
  "- We will cover the pitch, the aim, positions, throw-ins, corners, kick-offs, handballs, and how we spread out",
  "",
  "DO:",
  "- Have the team sitting where they can see the screen",
  "- Keep it conversational. Pause for questions between sections",
  "",
  "TEACHER NOTES:",
  "Single team-talk session for a mixed-experience boys' and girls' squad. Tone is supportive and beginner-safe. Experienced players go deeper through the prompts and the closing quick check.",
  "",
  "WATCH FOR:",
  "- Players who already know lots of this -- invite them to add detail when prompted",
  "- Players who are quiet or unsure -- normalise asking questions",
  "",
  "[Wellbeing/PE: Title | Establishing Purpose]",
].join("\n");

const NOTES_LI = [
  "SAY:",
  "- Today we are learning the basic rules and team shape so we can play together as a 4-4-2",
  "- By the end you should be able to show me your spot, take a throw-in, and know what to do at a corner or free kick",
  "",
  "DO:",
  "- Read the learning intention and each I can statement out loud together",
  "- Quick thumbs check: who feels confident with each one already?",
  "",
  "TEACHER NOTES:",
  "First criterion is achievable for everyone. The second covers the main set pieces. The third asks players to apply team shape during play -- this is the stretch.",
  "",
  "WATCH FOR:",
  "- Players who put thumbs down on the first one -- this is the priority for today",
  "",
  "[Wellbeing/PE: LI/SC]",
].join("\n");

const NOTES_PITCH = [
  "SAY:",
  "- This is what the pitch looks like from above",
  "- The thick line down the middle is the halfway line. Your half is the half you defend",
  "- The circle in the middle is the centre circle. Kick-offs happen from the spot in the middle",
  "- The big boxes at each end are the penalty area. Inside that box, if the defending team fouls the attacker, the attacker gets a penalty kick",
  "- The small box inside is the goal box. Goal kicks are taken from inside this box",
  "- The lines on the long sides are the touchlines. When the ball goes over a touchline it is a throw-in",
  "- The lines at each end are the by-lines or goal lines. Over the by-line gives a corner or a goal kick",
  "",
  "DO:",
  "- Point to each part as you name it",
  "- Ask: which line gives a throw-in? Which gives a corner? [touchline = throw-in; by-line = corner or goal kick]",
  "",
  "TEACHER NOTES:",
  "Most players will know some of this. Make sure everyone hears the words touchline, by-line, halfway line and penalty area at least once -- these are the words you will use during games.",
  "",
  "WATCH FOR:",
  "- Players mixing up touchline and by-line -- they often do. Re-point at each",
  "",
  "[Wellbeing/PE: Concept]",
].join("\n");

const NOTES_AIM = [
  "SAY:",
  "- The aim of the game is simple. Score more goals than the other team",
  "- A goal counts when the whole ball crosses the line between the goalposts and under the crossbar",
  "- You score in the goal you are attacking. That is the goal at the opposite end to your goalkeeper",
  "- Every player except the goalkeeper kicks the ball. The keeper can use their hands inside their own penalty area",
  "",
  "DO:",
  "- Use the diagram to point at the attacking goal and the defending goal",
  "- Ask: where can the keeper use their hands? [Inside their own penalty area only]",
  "",
  "TEACHER NOTES:",
  "Beginner players sometimes try to kick toward whichever goal is closer. Make the direction clear: we attack one way in the first half and we swap at half-time.",
  "",
  "WATCH FOR:",
  "- Players unsure which way they are attacking on game day -- remind them at the start of each half",
  "",
  "[Wellbeing/PE: Concept]",
].join("\n");

const NOTES_KICKOFF = [
  "SAY:",
  "- The game starts with a kick-off from the centre spot",
  "- One team takes the kick-off at the start. The other team takes it at the start of the second half",
  "- After a goal is scored the team that conceded takes the next kick-off",
  "- The ball can go forward, sideways or backwards at kick-off. The other team has to stay outside the centre circle until the ball is touched",
  "",
  "DO:",
  "- Point at the centre spot and the centre circle on the diagram",
  "- Ask: where do the other team stand at kick-off? [Outside the centre circle, in their own half]",
  "",
  "TEACHER NOTES:",
  "The rule that kick-offs can go in any direction is recent. Older players may remember it had to go forward. Either way, just touch the ball to start play.",
  "",
  "WATCH FOR:",
  "- Players from the non-kicking team drifting into the circle early -- ref will make them retake it",
  "",
  "[Wellbeing/PE: Set Piece]",
].join("\n");

const NOTES_FORMATION = [
  "SAY:",
  "- We are playing a 4-4-2. That means four defenders, four midfielders, two attackers, plus the goalkeeper",
  "- This is one of the most common shapes in soccer. It gives you cover at the back, control in the middle, and two players up front to score",
  "- Each line spreads across the pitch. Defenders cover the back, midfielders link the play, attackers stay higher up",
  "- Your job is to stay roughly in your line. You can move out of it, but you have to come back",
  "",
  "DO:",
  "- Point at each row of players on the diagram",
  "- Count out loud with the team: 4 at the back, 4 in the middle, 2 up front",
  "- Ask everyone to say their position out loud",
  "",
  "TEACHER NOTES:",
  "This is the visual anchor for the rest of the session. Players will refer back to it when we talk about spreading out and keeping shape. Keep the formation diagram up while you discuss roles in the next three slides.",
  "",
  "WATCH FOR:",
  "- Players who do not know which line they are in -- write names on a printed copy after the talk",
  "",
  "[Wellbeing/PE: I Do | Formation]",
].join("\n");

const NOTES_DEFENDERS = [
  "SAY:",
  "- Defenders are the back four. Your job is to stop the other team from scoring",
  "- Two centre-backs in the middle stay close to the keeper. Two full-backs play out wide, one on each side",
  "- Things to do: stay between the ball and your own goal, mark the attacker closest to you, clear the ball away from danger",
  "- Things to avoid: getting pulled out of position when you do not need to, dribbling near your own goal",
  "",
  "DO:",
  "- Point at the back four on the diagram on the right",
  "- Ask the defenders to identify themselves",
  "",
  "TEACHER NOTES:",
  "Defenders win games by being calm. Tell them: first job is to stop the goal, second job is to start the next attack. Keep wording simple for new players.",
  "",
  "WATCH FOR:",
  "- New defenders running upfield with the ball -- praise the intent but remind them to pass it forward",
  "",
  "[Wellbeing/PE: Position]",
].join("\n");

const NOTES_MIDS = [
  "SAY:",
  "- Midfielders are the four in the middle. You are the engine of the team",
  "- You link the defenders and the attackers. You run the most",
  "- Two of you stay a bit deeper to help defend. Two of you push up to help attack",
  "- Things to do: support the ball, pass forward when you can, get back when we lose it",
  "- Things to avoid: standing still, all four of you on the same side of the pitch",
  "",
  "DO:",
  "- Point at the middle four on the diagram",
  "- Ask the midfielders to identify themselves",
  "",
  "TEACHER NOTES:",
  "Midfielders shape the game. Help them see they have two jobs: attack support and defensive cover. The 'two deeper, two higher' picture is the easiest version to remember.",
  "",
  "WATCH FOR:",
  "- All four mids chasing the ball at once -- prompt: who is staying back?",
  "",
  "[Wellbeing/PE: Position]",
].join("\n");

const NOTES_ATTACKERS = [
  "SAY:",
  "- Attackers are the two at the top of the diagram. Your job is to score and to create chances",
  "- Stay high up the pitch when we have the ball. Be ready to run in behind the other team's defenders",
  "- When we lose the ball, do not give up. Press their defenders so they cannot pass easily",
  "- Things to do: stay onside, make runs into space, shoot when you get the chance",
  "- Things to avoid: dropping all the way back to defend, both of you in the same spot",
  "",
  "DO:",
  "- Point at the two attackers on the diagram",
  "- Ask the attackers to identify themselves",
  "",
  "TEACHER NOTES:",
  "Offside is when an attacker is closer to the opponent's goal line than both the ball and the second-last defender when the ball is played. Explain only as deep as your players need it. For beginners: stand level with the last defender until the pass is made.",
  "",
  "WATCH FOR:",
  "- Attackers crowding the same channel -- prompt: one to each side",
  "",
  "[Wellbeing/PE: Position]",
].join("\n");

const NOTES_THROWIN = [
  "SAY:",
  "- A throw-in happens when the ball goes fully over the touchline -- the long line on the side",
  "- The team that did not touch it last gets the throw",
  "- How to take one: stand behind the line, ball in both hands, bring it over and behind your head, throw with both hands, keep both feet on the ground",
  "- If you lift a foot, drop the ball, or throw it with one hand, the ref gives the throw to the other team",
  "",
  "DO:",
  "- Demonstrate the action: ball behind head, both hands, both feet planted",
  "- Ask one player to model it without a ball",
  "",
  "TEACHER NOTES:",
  "The foul-throw rule is the most common mistake new players make. Run a quick practice at training: line them up and have each player take a throw.",
  "",
  "WATCH FOR:",
  "- Players lifting the back foot when they throw -- the most common foul throw",
  "- Players throwing one-handed -- not allowed",
  "",
  "[Wellbeing/PE: Set Piece]",
].join("\n");

const NOTES_CORNER = [
  "SAY:",
  "- A corner kick happens when a defender is the last person to touch the ball before it crosses their own by-line -- the line at each end of the pitch",
  "- The attacking team takes the kick from the corner of the pitch on the side where it went out",
  "- The ball is placed in the small corner arc. The taker passes or crosses into the box",
  "- Defenders stand at least 9.15 metres away until the ball is kicked. That is about 10 big steps",
  "",
  "DO:",
  "- Point at the corner arc on the diagram",
  "- Ask: who takes the corner -- the team that kicked it out, or the other team? [The other team. The team that touched it last loses possession]",
  "",
  "TEACHER NOTES:",
  "Corners are a great chance to score. Remind attackers to attack the ball at the near post, far post, and the penalty spot. Remind defenders to mark a player or a space and clear hard.",
  "",
  "WATCH FOR:",
  "- Defenders standing too close to the kicker -- they have to be 9.15m back",
  "",
  "[Wellbeing/PE: Set Piece]",
].join("\n");

const NOTES_HANDBALL = [
  "SAY:",
  "- A handball is when an outfield player touches the ball with their hand or arm on purpose, or with their arm in an unnatural position",
  "- The other team gets a free kick from where the handball happened",
  "- If the handball happens inside the defending team's own penalty area, it is a penalty kick",
  "- Keepers can use their hands -- but only inside their own penalty area",
  "- For any free kick, the defending players must stand at least 9.15 metres back from the ball -- the same 10 big steps as a corner",
  "",
  "DO:",
  "- Hold one arm out to show what 'unnatural position' looks like",
  "- Pace out 9.15 metres on the floor so players can see how far back they have to stand",
  "",
  "TEACHER NOTES:",
  "Accidental ball-on-hand with arms tucked in is usually not a handball. Arm above shoulder, or arm out to make yourself bigger, usually is. Keep it simple for new players: hands by your sides if you can.",
  "",
  "WATCH FOR:",
  "- Players raising their arms to block a shot -- this is what gets called",
  "",
  "[Wellbeing/PE: Rule]",
].join("\n");

const NOTES_SPREAD = [
  "SAY:",
  "- The biggest beginner mistake in soccer is everyone running to the ball at once",
  "- We call that bunching. When you bunch, you give the other team space to play around you",
  "- Instead we want to keep our shape. That means staying in your line and your side of the pitch",
  "- If the ball is on the left, one or two of you go to help. The rest hold their spots",
  "- When we have the ball, spread out wide so we have options to pass",
  "- When we lose the ball, get back into your shape before you chase",
  "",
  "DO:",
  "- Point to the bunched diagram and the spread diagram side by side",
  "- Ask: which one is going to win the game? [The spread one]",
  "",
  "TEACHER NOTES:",
  "Spreading out is the most important tactic at this level. Practise it at training with a small-sided game and a rule: only the closest two players go to the ball.",
  "",
  "WATCH FOR:",
  "- Players who keep drifting back to the ball -- redirect to their spot",
  "",
  "[Wellbeing/PE: Tactic]",
].join("\n");

const NOTES_CFU = [
  "SAY:",
  "- Quick check before we wrap up",
  "- I will read each question. Hands up if you know the answer. Whole team can respond on the last one",
  "",
  "DO:",
  "- Read each question. Take answers from different players",
  "- Q1: What is a throw-in for? [Ball over the touchline]",
  "- Q2: How far back do defenders stand for a corner or free kick? [9.15 metres or about 10 big steps]",
  "- Q3: Where is the only place the goalkeeper can use their hands? [Inside their own penalty area]",
  "- Q4: What is the biggest mistake we want to avoid? [Bunching -- everyone chasing the ball]",
  "",
  "TEACHER NOTES:",
  "Use this as a quick informal check, not a test. The questions cover the four ideas most likely to matter on game day.",
  "",
  "WATCH FOR:",
  "- Players who can answer Q1 and Q3 but not Q2 and Q4 -- those are the ones to reinforce at training",
  "",
  "[Wellbeing/PE: CFU]",
].join("\n");

const NOTES_CLOSING = [
  "SAY:",
  "- Quick recap. Three things we want to remember every game",
  "- Know your spot in the 4-4-2",
  "- Spread out -- do not bunch on the ball",
  "- Throw-in: both feet on the ground, both hands behind your head",
  "- One more thing: have fun out there",
  "",
  "DO:",
  "- Read the three I can statements. Players show thumbs up, sideways or down",
  "- Finish by asking each player to say their position out loud one more time",
  "",
  "TEACHER NOTES:",
  "Close on a positive. The mixed-experience squad will respond better to encouragement than to drill-style correction at this stage. Save deeper tactics for training.",
  "",
  "WATCH FOR:",
  "- Any thumbs-down on the I can statements -- chat to those players individually after",
  "",
  "[Wellbeing/PE: Closing]",
].join("\n");

const NOTES_MATERIALS = [
  "SAY:",
  "- Here is what we will need today and at training this week",
  "",
  "DO:",
  "- Show the screen as a checklist",
  "- Have a soccer ball handy to demonstrate the throw-in and the kick-off",
  "",
  "TEACHER NOTES:",
  "No printed resources required. If you want, print one copy of the 4-4-2 diagram from slide 6 to write player names against each position after the talk.",
  "",
  "WATCH FOR:",
  "- Anyone who has not brought shin pads -- remind them for game day",
  "",
  "[Wellbeing/PE: Materials]",
].join("\n");

// ═══════════════════════════════════════════════════════════════
// Build
// ═══════════════════════════════════════════════════════════════

async function build() {
  const pres = new PptxGenJS();
  pres.layout = "LAYOUT_16x9";

  // ─── SLIDE 1: Title ─────────────────────────────────────────────
  titleSlide(
    pres,
    "Soccer Basics",
    "Our Team Talk",
    "How the game works | 4-4-2 | Set pieces | Team shape",
    NOTES_TITLE
  );

  // ─── SLIDE 2: Materials / What you need ─────────────────────────
  contentSlide(
    pres,
    "Today",
    C.SECONDARY,
    "What We Need",
    [
      "Soccer ball for the throw-in and kick-off demo",
      "Shin pads and boots for training and game day",
      "Drink bottle",
      "A spot to sit where you can see the screen",
      "Bring any questions -- there are no silly ones",
    ],
    NOTES_MATERIALS,
    FOOTER
  );

  // ─── SLIDE 3: LI/SC ─────────────────────────────────────────────
  liSlide(
    pres,
    ["We are learning the basic rules and team shape so we can play together as a 4-4-2"],
    [
      "I can name my position and show where I stand",
      "I can explain throw-ins, kick-offs, corners and handballs",
      "I can spread out and keep my shape during a game",
    ],
    NOTES_LI,
    FOOTER
  );

  // ─── SLIDE 4: The Pitch ─────────────────────────────────────────
  contentSlide(
    pres,
    "Concept",
    C.PRIMARY,
    "The Pitch",
    [
      "Halfway line splits the field in two",
      "Centre circle -- kick-offs happen here",
      "Penalty area -- big box at each end",
      "Goal box -- small box inside the penalty area",
      "Touchline -- long side; over it is a throw-in",
      "By-line -- short end; over it is a corner or goal kick",
    ],
    NOTES_PITCH,
    FOOTER,
    (slide, layoutGuide) => {
      const rx = layoutGuide.rightX;
      const rw = layoutGuide.rightW;
      const topY = layoutGuide.panelTopPadded;
      // Pitch panel
      const pitchH = 2.6;
      drawPitchHorizontal(slide, rx, topY + 0.05, rw, pitchH);
      // Caption labels
      slide.addText("Touchline", {
        x: rx, y: topY + pitchH + 0.10, w: rw, h: 0.28,
        fontSize: 11, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
        align: "center", margin: 0,
      });
      slide.addText("By-line   |   Halfway line   |   Centre spot", {
        x: rx, y: topY + pitchH + 0.38, w: rw, h: 0.28,
        fontSize: 10, fontFace: FONT_B, color: C.MUTED,
        align: "center", margin: 0,
      });
    }
  );

  // ─── SLIDE 5: Aim of the Game ───────────────────────────────────
  contentSlide(
    pres,
    "Aim",
    C.PRIMARY,
    "The Aim of the Game",
    [
      "Score more goals than the other team",
      "A goal counts when the WHOLE ball crosses the line",
      "Attack the goal at the OPPOSITE end to your keeper",
      "Swap ends at half-time -- now you attack the other way",
      "Only the goalkeeper can use their hands -- inside their own penalty area",
    ],
    NOTES_AIM,
    FOOTER,
    (slide, layoutGuide) => {
      const rx = layoutGuide.rightX;
      const rw = layoutGuide.rightW;
      const topY = layoutGuide.panelTopPadded;
      const pitchH = 2.4;
      const py = topY + 0.10;
      drawPitchHorizontal(slide, rx, py, rw, pitchH);
      // Arrow / attack labels
      slide.addText("Defend", {
        x: rx + 0.05, y: py + pitchH / 2 - 0.15, w: rw * 0.25, h: 0.3,
        fontSize: 11, fontFace: FONT_B, color: OWN_TEAM,
        bold: true, align: "center", margin: 0,
      });
      slide.addText("Attack -->", {
        x: rx + rw * 0.55, y: py + pitchH / 2 - 0.15, w: rw * 0.42, h: 0.3,
        fontSize: 11, fontFace: FONT_B, color: OPP_TEAM,
        bold: true, align: "center", margin: 0,
      });
      // Ball on centre spot
      drawBall(slide, rx + rw / 2, py + pitchH / 2);
      // Caption
      slide.addText("We attack one way, then swap at half-time", {
        x: rx, y: py + pitchH + 0.12, w: rw, h: 0.30,
        fontSize: 10, fontFace: FONT_B, color: C.MUTED,
        italic: true, align: "center", margin: 0,
      });
    }
  );

  // ─── SLIDE 6: Kick-Off ──────────────────────────────────────────
  contentSlide(
    pres,
    "Set Piece",
    C.ACCENT,
    "Kick-Off",
    [
      "The game starts with a kick-off from the centre spot",
      "One team kicks off at the start, the other team in the second half",
      "After a goal, the team that conceded kicks off",
      "The ball can go any direction once it is touched",
      "The other team must stay OUTSIDE the centre circle until it is touched",
    ],
    NOTES_KICKOFF,
    FOOTER,
    (slide, layoutGuide) => {
      const rx = layoutGuide.rightX;
      const rw = layoutGuide.rightW;
      const topY = layoutGuide.panelTopPadded;
      const pitchH = 2.5;
      const py = topY + 0.10;
      drawPitchHorizontal(slide, rx, py, rw, pitchH);
      drawBall(slide, rx + rw / 2, py + pitchH / 2);
      // Two kicking players on the centre circle
      drawPlayer(slide, rx + rw / 2 - 0.20, py + pitchH / 2 + 0.05, OWN_TEAM);
      drawPlayer(slide, rx + rw / 2 + 0.20, py + pitchH / 2 - 0.05, OWN_TEAM);
      slide.addText("Centre spot", {
        x: rx, y: py + pitchH + 0.12, w: rw, h: 0.28,
        fontSize: 11, fontFace: FONT_B, color: C.CHARCOAL,
        italic: true, align: "center", margin: 0,
      });
    }
  );

  // ─── SLIDE 7: Our 4-4-2 Formation (HERO custom slide) ───────────
  {
    const slide = pres.addSlide();
    addTopBar(slide, C.PRIMARY);
    addBadge(slide, "Formation", { color: C.PRIMARY });
    addTitle(slide, "Our 4-4-2: How We Line Up");

    // Vertical pitch centred
    const pitchX = 2.4, pitchY = CONTENT_TOP, pitchW = 3.4, pitchH = 3.8;
    drawPitchVertical(slide, pitchX, pitchY, pitchW, pitchH);

    const gkY = pitchY + pitchH - 0.32;
    drawPlayer(slide, pitchX + pitchW / 2, gkY, GK_GOLD, "GK", 0.34);

    // Formation rows — wider spread than the position slides because the pitch is larger
    const FORMATION_ROW_FRACS = [0.14, 0.38, 0.62, 0.86];
    const FORMATION_ATT_FRACS = [0.36, 0.64];
    const defY = pitchY + pitchH * 0.72;
    FORMATION_ROW_FRACS.forEach((f) => drawPlayer(slide, pitchX + pitchW * f, defY, OWN_TEAM, "D", 0.34));
    const midY = pitchY + pitchH * 0.50;
    FORMATION_ROW_FRACS.forEach((f) => drawPlayer(slide, pitchX + pitchW * f, midY, OWN_TEAM, "M", 0.34));
    const attY = pitchY + pitchH * 0.22;
    FORMATION_ATT_FRACS.forEach((f) => drawPlayer(slide, pitchX + pitchW * f, attY, OWN_TEAM, "A", 0.34));

    // Labels — left side
    const labelsX = 0.5;
    const labelsW = 1.7;
    const rowLabel = (title, color, count, y) => {
      slide.addShape("roundRect", {
        x: labelsX, y, w: labelsW, h: 0.7, rectRadius: 0.08,
        fill: { color },
      });
      slide.addText(title, {
        x: labelsX, y: y + 0.04, w: labelsW, h: 0.32,
        fontSize: 15, fontFace: FONT_H, color: "FFFFFF",
        bold: true, align: "center", valign: "middle", margin: 0,
      });
      slide.addText(count, {
        x: labelsX, y: y + 0.40, w: labelsW, h: 0.26,
        fontSize: 11, fontFace: FONT_B, color: "FFFFFF",
        align: "center", valign: "middle", margin: 0,
      });
    };
    rowLabel("Attackers", OPP_TEAM, "x 2 -- score and create", pitchY + pitchH * 0.22 - 0.35);
    rowLabel("Midfielders", C.PRIMARY, "x 4 -- the engine room", pitchY + pitchH * 0.50 - 0.35);
    rowLabel("Defenders", OWN_TEAM, "x 4 -- back four + GK", pitchY + pitchH * 0.72 - 0.35);
    rowLabel("Goalkeeper", GK_GOLD, "x 1 -- hands inside box", pitchY + pitchH - 0.65);

    // Right side legend / arrows showing attacking direction
    slide.addText("Attacking direction", {
      x: 6.0, y: pitchY + 0.05, w: 3.5, h: 0.30,
      fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL,
      bold: true, align: "left", margin: 0,
    });
    slide.addText("Opponent goal", {
      x: 6.0, y: pitchY + 0.35, w: 3.5, h: 0.30,
      fontSize: 11, fontFace: FONT_B, color: OPP_TEAM,
      italic: true, align: "left", margin: 0,
    });
    // Big up arrow
    slide.addShape("upArrow", {
      x: 6.5, y: pitchY + 0.75, w: 0.5, h: 1.6,
      fill: { color: C.SECONDARY },
      line: { color: C.PRIMARY, width: 0.5 },
    });
    slide.addText("Our team\nattacks UP", {
      x: 7.1, y: pitchY + 0.95, w: 2.4, h: 0.9,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL,
      bold: true, valign: "middle", margin: 0,
    });
    slide.addText("Our goalkeeper\ndefends the bottom", {
      x: 6.5, y: pitchY + 2.55, w: 3.0, h: 0.8,
      fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL,
      valign: "top", margin: 0,
    });

    addFooter(slide, FOOTER);
    slide.addNotes(NOTES_FORMATION);
    runSlideDiagnostics(slide, pres, { respectSafeBottom: false });
  }

  // ─── SLIDE 8: Defenders ─────────────────────────────────────────
  contentSlide(
    pres,
    "Position",
    OWN_TEAM,
    "Defenders (x 4)",
    [
      "Two centre-backs in the middle -- stay close to the keeper",
      "Two full-backs out wide -- one left, one right",
      "DO: stay between the ball and your own goal",
      "DO: mark the closest attacker, clear the ball away from danger",
      "AVOID: dribbling near your own goal",
      "AVOID: getting pulled out of position when you do not need to",
    ],
    NOTES_DEFENDERS,
    FOOTER,
    (slide, layoutGuide) => drawPositionPitch(slide, layoutGuide, {
      highlight: "def", highlightColor: OWN_TEAM, caption: "The back four",
    })
  );

  // ─── SLIDE 9: Midfielders ───────────────────────────────────────
  contentSlide(
    pres,
    "Position",
    C.PRIMARY,
    "Midfielders (x 4)",
    [
      "You are the engine room -- you run the most",
      "Two deeper midfielders help defend",
      "Two higher midfielders push up to attack",
      "DO: support the ball, pass forward when you can",
      "DO: get back into shape when we lose the ball",
      "AVOID: standing still or all four on the same side",
    ],
    NOTES_MIDS,
    FOOTER,
    (slide, layoutGuide) => drawPositionPitch(slide, layoutGuide, {
      highlight: "mid", highlightColor: C.PRIMARY, caption: "The middle four",
    })
  );

  // ─── SLIDE 10: Attackers ────────────────────────────────────────
  contentSlide(
    pres,
    "Position",
    OPP_TEAM,
    "Attackers (x 2)",
    [
      "Stay high up the pitch when we have the ball",
      "Be ready to run in behind the other team's defenders",
      "When we lose it, press their defenders -- do not give up",
      "DO: stay onside, make runs into space, shoot when you can",
      "AVOID: dropping all the way back to defend",
      "AVOID: both of you in the same channel",
    ],
    NOTES_ATTACKERS,
    FOOTER,
    (slide, layoutGuide) => drawPositionPitch(slide, layoutGuide, {
      highlight: "att", highlightColor: OPP_TEAM, caption: "The front two",
    })
  );

  // ─── SLIDE 11: Throw-ins ────────────────────────────────────────
  contentSlide(
    pres,
    "Set Piece",
    C.ACCENT,
    "Throw-Ins",
    [
      "When? Ball goes over the touchline (long side)",
      "The team that did not touch it last gets the throw",
      "Stand BEHIND the line",
      "Ball in BOTH hands, bring it OVER and BEHIND your head",
      "Throw with BOTH hands -- keep BOTH FEET on the ground",
      "Lift a foot or use one hand = foul throw, ball to the other team",
    ],
    NOTES_THROWIN,
    FOOTER,
    (slide, layoutGuide) => {
      const rx = layoutGuide.rightX;
      const rw = layoutGuide.rightW;
      const topY = layoutGuide.panelTopPadded;

      // Card showing the throw-in stick-figure
      addCard(slide, rx, topY, rw, 3.0, { fill: C.BG_CARD, strip: C.ACCENT });
      slide.addText("How to throw", {
        x: rx + 0.15, y: topY + 0.10, w: rw - 0.3, h: 0.30,
        fontSize: 13, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0,
      });

      // Simple stick figure
      const figX = rx + rw / 2;
      const headY = topY + 0.60;
      // Head
      slide.addShape("roundRect", {
        x: figX - 0.18, y: headY, w: 0.36, h: 0.36, rectRadius: 0.18,
        fill: { color: C.CHARCOAL }, line: { color: C.CHARCOAL, width: 0.5 },
      });
      // Body
      slide.addShape("line", {
        x: figX, y: headY + 0.36, w: 0, h: 0.95,
        line: { color: C.CHARCOAL, width: 4 },
      });
      // Arms over head, holding ball
      slide.addShape("line", {
        x: figX - 0.38, y: headY - 0.10, w: 0.76, h: 0,
        line: { color: C.CHARCOAL, width: 4 },
      });
      slide.addShape("line", {
        x: figX - 0.38, y: headY - 0.10, w: 0, h: 0.25,
        line: { color: C.CHARCOAL, width: 4 },
      });
      slide.addShape("line", {
        x: figX + 0.38, y: headY - 0.10, w: 0, h: 0.25,
        line: { color: C.CHARCOAL, width: 4 },
      });
      // Ball above head
      drawBall(slide, figX, headY - 0.18, 0.30);
      // Legs
      slide.addShape("line", {
        x: figX, y: headY + 1.31, w: -0.28, h: 0.55,
        line: { color: C.CHARCOAL, width: 4 },
      });
      slide.addShape("line", {
        x: figX, y: headY + 1.31, w: 0.28, h: 0.55,
        line: { color: C.CHARCOAL, width: 4 },
      });
      // Ground line
      slide.addShape("line", {
        x: rx + 0.3, y: headY + 1.92, w: rw - 0.6, h: 0,
        line: { color: C.SECONDARY, width: 2 },
      });
      // Feet labels
      slide.addText("Both feet planted -- behind the line", {
        x: rx + 0.15, y: headY + 1.96, w: rw - 0.3, h: 0.28,
        fontSize: 10, fontFace: FONT_B, color: C.MUTED,
        italic: true, align: "center", margin: 0,
      });
    }
  );

  // ─── SLIDE 12: Corner Kicks ─────────────────────────────────────
  contentSlide(
    pres,
    "Set Piece",
    C.ACCENT,
    "Corner Kicks",
    [
      "When? A defender is last to touch the ball before it crosses their own by-line",
      "The ATTACKING team takes the kick from the corner",
      "Place the ball inside the small corner arc",
      "Defenders stand at least 9.15 metres back (~10 big steps) until the ball is kicked",
      "Pass or cross the ball into the box -- good chance to score",
    ],
    NOTES_CORNER,
    FOOTER,
    (slide, layoutGuide) => {
      const rx = layoutGuide.rightX;
      const rw = layoutGuide.rightW;
      const topY = layoutGuide.panelTopPadded;
      const pitchH = 2.6;
      const py = topY + 0.05;
      drawPitchHorizontal(slide, rx, py, rw, pitchH);
      // Mark corner arc (top-right corner of pitch, attacking direction)
      const cornerX = rx + rw;
      const cornerY = py;
      // Corner arc indicator — small circle outline just inside the corner
      slide.addShape("roundRect", {
        x: cornerX - 0.30, y: cornerY - 0.02, w: 0.30, h: 0.30, rectRadius: 0.15,
        fill: { color: GRASS },
        line: { color: "FFFFFF", width: 1.2 },
      });
      // Ball at the corner
      drawBall(slide, cornerX - 0.14, cornerY + 0.14, 0.16);
      // Call-out label pointing at the corner
      slide.addText("Corner taken from here  →", {
        x: cornerX - 2.0, y: cornerY + 0.05, w: 1.85, h: 0.30,
        fontSize: 11, fontFace: FONT_B, color: "FFFFFF",
        bold: true, align: "right", margin: 0,
      });
      // 9.15m note
      slide.addText("Defenders 9.15m back", {
        x: rx, y: py + pitchH + 0.10, w: rw, h: 0.30,
        fontSize: 11, fontFace: FONT_B, color: C.CHARCOAL,
        italic: true, align: "center", margin: 0,
      });
    }
  );

  // ─── SLIDE 13: Handballs & Free Kicks ───────────────────────────
  contentSlide(
    pres,
    "Rule",
    C.ALERT,
    "Handballs and Free Kicks",
    [
      "Handball = touching the ball with your hand or arm on purpose",
      "Also given if your arm is in an unnatural position (arm out, above shoulder)",
      "Other team gets a free kick from where it happened",
      "Handball INSIDE your own penalty area = a penalty kick",
      "Defenders must stand 9.15 metres back from every free kick (~10 big steps)",
      "Keepers can use hands -- but only inside their own penalty area",
    ],
    NOTES_HANDBALL,
    FOOTER,
    (slide, layoutGuide) => {
      const rx = layoutGuide.rightX;
      const rw = layoutGuide.rightW;
      const topY = layoutGuide.panelTopPadded;
      const cardH = 3.0;
      addCard(slide, rx, topY, rw, cardH, { fill: C.BG_CARD, strip: C.ALERT });
      slide.addText("Free kick: the wall", {
        x: rx + 0.15, y: topY + 0.10, w: rw - 0.30, h: 0.30,
        fontSize: 13, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0,
      });

      // Draw small free-kick diagram inside the card
      const innerX = rx + 0.20;
      const innerY = topY + 0.55;
      const innerW = rw - 0.40;
      const innerH = 1.85;
      // Mini pitch strip
      slide.addShape("rect", {
        x: innerX, y: innerY, w: innerW, h: innerH,
        fill: { color: GRASS },
        line: { color: LINE, width: 1 },
      });
      // Ball on left
      const ballX = innerX + 0.30;
      const ballY = innerY + innerH / 2;
      drawBall(slide, ballX, ballY, 0.22);
      // Attacker behind ball
      drawPlayer(slide, ballX - 0.20, ballY + 0.05, OWN_TEAM, null, 0.26);
      // Wall of defenders (3) to the right
      const wallX = innerX + 1.50;
      const wallY = innerY + 0.30;
      [0, 1, 2].forEach((i) => {
        drawPlayer(slide, wallX, wallY + i * 0.45, OPP_TEAM, null, 0.26);
      });
      // 9.15m dashed line
      slide.addShape("line", {
        x: ballX + 0.15, y: innerY + 0.10, w: 0, h: innerH - 0.20,
        line: { color: "FFFFFF", width: 0.8, dashType: "dash" },
      });
      slide.addShape("line", {
        x: wallX - 0.15, y: innerY + 0.10, w: 0, h: innerH - 0.20,
        line: { color: "FFFFFF", width: 0.8, dashType: "dash" },
      });
      // Label
      slide.addText("9.15 m  (~10 big steps)", {
        x: ballX + 0.20, y: ballY - 0.16, w: wallX - ballX - 0.40, h: 0.32,
        fontSize: 10, fontFace: FONT_B, color: "FFFFFF",
        bold: true, align: "center", valign: "middle", margin: 0,
      });
      // Sub-caption
      slide.addText("Ball                                           Wall", {
        x: innerX, y: innerY + innerH + 0.05, w: innerW, h: 0.28,
        fontSize: 10, fontFace: FONT_B, color: C.MUTED,
        italic: true, align: "center", margin: 0,
      });
    }
  );

  // ─── SLIDE 14: Spreading Out ────────────────────────────────────
  contentSlide(
    pres,
    "Tactic",
    C.SUCCESS,
    "Spread Out -- Keep Our Shape",
    [
      "Biggest beginner mistake: everyone chases the ball at once",
      "Bunching gives the other team space around us",
      "Instead: stay in your line and your side of the pitch",
      "When WE have the ball, spread wide so passes are open",
      "When WE lose the ball, get back into shape before you chase",
      "Rule of thumb: only the closest two go to the ball",
    ],
    NOTES_SPREAD,
    FOOTER,
    (slide, layoutGuide) => {
      const rx = layoutGuide.rightX;
      const rw = layoutGuide.rightW;
      const topY = layoutGuide.panelTopPadded;
      const halfW = (rw - 0.15) / 2;
      const pitchH = 1.45;

      // LEFT mini: bunched
      const bx = rx, by = topY + 0.10;
      drawPitchVertical(slide, bx, by, halfW, pitchH);
      // Ball with all players around it (bunched)
      const bcx = bx + halfW / 2;
      const bcy = by + pitchH * 0.55;
      drawBall(slide, bcx, bcy, 0.16);
      const bunched = [
        [bcx - 0.20, bcy - 0.15], [bcx + 0.18, bcy - 0.12],
        [bcx - 0.12, bcy + 0.15], [bcx + 0.20, bcy + 0.10],
        [bcx, bcy - 0.30], [bcx, bcy + 0.30],
      ];
      bunched.forEach(([x, y]) => drawPlayer(slide, x, y, OWN_TEAM, null, 0.20));
      slide.addText("BUNCHED", {
        x: bx, y: by + pitchH + 0.08, w: halfW, h: 0.28,
        fontSize: 12, fontFace: FONT_B, color: C.ALERT,
        bold: true, align: "center", margin: 0,
      });
      slide.addText("(don't do this)", {
        x: bx, y: by + pitchH + 0.34, w: halfW, h: 0.24,
        fontSize: 9, fontFace: FONT_B, color: C.MUTED,
        italic: true, align: "center", margin: 0,
      });

      // RIGHT mini: spread
      const sx = rx + halfW + 0.15, sy = topY + 0.10;
      drawPitchVertical(slide, sx, sy, halfW, pitchH);
      // Spread players
      const sw = halfW;
      const sh = pitchH;
      const spread = [
        [sx + sw * 0.18, sy + sh * 0.75], [sx + sw * 0.40, sy + sh * 0.78],
        [sx + sw * 0.60, sy + sh * 0.78], [sx + sw * 0.82, sy + sh * 0.75],
        [sx + sw * 0.50, sy + sh * 0.50],
        [sx + sw * 0.35, sy + sh * 0.25], [sx + sw * 0.65, sy + sh * 0.25],
      ];
      spread.forEach(([x, y]) => drawPlayer(slide, x, y, OWN_TEAM, null, 0.20));
      drawBall(slide, sx + sw * 0.50, sy + sh * 0.50, 0.16);
      slide.addText("SPREAD", {
        x: sx, y: sy + pitchH + 0.08, w: halfW, h: 0.28,
        fontSize: 12, fontFace: FONT_B, color: C.SUCCESS,
        bold: true, align: "center", margin: 0,
      });
      slide.addText("(keep your shape)", {
        x: sx, y: sy + pitchH + 0.34, w: halfW, h: 0.24,
        fontSize: 9, fontFace: FONT_B, color: C.MUTED,
        italic: true, align: "center", margin: 0,
      });

      // Big take-home
      addTextOnShape(slide, "Only the closest two go to the ball", {
        x: rx, y: topY + 2.35, w: rw, h: 0.42, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, {
        fontSize: 13, fontFace: FONT_B, color: "FFFFFF",
        bold: true, align: "center", valign: "middle", margin: 0,
      });
    }
  );

  // ─── SLIDE 15: Quick Check (CFU) ────────────────────────────────
  cfuSlide(
    pres,
    "Check",
    "Quick Team Check",
    "Hands Up",
    "1. What is a throw-in for?\n2. How far back do defenders stand for a corner or free kick?\n3. Where can the goalkeeper use their hands?\n4. What is the biggest mistake we want to avoid?",
    NOTES_CFU,
    FOOTER
  );

  // ─── SLIDE 16: Closing ──────────────────────────────────────────
  closingSlide(
    pres,
    {
      reflectionPrompt: "Three things every game: Know your spot. Spread out. Both feet for throw-ins. Have fun out there.",
      scItems: [
        "I can name my position and show where I stand",
        "I can explain throw-ins, kick-offs, corners and handballs",
        "I can spread out and keep my shape during a game",
      ],
      selfAssessment: {
        prompt: "Thumbs up, sideways, or down for each one.",
        options: ["Got it", "Getting there", "Need more"],
      },
    },
    NOTES_CLOSING
  );

  // ── Write PPTX ──
  const pptxPath = path.join(LESSON_FOLDER, PPTX_NAME);
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to", pptxPath);
}

build().catch((err) => {
  console.error("Build failed:", err);
  process.exit(1);
});
