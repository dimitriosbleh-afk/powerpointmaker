# Lesson Spec Reference

A lesson spec is a JSON file that carries the CONTENT and INTENT of one
lesson. The pipeline makes every layout decision: which builder, what size,
where the visual sits, how the answer is revealed. You never write
coordinates.

```bash
node scripts/taught_log.js builds/<name>.json           # numeracy: what earlier lessons taught, for Daily Review
node scripts/check_spec_notes.js builds/<name>.json     # lint the teacher notes first
node scripts/build_and_check.js builds/<name>.json      # build + the seven QA gates
python scripts/pptx_to_images.py output/<folder>/<deck>.pptx   # then LOOK at it
```

Golden exemplars (copy their shape, not their content):

| Spec | Band | Shows |
|---|---|---|
| `builds/exemplar_foundation_numeracy_making_10.json` | Foundation numeracy | Daily Review with `from` keys, one-question-per-slide fluency, hero ten frames, diagnostic choice check, We Do practice round, You Do with worksheet and answer key |
| `builds/exemplar_year2_literacy_feeling_clues.json` | Year 2 literacy | pictogram launch, word card, text extracts with highlights, diagnostic choice check, We Do and You Do practice rounds, no printing |
| `builds/exemplar_year56_science_water_cycle.json` | Year 5/6 science | two word cards, cycle diagram I Do and faded We Do, table rehearsal, practice round of everyday cases, hinge CFU, standard Your turn, sectioned main sheet with Extension and supported sheets (megaprompt 85 outside maths) |
| `builds/exemplar_year56_numeracy_measuring_angles.json` | Year 5/6 numeracy | the maths planning team's rules (megaprompt 85): LI before launch, algorithm fluency with column sums, protractor diagrams, practice round, standard Your turn, sectioned main sheet with Extension and supported sheets |

Validation is strict. Every problem is printed with the field path and the
fix. Warnings (`ADVISORY [spec] ...`) are work not yet done.

For ordering and sequencing tasks, author the student items in a mixed order
and keep the solution separately in `notes.answer`, `reveal.answers` where
supported, and worksheet `items[].answer` / `answerVisual`. The builders
preserve the supplied item order; they do not shuffle it or detect an
already-solved task. Follow `IMPORTANT/MEGA_PROMPT.md` section 19a, including
its student-view QA, for slides, worksheet visuals and cut-out `cards`.

## Top level

```json
{
  "lesson":    { ... },          // required
  "materials": { ... },          // optional, feeds the Teacher Resources slide
  "resources": [ ... ],          // optional printed PDFs (default: none or one)
  "slides":    [ ... ]           // required, in teaching order
}
```

### lesson

| Field | Required | Notes |
|---|---|---|
| `subject` | yes | `literacy` `numeracy` `science` `inquiry` `wellbeing` |
| `yearLevel` | yes | `foundation` `grade1` `grade2` `grade34` `grade56` |
| `term` | yes | 1-4. With `week` and `session` it places the lesson in teaching order for the taught log |
| `week` | yes | 1-based; picks the palette variant. Every session of a unit uses the same week |
| `year` | | Four-digit year; defaults to the current year |
| `minutes` | | Session length, default 60. Sets the response floor (one per three minutes) |
| `plan` | yes | `{ curriculum, shape, criticalFeature, decisionPoints (2-4), anchor?, catchUp? }`: printed in the Teacher Resources notes under BEFORE TEACHING (megaprompt 84) |
| `variant` | alt | 0-5, overrides week |
| `session` | default 1 | Session number; names the resources folder and the `Session N` prefix |
| `title` | yes | Deck title |
| `subtitle`, `meta` | | Title slide lines. `meta` is the small pill, e.g. `"Year 2 Literacy | Reading"` |
| `footer` | | Defaults to `title | meta` |
| `outputFolder` | | Folder under `output/`; defaults to a slug of the title |
| `fileName` | | Defaults to `<title>.pptx` |
| `titleVisual` | | A visual spec for the cover instead of the subject glyph |

### materials

Arrays of short strings, each shown as a group on the Teacher Resources
slide: `manipulatives`, `studentTools`, `routineIcons`, `boardSetup`,
`videos`, `urls`, `ochre`. Name every manipulative the lesson uses.

## Slides

Every slide has `kind`, the fields for that kind, and `notes`. Unknown
fields are errors (they are almost always typos).

Fixed opening order (validated): `title`, `resources` (an `overview` may sit
between), then for numeracy `dailyReview`... `fluency`..., then `li`,
then the `launch`, then `keyWord` cards if any, then the body, `exitTicket`,
`closing` last.

| kind | Required | Optional | Builds |
|---|---|---|---|
| `title` | | | cover with subject glyph or `lesson.titleVisual` |
| `overview` | `lines` | `title` | teacher-facing overview (multi-session decks) |
| `resources` | | | Teacher Resources from `resources` + `materials` |
| `dailyReview` | `title`, `from` | `prompts`, `visual`, `reveal` | numeracy review; a visual with no prompts fills the slide. `from` names what it retrieves (see Taught log) |
| `fluency` | `title`, and `prompts` or `visual` | `reveal`, `label` | one numeral or fact, hero-sized; or a visual such as a column sum |
| `launch` | `title` | `lines`, `visual`, `label`, `prompt`, `reveal` | hero visual (no lines) or hero statement panel |
| `li` | `learningIntention`, `successCriteria` (exactly 3) | | LI and SC |
| `keyWord` | `word`, `meaning`, `pictogram` or `image` | `example`, `routine` | one word card with its picture |
| `heroVisual` | `badge`, `title`, `visual` | `label`, `prompt`, `link`, `badgeColor`, `reveal` | the representation IS the slide. `link` (a full http(s) URL) makes the `prompt` bar text clickable, e.g. a listen/watch slide that opens its video |
| `content` | `badge`, `title`, `lines` | `visual`, `badgeColor`, `reveal` | 1-3 short lines set as a hero panel; more lines as bullets; visual on the right |
| `workedExample` | `stage` (1-5), `title`, `steps` | `stageLabel`, `visual`, `reveal` | numeracy worked example with visual beside the steps |
| `choice` | `badge`, `title`, `options` (2-4) | `prompt`, `answer` (0-based), `letters` | Which one? cards; `answer` reveals a tick on click. With `answer`, every wrong option needs `misconception` |
| `cfu` | `title`, `technique`, `question` | `badge`, `reveal` | text check with the CHECK stamp |
| `practice` | `title`, `items` (3-8), `pivot` | `badge`, `routine`, `thinkTime`, `ask`, `followUp`, `badgeColor` | a practice round: one hero slide per item, answer on click, notes written for you (see Practice) |
| `youDo` | `title`, `task` | `steps` (max 3), `where`, `visual`, `visualLabel`, `frame`, `extendedTask` | task hero, First/Next/Then chips, mini model, sentence frame |
| `textExtract` | `badge`, `title`, `extract` | `highlights`, `source`, `prompt`, `reveal` | exact text, marker-highlighted phrases |
| `cycle` | `title`, `centerLabel`, `steps` (3-4) | `badge`, `promptTitle`, `promptLines`, `reveal` | science loop; `steps[].icon` names a pictogram; `label: ""` fades a name and keeps the `detail` clue |
| `process` | `title`, `steps` (2-6) | `badge`, `promptTitle`, `promptLines` | science ordered flow |
| `boardBuild` | `title`, `directive` | `promptText`, `prefilledHints` | blank build canvas |
| `scenario` | `title`, `scenario`, `questions` | `badge` | wellbeing scenario |
| `pairShare` | `title`, `questions` | | discussion cards |
| `exitTicket` | `questions` (1-3) | `title`, `visual`, `label` | with `visual`, a hero visual plus the first question as the prompt. Must be a new item, collected individually |
| `closing` | `reflectionPrompt` | `selfAssessment`, `takeaways` | review and reflect; SC come from the `li` slide |

`badge` is the student-facing stage label (`"I Do"`, `"We Do"`, `"CFU"`,
`"Launch"`, `"Notice"`). Numeracy decks get `Stage n |` prefixed for I Do,
We Do and You Do automatically. `badgeColor` is one of `primary`
`secondary` `accent` `alert` `success` `assess`; it defaults sensibly from the
badge text (CFU red, We Do secondary, You Do success).

### Checks and the exit ticket

A check is only decision-grade when every wrong answer tells the teacher
something (megaprompt 37). On a `choice` slide with an `answer`, each wrong
option names the misconception it catches:

```json
"options": [
  { "visual": { "type": "tensFrame", "filled": 3 }, "misconception": "counts the counters, not the empty boxes" },
  { "visual": { "type": "tensFrame", "filled": 7 } },
  { "visual": { "type": "tensFrame", "filled": 8 }, "misconception": "loses count of the empty boxes; touch each one" }
]
```

The build adds a `WRONG ANSWERS:` line to the slide's prep zone from these,
so leave room for it (the prep zone holds three lines).

The exit ticket is the lesson's evidence, so it must be a new item: a new
text, new numbers or a new context, never one already modelled, checked or
revealed (megaprompt 53). The validator fails an exit ticket that reuses a
character or place name from an earlier slide, repeats six or more words in
a row from one, repeats an earlier question with no new item, reuses a
modelled visual, collects its answer through partner talk or a choral
response, or leaves SC2 out of its tag. A `{ "type": "text" }` visual holds
a short new passage and sizes itself to fit.

### Practice

Every lesson meets two floors (megaprompt 82), and the build prints both
counts: planned whole-class responses, at least one per three minutes
(20 in 60 minutes), and independent items (Foundation 4, Years 1-2 6,
Years 3-6 8) from worksheet items or a practice round badged "You Do". A
You Do that is one extended task sets `youDo.extendedTask` instead.

A practice round is the main way there:

```json
{
  "kind": "practice",
  "badge": "We Do",
  "title": "How many more make 10?",
  "routine": "boards",
  "thinkTime": 8,
  "pivot": "count the empty boxes together on the board frame",
  "followUp": "cold call one board: how did you count?",
  "items": [
    { "visual": { "type": "tensFrame", "filled": 4 }, "label": "4 counters", "answer": "6 more. 4 and 6 make 10", "expect": "6" },
    { "visual": { "type": "tensFrame", "filled": 9 }, "label": "9 counters", "answer": "1 more. 9 and 1 make 10", "expect": "1" },
    { "visual": { "type": "tensFrame", "filled": 1 }, "label": "1 counter", "answer": "9 more. 1 and 9 make 10", "expect": "9" }
  ],
  "notes": {
    "say": ["Now it's your turn. Three quick frames,", "and everyone writes how many more on their board."],
    "trap": ["writing the counters seen.", "Fix: child touches each empty box, counts aloud, rewrites."],
    "stretch": "Say each pair both ways.",
    "help": "Real counters on a real frame.",
    "prep": "Guided practice from typical to edge cases. SC2.",
    "tag": "[We Do | Supported application | SC2 | HITS 3, 7]"
  }
}
```

Each item has exactly one of `visual`, `extract` (a short text, with
optional `source` and `highlights`) or `text` (one short line), plus
`answer` (shown on the answer bar) and optional `expect` (the EXPECT line,
default the answer), `say` (the reveal line) and `pivot`. `routine` is
`boards` (default), `fingers`, `point` or `choral`. The pipeline writes
each item's ASK, SCAN and REVEAL from the school cue scripts; the round's
`say` opens the first item, and EXTENSION and HELP sit on the second.

### Taught log

`build_and_check.js` records every lesson spec that passes all gates in
`records/taught_<yearLevel>_<subject>.json` (gitignored, one log per
machine; exemplars are never recorded; `TAUGHT_LOG_DIR` redirects it).
Entries are ordered by year, term, week and session. Each `dailyReview`
slide sets `from` to a log key (`"2026-T3-W8-S2"`), `"teacher"` when the
request named the focus, or `"before log"`. The validator checks keys
against the log: they must exist, come earlier, and at least one item
must reach back two weeks when the log has older learning (megaprompt 83).

```bash
node scripts/taught_log.js builds/<name>.json                     # earlier lessons by gap, with review items
node scripts/taught_log.js --list <yearLevel> <subject>
node scripts/taught_log.js --remove <yearLevel> <subject> <key>   # a lesson that was not taught
```

### Visual specs

Anywhere a `visual` is accepted:

```json
{ "type": "tensFrame", "filled": 7 }
{ "type": "fiveFrame", "filled": 3 }
{ "type": "doubleTensFrame", "filledTop": 10, "filledBottom": 8 }
{ "type": "dotCard", "count": 6 }                  { "type": "dotCards", "counts": [4, 6] }
{ "type": "numberTrack", "start": 1, "end": 10, "highlight": [7] }
{ "type": "numberLine", "start": 0, "end": 2, "step": 0.3333, "marked": [3] }
{ "type": "fractionStrips", "strips": [{ "denom": 4, "shaded": 3 }, { "denom": 4, "shaded": 0 }] }
{ "type": "array", "rows": 3, "cols": 4 }
{ "type": "baseTen", "hundreds": 1, "tens": 2, "ones": 3 }
{ "type": "groupedCounters", "groups": 3, "per": 4 }
{ "type": "ppwMat", "whole": 7, "partA": 4, "partB": null }
{ "type": "chips", "items": ["1/2", "3/4", "1/8"] }
{ "type": "pictogram", "name": "butterfly", "label": "butterfly" }
{ "type": "pictograms", "items": ["happy", "sad", "worried"], "labels": false }
{ "type": "text", "text": "9" }
{ "type": "table", "rows": [["Animal", "Legs"], ["Dog", "4"]] }
{ "type": "image", "path": "assets/unit/photo.jpg" }
{ "type": "angle", "rays": [0, 65], "protractor": true }
{ "type": "angle", "rays": [0, 130, 180], "arcs": [{ "from": 0, "to": 130, "label": "130°" }, { "from": 130, "to": 180, "label": "x" }] }
{ "type": "angle", "rays": [0, 90], "arcs": [{ "from": 0, "to": 90, "right": true }] }
{ "type": "columnSum", "numbers": [34567, 12345], "op": "+" }      op "-" or "×" too
{ "type": "shortDivision", "dividend": 4728, "divisor": 6 }        "quotient": "788", "remainder": 1 shows the answer
{ "type": "grid", "x": [0, 10], "y": [0, 10], "points": [{ "x": 3, "y": 4, "label": "A" }] }
{ "type": "grid", "x": [-5, 5], "y": [-5, 5], "polygon": [[1,1],[4,1],[4,3]], "arrows": [{ "from": [1,1], "to": [-3,-2] }] }
{ "type": "grid", "reference": true, "cols": 6, "rows": 5, "cells": [{ "col": "C", "row": 2, "label": "tree" }] }
{ "type": "chart", "style": "bar", "categories": ["Footy", "Netball"], "values": [8, 5], "title": "...", "xLabel": "...", "yLabel": "..." }
{ "type": "chart", "categories": ["Red", "Blue"], "series": [{ "name": "Expected", "values": [10, 10] }, { "name": "Observed", "values": [7, 13] }] }
{ "type": "spinner", "sectors": [{ "label": "red", "weight": 2 }, { "label": "blue" }], "pointer": 40 }
{ "type": "shape", "points": [[0,0],[8,0],[8,5],[0,5]], "labels": ["8 m", "5 m", null, null], "centerLabel": "A = ?" }
{ "type": "barModel", "parts": 4, "shaded": 1, "total": "$80", "labels": ["?", "", "", ""], "below": ["25%", "25%", "25%", "25%"] }
{ "type": "hundredGrid", "shaded": 35, "label": "35%" }
{ "type": "tally", "headers": ["Sport", "Tally", "Frequency"], "rows": [["Footy", 8], ["Netball", 5]] }   "counts": false leaves the frequency blank
{ "type": "fractionWall" }                                         halves to twelfths; "denoms": [1, 2, 4, 8], "shaded": { "4": 3 }, "labels": false
{ "type": "clock" }                                                blank face; "time": "3:45" draws the hands, "minutes": true, "digital": true
{ "type": "conversionChart", "measure": "length" }                 mass, capacity, time, or "units": ["m", "cm"], "factors": [100]
```

`grid`: coordinates sit on the lines (points, `polygon`/`polygons`, `path`,
dashed translation `arrows`, `showCoords`); a negative range draws four
quadrants; `reference: true` labels the spaces A, B, C and 1, 2, 3 instead.
`chart`: `style` is `bar` (default), `line` or `dot` (a dot plot, one dot per
count); `yStart` above 0 makes a misleading axis; `series` draws grouped bars
with a key; `showValues` prints each bar's value. `shape`: points in units
(x right, y up); `labels[i]` sits outside edge i; `grid: true` fills the shape
with unit squares; `splits` are dashed cut lines; `shapes: [...]` draws several
in one coordinate system; square corners get markers automatically.
`spinner`: sectors run clockwise from the top, and a colour word in the label
sets the fill. All of these are drawn once and look the same on slides and
paper.

`angle`: `rays` are directions in degrees, anticlockwise from pointing
right, from one vertex; `arcs` mark the angle anticlockwise from `from` to
`to` with an optional `label` (`right: true` draws the square marker);
`rotate` turns the figure; `protractor: true` overlays a protractor on the
first ray. Printed angles are true in degrees, so students can measure them
with a real protractor. On paper, `size` sets the width in points.

Pictogram names: `node -e 'console.log(require("./themes/factory").createTheme("science","grade2",0).listPictograms().join(" "))'`
or the sheet at the end of the Visual Catalogue. An unknown name fails the
build on purpose.

### reveal

```json
"reveal": { "answers": ["6 more. 4 and 6 make 10"] }
```

Adds a click-revealed answer bar under the content (megaprompt 20b). The
slide's own notes carry the `REVEAL after ...` beat. A slide with a reveal
cannot also have a `prompt` bar. For a genuinely different answer layout use
`"reveal": { "separate": true, "answers": [...], "notes": { "answer": "...", "beats": [...], "prep": "..." } }`
(a duplicate slide with its own post-reveal notes). On a `choice` slide use
`answer` instead.

### notes

Title and resources slides take a one-line string (the resources line is
followed automatically by the lesson plan). The closing takes a Glance
object: criteria read together, a self-assessment ASK, a RECORD beat that
reads it beside the exit evidence, and the reflection (megaprompt 52).
Every teaching slide takes a Glance object (megaprompt 45-47):

```json
"notes": {
  "answer": "4 more - 6 and 4 make 10",
  "beats": [
    ["POINT to the counters.", "SAY: Watch me. First I count the counters."],
    ["ASK: How many more make 10?", "8 sec. Cue: Write it... Chin it... Show me.", "EXPECT: 4"],
    ["SCAN boards, back row first.", "80%+ -> reveal, cold call: how did you count?", "Less -> count the empty boxes together, re-ask."],
    ["REVEAL after boards are scanned.", "SAY: Four more. Six and four make ten. Tick or fix."]
  ],
  "trap": ["writing 6, the counters seen.", "Fix: child touches each empty box, counts aloud, rewrites."],
  "stretch": "Say the pair both ways.",
  "help": "Child fills the empty boxes with real counters, then counts them.",
  "prep": "First guided try. SC2.",
  "tag": "[We Do | Supported application | SC2 | HITS 3, 7]"
}
```

Rules the linter and the build enforce: 2-5 beats; each string is one
physical line of at most 16 words (use an array for a multi-line beat); live
zone at most 150 words including the labels and numbers; every `ASK` carries
think time in seconds and one named routine (`Write it... Chin it... Show
me.`, `Everyone, together, on three`, `turn and tell`, `fingers up`,
`everyone points`); every `SCAN` has three lines: where to look, `80%+ ->`,
`Less -> ... re-ask`; `tag` is required. Omit `answer` when the slide asks
nothing. Run `check_spec_notes.js` until it prints "All notes within budget".

## Resources

```json
"resources": [
  {
    "kind": "worksheet",
    "label": "Make 10 Worksheet",
    "description": "You Do. Draw the counters that make each frame 10.",
    "title": "Make 10", "subtitle": "Draw counters to make 10.",
    "instructions": "Count the counters. Draw more until the frame shows 10.",
    "items": [
      { "prompt": "How many more make 10?", "visual": { "type": "tensFrame", "filled": 7 },
        "answerVisual": { "type": "tensFrame", "filled": 10 }, "answer": "3 more. 7 and 3 make 10.", "answerLabel": "more" }
    ],
    "tip": "Finished? Say each pair to a partner."
  },
  {
    "kind": "page",
    "label": "Water Cycle Scaffold",
    "description": "Enabling scaffold for the You Do.",
    "blocks": [
      { "heading": "The four stages" },
      { "visual": { "type": "pictograms", "items": [{ "name": "hot", "label": "evaporation" }] } },
      { "organiser": { "left": "Stage", "right": "What happens", "rows": 4, "rowH": 50, "leftContent": ["1. Evaporation", "2.", "3.", "4."] } },
      { "box": 190, "label": "Draw the cycle as a loop." },
      { "steps": ["Read", "Underline", "Write"] },
      { "checklist": ["Hook, context, position, preview", "A call to action"] },
      { "text": "I think ___ feels ___ because ___." },
      { "passage": "A text to read and highlight: larger type, wide line spacing." },
      { "lines": 3 },
      { "tip": "Extension: ..." }
    ]
  },
  { "kind": "cards", "label": "Feeling Cards", "cols": 2, "cards": [{ "text": "happy", "visual": { "type": "pictogram", "name": "happy" } }] },
  { "kind": "crossword", "label": "Persuasion Crossword", "wordBank": false,
    "words": [{ "answer": "rebuttal", "clue": "Answering the other side's argument" }, { "answer": "audience", "clue": "The people you are persuading" }] }
]
```

A `crossword` lays out its own grid from 4 to 16 `words` (`answer`, `clue`), longest first, each word crossing the others; a word that cannot cross is a validation error naming it. It gets an answer key automatically, and `wordBank: true` prints the answers as a bank for students who need it.

Names come out session-first (`Session 1 Make 10 Worksheet.pdf`) and the
Teacher Resources slide links them. A `worksheet` gets an answer key
automatically (`answerKey: false` to skip). Item fields: `prompt`, `visual`,
`answerVisual` (the filled-in state for the key), `answer`, `answerLabel`,
`kind` and the kind fields below. Visuals on paper are
limited to the types with a paper twin (the validator names them).

Keep to zero or one printed resource unless the lesson genuinely needs more
(megaprompt 0a item 7). Years 3-6 maths, literacy, science and inquiry are the exception: every session
carries a main sheet, an Extension and a supported sheet (megaprompt 85, 87).

### Sectioned worksheets

A worksheet may use `sections` instead of `items`. Each section prints a
coloured banner, an optional `intro`, a worked example with numbered steps,
then its questions, numbered across the sheet:

```json
{
  "kind": "worksheet", "role": "main", "label": "Worksheet",
  "title": "Measuring Angles", "subtitle": "Use your protractor.",
  "sections": [
    {
      "title": "Measure it", "proficiency": "fluency", "columns": 2,
      "example": { "prompt": "Measure the angle.", "visual": { "type": "angle", "rays": [0, 50] },
                   "steps": ["Put the centre on the corner.", "Line up 0 with one arm.", "Read the scale that starts at 0."], "answer": "50°" },
      "items": [ { "prompt": "Measure the angle.", "visual": { "type": "angle", "rays": [0, 35] }, "answer": "35°", "answerLabel": "Angle:" } ]
    }
  ]
}
```

`role` is `main`, `extension` or `supported`. `proficiency` (understanding,
fluency, problemSolving, reasoning) is internal and never printed. Any
item may carry `hint` (printed in italics under its prompt, for supported
sheets).

### Page layout and question kinds (megaprompt 86)

Every worksheet is laid out by `themes/lesson/worksheetLayout.js`. It fits
the sheet on one or two pages (two = one sheet printed double-sided, page 1
able to stand alone), balances the two pages, then gives every spare
millimetre back as squared working space. Three pages is a build error
unless the resource sets `maxPages` with a reason. Do not size working
space by hand; `workMin` (points) only sets a floor for one question.

`item.kind` picks the question format; each counts as work units towards
the practice floors:

| kind | fields | work units |
|---|---|---|
| `question` (default) | `prompt`, `visual`, `parts: [{prompt, answer, unit}]` (a/b/c on one card), `answer`, `answerLabel`, `unit`, `answerBox: false`, `lines: n` (ruled lines instead of squares), `working: "none"`, `workMin` | 1, or one per part |
| `table` | `columns`, `rows` (`""` = blank for the student), `answers` (same shape, for the key) | one per row |
| `sort` | `cards`, `groups`, `answer: { group: [cards] }` | one per two cards |
| `choice` | `options`, `answer` (index or indices to circle), `reason: true` adds "Because..." lines, `explain` (key) | 1 |
| `mistake` | `work` (the wrong working, one line per step), `answer` (the fix) | 2 |
| `open` | `answer` (a sample for the key), `workMin` | 2 |

A main sheet has at least 12 work units and at least three kinds; four
items in a row with the same kind and the same prompt is an error (a drill,
not a worksheet). Columns per section: `columns: 1`, `2` (default) or `3`. In a multi-column
section a `table` or `sort` takes a full-width row of its own; `span: "column"`
keeps one in its column, and `span: "full"` gives any item the full width.
