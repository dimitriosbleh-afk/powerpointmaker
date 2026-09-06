# Lesson Spec Reference

A lesson spec is a JSON file that carries the CONTENT and INTENT of one
lesson. The pipeline makes every layout decision: which builder, what size,
where the visual sits, how the answer is revealed. You never write
coordinates.

```bash
node scripts/check_spec_notes.js builds/<name>.json     # lint the teacher notes first
node scripts/build_and_check.js builds/<name>.json      # build + the seven QA gates
python scripts/pptx_to_images.py output/<folder>/<deck>.pptx   # then LOOK at it
```

Golden exemplars (copy their shape, not their content):

| Spec | Band | Shows |
|---|---|---|
| `builds/exemplar_foundation_numeracy_making_10.json` | Foundation numeracy | one-question-per-slide review and fluency, hero ten frames, choice check, You Do with worksheet and answer key |
| `builds/exemplar_year2_literacy_feeling_clues.json` | Year 2 literacy | pictogram launch, word card, text extracts with highlights, choice check, scaffold page |
| `builds/exemplar_year56_science_water_cycle.json` | Year 5/6 science | two word cards, cycle diagram I Do and faded We Do, table rehearsal, hinge CFU, scaffold |

Validation is strict. Every problem is printed with the field path and the
fix. Warnings (`ADVISORY [spec] ...`) are work not yet done.

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
| `week` | recommended | 1-based; picks the palette variant. Every session of a unit uses the same week |
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
between), then for numeracy `dailyReview`... `fluency`..., then a launch,
then `li`, then `keyWord` cards if any, then the body, `exitTicket`,
`closing` last.

| kind | Required | Optional | Builds |
|---|---|---|---|
| `title` | | | cover with subject glyph or `lesson.titleVisual` |
| `overview` | `lines` | `title` | teacher-facing overview (multi-session decks) |
| `resources` | | | Teacher Resources from `resources` + `materials` |
| `dailyReview` | `title` | `prompts`, `visual`, `reveal` | numeracy review; a visual with no prompts fills the slide |
| `fluency` | `title`, `prompts` | `reveal` | one numeral or fact, hero-sized |
| `launch` | `title` | `lines`, `visual`, `label`, `prompt`, `reveal` | hero visual (no lines) or hero statement panel |
| `li` | `learningIntention`, `successCriteria` (exactly 3) | | LI and SC |
| `keyWord` | `word`, `meaning`, `pictogram` or `image` | `example`, `routine` | one word card with its picture |
| `heroVisual` | `badge`, `title`, `visual` | `label`, `prompt`, `badgeColor`, `reveal` | the representation IS the slide |
| `content` | `badge`, `title`, `lines` | `visual`, `badgeColor`, `reveal` | 1-3 short lines set as a hero panel; more lines as bullets; visual on the right |
| `workedExample` | `stage` (1-5), `title`, `steps` | `stageLabel`, `visual`, `reveal` | numeracy worked example with visual beside the steps |
| `choice` | `badge`, `title`, `options` (2-4) | `prompt`, `answer` (0-based), `letters` | Which one? cards; `answer` reveals a tick on click |
| `cfu` | `title`, `technique`, `question` | `badge`, `reveal` | text check with the CHECK stamp |
| `youDo` | `title`, `task` | `steps` (max 3), `where`, `visual`, `visualLabel`, `frame` | task hero, First/Next/Then chips, mini model, sentence frame |
| `textExtract` | `badge`, `title`, `extract` | `highlights`, `source`, `prompt`, `reveal` | exact text, marker-highlighted phrases |
| `cycle` | `title`, `centerLabel`, `steps` (3-4) | `badge`, `promptTitle`, `promptLines`, `reveal` | science loop; `steps[].icon` names a pictogram; `label: ""` fades a name and keeps the `detail` clue |
| `process` | `title`, `steps` (2-6) | `badge`, `promptTitle`, `promptLines` | science ordered flow |
| `boardBuild` | `title`, `directive` | `promptText`, `prefilledHints` | blank build canvas |
| `scenario` | `title`, `scenario`, `questions` | `badge` | wellbeing scenario |
| `pairShare` | `title`, `questions` | | discussion cards |
| `exitTicket` | `questions` (1-3) | `title`, `visual`, `label` | with `visual`, a hero visual plus the first question as the prompt |
| `closing` | `reflectionPrompt` | `selfAssessment`, `takeaways` | review and reflect; SC come from the `li` slide |

`badge` is the student-facing stage label (`"I Do"`, `"We Do"`, `"CFU"`,
`"Launch"`, `"Notice"`). Numeracy decks get `Stage n |` prefixed for I Do,
We Do and You Do automatically. `badgeColor` is one of `primary`
`secondary` `accent` `alert` `success` `assess`; it defaults sensibly from the
badge text (CFU red, We Do secondary, You Do success).

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
```

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

Title, resources and closing slides take a one-line string. Every teaching
slide takes a Glance object (megaprompt 45-47):

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
zone at most 120 words including the labels and numbers; every `ASK` carries
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
      { "organiser": { "left": "Stage", "right": "What happens", "rows": 4, "leftContent": ["1. Evaporation", "2.", "3.", "4."] } },
      { "box": 190, "label": "Draw the cycle as a loop." },
      { "steps": ["Read", "Underline", "Write"] },
      { "text": "I think ___ feels ___ because ___." },
      { "lines": 3 },
      { "tip": "Stretch: ..." }
    ]
  },
  { "kind": "cards", "label": "Feeling Cards", "cols": 2, "cards": [{ "text": "happy", "visual": { "type": "pictogram", "name": "happy" } }] }
]
```

Names come out session-first (`Session 1 Make 10 Worksheet.pdf`) and the
Teacher Resources slide links them. A `worksheet` gets an answer key
automatically (`answerKey: false` to skip). Item fields: `prompt`, `visual`,
`answerVisual` (the filled-in state for the key), `answer`, `answerLines`,
`answerLabel`, `box` (drawing box height), `lined`. Visuals on paper are
limited to the types with a paper twin (the validator names them).

Keep to zero or one printed resource unless the lesson genuinely needs more
(megaprompt 0a item 7).
