# CLAUDE.md

Behavioral guidelines to reduce common LLM coding mistakes. Merge with project-specific instructions as needed.

**Tradeoff:** These guidelines bias toward caution over speed. For trivial tasks, use judgment.

## 1. Think Before Coding

**Don't assume. Don't hide confusion. Surface tradeoffs.**

Before implementing:
- State your assumptions explicitly. If uncertain, ask.
- If multiple interpretations exist, present them - don't pick silently.
- If a simpler approach exists, say so. Push back when warranted.
- If something is unclear, stop. Name what's confusing. Ask.

## 2. Simplicity First

**Minimum code that solves the problem. Nothing speculative.**

- No features beyond what was asked.
- No abstractions for single-use code.
- No "flexibility" or "configurability" that wasn't requested.
- No error handling for impossible scenarios.
- If you write 200 lines and it could be 50, rewrite it.

Ask yourself: "Would a senior engineer say this is overcomplicated?" If yes, simplify.

## 3. Surgical Changes

**Touch only what you must. Clean up only your own mess.**

When editing existing code:
- Don't "improve" adjacent code, comments, or formatting.
- Don't refactor things that aren't broken.
- Match existing style, even if you'd do it differently.
- If you notice unrelated dead code, mention it - don't delete it.

When your changes create orphans:
- Remove imports/variables/functions that YOUR changes made unused.
- Don't remove pre-existing dead code unless asked.

The test: Every changed line should trace directly to the user's request.

## 4. Goal-Driven Execution

**Define success criteria. Loop until verified.**

Transform tasks into verifiable goals:
- "Add validation" → "Write tests for invalid inputs, then make them pass"
- "Fix the bug" → "Write a test that reproduces it, then make it pass"
- "Refactor X" → "Ensure tests pass before and after"

For multi-step tasks, state a brief plan:
```
1. [Step] → verify: [check]
2. [Step] → verify: [check]
3. [Step] → verify: [check]
```

Strong success criteria let you loop independently. Weak criteria ("make it work") require constant clarification.

---

**These guidelines are working if:** fewer unnecessary changes in diffs, fewer rewrites due to overcomplication, and clarifying questions come before implementation rather than after mistakes.

# PPTX Lesson Generator

Node.js project using PptxGenJS to generate explicit teaching slide decks with companion PDF resources. Theme system provides 150 pre-built palettes across 5 subjects x 5 year levels x 6 weekly variants.

## Commands

```bash
node scripts/build_and_check.js builds/build_<unit>_lesson<n>.js  # Build + enforce 7 QA gates (see MEGA_PROMPT 59a)
node builds/build_<unit>_lesson<n>.js          # Build only (no automated checks)
node tests/test_theme.js <subject> <level> [variant] # Test a theme combo
python -m markitdown output/<file>.pptx        # Content QA - check text, order, typos (manual)
python scripts/check_lesson_quality.py output/<file>.pptx --profile literacy-60  # Lesson density/language QA
python scripts/pptx_to_images.py output/<file>.pptx  # Optional local preview - slides to slidetemp/*.jpg
python scripts/slide_montage.py                # Optional contact sheet from slidetemp/
python scripts/pptx_to_images.py --clean       # Delete slidetemp/ after QA
```

## Project Layout

```
themes/factory.js          # createTheme(subject, yearLevel, variant) - single entry point
themes/core/               # Shared utilities (layout, contrast, icons, shadows, elements, withReveal)
themes/builders/           # Slide builders by subject (base, literacy, numeracy, inquiry, wellbeing, science)
themes/palettes/           # Pure colour data (30 palettes per subject)
themes/pdf_helpers.js      # PDF resource generation (pdfkit)
builds/                    # One build script per lesson - writes to output/<LessonFolder>/
_archive/                  # Archived lesson scripts - historical only, not active exemplars
output/                    # Per-lesson folders (PPTX + companion PDFs)
IMPORTANT/MEGA_PROMPT.md   # Pedagogical framework - paste into conversation when planning lessons
docs/                      # Deep reference docs (read when needed, not every session)
```

**Never append below the `===== END OF MEGA-PROMPT. SHIFT CLICK HERE. =====` marker in `IMPORTANT/MEGA_PROMPT.md`.** Teachers select from the top of the file to that marker to copy the prompt. New sections go ABOVE it; the marker and the USER REQUEST block stay last.

## Theme API

```javascript
const { createTheme, weekToVariant } = require("../themes/factory");
const T = createTheme("literacy", "grade56", weekToVariant(3));
// T contains: C, FONT_H, FONT_B, all builders, withReveal, addCard, addFooter,
//   addTextOnShape, iconToBase64Png, getContrastColor, makeShadow, layout constants
```

Subjects: `literacy` | `numeracy` | `inquiry` | `wellbeing` | `science`
Year levels: `foundation` | `grade1` | `grade2` | `grade34` | `grade56`
Variants: `0`-`5` (use `weekToVariant(weekNumber)` for 1-based weeks)

**Design language lives in the theme, not the build script.** Palettes are retuned so every role colour is the brightest shade of its hue that still clears a per-band contrast floor (Foundation reads bright, Year 5/6 deeper); strong colours appear only on badges, pills, chips, reveals and ticks, while hero panels use the derived soft tints `C.PRIMARY_SOFT` / `C.ALERT_SOFT` / etc. (`T.softOf(hex)` for any colour). Title and closing slides carry one subject glyph, never decorative blobs. `addCard` has three variants (`white`, `tint`, `outline`). `contentSlide`, `cfuSlide` and `exitTicketSlide` set short content hero-sized automatically. Do not fight this with raw `addShape`/`addText`; if a look is wrong, fix the shared layer. To change palette hues, edit `themes/palettes/*.js` then run `node scripts/retune_palettes.js` to re-establish the floors. Read `docs/theme-system.md` "Design Language".

**Theme selection:** The explicit `Subject:` field in the user's prompt is authoritative for theme selection. Do not silently swap to a different theme because the content looks cross-curricular. If the content and subject seem misaligned, keep the theme aligned to the stated subject unless the user explicitly changes it.

**Theme cohesion: All lessons in the same unit MUST use the same variant.** Switching palettes between lessons in a unit looks confusing and unprofessional. Pick one variant for the unit (typically based on the week number) and use it for every lesson. Different variants are for different weeks or different units, not different lessons within the same unit.

**Subject builder overrides:** Subject builders (e.g. `createLiteracyBuilders`) must NOT re-export a stale copy of a base builder. If a subject builder shadows a base builder name (e.g. `annotatedModelSlide`), it must add genuine subject-specific behaviour that the base version cannot provide. If the override only duplicates the base logic — or is a frozen fork missing later improvements — remove it so the base version is used. The factory spreads `...subjectBuilders` after `...base`, so any name collision silently replaces the base version.

For builder signatures, palette schema, and full API: read `docs/theme-system.md`.

## PptxGenJS Rules

- NEVER use `#` in hex colours - causes file corruption. Use `"FF0000"` not `"#FF0000"`.
- NEVER encode opacity in hex strings (8-char like `"00000020"`). Use `opacity` property.
- NEVER reuse option objects across calls - PptxGenJS mutates in-place. Use factory functions.
- Every PptxGenJS text run must use a plain string `text` value. Never pass arrays, numbers, booleans, or objects as `text`; coerce dynamic values with `String(...)` before `addText()` or any builder/helper call that emits text runs.
- Use `bullet: true`, never unicode bullet characters (creates double bullets).
- Use `breakLine: true` between text array items.
- Use `addTextOnShape` instead of separate addShape + addText - validates contrast automatically.
- Set `margin: 0` on text boxes that must align precisely with shapes or icons.
- `rectRadius` only works with `ROUNDED_RECTANGLE`, not `RECTANGLE`.
- Avoid `lineSpacing` with bullets - causes excessive gaps; use `paraSpaceAfter` instead.
- Shadow `offset` must be non-negative - negative values corrupt the file.
- Hyperlinks must be passed in RUN options (`addText([{ text, options: { hyperlink } }], boxOpts)`), never at the `addText` options level. An options-level `hyperlink` makes PptxGenJS emit a shape-level `hlinkClick` as well as the run-level one, so the whole text box becomes clickable, not just the text. Include `color` in the same run options so the link keeps the theme colour.
- NEVER use em dashes, en dashes, smart quotes, ellipsis characters, or `--` anywhere in generated output: slide faces, notes, and PDFs. Use `-`, straight quotes, and `...`. The theme auto-converts these on `addText` (installSlideTextPatch) and `build_and_check.js` Gate 3 fails the build if any survive, but write clean strings at the source, especially for PDFs, which are not auto-sanitised.
- NEVER lay out slide text with runs of spaces (e.g. `"1/5    3/4    1/8"` or `"42        12        40"`). Gate 3 flags 3+ consecutive spaces. Use separate text elements, chips (`addTextOnShape` per item), `breakLine` runs, or `addRevealAnswerBar` with an array of answers (it joins with a visible `|` separator).
- Each presentation needs a fresh `new pptxgen()` instance.
- Write files with `await pres.writeFile({ fileName })`, never the deprecated `writeFile("path.pptx")` form.
- Always set `pres.layout = "LAYOUT_16x9"`. NEVER use `"LAYOUT_WIDE"` (wrong dimensions).

## Teacher Notes Rules (Glance Format v12.3)

Teacher notes are a live teleprompter and heads-up display: ~98% of the time the teacher reads them on an iPad mid-lesson, at a glance when confident, read aloud when not. Every teaching slide's notes use the Glance Format: a LIVE ZONE (max 8 logical units, ~120 words, no physical line over ~16 words, blank line between units), a `---` divider on its own line, then a PREP ZONE (max 3 lines). Budgets are RENDERED budgets and `composeGlanceNotes` + `build_and_check.js` Gate 4 enforce them as hard errors. Full spec and worked examples: `IMPORTANT/MEGA_PROMPT.md` sections 45-47.

- Author notes through `composeGlanceNotes({...})`. A beat may be a `string[]`: first line gets the number, continuation lines render indented. One idea per physical line - speech, think time + cue, and `EXPECT:` each on their own line.
- Live zone, fixed order: `ANSWER:` line first whenever the slide asks anything (the most-glanced fact, always in the same place), then 2-5 numbered beats in teaching order, then `TRAP:`, then `STRETCH:` and `HELP:` lines on core I Do / We Do / You Do slides, then `CARE:` for sensitive content only. OMIT the ANSWER line on slides that collect nothing - never boilerplate like `ANSWER: open - not taught`.
- Beats open with CAPS anchors so a glance finds the current moment: `POINT`, `SHOW`, `MODEL`, `DRAW`, `COVER`, `REVEAL`, `TIME`, `COLLECT`, `CIRCULATE`, plus `SAY:` and `ASK:`. The build pipeline bolds recognised anchors in the notes XML automatically.
- A `SAY:` line contains ONLY sayable words - never stage directions, think time, or scan targets fused into the speech. Multi-line ASK beat example: `["ASK: How many equal parts?", "10 sec. Cue: Write it... chin it... show me.", "EXPECT: eight"]`.
- `SCAN` is the decision beat, three short lines, only at genuine decision points: where to look / `80%+ -> [proceed]` / `Less -> [pivot with a different representation], re-ask`. Never one compound sentence.
- `TRAP:` is the error + `Fix:` ending with the student redoing the step, over one or two short lines. `TRAP: counting only shaded parts. Fix: hand on whole strip, count all, student recounts.`
- Reveal slides NEVER copy the base slide's notes. Pass `withReveal(buildFn, revealFn, { revealNotes: composeRevealNotes({ answer, beats, prep }) })` - a short post-reveal script (REVEALED: line, tick-and-fix beat, optional cold-call follow-up). Gate 4 fails the build on consecutive identical notes.
  - **Legacy fallback (do not rely on it):** when `revealNotes` is omitted, `withReveal` DERIVES post-reveal notes from the base slide's notes. It exists only so the 250+ decks written before the rule still build, and it prints an `ADVISORY`. A deck you write from scratch must always pass authored `revealNotes`.
- **Long note lines are wrapped, not rejected.** `composeGlanceNotes` splits any physical line over 16 words into indented continuations and prints an `ADVISORY`. The rendered guarantee is unchanged; only the failure mode is. The 120-word live-zone budget still throws, because that is a content problem - the slide is doing too much - and no rewrapping fixes it. Cut a beat or split the slide.
- **`ADVISORY` lines in build output are work not yet done.** They are non-failing on purpose (legacy decks depend on the fallbacks), but a new lesson should produce **zero**. If you see one, author the thing the pipeline just guessed at.
- Exponents in notes use the ASCII caret (`10^6`), never "10 to the 6" spelled out. Division stays in words (`12 divided by 3`).
- SAY voice is unchanged from what teachers trust: warm natural classroom talk, up to ~20 words per beat (one breath), never clipped fragments (`Watch me`), never presenter copy (`Today we are going to...`). On modelling beats script the think-aloud as connected teacher talk. Action segments are verb-first, <=10 words.
- Student-impact micro rules: every `ASK:` carries think time and ONE all-student response routine (boards, choral, fingers, turn and tell - never volunteer hands); `EXPECT:` is student voice, with `ACCEPT:` where a partial answer counts; explain prompts include a sentence stem (`I know it is ... because ...`); scripted feedback names the strategy, never bare `good job`; `REVEAL` segments state their protection (`REVEAL after boards scanned`).
- Routine tightness (MEGA_PROMPT section 75a): response routines use the school-standard cue scripts, identical in every deck - boards `Write it... Chin it... Show me.`, non-verbal signals cued `voices off`, choral `Everyone, together, on three`. First use in a deck carries the full cue, later beats may shorten to `boards up on cue`, and a broken routine gets a one-line scripted reset before the evidence is read. Hands up is never the sampling method - never write `take some answers` or `ask for volunteers`.
- Think ratio (MEGA_PROMPT section 75): on We Do, CFU and hinge slides, one ASK carries a targeted cold-call follow-up after the all-student response (probe `How do you know?`, bounce `Do you agree? Add one thing`, stretch, clarify), folded into the ASK beat or SCAN proceed clause; brisk routine slides carry none. Each lesson names its 2-3 decision-grade CFU points in the teacher-facing overview; between them pace stays brisk - secure answers get one line of feedback, never a re-teach, and brisk blocks carry a time-budget clause in the prep zone.
- `STRETCH` must deepen or transfer the same concept and be startable without teacher help; `HELP` must change the task form (manipulative, partial model, first step done, frame) and name the gap it targets. "Do more" is not stretch; "do fewer" is not help.
- Prep zone: one purpose/flow line (why the slide exists, assumption flags, internal SC focus) plus the single tag `[Stage | VTLM element | SC | HITS n]`, an optional `SOURCES:` line whenever a slide uses an external image or externally sourced claim, and an optional one-line `WHY:` misconception background. Nothing mid-lesson-critical goes below the divider.
- Formatting: plain text, ASCII-safe (straight quotes, `->`, `>=`, `x`, `^`; no smart quotes, em dashes, unicode bullets or arrows). No markdown. No bullet formatting - the typed numbers, CAPS anchors and blank lines between units carry the structure. No blank lines inside a unit.
- Non-teaching slides (title, credits, pure dividers) get one plain line of notes, no zones.
- If a slide uses `liSlide()`, the Learning Intention must be a single plain sentence and the Success Criteria must be exactly 3 simple `I can...` bullets. The first success criterion must be ultra-achievable for almost every student.
- Do not add a `PACING OVERVIEW` block. If timing genuinely matters, fold one short clause into the prep-zone purpose line.

## Cognitive Load Defaults

- `Lean` means fewer, better-taught moves, not less learning. Apply this across all sessions, not just literacy.
- Protect the high-yield parts of instruction: clear modelling, repeated practice, retrieval, CFU, guided practice, and independent application.
- If a lesson feels overcrowded, cut low-yield extras first: duplicate explanations, oversized vocab banks, unnecessary reveal pairs, long note essays, decorative transitions, and multiple competing objectives.
- Default future generations to `mixed readiness`, not assumed mastery. Avoid student-facing or `SAY:` phrasing such as `you already know`, `students know the routine`, `not new to you`, `we've done this`, or `by Week X students know` unless the user explicitly asked for a revision/review lesson.
- Beginner-safe prior-knowledge language is allowed: `Some of you may remember...`, `If this feels new, that's okay`, `We'll build this together`.
- Less on the slide does not mean less teaching. It means the teacher voice and the practice sequence carry the load instead of cluttered slide text.

## Lean Literacy Defaults

- Default future generations to `mixed readiness`, not assumed mastery. Avoid student-facing or `SAY:` phrasing such as `you already know`, `students know the routine`, `not new to you`, `we've done this`, or `by Week X students know` unless the user explicitly asked for a revision/review lesson.
- Beginner-safe prior-knowledge language is allowed: `Some of you may remember...`, `If this feels new, that's okay`, `We'll build this together`.
- Default a 60-minute literacy lesson to one reading/comprehension or craft focus plus one writing/language focus only.
- Default literacy lesson shape: title, Teacher Resources, hook or text launch, LI/SC, 0-2 explicit vocab slides if needed, up to 2 pause points, 1 craft/analysis slide, 1 CFU, 1 I Do, 1 We Do, 1 You Do, closing.
- Default budget for a 60-minute literacy deck is 10-14 unique slides. Above 14 means the lesson is probably too crowded. Above 16 requires an explicit reason from the user.
- Default reveal budget is 0-2 reveal pairs. Use reveals only when hiding the answer materially improves thinking. Do not use reveal pairs by default for every vocabulary, CFU, or We Do slide.
- Incidental vocabulary list slides are off by default. Only include them when the source text genuinely demands them or the user explicitly asks for them.
- Slide-face text should stay lean. Do not preload large definition banks, long explanation blocks, or multiple abstract objectives onto one lesson by default.

For full PptxGenJS API reference: read `docs/pptxgenjs-reference.md`.

## Layout Safety (10" x 5.625")

**Slide layout: Always use `pres.layout = "LAYOUT_16x9"` (10" x 5.625").** NEVER use `"LAYOUT_WIDE"` - it creates a 13.33" x 7.5" canvas but all theme builders and positioning constants assume 10" x 5.625", causing content to appear cropped/small in the top-left portion of the slide.

Content area: y 1.3"-5.1". Footer: y 5.3". NEVER place content below y 5.1".
For dynamic content, calculate total height and clamp to stay within the safe zone.
Console warnings during build = layout bugs. Fix before shipping.

### Two-Column Layout Rules (contentSlide / workedExSlide with drawRight)

- `contentSlide` left card: 4.5" wide (x 0.5-5.0). Right column: x 5.2, w 4.3.
- `workedExSlide` left card: 4.5" wide (x 0.5-5.0). Right column: x 5.3, w 4.2.
- **NEVER place right-column elements at x < 5.2.** The left card ends at x 5.0; anything placed before x 5.2 will overlap left-column text.
- When using `layoutGuide` from the `drawRight` callback, always start right-column content at `layoutGuide.rightX` or later.
- Both builders now auto-run `runSlideDiagnostics` when a `drawRight` callback is provided. Any overlap ERROR in the build output means content is visually hidden — treat it as a blocker.

### Text Box Sizing

- PptxGenJS text boxes do NOT clip overflow — text that exceeds the box height renders BELOW the box boundary and overlaps whatever is underneath.
- **Always size text boxes to fit their content.** Count lines, estimate height (fontSize × 0.022" per point × lines + padding), and verify the text box is large enough.
- Rule of thumb for body text height: at fontSize 12, each line needs ~0.22"; at fontSize 14, ~0.26"; at fontSize 16, ~0.30". Add ~0.1" padding.
- When placing a summary card with text inside and a separate element below it (e.g. a verdict banner), ensure `textY + textH` does not exceed `cardY + cardH`, and the element below has at least 0.15" clearance from the card's bottom edge.

### Reveal Bar Clearance

- `withReveal` now automatically checks the reveal layer against the base slide's RENDERED text and emits a gate-failing WARN if a reveal element covers it. Fix by shortening the question, raising its box, or moving the bar — never by deleting the check.
- When using `withReveal` and adding a reveal element (e.g. an answer bar) in the `revealFn`, ensure all content on the slide stops at least 0.15" ABOVE the reveal element's top edge.
- If the reveal bar is at y 4.25, the tallest content text box must end by y 4.1 at most.
- For factor-pair lists, prompts, or other variable-length content above a reveal bar, reduce the text box `h` to enforce this ceiling rather than letting it extend to `SAFE_BOTTOM`.

### Title Sizing

- Long titles that wrap to 2+ lines push content down. When a title exceeds ~45 characters, verify that the content below still fits without overlapping the footer zone.
- For custom slides, use `layoutGuide.panelTopPadded` (available from `contentSlide` and `workedExSlide` callbacks) as the starting y for right-column content when the title is long.
- Prefer concise titles (under 40 chars) for slides with dense two-column layouts.

## Shared Visual Helpers (Mandatory)

Every theme object (all subjects) carries grade-band-aware visual anchor helpers. **NEVER hand-draw a representation this table covers with raw addShape/addText** — hand-rolled versions are where fused grids, broken arrowheads and spacing hacks come from. If a lesson needs a variant a helper cannot draw, extend the helper in `themes/core/manipulatives.js` first, then use it.

| Need | Use |
|---|---|
| Tens frame / five frame | `addTensFrame(slide, x, y, w, filled)` / `addFiveFrame(...)` |
| Dot card (subitising) | `addDotCard(slide, x, y, size, count)` |
| Number track | `addNumberTrack(slide, x, y, w, start, end, [highlights])` |
| Number line | `addNumberLine(slide, x, y, w, labels[], marked[])` (`""` = unlabelled tick) |
| Fraction strips (separate wholes) | `addFractionStripSet(slide, x, y, w, h, [{denom, shaded, label, color}])` |
| Array of dots | `addArray(slide, x, y, rows, cols)` |
| MAB / base-10 | `addBaseTenBlocks(slide, x, y, hundreds, tens, ones)` |
| Row of choice chips | `addChipRow(slide, x, y, w, ["1/5", "3/4", ...])` — never space-separated inline text |
| "Groups of" counters | `addGroupedCounters(slide, x, y, groups, per)` |
| Part-part-whole mat | `addPartPartWholeMat(slide, x, y, w, h, {whole, partA, partB})` (`null` = blank box) |
| Answer reveal | `addRevealAnswerBar(slide, [answers], {y, h, fontSize})` inside `withReveal` revealFn — never a hand-placed success bar |
| Vocabulary word card | `keyWordSlide(pres, { word, meaning, example, pictogram }, notes, footer)` — ONE word per slide WITH its picture (`pictogram: "<name>"` or `image: path`). NEVER render vocabulary as a definition bullet list; call once per word (F-2: 1-3 words, Y3-4: 2-4, Y5-6: 2-5). A word card with no graphic prints an ADVISORY and is not finished |
| Any representation, sized to fill a frame | `drawVisual(slide, { type, ...values }, frame)` — declarative: `{ type: "tensFrame", filled: 7 }`, `{ type: "numberLine", start: 0, end: 2, step: 1/3, marked: [3] }`, `{ type: "fractionStrips", strips: [...] }`, `pictograms`, `table`, `text`, `image`, `custom`. Every `drawRight` slot (`contentSlide`, `workedExSlide`, `dailyReviewSlide`) accepts a spec instead of a callback |
| Visual-only teaching slide | `heroVisualSlide(pres, badge, title, visualSpec, notes, footer, { label, prompt })` — the representation fills a soft panel; use for F-2 concept slides and any slide whose purpose is the model |
| Which one? / A-B-C / example and non-example | `choiceSlide(pres, badge, title, prompt, [{ visual, text, caption }...], notes, footer)` then `clickBuild(s, [() => markChoice(s, correctIndex)])` — never hand-place option cards |
| You Do task | `youDoSlide(pres, title, task, ["First...", "Next...", "Then..."], notes, footer, { where, visual, frame })` — task is the hero, steps are chips |
| Text extract / read-aloud | `textExtractSlide(pres, badge, title, extract, notes, footer, { highlights, source, prompt })` — exact quoted text, marker-highlighted phrases |
| Picture for a word, stage, feeling or hook | `addPictogram(slide, name, x, y, size, { style, color, label })` / `addPictogramRow(slide, x, y, w, ["happy", "sad"])` — 200+ built-in names (`listPictograms()`, sheet in the Visual Catalogue). Science `cycleDiagramSlide` / `processFlowSlide` steps take `icon: "<name>"`. Unknown names fail the build |
| Data table | `addDataTable(slide, x, y, w, rows[][])` or `{ type: "table", rows }` — themed header, zebra rows, band-sized type |

PDF twins for worksheets/scaffolds (in `themes/pdf_helpers.js`): `addTenFramePdf`, `addFractionStripsPdf`, `addNumberLinePdf`, `addPpwMatPdf`, `addHundredGridPdf` (the paper twin of `addAreaModel` — 10x10 grid, cells fill column by column), plus `addCycleDiagramPdf` (cycles/loops — never hand-draw cycle arrows with doc.moveTo, they come out tangled) and `addPosterMockupPdf`/`addPosterPairPdf` (designed visuals on paper). Same rule: never hand-draw these with raw pdfkit primitives.

Visual reference deck: `node scripts/build_and_check.js builds/build_visual_catalogue.js` renders every helper per grade band, every pattern builder and the full pictogram sheet to `output/Visual_Catalogue/`. Rebuild and re-inspect it after ANY change to the theme's visual helpers, pictograms or builders, then run `npm run qa:theme` (includes `tests/test_visual_builders.js`).

`addRoutineBadge` and `addPictogram` are synchronous (resvg-js). Never `await` a builder to make an icon appear; if an icon is missing, the build printed a `WARN`.

## Key Conventions

- Palette uses semantic keys: PRIMARY, SECONDARY, ACCENT, ALERT, SUCCESS, BG_DARK, BG_LIGHT, BG_CARD, CHARCOAL, WHITE, MUTED. Backward-compatible aliases exist (C.NAVY, C.CREAM, C.TEAL).
- White text on coloured fills. Dark text on light fills. NEVER same colour for text and its background.
- White icons need a coloured circle background on light surfaces (that is what `addPictogram` style `circle` and `addRoutineBadge` do).
- Title slides take `opts.visual` (a visual spec) to show the lesson's own anchor instead of the subject glyph: `titleSlide(pres, title, subtitle, meta, notes, { visual: { type: "tensFrame", filled: 10 } })`. Never draw decorative shapes on title or closing slides.
- Images are opt-in instructional tools, not decoration. Use local lesson-cached or unit-cached assets only, and only when they directly support understanding.
- Distinguish `visual anchor` from `actual image`. A diagram, labelled mockup, source layout, or builder like `annotatedModelSlide(...)` counts as a visual anchor. A real local image is required when students are meant to interpret authentic visual evidence such as a photograph, map, artefact, poster, illustration, or source document itself.
- For literacy topics involving source analysis, text features, advertisements, posters, article layout, maps, artefacts, or compare-text-and-visual evidence, at least one core teaching slide should use a visual anchor such as `annotatedModelSlide(...)`, `addInstructionalImageCard(...)`, or another explicit source/feature layout. Text-only bullets are not sufficient by default for these cases.
- If the lesson is about structure or feature-spotting, a built visual mockup is usually sufficient. If the lesson is about inferring from or analysing a real source image, map, poster, artefact, or illustration, use an actual local instructional image rather than replacing it with a generic mockup.
- For structure/layout lessons, prefer clean wireframe-style mockups over pseudo-real scenic art. The mockup should clarify hierarchy, navigation, and information placement rather than trying to imitate illustration or photography unless the image itself is the instructional object.
- For visual-analysis lessons, preserve the visual object through the GRR where it remains the thing students are analysing. In We Do, fade labels or prompts first, not the visual itself. Do not replace a poster/map/source/diagram analysis task with a prose description if students still need to reason about visual features.
- When the We Do requires side-by-side comparison of two designed visuals, prefer `compareVisualSlide(...)` or another explicit dual-visual layout instead of two text-description cards.
- For poster, advertisement, article-layout, or similar designed-visual lessons, the mockup itself must look like the thing being analysed. Do not feed `annotatedModelSlide(...)` or `compareVisualSlide(...)` descriptive placeholder strings such as `Image: ...` or `Colour scheme: ...` inside the preview. Use structured mockups or real local images so students can infer from layout, emphasis, and visual hierarchy by looking.
- For newspaper front page, article layout, poster, infographic, and similar designed-visual I Do slides, prefer a structured `previewSpec` (poster spec with `components` array) over flat `previewBlocks` text when the builder supports it. Both `annotatedModelSlide(...)` and `compareVisualSlide(...)` support `previewSpec` via the shared `drawMockupPreview` path. Do not downgrade content to flat text to work around a stale builder override; fix the shared builder layer instead.
- **previewSpec consistency rule:** If a build script defines a structured mockup spec object (an object with a `components` array) for a designed visual, every builder call in the same lesson that renders that visual MUST use `previewSpec`, not `previewBlocks`. Do not define a spec and then pass `previewBlocks` to `annotatedModelSlide(...)` while passing `previewSpec` to `compareVisualSlide(...)` for the same visual — this produces an inconsistent visual fidelity between I Do and We Do. If a builder cannot render the spec, fix the shared layer rather than downgrading the lesson content.
- For science topics involving systems, cycles, sequences, life stages, or body processes, at least one core teaching slide MUST use a dedicated visual anchor such as `cycleDiagramSlide(...)`, `processFlowSlide(...)`, a labelled diagram, or a clearly instructional local image. Text-only cards are not sufficient by default for these topics.
- Water cycle, life cycle, digestive system, food chains, circuits, Earth-sun-moon systems, and similar content should be treated as mandatory visual cases unless there is a concrete reason not to.
- Persuasive posters, newspaper/article features, source-photo inference, historical nonfiction with maps/artefacts, and similar literacy lessons should also be treated as visual-anchor cases unless there is a concrete reason not to.
- Do not let sparse content float inside oversized full-height cards. If a slide only has a few short bullets or prompts, use a compact card or a two-column visual layout so the slide feels intentionally designed rather than underfilled.
- Student-facing instruction cards and prompt panels must start large enough for classroom viewing. For sparse prompts, target roughly 16-17 for the header and 14-15.5 for body lines, then shrink only if needed. Do not default to 12pt body text in roomy dialogue/instruction boxes.
- Prefer the shared `addInstructionCard` theme helper for left-hand "On your whiteboards" / "With your partner" cards and any similar sparse student-instruction panel so sizing is density-aware by default.
- When using `contentSlide(..., drawRight)` or numeracy `workedExSlide(..., drawRight)`, use the callback's second `layoutGuide` argument for custom right-column positions. Do not hardcode custom panels flush to `CONTENT_TOP` when the slide also has a long title; start from `layoutGuide.panelTopPadded` unless you have visually verified a tighter layout.
- Theme diagnostics are available for manual/custom slides: `runSlideDiagnostics(slide, pres)` plus the narrower `warnIfSlideHasOverlaps(...)` and `warnIfSlideElementsOutOfBounds(...)`. Use them before shipping any custom layout.
- `contentSlide` and `workedExSlide` now auto-run diagnostics when a `drawRight` callback is provided. Any ERROR or WARN in build output is a layout bug — fix it before shipping.
- Diagnostics also flag UNDERFILLED slides: if content stops in the top half of the content area, you get a WARN telling you to enlarge the hero task/visual or centre the layout. Do not shrink the check away — make the hero bigger (that is the fix the mega-prompt wants). `{ ignoreUnderfill: true }` is allowed only for deliberate visual-only white space (e.g. a lone Foundation ten frame).
- `cfuSlide`, `exitTicketSlide` and `contentSlide` are density-aware: short questions and one to three short lines render hero-sized on a soft tint panel, vertically centred, automatically. Prefer them over hand-built question slides.
- Standard footer text is ignored by diagnostics. If a custom footer-like element is falsely flagged, keep diagnostics enabled and use a narrow `ignoreIndices` override rather than disabling safe-bottom checks for the whole slide.
- Theme image helpers are available for local assets: `addImageWithCaption(...)` and `addInstructionalImageCard(...)`.
- `annotatedModelSlide(...)` is available on every theme object for labelled source features, poster/article structure, and "notice this part" teaching. Do not swap subjects just to reach it.
- `compareVisualSlide(...)` is available on every theme object for We Do comparison of two posters, layouts, advertisements, or similar designed visuals.
- Science process/system topics can also use the dedicated `processFlowSlide(...)` builder for ordered journeys, cycles, and body systems.
- Science cycle topics should prefer the dedicated `cycleDiagramSlide(...)` builder over manual text-plus-arrow layouts.
- **`clickBuild(slide, [step, step, ...])` is the preferred reveal mechanism** (MEGA_PROMPT §20b). Each step is a function adding the elements that appear on that click; anything added outside a step is visible from the start. It writes real PowerPoint entrance animations into the finished file, so one slide carries the whole build. Use it for I Do models built step by step, We Do answers, CFU reveals and Daily Review answers.
  ```js
  const s = contentSlide(pres, "12 divided by 3", [...], notes, footer);
  clickBuild(s, [
    () => { s.addText("3 groups", {...}); },
    () => { addRevealAnswerBar(s, ["4"], { y: 4.25 }); },
  ]);
  ```
  Shape ids are derived from `_slideObjects` position, so call `clickBuild` AFTER the base slide is built and do not remove elements afterwards — the build fails loudly if a step targets a missing element.
- `withReveal(buildFn, revealFn, { revealNotes })` duplicates the whole slide and is now the FALLBACK, for when the answer slide needs a genuinely different layout rather than extra elements on top. It doubles the slide count and forces separate reveal notes. Do NOT use for I Do (it breaks the model in half), exit tickets, or titles.
- Every lesson with companion PDFs gets a resource slide via `addResourceSlide()` from `pdf_helpers.js`.
- Per-lesson build output goes to `output/<LessonFolder>/` - PPTX at the root, companion PDFs in a `resources-session{N}/` subfolder where `N` is the session number within that week's sequence. This is the build step, not the delivery step.
- PptxGenJS hyperlinks use relative paths - include the subfolder prefix (e.g., `resources-session3/Session 3 Worksheet.pdf`).
- Resource names must be teacher-friendly and session-first: `Session 1 Worksheet`, `Session 1 Answer Key`, `Session 2 Enabling Scaffold`.
- Use the same human-readable name on the resource slide and in the PDF filename stem. Avoid codes like `WH4_L16`, `SR1`, `GO1`, `ET_Lesson5`, or similar.
- Do not use day names in resource filenames. Teachers run sessions on different days.
- Do not use underscores in teacher-facing PDF filenames. Use spaces.

## Orton-Gillingham (OG) Decks - Separate Pipeline

OG morphology session decks do NOT use the theme system, PptxGenJS, or `builds/`.
They are generated by cloning a locked master template (fonts, positions, icons and
click animations preserved exactly; type colours are normalised by the builder):
`python og_planner/build_og_week.py og_planner/weeks/<spec>.json`
(miniconda `python` - needs lxml/python-pptx). Full rules, content recipes, week-spec
schema and the user request format: `IMPORTANT/OG_MEGA_PROMPT.md`. The authoritative
photo transcriptions for the captured set are `og_planner/yoshimoto_cards_suffixes.json`,
`yoshimoto_cards_prefixes.json`, and `yoshimoto_cards_latin_roots.json`; they include
printed meanings, printed parts of speech where present, selected school-friendly
keywords, Australian-English associated words, and explicit exclusions. These fields
are locked for captured cards: do not guess or override them, do not automatically use
excluded words, and log any verified outside derivative using the exception schema in
`IMPORTANT/OG_MEGA_PROMPT.md` section 4. These catalogues win over the legacy
`morpheme_bank.json` and the unconfirmed `morpheme_meanings.json` reference catalogue.
The universal type-colour key overrides any older master colour: GREEN = root,
YELLOW = prefix, RED = suffix. The builder enforces it on morphology review-card
backgrounds, new/review morpheme-card backgrounds and Sound Bank boxes, and its
finished-PPTX gate rejects mismatches.
OG teacher notes follow the shared Glance Format: plain source text, one thought per
line, left-aligned paragraphs, eight live-zone lines maximum, and no markdown. Words
to Read Review is the exception: use separate cues for the whole-class read, exhaustive
playful groups for rows 1-4, and everyone reading row 5 so every student reads at least
two rows, followed by one retrieval question per line. The OG builder applies real bold
to recognised labels and underlines the spelling target where it appears inside every
`Sentence:` note line. Fixed-
answer new morphology activities require an immediate green answer slide titled `Tick
it or fix it - ...`; never use `Check and fix`. New-word grids must remain unbroken at
27 pt or larger. The first dictation uses the green meter and the second/trickier
dictation uses yellow; capitals render green, punctuation red, and spelling targets
bold/underlined. Raw Yoshimoto
reference library: `OG/`, navigated via `og_planner/OG_LIBRARY_INDEX.md` (scanned PDFs - render pages to images to read them).
Deliverable is one PPTX per session in `output/<week folder>/` - never merge OG decks.
The builder reopens each finished PPTX and hard-fails on dense or malformed notes,
missing real bold note labels, legacy `Derivative ask` output, undersized new-word
grids, missing/non-green answer checks, undersized grammar examples, or incorrect
dictation cue formatting. It also hard-fails when the Sound Bank or review cards
contain the day's focus morpheme or any morpheme this week teaches on a later day (the
bank is copied into books before that morpheme is taught), when a dictation target is
one of this week's new grid words / new learned words or is built on this week's focus
morphemes (dictation is revision from 2-3 weeks ago - students copy this week's words
from their books, school feedback Aug 2026), or when a You
Do morpheme sum uses an affix outside the taught set; the rules and the escape hatch
(`taught_morphemes`) are in OG_MEGA_PROMPT sections 2a/2c, 6 and 2f. Builder output uses two
levels: every `WARN` fails the build, a `NOTE` is advisory but still needs answering.
Run `python tests/test_og_builder_regressions.py` before
changing the OG builder, its sample specification, or its note rules.

## Auslan Decks - Two-Step Pipeline

Auslan units use two prompts in `IMPORTANT/`, numbered by step:
`AUSLAN_1_UNIT_PROMPT.md` (run in a chat, produces the unit planning document,
which always delivers as a Word file - author the Markdown, run its output
hygiene checks, then convert with `python scripts/md_to_docx.py in.md out.docx
"Title" "Subtitle" "Meta"`; there is no pandoc here, and QA the result via
LibreOffice + PyMuPDF because pdftoppm is not installed)
then `AUSLAN_2_SLIDES_PROMPT.md` (run in this repo with that document pasted,
builds the session decks + PDFs through the normal theme pipeline, literacy
theme). `AUSLAN_GAME_BANK.md` is an optional companion pasted alongside step 1:
a growing repository of voice-off games, teacher-supplied ones reproduced
faithfully and published mechanics rebuilt for a signing room. Add to it when a
game has been run and worked; do not reconstruct its entries from memory when it
is not pasted. The bank itself is gitignored, as is its published source in the
`reference/` folder (purchased copy, schools statutory educational licence,
internal use only, never committed) - the bank carries page-level citations to
that source, so pushing it to this public remote would be redistribution. It
lives on disk only, so on a fresh clone it will be absent and must be pasted
from the teacher's own copy. Hard rules that override everything else for Auslan builds: never draw,
generate, or mirror a sign image - images come only from the shared bank at
`assets/auslan_signs/`, populated by `python scripts/fetch_auslan_signs.py
--glosses ...` (Signbank frame-sequence strips, `<GLOSS>.jpg` plus `_2`/`_3`
variants; see the bank README for licensing), with a lookup-card fallback
linking `https://auslan.org.au/dictionary/search/?query=<word>` (never a guessed
Signbank entry URL); never compose multi-sign gloss strings on any slide or PDF
- students read plain English, the teacher models the Auslan. Every Auslan build
script prints a SIGN ASSETS report (found/missing glosses) which goes in the
final summary. `assets/auslan_signs/manifest.json` records the Signbank entry
and dictionary definition behind every image - check a sign's recorded sense
matches the lesson before shipping, because a plausible image of the wrong sense
is invisible in a rendered slide.
**The sign images, and the manifest, are gitignored on purpose**: they are
licensed for internal school teaching only and this repo has a public remote.
Only the recipe is tracked (README, `core_glosses.txt`, the fetcher). If the
bank is empty on a fresh clone, rebuild it with
`python scripts/fetch_auslan_signs.py --from-file assets/auslan_signs/core_glosses.txt`
(about 2 minutes). Never commit sign images and never publish a built Auslan deck.

## Multi-Session Unit Delivery (Required)

When the user requests more than one session in a single ask (a unit, a week, a multi-day sequence, "lessons 1 to N"), the delivered output MUST be one combined PowerPoint and one flat `Resources/` subfolder. Per-lesson folders are a build-step intermediate, not the deliverable.

**Workflow (do not skip the merge):**

1. Write one per-lesson build script per session in `builds/` as usual.
2. Write a manifest at `builds/manifests/<unit>.json` listing each lesson's `build_script`, `folder`, and `session` in teaching order, plus `unit_folder` and `unit_pptx_name`. Manifest format is documented in `scripts/merge_unit.py` and `docs/resource-system.md`.
3. Run `python scripts/build_unit.py builds/manifests/<unit>.json`. This builds every lesson through `build_and_check.js` (aborting on any gate failure), merges the decks and resources into `output/<unit_folder>/<unit_pptx_name>` + `output/<unit_folder>/Resources/<flat PDFs>`, then runs merged unit QA via `qa_unit.js --skip-build --skip-merge`.
4. The task is not "done" for a multi-session request until the combined unit folder exists. Do not claim completion after building per-lesson folders only.
5. For a single-session request, the per-lesson folder IS the deliverable — no merge needed.

If you fix one lesson later, rebuild just that lesson with `build_and_check.js`, then re-run `build_unit.py ... --skip-build` to re-merge and re-run merged unit QA without rebuilding the rest.

Resource filenames must be unique across the unit (the merge flattens all PDFs into one folder). The `Session N` prefix that `pdf_helpers.js` enforces handles this automatically — do not strip it.

**No Teacher Week Brief.** Do NOT generate a Teacher Week Brief, weekly summary PDF, or any standalone teacher-preparation document for a unit — school leadership has ruled these out (they encourage over-reliance instead of teachers reading the deck and notes). Never add a `teacher_brief` object to a manifest. Teacher preparation lives inside the deck: prep-zone note lines, the teacher-facing overview slide, and the Teacher Resources slide.

For resource generation details and PDF helper API: read `docs/resource-system.md`.
For ad-hoc (non-themed) presentation design guidance: read `docs/design-guide.md`.

**Scaffold quality:** An enabling scaffold must change the FORM of the task, not just the wording. It must draw a visual model, pre-fill intermediate steps, or provide a structural framework. If you claim "the model is drawn for you," draw the model with PDFKit primitives. Text that describes a visual is not a visual. Read `docs/resource-system.md` section "Scaffold Quality Rules" before writing an enabling scaffold PDF.
- For visual-analysis scaffold PDFs, include the visual object on paper as well. If students are comparing posters, advertisements, maps, or layouts, the PDF must show schematic or real versions of those visuals; prose descriptions are not an acceptable substitute.

## Build Script Authoring (Critical)

**NEVER delegate build script writing to agents/subagents.** Always write build scripts directly in the main conversation context. This is a hard rule learned from experience:

- Agents lack the accumulated context of PptxGenJS rendering quirks, builder signatures, layout constants, and the iterative build-inspect-fix discipline that produces correct output.
- Agents invent custom drawing helpers (flowcharts, Venn diagrams, tables) with hardcoded coordinates that haven't been tested. These consistently produce overlaps and misalignment.
- Agents favour manual `addShape`/`addText` with raw x/y/w/h values instead of using the tested theme builders (`contentSlide`, `workedExSlide`, etc.), which is fragile.
- The QA pipeline (markitdown, Google Slides review, and optional local preview images) catches content errors but cannot reliably catch the subtle layout regressions that agents introduce at scale across 1500+ line scripts.
- Previous builds that passed QA and rendered correctly in Google Slides were ALL written directly, never by agents.

**Use the tested theme builders** (`titleSlide`, `liSlide`, `contentSlide`, `cfuSlide`, `workedExSlide`, `exitTicketSlide`, `closingSlide`) for every slide that fits their signature. Only go manual for truly novel layouts, and test those individually.

**Archived scripts are not active exemplars.** Do not scan `_archive/` for nearby scripts to update or imitate by default. Treat that folder as historical reference only. If `builds/` is empty, build from the shared theme system, current docs, and the user brief rather than reviving archived lesson files.

Agents ARE useful for: research, reading reference files, visual QA inspection of rendered slide images, and content review. Just not for writing the build scripts themselves.

## QA (Required)

First render is almost never correct. After every build:
0. **Use `node scripts/build_and_check.js builds/build_<unit>_lesson<n>.js` as the default build command.** It runs seven gates: build, diagnostics, markitdown + forbidden markers, slide text hygiene, teacher notes format, hyperlink integrity, and lesson structure (resources placement, opening order, We Do vs You Do). If it exits non-zero, the build has failed — fix the issue before proceeding. Do NOT skip this step or ignore its output. **The gate script is the minimum automated bar, not a substitute for visual inspection.** Passing it means the build is structurally sound — it does NOT mean the slides look correct.
1. **Smoke build early.** If the script contains any manual/custom slide work, new helper usage, or new resource generation, run `build_and_check.js` after writing the PPTX-generating code but BEFORE writing companion PDFs. Do not write the entire script (slides + PDFs) in one pass and only build at the end. Catch API/signature errors while the change set is small and the fix is obvious.
2. The gate script covers markitdown automatically. If it reports FAIL on the markitdown gate, that is a blocker — do not dismiss it as "intermittent" or "environmental" without concrete evidence (e.g. markitdown works on other PPTX files in the same session).
3. **Visual QA is required after the gate passes.** Run `pptx_to_images.py` to generate slide previews, then inspect them directly. Look for: overlaps, text overflow, low contrast, uneven spacing, missing elements, text cut off, reveal mistakes, broken links, or elements below 5.1". The gate script cannot catch single-text-box overflow, reveal bar overlap, or visual imbalance — only eyes can.
4. Optional contact sheet (`slide_montage.py`) - generate a quick montage from `slidetemp/` when scanning many slides or sessions.
5. Final visual and compatibility QA in Google Slides - import the `.pptx` and inspect title, content, reveal, subject-specific, closing, and resource slides.
6. Fix issues, re-verify affected slides. One fix often creates another problem.
7. Repeat until a full pass reveals no new issues.
8. Clean up optional preview images: `python scripts/pptx_to_images.py --clean`

Do not say "QA passed" unless the Google Slides compatibility pass in step 5 is complete. Do not treat a passing gate script or local visual inspection as delivery-ready. If only the gate script ran, state that automated gates passed and visual review is still pending. If local visual QA ran but not Google Slides, state that local QA passed and Google Slides review is still pending.

## Dependencies

```bash
pip install "markitdown[pptx]" Pillow pymupdf   # Python: content + visual QA
npm install                                     # Node: pptxgenjs, pdfkit, react-icons, sharp, @resvg/resvg-js (sync icon rendering)
# LibreOffice (soffice) only needed for optional local image preview
```
