# Theme System Reference

> Lessons are authored as specs (`builds/<name>.json`, see `docs/lesson-spec.md`); the pipeline in `themes/lesson/` calls the builders below. This document is the reference for the builders themselves and for the rare custom slide.

The theme system uses a **factory pattern** with 150 pre-built colour palettes across 5 subjects, 5 year levels, and 6 weekly variants. Build scripts import a single factory function — all builders, element helpers, and palette colours are returned as a bound theme object.

## Factory API

```javascript
const { createTheme, weekToVariant } = require("../themes/factory");

// Create a theme for a specific subject + year level + week
const T = createTheme("literacy", "grade56", weekToVariant(3));

// Destructure — everything you need is on one object
const { C, FONT_H, FONT_B, titleSlide, liSlide, contentSlide,
        cfuSlide, closingSlide, annotatedModelSlide, compareVisualSlide, withReveal, addCard, addFooter,
        addTextOnShape, addImageWithCaption, runSlideDiagnostics,
        iconToBase64Png, getContrastColor } = T;

// Subject-specific builders are also on the theme object
const { vocabSlide, quoteSlide, modellingSlide } = T;  // literacy
const { workedExSlide, exitTicketSlide } = T;           // numeracy
const { experimentSlide, observationSlide } = T;        // science
const { scenarioSlide, reflectionSlide } = T;           // wellbeing
const { investigationSlide, findingsSlide } = T;        // inquiry
```

**Parameters:**
- `subject` — `"literacy"` | `"numeracy"` | `"inquiry"` | `"wellbeing"` | `"science"`
- `yearLevel` — `"foundation"` | `"grade1"` | `"grade2"` | `"grade34"` | `"grade56"`
- `variant` — `0`–`5` (weekly rotation index)

**`weekToVariant(weekNumber)`** converts a 1-based week number to a 0-based variant index (cycles 0–5).

Theme selection should follow the explicit `Subject:` field from the lesson prompt. Do not silently change themes because the content feels cross-curricular.

## Subjects and Their Builders

All subjects get the **base builders**: `titleSlide`, `liSlide`, `contentSlide`, `cfuSlide`, `closingSlide`, `keyWordSlide`, `exitTicketSlide`, `boardBuildSlide`, `annotatedModelSlide`, `compareVisualSlide`, plus the **pattern builders** `heroVisualSlide`, `choiceSlide` (+ `markChoice`), `youDoSlide`, `textExtractSlide`, and the helpers `addRevealAnswerBar`, `addDataTable`, `addPictogram`, `addPictogramRow`, `drawVisual`. `keyWordSlide(pres, { word, meaning, example, pictogram }, notes, footer)` is the vocabulary word card - one word per slide with its picture, never definition bullet lists. `cfuSlide`, `exitTicketSlide` and `contentSlide` are density-aware (one to three short lines render hero-sized on a soft tint panel, vertically centred).

| Subject | Extra Builders | Purpose |
|---------|---------------|---------|
| **Literacy** | `vocabSlide`, `quoteSlide`, `modellingSlide`, `pairShareSlide` | Text study, vocabulary, modelled writing |
| **Numeracy** | `workedExSlide`, `dailyReviewSlide`, `fluencySlide`, `addPlaceValueChart`, `addNumberLine`, `addAreaModel`, `addTenthsStrip`, `addDecimalDot`, `addStageBadge` | Worked examples, maths visuals |
| **Inquiry** | `investigationSlide`, `findingsSlide`, `pairShareSlide` | Question-driven, evidence gathering |
| **Wellbeing** | `scenarioSlide`, `reflectionSlide`, `pairShareSlide` | Social scenarios, discussion, reflection |
| **Science** | `experimentSlide`, `observationSlide`, `conclusionSlide`, `processFlowSlide`, `cycleDiagramSlide` | Scientific method structure, ordered processes, systems, cycles |

## Design Language (what every deck inherits)

The theme layer, not the build script, carries the look. A build script that uses the builders gets all of this for free; a build script that hand-draws with `addShape`/`addText` has to reproduce it, which is why hand-drawing is discouraged.

- **Colour does one job each.** Strong role colours (`PRIMARY`, `SECONDARY`, `ACCENT`, `ALERT`, `SUCCESS`, `ASSESS`) appear only on small, meaningful surfaces: the stage badge, the top bar, technique pills, answer bars, chips, ticks. Large surfaces (hero panels, question cards, option cards) use the derived **soft tint** `C.<ROLE>_SOFT` with a hairline `C.<ROLE>_LINE` border. `T.softOf(hex)` and `T.lineOf(hex)` derive the same tints for any colour a script chooses.
- **Palettes are bright enough to look like their year level.** Every role colour is the brightest shade of its hue that still clears a per-band contrast target against white (Foundation about 4.9:1 for `PRIMARY`, rising to about 6.8:1 by Year 5/6). `scripts/retune_palettes.js` re-establishes those floors after any hue edit, so a Foundation deck reads as royal blue and grass green rather than navy and bottle green, and white text on any role colour is always AA.
- **One motif, repeated.** The title slide carries the subject glyph in a soft circle (or the lesson's own visual anchor via `titleSlide(..., { visual })`); the closing slide repeats the glyph small. There are no decorative blobs, gradients or accent lines under titles.
- **Three card surfaces.** `addCard(..., { variant: "white" | "tint" | "outline", tone })`. White cards (soft shadow, optional left strip) hold supporting content; tint cards hold the hero task or question; outline cards hold option cards and reading panels.
- **Hero sizing is automatic.** Short content is set large and centred (`cfuSlide`, `exitTicketSlide`, `contentSlide`), representations are fitted to fill their frame (`drawVisual`), and pill badges shrink-fit long labels.
- **Pictures are built in.** 200+ pictograms (`addPictogram`) give vocabulary cards, science stages, feelings and launch hooks a picture without any hunt for images.

## Visual Specs (`drawVisual`)

A visual spec names a representation and its values; the theme sizes and centres it. Every builder that takes a right-column callback (`contentSlide`, `workedExSlide`, `dailyReviewSlide`) also accepts a spec in that slot, and `heroVisualSlide`, `choiceSlide`, `youDoSlide` and `titleSlide` take specs directly.

```javascript
drawVisual(slide, { type: "tensFrame", filled: 7 }, { x: 0.5, y: 1.3, w: 9, h: 3.8 });
```

| `type` | Fields | Draws |
|---|---|---|
| `tensFrame` / `fiveFrame` | `filled`, `color` | frame sized to the frame's width |
| `doubleTensFrame` | `filledTop`, `filledBottom` | two stacked frames (teen numbers) |
| `dotCard` / `dotCards` | `count` / `counts[]` | dice-pattern cards |
| `numberTrack` | `start`, `end`, `highlight[]` | numbered cells |
| `numberLine` | `start`, `end`, `step`, `labelEvery`, `marked[]` (or `labels[]`) | arrowed line with ticks |
| `fractionStrips` | `strips: [{ denom, shaded, label, color }]`, `showLabels` | separate wholes, gap between |
| `array` | `rows`, `cols` | dot array |
| `baseTen` | `hundreds`, `tens`, `ones` | MAB blocks |
| `groupedCounters` | `groups`, `per` | framed groups |
| `ppwMat` | `whole`, `partA`, `partB` (null = blank) | part-part-whole mat |
| `chips` | `items[]` | row of choice chips |
| `pictogram` | `name`, `label`, `style`, `color` | one large pictogram |
| `pictograms` | `items: ["happy", ...]` or `[{ name, label, color }]` | labelled row |
| `text` | `text`, `fontSize`, `card` | hero numeral or word in a bordered card |
| `table` | `rows[][]`, `header`, `colWidths` | themed data table (see `addDataTable`) |
| `image` | `path`, `frame` | local image, contained |
| `custom` | `draw(slide, frame)` | anything else, still fitted |

`drawVisual` returns the drawn bounds `{ x, y, w, h }`. Unknown types emit a `WARN`, which fails the build gate.

## Pictograms (`addPictogram`)

Simple flat glyphs (Phosphor Icons, MIT) rendered synchronously as a white glyph on a coloured circle or tile, or as a flat glyph in any theme colour. They are for naming a thing on a slide, not for illustration, and they never replace the mathematical or textual representation.

```javascript
addPictogram(slide, "butterfly", x, y, 1.2, { style: "circle", color: C.PRIMARY, label: "butterfly" });
addPictogramRow(slide, 0.5, 2.0, 9, ["happy", "calm", "worried", "sad"]);
keyWordSlide(pres, { word: "evaporate", meaning: "...", pictogram: "sun" }, notes, footer);
cycleDiagramSlide(..., [{ label: "Evaporation", detail: "...", icon: "sun" }, ...], ...);
```

`listPictograms()` returns every accepted name; the full sheet renders in `output/Visual_Catalogue/`. An unknown name emits `WARN [pictogram]` and fails the gate on purpose (a missing picture is invisible in the file). Styles: `circle` (default), `tile`, `flat`. Subject glyphs used on title and closing slides live in `SUBJECT_PICTOGRAMS`.

## Visual Anchor Helpers (ALL subjects, MANDATORY for their representation)

Every theme object carries these manipulative/visual helpers, sized by grade band.
**Never hand-draw a representation this table covers** — if a lesson needs a variant
the helper cannot draw, extend the helper in `themes/core/manipulatives.js` first.
Rendered reference: build `builds/build_visual_catalogue.js` and preview
`output/Visual_Catalogue/`.

| Representation | Helper | Signature |
|---|---|---|
| Tens frame | `addTensFrame` | `(slide, x, y, w, filled, opts)` |
| Five frame | `addFiveFrame` | `(slide, x, y, w, filled, opts)` |
| Dot card (subitising) | `addDotCard` | `(slide, x, y, size, count, opts)` |
| Number track | `addNumberTrack` | `(slide, x, y, w, start, end, highlight[], opts)` |
| Number line | `addNumberLine` | `(slide, x, y, w, labels[], markedPositions[], opts)` — use `""` for unlabelled ticks; on every subject (moved from numeracy in Sept 2026); `opts.markColor`, `opts.markSize` |
| Fraction strips (separate wholes) | `addFractionStripSet` | `(slide, x, y, w, h, [{denom, shaded, label, color}], opts)` |
| Array (rows x cols dots) | `addArray` | `(slide, x, y, rows, cols, opts)` |
| MAB / base-10 blocks | `addBaseTenBlocks` | `(slide, x, y, hundreds, tens, ones, opts)` |
| Choice chips (fractions, numbers, words) | `addChipRow` | `(slide, x, y, w, items[], opts)` — never inline text spaced with runs of spaces |
| Grouped counters ("groups of") | `addGroupedCounters` | `(slide, x, y, groups, per, opts)` |
| Part-part-whole mat | `addPartPartWholeMat` | `(slide, x, y, w, h, {whole, partA, partB}, opts)` — pass `null` for blank boxes |
| Answer reveal bar | `addRevealAnswerBar` | `(slide, answers, {y, h, fontSize, ...})` — use inside `withReveal` revealFn |

PDF twins in `themes/pdf_helpers.js` so worksheets and scaffolds show the SAME
representation on paper: `addTenFramePdf(doc, x, y, filled, opts)`,
`addFractionStripsPdf(doc, y, count, parts, opts)`, `addNumberLinePdf(doc, y, end, tickDenom, opts)`,
`addPpwMatPdf(doc, y, opts)`. Never hand-draw these with raw `doc.rect`/`doc.moveTo`
in a build script.

## Base Builder Signatures

| Function | Signature |
|----------|-----------|
| `titleSlide` | `(pres, title, subtitle, meta, notes, opts)` — `opts.visual` puts the lesson's visual anchor on the right instead of the subject glyph; `opts.glyph` picks another pictogram |
| `liSlide` | `(pres, liItems, scItems, notes, footer)` |
| `contentSlide` | `(pres, badgeText, badgeColor, title, bullets, notes, footer, drawRight)` — `bullets` may be one string; `drawRight` may be a callback or a visual spec; 1-3 short lines render as a hero panel |
| `cfuSlide` | `(pres, badgeText, title, technique, questionText, notes, footer)` |
| `closingSlide` | `(pres, reflectionPrompt, takeaways, notes)` |
| `annotatedModelSlide` | `(pres, badgeText, title, prompts, modelTitle, features, notes, footer, opts)` |
| `compareVisualSlide` | `(pres, badgeText, title, promptText, leftModel, rightModel, notes, footer, opts)` |
| `heroVisualSlide` | `(pres, badgeText, title, visualSpec, notes, footer, { label, prompt, badgeColor, panel })` — visual-only teaching slide; the spec fills a soft panel |
| `choiceSlide` | `(pres, badgeText, title, prompt, options[], notes, footer, { badgeColor, letters })` — 2-4 lettered option cards, each `{ visual, text, caption }`; returns slide with `choiceFrames` |
| `markChoice` | `(slide, index, { color })` — SUCCESS border + tick on the correct option; call inside a `clickBuild` step or `withReveal` revealFn |
| `youDoSlide` | `(pres, title, task, steps[], notes, footer, { where, visual, frame, badgeText, badgeColor, visualLabel })` — task is the hero; First/Next/Then chips; optional mini model |
| `textExtractSlide` | `(pres, badgeText, title, extract, notes, footer, { highlights[], source, prompt })` — reading panel with marker-highlighted phrases |
| `keyWordSlide` | `(pres, { word, meaning, example, routine, pictogram, image }, notes, footer)` — one word per slide; `pictogram` or `image` supplies the required graphic |
| `addDataTable` | `(slide, x, y, w, rows[][], { header, colWidths, fontSize, zebra })` — themed table |

Useful subject-specific signatures:

- Science: `processFlowSlide(pres, badgeText, title, promptItems, steps, notes, footer, opts)`
- Science: `cycleDiagramSlide(pres, badgeText, title, promptTitle, promptLines, centerLabel, steps, notes, footer)` — each step `{ label, detail, icon }`; `icon` is a pictogram name drawn in the stage chip
- Science: `processFlowSlide` steps also accept `icon`
- Numeracy: `workedExSlide` and `dailyReviewSlide` accept a visual spec in the `drawRight` slot

`contentSlide()` now sizes its main card to the amount of content instead of always stretching to the full safe height. Use it for standard content blocks, but if the slide is fundamentally a sequence, system, cycle, or journey, prefer a process/diagram layout rather than bullets alone.

For literacy topics that are fundamentally about noticing features in a source, structure, poster, advertisement, article layout, or visual evidence, prefer `annotatedModelSlide(...)` or a local instructional image rather than a plain bullet list. If students need to look at parts, labels, or evidence, the slide should show those parts visually.

Use a built visual mockup for structure/feature lessons. Use an actual local image when students are meant to infer from or analyse a real photograph, map, artefact, poster, illustration, or source document.

For structure/layout lessons, the default mockup style should be a clean wireframe that makes hierarchy, navigation, and information placement obvious. Do not simulate scenic artwork or faux photography unless the image itself is the thing students are meant to interpret.

For visual-analysis lessons, keep the visual object present into We Do if students are still analysing that visual. The normal fade is labelled visual -> unlabelled visual -> student-created or independent application. Do not fade from visual analysis to prose description unless the instructional target has genuinely shifted away from the visual itself.

When students need to compare two designed visuals, prefer `compareVisualSlide(...)` over custom text-description cards. It is designed for We Do comparison of posters, advertisements, article layouts, and similar side-by-side visual analysis tasks.

For poster, advertisement, article-layout, and similar designed-visual lessons, the preview itself must look like the designed object. Do not use placeholder text such as `Image: ...`, `Colour scheme: ...`, or `Layout: ...` inside the preview area where students are meant to infer from visual evidence. Use a structured mockup or a real local image instead.

`annotatedModelSlide(...)` and `compareVisualSlide(...)` both accept either the legacy `previewBlocks` array or a richer `previewSpec` object. Use `previewSpec` for poster/layout/infographic analysis when the preview needs to render a schematic visual rather than stacked text blocks. For newspaper front page, article layout, poster, infographic, and similar designed-visual I Do slides, always prefer `previewSpec` over flat `previewBlocks`. Do not downgrade content to flat text to work around a builder limitation; fix the shared builder layer instead.

**previewSpec consistency rule:** If a build script defines a structured mockup spec object (an object with a `components` array) for a designed visual, every builder call in the same lesson that renders that visual MUST use `previewSpec`, not `previewBlocks`. Do not define a spec and then pass `previewBlocks` to `annotatedModelSlide(...)` while passing `previewSpec` to `compareVisualSlide(...)` for the same visual — this creates an inconsistent visual fidelity between I Do and We Do. If the shared rendering path cannot handle the spec, fix the shared layer rather than downgrading the lesson content. A regression check at `tests/test_previewspec_consistency.js` scans build scripts for this mismatch.

For science topics that involve ordered systems or journeys, prefer `processFlowSlide(...)` over manual prompt-plus-list layouts. It is designed for digestive journeys, food chains, and similar content where order is part of the concept.

For science topics that are fundamentally cyclical, prefer `cycleDiagramSlide(...)` over manual text-plus-arrow layouts. It is designed for water cycles, life cycles, rock cycles, seasons, and similar content where the loop structure itself needs to be visible.

## Standardised Palette Schema

Every palette uses **semantic colour keys** (never topic-specific names like `C.MIDNIGHT`):

```
PRIMARY        — 60-70% weight: title bg, top bars, badges
SECONDARY      — 20-30%: alternate accents, card strips
ACCENT         — Highlight: SC cards, decorative elements
ALERT          — CFU / emphasis colour
SUCCESS        — Correct / enabling / You Do
ASSESS         — Exit ticket (optional, falls back to ALERT)
BG_DARK        — Title/closing slide background
BG_LIGHT       — Content slide background (cream/off-white)
BG_CARD        — Card fill (white)
WHITE          — Pure white ("FFFFFF")
CHARCOAL       — Body text on light backgrounds
MUTED          — Captions, footers
TEXT_ON_DARK   — Text on dark backgrounds
SUBTITLE       — Subtitle text on title slides
DECOR_1        — Decorative shape colour 1
DECOR_2        — Decorative shape colour 2
FONT_H         — Heading font name
FONT_B         — Body font name
```

Backward-compatible aliases are added by the factory: `C.NAVY → C.PRIMARY`, `C.CREAM → C.BG_LIGHT`, `C.TEAL → C.SECONDARY`.

Derived at theme creation (not stored in the palette files): `PRIMARY_SOFT`, `SECONDARY_SOFT`, `ACCENT_SOFT`, `ALERT_SOFT`, `SUCCESS_SOFT`, `ASSESS_SOFT` (card-fill washes), the matching `*_LINE` hairline tones, and `BG_DARK_PANEL` (a lighter panel for use on the dark title background). `DECOR_1` / `DECOR_2` are kept for compatibility but no longer drawn.

**Retuned palettes (September 2026).** The original palettes were authored far darker than the contrast floor needed (Foundation `PRIMARY` sat at 12-18:1 against white). `scripts/retune_palettes.js` lifts every role colour to the brightest shade of its hue that clears a per-band target and rewrites the palette files in place, preserving comments. Edit hues freely, then re-run the script; never hand-tune a colour below the floor.

## Year Level Font Pairings

| Level | Header Font | Body Font | Design Feel |
|-------|------------|-----------|-------------|
| Foundation | Arial Black | Calibri | Bold, saturated, playful |
| Grade 1 | Arial Black | Calibri | Bold, slightly softer |
| Grade 2 | Trebuchet MS | Calibri | Transitional |
| Grade 3/4 | Trebuchet MS or Georgia | Calibri | Balanced, sophisticated accents |
| Grade 5/6 | Georgia | Calibri | Sophisticated, muted, literary |

## 6 Weekly Variants Per Subject

Each subject has 6 named colour families that rotate weekly. Example for Literacy:

| Variant | Name | Feel |
|---------|------|------|
| 0 | Midnight Scholar | Deep blue / slate / dark gold / crimson |
| 1 | Plum & Honey | Plum / teal / dark honey / coral |
| 2 | Olive & Parchment | Deep olive / burgundy / dark gold / slate |
| 3 | Ink & Paper | Navy ink / charcoal blue / copper / dark red |
| 4 | Autumn Library | Deep brown / forest green / amber / maroon |
| 5 | Twilight Pages | Deep indigo / dusty rose / dark gold / teal |

## Theme Object Exports

The `createTheme()` return object includes everything a build script needs:

- **Palette:** `C`, `FONT_H`, `FONT_B`
- **Shadow factories:** `makeShadow`, `makeCardShadow`
- **Layout constants:** `SLIDE_W`, `SLIDE_H`, `SAFE_RIGHT`, `SAFE_BOTTOM`, `CONTENT_TOP`
- **Contrast utilities:** `hexToRgb`, `luminance`, `contrastRatio`, `validateContrast`, `getContrastColor`
- **Bounds validation:** `validateBounds`
- **Slide diagnostics:** `warnIfSlideHasOverlaps`, `warnIfSlideElementsOutOfBounds`, `runSlideDiagnostics`
- **Icon rendering:** `iconToBase64Png`
- **Element helpers:** `addTopBar`, `addBadge`, `addTitle`, `addCard` (variants white / tint / outline), `addInstructionCard`, `addFooter`, `addIconCircle`, `addTextOnShape`, `softOf`, `lineOf`
- **Colour tools:** `mixHex`, `lightenHex`, `darkenHex`
- **Image helpers:** `addImageWithCaption`, `addInstructionalImageCard`
- **Pictograms:** `addPictogram`, `addPictogramRow`, `listPictograms`, `PICTOGRAMS`, `renderPictogramPng`
- **Visual specs:** `drawVisual`, `isVisualSpec`, `addDataTable`, `VISUAL_TYPES`
- **Click-to-reveal:** `clickBuild` (preferred), `withReveal` (fallback)
- **Base builders:** `titleSlide`, `liSlide`, `contentSlide`, `cfuSlide`, `closingSlide`, `keyWordSlide`, `exitTicketSlide`, `boardBuildSlide`, `annotatedModelSlide`, `compareVisualSlide`, `heroVisualSlide`, `choiceSlide`, `markChoice`, `youDoSlide`, `textExtractSlide`
- **Subject-specific builders:** varies by subject (see table above)
- **Metadata:** `_subject`, `_yearLevel`, `_variant`, `_paletteName`

## Literacy Visual Cases

Treat these as visual-anchor cases by default:

- nonfiction lessons that compare text with a map, poster, photograph, artefact, diagram, or timeline
- lessons on persuasive advertisements, posters, and public-message texts
- lessons on article structure, newspaper features, captions, or text-feature spotting
- lessons where students must infer from a source or notice labelled parts of a model/example

Preferred tools for these cases:

- `annotatedModelSlide(...)` for built-in mockups, labelled source features, or structure-spotting
- `addInstructionalImageCard(...)` when a local photo, map, artefact, or source image genuinely teaches something
- `contentSlide(..., drawRight)` only when the right-hand visual is meaningful and not just decorative whitespace

## Adding a New Subject

1. Create `themes/builders/<subject>.js` — export `create<Subject>Builders(C, FONT_H, FONT_B, el)`
2. Create `themes/palettes/<subject>.js` — export `{ palettes }` with 5 year levels × 6 variants
3. Register in `themes/factory.js`: add to `SUBJECT_PALETTES` and `SUBJECT_BUILDER_FACTORIES`

**Subject builder overrides:** The factory spreads `...subjectBuilders` after `...base`, so any exported name that collides with a base builder silently replaces it. Subject builders must NOT re-export a stale copy of a base builder. If a subject override does not add genuine subject-specific behaviour that the base version cannot provide, remove it so the shared base version is used. A regression guard exists at `tests/test_no_stale_builder_overrides.js` — run it after adding or modifying subject builders.

## Adding New Variants

Append palette objects to the relevant year level array in `themes/palettes/<subject>.js` and update `VARIANTS_PER_LEVEL` in `themes/factory.js`.

## Adding a New Year Level

Add the key to `VALID_YEAR_LEVELS` in `themes/factory.js` and add corresponding entries in each palette file.

---

## Click-to-Reveal (`withReveal`)

Teachers often need to hide answers until after students have responded (CFU checks, problem pairs, worked example solutions). Since PptxGenJS has no animation API, we use **duplicate slides**: slide 1 shows the question, slide 2 shows question + answer. Clicking "next" in PowerPoint reveals the answer. The teacher experience is identical to a click-to-reveal animation.

### API

```javascript
withReveal(buildFn, revealFn)
```

- **`buildFn`** — zero-arg function that calls a slide builder and returns the slide. Called twice internally (once for the question slide, once for the answer slide).
- **`revealFn`** — `callback(slide)` that adds the answer/reveal content to the second slide.
- **Returns** the answer slide (the second slide).

### Usage

```javascript
// CFU slide with hidden answer
withReveal(
  () => cfuSlide(pres, "Check", "Quick Check", "Show Me Boards",
                 "What is 3 × 4?", notes, footer),
  (slide) => {
    addTextOnShape(slide, "Answer: 12", {
      x: 3.5, y: 4.2, w: 3, h: 0.6, rectRadius: 0.08,
      fill: { color: C.SECONDARY },
    }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
  }
);
// Creates 2 slides: question-only → question + answer

// Content slide with hidden definition
withReveal(
  () => contentSlide(pres, "Vocabulary", C.PRIMARY, "Key Term: Equivalent",
                     ["What does 'equivalent' mean in maths?"], notes, footer),
  (slide) => {
    addCard(slide, 0.5, 3.0, 9, 1.5, { strip: C.SECONDARY });
    slide.addText("Equivalent means equal in value, even if represented differently.\ne.g. 3/6 = 1/2", {
      x: 0.75, y: 3.15, w: 8.5, h: 1.2,
      fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });
  }
);
```

### When to Use

| Slide Type | Reveal Content | Use `withReveal`? |
|------------|---------------|-------------------|
| CFU slides | Expected student response / answer | **Yes** |
| We Do worked examples | Steps/answers the teacher cold calls for | **Yes** — students read instead of think if answers are visible |
| Problem pairs (We Do) | The solution | **Yes** |
| Hinge questions | Answer + explanation | **Yes** |
| Worked example (I Do) | Teacher narrates live — visual anchor | No — teacher is telling, not asking |
| Exit ticket (You Do) | Students work independently | No |
| Title / LI / Closing | No hidden content | No |

### Notes

- The reveal slide must carry its OWN post-reveal notes: pass `{ revealNotes: composeRevealNotes({ answer, beats, prep }) }` as the third argument. Without it the pipeline derives notes and prints an ADVISORY; consecutive identical notes fail the build gate.
- `clickBuild(slide, [step, ...])` is the preferred reveal mechanism (one slide, real entrance animations); `withReveal` is the fallback for an answer slide that needs a genuinely different layout.
- The question slide and answer slide are consecutive — no other slides should be inserted between them.
- The `revealFn` callback receives the full PptxGenJS slide object — you can add any element (text, shapes, images, charts).
- Available on every theme object: `T.withReveal` (or destructure as `withReveal` from `createTheme()`).

---

## Defensive Layout Helpers

These helpers are built into every theme object returned by `createTheme()`. They prevent common visual errors at build time.

### Bounds Validation

`addCard` and numeracy-specific helpers (`addPlaceValueChart`, `addNumberLine`, `addTenthsStrip`, `addAreaModel`) validate bounds automatically and print console warnings during build if elements overflow.

**Console warnings during build = layout bugs to fix.** Never ship a presentation with warnings.

### Slide Diagnostics

For manual/custom slides, the theme exposes slide-level diagnostics:

```javascript
const { runSlideDiagnostics } = T;
runSlideDiagnostics(slide, pres);
```

- `warnIfSlideHasOverlaps(slide, pres, opts)` catches unintended text/image collisions.
- `warnIfSlideElementsOutOfBounds(slide, pres, opts)` flags items outside the canvas or below the safe content zone.
- `runSlideDiagnostics(slide, pres, opts)` runs both checks together.

Use these on custom layouts before shipping. The overlap checker is intentionally conservative and focuses on text/image collisions so normal text-on-card layouts do not produce noise.

### Routine Badges

`addRoutineBadge(slide, routineKey, x, y, opts)` draws a classroom-routine icon in a coloured circle with a label (`miniWhiteboard`, `partnerTalk`, `thumbsUp`, `exitTicket`, ...). It is **synchronous** since September 2026; `await`-ing it is harmless but no longer needed. Before that it was async, and a forgotten `await` shipped a slide with no icon and no warning.

### Image Helpers

Use local lesson-cached or unit-cached assets only. The theme exposes:

```javascript
const { addImageWithCaption, addInstructionalImageCard } = T;

addImageWithCaption(slide, imagePath, {
  x: 5.7, y: 1.5, w: 3.2, h: 2.4,
  fit: "crop",
  caption: "Map of the local area",
  sourceLabel: "Source: local council",
});
```

- `addImageWithCaption(...)` places a local image with safe crop/contain sizing and optional caption/source label.
- `addInstructionalImageCard(...)` wraps that image in a theme card for use on content slides.
- These helpers do not fetch images from the web or manage a global asset library.

### Numeracy Visual Helpers

Available on numeracy themes only (via `createTheme("numeracy", ...)`):

- `addPlaceValueChart(slide, x, y, headers, values, opts)` — auto-sizing PV chart. Pass `{ totalW: 4.2 }`.
- `addNumberLine` is now on EVERY theme (see the shared visual anchor table above).
- `addAreaModel(slide, x, y, w, h, rows, cols, opts)` — grid-based area model.
- `addTenthsStrip(slide, x, y, w, h, filled, opts)` — tenths strip visual.
- `addDecimalDot(slide, geo, colIndex, opts)` — decimal dot positioned from chart geometry.

### Text on Shapes — `addTextOnShape`

**Always use this instead of separate addShape + addText calls.** Available on every theme. Guarantees `valign:"middle"`, `align:"center"`, `margin:0` and validates contrast:

```javascript
addTextOnShape(slide, "24 812", {
  x: 1, y: 2, w: 3, h: 0.5, rectRadius: 0.08,
  fill: { color: C.PRIMARY },
}, {
  fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true,
});
```

### Contrast Validation

Available on every theme object:

```javascript
// Auto-pick WHITE or CHARCOAL for a given background
const textColor = getContrastColor(someHexColor);

// Manual check — warns to console if contrast < 4.5:1 (WCAG AA)
validateContrast(textColor, bgColor, "my label badge");
```

**`addTextOnShape` runs contrast validation automatically.** For manual `addText` calls on coloured backgrounds, call `validateContrast` yourself or use `getContrastColor` to pick the text colour.
