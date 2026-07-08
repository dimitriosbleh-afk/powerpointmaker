# Theme System Reference

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

All subjects get the **base builders**: `titleSlide`, `liSlide`, `contentSlide`, `cfuSlide`, `closingSlide`, `keyWordSlide`, `exitTicketSlide`, `boardBuildSlide`, `annotatedModelSlide`, `compareVisualSlide`, plus `addRevealAnswerBar`. `keyWordSlide(pres, { word, meaning, example }, notes, footer)` is the vocabulary word card - one word per slide, never definition bullet lists. `cfuSlide` and `exitTicketSlide` are density-aware (short prompts render hero-sized and centred).

| Subject | Extra Builders | Purpose |
|---------|---------------|---------|
| **Literacy** | `vocabSlide`, `quoteSlide`, `modellingSlide`, `pairShareSlide` | Text study, vocabulary, modelled writing |
| **Numeracy** | `workedExSlide`, `dailyReviewSlide`, `fluencySlide`, `addPlaceValueChart`, `addNumberLine`, `addAreaModel`, `addTenthsStrip`, `addDecimalDot`, `addStageBadge` | Worked examples, maths visuals |
| **Inquiry** | `investigationSlide`, `findingsSlide`, `pairShareSlide` | Question-driven, evidence gathering |
| **Wellbeing** | `scenarioSlide`, `reflectionSlide`, `pairShareSlide` | Social scenarios, discussion, reflection |
| **Science** | `experimentSlide`, `observationSlide`, `conclusionSlide`, `processFlowSlide`, `cycleDiagramSlide` | Scientific method structure, ordered processes, systems, cycles |

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
| Number line | `addNumberLine` | `(slide, x, y, w, labels[], markedPositions[], opts)` — use `""` for unlabelled ticks |
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
| `titleSlide` | `(pres, title, subtitle, meta, notes)` |
| `liSlide` | `(pres, liItems, scItems, notes, footer)` |
| `contentSlide` | `(pres, badgeText, badgeColor, title, bullets, notes, footer, drawRight)` |
| `cfuSlide` | `(pres, badgeText, title, technique, questionText, notes, footer)` |
| `closingSlide` | `(pres, reflectionPrompt, takeaways, notes)` |
| `annotatedModelSlide` | `(pres, badgeText, title, prompts, modelTitle, features, notes, footer, opts)` |
| `compareVisualSlide` | `(pres, badgeText, title, promptText, leftModel, rightModel, notes, footer, opts)` |

Useful subject-specific signatures:

- Science: `processFlowSlide(pres, badgeText, title, promptItems, steps, notes, footer, opts)`
- Science: `cycleDiagramSlide(pres, badgeText, title, promptItems, stages, notes, footer, opts)`

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
- **Element helpers:** `addTopBar`, `addBadge`, `addTitle`, `addCard`, `addFooter`, `addIconCircle`, `addTextOnShape`
- **Image helpers:** `addImageWithCaption`, `addInstructionalImageCard`
- **Click-to-reveal:** `withReveal`
- **Base builders:** `titleSlide`, `liSlide`, `contentSlide`, `cfuSlide`, `closingSlide`, `annotatedModelSlide`, `compareVisualSlide`
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

- Both slides share the same teacher notes (the `notes` param is applied to both).
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
- `addNumberLine(slide, x, y, w, labels, opts)` — adaptive label width/font for dense lines.
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
