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
node scripts/build_and_check.js builds/build_<unit>_lesson<n>.js  # Build + enforce QA gates (diagnostics + markitdown)
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
- Each presentation needs a fresh `new pptxgen()` instance.
- Write files with `await pres.writeFile({ fileName })`, never the deprecated `writeFile("path.pptx")` form.
- Always set `pres.layout = "LAYOUT_16x9"`. NEVER use `"LAYOUT_WIDE"` (wrong dimensions).

## Teacher Notes Rules

- Speaker notes are plain text in PptxGenJS/PowerPoint. NEVER use markdown in teacher notes. Do not use `**bold**`, `_italics_`, backticks, markdown headings, or markdown lists.
- Teacher notes must be ASCII-safe for PowerPoint, PowerPoint for iPad, and exported notes views. Use straight quotes `' "`, hyphen bullets `-`, `...`, `->`, `>=`, `<=`, and `x`. Avoid smart quotes, em dashes, unicode bullets, unicode arrows, and other decorative symbols in notes.
- Keep teacher notes scannable. Prefer short bullets, not dense paragraphs. `SAY` should usually be 2-4 short bullets, `DO` 2-4 short bullets, `TEACHER NOTES` no more than 2 short sentences, and `WATCH FOR` 1-2 bullets.
- `SAY:` is teacher cue language, not script prose. Each bullet should sound like something a teacher could say immediately in class with little or no rewriting.
- Use a classroom-natural voice in `SAY:`: direct prompts, short follow-ups, question-led phrasing, and concise think-aloud cues. Light connective phrasing is allowed when it helps flow, but keep it brief and purposeful.
- Avoid slang, lesson-announcer phrasing, polished exposition, abstract briefing language, and over-explained transitions in `SAY:`. If a bullet sounds more like presenter copy than classroom talk, rewrite it.
- Section headers in notes should be plain uppercase text like `SAY:` and `DO:`. Do not try to force bold with markdown.
- Separate note sections with a blank line. PowerPoint on iPad reads paragraph breaks more reliably than dense text blocks.
- If a slide uses `liSlide()`, the Learning Intention must be a single plain sentence and the Success Criteria must be exactly 3 simple `I can...` bullets. The first success criterion must be ultra-achievable for almost every student.
- Do not add a `PACING OVERVIEW` block to speaker notes by default. If timing guidance is genuinely needed, keep it to one short sentence in `TEACHER NOTES`.
- End notes with at most one short framework/meta tag line. Do not stack multiple checklist tags.
- Use a plain-text `SOURCES:` section whenever a slide includes an external image or a non-trivial externally sourced factual claim.

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

- When using `withReveal` and adding a reveal element (e.g. an answer bar) in the `revealFn`, ensure all content on the slide stops at least 0.15" ABOVE the reveal element's top edge.
- If the reveal bar is at y 4.25, the tallest content text box must end by y 4.1 at most.
- For factor-pair lists, prompts, or other variable-length content above a reveal bar, reduce the text box `h` to enforce this ceiling rather than letting it extend to `SAFE_BOTTOM`.

### Title Sizing

- Long titles that wrap to 2+ lines push content down. When a title exceeds ~45 characters, verify that the content below still fits without overlapping the footer zone.
- For custom slides, use `layoutGuide.panelTopPadded` (available from `contentSlide` and `workedExSlide` callbacks) as the starting y for right-column content when the title is long.
- Prefer concise titles (under 40 chars) for slides with dense two-column layouts.

## Key Conventions

- Palette uses semantic keys: PRIMARY, SECONDARY, ACCENT, ALERT, SUCCESS, BG_DARK, BG_LIGHT, BG_CARD, CHARCOAL, WHITE, MUTED. Backward-compatible aliases exist (C.NAVY, C.CREAM, C.TEAL).
- White text on coloured fills. Dark text on light fills. NEVER same colour for text and its background.
- White icons need a coloured circle background on light surfaces.
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
- Standard footer text is ignored by diagnostics. If a custom footer-like element is falsely flagged, keep diagnostics enabled and use a narrow `ignoreIndices` override rather than disabling safe-bottom checks for the whole slide.
- Theme image helpers are available for local assets: `addImageWithCaption(...)` and `addInstructionalImageCard(...)`.
- `annotatedModelSlide(...)` is available on every theme object for labelled source features, poster/article structure, and "notice this part" teaching. Do not swap subjects just to reach it.
- `compareVisualSlide(...)` is available on every theme object for We Do comparison of two posters, layouts, advertisements, or similar designed visuals.
- Science process/system topics can also use the dedicated `processFlowSlide(...)` builder for ordered journeys, cycles, and body systems.
- Science cycle topics should prefer the dedicated `cycleDiagramSlide(...)` builder over manual text-plus-arrow layouts.
- `withReveal(buildFn, revealFn)` creates duplicate slide pairs for click-to-reveal. Use for CFU answers, We Do solutions, hinge questions. Do NOT use for I Do, exit tickets, or titles.
- Every lesson with companion PDFs gets a resource slide via `addResourceSlide()` from `pdf_helpers.js`.
- Per-lesson build output goes to `output/<LessonFolder>/` - PPTX at the root, companion PDFs in a `resources-session{N}/` subfolder where `N` is the session number within that week's sequence. This is the build step, not the delivery step.
- PptxGenJS hyperlinks use relative paths - include the subfolder prefix (e.g., `resources-session3/Session 3 Worksheet.pdf`).
- Resource names must be teacher-friendly and session-first: `Session 1 Worksheet`, `Session 1 Answer Key`, `Session 2 Enabling Scaffold`.
- Use the same human-readable name on the resource slide and in the PDF filename stem. Avoid codes like `WH4_L16`, `SR1`, `GO1`, `ET_Lesson5`, or similar.
- Do not use day names in resource filenames. Teachers run sessions on different days.
- Do not use underscores in teacher-facing PDF filenames. Use spaces.

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
0. **Use `node scripts/build_and_check.js builds/build_<unit>_lesson<n>.js` as the default build command.** It runs the build, checks for diagnostics errors/warnings, and runs markitdown. If it exits non-zero, the build has failed — fix the issue before proceeding. Do NOT skip this step or ignore its output. **The gate script is the minimum automated bar, not a substitute for visual inspection.** Passing it means the build is structurally sound — it does NOT mean the slides look correct.
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
npm install                                     # Node: pptxgenjs, pdfkit, react-icons, sharp
# LibreOffice (soffice) only needed for optional local image preview
```
