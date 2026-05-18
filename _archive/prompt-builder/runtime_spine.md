# PPTX Runtime Spine

Short runtime reminder for compaction recovery. This does not replace `CLAUDE.md` or `megapromptlean.md`.

## Non-negotiables

- Use `pres.layout = "LAYOUT_16x9"` only. Never use `LAYOUT_WIDE`.
- Keep content inside the 10" x 5.625" layout safety zone. Main content must stay above y 5.1".
- Prefer tested theme builders first. Do not drift into manual `addShape`/`addText` layouts unless the slide is genuinely novel.
- If you do go manual, run slide diagnostics and fix warnings before shipping.
- Every PptxGenJS text run must use a plain string `text` value. Coerce dynamic values with `String(...)`; never pass arrays, numbers, booleans, or objects as `text`.
- Write PPTX files with `await pres.writeFile({ fileName })`, not the deprecated string form.
- Keep theme selection aligned to the explicit `Subject:` field unless the user explicitly changes it.
- Subject builders must NOT shadow base builders with stale forks. If a subject override doesn't add genuine subject-specific behaviour, remove it so the base version is used. The factory spreads `...subjectBuilders` after `...base`, so collisions silently replace the base.
- Treat `builds_archive/` as historical only, not as an active exemplar source. If `builds/` is empty, build from the current theme system and docs instead of reviving archived scripts by default.
- **Two-column layout:** `contentSlide` left card is 4.5" (x 0.5-5.0), right column starts at x 5.2. `workedExSlide` left card is 4.5" (x 0.5-5.0), right column starts at x 5.3. NEVER place right-column elements at x < 5.2. Both builders auto-run diagnostics when `drawRight` is provided.
- **Text box overflow:** PptxGenJS does NOT clip text overflow — content that exceeds the box height renders below the boundary and overlaps elements underneath. Always size text boxes to fit: fontSize 12 ≈ 0.22"/line, fontSize 14 ≈ 0.26"/line, fontSize 16 ≈ 0.30"/line.
- **Reveal clearance:** Content must stop at least 0.15" above any reveal element added in `revealFn`. Do not let text boxes extend to `SAFE_BOTTOM` if a reveal bar will overlay.
- **Title length:** Titles over ~45 chars wrap to 2 lines and push content down. Use concise titles for dense two-column slides. Use `layoutGuide.panelTopPadded` for right-column content when titles are long.

## Notes

- Speaker notes must be plain ASCII-safe text. No markdown, smart punctuation, or decorative unicode.
- Default note sections are `SAY:`, `DO:`, and short supporting sections only when needed.
- Add a plain-text `SOURCES:` section whenever a slide uses an external image or a non-trivial externally sourced factual claim.

## Images

- Images are opt-in instructional tools, not decoration.
- Use local lesson-cached or unit-cached assets only. Do not assume a permanent global image library.
- Prefer source photos, diagrams, maps, labelled visuals, and artefacts that help students understand the content.
- Distinguish `visual anchor` from `actual image`. Mockups/diagrams/builders can satisfy a visual-anchor requirement. Use a real local image when students need to interpret an authentic source such as a photo, map, artefact, poster, illustration, or source document.
- For literacy topics involving source analysis, text features, advertisements, posters, article layout, maps, artefacts, or compare-text-and-visual evidence, use a visual anchor such as `annotatedModelSlide(...)` or a local instructional image on at least one core teaching slide.
- `annotatedModelSlide(...)` is available on every theme object. Keep the theme aligned to `Subject:` and use that shared builder instead of switching subjects to access it.
- For structure/layout lessons, prefer clean wireframe mockups over pseudo-real scenes. Built visuals should make hierarchy and placement obvious rather than pretending to be finished artwork.
- If the literacy lesson is about structure or feature-spotting, a built mockup is usually enough. If it is about analysing or inferring from a real source visual, use an actual local instructional image.
- For visual-analysis lessons, keep the visual object present into We Do if students are still analysing it. Fade labels first, not the poster/map/source/diagram itself.
- If students need to compare two designed visuals in We Do, prefer `compareVisualSlide(...)` or another explicit dual-visual layout instead of text-only descriptions.
- If the lesson uses a built poster/article/advertisement mockup, the mockup must look like the designed object itself. Do not use placeholder lines like `Image: ...` or `Colour scheme: ...` where students are meant to infer from visual evidence.
- For newspaper front page, article layout, poster, infographic, and similar designed-visual I Do slides, prefer a structured `previewSpec` over flat `previewBlocks` when the builder supports it. Do not downgrade content to fit a stale override; fix the shared builder layer instead.
- **previewSpec consistency:** If a build script defines a structured mockup spec for a designed visual, every builder call rendering that visual must use `previewSpec`. Do not define a spec object and then pass `previewBlocks` to one builder while passing `previewSpec` to another in the same lesson. Fix the shared layer instead of downgrading lesson content.
- For science topics involving systems, cycles, sequences, life stages, or body processes, use at least one dedicated visual anchor such as `cycleDiagramSlide(...)`, `processFlowSlide(...)`, a labelled diagram, or an instructional local image.
- Treat persuasive posters, newspaper/article features, source-photo inference, historical nonfiction with maps/artefacts, and similar literacy lessons as visual-anchor cases by default.
- Treat water cycle, life cycle, digestive system, food chains, circuits, and similar topics as mandatory visual cases by default.
- Avoid oversized empty cards. If a slide only has a few short bullets or prompts, use a compact card or a two-column/visual layout instead of stretching text to fill space.
- If an image does not clearly improve learning, skip it.

## QA

0. **Default build command: `node scripts/build_and_check.js builds/build_<unit>_lesson<n>.js`** — runs build, checks diagnostics (zero errors/warnings required), and runs markitdown. Non-zero exit = failed gate. Fix before visual QA. **The gate script is the minimum bar, not the finish line.**
1. **Smoke build early.** Run `build_and_check.js` after writing PPTX code but BEFORE writing companion PDFs. Do not write the entire script in one pass.
2. If the gate script reports markitdown FAIL, treat it as a blocker. Do not dismiss as "intermittent" without concrete evidence.
3. **Visual QA is required after the gate passes.** Run `pptx_to_images.py`, inspect the slides. The gate cannot catch text-box overflow, reveal overlap, or visual imbalance — only eyes can.
4. Optionally create a contact sheet: `python scripts/slide_montage.py`
5. Final compatibility check in Google Slides for delivered decks.

If diagnostics only complain about the footer zone on a custom slide, keep them enabled and use `runSlideDiagnostics(slide, pres, { respectSafeBottom: false })` rather than deleting the diagnostics call.

Do not call QA "passed" unless the Google Slides compatibility check (step 5) is complete. Do not treat a passing gate script or local visual inspection as delivery-ready. If only local checks ran, say local QA passed and Google Slides review is still pending.

When using `contentSlide(..., drawRight)` or numeracy `workedExSlide(..., drawRight)`, use the callback's second `layoutGuide` argument for custom right-column panels. Start custom right-column cards from `layoutGuide.panelTopPadded` by default rather than hardcoding `CONTENT_TOP`.

## Resources

- Resource slide links must use relative paths with the `resources-sessionN/` prefix.
- Teacher-facing PDF names must stay human-readable and session-first.
- For visual-analysis scaffold PDFs, include the visual object on paper too. Text that describes a visual is not a visual scaffold.

## Multi-session delivery

- Multi-session requests (unit, week, "lessons 1 to N") are not "done" until one combined PPTX and one flat `Resources/` folder exist at `output/<unit_folder>/`.
- Workflow: per-lesson build scripts in `builds/` → manifest at `builds/manifests/<unit>.json` → `python scripts/build_unit.py builds/manifests/<unit>.json`. The wrapper runs `build_and_check.js` per lesson then merges. Use `--skip-build` to re-merge after fixing one lesson.
- Single-session requests do not need the merge. The per-lesson folder is the deliverable.
