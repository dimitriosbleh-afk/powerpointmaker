# Archived lesson scripts (pre v12.6)

These are the 300 per-lesson build scripts written before the September 2026
theme redesign, plus their shared libs and unit manifests. They are kept for
history and for regenerating an old deck if a teacher needs one. They are
NOT exemplars: they hand-place visuals at small sizes, use definition-list
vocabulary, and many exceed the current notes budgets (32 of 299 fail the
build gate). Do not copy patterns from here.

The live exemplars are the spec files in `builds/` (see
`docs/lesson-spec.md` and megaprompt section 15j).

Scripts still build from this folder: `_archive/themes/` forwards their
`../themes/...` requires to the live theme. Output lands wherever the script
resolves `output/` (most write to the repo's `output/`). Example:

    node scripts/build_and_check.js _archive/lessons/build_mk10_lesson1.js

Unit manifests in `manifests/` reference `builds/build_*.js` paths and would
need those paths rewritten to `_archive/lessons/` before `build_unit.py` can
run them again.
