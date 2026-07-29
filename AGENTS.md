# AGENTS.md

This file used to be a full copy of the project instructions. It drifted 227
lines behind `CLAUDE.md` and was still teaching the pre-v11 sectioned
`SAY:` / `DO:` teacher-notes format, which the pipeline stopped accepting at
v11.0. Anything reading it got contradicted guidance.

It is now a pointer, so there is one source of truth instead of two copies that
disagree.

## Read these instead

- **`CLAUDE.md`** - project instructions: build commands, PptxGenJS rules,
  teacher notes rules, layout safety, QA requirements, multi-session delivery.
  Start here.
- **`IMPORTANT/MEGA_PROMPT.md`** - the pedagogical framework and the runtime
  prompt for building a lesson. The authority on lesson design.
- **`IMPORTANT/TEACHER_NOTES.md`** - adding notes to an existing deck.
- **`IMPORTANT/OG_MEGA_PROMPT.md`** - Orton-Gillingham decks, which use a
  separate template-locked pipeline.
- **`docs/`** - deep reference: theme system, resource system, PptxGenJS API,
  design guide.

## If you are tempted to add project rules here

Put them in `CLAUDE.md`. A rule that lives in two files is a rule that will be
wrong in one of them.
