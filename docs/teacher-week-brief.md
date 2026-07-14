# Teacher Week Brief

The Teacher Week Brief is a one-page preparation PDF for a week or multi-session sequence. It is generated from `teacher_brief` data in the unit manifest and written to the final flat `Resources/` folder.

Its purpose is intellectual preparation. A teacher should be able to read it in about three minutes and know:

- the learning throughline across the week
- the three highest-leverage teaching moves
- what students will do in each session
- the decision-grade check in each session
- the prepared response if students are not ready
- the essential content the teacher must understand
- the most likely misconceptions
- the short section worth rehearsing aloud
- the materials that must be ready

The brief is not a compressed slide deck, a research summary or a replacement for teacher judgement. It must be generated from the actual lesson sequence and its checks for understanding.

## Manifest configuration

Add a `teacher_brief` object beside `unit_folder`, `unit_pptx_name` and `lessons`:

```json
{
  "unit_folder": "Fractions_Week_1",
  "unit_pptx_name": "Fractions Week 1.pptx",
  "teacher_brief": {
    "file_name": "Teacher Week Brief.pdf",
    "title": "Teacher Week Brief",
    "unit": "Fractions as numbers",
    "grade": "Years 3-4",
    "subject": "Mathematics",
    "curriculum": "Mathematics 2.0, Number",
    "throughline": "Students move from naming equal parts to locating, comparing and explaining fractions on shared visual models.",
    "high_leverage_moves": [
      "Pre-cue the question, give think time, then require every student to show an answer.",
      "Keep the fraction strip visible while language and symbols are connected.",
      "Use each check to choose the next move, not simply the next slide."
    ],
    "sessions": [
      {
        "session": 1,
        "focus": "Name unit fractions",
        "students": "Build and label equal parts with strips.",
        "check": "All boards show the denominator as the total equal parts.",
        "respond": "If mixed, rebuild one whole and count every part before rechecking."
      },
      {
        "session": 2,
        "focus": "Locate fractions",
        "students": "Place fractions on a number line.",
        "check": "Students use equal intervals and start from zero.",
        "respond": "If spacing drifts, fold a strip first and transfer the marks."
      }
    ],
    "essential_knowledge": [
      "The denominator names how many equal parts make one whole.",
      "Comparisons are valid only when the wholes are the same size."
    ],
    "misconceptions": [
      "A larger denominator means a larger piece.",
      "The tick marks, rather than the intervals, are counted on a number line."
    ],
    "response_rule": "Secure and complete responses -> move on. Mixed responses -> one more guided pair. A shared misconception -> re-model with a different representation, then recheck everyone.",
    "rehearse": [
      "Practise the first model and the exact hinge question aloud.",
      "Decide where you will stand to scan every board quickly."
    ],
    "materials": [
      "Fraction strips",
      "Mini-whiteboards",
      "Printed practice sheet"
    ]
  },
  "lessons": []
}
```

## Content rules

- Use exactly three high-leverage moves. These are the small set that most strongly influences the week's success.
- Include one session row for every manifest lesson. Session numbers must match exactly.
- Write the check as visible evidence, not an activity. "Boards show equal intervals from zero" is evidence. "Use mini-whiteboards" is only a routine.
- Write the response as a prepared teaching move. Name the different model, guided pair, scaffold or recheck.
- Keep the learning goal fixed when describing a response. Adapt pace, representation or support from evidence.
- Write essential knowledge for the teacher, not student-facing definitions.
- Use plain language. Do not fill the brief with VTLM, HITS or research terminology.
- Keep every field within the generator's length limits. The build fails instead of shrinking the brief into unreadable text.

## Build behaviour

`python scripts/build_unit.py builds/manifests/<unit>.json` now:

1. builds each lesson
2. merges the unit deck and session PDFs
3. generates `Resources/Teacher Week Brief.pdf` when `teacher_brief` is present
4. validates the delivered PDF during merged unit QA

The generator enforces one page. A configuration that cannot fit fails with the field or card that must be shortened.
