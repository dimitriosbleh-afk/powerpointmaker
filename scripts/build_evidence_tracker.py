"""Build the Auslan evidence tracker workbook for a unit.

One sheet per class. The columns are the ACARA F-6 (L2) Skills and Knowledge
Checklist rows for the cohort's year levels, because those rows ARE the report
descriptors: the same five per year, every semester, rated Emerging,
Consolidating or Proficient. The sheet is therefore the report as a class grid,
and it imports into iDoceo unchanged.

It also carries the hidden exit-ticket rotation, so the teacher knows which
quarter of the class is due each week without opening a second document, and a
Key sheet holding the unit's evidence map in full.

    python scripts/build_evidence_tracker.py \
        --unit "Deaf Sports in Australia" --cohort Enrichment --years 5 6 \
        --classes 5A:25 5B:25 6A:26 --map map.json \
        --out "output/.../Deaf Sports Enrichment Evidence Tracker.xlsx"

The checklist rows come from the teacher's own document, parsed from
reference/auslan/acara/acara_checklist_F-6_L2.md. If that file is not there the
script stops rather than inventing rows.

The evidence map (--map) is section 13.6 of the unit document as JSON:

    {"Y5.1": {"lessons": "Lessons 2, 5 and 6", "evidence": "Piece 1",
              "acara": "AC9L2AU6C02", "vic": "VCASFC164",
              "emerging": "...", "consolidating": "...", "proficient": "..."}}

Without it the Key sheet still lists the rows, with the map columns blank.
"""

import argparse
import json
import os
import re
import sys

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CHECKLIST = os.path.join(ROOT, "reference", "auslan", "acara",
                         "acara_checklist_F-6_L2.md")

LEVELS = ["E", "C", "P"]
LEVEL_NOTE = "E = Emerging, C = Consolidating, P = Proficient"
HEAD_FILL = PatternFill("solid", fgColor="DCE4F0")
BAND_FILL = PatternFill("solid", fgColor="EEF3FA")


def parse_checklist(path):
    """Return {year_key: [(row_id, statement), ...]} from the teacher's checklist."""
    if not os.path.exists(path):
        sys.exit(
            "The ACARA checklist is not on disk: {}\n"
            "It is the teacher's document and assets are gitignored. Put it back "
            "before building a tracker; the rows must not be invented.".format(path)
        )
    text = open(path, encoding="utf-8").read()
    years = {}
    blocks = re.split(r"^# ", text, flags=re.M)
    for b in blocks:
        m = re.match(r"(Foundation|Year (\d)), L2 ", b)
        if not m:
            continue
        key = "F" if m.group(1) == "Foundation" else "Y" + m.group(2)
        rows = []
        for line in b.split("\n"):
            if not line.startswith("|"):
                continue
            first = line.strip().strip("|").split("|")[0].strip()
            if not first or first.startswith("---") or first == "Skill / Knowledge Area":
                continue
            rows.append(first)
        years[key] = [("{}.{}".format(key, i + 1), s) for i, s in enumerate(rows)]
    if not years:
        sys.exit("No year blocks found in {}".format(path))
    return years


def rotation_plan(n_groups, lessons, override=None, first_override=None):
    """Which group is due in which lesson, and when each comes round again.

    A quarter of the class in each of the first n_groups lessons, then the
    groups come round again over whatever lessons are left, so every student is
    seen at least twice before evidence closes.

    A unit that skips a rotation in one lesson, because another piece of
    evidence is running that week, passes `override` and the guess is not used.
    """
    first = dict(first_override) if first_override else {}
    for g in range(1, n_groups + 1):
        first.setdefault(g, g)
    if override:
        second = dict(override)
        for g in range(1, n_groups + 1):
            second.setdefault(g, lessons)
        return first, second

    later = [w for w in range(n_groups + 1, lessons + 1)] or [lessons]
    second = {}
    for i in range(n_groups):
        # Spread the groups as evenly as the remaining lessons allow.
        second[i + 1] = later[min(len(later) - 1, i * len(later) // n_groups)]
    return first, second


def parse_second_pass(text):
    """1:5,2:5,3:7,4:7 -> {1: 5, 2: 5, 3: 7, 4: 7}

    Also used for --first-pass, which has the same shape.
    """
    out = {}
    for part in text.split(","):
        part = part.strip()
        if not part:
            continue
        g, _, w = part.partition(":")
        out[int(g)] = int(w)
    return out


def build(args):
    years = parse_checklist(CHECKLIST)
    wanted = ["F" if y.upper() in ("F", "FOUNDATION") else "Y" + str(y) for y in args.years]
    missing = [y for y in wanted if y not in years]
    if missing:
        sys.exit("The checklist has no rows for: {}".format(", ".join(missing)))
    rows = [r for y in wanted for r in years[y]]

    emap = {}
    if args.map:
        with open(args.map, encoding="utf-8") as fh:
            emap = json.load(fh)

    classes = []
    for spec in args.classes:
        name, _, size = spec.partition(":")
        classes.append((name.strip(), int(size or 25)))

    override = parse_second_pass(args.second_pass) if args.second_pass else None
    first_override = parse_second_pass(args.first_pass) if args.first_pass else None
    first_pass, second_pass = rotation_plan(args.groups, args.lessons, override, first_override)

    wb = Workbook()
    wb.remove(wb.active)

    dv_levels = None
    for cls_name, size in classes:
        ws = wb.create_sheet(cls_name[:31])
        headers = ["Student", "First rotation", "Second rotation"]
        headers += [rid for rid, _ in rows]
        headers += ["Piece 1 filmed", "Piece 2 score", "Cultural reflections",
                    "Exit rotation dates", "Notes"]
        ws.append(headers)

        for c, head in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=c)
            cell.font = Font(bold=True)
            cell.fill = HEAD_FILL
            cell.alignment = Alignment(wrap_text=True, vertical="center")
        # The full checklist statement lives in a comment on its header, so the
        # column can stay three characters wide and still be readable.
        for i, (rid, statement) in enumerate(rows):
            cell = ws.cell(row=1, column=4 + i)
            cell.comment = Comment("{}\n\n{}".format(statement, LEVEL_NOTE), "Auslan tracker")

        for i in range(size):
            group = (i % args.groups) + 1
            ws.append([None, first_pass[group], second_pass[group]])

        ws.freeze_panes = "B2"
        ws.column_dimensions["A"].width = 26
        ws.column_dimensions["B"].width = 9
        ws.column_dimensions["C"].width = 9
        for i in range(len(rows)):
            ws.column_dimensions[get_column_letter(4 + i)].width = 7
        tail = 4 + len(rows)
        for w, col in zip((13, 12, 13, 18, 34), range(tail, tail + 5)):
            ws.column_dimensions[get_column_letter(col)].width = w
        ws.row_dimensions[1].height = 42

        last_row = 1 + size
        dv_levels = DataValidation(
            type="list", formula1='"{}"'.format(",".join(LEVELS)), allow_blank=True
        )
        dv_levels.error = LEVEL_NOTE
        dv_levels.prompt = LEVEL_NOTE
        ws.add_data_validation(dv_levels)
        first_col = get_column_letter(4)
        last_col = get_column_letter(4 + len(rows) - 1)
        dv_levels.add("{}2:{}{}".format(first_col, last_col, last_row))
        piece1 = get_column_letter(tail)
        dv_levels.add("{}2:{}{}".format(piece1, piece1, last_row))

        for r in range(2, last_row + 1):
            if r % 2 == 0:
                for c in range(1, len(headers) + 1):
                    ws.cell(row=r, column=c).fill = BAND_FILL

    key = wb.create_sheet("Key")
    key.append(["{} | {} | rated E, C or P".format(args.unit, args.cohort)])
    key.cell(row=1, column=1).font = Font(bold=True, size=13)
    key.append([])
    key.append(["These rows are the report descriptors. They are the ACARA F-6 (L2) "
                "Skills and Knowledge Checklist for the student's year, the same five "
                "every semester, so parents track the same skills."])
    key.append([LEVEL_NOTE])
    key.append(["Designed to import into iDoceo unchanged."])
    key.append([])
    head = ["Row", "Checklist statement", "Where it is taught", "Evidence",
            "Emerging", "Consolidating", "Proficient", "ACARA", "Victorian"]
    key.append(head)
    hrow = key.max_row
    for c in range(1, len(head) + 1):
        cell = key.cell(row=hrow, column=c)
        cell.font = Font(bold=True)
        cell.fill = HEAD_FILL
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    for rid, statement in rows:
        m = emap.get(rid, {})
        key.append([rid, statement, m.get("lessons", ""), m.get("evidence", ""),
                    m.get("emerging", ""), m.get("consolidating", ""),
                    m.get("proficient", ""), m.get("acara", ""), m.get("vic", "")])
    for col, width in zip("ABCDEFGHI", (7, 48, 20, 22, 30, 30, 30, 15, 13)):
        key.column_dimensions[col].width = width
    for r in range(hrow, key.max_row + 1):
        for c in range(1, len(head) + 1):
            key.cell(row=r, column=c).alignment = Alignment(wrap_text=True, vertical="top")
    key.freeze_panes = "A{}".format(hrow + 1)

    os.makedirs(os.path.dirname(os.path.abspath(args.out)) or ".", exist_ok=True)
    wb.save(args.out)
    print("Written: {}".format(args.out))
    print("  {} class sheet(s), {} checklist rows ({}), rotation over {} lessons in {} groups".format(
        len(classes), len(rows), ", ".join(wanted), args.lessons, args.groups))
    if not emap:
        print("  Key sheet has no evidence map. Pass --map to fill it from section 13.6.")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--unit", required=True)
    ap.add_argument("--cohort", required=True, help="Foundation, Discovery, Challenge or Enrichment")
    ap.add_argument("--years", nargs="+", required=True,
                    help="the year levels in this cohort, for example 5 6")
    ap.add_argument("--classes", nargs="+", required=True,
                    help="NAME:SIZE per class, for example 5A:25 5B:26")
    ap.add_argument("--lessons", type=int, default=7, help="teaching lessons in the unit")
    ap.add_argument("--groups", type=int, default=4, help="rotation groups, a quarter each")
    ap.add_argument("--first-pass",
                    help="when each group is first called, as GROUP:LESSON pairs, "
                         "for example 1:2,2:3,3:4,4:5. Use it when the rotation does "
                         "not start in lesson 1, because that lesson teaches a routine "
                         "rather than vocabulary")
    ap.add_argument("--second-pass",
                    help="when each group comes round again, as GROUP:LESSON pairs, "
                         "for example 1:5,2:5,3:7,4:7. Use it when a lesson has no "
                         "rotation because another piece of evidence runs that week")
    ap.add_argument("--map", help="the section 13.6 evidence map as JSON")
    ap.add_argument("--out", required=True)
    build(ap.parse_args())


if __name__ == "__main__":
    main()
