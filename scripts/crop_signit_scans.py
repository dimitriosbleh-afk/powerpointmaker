"""Cut sign illustrations out of scanned Sign It! pages into the sign bank.

The teacher scans the pages his classes have open on their desks. This turns
those scans into one image per gloss at assets/auslan_signs/signit/<GLOSS>.png,
so a slide can show the same illustration the student is looking at, captioned
with its page number.

It will not guess which illustration is which sign. Auto-detecting picture
blocks on a textbook page is easy; getting one wrong is a whole class taught the
wrong sign for a term. So the boxes are named by a person, once, and after that
the cropping is mechanical and repeatable.

Three steps:

1. Overlay a percent grid on the scan so you can read the boxes off it.

       python scripts/crop_signit_scans.py --page scans/p34.png --grid

   Writes scans/p34_grid.png. Read each illustration's box as left, top, width,
   height in percent of the page.

2. Write those boxes into a crops file, one page at a time. JSON:

       {"page": 34,
        "crops": {"SPORT": [8, 12, 20, 22], "TEAM": [34, 12, 20, 22]}}

   Values of 100 or less are percentages of the page. Larger values are pixels.

3. Cut them.

       python scripts/crop_signit_scans.py --page scans/p34.png --crops p34.json

   Writes assets/auslan_signs/signit/SPORT.png and TEAM.png, records the page
   number in signit_manifest.json, and writes a contact sheet next to the scan
   so you can check every crop at a glance before the build uses them.

Never mirror, flip or recolour a crop: that reverses handedness and teaches the
sign wrong. The script only cuts and pads.

Sign It! is Auslan Hub's copyright. Scans sit under the Australian schools
statutory educational licence, section 113P: internal school teaching only,
never redistributed. assets/ is gitignored for that reason.
"""

import argparse
import json
import os
import sys

from PIL import Image, ImageDraw, ImageFont

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT_DIR = os.path.join(ROOT, "assets", "auslan_signs", "signit")
MANIFEST = os.path.join(ROOT, "assets", "auslan_signs", "signit_manifest.json")

FONT_CANDIDATES = [
    r"C:\Windows\Fonts\arialbd.ttf",
    r"C:\Windows\Fonts\calibrib.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
]


def load_font(px):
    for path in FONT_CANDIDATES:
        if os.path.exists(path):
            return ImageFont.truetype(path, px)
    return ImageFont.load_default()


def write_grid(page_path):
    """Overlay a labelled ten-percent grid so boxes can be read off the scan."""
    im = Image.open(page_path).convert("RGB")
    w, h = im.size
    d = ImageDraw.Draw(im)
    font = load_font(max(11, w // 70))
    for pct in range(0, 101, 5):
        x = round(w * pct / 100.0)
        y = round(h * pct / 100.0)
        major = pct % 10 == 0
        col = (220, 40, 40) if major else (150, 190, 230)
        d.line([(x, 0), (x, h)], fill=col, width=2 if major else 1)
        d.line([(0, y), (w, y)], fill=col, width=2 if major else 1)
        if major and pct:
            d.text((x + 3, 3), str(pct), fill=(220, 40, 40), font=font)
            d.text((3, y + 3), str(pct), fill=(220, 40, 40), font=font)
    out = os.path.splitext(page_path)[0] + "_grid.png"
    im.save(out)
    return out, im.size


def to_pixels(box, size):
    """A box is percentages of the page unless any value is over 100."""
    w, h = size
    x, y, bw, bh = [float(v) for v in box]
    if max(x, y, bw, bh) <= 100:
        return (round(w * x / 100.0), round(h * y / 100.0),
                round(w * bw / 100.0), round(h * bh / 100.0))
    return (round(x), round(y), round(bw), round(bh))


def contact_sheet(crops, out_path, cols=4, cell=260):
    items = sorted(crops.items())
    rows = (len(items) + cols - 1) // cols
    label_h = 30
    sheet = Image.new("RGB", (cols * cell, rows * (cell + label_h)), "white")
    d = ImageDraw.Draw(sheet)
    font = load_font(18)
    for i, (gloss, im) in enumerate(items):
        cx = (i % cols) * cell
        cy = (i // cols) * (cell + label_h)
        thumb = im.copy()
        thumb.thumbnail((cell - 12, cell - 12))
        sheet.paste(thumb, (cx + (cell - thumb.width) // 2, cy + (cell - thumb.height) // 2))
        d.rectangle([cx + 2, cy + 2, cx + cell - 2, cy + cell - 2], outline=(200, 200, 200))
        d.text((cx + 8, cy + cell + 4), gloss, fill=(20, 20, 20), font=font)
    sheet.save(out_path)
    return out_path


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--page", required=True, help="the scanned page image")
    ap.add_argument("--grid", action="store_true",
                    help="write a grid overlay to read boxes off, and stop")
    ap.add_argument("--crops", help="the JSON file of gloss -> box for this page")
    ap.add_argument("--pad", type=float, default=1.5,
                    help="percent of the box added as a margin on every side")
    args = ap.parse_args()

    if not os.path.exists(args.page):
        ap.error("no such page image: {}".format(args.page))

    if args.grid or not args.crops:
        out, size = write_grid(args.page)
        print("Grid written: {}  (page is {}x{} pixels)".format(out, *size))
        print("Read each illustration as left, top, width, height in percent,")
        print("put them in a crops file, then run again with --crops.")
        return

    with open(args.crops, encoding="utf-8") as fh:
        spec = json.load(fh)
    page_no = spec.get("page")
    boxes = spec.get("crops") or {}
    if not boxes:
        ap.error("the crops file has no 'crops' object")
    if page_no is None:
        print("WARNING: no 'page' in the crops file, so no page number is recorded. "
              "Slides caption the illustration with it, so add one.", file=sys.stderr)

    im = Image.open(args.page).convert("RGB")
    os.makedirs(OUT_DIR, exist_ok=True)

    cut = {}
    for gloss, box in boxes.items():
        gloss = gloss.strip().upper()
        x, y, w, h = to_pixels(box, im.size)
        px, py = round(w * args.pad / 100.0), round(h * args.pad / 100.0)
        x0 = max(0, x - px)
        y0 = max(0, y - py)
        x1 = min(im.width, x + w + px)
        y1 = min(im.height, y + h + py)
        if x1 - x0 < 20 or y1 - y0 < 20:
            print("  SKIP {}: box is smaller than 20 pixels".format(gloss), file=sys.stderr)
            continue
        crop = im.crop((x0, y0, x1, y1))
        out = os.path.join(OUT_DIR, "{}.png".format(gloss))
        crop.save(out)
        cut[gloss] = crop
        print("  {:<16} {}x{}  ->  {}".format(gloss, crop.width, crop.height,
                                              os.path.relpath(out, ROOT)))

    if not cut:
        print("Nothing cut.", file=sys.stderr)
        sys.exit(1)

    sheet = contact_sheet(cut, os.path.splitext(args.page)[0] + "_crops.png")

    record = {}
    if os.path.exists(MANIFEST):
        try:
            with open(MANIFEST, encoding="utf-8") as fh:
                record = json.load(fh).get("signs", {})
        except Exception:
            record = {}
    for gloss in cut:
        record[gloss] = {"page": page_no, "file": "{}.png".format(gloss),
                         "scan": os.path.basename(args.page)}
    with open(MANIFEST, "w", encoding="utf-8") as fh:
        json.dump({
            "source": "Sign It! (Auslan Hub), scanned by the teacher",
            "licence": "Australian schools statutory educational licence, s113P. "
                       "Internal school teaching only, never redistributed.",
            "note": "The page number is what a slide captions the illustration with, "
                    "so students can open the same page.",
            "signs": dict(sorted(record.items())),
        }, fh, indent=2)

    print("\n{} illustration(s) cut. Check them: {}".format(len(cut), sheet))
    print("Page numbers recorded in {}".format(os.path.relpath(MANIFEST, ROOT)))


if __name__ == "__main__":
    main()
