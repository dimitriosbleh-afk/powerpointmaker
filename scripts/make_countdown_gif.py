"""Make a countdown timer as an animated GIF, for slides that time a task.

PowerPoint and Google Slides cannot run a live timer without an add-in, so the
timer is a GIF: one frame per second, a ring that empties, and the time left in
the middle. It starts when the slide appears in slideshow, which is why the
teacher notes for a timed slide say "advance to start the clock".

It plays once by default. A countdown that loops back to full is worse than no
timer at all.

Usage:
    python scripts/make_countdown_gif.py 120
    python scripts/make_countdown_gif.py 180 --color 1B3F94
    python scripts/make_countdown_gif.py 120 --color 1B3F94 --bg FAF7F0 --size 320

Writes assets/timers/countdown_<seconds>s_<COLOR>.gif and prints the path. The
file is cached: the same seconds and colour never renders twice, so a build
script can call this for every timed slide without cost.
"""

import argparse
import os
import sys

from PIL import Image, ImageDraw, ImageFont

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT_DIR = os.path.join(ROOT, "assets", "timers")

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


def hex_to_rgb(h):
    h = h.lstrip("#").strip()
    if len(h) != 6:
        raise ValueError("colour must be 6 hex digits with no hash, for example 1B3F94")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def mix(a, b, t):
    return tuple(round(a[i] + (b[i] - a[i]) * t) for i in range(3))


def label(seconds):
    if seconds >= 60:
        return "{}:{:02d}".format(seconds // 60, seconds % 60)
    return str(seconds)


def render(seconds, color, bg, size, alert_at):
    """One frame per second, counting down to zero, plus a held zero frame."""
    fg = hex_to_rgb(color)
    back = hex_to_rgb(bg)
    track = mix(fg, back, 0.82)          # the ring's unfilled part
    alert = (197, 48, 48)

    pad = round(size * 0.05)
    box = (pad, pad, size - pad, size - pad)
    ring = round(size * 0.085)
    inner = (box[0] + ring, box[1] + ring, box[2] - ring, box[3] - ring)
    font = load_font(round(size * 0.30))

    frames = []
    for left in range(seconds, -1, -1):
        im = Image.new("RGB", (size, size), back)
        d = ImageDraw.Draw(im)
        d.ellipse(box, fill=track)
        frac = left / float(seconds) if seconds else 0.0
        if frac > 0:
            # PIL measures from 3 o'clock and sweeps clockwise, so 12 o'clock is
            # -90. The remaining arc ENDS at 12: that way the gap opens at 12 and
            # grows clockwise, which is how a clock drains. Starting the arc at
            # 12 instead makes the gap grow anticlockwise and reads backwards.
            d.pieslice(box, -90 + 360 * (1 - frac), 270, fill=fg)
        d.ellipse(inner, fill=back)

        text = label(left)
        col = alert if (left <= alert_at and left > 0) else fg
        if left == 0:
            col = alert
        l, t, r, b = d.textbbox((0, 0), text, font=font)
        d.text(((size - (r - l)) / 2 - l, (size - (b - t)) / 2 - t), text, font=font, fill=col)
        frames.append(im)

    # One palette across every frame, so only the arc and the digits are
    # re-encoded. The palette has to be learned from frames that between them
    # use every colour: built from the first frame alone it has no red in it,
    # and every alert-coloured number silently quantises to the nearest blue.
    sample_ids = sorted({0, len(frames) // 2, max(0, len(frames) - alert_at - 1), len(frames) - 1})
    samples = [frames[i] for i in sample_ids]
    strip = Image.new("RGB", (size * len(samples), size), back)
    for i, f in enumerate(samples):
        strip.paste(f, (i * size, 0))
    base_pal = strip.quantize(colors=32, method=Image.MEDIANCUT)
    return [f.quantize(palette=base_pal, dither=Image.NONE) for f in frames]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("seconds", type=int, help="how long the task runs")
    ap.add_argument("--color", default="1B3F94",
                    help="ring colour, 6 hex digits, no hash. Pass the theme's "
                         "C.PRIMARY so the timer matches the deck")
    ap.add_argument("--bg", default="FFFFFF", help="background colour behind the ring")
    ap.add_argument("--size", type=int, default=300, help="pixels square")
    ap.add_argument("--alert-at", type=int, default=10,
                    help="seconds left at which the number turns red")
    ap.add_argument("--loop", action="store_true",
                    help="loop forever. Off by default: a countdown that resets "
                         "itself is worse than no timer")
    ap.add_argument("--out", help="write here instead of the cache path")
    ap.add_argument("--force", action="store_true", help="rebuild even if cached")
    args = ap.parse_args()

    if args.seconds < 5 or args.seconds > 1800:
        ap.error("seconds must be between 5 and 1800")

    os.makedirs(OUT_DIR, exist_ok=True)
    out = args.out or os.path.join(
        OUT_DIR, "countdown_{}s_{}.gif".format(args.seconds, args.color.upper().lstrip("#"))
    )
    if os.path.exists(out) and not args.force:
        print(out)
        return

    frames = render(args.seconds, args.color, args.bg, args.size, args.alert_at)
    durations = [1000] * len(frames)
    durations[-1] = 3000                      # hold on zero
    save = dict(save_all=True, append_images=frames[1:], duration=durations,
                optimize=True, disposal=1)
    if args.loop:
        save["loop"] = 0
    frames[0].save(out, **save)
    print(out)
    print("  {} frames, {} KB".format(len(frames), round(os.path.getsize(out) / 1024.0)),
          file=sys.stderr)


if __name__ == "__main__":
    main()
