"""Populate the Auslan sign image bank from Auslan Signbank.

For each gloss, finds the Signbank entry, downloads the sign video, and writes
either an animated GIF of the sign (--gif, the default for new units) or a
still sequence strip, plus a manifest recording exactly which entry the asset
came from, so the teacher can verify every sign before teaching it.

The teacher's verdict on the strips was that they are not clear enough to
reproduce a sign from, so --gif is what decks use. The strip mode is kept for
printables, which cannot animate.

Nothing here is a verified sign. The manifest is the verification worklist.

Usage:
    python scripts/fetch_auslan_signs.py --gif --glosses TEAM SCHOOL AGAIN
    python scripts/fetch_auslan_signs.py --gif --from-file glosses.txt
    python scripts/fetch_auslan_signs.py --from-file glosses.txt --refetch
    python scripts/fetch_auslan_signs.py --gif --links vetted_links.json

A links file is how the teacher's vetted entry wins over a search. It is JSON
({"TEAM": "https://auslan.org.au/dictionary/words/team-1.html"}) or one
`GLOSS <tab or space> url` per line. A gloss listed there is fetched from that
exact entry and marked vetted in the manifest; anything not listed falls back
to the search, and the manifest says so.

Source: Auslan Signbank (auslan.org.au), CC BY-NC-ND 4.0. Stills are extracted
for internal school teaching use under the Australian schools statutory
educational licence. See assets/auslan_signs/README.md.
"""

import argparse
import concurrent.futures as cf
import glob
import html
import json
import os
import re
import sys
import threading
import time
import urllib.parse
import urllib.request

import cv2
import numpy as np
from PIL import Image

BASE = "https://auslan.org.au"
UA = "Mozilla/5.0 (compatible; school Auslan lesson resource builder)"
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT_DIR = os.path.join(ROOT, "assets", "auslan_signs")
CACHE_DIR = os.path.join(ROOT, "tmp", "auslan_video_cache")

# Glosses whose English search term is not just the gloss with hyphens removed.
SEARCH_OVERRIDES = {
    "THANK-YOU": "thank you",
    "SLOW-DOWN": "slow down",
    "FLASHING-LIGHT": "flashing light",
    "CAPTIONS": "caption",
    "FLASHING-LIGHT": "alarm (flashing)",
    "GOODBYE": "bye",
    "FS": None,  # fingerspelling is not a lexical sign
}

STRIP_H = 460          # height of each frame panel in the output strip
PANEL_GAP = 14
MARGIN = 16
MAX_PANELS = 3

# GIF settings, tuned by measuring. 280px tall reads clearly on a projector at
# the size a sign card uses. Every frame shares one palette and nothing is
# disposed between frames, so the unchanging backdrop is encoded once and the
# file is roughly half what per-frame palettes cost. 64 colours with no dither
# keeps handshape and face clean; the backdrop is flat, so the colours are not
# doing much work. A sign lands around 300KB, so a deck of a dozen is about 4MB.
GIF_H = 280
GIF_MAX_FRAMES = 14
GIF_MS = 90            # per frame; about 11 frames a second
GIF_HOLD_MS = 550      # hold on the first and last frame so the sign reads
GIF_COLORS = 64


def fetch(url, timeout=30):
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.geturl(), r.read()


def fetch_text(url):
    try:
        final, raw = fetch(url)
        return final, raw.decode("utf-8", "replace")
    except Exception as exc:
        return None, str(exc)


def find_existing(gloss, ext=None):
    """Assets already in the bank for this gloss.

    With `ext` given, only that kind counts: a bank full of strips must not
    make a --gif run think there is nothing to do.
    """
    exts = [ext] if ext else ["gif", "jpg", "png"]
    out = []
    for e in exts:
        out += glob.glob(os.path.join(OUT_DIR, "{}.{}".format(gloss, e)))
    return sorted(out)


def gloss_to_term(gloss):
    if gloss in SEARCH_OVERRIDES:
        return SEARCH_OVERRIDES[gloss]
    return gloss.replace("-", " ").lower()


def word_url(term, index):
    return "{}/dictionary/words/{}-{}.html".format(
        BASE, urllib.parse.quote(term), index
    )


def search_best_match(term):
    """Search Signbank and return the headword whose entry best matches `term`.

    Signbank keeps the headword's own capitalisation and punctuation, so
    constructed URLs miss proper nouns (Australia) and bracketed senses
    (alarm (flashing)). Reading the result list avoids guessing.
    """
    final, text = fetch_text(
        "{}/dictionary/search/?query={}".format(BASE, urllib.parse.quote(term))
    )
    if final is None:
        return None, "network error: {}".format(text[:80])
    if "Sign Definition" in text:
        return final, None  # search redirected straight to the entry

    found = re.findall(r"/dictionary/words/([^\"']+?)-(\d+)\.html", text)
    if not found:
        return None, "no Signbank entry for '{}'".format(term)

    tl = term.lower()
    best, best_score = None, None
    for enc, _ in found:
        word = urllib.parse.unquote(enc)
        wl = word.lower()
        if wl == tl:
            score = (0, 0)
        elif wl.startswith(tl):
            score = (1, len(wl))
        elif tl in wl:
            score = (2, len(wl))
        else:
            score = (3, len(wl))
        if best_score is None or score < best_score:
            best, best_score = enc, score
    if best is None:
        return None, "no Signbank entry for '{}'".format(term)
    return "{}/dictionary/words/{}-1.html".format(BASE, best), None


def parse_entry(htmltext):
    """Pull the video URL, keywords, definition and variant count off an entry page."""
    video = None
    m = re.search(r'<source src="(https://[^"]+\.mp4)"', htmltext)
    if m:
        video = m.group(1)

    keywords = ""
    m = re.search(r"Keywords:\s*</strong>(.*?)</p>", htmltext, re.S)
    if m:
        keywords = html.unescape(
            re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", m.group(1)))
        ).strip()

    # Definitions sit in .definition-entry blocks, grouped under a part-of-speech
    # panel title. The sense number is its own span, so strip tags then collapse.
    defs = []
    for pm in re.finditer(
        r"panel-title'>([^<]+)</h3>(.*?)(?=panel-title'|Sign Distribution|</body>)",
        htmltext, re.S,
    ):
        kind = pm.group(1).strip()
        for dm in re.finditer(r'definition-entry"?>(.*?)</div>\s*</div>', pm.group(2), re.S):
            txt = html.unescape(
                re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", dm.group(1)))
            ).strip()
            if txt:
                defs.append("{}: {}".format(kind, txt))
    definition = " | ".join(defs[:4])

    # "Matches for the word team" is followed by a button group, one numbered
    # button or relative link per variant (href="team-2.html").
    variants = 1
    idx = htmltext.find("Matches for the word")
    if idx >= 0:
        seg = htmltext[idx: idx + 1200]
        nums = re.findall(r">\s*(\d+)\s*</(?:button|a)>", seg)
        if nums:
            variants = max(int(n) for n in nums)

    return {
        "video": video,
        "keywords": keywords,
        "definition": definition,
        "variants": min(variants, 12),
    }


VETTED_LINKS = {}


def load_links(path):
    """Read the teacher's vetted gloss -> Signbank entry URL map."""
    with open(path, encoding="utf-8") as fh:
        raw = fh.read().strip()
    if raw.startswith("{"):
        data = json.loads(raw)
    else:
        data = {}
        for line in raw.splitlines():
            line = line.split("#")[0].strip()
            if not line:
                continue
            parts = line.replace("\t", " ").split(None, 1)
            if len(parts) == 2:
                data[parts[0]] = parts[1].strip()
    clean = {}
    for gloss, url in data.items():
        url = str(url).strip()
        if not url.startswith(BASE + "/dictionary/words/"):
            print("  ignoring {}: not a Signbank entry URL ({})".format(gloss, url[:60]))
            continue
        clean[gloss.strip().upper()] = url
    return clean


def resolve_vetted(gloss):
    """Build the entry list from the teacher's own link, no search involved."""
    url = VETTED_LINKS[gloss]
    final, text = fetch_text(url)
    if final is None or "Sign Definition" not in text:
        return [], "vetted link did not load: {}".format(url)
    entry = parse_entry(text)
    if not entry["video"]:
        return [], "vetted entry has no video: {}".format(url)
    entry["url"] = final
    entry["vetted"] = True
    return [entry], None


def resolve(gloss):
    """Return a list of entry dicts for this gloss, best match first."""
    if gloss in VETTED_LINKS:
        return resolve_vetted(gloss)
    term = gloss_to_term(gloss)
    if not term:
        return [], "no lexical sign expected for this gloss"

    final, text = fetch_text(word_url(term.lower(), 1))
    if final is None or "Sign Definition" not in text:
        # Signbank keeps headword capitalisation and bracketed senses, so read
        # the search results rather than guessing a URL.
        hit, err = search_best_match(term)
        if err:
            return [], err
        final, text = fetch_text(hit)
        if final is None or "Sign Definition" not in text:
            return [], "no usable entry page for '{}'".format(term)
        m = re.search(r"/dictionary/words/([^\"'/]+)-\d+\.html", final)
        if m:
            term = urllib.parse.unquote(m.group(1))

    first = parse_entry(text)
    if not first["video"]:
        return [], "entry found but no video for '{}'".format(term)
    first["url"] = final
    entries = [first]

    for i in range(2, min(first["variants"], MAX_PANELS + 1) + 1):
        time.sleep(0.4)
        f2, t2 = fetch_text(word_url(term, i))
        if f2 is None or "Sign Definition" not in t2:
            break
        e = parse_entry(t2)
        if not e["video"]:
            break
        e["url"] = f2
        entries.append(e)

    return entries, None


def download_video(url):
    os.makedirs(CACHE_DIR, exist_ok=True)
    name = re.sub(r"[^A-Za-z0-9_.-]", "_", url.split("/")[-1])
    path = os.path.join(CACHE_DIR, name)
    if os.path.exists(path) and os.path.getsize(path) > 1000:
        return path
    _, raw = fetch(url, timeout=90)
    # Write then rename: two workers can want the same clip (synonyms share a
    # video) and a half-written file would fail to decode.
    tmp = "{}.{}.part".format(path, os.getpid())
    with open(tmp, "wb") as fh:
        fh.write(raw)
    os.replace(tmp, path)
    return path


def read_frames(path):
    cap = cv2.VideoCapture(path)
    frames = []
    while True:
        ok, fr = cap.read()
        if not ok:
            break
        frames.append(fr)
    cap.release()
    return frames


def _thumb(fr):
    return cv2.cvtColor(cv2.resize(fr, (160, 90)), cv2.COLOR_BGR2GRAY).astype(np.float32)


def _sharpness(fr):
    g = cv2.cvtColor(cv2.resize(fr, (320, 180)), cv2.COLOR_BGR2GRAY)
    return float(cv2.Laplacian(g, cv2.CV_64F).var())


def sign_window(frames):
    """Return (lo, hi, diffs, thresh) for the stretch that is actually the sign.

    Frame 0 and the last frame are rest poses. A frame resembling either one is
    the signer arriving or leaving, not signing. The longest contiguous run of
    genuinely-signing frames is the window. Returns lo == hi == -1 when the clip
    never moves enough to call it.
    """
    if len(frames) < 3:
        return -1, -1, [], 0.0
    rest_in, rest_out = _thumb(frames[0]), _thumb(frames[-1])
    diffs = []
    for fr in frames:
        t = _thumb(fr)
        diffs.append(min(float(np.abs(t - rest_in).mean()),
                         float(np.abs(t - rest_out).mean())))
    peak = max(diffs) if diffs else 0.0
    if peak < 1.5:
        return -1, -1, diffs, 0.0
    thresh = peak * 0.5
    best = cur = None
    for i, d in enumerate(diffs):
        if d >= thresh:
            cur = (cur[0], i) if cur else (i, i)
            if not best or (cur[1] - cur[0]) > (best[1] - best[0]):
                best = cur
        else:
            cur = None
    if not best:
        i = int(np.argmax(diffs))
        return i, i, diffs, thresh
    return best[0], best[1], diffs, thresh


def pick_frames(frames):
    """Choose up to MAX_PANELS frames that all sit inside the sign itself.

    Each panel is nudged to the sharpest nearby frame so fast signs do not come
    out motion-blurred.
    """
    if len(frames) < 3:
        return frames[:1]
    lo, hi, diffs, thresh = sign_window(frames)
    if lo < 0:
        return [frames[len(frames) // 2]]
    if hi - lo < 3:
        return [frames[(lo + hi) // 2]]

    picks = []
    for f in (0.08, 0.5, 0.92):
        target = int(round(lo + (hi - lo) * f))
        window = [i for i in range(max(lo, target - 2), min(hi, target + 2) + 1)
                  if diffs[i] >= thresh]
        if not window:
            continue
        picks.append(max(window, key=lambda i: _sharpness(frames[i])))

    # Drop near-duplicates so a held handshape does not print three times.
    kept = []
    for i in sorted(set(picks)):
        t = _thumb(frames[i])
        if all(np.abs(t - _thumb(frames[k])).mean() > 2.5 for k in kept):
            kept.append(i)
    return [frames[i] for i in kept] or [frames[(lo + hi) // 2]]


def signer_bbox(frames):
    """Bounding box of the signer across all chosen frames.

    Signbank was filmed over several decades against different backdrops -
    bright blue, dark navy, green. Rather than assume a colour, learn it from
    the frame corners and the strip above the signer's head, which are always
    backdrop, then treat anything far from that colour as the signer.
    """
    h, w = frames[0].shape[:2]
    # Match on hue and saturation, not RGB distance: the backdrops are lit
    # unevenly, so a brightness gradient must not read as "signer".
    hsv0 = cv2.cvtColor(frames[0], cv2.COLOR_BGR2HSV)
    band = max(2, h // 14)
    ring = np.concatenate([
        hsv0[0:band, :].reshape(-1, 3),
        hsv0[:, 0:band].reshape(-1, 3),
        hsv0[:, w - band:w].reshape(-1, 3),
    ])
    bg_h = float(np.median(ring[:, 0]))
    bg_s = float(np.median(ring[:, 1]))

    mask_total = np.zeros((h, w), dtype=bool)
    for fr in frames:
        hsv = cv2.cvtColor(fr, cv2.COLOR_BGR2HSV)
        dh = np.abs(hsv[:, :, 0].astype(np.float32) - bg_h)
        dh = np.minimum(dh, 180.0 - dh)  # hue is circular
        is_bg = (dh < 14) & (hsv[:, :, 1].astype(np.float32) > bg_s * 0.45)
        mask_total |= ~is_bg

    # Ignore stray specks (compression noise, station idents burnt into old clips).
    mask_u8 = cv2.morphologyEx(
        mask_total.astype(np.uint8), cv2.MORPH_OPEN, np.ones((5, 5), np.uint8)
    )
    col_counts = mask_u8.sum(axis=0)
    row_counts = mask_u8.sum(axis=1)
    cols = np.where(col_counts > h * 0.02)[0]
    rows = np.where(row_counts > w * 0.02)[0]
    if len(cols) == 0 or len(rows) == 0:
        return 0, 0, w, h
    x0, x1 = int(cols[0]), int(cols[-1])
    y0, y1 = int(rows[0]), int(rows[-1])
    padx = int((x1 - x0) * 0.08) + 8
    pady = int((y1 - y0) * 0.06) + 8
    x0 = max(0, x0 - padx)
    x1 = min(w - 1, x1 + padx)
    y0 = max(0, y0 - pady)
    y1 = min(h - 1, y1 + pady)
    return x0, y0, x1 - x0 + 1, y1 - y0 + 1


def compose_strip(frames, out_path):
    x, y, w, h = signer_bbox(frames)
    panels = []
    for fr in frames:
        crop = fr[y:y + h, x:x + w]
        scale = STRIP_H / crop.shape[0]
        panels.append(cv2.resize(crop, (max(1, int(crop.shape[1] * scale)), STRIP_H),
                                 interpolation=cv2.INTER_AREA))

    total_w = sum(p.shape[1] for p in panels) + PANEL_GAP * (len(panels) - 1) + MARGIN * 2
    total_h = STRIP_H + MARGIN * 2
    canvas = np.full((total_h, total_w, 3), 255, dtype=np.uint8)
    cx = MARGIN
    for p in panels:
        canvas[MARGIN:MARGIN + STRIP_H, cx:cx + p.shape[1]] = p
        cx += p.shape[1] + PANEL_GAP
    # JPEG, not PNG: these are photographic stills and PNG costs ~10x the bytes
    # for no visible gain. Line-art scans added by hand may stay PNG.
    cv2.imwrite(out_path, canvas, [cv2.IMWRITE_JPEG_QUALITY, 88])
    return len(panels), total_w, total_h


def compose_gif(frames, out_path):
    """Write the sign as a looping GIF, cropped to the signer.

    Padded a little either side of the movement window so the sign starts from
    rest and returns to it, which is how a learner needs to see it. The first
    and last frames are held so the start and end handshapes read before the
    loop comes round again.
    """
    lo, hi, _, _ = sign_window(frames)
    if lo < 0:
        lo, hi = 0, len(frames) - 1
    pad = max(2, (hi - lo) // 6)
    lo = max(0, lo - pad)
    hi = min(len(frames) - 1, hi + pad)
    window = frames[lo:hi + 1] or frames

    step = max(1, int(round(len(window) / float(GIF_MAX_FRAMES))))
    chosen = window[::step][:GIF_MAX_FRAMES]
    if len(chosen) < 2:
        chosen = window[:2] or window

    x, y, w, h = signer_bbox(chosen)
    pil = []
    for fr in chosen:
        crop = fr[y:y + h, x:x + w]
        scale = GIF_H / float(crop.shape[0])
        small = cv2.resize(
            crop, (max(1, int(crop.shape[1] * scale)), GIF_H), interpolation=cv2.INTER_AREA
        )
        pil.append(Image.fromarray(cv2.cvtColor(small, cv2.COLOR_BGR2RGB)))

    # One palette for every frame, so GIF can encode only what changed between
    # them. Per-frame palettes force a full frame each time and cost about
    # double for no visible gain on a flat backdrop.
    base = pil[0].quantize(colors=GIF_COLORS, method=Image.MEDIANCUT)
    quant = [base] + [p.quantize(palette=base, dither=Image.NONE) for p in pil[1:]]

    durations = [GIF_MS] * len(quant)
    durations[0] = GIF_HOLD_MS
    durations[-1] = GIF_HOLD_MS
    quant[0].save(
        out_path, save_all=True, append_images=quant[1:],
        duration=durations, loop=0, optimize=True, disposal=1,
    )
    return len(quant), quant[0].width, quant[0].height


def build_gloss(gloss, refetch=False, as_gif=False):
    entries, err = resolve(gloss)
    if err:
        return {"gloss": gloss, "status": "MISSING", "reason": err}

    ext = "gif" if as_gif else "jpg"
    made = []
    for i, entry in enumerate(entries[:MAX_PANELS]):
        suffix = "" if i == 0 else "_{}".format(i + 1)
        out = os.path.join(OUT_DIR, "{}{}.{}".format(gloss, suffix, ext))
        if os.path.exists(out) and not refetch:
            made.append({"file": os.path.basename(out), "skipped": True, "entry": entry["url"]})
            continue
        try:
            vid = download_video(entry["video"])
            frames = read_frames(vid)
            if not frames:
                continue
            if as_gif:
                n, w, h = compose_gif(frames, out)
            else:
                n, w, h = compose_strip(pick_frames(frames), out)
        except Exception as exc:
            return {"gloss": gloss, "status": "ERROR", "reason": str(exc)[:120]}
        made.append({
            "file": os.path.basename(out),
            "kind": "gif" if as_gif else "strip",
            "frames" if as_gif else "panels": n,
            "size": "{}x{}".format(w, h),
            "kb": round(os.path.getsize(out) / 1024.0),
            "entry": entry["url"],
            "vetted": bool(entry.get("vetted")),
            "keywords": entry["keywords"],
            "definition": entry["definition"],
            "video": entry["video"],
        })
        time.sleep(0.5)

    if not made:
        return {"gloss": gloss, "status": "MISSING", "reason": "no usable video"}
    return {"gloss": gloss, "status": "OK", "images": made,
            "variants_available": entries[0]["variants"]}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--glosses", nargs="*", default=[])
    ap.add_argument("--from-file")
    ap.add_argument("--refetch", action="store_true")
    ap.add_argument("--gif", action="store_true",
                    help="write an animated GIF of the whole sign instead of a "
                         "still strip. This is what decks use")
    ap.add_argument("--links",
                    help="the teacher's vetted gloss -> Signbank entry URL map "
                         "(JSON, or one 'GLOSS url' per line). Listed glosses "
                         "are fetched from that entry, not from a search")
    ap.add_argument("--workers", type=int, default=4,
                    help="parallel fetches (default 4; Signbank is run by a "
                         "charity, so do not raise this much)")
    args = ap.parse_args()

    glosses = list(args.glosses)
    if args.from_file:
        with open(args.from_file, encoding="utf-8") as fh:
            for line in fh:
                line = line.split("#")[0].strip()
                if line:
                    glosses.append(line.upper())
    if args.links:
        VETTED_LINKS.update(load_links(args.links))
        print("Vetted links loaded for {} gloss(es).".format(len(VETTED_LINKS)))
        if not glosses:
            glosses = sorted(VETTED_LINKS)
    if not glosses:
        ap.error("give --glosses, --from-file or --links")

    os.makedirs(OUT_DIR, exist_ok=True)
    ext = "gif" if args.gif else "jpg"

    # Skip glosses already in the bank before spending any requests on them.
    # After a few units most of a new unit's vocabulary is already here, which
    # is the whole point of sharing one bank.
    todo, already = [], []
    for g in dict.fromkeys(glosses):  # de-duplicate, keep order
        if not args.refetch and find_existing(g, ext):
            already.append(g)
        else:
            todo.append(g)
    if already:
        print("Already in the bank, skipping {}: {}".format(
            len(already), ", ".join(already)))
    if not todo:
        print("Nothing to fetch.")
        return

    results = []
    done = 0
    lock = threading.Lock()
    print("Fetching {} gloss(es) with {} workers...".format(len(todo), args.workers))
    with cf.ThreadPoolExecutor(max_workers=max(1, args.workers)) as pool:
        futures = {pool.submit(build_gloss, g, args.refetch, args.gif): g for g in todo}
        for fut in cf.as_completed(futures):
            g = futures[fut]
            try:
                res = fut.result()
            except Exception as exc:
                res = {"gloss": g, "status": "ERROR", "reason": str(exc)[:120]}
            with lock:
                done += 1
                results.append(res)
                if res["status"] == "OK":
                    detail = ", ".join(m["file"] for m in res["images"])
                else:
                    detail = res.get("reason", "")
                print("[{:>3}/{}] {:<7} {:<16} {}".format(
                    done, len(todo), res["status"], g, detail))
                sys.stdout.flush()
    results.sort(key=lambda r: r["gloss"])

    manifest_path = os.path.join(OUT_DIR, "manifest.json")
    existing = {}
    if os.path.exists(manifest_path):
        try:
            with open(manifest_path, encoding="utf-8") as fh:
                existing = {r["gloss"]: r for r in json.load(fh).get("signs", [])}
        except Exception:
            existing = {}
    for r in results:
        existing[r["gloss"]] = r

    with open(manifest_path, "w", encoding="utf-8") as fh:
        json.dump({
            "source": "Auslan Signbank, auslan.org.au",
            "licence": "CC BY-NC-ND 4.0; stills extracted for internal school "
                       "teaching use under the Australian schools statutory "
                       "educational licence",
            "warning": "Nothing here is a verified sign. Check each entry link "
                       "and rehearse before teaching.",
            "signs": sorted(existing.values(), key=lambda r: r["gloss"]),
        }, fh, indent=2)

    ok = sum(1 for r in results if r["status"] == "OK")
    kb = sum(m.get("kb", 0) for r in results if r["status"] == "OK" for m in r["images"])
    print("\n{} of {} glosses imaged as {}s, {} KB total. Manifest: {}".format(
        ok, len(results), ext, kb, manifest_path))
    unvetted = [r["gloss"] for r in results if r["status"] == "OK"
                and not any(m.get("vetted") for m in r["images"])]
    if unvetted and VETTED_LINKS:
        print("Found by search, not from a vetted link: " + ", ".join(unvetted))
    missing = [r["gloss"] for r in results if r["status"] != "OK"]
    if missing:
        print("Not found (deck will use lookup cards): " + ", ".join(missing))


if __name__ == "__main__":
    main()
