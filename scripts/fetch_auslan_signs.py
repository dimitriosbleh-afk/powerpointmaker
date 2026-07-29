"""Populate the Auslan sign image bank from Auslan Signbank.

For each gloss, finds the Signbank entry, downloads the sign video, and composes
a still sequence strip that shows the sign's movement (the way the printed
references do). Writes assets/auslan_signs/<GLOSS>.png plus a manifest recording
exactly which entry each image came from, so the teacher can verify every sign
before teaching it.

Nothing here is a verified sign. The manifest is the verification worklist.

Usage:
    python scripts/fetch_auslan_signs.py --glosses TEAM SCHOOL AGAIN
    python scripts/fetch_auslan_signs.py --from-file glosses.txt
    python scripts/fetch_auslan_signs.py --from-file glosses.txt --refetch

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


def find_existing(gloss):
    """Images already in the bank for this gloss, either extension."""
    return sorted(
        glob.glob(os.path.join(OUT_DIR, "{}.jpg".format(gloss)))
        + glob.glob(os.path.join(OUT_DIR, "{}.png".format(gloss)))
    )


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


def resolve(gloss):
    """Return a list of entry dicts for this gloss, best match first."""
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


def pick_frames(frames):
    """Choose up to MAX_PANELS frames that all sit inside the sign itself.

    Frame 0 is the rest pose. Frames far from it are the sign. Panels are taken
    only from the sustained-movement window, so the return to rest never prints
    as if it were part of the sign, and each panel is nudged to the sharpest
    nearby frame so fast signs do not come out motion-blurred.
    """
    if len(frames) < 3:
        return frames[:1]

    # Score each frame against BOTH rest poses. A frame that resembles either the
    # opening or the closing rest is the signer arriving or leaving, not signing.
    rest_in, rest_out = _thumb(frames[0]), _thumb(frames[-1])
    diffs = []
    for fr in frames:
        t = _thumb(fr)
        diffs.append(min(float(np.abs(t - rest_in).mean()),
                         float(np.abs(t - rest_out).mean())))
    peak = max(diffs) if diffs else 0.0
    if peak < 1.5:
        return [frames[len(frames) // 2]]

    # Longest contiguous run of genuinely-signing frames.
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
        return [frames[int(np.argmax(diffs))]]

    lo, hi = best
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


def build_gloss(gloss, refetch=False):
    entries, err = resolve(gloss)
    if err:
        return {"gloss": gloss, "status": "MISSING", "reason": err}

    made = []
    for i, entry in enumerate(entries[:MAX_PANELS]):
        suffix = "" if i == 0 else "_{}".format(i + 1)
        out = os.path.join(OUT_DIR, "{}{}.jpg".format(gloss, suffix))
        if os.path.exists(out) and not refetch:
            made.append({"file": os.path.basename(out), "skipped": True, "entry": entry["url"]})
            continue
        try:
            vid = download_video(entry["video"])
            frames = read_frames(vid)
            if not frames:
                continue
            picks = pick_frames(frames)
            n, w, h = compose_strip(picks, out)
        except Exception as exc:
            return {"gloss": gloss, "status": "ERROR", "reason": str(exc)[:120]}
        made.append({
            "file": os.path.basename(out),
            "panels": n,
            "size": "{}x{}".format(w, h),
            "entry": entry["url"],
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
    if not glosses:
        ap.error("give --glosses or --from-file")

    os.makedirs(OUT_DIR, exist_ok=True)

    # Skip glosses already in the bank before spending any requests on them.
    # After a few units most of a new unit's vocabulary is already here, which
    # is the whole point of sharing one bank.
    todo, already = [], []
    for g in dict.fromkeys(glosses):  # de-duplicate, keep order
        if not args.refetch and find_existing(g):
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
        futures = {pool.submit(build_gloss, g, args.refetch): g for g in todo}
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
    print("\n{} of {} glosses imaged. Manifest: {}".format(ok, len(results), manifest_path))
    missing = [r["gloss"] for r in results if r["status"] != "OK"]
    if missing:
        print("Not found (deck will use lookup cards): " + ", ".join(missing))


if __name__ == "__main__":
    main()
