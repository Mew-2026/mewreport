"""Read a handwritten loading sheet with renamed.to by cutting it into strips.

renamed.to only returns a filename, so each strip is small enough that its
rows fit in one filename. Rows come back as "N.MODEL-QTY"; overlapping strips
are merged by row number.

Usage: python loading_sheet_ocr.py <image> [out.json]
"""
import json
import os
import re
import sys
import tempfile

import cv2
import requests
from rapidfuzz import fuzz

API = "https://www.renamed.to/api/v1/rename"
PROMPT = (
    "Write every fully visible numbered handwritten row of this image into the filename, "
    "in order, as ROWNUMBER.MODEL-QTY joined by underscore, e.g. 1.2042STJ-3_2.2042TT-2. "
    "QTY is the number after the arrow. MODEL has no spaces or hyphens. "
    "Copy each MODEL exactly as written, all letters and digits, never shorten it "
    "(e.g. TC13548ST, RC18000STP, LPTT12150H). The handwritten TT looks like the symbol pi: write it as TT. "
    "Skip rows that are cut off at the top or bottom edge. Skip crossed-out rows. "
    "No date, no other words."
)
STRIP_H = 180     # px, ~5 rows at 1280px page height
OVERLAP = 55      # px, so every row is whole in at least one strip
ROW_START = re.compile(r"(?:^|[_\s])(\d{1,3})\.")   # "_12." marks the start of row 12
QTY_END = re.compile(r"[-\s_]+(\d{1,4})\W*$")       # trailing "-7" is the quantity


def count_rows(name):
    return len(ROW_START.findall(os.path.splitext(name)[0]))


def parse_rows(name):
    """'13.ASTT_12150H-2_14.BC16048ST-48.jpg' -> [(13, 'ASTT12150H', 2), (14, 'BC16048ST', 48)]"""
    stem = os.path.splitext(name)[0]
    starts = list(ROW_START.finditer(stem))
    rows = []
    for i, m in enumerate(starts):
        chunk = stem[m.end(): starts[i + 1].start() if i + 1 < len(starts) else len(stem)]
        q = QTY_END.search(chunk)
        if not q:
            continue  # model without a quantity: unusable
        model = re.sub(r"[^A-Za-z0-9]", "", chunk[: q.start()]).upper()
        row = (int(m.group(1)), model, int(q.group(1)))
        # renamed.to sometimes writes a row twice in a row; drop the repeat
        if model and not (rows and rows[-1][1:] == row[1:]):
            rows.append(row)
    return rows


def load_key():
    key = os.environ.get("RENAMED_API_KEY")
    env = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
    if not key and os.path.exists(env):
        for line in open(env, encoding="utf-8"):
            if line.startswith("RENAMED_API_KEY="):
                key = line.split("=", 1)[1].strip()
    if not key:
        sys.exit("RENAMED_API_KEY not set (put it in .env)")
    return key


def strips(img):
    """Yield crops of the page: left column and right column, cut into overlapping strips."""
    h, w = img.shape[:2]
    # Column layout measured on a 879x1280 sheet; scale to the actual image.
    sx, sy = w / 879, h / 1280
    columns = [(0, 350, 75, 1240), (330, 879, 85, 380)]
    for col, (x0, x1, y0, y1) in enumerate(columns):
        x0, x1, y0, y1 = int(x0 * sx), int(x1 * sx), int(y0 * sy), int(y1 * sy)
        step, sh = int((STRIP_H - OVERLAP) * sy), int(STRIP_H * sy)
        y = y0
        while y < y1:
            # second crop is the same strip shifted down a third, used when the first comes back empty
            shifted = img[min(y + sh // 3, y1 - 1):min(y + sh + sh // 3, y1), x0:x1]
            yield col, img[y:min(y + sh, y1), x0:x1], shifted
            if y + sh >= y1:
                break
            y += step


def rename(key, path):
    with open(path, "rb") as f:
        r = requests.post(
            API,
            headers={"Authorization": f"Bearer {key}"},
            files={"file": (os.path.basename(path), f, "image/jpeg")},
            data={"strategy": "follow_custom_prompt", "templateMode": "custom", "customTemplate": PROMPT},
            timeout=180,
        )
    r.raise_for_status()
    return r.json().get("suggestedFilename", "")


def same_model(a, b):
    # a cut-off reply can shorten a model ("NEO" for "NEO18030"), so a prefix also counts
    return fuzz.ratio(a, b) >= 80 or (min(len(a), len(b)) >= 3 and (a.startswith(b) or b.startswith(a)))


def same_row(a, b):
    return a[2] == b[2] and same_model(a[1], b[1])


def overlaps(tail, head):
    """tail (end of earlier strip) shows the same rows as head (start of next strip).

    The earlier strip's last row may be cut off, so its quantity only has to be a prefix
    of the newer reading (RC25000P-1 vs RC25000P-14)."""
    if len(tail) == 1:
        return same_row(tail[0], head[0])
    *rest, (_, tm, tq) = tail
    _, hm, hq = head[-1]
    return all(same_row(x, y) for x, y in zip(rest, head)) and same_model(tm, hm) and str(hq).startswith(str(tq))


def combine(old, new):
    return [(n, max(om, nm, key=len), nq) for (_, om, _), (n, nm, nq) in zip(old, new)]


def merge(seq, new):
    """Append strip rows `new` to `seq`, dropping the rows that overlap seq's tail.

    The overlap is the longest tail of seq that matches the head of new; inside it the newer
    strip's quantity wins and the longer model text is kept. If seq's tail has one extra row the newer strip left out (usually a
    struck-out row read by only one strip), that row is dropped.
    """
    for k in range(min(len(seq), len(new)), 0, -1):
        if overlaps(seq[-k:], new[:k]):
            return seq[:-k] + combine(seq[-k:], new[:k]) + new[k:]
        tail = seq[-k - 1:]
        if len(tail) == k + 1:
            for skip in range(1, k + 1):  # never skip the tail's first row: that's a plain k+1 overlap
                if overlaps(tail[:skip] + tail[skip + 1:], new[:k]):
                    return seq[:-k - 1] + combine(tail[:skip] + tail[skip + 1:], new[:k]) + new[k:]
    return seq + new


def read_sheet(image_path):
    key = load_key()
    img = cv2.imread(image_path)
    if img is None:
        sys.exit(f"cannot read image {image_path}")
    strip_rows = []
    with tempfile.TemporaryDirectory() as tmp:
        p = os.path.join(tmp, "strip.jpg")
        for i, (col, crop, shifted) in enumerate(strips(img)):
            # Replies vary call to call: sometimes empty, sometimes rows without a quantity.
            # Retry until every row in the reply parses; keep the reply with the most rows.
            best_name, best = "", []
            for part in (crop, crop, crop, shifted):
                cv2.imwrite(p, cv2.resize(part, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC))
                name = rename(key, p)
                print(f"strip {i}: {name}", file=sys.stderr)
                found = parse_rows(name)
                if len(found) > len(best):
                    best_name, best = name, found
                if found and len(found) == count_rows(name):
                    break
            strip_rows.append({"column": col, "filename": best_name, "rows": best})
    return build(strip_rows)


def build(strip_rows):
    rows = []
    for col in sorted({s["column"] for s in strip_rows}):
        seq = []
        for s in strip_rows:
            if s["column"] == col:
                seq = merge(seq, [tuple(r) for r in s["rows"]])
        rows += seq
    return {
        "rows": [{"line": i + 1, "sheet_row": n, "model": m, "qty": q} for i, (n, m, q) in enumerate(rows)],
        "total": sum(q for _, _, q in rows),
        "strips": strip_rows,
    }


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    result = read_sheet(sys.argv[1])
    out = json.dumps(result, indent=1)
    if len(sys.argv) > 2:
        open(sys.argv[2], "w", encoding="utf-8").write(out)
    print(out)
