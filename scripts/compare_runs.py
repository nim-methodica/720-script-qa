# -*- coding: utf-8 -*-
"""Compare two QA runs of the same script (בקרה חוזרת) → תוקן / נשאר / חדש.

Usage:
    PYTHONUTF8=1 python compare_runs.py <old_merged.json> <new_merged.json>

Inputs are merge_findings.py outputs (each has a "findings" list). Matching
is primarily by (check, normalized-quote-prefix) — quotes survive slide
insertions/deletions between versions, slide numbers do not. A same-check
finding whose quote changed slightly still matches via a loose token
overlap, so rewordings of the SAME defect don't show up as fixed+new.

Prints a ready-to-paste Markdown "דוח סבב" block + summary counts.
"""
import io
import json
import re
import sys


def load(path):
    with io.open(path, encoding="utf-8-sig") as f:
        return json.load(f)


def norm(q):
    return re.sub(r"\s+", " ", str(q)).strip()[:60]


def tokens(q):
    return set(re.findall(r"[\w֐-׿%]+", str(q)))


def match(f, candidates):
    """Best candidate with same check: exact quote-prefix, else >=60% token overlap."""
    nq, tq = norm(f.get("quote")), tokens(f.get("quote"))
    best = None
    for c in candidates:
        if c.get("check") != f.get("check"):
            continue
        if norm(c.get("quote")) == nq:
            return c
        tc = tokens(c.get("quote"))
        if tq and tc:
            ov = len(tq & tc) / max(len(tq), len(tc))
            if ov >= 0.6 and (best is None or ov > best[0]):
                best = (ov, c)
    return best[1] if best else None


def line(f, slide_key="slide"):
    return "בדיקה #%s · שקף %s · 「%s」" % (f.get("check"), f.get(slide_key), norm(f.get("quote")))


def main(argv):
    old, new = load(argv[0]), load(argv[1])
    old_f = list(old.get("findings", []))
    new_f = list(new.get("findings", []))

    remaining, fixed, added = [], [], []
    matched_new = set()
    for f in old_f:
        m = match(f, [n for i, n in enumerate(new_f) if i not in matched_new])
        if m is not None:
            matched_new.add(new_f.index(m))
            remaining.append((f, m))
        else:
            fixed.append(f)
    for i, n in enumerate(new_f):
        if i not in matched_new:
            added.append(n)

    print("## דוח סבב — השוואה לבדיקה הקודמת")
    print()
    print("**תיקנת %d מתוך %d · נשארו %d · נוספו %d חדשים**" % (
        len(fixed), len(old_f), len(remaining), len(added)))
    print()
    if fixed:
        print("### ✅ תוקנו (%d)" % len(fixed))
        for f in fixed:
            print("- " + line(f))
        print()
    if remaining:
        print("### ⏳ נשארו (%d)" % len(remaining))
        for f, m in remaining:
            moved = "" if f.get("slide") == m.get("slide") else " (עבר משקף %s לשקף %s)" % (
                f.get("slide"), m.get("slide"))
            print("- %s %s%s" % (m.get("severity", ""), line(m), moved))
        print()
    if added:
        print("### 🆕 חדשים (%d)" % len(added))
        for n in added:
            print("- %s %s" % (n.get("severity", ""), line(n)))
    return 0


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print(__doc__)
        sys.exit(2)
    sys.exit(main(sys.argv[1:]))
