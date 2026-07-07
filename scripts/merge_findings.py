# -*- coding: utf-8 -*-
"""Merge subagent findings JSONs into verified, counted, sorted data.

Usage:
    PYTHONUTF8=1 python merge_findings.py <total_slides> <f1.json> [f2.json ...] [-o merged.json]

Each input file is one subagent's return value:
{
  "range": [1, 30],                 # the slide range the agent was assigned
  "reviewed_slides": [1, 2, ...],   # every slide it actually read (MUST cover range)
  "findings": [
    {"slide": 17,                   # primary slide (int)
     "slides_extra": [34, 35],      # optional: other slides of the same finding
     "check": 29,                   # 1..36
     "severity": "🔴",              # 🔴 / 🟠 / 🔵
     "quote": "...",                # verbatim quote from the slide (required)
     "recommendation": "...",       # full, paste-ready recommendation (required)
     "fixes": "..."                 # what the fix solves (required)
    }, ...
  ],
  "passed_checks": [{"check": 5, "note": "..."}]   # optional
}

Output (merged.json + stdout summary):
- counts per severity (the report header numbers — never hand-tallied again)
- missing_slides: gaps in 1..N coverage  → dispatch a complement agent before reporting
- invalid: findings dropped for schema violations (no quote / bad check number)
- duplicates: same (slide, check) reported by more than one agent (kept once)
- findings: sorted by slide, then severity
"""
import io
import json
import sys

SEVS = ["🔴", "🟠", "🔵"]
SEV_ORDER = {s: i for i, s in enumerate(SEVS)}
N_CHECKS = 36


def load(path):
    with io.open(path, encoding="utf-8-sig") as f:
        return json.load(f)


def main(argv):
    args = [a for a in argv if a != "-o"]
    out_path = None
    if "-o" in argv:
        out_path = argv[argv.index("-o") + 1]
        args = argv[:argv.index("-o")]
    total = int(args[0])
    files = args[1:]
    if not files:
        print("usage: merge_findings.py <total_slides> <f1.json> ... [-o merged.json]")
        return 2

    reviewed, findings, passed, invalid, dups = set(), [], [], [], []
    seen = {}
    for path in files:
        data = load(path)
        reviewed.update(int(s) for s in data.get("reviewed_slides", []))
        for p in data.get("passed_checks", []):
            passed.append(p)
        for f in data.get("findings", []):
            problems = []
            slide = f.get("slide")
            check = f.get("check")
            if not isinstance(slide, int) or not (1 <= slide <= total):
                problems.append("slide")
            if not isinstance(check, int) or not (1 <= check <= N_CHECKS):
                problems.append("check")
            if f.get("severity") not in SEVS:
                problems.append("severity")
            for key in ("quote", "recommendation", "fixes"):
                if not str(f.get(key, "")).strip():
                    problems.append(key)   # bli tsitut ein mimtsa
            if problems:
                invalid.append({"file": path, "finding": f, "problems": problems})
                continue
            k = (slide, check, str(f.get("quote"))[:40])
            if k in seen:
                dups.append({"slide": slide, "check": check, "files": [seen[k], path]})
                continue
            seen[k] = path
            f["_src"] = path
            findings.append(f)

    missing = sorted(set(range(1, total + 1)) - reviewed)
    findings.sort(key=lambda f: (f["slide"], SEV_ORDER[f["severity"]], f["check"]))
    counts = {s: sum(1 for f in findings if f["severity"] == s) for s in SEVS}

    merged = {
        "total_slides": total,
        "coverage_ok": not missing,
        "missing_slides": missing,
        "counts": {"blocker": counts["🔴"], "warning": counts["🟠"], "polish": counts["🔵"]},
        "duplicates": dups,
        "invalid": invalid,
        "findings": findings,
        "passed_checks": passed,
    }
    if out_path:
        with io.open(out_path, "w", encoding="utf-8") as f:
            json.dump(merged, f, ensure_ascii=False, indent=1)

    print("# merge summary")
    print("slides: %d/%d reviewed%s" % (total - len(missing), total,
          "" if not missing else "  !! MISSING: %s" % missing))
    print("counts: 🔴 %d · 🟠 %d · 🔵 %d" % (counts["🔴"], counts["🟠"], counts["🔵"]))
    if dups:
        print("duplicates dropped: %d  %s" % (len(dups), [(d["slide"], d["check"]) for d in dups]))
    if invalid:
        print("INVALID (dropped, fix the agent output!): %d" % len(invalid))
        for iv in invalid:
            print("  - %s: slide=%s check=%s problems=%s" % (
                iv["file"], iv["finding"].get("slide"), iv["finding"].get("check"), iv["problems"]))
    return 0 if (not missing and not invalid) else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
