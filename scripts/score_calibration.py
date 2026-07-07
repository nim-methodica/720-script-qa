# -*- coding: utf-8 -*-
"""Score a calibration run of the 720-script-qa skill.

Usage:
    PYTHONUTF8=1 python score_calibration.py tests/calibration/expected.json <agent_output.json>

<agent_output.json> is a subagent findings file in the merge_findings.py schema
(or a merged.json — both have a "findings" list).

A planted item counts as DETECTED when some finding matches its slide and its
check number (or one of the allowed "alt" check numbers — adjacent checks
overlap by design, e.g. #5/#14/#36 on a wrong-feedback plant).
A trap is VIOLATED when some finding matches its slide+check (and, if
"trap_quote_must_not_contain" is set, only when the quote contains that text —
this separates a legitimate plant on the same slide+check from the trap).

PASS: >= 13/15 planted detected AND 0 trap violations.
"""
import io
import json
import sys

PASS_MIN_DETECTED = 13


def load(path):
    with io.open(path, encoding="utf-8-sig") as f:
        return json.load(f)


def main(argv):
    if len(argv) != 2:
        print(__doc__)
        return 2
    expected, output = load(argv[0]), load(argv[1])
    findings = output.get("findings", [])

    def slides_of(f):
        return [f.get("slide")] + list(f.get("slides_extra", []))

    detected, missed = [], []
    for p in expected["planted"]:
        ok_checks = [p["check"]] + p.get("alt", [])
        hit = any(p["slide"] in slides_of(f) and f.get("check") in ok_checks
                  for f in findings)
        (detected if hit else missed).append(p)

    violations = []
    for t in expected["traps"]:
        for f in findings:
            if t["slide"] in slides_of(f) and f.get("check") == t["check"]:
                must = t.get("trap_quote_must_not_contain")
                if must and must not in str(f.get("quote", "")):
                    continue
                violations.append({"trap": t["desc"], "finding_quote": f.get("quote")})

    print("# calibration score")
    print("detected: %d/%d planted" % (len(detected), len(expected["planted"])))
    for m in missed:
        print("  MISSED  שקף %-3s #%-3s %s" % (m["slide"], m["check"], m["desc"]))
    print("trap violations: %d" % len(violations))
    for v in violations:
        print("  VIOLATED  %s  (quote: %s)" % (v["trap"], str(v["finding_quote"])[:60]))
    ok = len(detected) >= PASS_MIN_DETECTED and not violations
    print("verdict: %s" % ("PASS" if ok else "FAIL"))
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
