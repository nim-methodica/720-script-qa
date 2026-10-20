# -*- coding: utf-8 -*-
"""בדיקת-יחידה ל-template_scan.py: בונה את ה-fixture, מריץ את הסורק, ובודק שכל פגם מושתל
נתפס (HIT) ושאף מלכודת לא מדווחת (TRAP). הרצה (מתיקיית הסקיל):
  PYTHONUTF8=1 python tests/template/test_template_scan.py
קוד-יציאה 0 = PASS."""
import os
import re
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
SK = os.path.dirname(os.path.dirname(HERE))
FIX = os.path.join(HERE, "fixture.pptx")
subprocess.check_call([sys.executable, os.path.join(HERE, "build_fixture.py"), FIX], stdout=subprocess.DEVNULL)
out = subprocess.check_output([sys.executable, os.path.join(SK, "scripts", "template_scan.py"), FIX],
                              env={**os.environ, "PYTHONUTF8": "1"}).decode("utf-8")

sections = {}
cur = None
for ln in out.splitlines():
    m = re.match(r"### שקף (\d+)", ln)
    if m:
        cur = int(m.group(1))
        sections[cur] = []
    elif cur and ln.startswith("- #"):
        sections[cur].append(ln)
    elif ln.startswith("## "):
        cur = None


def has(slide, check, frag=""):
    return any(l.startswith("- #%s " % check) and frag in l for l in sections.get(slide, []))


PLANTED = [
    (1, 37, "המסה תלויה במיקום", "#37 נקודה בסוף מסיח"),
    (1, 2, "␣␣", "#2 רווח כפול במסיח"),
    (1, 44, "אילו מהמשפטים", "#44 אין רווח לפני השאלה"),
    (1, 40, "טקסט נסתר בלבן", "#40 טקסט לבן"),
    (1, 45, "התשובה הנכונה", "#45 תשובה בהערות"),
    (2, 38, "אפשר רמז", "#38 רמז בשאלת 2 מסיחים"),
    (2, 42, "0 תיבות-משוב", "#42 אין תיבות-משוב"),
    (2, 40, "אנימציה", "#40 אנימציה על המסך"),
    (2, 40, "למפתח", "#40 הנחיית 'למפתחת'"),
    (2, 41, "טבלה", "#41 תמונה-טבלה"),
    (4, 42, "זהה לשקף הקודם", "#42 שקף-מצב נפרד"),
    (4, 43, "מילה בודדת", "#43 מילה בודדת בשורה"),
]
TRAPS = [
    (3, 37, "", "מסיחים ללא נקודה — תקין"),
    (3, 38, "", "שאלת 2 מסיחים ללא רמז — תקין"),
    (3, 40, "", "לבן על רקע כהה / 'מקום לגיף' / 'איריס, …' — תקינים"),
    (3, 41, "", "טבלה חיה — תקין"),
    (3, 42, "", "שני משובים — תקין"),
    (3, 43, "", "'שאלה 1 מתוך 5' — תווית-תבנית"),
    (3, 44, "", "רווח בין מקדים לשאלה — תקין"),
    (3, 45, "", "הערות ללא תשובה — תקין"),
    (1, 42, "", "שאלת 4 מסיחים עם 2 משובים — תקין"),
]
fails = 0
for s, c, frag, d in PLANTED:
    ok = has(s, c, frag)
    fails += not ok
    print(("HIT    " if ok else "MISSED ") + "שקף %d #%d %s" % (s, c, d))
for s, c, frag, d in TRAPS:
    bad = has(s, c, frag)
    fails += bad
    print(("TRAP-VIOLATED " if bad else "trap-ok       ") + "שקף %d #%d %s" % (s, c, d))
print("verdict: %s" % ("PASS" if not fails else "FAIL (%d)" % fails))
sys.exit(1 if fails else 0)
