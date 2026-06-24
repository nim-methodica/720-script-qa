#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
scan.py — סורק דטרמיניסטי ממצה (read-only) על תסריט מחולץ (markdown מ-markitdown).
מבטיח כיסוי מלא של הבדיקות המכניות — כל מופע, בכל שקף. לא מתקן, רק מדווח.

שימוש:  python scan.py <extracted.md>
פלט:    דוח דטרמיניסטי לפי שקף (stdout) — להטמיע/להצליב מול הדוח הסופי.

הרצה מומלצת ל-UTF-8:  PYTHONUTF8=1 python scan.py <extracted.md>
"""
import re, sys, io

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

if len(sys.argv) < 2:
    print("usage: python scan.py <extracted.md>"); sys.exit(1)

text = open(sys.argv[1], encoding="utf-8").read()
parts = re.split(r"<!-- Slide number: (\d+) -->", text)
slides = [(parts[i], parts[i + 1]) for i in range(1, len(parts), 2)]
N = len(slides)

REPEAT_OK = {"שלב שלב", "סוף סוף", "מאוד מאוד", "לאט לאט", "יום יום", "טוב טוב", "חבל חבל"}

# --- רובד לשוני (#1/#21): סף משפט, watchlist אקדמי/מליצי, תבניות יתירות ---
SENTENCE_MAX = 25   # מילים למשפט — מעל זה דגל (גיל 13–14)
ACADEMIC = [        # מילים/ביטויים ברובד גבוה מדי לגיל ז'–ח' (העדפה לחלופה יומיומית)
    "הינו", "הינה", "הינם", "כתוצאה מכך", "יתרה מזאת", "יתרה מכך", "נוכח העובדה",
    "אי לכך", "חרף", "על אודות", "הללו", "בטרם", "כאמור", "שמא", "לעת עתה",
    "במידה ו", "על מנת", "אשר", "בעבור", "כיוון לכך",
]
ACADEMIC_RE = re.compile(r"(?<![א-ת])(" + "|".join(re.escape(w) for w in ACADEMIC) + r")(?![א-ת])")
REDUNDANCY = [      # תבניות יתירות נפוצות (זיהוי + הצעה)
    (re.compile(r"(?<![א-ת])(\S{2,})\s*,\s*\1(?![א-ת])"), "מילה כפולה עם פסיק"),
    (re.compile(r"(?<![א-ת])בעצם(?![א-ת])"), "מילת-מילוי 'בעצם'"),
    (re.compile(r"(?<![א-ת])ממש ממש(?![א-ת])"), "הגזמה כפולה 'ממש ממש'"),
    (re.compile(r"חזרה\s+בחזרה|יחד\s+ביחד"), "כפל מיותר"),
]
findings = {}   # slide -> list[(check, detail)]

def add(sl, check, detail):
    findings.setdefault(sl, []).append((check, detail))

IMG_RE = re.compile(r"!\[[^\]]*\]\([^)]*\)")

def body_lines(body):
    """טקסט גלוי בלבד. פושט תגיות-תמונה (markitdown לעיתים מצמיד טקסט-תלמיד אחרי תמונה),
    ומדלג על טבלאות/כותרות-md/תגיות. מחזיר את הטקסט הנקי."""
    for ln in body.splitlines():
        s = IMG_RE.sub(" ", ln).strip()
        if not s or s.startswith(("|", "#", "<", "---")):
            continue
        yield s, ln

def ctx(s, i, w=22):
    seg = s[max(0, i - w): i + w]
    return seg.replace("  ", "␣␣")

for sl, body in slides:
    notes = ""
    m = re.search(r"### Notes:(.*)$", body, re.S)
    if m:
        notes = m.group(1)
    visible = body.split("### Notes:")[0]   # גוף השקף ללא הערות-דובר

    for s, raw in body_lines(visible):
        if "methodica" in s:   # שורת מזהה — לא טקסט גלוי לתלמיד
            continue
        # #2 רווחים כפולים: 2-3 = שגיאת ניסוח; 4+ = כנראה תיבת-מילוי (לבדוק)
        for mm in re.finditer(r"(?<=\S)( {2,})(?=\S)", s):
            if len(mm.group(1)) <= 3:
                add(sl, "#2 רווח כפול", f'"…{ctx(s, mm.start())}…"')
            else:
                add(sl, "רצף רווחים ארוך (תיבת-מילוי? לבדוק)", f'"…{ctx(s, mm.start(), 16)}…"')
        # #4 אחוז עם רווח
        for mm in re.finditer(r"\d\s+%", s):
            add(sl, "#4 אחוז עם רווח", f'"…{ctx(s, mm.start())}…"')
        # #4 כפל לטיני
        for mm in re.finditer(r"\d\s*[xX*]\s*\d", s):
            add(sl, "#4 כפל לטיני (→ ·)", f'"…{ctx(s, mm.start())}…"')
        # #4 מקף-מקלדת בין מספרים עצמאיים (טווח → en-dash) — לא מזהים/תאריכים
        for mm in re.finditer(r"(?<![\w./-])\d{1,4}\s*-\s*\d{1,4}(?![\w./:-])", s):
            add(sl, "#4 מקף בין-מספרי (→ –/−)", f'"…{ctx(s, mm.start())}…"')
        # #1 מילה כפולה (לא אידיום)
        for mm in re.finditer(r"\b([א-ת]{2,})\s+\1\b", s):
            if mm.group(0) not in REPEAT_OK:
                add(sl, "#1 מילה כפולה", f'"{mm.group(0)}"')
        # #3 פנייה ממוגדרת יחיד
        for mm in re.finditer(r"\bאתה\b|\bתבחר\b|\bבחר\b(?!ו)|\bהקלד\b(?!ו)|\bלחץ\b(?!ו)|\bענה\b(?!ו)|\bסמן\b(?!ו)", s):
            add(sl, "#3 פנייה ממוגדרת", f'"…{ctx(s, mm.start())}…"')
        # #1 רובד גבוה/מליצי (watchlist — לבדוק register מול language-and-register.md)
        for mm in ACADEMIC_RE.finditer(s):
            add(sl, "#1 רובד גבוה (לבדוק register)", f'"{mm.group(1)}" — …{ctx(s, mm.start())}…')
        # #1 תבניות יתירות
        for pat, lab in REDUNDANCY:
            for mm in pat.finditer(s):
                add(sl, f"#1 יתירות — {lab}", f'"…{ctx(s, mm.start())}…"')

    # #1 משפט ארוך (>SENTENCE_MAX מילים) — פר-משפט, בנוסף לעומס-השקף (#25)
    blob = " ".join(s for s, _ in body_lines(visible) if "methodica" not in s)
    for sent in re.split(r"(?<=[.!?])\s+|[\n]+", blob):
        nw = len(sent.split())
        if nw > SENTENCE_MAX:
            add(sl, f"#1 משפט ארוך (>{SENTENCE_MAX} מילים)", f'{nw} מילים: "…{sent.strip()[:48]}…"')

    # #25 עומס קריאה — ספירת מילים בטקסט גלוי (סף >70=🟠, >120=🔴)
    words = sum(len(s.split()) for s, _ in body_lines(visible) if "methodica" not in s)
    if words > 120:
        add(sl, "#25 עומס קריאה (🔴 >120)", f"~{words} מילים בשקף — לפצל/לקצר")
    elif words > 70:
        add(sl, "#25 עומס קריאה (🟠 >70)", f"~{words} מילים בשקף — לשקול קיצור/פיצול")

    # #11 'לא בתבנית' בלי הערות-דובר
    if "לא בתבנית" in visible and len(notes.strip()) < 10:
        add(sl, "#11 'לא בתבנית' ללא הערת הפקה", "הערות דובר ריקות/חסרות")

# ---- אגרגציות ברמת היחידה ----
ID_RE = re.compile(r"methodica(?:-[a-z]+)+(?:-\d+)+")
COMP_RE = re.compile(r"^methodica(?:-[a-z]+)+-\d+-\d+$")   # מזהה-רכיב = בדיוק 2 מקטעים מספריים
ids = sorted(set(ID_RE.findall(text)))
comp_ids = sorted(i for i in ids if COMP_RE.match(i))
# חוצץ-רכיב = שקף קצר (≤3 שורות גלויות, לא שקף-מטא עם טבלה גדולה) המכיל מזהה-רכיב
dividers = []
for sl, b in slides:
    vis = b.split("### Notes:")[0]
    if not any(c in vis for c in comp_ids):
        continue
    vlines = [s for s, _ in body_lines(vis) if "methodica" not in s]
    table_rows = sum(1 for ln in vis.splitlines() if ln.strip().startswith("|"))
    if len(vlines) <= 3 and table_rows < 4:
        dividers.append(sl)
ans = []
for sl, b in slides:
    mm = re.search(r"התשובה הנכונה\s*[-–:]\s*([אבגד1-9])", b)
    if mm:
        ans.append((sl, mm.group(1)))

print(f"# סריקה דטרמיניסטית — {N} שקפים\n")
print(f"שקפי-חוצץ של רכיבים: {dividers or 'לא זוהו'}")
print(f"מזהי-רכיב ({len(comp_ids)}): {comp_ids}")
print(f"סה\"כ מזהים (רכיב+פריט): {len(ids)}")
if ans:
    from collections import Counter
    pos = Counter(a for _, a in ans)
    print(f"מפתחות תשובה: {ans}  → התפלגות {dict(pos)}")
print(f"\nשקפים עם ממצא מכני: {len(findings)} מתוך {N}\n")

print("## ממצאים מכניים לפי שקף (ממצה — כל מופע)\n")
for sl in sorted(findings, key=lambda x: int(x)):
    print(f"### שקף {sl}")
    for check, detail in findings[sl]:
        print(f"- {check}: {detail}")
    print()

# סיכום לפי בדיקה
from collections import Counter
tally = Counter(c for v in findings.values() for c, _ in v)
print("## סיכום מכני לפי בדיקה")
for c, n in tally.most_common():
    print(f"- {c}: {n} מופעים")
