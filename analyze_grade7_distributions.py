"""
התפלגות רמות לכל שאלה (Q50-Q59) – כלל האוכלוסייה כיתה ז
N = 1,250 תלמידים | מחוון פעלנות לומדים (6 ממדים × 5 רמות)
"""

import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import random

random.seed(2024)
np.random.seed(2024)

# ============================================================
# מחוון פעלנות לומדים – זהה למחוון_מפורט_סימולציה
# ============================================================
RUBRIC = {
    "motivation": {
        "he": "א. מוטיבציה ורלוונטיות",
        "questions": ["Q50", "Q51"],
        "levels": {
            1: "אינו מוצא קשר בין הלמידה לחייו. לא מעוניין במשימות מאתגרות. פועל רק מדחף חיצוני.",
            2: "מזדהה לעיתים עם החומר. מוכן לאתגר בתנאי שיש תמריץ מיידי (ציון).",
            3: "מתעניין בחלק מהנושאים. מגלה מוכנות לאתגר כשהחומר נוגע אליו.",
            4: "רואה רלוונטיות ברוב הנושאים. משקיע מעבר לדרישות המינימום.",
            5: "מחובר עמוקות ללמידה; רואה כל נושא כהזדמנות. יוזם למידה עצמאית מרצון.",
        },
    },
    "growth_mindset": {
        "he": "ב. תודעת צמיחה",
        "questions": ["Q52", "Q53"],
        "levels": {
            1: "תופס יכולת כקבועה. מתייאש בפני כישלון. נוטה לוותר מהר.",
            2: "מאמין בשיפור תיאורטי אך מתקשה ליישם. קשה לו לראות בטעות הזדמנות.",
            3: "מסוגל להתמיד כשהקושי אינו גדול מדי. מבין שמאמץ עוזר.",
            4: "רואה בקשיים אתגרים. מתאושש מכישלונות ולומד מהם.",
            5: "חי לפי עקרון הצמיחה; כישלון הוא מנוף. מחפש אתגרים באופן יזום.",
        },
    },
    "initiative": {
        "he": "ג. יוזמה ואחריות",
        "questions": ["Q54"],
        "levels": {
            1: "ממתין להוראות בלבד. אינו לוקח אחריות. תלוי לחלוטין במורה.",
            2: "לפעמים מציע רעיונות אך נסוג. קצת אחריות בלחץ חיצוני.",
            3: "לוקח אחריות בסיסית. פועל עצמאית במשימות שגרתיות.",
            4: "יוזם פעולות ומציע שינויים. לוקח אחריות גם כשדברים אינם הולכים כשורה.",
            5: "מוביל, מחדש ומשפיע על הסביבה. אחריות עמוקה גם מעבר לתחום האישי.",
        },
    },
    "self_regulation": {
        "he": "ד. ויסות עצמי",
        "questions": ["Q55", "Q56"],
        "levels": {
            1: "אינו מתכנן. נסחף בלחצים. מתקשה להמשיך תחת תסכול.",
            2: "מתכנן לפעמים אך לא ממשיך. ויסות רגשי חלקי.",
            3: "מתכנן משימות בסיסיות. מצליח להמשיך לרוב גם תחת לחץ.",
            4: "מנהל זמן ביעילות. מתמודד טוב עם תסכול ושומר על יציבות.",
            5: "ויסות עצמי מלא: תכנון אסטרטגי, שליטה רגשית, גמישות תחת לחץ.",
        },
    },
    "self_awareness": {
        "he": "ה. מודעות עצמית",
        "questions": ["Q57", "Q58"],
        "levels": {
            1: "אינו מכיר חוזקות/חולשות. לא מפיק לקחים מניסיונות.",
            2: "מודע חלקית לחוזקות. לפעמים חושב על כישלונות אך לא מסיק מסקנות.",
            3: "מכיר את עצמו בתחומים מרכזיים. מפיק לקחים בסיסיים.",
            4: "מודע לדפוסים אישיים. מפיק לקחים ומיישם בפועל.",
            5: "מודעות עצמית עמוקה; מנתח התנהגות, מזהה צמיחה ומתאים גישה.",
        },
    },
    "support": {
        "he": "ו. תמיכה וחוויות רגשיות",
        "questions": ["Q59"],
        "levels": {
            1: "לא מרגיש תמיכה. אינו פונה לעזרה גם כשצריך.",
            2: "לפעמים מבקש עזרה אך חש שאין לו תמיכה אמיתית.",
            3: "מרגיש תמיכה בסיסית. פונה לעזרה כשקשה מאוד.",
            4: "מרגיש תמיכה ממקורות מגוונים. פונה לעזרה כחלק שוטף.",
            5: "רשת תמיכה חזקה. מנצל ומספק עזרה הדדית. חלק מקהילה לומדת.",
        },
    },
}

DIM_CODES = list(RUBRIC.keys())

LEVEL_NAMES_HE = {
    1: "בתחילת הדרך",
    2: "מתפתח/ת",
    3: "מתקדם/ת",
    4: "מיומן/ת",
    5: "מומחה/ית",
}

QUESTION_TEXTS = {
    "Q50": "עד כמה השיעורים מעניינים אותך ואתה מרגיש שהלמידה רלוונטית לחיים שלך?",
    "Q51": "עד כמה אתה מוכן להתמודד עם משימות מאתגרות ומשקיע בלמידה מעבר לציונים?",
    "Q52": "עד כמה אתה מאמין שאתה יכול להשתפר בכל דבר באמצעות מאמץ והתמדה?",
    "Q53": "כשאתה נתקל בקושי או טועה, עד כמה אתה רואה בזה הזדמנות ללמוד ולא מוותר?",
    "Q54": "עד כמה אתה לוקח אחריות על הלמידה שלך, פועל בעצמאות ומציע יוזמות?",
    "Q55": "עד כמה אתה מתכנן משימות מראש ומנהל את הזמן שלך בצורה יעילה?",
    "Q56": "כשאתה מתוסכל או לחוץ, עד כמה אתה מצליח להרגיע את עצמך ולהמשיך ללמוד?",
    "Q57": "עד כמה אתה מכיר את החוזקות והחולשות שלך בלמידה?",
    "Q58": "עד כמה אתה חושב על ההצלחות והכישלונות שלך ומפיק לקחים?",
    "Q59": "עד כמה אתה יודע לבקש עזרה כשצריך ומרגיש שיש לך תמיכה (ממורים, חברים, משפחה)?",
}

QUESTION_DIM = {
    "Q50": "motivation", "Q51": "motivation",
    "Q52": "growth_mindset", "Q53": "growth_mindset",
    "Q54": "initiative",
    "Q55": "self_regulation", "Q56": "self_regulation",
    "Q57": "self_awareness", "Q58": "self_awareness",
    "Q59": "support",
}

QUESTIONS_ORDER = ["Q50", "Q51", "Q52", "Q53", "Q54", "Q55", "Q56", "Q57", "Q58", "Q59"]

AI_TOOLS = ["ChatGPT", "Claude", "Gemini", "Google Bard", "Copilot", "לא משתמש/ת"]
AI_WEIGHTS = [0.30, 0.12, 0.18, 0.10, 0.08, 0.22]

N = 1250

DIM_DIST = {
    "motivation":      [0.06, 0.16, 0.38, 0.28, 0.12],
    "growth_mindset":  [0.12, 0.22, 0.36, 0.22, 0.08],
    "initiative":      [0.05, 0.13, 0.32, 0.32, 0.18],
    "self_regulation": [0.08, 0.17, 0.36, 0.27, 0.12],
    "self_awareness":  [0.09, 0.19, 0.38, 0.25, 0.09],
    "support":         [0.11, 0.20, 0.37, 0.22, 0.10],
}


def assign_level(score):
    if score < 1.5:
        return 1
    elif score < 2.5:
        return 2
    elif score < 3.5:
        return 3
    elif score < 4.5:
        return 4
    return 5


def generate_students(n):
    rows = []
    student_id = 1001
    for i in range(n):
        ai_tool = np.random.choice(AI_TOOLS, p=AI_WEIGHTS)
        ai_user = ai_tool != "לא משתמש/ת"

        dim_scores = {}
        for dim in DIM_CODES:
            probs = DIM_DIST[dim].copy()
            if ai_user and dim in ("initiative", "self_regulation"):
                probs = [max(0, p - 0.02) for p in probs]
                probs[2] -= 0.02
                probs[3] += 0.03
                probs[4] += 0.03
                s = sum(probs)
                probs = [p / s for p in probs]
            lvl = np.random.choice([1, 2, 3, 4, 5], p=probs)
            lo, hi = lvl - 0.49, lvl + 0.49
            score = np.clip(np.random.uniform(lo, hi), 1.0, 5.0)
            dim_scores[dim] = round(score, 1)

        q_answers = {}
        for q, dim in QUESTION_DIM.items():
            base = dim_scores[dim]
            noise = random.choice([-0.5, 0, 0, 0, 0.5])
            ans = int(np.clip(round(base + noise), 1, 5))
            q_answers[q] = ans

        dim_avgs = {}
        for dim in DIM_CODES:
            qs = RUBRIC[dim]["questions"]
            dim_avgs[dim] = round(np.mean([q_answers[q] for q in qs]), 2)

        overall = round(np.mean(list(dim_avgs.values())), 2)
        overall_level = assign_level(overall)

        row = {
            "מזהה תלמיד": student_id,
            "כיתה": "ז",          # כל התלמידים – כיתה ז
            "כלי AI בשימוש": ai_tool,
        }
        for q in QUESTIONS_ORDER:
            row[q] = q_answers[q]
        for dim in DIM_CODES:
            row[f"ממוצע_{dim}"] = dim_avgs[dim]
            row[f"רמה_{dim}"] = assign_level(dim_avgs[dim])
        row["ממוצע_כולל"] = overall
        row["רמה_כללית"] = overall_level
        row["שם_רמה_כללית"] = LEVEL_NAMES_HE[overall_level]
        rows.append(row)
        student_id += 1

    return pd.DataFrame(rows)


print("⏳ יוצר 1,250 תלמידי כיתה ז...")
df = generate_students(N)
print(f"✅ נוצרו {len(df)} תלמידים (כיתה ז בלבד)")

# ============================================================
# חישוב התפלגויות
# ============================================================

# לכל שאלה – ספירה לפי רמה (1-5)
q_dist = {}
for q in QUESTIONS_ORDER:
    counts = df[q].value_counts().sort_index()
    q_dist[q] = {lvl: int(counts.get(lvl, 0)) for lvl in range(1, 6)}

# לכל ממד – ספירה לפי רמה
dim_dist = {}
for dim in DIM_CODES:
    counts = df[f"רמה_{dim}"].value_counts().sort_index()
    dim_dist[dim] = {lvl: int(counts.get(lvl, 0)) for lvl in range(1, 6)}

# רמה כללית
overall_dist = {lvl: int(df["רמה_כללית"].value_counts().sort_index().get(lvl, 0))
                for lvl in range(1, 6)}

# ============================================================
# עיצוב Excel
# ============================================================

LEVEL_COLORS = {
    1: "FFCDD2",  # אדום בהיר
    2: "FFE0B2",  # כתום בהיר
    3: "FFF9C4",  # צהוב בהיר
    4: "C8E6C9",  # ירוק בהיר
    5: "B3E5FC",  # כחול בהיר
}

DIM_BG = {
    "motivation":      "E8EAF6",
    "growth_mindset":  "E8F5E9",
    "initiative":      "FFF8E1",
    "self_regulation": "E0F2F1",
    "self_awareness":  "FCE4EC",
    "support":         "F3E5F5",
}

wb = openpyxl.Workbook()
wb.remove(wb.active)


def is_rtl(text):
    return any('\u0590' <= ch <= '\u05FF' for ch in str(text))


def cs(ws, row, col, value, bold=False, size=10, font_color="000000",
       bg=None, wrap=False, align_h="center", align_v="center",
       merge_to=None, italic=False):
    c = ws.cell(row=row, column=col, value=value)
    c.font = Font(bold=bold, size=size, color=font_color, italic=italic,
                  name="David" if is_rtl(str(value)) else "Calibri")
    if bg:
        c.fill = PatternFill("solid", fgColor=bg)
    c.alignment = Alignment(horizontal=align_h, vertical=align_v,
                             wrap_text=wrap)
    thin = Side(style="thin", color="BDBDBD")
    c.border = Border(left=thin, right=thin, top=thin, bottom=thin)
    if merge_to:
        ws.merge_cells(start_row=row, start_column=col,
                       end_row=row, end_column=merge_to)
    return c


def hdr(ws, row, col, text, size=11, merge_to=None, bg="1A237E", fc="FFFFFF", bold=True):
    return cs(ws, row, col, text, bold=bold, size=size,
              font_color=fc, bg=bg, merge_to=merge_to, wrap=True)


# ====================================================================
# גיליון 1: התפלגות לפי שאלה – טבלה ראשית
# ====================================================================
ws1 = wb.create_sheet("התפלגות לפי שאלה")
ws1.sheet_view.rightToLeft = True

hdr(ws1, 1, 1, "התפלגות רמות לכל שאלה – כלל האוכלוסייה", size=15, merge_to=11)
hdr(ws1, 2, 1,
    f"N = {N:,} תלמידי כיתה ז  |  10 שאלות  |  5 רמות  |  מחוון פעלנות לומדים",
    size=10, bg="37474F", merge_to=11)
hdr(ws1, 3, 1,
    "ממד – שאלה  |  טקסט השאלה  |  ממוצע  |  רמה 1 עד רמה 5 (N וּ-%)",
    size=9, bg="546E7A", merge_to=11)

row = 5

# כותרות עמודות
for c_val, col_i, bg in [
    ("ממד",        1, "263238"),
    ("שאלה",       2, "263238"),
    ("טקסט השאלה", 3, "263238"),
    ("ממוצע",      5, "263238"),
]:
    hdr(ws1, row, col_i, c_val, size=10, bg=bg)

ws1.merge_cells(start_row=row, start_column=3, end_row=row, end_column=4)

for lvl in range(1, 6):
    hdr(ws1, row, 5 + lvl,
        f"רמה {lvl}\n{LEVEL_NAMES_HE[lvl]}",
        bg=LEVEL_COLORS[lvl], fc="000000", size=9)

row += 1

current_dim = None
for q in QUESTIONS_ORDER:
    dim = QUESTION_DIM[q]
    bg = DIM_BG[dim]
    dim_name = RUBRIC[dim]["he"]
    q_text = QUESTION_TEXTS[q]
    mean_val = round(df[q].mean(), 2)

    # ממד – מיזוג שתי שורות לאותו ממד
    if dim != current_dim:
        cs(ws1, row, 1, dim_name, bold=True, size=10, bg=bg, wrap=True)
        current_dim = dim
    else:
        cs(ws1, row, 1, "", bg=bg)

    cs(ws1, row, 2, q, bold=True, size=11, bg=bg)
    cs(ws1, row, 3, q_text, size=9, bg=bg, wrap=True, align_h="right", merge_to=4)
    cs(ws1, row, 5, mean_val, bold=True, size=11, bg="ECEFF1")

    for lvl in range(1, 6):
        count = q_dist[q][lvl]
        pct = count / N * 100
        cs(ws1, row, 5 + lvl,
           f"{count}\n({pct:.1f}%)",
           size=9, bg=LEVEL_COLORS[lvl], wrap=True)

    ws1.row_dimensions[row].height = 50
    row += 1

# רוחב עמודות
ws1.column_dimensions["A"].width = 24
ws1.column_dimensions["B"].width = 8
ws1.column_dimensions["C"].width = 32
ws1.column_dimensions["D"].width = 8
ws1.column_dimensions["E"].width = 10
for letter in ["F", "G", "H", "I", "J", "K"]:
    ws1.column_dimensions[letter].width = 14


# ====================================================================
# גיליון 2: מחוון מפורט – תיאור רמות לכל ממד + N
# ====================================================================
ws2 = wb.create_sheet("מחוון מפורט")
ws2.sheet_view.rightToLeft = True

hdr(ws2, 1, 1, "מחוון פעלנות לומדים – תיאור רמות לכל ממד", size=14, merge_to=7)
hdr(ws2, 2, 1,
    "N = 1,250 תלמידי כיתה ז  |  Q50–Q59  |  6 ממדים × 5 רמות",
    size=10, bg="37474F", merge_to=7)

row = 4
hdr(ws2, row, 1, "ממד", bg="263238", size=10)
hdr(ws2, row, 2, "שאלות", bg="263238", size=10)
for lvl in range(1, 6):
    hdr(ws2, row, 2 + lvl,
        f"רמה {lvl}\n{LEVEL_NAMES_HE[lvl]}",
        bg=LEVEL_COLORS[lvl], fc="000000", size=9)
row += 1

for dim in DIM_CODES:
    bg = DIM_BG[dim]
    dim_name = RUBRIC[dim]["he"]
    qs_str = "  |  ".join(RUBRIC[dim]["questions"])

    cs(ws2, row, 1, dim_name, bold=True, size=10, bg=bg, wrap=True)
    cs(ws2, row, 2, qs_str, size=9, bg=bg, wrap=True)

    for lvl in range(1, 6):
        level_text = RUBRIC[dim]["levels"][lvl]
        count = dim_dist[dim][lvl]
        pct = count / N * 100
        full_text = f"{level_text}\n\nN = {count}  ({pct:.1f}%)"
        cs(ws2, row, 2 + lvl, full_text,
           size=9, bg=LEVEL_COLORS[lvl], wrap=True, align_h="right")

    ws2.row_dimensions[row].height = 90
    row += 1

ws2.column_dimensions["A"].width = 24
ws2.column_dimensions["B"].width = 12
for letter in ["C", "D", "E", "F", "G"]:
    ws2.column_dimensions[letter].width = 30


# ====================================================================
# גיליון 3: סיכום כולל – רמה כללית + לפי ממד
# ====================================================================
ws3 = wb.create_sheet("סיכום כלל האוכלוסייה")
ws3.sheet_view.rightToLeft = True

hdr(ws3, 1, 1, "סיכום – כל 1,250 תלמידי כיתה ז", size=14, merge_to=7)
hdr(ws3, 2, 1,
    f"ממוצע כולל: {df['ממוצע_כולל'].mean():.2f}/5  |  כיתה ז בלבד",
    size=10, bg="37474F", merge_to=7)

row = 4
hdr(ws3, row, 1, "התפלגות רמות כללית", size=12, merge_to=7, bg="455A64")
row += 1

hdr(ws3, row, 1, "רמה",            bg="263238")
hdr(ws3, row, 2, "שם הרמה",        bg="263238")
hdr(ws3, row, 3, "מספר תלמידים",   bg="263238")
hdr(ws3, row, 4, "% מ-1,250",      bg="263238")
hdr(ws3, row, 5, "גרף",            bg="263238", merge_to=7)
row += 1

for lvl in range(1, 6):
    count = overall_dist[lvl]
    pct = count / N * 100
    bar = "█" * max(1, int(pct / 2))
    cs(ws3, row, 1, f"רמה {lvl}", bold=True, bg=LEVEL_COLORS[lvl])
    cs(ws3, row, 2, LEVEL_NAMES_HE[lvl], bg=LEVEL_COLORS[lvl])
    cs(ws3, row, 3, count, bold=True, bg=LEVEL_COLORS[lvl])
    cs(ws3, row, 4, f"{pct:.1f}%", bold=True, bg=LEVEL_COLORS[lvl])
    cs(ws3, row, 5, bar, bg=LEVEL_COLORS[lvl], align_h="left", merge_to=7)
    row += 1

cs(ws3, row, 1, 'סה"כ',         bold=True, bg="ECEFF1")
cs(ws3, row, 2, "כלל התלמידים", bold=True, bg="ECEFF1")
cs(ws3, row, 3, N,              bold=True, bg="ECEFF1")
cs(ws3, row, 4, "100.0%",       bold=True, bg="ECEFF1")
cs(ws3, row, 5, "",              bg="ECEFF1", merge_to=7)

row += 2

# התפלגות לפי ממד
hdr(ws3, row, 1, "התפלגות לפי ממד", size=12, merge_to=7, bg="455A64")
row += 1

hdr(ws3, row, 1, "ממד",    bg="263238")
hdr(ws3, row, 2, "ממוצע", bg="263238")
for lvl in range(1, 6):
    hdr(ws3, row, 2 + lvl,
        f"רמה {lvl}\n{LEVEL_NAMES_HE[lvl]}",
        bg=LEVEL_COLORS[lvl], fc="000000", size=9)
row += 1

for dim in DIM_CODES:
    bg = DIM_BG[dim]
    dim_name = RUBRIC[dim]["he"]
    mean_val = round(df[f"ממוצע_{dim}"].mean(), 2)
    cs(ws3, row, 1, dim_name, bold=True, size=10, bg=bg, wrap=True)
    cs(ws3, row, 2, f"{mean_val}/5", bold=True, bg=bg)
    for lvl in range(1, 6):
        count = dim_dist[dim][lvl]
        pct = count / N * 100
        cs(ws3, row, 2 + lvl,
           f"{count}\n({pct:.1f}%)",
           size=9, bg=LEVEL_COLORS[lvl], wrap=True)
    ws3.row_dimensions[row].height = 35
    row += 1

for letter, width in zip(["A", "B", "C", "D", "E", "F", "G"],
                          [26, 12, 14, 14, 14, 14, 14]):
    ws3.column_dimensions[letter].width = width


# ====================================================================
# שמירה
# ====================================================================
OUTPUT = "התפלגות_שאלות_כיתה_ז_N1250.xlsx"
wb.save(OUTPUT)
print(f"\n✅ נשמר: {OUTPUT}")

# ============================================================
# הדפסת סיכום לקונסול
# ============================================================
print("\n" + "=" * 65)
print(f"התפלגות רמות כללית  (N={N} תלמידי כיתה ז)")
print("=" * 65)
for lvl in range(1, 6):
    count = overall_dist[lvl]
    pct = count / N * 100
    print(f"  רמה {lvl} – {LEVEL_NAMES_HE[lvl]:<12}: {count:5d}  ({pct:.1f}%)")

print("\nממוצעים לפי שאלה:")
print("-" * 45)
for q in QUESTIONS_ORDER:
    dim = QUESTION_DIM[q]
    dim_he = RUBRIC[dim]["he"]
    mean = df[q].mean()
    print(f"  {q}  [{dim_he}]  →  {mean:.2f}/5")

print("\nממוצעים לפי ממד:")
print("-" * 45)
for dim in DIM_CODES:
    mean = df[f"ממוצע_{dim}"].mean()
    print(f"  {RUBRIC[dim]['he']:<26}: {mean:.2f}/5")

print("\nהתפלגות לפי שאלה (N לכל רמה):")
print("-" * 65)
header = f"{'שאלה':<6}  " + "  ".join(
    [f"{'רמה ' + str(l):<10}" for l in range(1, 6)] + ["ממוצע"]
)
print(header)
for q in QUESTIONS_ORDER:
    parts = [f"{q:<6}"]
    for lvl in range(1, 6):
        count = q_dist[q][lvl]
        pct = count / N * 100
        parts.append(f"{count:3d}({pct:4.1f}%)")
    parts.append(f"{df[q].mean():.2f}")
    print("  ".join(parts))
