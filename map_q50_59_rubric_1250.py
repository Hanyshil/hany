"""
מיפוי תשובות תלמידים לשאלות 50-59 לפי מחוון_מפורט_סימולציה
Mapping 1250 Student Responses (Q50-Q59) to Rubric Levels

מחוון 6 ממדים × 5 רמות:
א. מוטיבציה ורלוונטיות (Q50, Q51)
ב. תודעת צמיחה (Q52, Q53)
ג. יוזמה ואחריות (Q54)
ד. ויסות עצמי (Q55, Q56)
ה. מודעות עצמית (Q57, Q58)
ו. תמיכה וחוויות רגשיות (Q59)

N = 1,250 תלמידים
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
# A) RUBRIC DEFINITIONS  (מחוון_מפורט_סימולציה)
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

DIM_OF_Q = {
    "Q50": "motivation", "Q51": "motivation",
    "Q52": "growth_mindset", "Q53": "growth_mindset",
    "Q54": "initiative",
    "Q55": "self_regulation", "Q56": "self_regulation",
    "Q57": "self_awareness", "Q58": "self_awareness",
    "Q59": "support",
}

AI_TOOLS = ["ChatGPT", "Claude", "Gemini", "Google Bard", "Copilot", "לא משתמש/ת"]
AI_WEIGHTS = [0.30, 0.12, 0.18, 0.10, 0.08, 0.22]

GRADES = ["ט", "י", "יא", "יב"]
GRADE_WEIGHTS = [0.25, 0.27, 0.25, 0.23]

# ============================================================
# B)  GENERATE 1,250 STUDENTS
# ============================================================
N = 1250

# Realistic dimension distributions:
# Overall roughly normal around 3, slightly right-skewed for some dims
# Based on mean avgs from rubric: motiv=3.10, growth=2.76, init=3.40,
#   reg=3.18, aware=3.04, support=2.96
DIM_DIST = {
    "motivation":       [0.06, 0.16, 0.38, 0.28, 0.12],  # p(lvl1..5)
    "growth_mindset":   [0.12, 0.22, 0.36, 0.22, 0.08],
    "initiative":       [0.05, 0.13, 0.32, 0.32, 0.18],
    "self_regulation":  [0.08, 0.17, 0.36, 0.27, 0.12],
    "self_awareness":   [0.09, 0.19, 0.38, 0.25, 0.09],
    "support":          [0.11, 0.20, 0.37, 0.22, 0.10],
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
        grade = np.random.choice(GRADES, p=GRADE_WEIGHTS)
        ai_tool = np.random.choice(AI_TOOLS, p=AI_WEIGHTS)
        ai_user = ai_tool != "לא משתמש/ת"

        # Draw a base level per dimension
        dim_scores = {}
        for dim in DIM_CODES:
            probs = DIM_DIST[dim].copy()
            # AI users: slightly higher initiative & self_regulation
            if ai_user and dim in ("initiative", "self_regulation"):
                # shift distribution up slightly
                probs = [max(0, p - 0.02) for p in probs]
                probs[2] -= 0.02
                probs[3] += 0.03
                probs[4] += 0.03
                s = sum(probs)
                probs = [p / s for p in probs]
            lvl = np.random.choice([1, 2, 3, 4, 5], p=probs)
            # Within-level score (continuous)
            lo, hi = lvl - 0.49, lvl + 0.49
            score = np.clip(np.random.uniform(lo, hi), 1.0, 5.0)
            dim_scores[dim] = round(score, 1)

        # Q answers: each question gets the parent dim level ± small noise
        q_answers = {}
        for q, dim in DIM_OF_Q.items():
            base = dim_scores[dim]
            noise = random.choice([-0.5, 0, 0, 0, 0.5])
            ans = int(np.clip(round(base + noise), 1, 5))
            q_answers[q] = ans

        # Dimension averages from raw Q answers (re-compute from questions)
        dim_avgs = {}
        for dim in DIM_CODES:
            qs = RUBRIC[dim]["questions"]
            dim_avgs[dim] = round(np.mean([q_answers[q] for q in qs]), 2)

        overall = round(np.mean(list(dim_avgs.values())), 2)
        overall_level = assign_level(overall)

        row = {
            "מזהה תלמיד": student_id,
            "כיתה": grade,
            "כלי AI בשימוש": ai_tool,
        }
        for q in ["Q50", "Q51", "Q52", "Q53", "Q54", "Q55", "Q56", "Q57", "Q58", "Q59"]:
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


print("⏳ יוצר 1,250 תלמידים...")
df = generate_students(N)
print(f"✅ נוצרו {len(df)} תלמידים")

# ============================================================
# C)  SUMMARY CALCULATIONS
# ============================================================

# Overall level distribution
level_dist = df["רמה_כללית"].value_counts().sort_index()
level_pct = (level_dist / N * 100).round(2)

# Per-dimension level distribution
dim_level_counts = {}
dim_level_pcts = {}
for dim in DIM_CODES:
    col = f"רמה_{dim}"
    counts = df[col].value_counts().sort_index()
    # fill missing levels with 0
    counts = counts.reindex([1, 2, 3, 4, 5], fill_value=0)
    dim_level_counts[dim] = counts
    dim_level_pcts[dim] = (counts / N * 100).round(2)

# AI usage stats
ai_stats = df.groupby("כלי AI בשימוש")["ממוצע_כולל"].agg(["count", "mean"]).round(2)
ai_stats.columns = ["מספר תלמידים", "ממוצע פעלנות"]
ai_stats["אחוז מסה\"כ"] = (ai_stats["מספר תלמידים"] / N * 100).round(1)

# Grade stats
grade_stats = df.groupby("כיתה")["ממוצע_כולל"].agg(["count", "mean"]).round(2)

# ============================================================
# D)  EXCEL OUTPUT
# ============================================================
COLORS = {
    "dark_nav":  "1A237E",
    "mid_blue":  "1565C0",
    "light_blue":"E3F2FD",
    "level1":    "FFCDD2",
    "level2":    "FFE0B2",
    "level3":    "FFF9C4",
    "level4":    "C8E6C9",
    "level5":    "B3E5FC",
    "white":     "FFFFFF",
    "gray_light":"F5F5F5",
    "purple":    "4A148C",
    "purple_lt": "F3E5F5",
    "teal":      "004D40",
    "teal_lt":   "E0F2F1",
}

LC = {1: COLORS["level1"], 2: COLORS["level2"], 3: COLORS["level3"],
      4: COLORS["level4"], 5: COLORS["level5"]}


def thin_border():
    s = Side(style="thin")
    return Border(left=s, right=s, top=s, bottom=s)


def write_cell(ws, row, col, value, bold=False, size=10, font_color="000000",
               bg=COLORS["white"], align="right", wrap=True, merge_to=None,
               italic=False):
    c = ws.cell(row=row, column=col, value=value)
    c.font = Font(bold=bold, size=size, color=font_color, italic=italic,
                  name="David" if any('\u0590' <= ch <= '\u05FF' for ch in str(value)) else "Calibri")
    c.fill = PatternFill("solid", fgColor=bg)
    c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    c.border = thin_border()
    if merge_to:
        ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=merge_to)
    return c


def hdr(ws, row, col, text, bold=True, size=11, font_color="FFFFFF",
        bg=COLORS["dark_nav"], merge_to=None):
    return write_cell(ws, row, col, text, bold=bold, size=size,
                      font_color=font_color, bg=bg, align="center", merge_to=merge_to)


wb = openpyxl.Workbook()
wb.remove(wb.active)

# ================================================================
# SHEET 1: סיכום התפלגות רמות (Distribution Summary)
# ================================================================
ws1 = wb.create_sheet("סיכום התפלגות רמות")
ws1.sheet_view.rightToLeft = True
ws1.sheet_view.showGridLines = False

col_widths_s1 = [32, 10, 12, 12, 12, 12, 12, 18, 18]
for i, w in enumerate(col_widths_s1, 1):
    ws1.column_dimensions[get_column_letter(i)].width = w

# ---- Title ----
hdr(ws1, 1, 1, "מיפוי תשובות תלמידים לשאלות 50–59 לפי מחוון פעלנות לומדים",
    size=15, merge_to=9)
hdr(ws1, 2, 1, f"N = {N:,} תלמידים | 6 ממדים | 5 רמות | שאלות Q50–Q59",
    size=10, bg=COLORS["mid_blue"], merge_to=9)
hdr(ws1, 3, 1, "כלים: ChatGPT, Claude, Gemini, Google Bard, Copilot | גיליון סימולציה עברית",
    size=10, bg="37474F", font_color="ECEFF1", merge_to=9)
ws1.row_dimensions[1].height = 34
ws1.row_dimensions[2].height = 22
ws1.row_dimensions[3].height = 20

# ---- A. Overall Level Distribution ----
row = 5
hdr(ws1, row, 1, "א. התפלגות רמות כללית – פעלנות לומדים (N=1,250)",
    size=12, bg=COLORS["purple"], merge_to=9)
ws1.row_dimensions[row].height = 28
row += 1

hdr(ws1, row, 1, "רמה", size=10, bg=COLORS["mid_blue"])
hdr(ws1, row, 2, "שם הרמה", size=10, bg=COLORS["mid_blue"])
hdr(ws1, row, 3, "מספר תלמידים", size=10, bg=COLORS["mid_blue"])
hdr(ws1, row, 4, "% מ-1,250", size=10, bg=COLORS["mid_blue"])
hdr(ws1, row, 5, "גרף עמודות", size=10, bg=COLORS["mid_blue"], merge_to=9)
ws1.row_dimensions[row].height = 24
row += 1

for lvl in range(1, 6):
    cnt = int(level_dist.get(lvl, 0))
    pct = float(level_pct.get(lvl, 0))
    bar = "█" * int(pct / 2)
    ws1.row_dimensions[row].height = 22
    write_cell(ws1, row, 1, f"רמה {lvl}", bold=True, size=10, bg=LC[lvl], align="center")
    write_cell(ws1, row, 2, LEVEL_NAMES_HE[lvl], bold=True, size=10, bg=LC[lvl])
    write_cell(ws1, row, 3, cnt, size=10, bg=LC[lvl], align="center")
    write_cell(ws1, row, 4, f"{pct:.1f}%", bold=True, size=10, bg=LC[lvl], align="center")
    write_cell(ws1, row, 5, bar, size=9, bg=LC[lvl], merge_to=9,
               font_color="555555", align="left")
    row += 1

# Total row
write_cell(ws1, row, 1, "סה\"כ", bold=True, size=11, bg=COLORS["mid_blue"],
           font_color="FFFFFF", align="center")
write_cell(ws1, row, 2, "כלל התלמידים", bold=True, size=10, bg=COLORS["mid_blue"],
           font_color="FFFFFF")
write_cell(ws1, row, 3, N, bold=True, size=11, bg=COLORS["mid_blue"],
           font_color="FFFFFF", align="center")
write_cell(ws1, row, 4, "100.0%", bold=True, size=11, bg=COLORS["mid_blue"],
           font_color="FFFFFF", align="center")
write_cell(ws1, row, 5, "", bg=COLORS["mid_blue"], merge_to=9)
ws1.row_dimensions[row].height = 24
row += 2

# ---- B. Per-Dimension Level Distribution ----
hdr(ws1, row, 1, "ב. התפלגות רמות לפי ממד (N=1,250 לכל ממד)",
    size=12, bg=COLORS["teal"], merge_to=9)
ws1.row_dimensions[row].height = 28
row += 1

# Header
hdr(ws1, row, 1, "ממד", size=10, bg=COLORS["mid_blue"])
hdr(ws1, row, 2, "ממוצע\nכיתתי", size=10, bg=COLORS["mid_blue"])
for lvl in range(1, 6):
    hdr(ws1, row, 2 + lvl, f"רמה {lvl}\n{LEVEL_NAMES_HE[lvl]}\nN | %",
        size=9, bg=LC[lvl],
        font_color={1: "B71C1C", 2: "E65100", 3: "827717",
                    4: "1B5E20", 5: "01579B"}[lvl])
hdr(ws1, row, 8, "ממד חזק?", size=10, bg=COLORS["mid_blue"])
hdr(ws1, row, 9, "המלצה", size=10, bg=COLORS["mid_blue"])
ws1.row_dimensions[row].height = 46
row += 1

dim_avg_overall = {}
for dim in DIM_CODES:
    avg_val = round(df[f"ממוצע_{dim}"].mean(), 2)
    dim_avg_overall[dim] = avg_val

strongest_dim = max(DIM_CODES, key=lambda d: dim_avg_overall[d])
weakest_dim = min(DIM_CODES, key=lambda d: dim_avg_overall[d])

recommendations_short = {
    "motivation": "חיזוק רלוונטיות – קשר לחיי התלמיד",
    "growth_mindset": "⚠ דחיפות: תרגול תודעת צמיחה",
    "initiative": "✅ נצל חוזקה – פרויקטים עצמאיים",
    "self_regulation": "סדנאות ניהול זמן וויסות רגשי",
    "self_awareness": "פורטפוליו רפלקטיבי + שיחות מעקב",
    "support": "חיזוק מערך עזרה הדדית בכיתה",
}

for dim in DIM_CODES:
    avg_val = dim_avg_overall[dim]
    avg_lvl = assign_level(avg_val)
    row_bg = LC[avg_lvl]
    ws1.row_dimensions[row].height = 24
    write_cell(ws1, row, 1, RUBRIC[dim]["he"], bold=True, size=10, bg=row_bg)
    write_cell(ws1, row, 2, f"{avg_val:.2f}", bold=True, size=10, bg=LC[avg_lvl],
               align="center")
    for lvl in range(1, 6):
        cnt = int(dim_level_counts[dim][lvl])
        pct = float(dim_level_pcts[dim][lvl])
        write_cell(ws1, row, 2 + lvl, f"{cnt:,}\n({pct:.1f}%)", size=9, bg=LC[lvl],
                   align="center")
    flag = "✅ חזק ביותר" if dim == strongest_dim else ("⚠ חלש ביותר" if dim == weakest_dim else "")
    write_cell(ws1, row, 8, flag, bold=(flag != ""), size=9,
               bg=LC[4] if dim == strongest_dim else (LC[1] if dim == weakest_dim else COLORS["white"]),
               align="center")
    write_cell(ws1, row, 9, recommendations_short[dim], size=9, bg=COLORS["gray_light"])
    row += 1

row += 1

# ---- C. AI Usage Distribution ----
hdr(ws1, row, 1, "ג. התפלגות שימוש בכלי AI (N=1,250) וקשר לפעלנות",
    size=12, bg="263238", merge_to=9)
ws1.row_dimensions[row].height = 28
row += 1

hdr(ws1, row, 1, "כלי AI", size=10, bg=COLORS["mid_blue"])
hdr(ws1, row, 2, "מספר תלמידים", size=10, bg=COLORS["mid_blue"])
hdr(ws1, row, 3, "% מ-1,250", size=10, bg=COLORS["mid_blue"])
hdr(ws1, row, 4, "ממוצע פעלנות", size=10, bg=COLORS["mid_blue"])
hdr(ws1, row, 5, "השוואה", size=10, bg=COLORS["mid_blue"], merge_to=9)
ws1.row_dimensions[row].height = 24
row += 1

tool_colors_map = {
    "ChatGPT": "E3F2FD", "Claude": "F3E5F5", "Gemini": "E8F5E9",
    "Google Bard": "FFF8E1", "Copilot": "FCE4EC", "לא משתמש/ת": "F5F5F5",
}
overall_ai_avg = df[df["כלי AI בשימוש"] != "לא משתמש/ת"]["ממוצע_כולל"].mean()
overall_no_ai_avg = df[df["כלי AI בשימוש"] == "לא משתמש/ת"]["ממוצע_כולל"].mean()

for tool in AI_TOOLS:
    sub = df[df["כלי AI בשימוש"] == tool]
    cnt = len(sub)
    pct = cnt / N * 100
    avg_p = sub["ממוצע_כולל"].mean() if cnt > 0 else 0
    tcolor = tool_colors_map.get(tool, COLORS["white"])
    ws1.row_dimensions[row].height = 20
    write_cell(ws1, row, 1, tool, bold=True, size=10, bg=tcolor)
    write_cell(ws1, row, 2, cnt, size=10, bg=tcolor, align="center")
    write_cell(ws1, row, 3, f"{pct:.1f}%", size=10, bg=tcolor, align="center")
    write_cell(ws1, row, 4, f"{avg_p:.2f}/5", bold=True, size=10, bg=tcolor, align="center")
    if tool == "לא משתמש/ת":
        cmp = "בסיס השוואה"
    else:
        diff = avg_p - overall_no_ai_avg
        cmp = f"{'↑' if diff > 0 else '↓'} {abs(diff):.2f} מעל לא-משתמשים"
    write_cell(ws1, row, 5, cmp, size=9, bg=tcolor, merge_to=9)
    row += 1

row += 1
write_cell(ws1, row, 1, f"ממוצע משתמשי AI: {overall_ai_avg:.2f}/5 | ממוצע לא-משתמשים: {overall_no_ai_avg:.2f}/5",
           bold=True, size=10, bg=COLORS["light_blue"], merge_to=9, align="center",
           font_color="1A237E")
row += 2

# ---- D. Grade distribution ----
hdr(ws1, row, 1, "ד. ממוצע פעלנות לפי כיתה", size=12, bg=COLORS["purple"], merge_to=9)
ws1.row_dimensions[row].height = 26
row += 1
hdr(ws1, row, 1, "כיתה", size=10, bg=COLORS["mid_blue"])
hdr(ws1, row, 2, "מספר תלמידים", size=10, bg=COLORS["mid_blue"])
hdr(ws1, row, 3, "ממוצע כולל", size=10, bg=COLORS["mid_blue"])
hdr(ws1, row, 4, "רמה שלטת", size=10, bg=COLORS["mid_blue"])
hdr(ws1, row, 5, "", bg=COLORS["mid_blue"], merge_to=9)
ws1.row_dimensions[row].height = 24
row += 1
grade_colors_map = {"ט": "E8EAF6", "י": "E8F5E9", "יא": "FFF8E1", "יב": "FCE4EC"}
for grade in GRADES:
    sub = df[df["כיתה"] == grade]
    cnt = len(sub)
    avg_g = sub["ממוצע_כולל"].mean()
    dominant_lvl = sub["רמה_כללית"].mode()[0]
    gc = grade_colors_map.get(grade, COLORS["white"])
    ws1.row_dimensions[row].height = 20
    write_cell(ws1, row, 1, f"כיתה {grade}", bold=True, size=10, bg=gc)
    write_cell(ws1, row, 2, cnt, size=10, bg=gc, align="center")
    write_cell(ws1, row, 3, f"{avg_g:.2f}/5", bold=True, size=10, bg=gc, align="center")
    write_cell(ws1, row, 4, f"רמה {dominant_lvl} – {LEVEL_NAMES_HE[dominant_lvl]}",
               size=9, bg=LC[dominant_lvl])
    write_cell(ws1, row, 5, "", bg=gc, merge_to=9)
    row += 1

print("  ✅ Sheet 1: סיכום התפלגות רמות")

# ================================================================
# SHEET 2: תשובות תלמידים לפי רמה (All Answers Sorted by Level)
# ================================================================
ws2 = wb.create_sheet("תשובות תלמידים לפי רמה")
ws2.sheet_view.rightToLeft = True
ws2.sheet_view.showGridLines = False

hdr(ws2, 1, 1, "טבלת תשובות תלמידים לפי רמת פעלנות כללית (מסונן רמה 1 → 5)",
    size=14, merge_to=17)
hdr(ws2, 2, 1, f"N = {N:,} תלמידים | שאלות Q50–Q59 | ציון 1–5 לכל שאלה",
    size=10, bg=COLORS["mid_blue"], merge_to=17)
ws2.row_dimensions[1].height = 30
ws2.row_dimensions[2].height = 20

col_widths_s2 = [8, 6, 14, 6, 6, 6, 6, 6, 6, 6, 6, 6, 6, 8, 8, 8, 14]
for i, w in enumerate(col_widths_s2, 1):
    ws2.column_dimensions[get_column_letter(i)].width = w

row = 3
hdr(ws2, row, 1, "רמה כללית", size=9, bg=COLORS["mid_blue"])
hdr(ws2, row, 2, "שם רמה", size=9, bg=COLORS["mid_blue"])
hdr(ws2, row, 3, "מזהה תלמיד", size=9, bg=COLORS["mid_blue"])
hdr(ws2, row, 4, "כיתה", size=9, bg=COLORS["mid_blue"])
for qi, q in enumerate(["Q50", "Q51", "Q52", "Q53", "Q54", "Q55", "Q56", "Q57", "Q58", "Q59"], 5):
    hdr(ws2, row, qi, q, size=9, bg=COLORS["mid_blue"])
hdr(ws2, row, 15, "ממוצע\nכולל", size=9, bg=COLORS["mid_blue"])
hdr(ws2, row, 16, "כלי AI", size=9, bg=COLORS["mid_blue"])
hdr(ws2, row, 17, "דפוס", size=9, bg=COLORS["mid_blue"])
ws2.row_dimensions[row].height = 28
row += 1

sorted_df = df.sort_values(["רמה_כללית", "ממוצע_כולל"], ascending=[True, False])

prev_level = None
for _, r in sorted_df.iterrows():
    lvl = int(r["רמה_כללית"])
    bg = LC[lvl]

    # Level separator header
    if lvl != prev_level:
        ws2.row_dimensions[row].height = 22
        hdr(ws2, row, 1, f"רמה {lvl} – {LEVEL_NAMES_HE[lvl]}",
            size=11, bg={1: "C62828", 2: "E65100", 3: "F9A825",
                         4: "2E7D32", 5: "01579B"}[lvl],
            merge_to=17)
        cnt_lvl = int(level_dist.get(lvl, 0))
        row += 1
        hdr(ws2, row, 1,
            f"({cnt_lvl:,} תלמידים = {level_pct.get(lvl,0):.1f}% מ-N=1,250)",
            size=9, bg=bg, font_color="333333", merge_to=17)
        ws2.row_dimensions[row].height = 18
        row += 1
        prev_level = lvl

    ws2.row_dimensions[row].height = 16
    write_cell(ws2, row, 1, lvl, bold=True, size=9, bg=bg, align="center")
    write_cell(ws2, row, 2, LEVEL_NAMES_HE[lvl], size=8, bg=bg)
    write_cell(ws2, row, 3, int(r["מזהה תלמיד"]), size=8, bg=bg, align="center")
    write_cell(ws2, row, 4, r["כיתה"], size=8, bg=bg, align="center")
    for qi, q in enumerate(["Q50", "Q51", "Q52", "Q53", "Q54", "Q55", "Q56", "Q57", "Q58", "Q59"], 5):
        ans = int(r[q])
        # Color the cell by the score value
        score_bg = {1: "FFCDD2", 2: "FFE0B2", 3: "FFF9C4", 4: "C8E6C9", 5: "B3E5FC"}[ans]
        write_cell(ws2, row, qi, ans, size=8, bg=score_bg, align="center")
    write_cell(ws2, row, 15, f"{r['ממוצע_כולל']:.2f}", bold=True, size=8, bg=bg, align="center")
    write_cell(ws2, row, 16, r["כלי AI בשימוש"], size=8, bg=bg)

    # Pattern tag
    tags = []
    if r["ממוצע_motivation"] >= 3.5 and r["ממוצע_self_regulation"] <= 2.5:
        tags.append("מוטיב↑ ויסות↓")
    if r["ממוצע_growth_mindset"] >= 3.5 and r["ממוצע_initiative"] <= 2.0:
        tags.append("צמיחה↑ יוזמה↓")
    if not tags:
        tags.append("מאוזן")
    write_cell(ws2, row, 17, " | ".join(tags), size=7, bg=bg)
    row += 1

print("  ✅ Sheet 2: תשובות תלמידים לפי רמה")

# ================================================================
# SHEET 3: טבלת מיפוי רמות × ממדים
# ================================================================
ws3 = wb.create_sheet("מיפוי רמות לפי ממד")
ws3.sheet_view.rightToLeft = True
ws3.sheet_view.showGridLines = False

hdr(ws3, 1, 1, "טבלת מיפוי: ספירת תלמידים ואחוזים לפי רמה ×  ממד",
    size=14, merge_to=14)
hdr(ws3, 2, 1, f"כל תא: מספר תלמידים (% מ-{N:,}) | הדגשת מקסימום בכל ממד",
    size=10, bg=COLORS["mid_blue"], merge_to=14)
ws3.row_dimensions[1].height = 30
ws3.row_dimensions[2].height = 20

col_widths_s3 = [28] + [18] * 6 + [18, 18, 18, 18, 18, 18, 18]
for i, w in enumerate(col_widths_s3, 1):
    ws3.column_dimensions[get_column_letter(i)].width = min(w, 22)

row = 3
hdr(ws3, row, 1, "ממד / רמה", size=10, bg=COLORS["dark_nav"])
for lvl in range(1, 6):
    hdr(ws3, row, 1 + lvl,
        f"רמה {lvl}\n{LEVEL_NAMES_HE[lvl]}", size=9, bg=LC[lvl],
        font_color={1: "B71C1C", 2: "E65100", 3: "827717",
                    4: "1B5E20", 5: "01579B"}[lvl])
hdr(ws3, row, 7, "ממוצע\nכיתתי", size=9, bg=COLORS["mid_blue"])
hdr(ws3, row, 8, "רמה שלטת", size=9, bg=COLORS["mid_blue"])
ws3.row_dimensions[row].height = 36
row += 1

for dim in DIM_CODES:
    avg_dim = dim_avg_overall[dim]
    dom_lvl = assign_level(avg_dim)
    max_cnt = max(dim_level_counts[dim].values)
    ws3.row_dimensions[row].height = 28
    write_cell(ws3, row, 1, RUBRIC[dim]["he"], bold=True, size=10,
               bg=COLORS["light_blue"])
    for lvl in range(1, 6):
        cnt = int(dim_level_counts[dim][lvl])
        pct = float(dim_level_pcts[dim][lvl])
        is_max = (cnt == max_cnt)
        write_cell(ws3, row, 1 + lvl,
                   f"{cnt:,}\n({pct:.1f}%)",
                   bold=is_max, size=9, bg=LC[lvl],
                   align="center",
                   font_color="1A237E" if is_max else "000000")
    write_cell(ws3, row, 7, f"{avg_dim:.2f}", bold=True, size=10, bg=LC[dom_lvl],
               align="center")
    write_cell(ws3, row, 8, f"רמה {dom_lvl}\n{LEVEL_NAMES_HE[dom_lvl]}", size=9,
               bg=LC[dom_lvl], align="center")
    row += 1

# Totals row
ws3.row_dimensions[row].height = 26
write_cell(ws3, row, 1, "סה\"כ (רמה כללית)", bold=True, size=11, bg=COLORS["dark_nav"],
           font_color="FFFFFF")
for lvl in range(1, 6):
    cnt = int(level_dist.get(lvl, 0))
    pct = float(level_pct.get(lvl, 0.0))
    write_cell(ws3, row, 1 + lvl,
               f"{cnt:,}\n({pct:.1f}%)",
               bold=True, size=10, bg=LC[lvl], align="center",
               font_color="1A237E")
overall_mean = round(df["ממוצע_כולל"].mean(), 2)
overall_dom = assign_level(overall_mean)
write_cell(ws3, row, 7, f"{overall_mean:.2f}", bold=True, size=11, bg=LC[overall_dom],
           align="center")
write_cell(ws3, row, 8, f"רמה {overall_dom}\n{LEVEL_NAMES_HE[overall_dom]}", bold=True,
           size=9, bg=LC[overall_dom], align="center")

row += 3

# ---- Rubric text per level × dimension ----
hdr(ws3, row, 1, "מחוון_מפורט_סימולציה – תיאורי רמות לפי ממד",
    size=12, bg=COLORS["purple"], merge_to=8)
ws3.row_dimensions[row].height = 26
row += 1
hdr(ws3, row, 1, "ממד", size=10, bg=COLORS["mid_blue"])
for lvl in range(1, 6):
    hdr(ws3, row, 1 + lvl,
        f"רמה {lvl} – {LEVEL_NAMES_HE[lvl]}", size=9, bg=LC[lvl],
        font_color={1: "B71C1C", 2: "E65100", 3: "827717", 4: "1B5E20", 5: "01579B"}[lvl])
ws3.row_dimensions[row].height = 28
row += 1

for dim in DIM_CODES:
    ws3.row_dimensions[row].height = 70
    write_cell(ws3, row, 1, RUBRIC[dim]["he"], bold=True, size=10, bg=COLORS["light_blue"])
    for lvl in range(1, 6):
        write_cell(ws3, row, 1 + lvl, RUBRIC[dim]["levels"][lvl], size=8, bg=LC[lvl])
    row += 1

print("  ✅ Sheet 3: מיפוי רמות לפי ממד")

# ================================================================
# SHEET 4: תשובות לפי ממד ורמה (Answers per Dimension per Level)
# ================================================================
ws4 = wb.create_sheet("תשובות לפי ממד ורמה")
ws4.sheet_view.rightToLeft = True
ws4.sheet_view.showGridLines = False

hdr(ws4, 1, 1, "ניתוח תשובות לפי ממד ורמה – מחוון_מפורט_סימולציה",
    size=14, merge_to=10)
hdr(ws4, 2, 1, f"לכל ממד: מספר תלמידים, ממוצע תשובה, אחוז מ-{N:,}",
    size=10, bg=COLORS["mid_blue"], merge_to=10)
ws4.row_dimensions[1].height = 30
ws4.row_dimensions[2].height = 20

col_widths_s4 = [28, 22, 10, 10, 10, 10, 10, 14, 14, 22]
for i, w in enumerate(col_widths_s4, 1):
    ws4.column_dimensions[get_column_letter(i)].width = w

row = 3
for dim in DIM_CODES:
    # Section header
    hdr(ws4, row, 1, RUBRIC[dim]["he"] + " | שאלות: " + ", ".join(RUBRIC[dim]["questions"]),
        size=12, bg=COLORS["purple"], merge_to=10)
    ws4.row_dimensions[row].height = 26
    row += 1

    # Column headers
    hdr(ws4, row, 1, "רמה", size=9, bg=COLORS["mid_blue"])
    hdr(ws4, row, 2, "תיאור הרמה (מחוון)", size=9, bg=COLORS["mid_blue"])
    hdr(ws4, row, 3, "N תלמידים", size=9, bg=COLORS["mid_blue"])
    hdr(ws4, row, 4, "% מ-1,250", size=9, bg=COLORS["mid_blue"])
    hdr(ws4, row, 5, "ממוצע\nתשובה", size=9, bg=COLORS["mid_blue"])
    for qi, q in enumerate(RUBRIC[dim]["questions"], 6):
        hdr(ws4, row, qi, f"{q}\nממוצע", size=9, bg=COLORS["mid_blue"])
    hdr(ws4, row, 6 + len(RUBRIC[dim]["questions"]),
        "% שימוש AI", size=9, bg=COLORS["mid_blue"])
    hdr(ws4, row, 7 + len(RUBRIC[dim]["questions"]),
        "כלי נפוץ", size=9, bg=COLORS["mid_blue"])
    ws4.row_dimensions[row].height = 30
    row += 1

    level_col = f"רמה_{dim}"
    for lvl in range(1, 6):
        sub = df[df[level_col] == lvl]
        cnt = len(sub)
        pct = cnt / N * 100
        avg_dim_score = sub[f"ממוצע_{dim}"].mean() if cnt > 0 else 0
        ai_pct = len(sub[sub["כלי AI בשימוש"] != "לא משתמש/ת"]) / cnt * 100 if cnt > 0 else 0
        top_tool = sub["כלי AI בשימוש"].mode()[0] if cnt > 0 else "–"

        ws4.row_dimensions[row].height = 60
        write_cell(ws4, row, 1, f"רמה {lvl}\n{LEVEL_NAMES_HE[lvl]}",
                   bold=True, size=9, bg=LC[lvl], align="center")
        write_cell(ws4, row, 2, RUBRIC[dim]["levels"][lvl], size=8, bg=LC[lvl])
        write_cell(ws4, row, 3, cnt, size=10, bg=LC[lvl], align="center", bold=(cnt > 0))
        write_cell(ws4, row, 4, f"{pct:.1f}%", bold=True, size=10, bg=LC[lvl], align="center")
        write_cell(ws4, row, 5, f"{avg_dim_score:.2f}" if cnt > 0 else "–",
                   size=10, bg=LC[lvl], align="center")
        for qi, q in enumerate(RUBRIC[dim]["questions"], 6):
            avg_q = sub[q].mean() if cnt > 0 else 0
            write_cell(ws4, row, qi, f"{avg_q:.2f}" if cnt > 0 else "–",
                       size=9, bg=LC[lvl], align="center")
        write_cell(ws4, row, 6 + len(RUBRIC[dim]["questions"]),
                   f"{ai_pct:.1f}%", size=9, bg=LC[lvl], align="center")
        write_cell(ws4, row, 7 + len(RUBRIC[dim]["questions"]),
                   top_tool, size=8, bg=LC[lvl])
        row += 1

    row += 1

print("  ✅ Sheet 4: תשובות לפי ממד ורמה")

# ================================================================
# SHEET 5: דפוסי שימוש ב-AI (AI Usage Patterns × Level)
# ================================================================
ws5 = wb.create_sheet("דפוסי שימוש AI × רמה")
ws5.sheet_view.rightToLeft = True
ws5.sheet_view.showGridLines = False

hdr(ws5, 1, 1, "דפוסי שימוש ב-AI לפי רמת פעלנות לומדים (Q50–Q59)",
    size=14, merge_to=9)
hdr(ws5, 2, 1, f"N={N:,} | אחוזים מכלל {N:,} תלמידים שהשיבו לשאלות על פעלנות ודפוסי שימוש ב-AI",
    size=10, bg=COLORS["mid_blue"], merge_to=9)
ws5.row_dimensions[1].height = 30
ws5.row_dimensions[2].height = 22

col_widths_s5 = [18, 14, 10, 10, 10, 12, 12, 12, 18]
for i, w in enumerate(col_widths_s5, 1):
    ws5.column_dimensions[get_column_letter(i)].width = w

row = 3

# ---- A. Cross-table: AI Tool × Overall Level ----
hdr(ws5, row, 1, "א. מטריצת כלי AI × רמת פעלנות כללית (N | %)",
    size=11, bg=COLORS["teal"], merge_to=9)
ws5.row_dimensions[row].height = 26
row += 1

hdr(ws5, row, 1, "כלי AI \\ רמה", size=9, bg=COLORS["mid_blue"])
for lvl in range(1, 6):
    hdr(ws5, row, 1 + lvl, f"רמה {lvl}\n{LEVEL_NAMES_HE[lvl]}", size=9, bg=LC[lvl],
        font_color={1: "B71C1C", 2: "E65100", 3: "827717", 4: "1B5E20", 5: "01579B"}[lvl])
hdr(ws5, row, 7, "סה\"כ", size=9, bg=COLORS["mid_blue"])
hdr(ws5, row, 8, "% מ-1,250", size=9, bg=COLORS["mid_blue"])
hdr(ws5, row, 9, "ממוצע פעלנות", size=9, bg=COLORS["mid_blue"])
ws5.row_dimensions[row].height = 36
row += 1

for tool in AI_TOOLS:
    sub = df[df["כלי AI בשימוש"] == tool]
    total_tool = len(sub)
    tcolor = tool_colors_map.get(tool, COLORS["white"])
    ws5.row_dimensions[row].height = 20
    write_cell(ws5, row, 1, tool, bold=True, size=9, bg=tcolor)
    for lvl in range(1, 6):
        cnt_tl = len(sub[sub["רמה_כללית"] == lvl])
        pct_tl = cnt_tl / N * 100
        write_cell(ws5, row, 1 + lvl,
                   f"{cnt_tl}\n({pct_tl:.1f}%)", size=8, bg=LC[lvl], align="center")
    write_cell(ws5, row, 7, total_tool, bold=True, size=9, bg=tcolor, align="center")
    write_cell(ws5, row, 8, f"{total_tool/N*100:.1f}%", size=9, bg=tcolor, align="center")
    avg_tool = sub["ממוצע_כולל"].mean() if total_tool > 0 else 0
    write_cell(ws5, row, 9, f"{avg_tool:.2f}/5", bold=True, size=9, bg=tcolor, align="center")
    row += 1

# Totals
ws5.row_dimensions[row].height = 24
write_cell(ws5, row, 1, "סה\"כ", bold=True, size=10, bg=COLORS["mid_blue"],
           font_color="FFFFFF", align="center")
for lvl in range(1, 6):
    cnt_l = int(level_dist.get(lvl, 0))
    pct_l = float(level_pct.get(lvl, 0.0))
    write_cell(ws5, row, 1 + lvl, f"{cnt_l:,}\n({pct_l:.1f}%)",
               bold=True, size=9, bg=LC[lvl], align="center")
write_cell(ws5, row, 7, N, bold=True, size=10, bg=COLORS["mid_blue"],
           font_color="FFFFFF", align="center")
write_cell(ws5, row, 8, "100.0%", bold=True, size=10, bg=COLORS["mid_blue"],
           font_color="FFFFFF", align="center")
write_cell(ws5, row, 9, f"{df['ממוצע_כולל'].mean():.2f}/5", bold=True, size=10,
           bg=COLORS["mid_blue"], font_color="FFFFFF", align="center")
row += 2

# ---- B. Q-level averages by AI tool ----
hdr(ws5, row, 1, "ב. ממוצע תשובה לכל שאלה (Q50-Q59) לפי כלי AI",
    size=11, bg=COLORS["purple"], merge_to=9)
ws5.row_dimensions[row].height = 26
row += 1

hdr(ws5, row, 1, "כלי AI", size=9, bg=COLORS["mid_blue"])
for qi, q in enumerate(["Q50", "Q51", "Q52", "Q53", "Q54"], 2):
    hdr(ws5, row, qi, q, size=9, bg=COLORS["mid_blue"])
for qi, q in enumerate(["Q55", "Q56", "Q57", "Q58", "Q59"], 7):
    pass
hdr(ws5, row, 7, "Q55", size=9, bg=COLORS["mid_blue"])
hdr(ws5, row, 8, "Q56-Q59\n(ממוצע)", size=9, bg=COLORS["mid_blue"])
hdr(ws5, row, 9, "ממוצע כולל", size=9, bg=COLORS["mid_blue"])
ws5.row_dimensions[row].height = 30
row += 1

for tool in AI_TOOLS:
    sub = df[df["כלי AI בשימוש"] == tool]
    tcolor = tool_colors_map.get(tool, COLORS["white"])
    ws5.row_dimensions[row].height = 20
    write_cell(ws5, row, 1, tool, bold=True, size=8, bg=tcolor)
    for qi, q in enumerate(["Q50", "Q51", "Q52", "Q53", "Q54"], 2):
        avg_q = sub[q].mean() if len(sub) > 0 else 0
        write_cell(ws5, row, qi, f"{avg_q:.2f}", size=8, bg=tcolor, align="center")
    avg_q55 = sub["Q55"].mean() if len(sub) > 0 else 0
    write_cell(ws5, row, 7, f"{avg_q55:.2f}", size=8, bg=tcolor, align="center")
    avg_56_59 = sub[["Q56", "Q57", "Q58", "Q59"]].mean().mean() if len(sub) > 0 else 0
    write_cell(ws5, row, 8, f"{avg_56_59:.2f}", size=8, bg=tcolor, align="center")
    avg_all = sub["ממוצע_כולל"].mean() if len(sub) > 0 else 0
    write_cell(ws5, row, 9, f"{avg_all:.2f}/5", bold=True, size=9, bg=tcolor, align="center")
    row += 1

row += 2

# ---- C. Key patterns ----
hdr(ws5, row, 1, "ג. דפוסי מפתח שנמצאו בנתונים (N=1,250)",
    size=11, bg="263238", font_color="FFFFFF", merge_to=9)
ws5.row_dimensions[row].height = 26
row += 1

patterns_data = [
    ("מוטיבציה גבוהה + ויסות עצמי נמוך",
     len(df[(df["ממוצע_motivation"] >= 3.5) & (df["ממוצע_self_regulation"] <= 2.5)]),
     "תרגול ניהול זמן + מיינדפולנס"),
    ("תודעת צמיחה גבוהה + יוזמה נמוכה",
     len(df[(df["ממוצע_growth_mindset"] >= 3.5) & (df["ממוצע_initiative"] <= 2.0)]),
     "העצמה לפרויקטים עצמאיים"),
    ("מודעות עצמית נמוכה + תמיכה גבוהה",
     len(df[(df["ממוצע_self_awareness"] <= 2.5) & (df["ממוצע_support"] >= 4.0)]),
     "פורטפוליו רפלקטיבי אישי"),
    ("פרופיל מאוזן-גבוה (כל הממדים ≥ 3.5)",
     len(df[(df[[f'ממוצע_{d}' for d in DIM_CODES]] >= 3.5).all(axis=1)]),
     "מנהיגות עמיתים + למידה מורחבת"),
    ("פרופיל מאוזן-נמוך (כל הממדים ≤ 2.5)",
     len(df[(df[[f'ממוצע_{d}' for d in DIM_CODES]] <= 2.5).all(axis=1)]),
     "⚠ התערבות מיידית + ליווי אישי"),
]

hdr(ws5, row, 1, "דפוס", size=9, bg=COLORS["mid_blue"])
hdr(ws5, row, 2, "N תלמידים", size=9, bg=COLORS["mid_blue"])
hdr(ws5, row, 3, "% מ-1,250", size=9, bg=COLORS["mid_blue"])
hdr(ws5, row, 4, "המלצה", size=9, bg=COLORS["mid_blue"], merge_to=9)
ws5.row_dimensions[row].height = 24
row += 1

for desc, cnt, rec in patterns_data:
    pct = cnt / N * 100
    bg = COLORS["level2"] if cnt / N > 0.15 else COLORS["gray_light"]
    ws5.row_dimensions[row].height = 22
    write_cell(ws5, row, 1, desc, size=9, bg=bg)
    write_cell(ws5, row, 2, cnt, bold=True, size=10, bg=bg, align="center")
    write_cell(ws5, row, 3, f"{pct:.1f}%", bold=True, size=10, bg=bg, align="center")
    write_cell(ws5, row, 4, rec, size=9, bg=COLORS["gray_light"], merge_to=9)
    row += 1

print("  ✅ Sheet 5: דפוסי שימוש AI × רמה")

# ================================================================
# SAVE
# ================================================================
output_path = "/home/user/hany/מיפוי_תשובות_Q50_59_מחוון_1250.xlsx"
wb.save(output_path)
print(f"\n✅ הקובץ נשמר: {output_path}")
print(f"   גיליונות: {[ws.title for ws in wb.worksheets]}")

# ---- Print summary to terminal ----
print("\n" + "=" * 70)
print("  תוצאות עיקריות – N = 1,250 תלמידים")
print("=" * 70)
print(f"\n  התפלגות רמות כללית:")
for lvl in range(1, 6):
    cnt = int(level_dist.get(lvl, 0))
    pct = float(level_pct.get(lvl, 0.0))
    bar = "█" * int(pct / 2)
    print(f"    רמה {lvl} – {LEVEL_NAMES_HE[lvl]:<14}: {cnt:>4,} תלמידים ({pct:>5.1f}%) {bar}")

print(f"\n  ממוצע ממדים:")
for dim in DIM_CODES:
    avg = dim_avg_overall[dim]
    dom = assign_level(avg)
    marker = " ← חזק ביותר" if dim == strongest_dim else (" ← חלש ביותר" if dim == weakest_dim else "")
    print(f"    {RUBRIC[dim]['he']:<28}: {avg:.2f}/5 (רמה {dom} – {LEVEL_NAMES_HE[dom]}){marker}")

print(f"\n  שימוש ב-AI:")
for tool in AI_TOOLS:
    sub = df[df["כלי AI בשימוש"] == tool]
    print(f"    {tool:<16}: {len(sub):>4,} ({len(sub)/N*100:.1f}%) | ממוצע פעלנות: {sub['ממוצע_כולל'].mean():.2f}")

print(f"\n  ממוצע כולל: {df['ממוצע_כולל'].mean():.2f}/5")
print("=" * 70)
