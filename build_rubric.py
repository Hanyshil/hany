"""
מחוון הערכה לאוריינות בינה מלאכותית - דפוסי שימוש
בניית מחוון מלא עם ממצאים, הערות והצעות לשיפור
מבוסס על ניתוח נתוני שאלות 50-59
"""

import openpyxl
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, GradientFill
)
from openpyxl.utils import get_column_letter
import pandas as pd
import numpy as np

# ===== Load Data =====
df_q = pd.read_excel("student_proactiveness_q50_59.xlsx")
df_ai = pd.read_excel("student_ai_usage.xlsx")

# Merge on student name
df = pd.merge(df_q, df_ai, on=["שם התלמיד/ה", "כיתה", "כלי AI בשימוש"], how="left")

q_cols = [c for c in df_q.columns if c.startswith("Q5")]

dimension_map = {
    "Q50": "motivation", "Q51": "motivation",
    "Q52": "growth_mindset", "Q53": "growth_mindset",
    "Q54": "initiative",
    "Q55": "self_regulation", "Q56": "self_regulation",
    "Q57": "self_awareness", "Q58": "self_awareness",
    "Q59": "support"
}

dim_codes = ["motivation", "growth_mindset", "initiative", "self_regulation", "self_awareness", "support"]

dim_names_he = {
    "motivation": "א. מוטיבציה ורלוונטיות",
    "growth_mindset": "ב. תודעת צמיחה",
    "initiative": "ג. יוזמה ואחריות",
    "self_regulation": "ד. ויסות עצמי",
    "self_awareness": "ה. מודעות עצמית",
    "support": "ו. תמיכה וחוויות רגשיות"
}

dim_questions = {
    "motivation": ["Q50", "Q51"],
    "growth_mindset": ["Q52", "Q53"],
    "initiative": ["Q54"],
    "self_regulation": ["Q55", "Q56"],
    "self_awareness": ["Q57", "Q58"],
    "support": ["Q59"]
}

dim_q_labels = {
    "motivation": "Q50: רלוונטיות וענין | Q51: מוכנות להתמודדות",
    "growth_mindset": "Q52: אמונה ביכולת שיפור | Q53: ראיית קושי כהזדמנות",
    "initiative": "Q54: אחריות ויוזמה",
    "self_regulation": "Q55: תכנון וניהול זמן | Q56: ויסות רגשי תחת לחץ",
    "self_awareness": "Q57: הכרת חוזקות/חולשות | Q58: הפקת לקחים",
    "support": "Q59: בקשת עזרה ותמיכה חברתית"
}

# Compute dimension scores
for dim in dim_codes:
    dim_q = [c for c in q_cols if dimension_map.get(c.split(":")[0]) == dim]
    df[f"dim_{dim}"] = df[dim_q].mean(axis=1).round(1)

dim_cols = [f"dim_{d}" for d in dim_codes]
df["overall"] = df[dim_cols].mean(axis=1).round(1)

def level(score):
    if score < 1.5: return 1
    elif score < 2.5: return 2
    elif score < 3.5: return 3
    elif score < 4.5: return 4
    else: return 5

df["overall_level"] = df["overall"].apply(level)
for dim in dim_codes:
    df[f"level_{dim}"] = df[f"dim_{dim}"].apply(level)

# ===== Style Helpers =====
COLORS = {
    "header_dark":   "1A237E",  # dark navy
    "header_mid":    "1565C0",  # medium blue
    "header_light":  "E3F2FD",  # light blue
    "level1":        "FFCDD2",  # red-pink
    "level2":        "FFE0B2",  # orange-light
    "level3":        "FFF9C4",  # yellow
    "level4":        "C8E6C9",  # light green
    "level5":        "B3E5FC",  # light blue
    "section":       "F3E5F5",  # purple-light
    "gold":          "F9A825",
    "white":         "FFFFFF",
    "gray_light":    "F5F5F5",
    "gray_mid":      "BDBDBD",
    "warning":       "FF6F00",
    "success":       "2E7D32",
}

level_colors = {
    1: COLORS["level1"],
    2: COLORS["level2"],
    3: COLORS["level3"],
    4: COLORS["level4"],
    5: COLORS["level5"],
}

level_names_he = {
    1: "בתחילת הדרך",
    2: "מתפתח/ת",
    3: "מתקדם/ת",
    4: "מיומן/ת",
    5: "מומחה/ית"
}

def make_border(style="thin"):
    s = Side(style=style)
    return Border(left=s, right=s, top=s, bottom=s)

def hdr(ws, row, col, text, bold=True, size=11, color="FFFFFF", bg=COLORS["header_dark"],
        align="center", wrap=True, merge_to=None):
    cell = ws.cell(row=row, column=col, value=text)
    cell.font = Font(bold=bold, size=size, color=color,
                     name="David" if any('\u0590' <= c <= '\u05FF' for c in str(text)) else "Calibri")
    cell.fill = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap, reading_order=2)
    cell.border = make_border()
    if merge_to:
        ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=merge_to)
    return cell

def cell(ws, row, col, text, bold=False, size=10, color="000000", bg=COLORS["white"],
         align="right", wrap=True, border=True):
    c = ws.cell(row=row, column=col, value=text)
    c.font = Font(bold=bold, size=size, color=color,
                  name="David" if any('\u0590' <= c <= '\u05FF' for c in str(text)) else "Calibri")
    c.fill = PatternFill("solid", fgColor=bg)
    c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap, reading_order=2)
    if border:
        c.border = make_border()
    return c

# ===================================================================
# CREATE WORKBOOK
# ===================================================================
wb = openpyxl.Workbook()
wb.remove(wb.active)

# ===================================================================
# SHEET 1: מחוון ראשי - הגדרות רמות
# ===================================================================
ws1 = wb.create_sheet("מחוון ראשי")
ws1.sheet_view.rightToLeft = True
ws1.sheet_view.showGridLines = False

# Title
hdr(ws1, 1, 1, "מחוון הערכה – אוריינות בינה מלאכותית ופעלנות לומדים", bold=True, size=16,
    bg=COLORS["header_dark"], merge_to=9)
hdr(ws1, 2, 1, "מבוסס על ניתוח תשובות 25 תלמידים לשאלות 50–59 | כלים: ChatGPT, Claude, Gemini, Google Bard, Copilot",
    bold=False, size=10, bg=COLORS["header_mid"], merge_to=9)
hdr(ws1, 3, 1, "ממוצע כיתתי כללי: 3.08/5 | רמה שלטת: מתקדם/ת (68% מהתלמידים) | נתוני הסימולציות 50–59",
    bold=False, size=10, bg="37474F", color="ECEFF1", merge_to=9)

ws1.row_dimensions[1].height = 35
ws1.row_dimensions[2].height = 22
ws1.row_dimensions[3].height = 20

# Column widths
col_widths = [28, 8, 28, 28, 28, 28, 28, 20, 30]
for i, w in enumerate(col_widths, 1):
    ws1.column_dimensions[get_column_letter(i)].width = w

# Table header row 4
row = 4
headers = ["ממד הערכה", "קוד", "רמה 1\nבתחילת הדרך", "רמה 2\nמתפתח/ת",
           "רמה 3\nמתקדם/ת", "רמה 4\nמיומן/ת", "רמה 5\nמומחה/ית",
           "ממוצע\nבכיתה", "שאלות המחוון"]
bg_hdrs = [COLORS["header_mid"]] * 2 + [COLORS["level1"], COLORS["level2"], COLORS["level3"],
            COLORS["level4"], COLORS["level5"], COLORS["header_mid"], COLORS["header_mid"]]
txt_colors = ["FFFFFF"] * 2 + ["B71C1C", "E65100", "827717", "1B5E20", "01579B",
               "FFFFFF", "FFFFFF"]

for c_idx, (h, bg, tc) in enumerate(zip(headers, bg_hdrs, txt_colors), 1):
    hdr(ws1, row, c_idx, h, bold=True, size=10, bg=bg, color=tc)
ws1.row_dimensions[row].height = 36

# ===== RUBRIC CONTENT PER DIMENSION =====
rubric_content = {
    "motivation": {
        "he": "א. מוטיבציה\nורלוונטיות",
        "avg": 3.10,
        "l1": "אינו מוצא קשר בין הלמידה לחייו. לא מעוניין במשימות מאתגרות. פועל רק מדחף חיצוני.",
        "l2": "מזדהה לעיתים עם החומר. מוכן לאתגר בתנאי שיש תמריץ מיידי (ציון).",
        "l3": "מתעניין בחלק מהנושאים. מגלה מוכנות לאתגר כשהחומר נוגע אליו.",
        "l4": "רואה רלוונטיות ברוב הנושאים. משקיע מעבר לדרישות המינימום.",
        "l5": "מחובר עמוקות ללמידה; רואה כל נושא כהזדמנות. יוזם למידה עצמאית מרצון."
    },
    "growth_mindset": {
        "he": "ב. תודעת\nצמיחה",
        "avg": 2.76,
        "l1": "תופס יכולת כקבועה. מתייאש בפני כישלון. נוטה לוותר מהר.",
        "l2": "מאמין בשיפור תיאורטי אך מתקשה ליישם. קשה לו לראות בטעות הזדמנות.",
        "l3": "מסוגל להתמיד כשהקושי אינו גדול מדי. מבין שמאמץ עוזר.",
        "l4": "רואה בקשיים אתגרים. מתאושש מכישלונות ולומד מהם.",
        "l5": "חי לפי עקרון הצמיחה; כישלון הוא מנוף. מחפש אתגרים באופן יזום."
    },
    "initiative": {
        "he": "ג. יוזמה\nואחריות",
        "avg": 3.40,
        "l1": "ממתין להוראות בלבד. אינו לוקח אחריות. תלוי לחלוטין במורה.",
        "l2": "לפעמים מציע רעיונות אך נסוג. קצת אחריות בלחץ חיצוני.",
        "l3": "לוקח אחריות בסיסית. פועל עצמאית במשימות שגרתיות.",
        "l4": "יוזם פעולות ומציע שינויים. לוקח אחריות גם כשדברים אינם הולכים כשורה.",
        "l5": "מוביל, מחדש ומשפיע על הסביבה. אחריות עמוקה גם מעבר לתחום האישי."
    },
    "self_regulation": {
        "he": "ד. ויסות\nעצמי",
        "avg": 3.18,
        "l1": "אינו מתכנן. נסחף בלחצים. מתקשה להמשיך תחת תסכול.",
        "l2": "מתכנן לפעמים אך לא ממשיך. ויסות רגשי חלקי.",
        "l3": "מתכנן משימות בסיסיות. מצליח להמשיך לרוב גם תחת לחץ.",
        "l4": "מנהל זמן ביעילות. מתמודד טוב עם תסכול ושומר על יציבות.",
        "l5": "ויסות עצמי מלא: תכנון אסטרטגי, שליטה רגשית, גמישות תחת לחץ."
    },
    "self_awareness": {
        "he": "ה. מודעות\nעצמית",
        "avg": 3.04,
        "l1": "אינו מכיר חוזקות/חולשות. לא מפיק לקחים מניסיונות.",
        "l2": "מודע חלקית לחוזקות. לפעמים חושב על כישלונות אך לא מסיק מסקנות.",
        "l3": "מכיר את עצמו בתחומים מרכזיים. מפיק לקחים בסיסיים.",
        "l4": "מודע לדפוסים אישיים. מפיק לקחים ומיישם בפועל.",
        "l5": "מודעות עצמית עמוקה; מנתח התנהגות, מזהה צמיחה ומתאים גישה."
    },
    "support": {
        "he": "ו. תמיכה\nוחוויות רגשיות",
        "avg": 2.96,
        "l1": "לא מרגיש תמיכה. אינו פונה לעזרה גם כשצריך.",
        "l2": "לפעמים מבקש עזרה אך חש שאין לו תמיכה אמיתית.",
        "l3": "מרגיש תמיכה בסיסית. פונה לעזרה כשקשה מאוד.",
        "l4": "מרגיש תמיכה ממקורות מגוונים. פונה לעזרה כחלק שוטף.",
        "l5": "רשת תמיכה חזקה. מנצל ומספק עזרה הדדית. חלק מקהילה לומדת."
    }
}

row = 5
for dim_code, content in rubric_content.items():
    ws1.row_dimensions[row].height = 60
    cell(ws1, row, 1, content["he"], bold=True, size=11, bg=COLORS["header_light"],
         align="center")
    cell(ws1, row, 2, dim_code[:1].upper(), bold=True, size=10, bg=COLORS["header_light"],
         align="center")
    for lvl in range(1, 6):
        key = f"l{lvl}"
        bg_color = level_colors[lvl]
        cell(ws1, row, 2 + lvl, content[key], size=9, bg=bg_color, align="right")
    avg_val = content["avg"]
    avg_level = level(avg_val)
    cell(ws1, row, 8, f"{avg_val:.2f}/5\n({level_names_he[avg_level]})",
         bold=True, size=10, bg=level_colors[avg_level], align="center",
         color="000000")
    q_text = dim_q_labels[dim_code]
    cell(ws1, row, 9, q_text, size=9, bg=COLORS["gray_light"], align="right")
    row += 1

# Overall row
ws1.row_dimensions[row].height = 30
hdr(ws1, row, 1, "ממוצע כיתתי כולל", bold=True, size=11, bg="37474F", color="ECEFF1")
cell(ws1, row, 2, "–", bold=True, size=11, bg="37474F", color="ECEFF1", align="center")
for lvl in range(1, 6):
    cell(ws1, row, 2 + lvl, "", bg="37474F")
cell(ws1, row, 8, "3.08/5\n(מתקדם/ת)", bold=True, size=11, bg=COLORS["level3"],
     align="center", color="333333")
cell(ws1, row, 9, "ממוצע כל 10 השאלות", size=10, bg=COLORS["gray_light"])
row += 1

# Notes below rubric
row += 1
hdr(ws1, row, 1, "📋 הערות פרשניות וממצאי מפתח מהנתונים", bold=True, size=12,
    bg="4A148C", color="FFFFFF", merge_to=9)
ws1.row_dimensions[row].height = 24
row += 1

notes = [
    ("ממד חזק ביותר: יוזמה ואחריות (ג)",
     "ממוצע 3.40/5 – 60% מהתלמידים ברמה 4 ומעלה. ✅ חוזקה של הכיתה.",
     COLORS["level4"]),
    ("ממד חלש ביותר: תודעת צמיחה (ב)",
     "ממוצע 2.76/5 – 40% ברמות 1-2. ⚠ דורש התערבות פדגוגית ממוקדת.",
     COLORS["level1"]),
    ("דפוס שכיח: מוטיבציה גבוהה + ויסות עצמי נמוך",
     "7 תלמידים (28%) – רוצים ללמוד אך לא מצליחים לנהל זמן ולווסת רגשות. נדרשת תמיכה בניהול למידה.",
     COLORS["level2"]),
    ("דפוס ייחודי: פערים גדולים בין ממדים",
     "17 תלמידים (68%) מציגים פערים ≥3 נקודות בין הממד החזק לחלש. מצביע על פרופיל שאינו מאוזן.",
     COLORS["level2"]),
    ("שימוש ב-AI ופעלנות",
     "משתמשי AI: 3.18/5 | לא-משתמשים: 2.58/5. קשר חיובי, אך לא בהכרח סיבתי.",
     COLORS["level3"]),
    ("כיתה י' בולטת לחיוב",
     "ממוצע 3.22/5 – הגבוה ביותר מבין כל הכיתות. מצב ויסות עצמי טוב (3.55).",
     COLORS["level4"]),
    ("כיתות יא'-יב': אתגרי ויסות עצמי",
     "שתי הכיתות הגבוהות מציגות ויסות עצמי ומודעות עצמית ירודים יחסית לגיל.",
     COLORS["level2"]),
]

for title, body, bg in notes:
    ws1.row_dimensions[row].height = 40
    cell(ws1, row, 1, title, bold=True, size=10, bg=bg, align="right")
    c = ws1.cell(row=row, column=2, value=body)
    c.font = Font(size=9, name="David")
    c.fill = PatternFill("solid", fgColor=bg)
    c.alignment = Alignment(horizontal="right", vertical="center", wrap_text=True, reading_order=2)
    c.border = make_border()
    ws1.merge_cells(start_row=row, start_column=2, end_row=row, end_column=9)
    row += 1

# ===================================================================
# SHEET 2: פרופיל תלמידים
# ===================================================================
ws2 = wb.create_sheet("פרופיל תלמידים")
ws2.sheet_view.rightToLeft = True
ws2.sheet_view.showGridLines = False

hdr(ws2, 1, 1, "פרופיל תלמידים – ציוני ממדים ורמות", bold=True, size=14,
    bg=COLORS["header_dark"], merge_to=13)
ws2.row_dimensions[1].height = 30

col_w2 = [18, 5, 10, 12, 12, 12, 12, 12, 12, 10, 12, 30, 20]
for i, w in enumerate(col_w2, 1):
    ws2.column_dimensions[get_column_letter(i)].width = w

row = 2
headers2 = ["שם", "כיתה", "כלי AI", "מוטיבציה", "תודעת\nצמיחה", "יוזמה\nואחריות",
            "ויסות\nעצמי", "מודעות\nעצמית", "תמיכה", "ממוצע\nכולל", "רמה כוללת",
            "ממד חזק ביותר", "דפוס מזוהה"]
bg2 = [COLORS["header_mid"]] * 13

for c_i, h in enumerate(headers2, 1):
    hdr(ws2, row, c_i, h, bold=True, size=9, bg=COLORS["header_mid"])
ws2.row_dimensions[row].height = 32
row += 1

sorted_df = df.sort_values("overall", ascending=False)

for _, r in sorted_df.iterrows():
    ws2.row_dimensions[row].height = 22
    overall_lvl = r["overall_level"]
    row_bg = level_colors[overall_lvl]

    cell(ws2, row, 1, r["שם התלמיד/ה"], bold=True, size=9, bg=row_bg)
    cell(ws2, row, 2, r["כיתה"], size=9, bg=row_bg, align="center")
    ai_tool = r.get("כלי AI בשימוש", "–")
    cell(ws2, row, 3, str(ai_tool), size=8, bg=row_bg)

    dim_scores = {}
    for col_idx, dim in enumerate(dim_codes, 4):
        val = r[f"dim_{dim}"]
        dim_scores[dim] = val
        lvl = level(val)
        cell(ws2, row, col_idx, f"{val:.1f}", size=9, bg=level_colors[lvl], align="center")

    cell(ws2, row, 10, f"{r['overall']:.1f}", bold=True, size=10, bg=row_bg, align="center")
    cell(ws2, row, 11, level_names_he[overall_lvl], size=9, bg=row_bg, align="center")

    strongest_dim = max(dim_codes, key=lambda d: r[f"dim_{d}"])
    cell(ws2, row, 12, dim_names_he[strongest_dim], size=8, bg=row_bg)

    # Identify pattern
    patterns = []
    if r["dim_motivation"] >= 3.5 and r["dim_self_regulation"] <= 2.5:
        patterns.append("מוטיבציה גבוהה+ויסות נמוך")
    if r["dim_growth_mindset"] >= 3.5 and r["dim_initiative"] <= 2.0:
        patterns.append("צמיחה גבוהה+יוזמה נמוכה")
    if r["dim_self_awareness"] <= 2.5 and r["dim_support"] >= 4.0:
        patterns.append("מודעות נמוכה+תמיכה גבוהה")
    max_d = max(dim_codes, key=lambda d: r[f"dim_{d}"])
    min_d = min(dim_codes, key=lambda d: r[f"dim_{d}"])
    gap = r[f"dim_{max_d}"] - r[f"dim_{min_d}"]
    if gap >= 3:
        patterns.append(f"פער גבוה ({gap:.0f} נק')")
    if not patterns:
        patterns.append("פרופיל מאוזן יחסית")

    cell(ws2, row, 13, " | ".join(patterns), size=8, bg=row_bg)
    row += 1

# Add legend
row += 1
hdr(ws2, row, 1, "מקרא רמות", bold=True, size=10, bg=COLORS["header_mid"], merge_to=3)
row += 1
for lvl, name in level_names_he.items():
    cell(ws2, row, 1, f"רמה {lvl}: {name}", bold=True, size=9, bg=level_colors[lvl])
    ws2.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
    row += 1

# ===================================================================
# SHEET 3: המלצות ושינויים
# ===================================================================
ws3 = wb.create_sheet("המלצות ושינויים")
ws3.sheet_view.rightToLeft = True
ws3.sheet_view.showGridLines = False

ws3.column_dimensions["A"].width = 30
ws3.column_dimensions["B"].width = 50
ws3.column_dimensions["C"].width = 50
ws3.column_dimensions["D"].width = 30
ws3.column_dimensions["E"].width = 20

hdr(ws3, 1, 1, "המלצות, הערות ושינויים מוצעים למחוון", bold=True, size=14,
    bg=COLORS["header_dark"], merge_to=5)
hdr(ws3, 2, 1, "מבוסס על ניתוח דפוסי שימוש ב-AI וממצאי שאלות 50–59",
    bold=False, size=10, bg=COLORS["header_mid"], merge_to=5)
ws3.row_dimensions[1].height = 30
ws3.row_dimensions[2].height = 20

row = 3
hdr(ws3, row, 1, "תחום / ממד", bold=True, size=10, bg="4A148C", color="FFFFFF")
hdr(ws3, row, 2, "ממצא מהנתונים", bold=True, size=10, bg="4A148C", color="FFFFFF")
hdr(ws3, row, 3, "הצעה לשינוי / המלצה פדגוגית", bold=True, size=10, bg="4A148C", color="FFFFFF")
hdr(ws3, row, 4, "עדיפות", bold=True, size=10, bg="4A148C", color="FFFFFF")
hdr(ws3, row, 5, "אחראי מוצע", bold=True, size=10, bg="4A148C", color="FFFFFF")
ws3.row_dimensions[row].height = 24
row += 1

recommendations = [
    # (area, finding, recommendation, priority, responsible)
    (
        "ב. תודעת צמיחה\n(הממד החלש ביותר)",
        "ממוצע 2.76/5 – החלש מכל הממדים.\n40% ברמות 1-2.\nהתלמידים אינם רואים בקשי הזדמנות ללמוד.",
        "✅ הוסף לשיעורים תרגילים ספציפיים לתודעת צמיחה:\n"
        "• שיתוף סיפורי כישלון→הצלחה\n"
        "• תגמול על מאמץ ולא רק על תוצאה\n"
        "• שינוי שפה: 'עדיין לא' במקום 'לא יכול'\n"
        "• יומן תהליך אישי שבועי",
        "🔴 גבוהה", "מחנך/ת כיתה"
    ),
    (
        "דפוס א: מוטיבציה גבוהה\n+ ויסות עצמי נמוך\n(7 תלמידים = 28%)",
        "7 תלמידים רוצים ללמוד (מוטיבציה ≥3.5)\nאך לא מצליחים לנהל זמן ולווסת עצמם\n(ויסות ≤2.5).",
        "✅ תכנית התערבות ממוקדת:\n"
        "• סדנאות ניהול זמן ותכנון שבועי\n"
        "• כלים לוויסות רגשי (מיינדפולנס בסיסי)\n"
        "• שימוש ב-AI כעוזר תכנון (לא כתחליף חשיבה)\n"
        "• חונכות עמיתים עם תלמידים מויסתים",
        "🔴 גבוהה", "יועץ/ת חינוכי/ת"
    ),
    (
        "דפוס ו: פערים גדולים\nbetween ממדים\n(17 תלמידים = 68%)",
        "רוב התלמידים (68%) מציגים פערים ≥3\nנקודות בין הממד החזק לחלש.\nהפרופיל אינו מאוזן.",
        "✅ בנה תכנית פיתוח אישית לכל תלמיד:\n"
        "• זיהוי ממד חלש ספציפי לכל תלמיד\n"
        "• משימות מותאמות אישית לחיזוק הממד החלש\n"
        "• שיחת מעקב אישית מדי חודש\n"
        "• עדכון מחוון אישי ברמה חצי-שנתית",
        "🟠 בינונית-גבוהה", "מחנך/ת + יועץ/ת"
    ),
    (
        "ג. יוזמה ואחריות\n(הממד החזק ביותר)",
        "ממוצע 3.40/5 – הגבוה מכל הממדים.\n60% ברמות 4-5.\nחוזקה ברורה של הכיתה.",
        "✅ נצל את היוזמה כמנוף:\n"
        "• העצם תלמידים לנהל פרויקטים עצמאיים\n"
        "• צור מסגרות של 'שגרירי למידה' בכיתה\n"
        "• חבר יוזמה לתחומי ממד חלש (צמיחה, ויסות)\n"
        "• הרחב את שאלת Q54 לשני מדדים נפרדים",
        "🟡 בינונית", "מורה מקצועי/ת"
    ),
    (
        "שימוש ב-AI\nושאלת אוריינות",
        "משתמשי AI: 3.18/5 | לא-משתמשים: 2.58/5.\nקשר חיובי בין שימוש ב-AI לפעלנות.\nאך 4 תלמידים לא משתמשים כלל.",
        "✅ המלצות לאוריינות AI:\n"
        "• הוסף שאלות על איכות שימוש ב-AI (לא רק תדירות)\n"
        "• שאל: 'כיצד AI עוזר לחשיבה שלך?'\n"
        "• לא-משתמשים: בדוק חסמים ותמוך בחשיפה\n"
        "• הוסף ממד 'אוריינות AI' ייעודי למחוון",
        "🟠 בינונית-גבוהה", "רכז/ת טכנולוגיה"
    ),
    (
        "ו. תמיכה וחוויות רגשיות\n(ממד שני בחולשה)",
        "ממוצע 2.96/5.\n40% ברמות 1-2.\nתלמידים אינם פונים לעזרה ביעילות.",
        "✅ חיזוק מערך התמיכה:\n"
        "• צור שגרה של 'שעת שאלות' בכיתה\n"
        "• הכשר תלמידים כ'חברי עזרה' הדדיים\n"
        "• בנה ערוצי תקשורת מגוונים (מורה, עמיתים, AI)\n"
        "• שקול הוספת מדד 'נגישות לעזרה' בשאלון",
        "🟡 בינונית", "יועץ/ת חינוכי/ת"
    ),
    (
        "ה. מודעות עצמית\nלפי כיתה",
        "כיתה יב': ממוצע 2.30 – הנמוך ביותר.\nכיתה י': ממוצע 3.25 – הגבוה ביותר.\nירידה במודעות עצמית עם עליית הכיתה.",
        "✅ שים לב לפרדוקס כיתות גבוהות:\n"
        "• ייתכן לחץ בגרות גורם לפחות הרהור עצמי\n"
        "• הוסף בכיתות יא'-יב' פורטפוליו רפלקטיבי\n"
        "• שלב שאלות מודעות עצמית בפגישות הכנה לבגרות\n"
        "• חקור האם AI מפחית הרהור עצמי אצל תלמידים בוגרים",
        "🟠 בינונית-גבוהה", "מחנך/ת יא'-יב'"
    ),
    (
        "שאלון Q50-Q59:\nהצעות לשיפור כלי המדידה",
        "שאלון 5-נקודות עם 10 שאלות.\nממד יוזמה (Q54) מכיל שאלה אחת בלבד.\nאין שאלה ישירה על אוריינות AI.",
        "✅ שינויים מוצעים בשאלון:\n"
        "• פצל Q54 לשתי שאלות: יוזמה + אחריות\n"
        "• הוסף Q60: 'עד כמה אתה משתמש ב-AI בצורה ביקורתית?'\n"
        "• הוסף Q61: 'עד כמה AI עוזר לך לחשוב – לא לחשוב במקומך?'\n"
        "• שקול פורמט פתוח לשאלה אחת לפחות\n"
        "• הוסף מדד השוואה (לפני/אחרי יחידת לימוד)",
        "🔴 גבוהה", "מפתח/ת תוכנית"
    ),
    (
        "תלמידים הזקוקים\nלתמיכה מיידית",
        "3 תלמידים עם ממוצע ≤2.5:\n• מיכל ברק (2.3)\n• איתי פרידמן (2.5)\n• נויה רוזנברג (2.5)\nכולם חלשים בוויסות עצמי.",
        "✅ התערבות מיידית:\n"
        "• שיחת בירור אישית עם כל אחד מהשלושה\n"
        "• תכנית ויסות עצמי אישית (8 שבועות)\n"
        "• בדוק גורמים חיצוניים (קשיי בית, חרדה)\n"
        "• קשר עם הורים אם נדרש",
        "🔴 דחופה", "מחנך/ת + יועץ/ת"
    ),
    (
        "שינוי מבני מוצע\nבמחוון הסימולציות",
        "המחוון הנוכחי: 6 ממדים, 5 רמות.\nחסר ממד AI ספציפי.\nאין התייחסות לאיכות שימוש.",
        "✅ הצעה למחוון 7 ממדים:\n"
        "א-ו: ממדים קיימים (שמור)\n"
        "ז. אוריינות AI: שימוש ביקורתי, הבנת מגבלות,\n"
        "   שילוב AI בתהליך חשיבה ולא כתחליף\n"
        "• רמות: 1=שימוש עיוור | 3=שימוש מודע | 5=שותפות ביקורתית",
        "🟠 בינונית-גבוהה", "צוות חינוכי רחב"
    ),
]

priority_colors = {
    "🔴 גבוהה":           "FFCDD2",
    "🔴 דחופה":           "FF8A80",
    "🟠 בינונית-גבוהה":   "FFE0B2",
    "🟡 בינונית":         "FFF9C4",
}

for area, finding, rec, priority, responsible in recommendations:
    ws3.row_dimensions[row].height = 80
    p_color = priority_colors.get(priority, COLORS["white"])
    cell(ws3, row, 1, area, bold=True, size=10, bg=COLORS["header_light"])
    cell(ws3, row, 2, finding, size=9, bg=COLORS["gray_light"])
    cell(ws3, row, 3, rec, size=9, bg="E8F5E9")
    cell(ws3, row, 4, priority, bold=True, size=10, bg=p_color, align="center")
    cell(ws3, row, 5, responsible, size=9, bg=COLORS["gray_light"], align="center")
    row += 1

# ===================================================================
# SHEET 4: דפוסי שימוש ב-AI
# ===================================================================
ws4 = wb.create_sheet("דפוסי שימוש ב-AI")
ws4.sheet_view.rightToLeft = True
ws4.sheet_view.showGridLines = False

ws4.column_dimensions["A"].width = 18
ws4.column_dimensions["B"].width = 14
ws4.column_dimensions["C"].width = 18
ws4.column_dimensions["D"].width = 16
ws4.column_dimensions["E"].width = 22
ws4.column_dimensions["F"].width = 14
ws4.column_dimensions["G"].width = 22
ws4.column_dimensions["H"].width = 12
ws4.column_dimensions["I"].width = 12
ws4.column_dimensions["J"].width = 14

hdr(ws4, 1, 1, "דפוסי שימוש ב-AI – ניתוח מלא", bold=True, size=14,
    bg=COLORS["header_dark"], merge_to=10)
ws4.row_dimensions[1].height = 30

row = 2
headers4 = ["שם", "כיתה", "כלי AI", "תדירות", "מטרת שימוש",
            "שעות/שבוע", "רמת תועלת", "חששות", "ציון לפני AI", "ציון אחרי AI"]
for c_i, h in enumerate(headers4, 1):
    hdr(ws4, row, c_i, h, bold=True, size=9, bg=COLORS["header_mid"])
ws4.row_dimensions[row].height = 24
row += 3 + 1

# Sort by AI tool
sorted_ai = df_ai.sort_values("כלי AI בשימוש")
row = 3
for _, r in sorted_ai.iterrows():
    ws4.row_dimensions[row].height = 20
    tool = str(r.get("כלי AI בשימוש", "–"))

    # Color by tool
    tool_colors = {
        "ChatGPT":     "E3F2FD",
        "Claude":      "F3E5F5",
        "Gemini":      "E8F5E9",
        "Google Bard": "FFF8E1",
        "Copilot":     "FCE4EC",
        "לא משתמש/ת": "F5F5F5",
    }
    bg = tool_colors.get(tool, COLORS["white"])

    score_before = r.get("ציון עצמי לפני AI (0-100)", 0)
    score_after = r.get("ציון עצמי אחרי AI (0-100)", 0)
    delta = score_after - score_before
    delta_color = COLORS["level4"] if delta > 0 else (COLORS["level1"] if delta < 0 else COLORS["white"])

    cell(ws4, row, 1, r.get("שם התלמיד/ה", ""), bold=True, size=9, bg=bg)
    cell(ws4, row, 2, r.get("כיתה", ""), size=9, bg=bg, align="center")
    cell(ws4, row, 3, tool, size=9, bg=bg)
    cell(ws4, row, 4, str(r.get("תדירות שימוש", "")), size=9, bg=bg)
    cell(ws4, row, 5, str(r.get("מטרת השימוש העיקרית", "")), size=9, bg=bg)
    cell(ws4, row, 6, str(r.get("שעות שימוש שבועיות", "")), size=9, bg=bg, align="center")
    cell(ws4, row, 7, str(r.get("רמת תועלת", "")), size=9, bg=bg)
    cell(ws4, row, 8, str(r.get("חששות", "")), size=9, bg=bg)
    cell(ws4, row, 9, str(score_before), size=9, bg=bg, align="center")
    after_txt = f"{score_after} ({'↑' if delta > 0 else '↓' if delta < 0 else '→'}{abs(delta)})"
    cell(ws4, row, 10, after_txt, size=9, bg=delta_color, align="center")
    row += 1

# Summary stats
row += 1
hdr(ws4, row, 1, "סיכום שימוש לפי כלי AI", bold=True, size=11,
    bg="37474F", color="FFFFFF", merge_to=10)
row += 1
hdr(ws4, row, 1, "כלי AI", size=10, bg=COLORS["header_mid"])
hdr(ws4, row, 2, "מספר\nתלמידים", size=10, bg=COLORS["header_mid"])
hdr(ws4, row, 3, "ממוצע שעות/שבוע", size=10, bg=COLORS["header_mid"])
hdr(ws4, row, 4, "ממוצע שינוי בציון", size=10, bg=COLORS["header_mid"])
hdr(ws4, row, 5, "חשש נפוץ", size=10, bg=COLORS["header_mid"])
ws4.row_dimensions[row].height = 28
row += 1

for tool in sorted(df_ai["כלי AI בשימוש"].unique()):
    sub = df_ai[df_ai["כלי AI בשימוש"] == tool]
    avg_hrs = sub["שעות שימוש שבועיות"].mean() if "שעות שימוש שבועיות" in sub else 0
    delta_avg = (sub["ציון עצמי אחרי AI (0-100)"] - sub["ציון עצמי לפני AI (0-100)"]).mean()
    top_concern = sub["חששות"].mode()[0] if len(sub) > 0 and "חששות" in sub else "–"

    tool_colors = {
        "ChatGPT": "E3F2FD", "Claude": "F3E5F5", "Gemini": "E8F5E9",
        "Google Bard": "FFF8E1", "Copilot": "FCE4EC", "לא משתמש/ת": "F5F5F5",
    }
    bg = tool_colors.get(tool, COLORS["white"])
    ws4.row_dimensions[row].height = 20
    cell(ws4, row, 1, tool, bold=True, size=9, bg=bg)
    cell(ws4, row, 2, str(len(sub)), size=9, bg=bg, align="center")
    cell(ws4, row, 3, f"{avg_hrs:.1f}", size=9, bg=bg, align="center")
    delta_bg = COLORS["level4"] if delta_avg > 0 else (COLORS["level1"] if delta_avg < 0 else COLORS["white"])
    cell(ws4, row, 4, f"{'↑' if delta_avg > 0 else '↓' if delta_avg < 0 else '→'}{abs(delta_avg):.1f}",
         size=9, bg=delta_bg, align="center")
    cell(ws4, row, 5, top_concern, size=9, bg=bg)
    row += 1

# ===================================================================
# SHEET 5: מחוון אוריינות AI מוצע
# ===================================================================
ws5 = wb.create_sheet("מחוון אוריינות AI מוצע")
ws5.sheet_view.rightToLeft = True
ws5.sheet_view.showGridLines = False

ws5.column_dimensions["A"].width = 30
ws5.column_dimensions["B"].width = 30
ws5.column_dimensions["C"].width = 30
ws5.column_dimensions["D"].width = 30
ws5.column_dimensions["E"].width = 30

hdr(ws5, 1, 1, "מחוון מוצע: ממד אוריינות בינה מלאכותית (ממד ז' חדש)", bold=True, size=14,
    bg="1A237E", merge_to=5)
hdr(ws5, 2, 1, "הצעה להרחבת המחוון הקיים (6 ממדים) בממד שביעי ייעודי לאוריינות AI",
    bold=False, size=10, bg="1565C0", merge_to=5)
ws5.row_dimensions[1].height = 30
ws5.row_dimensions[2].height = 20

col_widths5 = [30, 30, 30, 30, 30]
for i, w in enumerate(col_widths5, 1):
    ws5.column_dimensions[get_column_letter(i)].width = w

row = 3
sub_dims = [
    "ז1. שימוש ביקורתי ב-AI",
    "ז2. הבנת מגבלות AI",
    "ז3. שמירה על חשיבה עצמאית",
    "ז4. הערכת מידע מ-AI",
    "ז5. שיתוף פעולה אתי עם AI",
]

for sd in sub_dims:
    hdr(ws5, row, 1, sd, bold=True, size=10, bg="4527A0", color="EDE7F6", merge_to=5)
    ws5.row_dimensions[row].height = 24
    row += 1

    level_descs = {
        "ז1. שימוש ביקורתי ב-AI": {
            1: "מקבל כל תשובה של AI כאמת. לא מעלה שאלות על תוכן.",
            2: "לפעמים מפקפק אך בעיקר סומך על AI.",
            3: "בודק חלק מהתשובות. שואל AI שאלות ממוקדות.",
            4: "מאמת מידע בעקביות. מנסח שאלות ביקורתיות.",
            5: "שותפות ביקורתית: מנתח, בוחן, ומשפר תשובות AI."
        },
        "ז2. הבנת מגבלות AI": {
            1: "לא מכיר מגבלות AI. מאמין שה-AI תמיד צודק.",
            2: "יודע ש-AI טועה לפעמים אך אינו יודע מתי.",
            3: "מכיר מגבלות בסיסיות: hallucinations, תאריכים.",
            4: "מזהה תחומים שבהם AI חלש (שפה, הקשר, רגש).",
            5: "מומחה: מכיר ארכיטקטורת AI ומגבלותיה לעומק."
        },
        "ז3. שמירה על חשיבה עצמאית": {
            1: "מסתמך על AI לכל משימה. לא חושב לפני פנייה.",
            2: "מנסה לפתור לבד לפעמים אך נכנע מהר.",
            3: "מנסה לפתור לבד ואז בודק עם AI.",
            4: "AI הוא שלב בתהליך, לא תחליף. חשיבה עצמאית ברורה.",
            5: "AI כמגבר ולא כתחליף. מובל על ידי שאלות עצמיות."
        },
        "ז4. הערכת מידע מ-AI": {
            1: "מעתיק תשובות AI ישירות. אין סינון.",
            2: "מסנן מעט אך ברוב הפעמים מעתיק.",
            3: "עורך ומתאים תשובות AI לצרכיו.",
            4: "משלב מידע AI עם מקורות נוספים.",
            5: "יוצר מוצר מקורי המבוסס על AI + חשיבה עצמית."
        },
        "ז5. שיתוף פעולה אתי עם AI": {
            1: "לא מודע לשאלות אתיות (הגנת פרטיות, זכויות יוצרים).",
            2: "מודע לאתיקה בצורה כללית אך לא מיישם.",
            3: "נמנע משיתוף מידע אישי. מציין שימוש ב-AI.",
            4: "מיישם עקרונות אתיים בעקביות.",
            5: "מדגים אחריות אתית ומשפיע על עמיתים."
        }
    }

    if sd in level_descs:
        hdr(ws5, row, 1, "רמה 1\nבתחילת הדרך", bold=True, size=9, bg=COLORS["level1"], color="B71C1C")
        hdr(ws5, row, 2, "רמה 2\nמתפתח/ת", bold=True, size=9, bg=COLORS["level2"], color="E65100")
        hdr(ws5, row, 3, "רמה 3\nמתקדם/ת", bold=True, size=9, bg=COLORS["level3"], color="827717")
        hdr(ws5, row, 4, "רמה 4\nמיומן/ת", bold=True, size=9, bg=COLORS["level4"], color="1B5E20")
        hdr(ws5, row, 5, "רמה 5\nמומחה/ית", bold=True, size=9, bg=COLORS["level5"], color="01579B")
        ws5.row_dimensions[row].height = 28
        row += 1
        ws5.row_dimensions[row].height = 60
        for lvl_idx, lvl in enumerate(range(1, 6), 1):
            cell(ws5, row, lvl_idx, level_descs[sd][lvl], size=9, bg=level_colors[lvl])
        row += 1

    row += 1

# Add proposed new questions
row += 1
hdr(ws5, row, 1, "שאלות מוצעות לממד האוריינות ב-AI (Q60-Q65)", bold=True, size=12,
    bg="37474F", color="FFFFFF", merge_to=5)
row += 1
proposed_qs = [
    ("Q60", "עד כמה אתה בודק את נכונות המידע שקיבלת מ-AI לפני שאתה משתמש בו?"),
    ("Q61", "עד כמה אתה מנסה לחשוב על פתרון לבד לפני שאתה פונה ל-AI?"),
    ("Q62", "עד כמה אתה מבין את המגבלות וטעויות האפשריות של AI?"),
    ("Q63", "עד כמה אתה מציין בעבודות שלך כשהשתמשת ב-AI?"),
    ("Q64", "עד כמה AI עוזר לך לחשוב עמוק יותר – ולא לדלג על תהליך?"),
    ("Q65", "עד כמה אתה מרגיש שימוש ב-AI שיפר את ההבנה שלך – לא רק את הציון?"),
]

for q_code, q_text in proposed_qs:
    ws5.row_dimensions[row].height = 30
    cell(ws5, row, 1, q_code, bold=True, size=10, bg="E8EAF6", align="center")
    c = ws5.cell(row=row, column=2, value=q_text)
    c.font = Font(size=10, name="David")
    c.fill = PatternFill("solid", fgColor="E8EAF6")
    c.alignment = Alignment(horizontal="right", vertical="center", wrap_text=True, reading_order=2)
    c.border = make_border()
    ws5.merge_cells(start_row=row, start_column=2, end_row=row, end_column=5)
    row += 1

# ===================================================================
# SAVE
# ===================================================================
output_path = "/home/user/hany/מחוון_אוריינות_AI_מלא.xlsx"
wb.save(output_path)
print(f"✅ המחוון נשמר: {output_path}")
print(f"   גיליונות: {[ws.title for ws in wb.worksheets]}")
