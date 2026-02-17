"""
יצירת קובץ Excel לדוגמה עם נתוני שימוש בבינה מלאכותית של תלמידים
"""

import pandas as pd
import random

random.seed(42)

students = [
    "יעל כהן", "נועם לוי", "מאיה אברהם", "איתי פרידמן", "שירה דוד",
    "עומר גולן", "תמר מזרחי", "אדם ביטון", "נויה רוזנברג", "דניאל שלום",
    "רוני חן", "ליאור אלון", "עדי שרון", "אורי קצב", "מיכל ברק",
    "גיא סגל", "הילה וייס", "ניר אזולאי", "שקד תורג'מן", "אלה נחום",
    "עידו מלכה", "ליה גרינברג", "יונתן ספיר", "נעמה הלל", "רועי פינטו"
]

grades = ["ט", "י", "יא", "יב"]

ai_tools = ["ChatGPT", "Google Bard", "Claude", "Copilot", "Gemini", "לא משתמש/ת"]

usage_purposes = [
    "עזרה בשיעורי בית",
    "הבנת חומר לימודי",
    "כתיבת חיבורים",
    "פתרון בעיות במתמטיקה",
    "תרגום טקסטים",
    "הכנה למבחנים",
    "פרויקטים בתכנות",
    "מחקר לעבודות",
    "יצירת מצגות",
    "סיכום טקסטים"
]

frequency = ["כל יום", "כמה פעמים בשבוע", "פעם בשבוע", "כמה פעמים בחודש", "לעיתים רחוקות", "אף פעם"]

subjects = ["מתמטיקה", "אנגלית", "היסטוריה", "מדעים", "ספרות", "מדעי המחשב", "אזרחות", "פיזיקה"]

helpfulness = ["מאוד עוזר", "עוזר במידה מסוימת", "לא באמת עוזר", "מזיק ללמידה"]

concerns = [
    "חשש מהעתקה",
    "תלות יתר בטכנולוגיה",
    "חוסר הבנה עצמית",
    "פגיעה בחשיבה ביקורתית",
    "אין חששות",
    "חוסר דיוק במידע",
    "פגיעה ביצירתיות"
]

data = []
for student in students:
    grade = random.choice(grades)
    tool = random.choice(ai_tools)

    if tool == "לא משתמש/ת":
        freq = "אף פעם"
        purpose = "לא רלוונטי"
        subject = "לא רלוונטי"
        helpful = "לא רלוונטי"
        hours = 0
    else:
        freq = random.choice(frequency[:5])
        purpose = random.choice(usage_purposes)
        subject = random.choice(subjects)
        helpful = random.choice(helpfulness)
        if freq == "כל יום":
            hours = random.uniform(1, 4)
        elif freq == "כמה פעמים בשבוע":
            hours = random.uniform(0.5, 2.5)
        elif freq == "פעם בשבוע":
            hours = random.uniform(0.25, 1.5)
        else:
            hours = random.uniform(0.1, 0.75)

    concern = random.choice(concerns)

    self_score_before = random.randint(50, 95)
    if tool != "לא משתמש/ת" and helpful == "מאוד עוזר":
        self_score_after = min(100, self_score_before + random.randint(5, 20))
    elif tool != "לא משתמש/ת" and helpful == "מזיק ללמידה":
        self_score_after = max(30, self_score_before - random.randint(5, 15))
    else:
        self_score_after = self_score_before + random.randint(-5, 10)

    data.append({
        "שם התלמיד/ה": student,
        "כיתה": grade,
        "כלי AI בשימוש": tool,
        "תדירות שימוש": freq,
        "מטרת השימוש העיקרית": purpose,
        "מקצוע עיקרי לשימוש": subject,
        "שעות שימוש שבועיות": round(hours, 1),
        "רמת תועלת": helpful,
        "חששות": concern,
        "ציון עצמי לפני AI (0-100)": self_score_before,
        "ציון עצמי אחרי AI (0-100)": self_score_after
    })

df = pd.DataFrame(data)
df.to_excel("/home/user/hany/student_ai_usage.xlsx", index=False, engine="openpyxl")
print(f"Created Excel file with {len(df)} rows and {len(df.columns)} columns")
print(f"Columns: {list(df.columns)}")
print(df.head())
