"""
create_mock_data.py
יוצר קובץ mipuim_simulation.csv עם 1200 שורות ומבנה זהה לקובץ האמיתי.
כשהקובץ האמיתי יועלה – הוא יחליף אותו אוטומטית.
"""
import csv
import random
from datetime import datetime, timedelta

random.seed(42)

PILOTS = ["פיילוט א", "פיילוט ב", "פיילוט ג"]
MOSADOT = [
    ("1001", "בית ספר אורט נהריה"),
    ("1002", "בית ספר ריאלי חיפה"),
    ("1003", "תיכון הרצליה תל אביב"),
    ("1004", "בית ספר דה שליט רחובות"),
    ("1005", "תיכון עירוני ד׳ תל אביב"),
    ("1006", "בית ספר נעמן עכו"),
    ("1007", "תיכון מקיף גבעתיים"),
    ("1008", "בית ספר עמל פתח תקווה"),
]
KITOT = [
    ("ט1", "כיתה ט׳1"),
    ("ט2", "כיתה ט׳2"),
    ("י1",  "כיתה י׳1"),
    ("י2",  "כיתה י׳2"),
    ("יא1", "כיתה יא׳1"),
    ("יא2", "כיתה יא׳2"),
    ("יב1", "כיתה יב׳1"),
]
SHMOT_MISHPACHA = [
    "כהן","לוי","מזרחי","פרץ","ביטון","אברהם","פרידמן","שמש",
    "דהן","גרוס","ברק","אלון","שפירא","אלמוג","זיו","נחמיאס",
    "בן דוד","בן יוסף","שלום","כץ"
]
SHMOT_PRATI = [
    "נועה","יוסף","מיכל","אורי","שירה","דוד","רונית","אייל",
    "תמר","גיל","אדר","ליאור","יעל","עמית","רחל","ניר",
    "הדס","יונתן","גלית","שון"
]

# הגדרת פרופילי תשובות לכל דפוס שימוש
PROFILES = {
    "active_critical":  [4, 4, 5, 4, 5, 3, 4, 5, 4],  # משתמש פעיל-ביקורתי
    "passive_shortcut": [3, 2, 3, 2, 3, 4, 2, 3, 2],  # משתמש פסיבי-קיצורי
    "low_engagement":   [2, 1, 2, 2, 1, 3, 1, 2, 1],  # מעורבות נמוכה
}

PROFILE_DIST = ["active_critical"] * 35 + ["passive_shortcut"] * 40 + ["low_engagement"] * 25

start_date = datetime(2024, 9, 1)

rows = []
for i in range(1, 1201):
    mosad = random.choice(MOSADOT)
    kita  = random.choice(KITOT)
    pilot = random.choice(PILOTS)
    profile_name = random.choice(PROFILE_DIST)
    base = PROFILES[profile_name]
    q_vals = [max(1, min(5, v + random.randint(-1, 1))) for v in base]
    date_created = start_date + timedelta(days=random.randint(0, 180),
                                          hours=random.randint(8, 17),
                                          minutes=random.randint(0, 59))
    rows.append({
        "UID": f"STU{i:04d}",
        "פיילוט": pilot,
        "dateTimeCreate": date_created.strftime("%Y-%m-%d %H:%M:%S"),
        "shem_mishpacha_talmid": random.choice(SHMOT_MISHPACHA),
        "shem_male_talmid":     random.choice(SHMOT_PRATI),
        "mosad.code":  mosad[0],
        "mosad.teur":  mosad[1],
        "kita.code":   kita[0],
        "kita.teur":   kita[1],
        "Q1": q_vals[0],
        "Q2": q_vals[1],
        "Q3": q_vals[2],
        "Q4": q_vals[3],
        "Q5": q_vals[4],
        "Q6": q_vals[5],
        "Q7": q_vals[6],
        "Q8": q_vals[7],
        "Q9": q_vals[8],
    })

fieldnames = [
    "UID","פיילוט","dateTimeCreate",
    "shem_mishpacha_talmid","shem_male_talmid",
    "mosad.code","mosad.teur","kita.code","kita.teur",
    "Q1","Q2","Q3","Q4","Q5","Q6","Q7","Q8","Q9"
]

with open("/home/user/hany/mipuim_simulation.csv", "w", newline="", encoding="utf-8-sig") as f:
    writer = csv.DictWriter(f, fieldnames=fieldnames)
    writer.writeheader()
    writer.writerows(rows)

print(f"נוצר קובץ mipuim_simulation.csv עם {len(rows)} שורות.")
print("עמודות:", ", ".join(fieldnames))
