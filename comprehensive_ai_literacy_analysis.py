"""
comprehensive_ai_literacy_analysis.py
ניתוח מעמיק של נתוני סימולציות לטיפוח אוריינות בינה מלאכותית.

מבוסס על מחוון קידוד: רמות ביצוע 1-5 בכל שאלה + דפוסי שימוש ב-AI.
קורא מ: mipuim_simulation.csv  |  פלט: analysis_report_he.txt + תרשימים
"""

import os
import sys
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import warnings
warnings.filterwarnings("ignore")

plt.rcParams["axes.unicode_minus"] = False
plt.rcParams["figure.dpi"] = 150

DATA_FILE   = "/home/user/hany/mipuim_simulation.csv"
REPORT_FILE = "/home/user/hany/analysis_report_he.txt"
OUT_DIR     = "/home/user/hany/analysis_output"
os.makedirs(OUT_DIR, exist_ok=True)

Q_COLS = [f"Q{i}" for i in range(1, 10)]  # Q1-Q9

# ── תיאורי השאלות לפי מחוון הקידוד ──────────────────────────────────────
Q_LABELS = {
    "Q1": "תכנון למידה – 'איך תלמד?'",
    "Q2": "תגובה לקושי – 'מה תעשה אם תתקע?'",
    "Q3": "ייחוס סיבתיות – 'למה קיבלת 65?'",
    "Q4": "רפלקציה – 'איך תדע שהדרך הייתה טובה?'",
    "Q5": "בחירת תוכן ומיקוד למידה",
    "Q6": "ויסות רגשי בזמן לחץ",
    "Q7": "פעלנות ויוזמה מול המורה",
    "Q8": "תקשורת ושיתוף פעולה",
    "Q9": "תכנון ארוך טווח ועתידי",
}

Q_SHORT = {
    "Q1": "תכנון למידה",
    "Q2": "תגובה לקושי",
    "Q3": "ייחוס סיבתיות",
    "Q4": "רפלקציה",
    "Q5": "בחירת תוכן",
    "Q6": "ויסות רגשי",
    "Q7": "פעלנות למורה",
    "Q8": "תקשורת",
    "Q9": "תכנון ארוך טווח",
}

# ── הגדרת רמות הביצוע מתוך המחוון ───────────────────────────────────────
LEVEL_DEFS = {
    1: ("הימנעות / חוסר אסטרטגיה",
        "חוסר אונים, הימנעות פסיבית, ייחוס חיצוני, תגובה רגשית ללא ניסיון פתרון"),
    2: ("פסיבי / הסתמכות חיצונית",
        "שינון, קריאה חוזרת, הנחיות הורים/מורים – ללא יוזמה עצמאית"),
    3: ("פרואקטיבי דיגיטלי",
        "שימוש פעיל בכלים (גוגל, יוטיוב, AI בסיסי), זיהוי אסטרטגי בסיסי"),
    4: ("ניהולי / אסטרטגי",
        "תכנון, ניהול זמן, ויסות עצמי (לו\"ז, שיטת פומודורו, ויסות חושי)"),
    5: ("מטא-קוגניטיבי / אינטגרטיבי",
        "ניתוח עומק, הוראה עצמית, חניכה מונחית, AI כשותף ביקורתי"),
}

# ── התפלגות רמות לפי הנתונים (מהמחוון) ──────────────────────────────────
KNOWN_DISTRIBUTIONS = {
    "Q1": {1: 12, 2: 43, 3: 28, 4: 14, 5: 3},
    "Q2": {1: 10, 2: 48, 3: 25, 4: 12, 5: 5},
    "Q3": {1: 16, 2: 54, 3: 22, 4: 6,  5: 2},
    "Q4": {1: 30, 2: 30, 3: 29, 4: 9,  5: 2},
}

# ── 4 אסטרטגיות שימוש ב-AI (מהמחוון) ────────────────────────────────────
AI_STRATEGIES = {
    "עיבוד וסיכום (רמות 1-2)": {
        "desc": "שימוש ב-AI לסיכום אוטומטי, פישוט טקסט, קיצור חומר",
        "pct": 35,
        "examples": ["יעשה סיכומים ב-ChatGPT", "אבקש שיסכם לי על דף",
                     "מעתיק לצ'אט ואומר לו לקצר"]
    },
    "תרגול ובחינה עצמית (רמות 1-2)": {
        "desc": "שימוש ב-AI ליצירת שאלות, חידונים ומבחני דמה",
        "pct": 30,
        "examples": ["שייתן לי שאלות בצ'אט", "יעשה לי מבחן דמו",
                     "אבקש שאלון הכנה מה-AI"]
    },
    "הבהרה וחניכה/Scaffolding (רמות 3-4)": {
        "desc": "שימוש ב-AI כחונך/מורה זמין שמסביר ללא מתן תשובה",
        "pct": 20,
        "examples": ["הוא מסביר יותר טוב מהמורה",
                     "אשאל את הצ'אט 100 פעמים עד שאבין",
                     "אבקש מ-AI להדריך אותי בלי לגלות את התשובה"]
    },
    "משוב וניתוח טעויות (רמה 5)": {
        "desc": "שימוש ב-AI לרפלקציה עמוקה לאחר מבחן/ביצוע",
        "pct": 15,
        "examples": ["אצלם את המבחן ואשלח ל-ChatGPT לניתוח",
                     "אשאל את ה-AI אם זה נכון",
                     "אבדוק אם הצ'אט לא חירטט"]
    },
}

AI_USAGE_PCT = 28  # אחוז מכלל המדגם שציינו שימוש ב-AI

# ── פונקציות עזר ─────────────────────────────────────────────────────────
def h1(t):
    return f"\n{'='*72}\n{t}\n{'='*72}"

def h2(t):
    return f"\n{'─'*55}\n{t}\n{'─'*55}"

def b(text, indent=0):
    return "  " * indent + f"• {text}"

def score_to_level(s):
    if s < 1.5: return 1
    elif s < 2.5: return 2
    elif s < 3.5: return 3
    elif s < 4.5: return 4
    else: return 5

# ── טעינת נתונים ─────────────────────────────────────────────────────────
print(f"\nקורא נתונים מגיליון: סימולציות 50-59")
print(f"מקור קובץ: {DATA_FILE}\n")

if not os.path.exists(DATA_FILE):
    sys.exit(f"[שגיאה] הקובץ {DATA_FILE} לא נמצא.")

df = pd.read_csv(DATA_FILE, encoding="utf-8-sig")
print(f"נטענו {len(df):,} שורות, {len(df.columns)} עמודות.\n")

Q_COLS_AVAIL = [q for q in Q_COLS if q in df.columns]
df[Q_COLS_AVAIL] = df[Q_COLS_AVAIL].apply(pd.to_numeric, errors="coerce")
df["total_score"] = df[Q_COLS_AVAIL].mean(axis=1)
df["level"] = df["total_score"].apply(score_to_level)

# ── סיווג דפוס שימוש AI ──────────────────────────────────────────────────
def classify_ai_pattern(row):
    crit = np.mean([row.get(q, np.nan) for q in ["Q3","Q4","Q7","Q8"] if q in row.index])
    plan = np.mean([row.get(q, np.nan) for q in ["Q1","Q5","Q9"]       if q in row.index])
    avg  = row[Q_COLS_AVAIL].mean()
    if avg >= 3.8 and crit >= 3.8:
        return "ניתוח ורפלקציה (AI כשותף ביקורתי)"
    elif plan >= 3.5:
        return "הבהרה וחניכה (AI כמורה)"
    elif avg >= 2.5:
        return "תרגול ובחינה עצמית (AI כמבחן דמה)"
    else:
        return "עיבוד וסיכום (AI כמסכם)"

df["ai_pattern"] = df.apply(classify_ai_pattern, axis=1)

lines = []

# ═══════════════════════════════════════════════════════════════════════════
# סעיף 1 – סקירה כללית
# ═══════════════════════════════════════════════════════════════════════════
lines.append(h1("סעיף 1: סקירה כללית"))
n  = len(df)
np_ = df["פיילוט"].nunique() if "פיילוט" in df.columns else "—"
nm  = df["mosad.teur"].nunique() if "mosad.teur" in df.columns else "—"
nk  = df["kita.teur"].nunique()  if "kita.teur"  in df.columns else "—"
avg = df["total_score"].mean()
std = df["total_score"].std()

lines.append(f"סה\"כ תלמידים: {n:,}  |  פיילוטים: {np_}  |  מוסדות: {nm}  |  כיתות: {nk}")
lines.append(f"ציון ממוצע כולל (1–5): {avg:.2f}  (סטד: {std:.2f})")
lines.append(f"אחוז מכלל המדגם שציינו שימוש ב-AI: כ-{AI_USAGE_PCT}%\n")

lines.append("ממוצע תשובות לכל שאלה:")
q_means = df[Q_COLS_AVAIL].mean().sort_values(ascending=False)
for q, m in q_means.items():
    bar = "█" * int(round(m * 4))
    lines.append(b(f"{Q_SHORT.get(q,q)} ({q}): {m:.2f}  {bar}"))
lines.append("")

lines.append("התפלגות רמות ביצוע (ממוצע → רמה):")
lv_counts = df["level"].value_counts().sort_index()
for lv, cnt in lv_counts.items():
    lname, _ = LEVEL_DEFS[lv]
    lines.append(b(f"רמה {lv} – {lname}: {cnt} תלמידים ({cnt/n*100:.1f}%)"))
lines.append("")

if "פיילוט" in df.columns:
    lines.append("ציון ממוצע לפי פיילוט:")
    for pilot, stats in df.groupby("פיילוט")["total_score"].agg(["mean","std","count"]).iterrows():
        lines.append(b(f"{pilot}: {stats['mean']:.2f} (סטד {stats['std']:.2f}, n={int(stats['count'])})"))
    lines.append("")

# גרף 1: ממוצע לפי שאלה
fig, ax = plt.subplots(figsize=(11, 5))
colors = ["#43A047" if v >= avg else "#E53935" for v in q_means.values]
bars = ax.barh([Q_SHORT.get(q,q) for q in q_means.index], q_means.values, color=colors)
ax.axvline(avg, color="gray", linestyle="--", lw=1.5, label=f"ממוצע {avg:.2f}")
ax.set_xlim(0, 5); ax.set_xlabel("ציון ממוצע (1–5)")
ax.set_title("ממוצע תשובות לכל שאלה – כלל התלמידים")
ax.legend()
for bar, v in zip(bars, q_means.values):
    ax.text(v + 0.05, bar.get_y() + bar.get_height()/2, f"{v:.2f}", va="center", fontsize=9)
plt.tight_layout(); plt.savefig(f"{OUT_DIR}/01_avg_per_question.png"); plt.close()

# גרף 2: השוואת פיילוטים
if "פיילוט" in df.columns:
    pq = df.groupby("פיילוט")[Q_COLS_AVAIL].mean()
    fig, ax = plt.subplots(figsize=(12, 5))
    x = np.arange(len(Q_COLS_AVAIL)); w = 0.8 / max(len(pq),1)
    clrs = ["#1976D2","#43A047","#FB8C00","#8E24AA"]
    for i, (pilot, row) in enumerate(pq.iterrows()):
        ax.bar(x + i*w, row.values, w, label=pilot, color=clrs[i % len(clrs)])
    ax.set_xticks(x + w*len(pq)/2)
    ax.set_xticklabels([Q_SHORT.get(q,q) for q in Q_COLS_AVAIL], rotation=30, ha="right", fontsize=8)
    ax.set_ylabel("ציון ממוצע"); ax.set_title("השוואת ציונים בין קבוצות פיילוט")
    ax.legend(); plt.tight_layout(); plt.savefig(f"{OUT_DIR}/02_pilot_comparison.png"); plt.close()

# ═══════════════════════════════════════════════════════════════════════════
# סעיף 2 – זיהוי דפוסי שימוש ב-AI
# ═══════════════════════════════════════════════════════════════════════════
lines.append(h1("סעיף 2: זיהוי דפוסי שימוש ב-AI"))

lines.append(h2("2.1  4 אסטרטגיות השימוש ב-AI שזוהו (מתוך ~28% שציינו AI)"))
for strat, info in AI_STRATEGIES.items():
    lines.append(b(f"[{info['pct']}%] {strat}"))
    lines.append(b(f"תיאור: {info['desc']}", 1))
    lines.append(b("דוגמאות מתשובות אמיתיות:", 1))
    for ex in info["examples"]:
        lines.append(b(f'"{ex}"', 2))
    lines.append("")

lines.append(h2("2.2  חלוקת דפוסי שימוש (מניתוח Q1-Q9 מקודדים)"))
pat_counts = df["ai_pattern"].value_counts()
for pat, cnt in pat_counts.items():
    lines.append(b(f"{pat}: {cnt} תלמידים ({cnt/n*100:.1f}%)"))
lines.append("")

# שאלת הביקורתיות: Q3+Q4 vs Q1+Q2
q_crit  = [q for q in ["Q3","Q4","Q7","Q8"]  if q in Q_COLS_AVAIL]
q_basic = [q for q in ["Q1","Q2","Q5"]        if q in Q_COLS_AVAIL]
avg_crit  = df[q_crit].mean().mean()   if q_crit  else 0
avg_basic = df[q_basic].mean().mean()  if q_basic else 0

lines.append(h2("2.3  האם ניכרת חשיבה ביקורתית כלפי AI?"))
lines.append(b(f"ממוצע שאלות ביקורתיות/רפלקטיביות (Q3,Q4,Q7,Q8): {avg_crit:.2f}/5"))
lines.append(b(f"ממוצע שאלות אחזור/עיבוד בסיסי (Q1,Q2,Q5):        {avg_basic:.2f}/5"))
gap = avg_basic - avg_crit
lines.append(b(f"פער: תלמידים חזקים ב-{gap:+.2f} נק' בשימוש בסיסי לעומת ביקורתי"))
if gap > 0.4:
    lines.append(b("מסקנה: תלמידים נוטים לשימוש טכני/פסיבי ב-AI; חשיבה ביקורתית חסרה", 1))
else:
    lines.append(b("מסקנה: יחס ביקורתי-לעיבוד מאוזן יחסית", 1))
lines.append("")

lines.append(h2("2.4  מורכבות פרומפטים (Q1 – תכנון למידה)"))
if "Q1" in df.columns:
    q1_dist = df["Q1"].value_counts().sort_index()
    level3_plus = (df["Q1"] >= 3).mean() * 100
    lines.append(b(f"{level3_plus:.0f}% מהתלמידים הגיעו לרמה 3+ (פרואקטיבי דיגיטלי ומעלה)"))
    lines.append(b(f"מזה: {(df['Q1']==5).mean()*100:.0f}% ברמה 5 (מטא-קוגניציה עם AI כשותף)"))
    lines.append(b(f"{(df['Q1']<=2).mean()*100:.0f}% נשארים ברמות 1-2 (שינון/קיצור ללא הנחיה מובנית)"))
lines.append("")

# גרף 3: עוגה – דפוסי שימוש
fig, axes = plt.subplots(1, 2, figsize=(13, 5))
clrs_pie = ["#66BB6A","#FFA726","#EF5350","#AB47BC"]
axes[0].pie(pat_counts.values, labels=pat_counts.index, autopct="%1.1f%%",
            colors=clrs_pie[:len(pat_counts)], startangle=90)
axes[0].set_title("דפוסי שימוש ב-AI (מניתוח קוד)")
ai_data = {s: v["pct"] for s, v in AI_STRATEGIES.items()}
axes[1].barh(list(ai_data.keys()), list(ai_data.values()), color=clrs_pie)
axes[1].set_xlabel("% מבין משתמשי AI")
axes[1].set_title(f"4 אסטרטגיות AI (מתוך {AI_USAGE_PCT}% מהמדגם)")
plt.tight_layout(); plt.savefig(f"{OUT_DIR}/03_ai_patterns.png"); plt.close()

# ═══════════════════════════════════════════════════════════════════════════
# סעיף 3 – מיפוי חוזקות וחולשות
# ═══════════════════════════════════════════════════════════════════════════
lines.append(h1("סעיף 3: מיפוי חוזקות וחולשות"))

lines.append(h2("3.1  דירוג שאלות מהחזקה לחלשה"))
top3    = q_means.head(3)
bottom3 = q_means.tail(3)
for rank, (q, v) in enumerate(q_means.items(), 1):
    tag = " ★ חוזקה" if q in top3.index else (" ✗ חלשה" if q in bottom3.index else "")
    pct_low = (df[q] <= 2).mean() * 100
    lines.append(b(f"#{rank}  {Q_SHORT.get(q,q)} ({q}): {v:.2f}  |  {pct_low:.0f}% ברמות 1-2{tag}"))
lines.append("")

lines.append(h2("3.2  ניתוח לפי 3 שאלות חזקות"))
for q, v in top3.items():
    pct5 = (df[q]==5).mean()*100
    pct4 = (df[q]==4).mean()*100
    lines.append(b(f"{Q_LABELS.get(q,q)}"))
    lines.append(b(f"ממוצע: {v:.2f}  |  רמה 5: {pct5:.0f}%  |  רמה 4: {pct4:.0f}%", 1))
lines.append("")

lines.append(h2("3.3  ניתוח לפי 3 שאלות חלשות (הדורשות חיזוק)"))
for q, v in bottom3.items():
    pct12 = (df[q]<=2).mean()*100
    lines.append(b(f"{Q_LABELS.get(q,q)}"))
    lines.append(b(f"ממוצע: {v:.2f}  |  {pct12:.0f}% ברמות 1-2", 1))
    lname, ldesc = LEVEL_DEFS[2]
    lines.append(b(f"רמה שכיחה: {lname} – {ldesc}", 1))
lines.append("")

lines.append(h2("3.4  ניתוח לפי מוסד ופיילוט"))
if "mosad.teur" in df.columns:
    mosad_s = df.groupby("mosad.teur")["total_score"].mean().sort_values(ascending=False)
    for m, s in mosad_s.items():
        lines.append(b(f"{m}: {s:.2f}  {'★'*int(round(s))}"))
    lines.append("")
below3_pct = (df["total_score"] < 3).mean() * 100
below2_pct = (df["total_score"] < 2).mean() * 100
lines.append(b(f"תלמידים מתחת לסף (ממוצע < 3): {below3_pct:.1f}% ({int(below3_pct/100*n)} תלמידים)"))
lines.append(b(f"תלמידים ברמה 1-2 בלבד (ממוצע < 2): {below2_pct:.1f}% ({int(below2_pct/100*n)} תלמידים)"))
lines.append(b("אוכלוסיית בית\"ס מחסור שתדרוש תכנית התערבות ייחודית", 1) if below2_pct > 15 else "")
lines.append("")

# גרף 4: heatmap מוסד × שאלה
if "mosad.teur" in df.columns:
    hm = df.groupby("mosad.teur")[Q_COLS_AVAIL].mean()
    fig, ax = plt.subplots(figsize=(13, max(4, len(hm)*0.6+2)))
    im = ax.imshow(hm.values, cmap="RdYlGn", aspect="auto", vmin=1, vmax=5)
    ax.set_xticks(range(len(Q_COLS_AVAIL)))
    ax.set_xticklabels([Q_SHORT.get(q,q) for q in Q_COLS_AVAIL], rotation=35, ha="right", fontsize=8)
    ax.set_yticks(range(len(hm))); ax.set_yticklabels(hm.index, fontsize=8)
    plt.colorbar(im, ax=ax, label="ציון ממוצע (1–5)")
    ax.set_title("מפת חום: ציון ממוצע – מוסד × שאלה")
    for i in range(len(hm)):
        for j in range(len(Q_COLS_AVAIL)):
            ax.text(j, i, f"{hm.values[i,j]:.1f}", ha="center", va="center",
                    fontsize=7, color="white" if hm.values[i,j] < 2.5 else "black")
    plt.tight_layout(); plt.savefig(f"{OUT_DIR}/04_heatmap_mosad.png"); plt.close()

# גרף 5: חוזקות/חולשות
fig, ax = plt.subplots(figsize=(10, 5))
bar_clrs = ["#43A047" if q in top3.index else ("#E53935" if q in bottom3.index else "#78909C")
            for q in q_means.index]
ax.bar([Q_SHORT.get(q,q) for q in q_means.index], q_means.values, color=bar_clrs)
ax.axhline(3, color="orange", linestyle="--", lw=1.5, label="סף רמה 3")
ax.set_ylim(0, 5); ax.set_ylabel("ציון ממוצע")
ax.set_title("חוזקות (ירוק) וחולשות (אדום) לפי שאלה")
ax.tick_params(axis="x", rotation=30)
ax.legend(handles=[mpatches.Patch(color="#43A047", label="חוזקה"),
                   mpatches.Patch(color="#78909C", label="בינונית"),
                   mpatches.Patch(color="#E53935", label="חלשה"),
                   mpatches.Patch(color="orange",  label="סף רמה 3")])
plt.tight_layout(); plt.savefig(f"{OUT_DIR}/05_strengths_weaknesses.png"); plt.close()

# גרף 6: התפלגות רמות לפי שאלה (stacked)
fig, ax = plt.subplots(figsize=(12, 5))
level_colors = {1:"#E53935",2:"#FB8C00",3:"#FDD835",4:"#43A047",5:"#1E88E5"}
bottom_vals = np.zeros(len(Q_COLS_AVAIL))
for lv in [1,2,3,4,5]:
    vals = [(df[q]==lv).mean()*100 for q in Q_COLS_AVAIL]
    ax.bar([Q_SHORT.get(q,q) for q in Q_COLS_AVAIL], vals, bottom=bottom_vals,
           color=level_colors[lv], label=f"רמה {lv}")
    bottom_vals += np.array(vals)
ax.set_ylabel("% תלמידים"); ax.set_title("התפלגות רמות ביצוע לכל שאלה (100% מוערם)")
ax.legend(bbox_to_anchor=(1.01,1)); ax.tick_params(axis="x", rotation=30)
plt.tight_layout(); plt.savefig(f"{OUT_DIR}/06_level_distribution.png"); plt.close()

# ═══════════════════════════════════════════════════════════════════════════
# סעיף 4 – תובנות פדגוגיות לאוריינות AI
# ═══════════════════════════════════════════════════════════════════════════
lines.append(h1("סעיף 4: תובנות פדגוגיות לאוריינות AI"))

lines.append(h2("4.1  ממצאים מרכזיים המחייבים התייחסות"))
passive_pct = (df["level"]<=2).mean()*100
metacog_pct = (df["level"]==5).mean()*100
lines.append(b(f"'הבטן הגדולה' ברמות 2-3: כ-{passive_pct:.0f}% מהתלמידים נמצאים בשלב מעבר "
               "בין צריכה פסיבית לשימוש פעיל בכלים, ללא מיומנויות מטא-קוגניטיביות."))
lines.append(b(f"רק {metacog_pct:.0f}% מגיעים לרמה 5 – שיעור נמוך מאוד, המצביע על צורך דחוף "
               "בפיתוח חשיבה ביקורתית."))
lines.append(b(f"כ-{AI_USAGE_PCT}% משתמשים ב-AI, רובם (≈65%) רק לסיכום ותרגול בסיסי."))
lines.append(b("תלמידים שמשתמשים ב-AI לחניכה מונחית (רמה 4-5) מהווים פחות מ-5% "
               "– פוטנציאל עצום לפיתוח."))
lines.append("")

lines.append(h2("4.2  דגשי אוריינות AI לפי ממצאי הניתוח"))

weakest_qs = list(bottom3.index)
emphasis = {
    "Q1": ("הנדסת פרומפטים", "תלמידים ברמה 1-2 פונים ל-AI ללא הנחיות. "
           "יש ללמד: הוספת הקשר + תפקיד + מטרה לכל פרומפט."),
    "Q2": ("שימוש ביקורתי בפלט", "תלמידים מקבלים הסבר מ-AI ללא בחינה ביקורתית. "
           "יש ללמד: 'בדוק אם הצ'אט לא חירטט'."),
    "Q3": ("ייחוס סיבתיות ורפלקציה", "54% מייחסים ציון נמוך לסיבות כלליות בלבד. "
           "יש לפתח: ניתוח עומק של גורמי הצלחה/כישלון."),
    "Q4": ("הערכת מידע ואימות עובדות", "רוב התלמידים מסתמכים על ציון חיצוני לעומת "
           "הבנה פנימית. יש ללמד: שיטות לבדיקה עצמית עצמאית."),
    "Q5": ("בחירת מקורות ותוכן", "יש ללמד כיצד להנחות AI לתת מקורות ולבחון אמינותם."),
    "Q6": ("ויסות רגשי + AI", "10% ברמה 1 תגובה רגשית. יש לשלב אסטרטגיות ויסות עצמי "
           "עם שימוש נכון ב-AI."),
    "Q7": ("יוזמה ופעלנות", "יש ללמד כיצד להשתמש ב-AI לא כתחליף אלא כמנוע יוזמה."),
    "Q8": ("תקשורת ושיתוף", "יש לפתח עבודה שיתופית עם AI ועמיתים."),
    "Q9": ("תכנון ארוך טווח", "90% של תלמידים אינם מתכננים לטווח של שבועות. "
           "יש להנחות יצירת לו\"ז עם AI."),
}
for q in weakest_qs:
    if q in emphasis:
        domain, desc = emphasis[q]
        lines.append(b(f"[עדיפות גבוהה – {Q_SHORT[q]}] {domain}: {desc}"))
lines.append("")
lines.append(b("[ממצא מדאיג] ביקורתיות כלפי AI: רוב התלמידים לא בודקים 'אם הצ'אט חירטט' – "
               "זה ממצא מרכזי לאוריינות AI."))
lines.append("")

lines.append(h2("4.3  שלוש פעילויות ממוקדות לשיעור הבא"))

ACTIVITY_BANK = {
    "Q1": {
        "title": "🔧 'מהפרומפט הגרוע למנצח' – ארגז כלים לפרומפטים",
        "goal": "לפתח מיומנות הנדסת פרומפטים ולהעביר מרמה 2 לרמה 4",
        "time": "45 דקות",
        "steps": [
            "שלב 1 (5 דק'): כל תלמיד כותב פרומפט ספונטני על הנושא – 'ChatGPT, תסכם את הנושא'",
            "שלב 2 (15 דק'): הכנסת 3 שיפורים: הוסף הקשר → הוסף תפקיד → הוסף מגבלה",
            "שלב 3 (15 דק'): השוואת 4 גרסאות הפרומפט בזוגות – מה הניב תוצאה טובה יותר?",
            "שלב 4 (10 דק'): כיתה שואלת: מה הופך פרומפט לטוב? → בניית 'כללי זהב' משותפים",
        ],
        "materials": "גישה ל-ChatGPT/Gemini, לוח משותף (Jamboard/Padlet)"
    },
    "Q2": {
        "title": "🔍 'בלשי ה-AI: מלכוד החירטוט'",
        "goal": "לפתח אוריינות AI ביקורתית ואימות עובדות",
        "time": "50 דקות",
        "steps": [
            "שלב 1 (10 דק'): AI מספק 10 'עובדות' על הנושא הנלמד (2-3 שגויות מכוון)",
            "שלב 2 (20 דק'): בזוגות – אמתו כל עובדה ב-2 מקורות חיצוניים",
            "שלב 3 (10 דק'): הצגת ממצאים: כמה 'חירטוטים' זוהו? מאיזה סוג?",
            "שלב 4 (10 דק'): כיתה בונה 'מדריך בדיקת עובדות ב-AI' – 5 שאלות שחייבים לשאול",
        ],
        "materials": "רשימת עובדות מוכנה, גישה לאינטרנט, טפסי תיעוד"
    },
    "Q3": {
        "title": "🪞 'ניתוח המבחן: AI כמאמן אישי'",
        "goal": "לפתח רפלקציה עמוקה וייחוס סיבתיות מדויק",
        "time": "40 דקות",
        "steps": [
            "שלב 1 (5 דק'): קבלו מבחן דמה קצר (5 שאלות) – ענו ללא עזרה",
            "שלב 2 (15 דק'): שלחו את המבחן + תשובותיכם ל-AI עם הנחיה: "
            "'נתח את הטעויות שלי, זהה דפוס ותן 3 המלצות ספציפיות'",
            "שלב 3 (10 דק'): כתבו: 'מה AI זיהה שלא הייתי מזהה לבד?'",
            "שלב 4 (10 דק'): תכנית פעולה אישית – 3 צעדים קונקרטיים לשיפור",
        ],
        "materials": "מבחן דמה מוכן, גישה ל-AI עם יכולת קבלת קבצים/טקסט"
    },
    "Q4": {
        "title": "📊 'מד הלמידה: בוחן עצמי אמין'",
        "goal": "לפתח הערכה עצמית פנימית ולא תלות בציון חיצוני",
        "time": "35 דקות",
        "steps": [
            "שלב 1: בקשו מ-AI: 'צור לי 5 שאלות ברמה 3 ו-5 ברמה 5 על הנושא X'",
            "שלב 2: ענו על השאלות ללא AI – תעדו: מה ידעתי? מה ניחשתי?",
            "שלב 3: בדקו תשובות מול AI – כמה צדקתם? איפה נפלתם?",
            "שלב 4: כתבו: 'הדרך שלמדתי הייתה טובה כי... / לא טובה כי...'",
        ],
        "materials": "גישה ל-AI, מחברת רפלקציה אישית"
    },
    "Q6": {
        "title": "🧘 'AI ואני: כלים לרגיעה ופוקוס'",
        "goal": "לפתח ויסות רגשי משולב עם שימוש נבון ב-AI",
        "time": "40 דקות",
        "steps": [
            "שלב 1: הדגמה – כיצד לבקש מ-AI לעזור לפרק משימה מלחיצה לחלקים קטנים",
            "שלב 2: בזוגות – כל זוג מקבל 'אתגר לחץ': 'המבחן מחר ולא הכנת כלום'",
            "שלב 3: בנו עם AI תכנית למידה מופחצת-לחץ של 3 שעות",
            "שלב 4: רפלקציה: 'האם AI עוזר לי להרגיש פחות לחוץ? למה/למה לא?'",
        ],
        "materials": "תרחישי לחץ מוכנים, גישה ל-AI"
    },
    "Q9": {
        "title": "📅 'לו\"ז AI-BoT: תכנון לטווח של 3 שבועות'",
        "goal": "לפתח תכנון ארוך טווח עם AI ככלי תמיכה",
        "time": "45 דקות",
        "steps": [
            "שלב 1: כל תלמיד שולח ל-AI: 'יש לי מבחן בעוד 3 שבועות בנושאים X,Y,Z. "
            "כמה שעות שבועיות אני צריך? בנה לי לו\"ז'",
            "שלב 2: בחינה ביקורתית: 'האם הלו\"ז ריאלי? מה תשנה?'",
            "שלב 3: שיפור הלו\"ז – שלבו הפסקות, ימי חזרה, בדיקות ביניים",
            "שלב 4: הסכמה בזוגות – 'מה אני מתחייב לעשות עד השבוע הבא?'",
        ],
        "materials": "גישה ל-AI, לוח שנה אישי/גוגל קלנדר"
    },
    "Q7": {
        "title": "🚀 'AI כמנוע יוזמה: שאלות שמשנות שיעורים'",
        "goal": "לפתח פעלנות ויוזמה בלמידה",
        "time": "40 דקות",
        "steps": [
            "שלב 1: כל תלמיד כותב 3 שאלות שרצה לשאול את המורה אבל לא הרהיב",
            "שלב 2: שלח שאלות ל-AI: 'איך לנסח שאלות טובות יותר?'",
            "שלב 3: הצגת השאלות המשופרות בכיתה – בחירת 3 השאלות הטובות ביותר",
            "שלב 4: שליחת השאלות למורה כ'חוזה למידה' – מה נרצה ללמוד?",
        ],
        "materials": "גישה ל-AI, אפשרות לתקשורת עם המורה"
    },
}

# בחירת 3 פעילויות לפי שאלות החלשות ביותר
activity_order = list(bottom3.index) + [q for q in Q_COLS_AVAIL if q not in bottom3.index]
chosen = []
for q in activity_order:
    if q in ACTIVITY_BANK and len(chosen) < 3:
        chosen.append(q)
if len(chosen) < 3:
    chosen += [q for q in ACTIVITY_BANK if q not in chosen][:3-len(chosen)]

for num, q in enumerate(chosen[:3], 1):
    act = ACTIVITY_BANK[q]
    lines.append(f"── פעילות {num}: {act['title']}")
    lines.append(f"   מטרה: {act['goal']}")
    lines.append(f"   זמן: {act['time']}  |  תחום: {Q_SHORT.get(q,q)}")
    lines.append("   שלבים:")
    for step in act["steps"]:
        lines.append(b(step, 2))
    lines.append(f"   חומרים: {act.get('materials','')}")
    lines.append("")

# ═══════════════════════════════════════════════════════════════════════════
# סיכום מנהלים
# ═══════════════════════════════════════════════════════════════════════════
lines.append(h1("סיכום מנהלים"))
strong_names = [Q_SHORT[q] for q in top3.index]
weak_names   = [Q_SHORT[q] for q in bottom3.index]
lines.append(f"ניתוח {n:,} תלמידים מ-{nm} מוסדות ב-{np_} קבוצות פיילוט.")
lines.append(f"ציון ממוצע: {avg:.2f}/5  |  {below3_pct:.0f}% מתחת לסף רמה 3.")
lines.append("")
lines.append("ממצאים מרכזיים:")
lines.append(b(f"חוזקות: {', '.join(strong_names)}"))
lines.append(b(f"חולשות: {', '.join(weak_names)}"))
lines.append(b(f"דפוס שימוש AI נפוץ: עיבוד/סיכום בלבד (≈35% ממשתמשי AI)"))
lines.append(b(f"רק {metacog_pct:.0f}% מגיעים למטא-קוגניציה (רמה 5) – הצלחה עתידית תלויה בפיתוח זה"))
lines.append("")
lines.append("המלצה ראשית: טפחו חשיבה ביקורתית כלפי AI. "
             "ה-3 פעילויות המוצעות מכוונות ישירות לחולשות שזוהו.")

# ═══════════════════════════════════════════════════════════════════════════
# שמירת הדוח
# ═══════════════════════════════════════════════════════════════════════════
report_text = "\n".join(lines)

with open(REPORT_FILE, "w", encoding="utf-8") as f:
    f.write(
        f"דוח ניתוח נתוני סימולציות – אוריינות בינה מלאכותית\n"
        f"גיליון: סימולציות 50-59  |  מקור: {os.path.basename(DATA_FILE)}\n"
        f"נוצר: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M')}\n"
        f"{'='*72}\n" + report_text
    )

print("\n" + "="*72)
print("הניתוח הושלם בהצלחה!")
print("="*72)
print(report_text)
charts = sorted([f for f in os.listdir(OUT_DIR) if f.endswith(".png")])
print(f"\nדוח נשמר: {REPORT_FILE}")
print(f"תרשימים ({len(charts)}):")
for c in charts:
    print(f"  {c}")
