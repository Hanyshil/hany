"""
סימולציית שאלות 50-59 מבוססות על מחוון הערכת פעלנות לומדים
Simulation of Questions 50-59 based on Student Learner Proactiveness Rubric

ממדים (6 ממדי המחוון):
א. מוטיבציה ורלוונטיות (שאלות 50-51)
ב. תודעת צמיחה (שאלות 52-53)
ג. יוזמה ואחריות (שאלה 54)
ד. ויסות עצמי (שאלות 55-56)
ה. מודעות עצמית (שאלות 57-58)
ו. תמיכה וחוויות רגשיות (שאלה 59)

רמות (1-5):
1 = בתחילת הדרך
2 = מתפתח/ת
3 = מתקדם/ת
4 = מיומן/ת
5 = מומחה/ית
"""

import pandas as pd
import numpy as np
import random

random.seed(42)
np.random.seed(42)

# Read existing student data
df_existing = pd.read_excel("/home/user/hany/student_ai_usage.xlsx", engine="openpyxl")
students = df_existing["שם התלמיד/ה"].tolist()
grades = df_existing["כיתה"].tolist()
ai_tools = df_existing["כלי AI בשימוש"].tolist()
ai_hours = df_existing["שעות שימוש שבועיות"].tolist()

# Define questions 50-59 mapped to rubric dimensions
questions = {
    "Q50": {
        "text": "עד כמה השיעורים מעניינים אותך ואתה מרגיש שהלמידה רלוונטית לחיים שלך?",
        "dimension": "א. מוטיבציה ורלוונטיות",
        "dimension_code": "motivation"
    },
    "Q51": {
        "text": "עד כמה אתה מוכן להתמודד עם משימות מאתגרות ומשקיע בלמידה מעבר לציונים?",
        "dimension": "א. מוטיבציה ורלוונטיות",
        "dimension_code": "motivation"
    },
    "Q52": {
        "text": "עד כמה אתה מאמין שאתה יכול להשתפר בכל דבר באמצעות מאמץ והתמדה?",
        "dimension": "ב. תודעת צמיחה",
        "dimension_code": "growth_mindset"
    },
    "Q53": {
        "text": "כשאתה נתקל בקושי או טועה, עד כמה אתה רואה בזה הזדמנות ללמוד ולא מוותר?",
        "dimension": "ב. תודעת צמיחה",
        "dimension_code": "growth_mindset"
    },
    "Q54": {
        "text": "עד כמה אתה לוקח אחריות על הלמידה שלך, פועל בעצמאות ומציע יוזמות?",
        "dimension": "ג. יוזמה ואחריות",
        "dimension_code": "initiative"
    },
    "Q55": {
        "text": "עד כמה אתה מתכנן משימות מראש ומנהל את הזמן שלך בצורה יעילה?",
        "dimension": "ד. ויסות עצמי",
        "dimension_code": "self_regulation"
    },
    "Q56": {
        "text": "כשאתה מתוסכל או לחוץ, עד כמה אתה מצליח להרגיע את עצמך ולהמשיך ללמוד?",
        "dimension": "ד. ויסות עצמי",
        "dimension_code": "self_regulation"
    },
    "Q57": {
        "text": "עד כמה אתה מכיר את החוזקות והחולשות שלך בלמידה?",
        "dimension": "ה. מודעות עצמית",
        "dimension_code": "self_awareness"
    },
    "Q58": {
        "text": "עד כמה אתה חושב על ההצלחות והכישלונות שלך ומפיק לקחים?",
        "dimension": "ה. מודעות עצמית",
        "dimension_code": "self_awareness"
    },
    "Q59": {
        "text": "עד כמה אתה יודע לבקש עזרה כשצריך ומרגיש שיש לך תמיכה (ממורים, חברים, משפחה)?",
        "dimension": "ו. תמיכה וחוויות רגשיות",
        "dimension_code": "support"
    }
}

# Student personality profiles that create realistic patterns
# Each student gets a base profile that influences their responses
# Profile: (motivation_base, growth_base, initiative_base, regulation_base, awareness_base, support_base)
def generate_student_profile(student_idx, ai_tool, hours):
    """Generate a personality profile based on student characteristics."""

    # Base level with some randomness (realistic distribution - most students are 2-4)
    base = np.random.choice([1, 2, 2, 3, 3, 3, 3, 4, 4, 5], size=6)

    # AI heavy users tend to have higher self-regulation and initiative
    if hours >= 2.0:
        base[2] = min(5, base[2] + 1)  # initiative
        base[3] = min(5, base[3] + 1)  # regulation

    # Non-users might have different patterns
    if ai_tool == "לא משתמש/ת":
        # Some non-users are very strong learners who don't need AI
        if random.random() > 0.5:
            base[0] = min(5, base[0] + 1)  # higher motivation
            base[4] = min(5, base[4] + 1)  # higher self-awareness
        # Others might be less engaged
        else:
            base[0] = max(1, base[0] - 1)

    # Add individual variation per dimension
    for i in range(6):
        base[i] = max(1, min(5, base[i] + random.choice([-1, 0, 0, 0, 1])))

    return base

# Generate responses
data = []
for idx, (student, grade, tool, hours) in enumerate(zip(students, grades, ai_tools, ai_hours)):
    profile = generate_student_profile(idx, tool, hours)

    # Map: motivation(0,1), growth(2,3), initiative(4), regulation(5,6), awareness(7,8), support(9)
    dimension_bases = {
        "motivation": profile[0],
        "growth_mindset": profile[1],
        "initiative": profile[2],
        "self_regulation": profile[3],
        "self_awareness": profile[4],
        "support": profile[5]
    }

    row = {
        "שם התלמיד/ה": student,
        "כיתה": grade,
        "כלי AI בשימוש": tool,
    }

    for q_id, q_info in questions.items():
        dim_code = q_info["dimension_code"]
        base = dimension_bases[dim_code]
        # Add small noise per question (±1) for realism
        response = max(1, min(5, base + random.choice([-1, 0, 0, 1])))
        row[f"{q_id}: {q_info['text']}"] = response

    data.append(row)

df_sim = pd.DataFrame(data)

# Save to Excel
output_path = "/home/user/hany/student_proactiveness_q50_59.xlsx"
df_sim.to_excel(output_path, index=False, engine="openpyxl")

print(f"Created: {output_path}")
print(f"Shape: {df_sim.shape}")
print(f"\nQuestions (Q50-Q59):")
for q_id, q_info in questions.items():
    print(f"  {q_id} [{q_info['dimension']}]: {q_info['text']}")

# Show response distribution
print(f"\n--- Response Distribution (1-5) ---")
q_cols = [c for c in df_sim.columns if c.startswith("Q5")]
for col in q_cols:
    q_num = col.split(":")[0]
    mean_val = df_sim[col].mean()
    dist = df_sim[col].value_counts().sort_index()
    print(f"  {q_num}: mean={mean_val:.1f}  |  {dict(dist)}")

print(f"\n--- Sample Data (first 5 students) ---")
display_cols = ["שם התלמיד/ה", "כיתה"] + q_cols[:3]
print(df_sim[display_cols].head().to_string())
