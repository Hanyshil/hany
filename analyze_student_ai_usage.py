"""
ניתוח דפוסי שימוש בבינה מלאכותית של תלמידים
Analysis of Student AI Usage Patterns
"""

import pandas as pd
import matplotlib.pyplot as plt
import matplotlib
import os

# Hebrew/RTL font support
matplotlib.rcParams['font.family'] = 'DejaVu Sans'
matplotlib.rcParams['axes.unicode_minus'] = False

OUTPUT_DIR = "/home/user/hany/analysis_output"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ===== 1. Read the Excel file =====
df = pd.read_excel("/home/user/hany/student_ai_usage.xlsx", engine="openpyxl")

print("=" * 60)
print("  ניתוח דפוסי שימוש בבינה מלאכותית - תלמידים")
print("  Student AI Usage Pattern Analysis")
print("=" * 60)
print(f"\nTotal students surveyed: {len(df)}")
print(f"Columns: {len(df.columns)}")
print(f"\nColumn names:\n")
for i, col in enumerate(df.columns, 1):
    print(f"  {i}. {col}")

# ===== 2. Basic Statistics =====
print("\n" + "=" * 60)
print("  סטטיסטיקה בסיסית / Basic Statistics")
print("=" * 60)

# Grade distribution
print("\n--- Distribution by Grade (כיתה) ---")
grade_dist = df["כיתה"].value_counts()
for grade, count in grade_dist.items():
    print(f"  Grade {grade}: {count} students ({count/len(df)*100:.0f}%)")

# AI tool usage
print("\n--- AI Tools Used (כלי AI בשימוש) ---")
tool_dist = df["כלי AI בשימוש"].value_counts()
for tool, count in tool_dist.items():
    print(f"  {tool}: {count} students ({count/len(df)*100:.0f}%)")

# Usage frequency
print("\n--- Usage Frequency (תדירות שימוש) ---")
freq_dist = df["תדירות שימוש"].value_counts()
for freq, count in freq_dist.items():
    print(f"  {freq}: {count} students ({count/len(df)*100:.0f}%)")

# ===== 3. Deeper Analysis =====
print("\n" + "=" * 60)
print("  ניתוח מעמיק / Deeper Analysis")
print("=" * 60)

# Filter only active AI users
active_users = df[df["כלי AI בשימוש"] != "לא משתמש/ת"]
non_users = df[df["כלי AI בשימוש"] == "לא משתמש/ת"]
print(f"\nActive AI users: {len(active_users)} ({len(active_users)/len(df)*100:.0f}%)")
print(f"Non-users: {len(non_users)} ({len(non_users)/len(df)*100:.0f}%)")

# Average weekly hours
if len(active_users) > 0:
    avg_hours = active_users["שעות שימוש שבועיות"].mean()
    max_hours = active_users["שעות שימוש שבועיות"].max()
    min_hours = active_users["שעות שימוש שבועיות"].min()
    print(f"\nWeekly usage hours (active users):")
    print(f"  Average: {avg_hours:.1f} hours")
    print(f"  Max: {max_hours:.1f} hours")
    print(f"  Min: {min_hours:.1f} hours")

# Main purposes
print("\n--- Main Usage Purposes (מטרת השימוש) ---")
purpose_dist = active_users["מטרת השימוש העיקרית"].value_counts()
for purpose, count in purpose_dist.items():
    print(f"  {purpose}: {count} students")

# Subjects
print("\n--- Main Subjects for AI Use (מקצוע עיקרי) ---")
subject_dist = active_users["מקצוע עיקרי לשימוש"].value_counts()
for subject, count in subject_dist.items():
    print(f"  {subject}: {count} students")

# Helpfulness
print("\n--- Perceived Helpfulness (רמת תועלת) ---")
help_dist = active_users["רמת תועלת"].value_counts()
for level, count in help_dist.items():
    print(f"  {level}: {count} students ({count/len(active_users)*100:.0f}%)")

# Concerns
print("\n--- Concerns (חששות) ---")
concern_dist = df["חששות"].value_counts()
for concern, count in concern_dist.items():
    print(f"  {concern}: {count} students")

# Score changes
print("\n--- Self-Score Changes (שינוי ציון עצמי) ---")
df["score_change"] = df["ציון עצמי אחרי AI (0-100)"] - df["ציון עצמי לפני AI (0-100)"]
avg_change_all = df["score_change"].mean()
print(f"  Average change (all students): {avg_change_all:+.1f}")

if len(active_users) > 0:
    active_change = active_users["ציון עצמי אחרי AI (0-100)"].values - active_users["ציון עצמי לפני AI (0-100)"].values
    print(f"  Average change (AI users): {active_change.mean():+.1f}")

if len(non_users) > 0:
    non_user_change = non_users["ציון עצמי אחרי AI (0-100)"].values - non_users["ציון עצמי לפני AI (0-100)"].values
    print(f"  Average change (non-users): {non_user_change.mean():+.1f}")

# Per-student summary
print("\n" + "=" * 60)
print("  דפוסי שימוש לפי תלמיד / Per-Student Patterns")
print("=" * 60)
for _, row in df.iterrows():
    name = row["שם התלמיד/ה"]
    grade = row["כיתה"]
    tool = row["כלי AI בשימוש"]
    freq = row["תדירות שימוש"]
    hours = row["שעות שימוש שבועיות"]
    change = row["ציון עצמי אחרי AI (0-100)"] - row["ציון עצמי לפני AI (0-100)"]
    print(f"  {name} (כיתה {grade}): {tool} | {freq} | {hours}h/week | score change: {change:+d}")


# ===== 4. Generate Graphs =====
print("\n" + "=" * 60)
print("  יצירת גרפים / Generating Graphs")
print("=" * 60)

# --- Graph 1: AI Tool Popularity ---
fig, ax = plt.subplots(figsize=(10, 6))
colors = ['#4CAF50', '#2196F3', '#FF9800', '#9C27B0', '#F44336', '#607D8B']
tool_counts = df["כלי AI בשימוש"].value_counts()
bars = ax.bar(range(len(tool_counts)), tool_counts.values, color=colors[:len(tool_counts)])
ax.set_xticks(range(len(tool_counts)))
ax.set_xticklabels(tool_counts.index, fontsize=11)
ax.set_ylabel("Number of Students", fontsize=12)
ax.set_title("AI Tools Used by Students\n(כלי AI בשימוש התלמידים)", fontsize=14)
for bar, val in zip(bars, tool_counts.values):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.1,
            str(val), ha='center', va='bottom', fontsize=12, fontweight='bold')
ax.set_ylim(0, max(tool_counts.values) + 2)
plt.tight_layout()
fig.savefig(f"{OUTPUT_DIR}/01_ai_tools_popularity.png", dpi=150)
print("  Saved: 01_ai_tools_popularity.png")
plt.close()

# --- Graph 2: Usage Frequency ---
fig, ax = plt.subplots(figsize=(10, 6))
freq_order = ["כל יום", "כמה פעמים בשבוע", "פעם בשבוע", "כמה פעמים בחודש", "לעיתים רחוקות", "אף פעם"]
freq_counts = df["תדירות שימוש"].value_counts().reindex(freq_order).dropna()
colors_freq = ['#D32F2F', '#F57C00', '#FBC02D', '#388E3C', '#1976D2', '#757575']
wedges, texts, autotexts = ax.pie(
    freq_counts.values,
    labels=freq_counts.index,
    autopct='%1.0f%%',
    colors=colors_freq[:len(freq_counts)],
    startangle=90,
    textprops={'fontsize': 10}
)
for autotext in autotexts:
    autotext.set_fontsize(11)
    autotext.set_fontweight('bold')
ax.set_title("Usage Frequency Distribution\n(התפלגות תדירות השימוש)", fontsize=14)
plt.tight_layout()
fig.savefig(f"{OUTPUT_DIR}/02_usage_frequency.png", dpi=150)
print("  Saved: 02_usage_frequency.png")
plt.close()

# --- Graph 3: Usage Purpose ---
fig, ax = plt.subplots(figsize=(12, 6))
purpose_counts = active_users["מטרת השימוש העיקרית"].value_counts()
bars = ax.barh(range(len(purpose_counts)), purpose_counts.values, color='#42A5F5')
ax.set_yticks(range(len(purpose_counts)))
ax.set_yticklabels(purpose_counts.index, fontsize=11)
ax.set_xlabel("Number of Students", fontsize=12)
ax.set_title("Main Purposes for AI Usage\n(מטרות עיקריות לשימוש ב-AI)", fontsize=14)
for bar, val in zip(bars, purpose_counts.values):
    ax.text(bar.get_width() + 0.05, bar.get_y() + bar.get_height()/2,
            str(val), ha='left', va='center', fontsize=11, fontweight='bold')
plt.tight_layout()
fig.savefig(f"{OUTPUT_DIR}/03_usage_purposes.png", dpi=150)
print("  Saved: 03_usage_purposes.png")
plt.close()

# --- Graph 4: Subject Distribution ---
fig, ax = plt.subplots(figsize=(10, 6))
subject_counts = active_users["מקצוע עיקרי לשימוש"].value_counts()
colors_subj = plt.cm.Set3(range(len(subject_counts)))
bars = ax.bar(range(len(subject_counts)), subject_counts.values, color=colors_subj)
ax.set_xticks(range(len(subject_counts)))
ax.set_xticklabels(subject_counts.index, fontsize=11, rotation=30, ha='right')
ax.set_ylabel("Number of Students", fontsize=12)
ax.set_title("AI Usage by Subject\n(שימוש ב-AI לפי מקצוע)", fontsize=14)
for bar, val in zip(bars, subject_counts.values):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.1,
            str(val), ha='center', va='bottom', fontsize=11, fontweight='bold')
plt.tight_layout()
fig.savefig(f"{OUTPUT_DIR}/04_subject_distribution.png", dpi=150)
print("  Saved: 04_subject_distribution.png")
plt.close()

# --- Graph 5: Helpfulness Rating ---
fig, ax = plt.subplots(figsize=(8, 6))
help_counts = active_users["רמת תועלת"].value_counts()
help_colors = {'מאוד עוזר': '#4CAF50', 'עוזר במידה מסוימת': '#8BC34A',
               'לא באמת עוזר': '#FF9800', 'מזיק ללמידה': '#F44336'}
colors_h = [help_colors.get(h, '#9E9E9E') for h in help_counts.index]
wedges, texts, autotexts = ax.pie(
    help_counts.values,
    labels=help_counts.index,
    autopct='%1.0f%%',
    colors=colors_h,
    startangle=90,
    textprops={'fontsize': 11}
)
for autotext in autotexts:
    autotext.set_fontsize(12)
    autotext.set_fontweight('bold')
ax.set_title("Perceived Helpfulness of AI\n(רמת תועלת נתפסת של AI)", fontsize=14)
plt.tight_layout()
fig.savefig(f"{OUTPUT_DIR}/05_helpfulness_rating.png", dpi=150)
print("  Saved: 05_helpfulness_rating.png")
plt.close()

# --- Graph 6: Weekly Hours by Tool ---
fig, ax = plt.subplots(figsize=(10, 6))
hours_by_tool = active_users.groupby("כלי AI בשימוש")["שעות שימוש שבועיות"].mean().sort_values(ascending=False)
bars = ax.bar(range(len(hours_by_tool)), hours_by_tool.values, color='#AB47BC')
ax.set_xticks(range(len(hours_by_tool)))
ax.set_xticklabels(hours_by_tool.index, fontsize=11)
ax.set_ylabel("Average Weekly Hours", fontsize=12)
ax.set_title("Average Weekly Hours by AI Tool\n(שעות שבועיות ממוצעות לפי כלי)", fontsize=14)
for bar, val in zip(bars, hours_by_tool.values):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.02,
            f'{val:.1f}h', ha='center', va='bottom', fontsize=11, fontweight='bold')
plt.tight_layout()
fig.savefig(f"{OUTPUT_DIR}/06_weekly_hours_by_tool.png", dpi=150)
print("  Saved: 06_weekly_hours_by_tool.png")
plt.close()

# --- Graph 7: Before vs After Self-Scores ---
fig, ax = plt.subplots(figsize=(12, 7))
x = range(len(df))
width = 0.35
bars1 = ax.bar([i - width/2 for i in x], df["ציון עצמי לפני AI (0-100)"].values, width,
               label='Before AI', color='#FFAB91', edgecolor='white')
bars2 = ax.bar([i + width/2 for i in x], df["ציון עצמי אחרי AI (0-100)"].values, width,
               label='After AI', color='#80CBC4', edgecolor='white')
ax.set_xticks(list(x))
ax.set_xticklabels(df["שם התלמיד/ה"].values, rotation=60, ha='right', fontsize=8)
ax.set_ylabel("Self Score (0-100)", fontsize=12)
ax.set_title("Self-Score Before vs After AI Usage\n(ציון עצמי לפני ואחרי שימוש ב-AI)", fontsize=14)
ax.legend(fontsize=11)
ax.set_ylim(0, 110)
plt.tight_layout()
fig.savefig(f"{OUTPUT_DIR}/07_before_after_scores.png", dpi=150)
print("  Saved: 07_before_after_scores.png")
plt.close()

# --- Graph 8: Concerns Distribution ---
fig, ax = plt.subplots(figsize=(12, 6))
concern_counts = df["חששות"].value_counts()
bars = ax.barh(range(len(concern_counts)), concern_counts.values, color='#EF5350')
ax.set_yticks(range(len(concern_counts)))
ax.set_yticklabels(concern_counts.index, fontsize=11)
ax.set_xlabel("Number of Students", fontsize=12)
ax.set_title("Student Concerns About AI\n(חששות תלמידים לגבי AI)", fontsize=14)
for bar, val in zip(bars, concern_counts.values):
    ax.text(bar.get_width() + 0.05, bar.get_y() + bar.get_height()/2,
            str(val), ha='left', va='center', fontsize=11, fontweight='bold')
plt.tight_layout()
fig.savefig(f"{OUTPUT_DIR}/08_concerns.png", dpi=150)
print("  Saved: 08_concerns.png")
plt.close()

# --- Graph 9: Score Change Distribution ---
fig, ax = plt.subplots(figsize=(10, 6))
changes = df["score_change"]
colors_change = ['#4CAF50' if c >= 0 else '#F44336' for c in changes]
bars = ax.bar(range(len(df)), changes.values, color=colors_change)
ax.set_xticks(range(len(df)))
ax.set_xticklabels(df["שם התלמיד/ה"].values, rotation=60, ha='right', fontsize=8)
ax.set_ylabel("Score Change", fontsize=12)
ax.set_title("Self-Score Change per Student\n(שינוי ציון עצמי לפי תלמיד)", fontsize=14)
ax.axhline(y=0, color='black', linewidth=0.8)
plt.tight_layout()
fig.savefig(f"{OUTPUT_DIR}/09_score_changes.png", dpi=150)
print("  Saved: 09_score_changes.png")
plt.close()

# --- Graph 10: Usage by Grade ---
fig, axes = plt.subplots(1, 2, figsize=(14, 6))

grade_tool = pd.crosstab(df["כיתה"], df["כלי AI בשימוש"])
grade_tool.plot(kind='bar', ax=axes[0], colormap='Set2')
axes[0].set_title("AI Tool by Grade\n(כלי AI לפי כיתה)", fontsize=13)
axes[0].set_xlabel("Grade", fontsize=11)
axes[0].set_ylabel("Students", fontsize=11)
axes[0].legend(fontsize=8, loc='upper right')
axes[0].tick_params(axis='x', rotation=0)

grade_hours = df.groupby("כיתה")["שעות שימוש שבועיות"].mean()
bars = axes[1].bar(grade_hours.index, grade_hours.values, color='#7E57C2')
axes[1].set_title("Avg Weekly Hours by Grade\n(שעות שבועיות ממוצעות לפי כיתה)", fontsize=13)
axes[1].set_xlabel("Grade", fontsize=11)
axes[1].set_ylabel("Average Hours", fontsize=11)
for bar, val in zip(bars, grade_hours.values):
    axes[1].text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.02,
                 f'{val:.1f}h', ha='center', fontsize=11, fontweight='bold')

plt.tight_layout()
fig.savefig(f"{OUTPUT_DIR}/10_usage_by_grade.png", dpi=150)
print("  Saved: 10_usage_by_grade.png")
plt.close()


# ===== 5. Summary & Insights =====
print("\n" + "=" * 60)
print("  סיכום ותובנות / Summary & Key Insights")
print("=" * 60)

# Most popular tool
top_tool = tool_dist.index[0]
top_tool_pct = tool_dist.values[0] / len(df) * 100
print(f"\n1. MOST POPULAR AI TOOL: {top_tool} ({top_tool_pct:.0f}% of students)")

# Usage rate
usage_rate = len(active_users) / len(df) * 100
print(f"\n2. AI ADOPTION RATE: {usage_rate:.0f}% of students use AI tools")

# Most common purpose
if len(purpose_dist) > 0:
    top_purpose = purpose_dist.index[0]
    print(f"\n3. TOP USAGE PURPOSE: {top_purpose} ({purpose_dist.values[0]} students)")

# Average hours
if len(active_users) > 0:
    print(f"\n4. AVERAGE WEEKLY USAGE: {avg_hours:.1f} hours per week (among users)")

# Most common subject
if len(subject_dist) > 0:
    top_subject = subject_dist.index[0]
    print(f"\n5. TOP SUBJECT FOR AI: {top_subject} ({subject_dist.values[0]} students)")

# Helpfulness
if len(help_dist) > 0:
    helpful_pct = 0
    for level in ["מאוד עוזר", "עוזר במידה מסוימת"]:
        if level in help_dist.index:
            helpful_pct += help_dist[level]
    helpful_pct = helpful_pct / len(active_users) * 100
    print(f"\n6. POSITIVE PERCEPTION: {helpful_pct:.0f}% of users find AI helpful")

# Score impact
print(f"\n7. SCORE IMPACT:")
print(f"   - AI users average score change: {active_change.mean():+.1f}")
if len(non_users) > 0:
    print(f"   - Non-users average score change: {non_user_change.mean():+.1f}")

# Top concern
top_concern = concern_dist.index[0]
print(f"\n8. TOP CONCERN: {top_concern} ({concern_dist.values[0]} students)")

# Heavy users
heavy_users = active_users[active_users["שעות שימוש שבועיות"] >= 2.0]
print(f"\n9. HEAVY USERS (2+ hours/week): {len(heavy_users)} students")
for _, row in heavy_users.iterrows():
    print(f"   - {row['שם התלמיד/ה']}: {row['שעות שימוש שבועיות']}h/week using {row['כלי AI בשימוש']}")

print(f"\n10. GRAPHS GENERATED: 10 charts saved to {OUTPUT_DIR}/")

print("\n" + "=" * 60)
print("  ✓ Analysis complete!")
print("=" * 60)
