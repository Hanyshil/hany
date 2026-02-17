"""
ניתוח דפוסי פעלנות לומדים - שאלות 50-59
Analysis of Student Learner Proactiveness Patterns - Questions 50-59

מבוסס על מחוון הערכה עם 6 ממדים ו-5 רמות:
א. מוטיבציה ורלוונטיות (Q50, Q51)
ב. תודעת צמיחה (Q52, Q53)
ג. יוזמה ואחריות (Q54)
ד. ויסות עצמי (Q55, Q56)
ה. מודעות עצמית (Q57, Q58)
ו. תמיכה וחוויות רגשיות (Q59)
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib
import os

matplotlib.rcParams['font.family'] = 'DejaVu Sans'
matplotlib.rcParams['axes.unicode_minus'] = False

OUTPUT_DIR = "/home/user/hany/proactiveness_output"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ===== Load Data =====
df = pd.read_excel("/home/user/hany/student_proactiveness_q50_59.xlsx", engine="openpyxl")

# Identify question columns
q_cols = [c for c in df.columns if c.startswith("Q5")]
q_nums = [c.split(":")[0] for c in q_cols]

# Dimension mapping
dimension_map = {
    "Q50": "motivation",
    "Q51": "motivation",
    "Q52": "growth_mindset",
    "Q53": "growth_mindset",
    "Q54": "initiative",
    "Q55": "self_regulation",
    "Q56": "self_regulation",
    "Q57": "self_awareness",
    "Q58": "self_awareness",
    "Q59": "support"
}

dimension_names_he = {
    "motivation": "א. מוטיבציה ורלוונטיות",
    "growth_mindset": "ב. תודעת צמיחה",
    "initiative": "ג. יוזמה ואחריות",
    "self_regulation": "ד. ויסות עצמי",
    "self_awareness": "ה. מודעות עצמית",
    "support": "ו. תמיכה וחוויות רגשיות"
}

dimension_names_en = {
    "motivation": "Motivation & Relevance",
    "growth_mindset": "Growth Mindset",
    "initiative": "Initiative & Responsibility",
    "self_regulation": "Self-Regulation",
    "self_awareness": "Self-Awareness",
    "support": "Support & Emotions"
}

level_names = {
    1: "בתחילת הדרך",
    2: "מתפתח/ת",
    3: "מתקדם/ת",
    4: "מיומן/ת",
    5: "מומחה/ית"
}

level_names_en = {
    1: "Beginning",
    2: "Developing",
    3: "Advancing",
    4: "Proficient",
    5: "Expert"
}

# ===== Calculate Dimension Scores =====
for dim_code in dimension_names_he:
    dim_q_cols = [c for c in q_cols if dimension_map.get(c.split(":")[0]) == dim_code]
    if dim_q_cols:
        df[f"dim_{dim_code}"] = df[dim_q_cols].mean(axis=1).round(1)

# Overall proactiveness score
dim_cols = [c for c in df.columns if c.startswith("dim_")]
df["overall_proactiveness"] = df[dim_cols].mean(axis=1).round(1)

# Assign rubric level
def assign_level(score):
    if score < 1.5:
        return 1
    elif score < 2.5:
        return 2
    elif score < 3.5:
        return 3
    elif score < 4.5:
        return 4
    else:
        return 5

df["overall_level"] = df["overall_proactiveness"].apply(assign_level)
for dim_code in dimension_names_he:
    df[f"level_{dim_code}"] = df[f"dim_{dim_code}"].apply(assign_level)

# =========================================================================
#  REPORT
# =========================================================================
print("=" * 70)
print("  ניתוח דפוסי פעלנות לומדים - שאלות 50-59")
print("  Student Learner Proactiveness Analysis - Q50-Q59")
print("=" * 70)
print(f"\n  Students: {len(df)} | Questions: {len(q_cols)} | Dimensions: {len(dimension_names_he)}")

# --- 1. Overall Class Statistics ---
print("\n" + "=" * 70)
print("  1. סטטיסטיקה כללית לפי ממד / Overall Dimension Statistics")
print("=" * 70)
print(f"\n  {'Dimension':<30} {'Mean':>6} {'Median':>6} {'Min':>4} {'Max':>4} {'StdDev':>7}")
print("  " + "-" * 60)
for dim_code, dim_he in dimension_names_he.items():
    col = f"dim_{dim_code}"
    mean = df[col].mean()
    median = df[col].median()
    mn = df[col].min()
    mx = df[col].max()
    std = df[col].std()
    dim_en = dimension_names_en[dim_code]
    print(f"  {dim_en:<30} {mean:>6.2f} {median:>6.1f} {mn:>4.1f} {mx:>4.1f} {std:>7.2f}")

print(f"\n  {'OVERALL PROACTIVENESS':<30} {df['overall_proactiveness'].mean():>6.2f} "
      f"{df['overall_proactiveness'].median():>6.1f} {df['overall_proactiveness'].min():>4.1f} "
      f"{df['overall_proactiveness'].max():>4.1f} {df['overall_proactiveness'].std():>7.2f}")

# --- 2. Level Distribution per Dimension ---
print("\n" + "=" * 70)
print("  2. התפלגות רמות לפי ממד / Level Distribution per Dimension")
print("=" * 70)

for dim_code, dim_he in dimension_names_he.items():
    level_col = f"level_{dim_code}"
    print(f"\n  {dim_he} ({dimension_names_en[dim_code]}):")
    for lvl in range(1, 6):
        count = (df[level_col] == lvl).sum()
        pct = count / len(df) * 100
        bar = "█" * int(pct / 3)
        print(f"    Level {lvl} ({level_names[lvl]:>14}): {count:>2} students ({pct:>5.1f}%) {bar}")

# --- 3. Overall Level Distribution ---
print("\n" + "=" * 70)
print("  3. התפלגות רמות כללית / Overall Level Distribution")
print("=" * 70)
for lvl in range(1, 6):
    count = (df["overall_level"] == lvl).sum()
    pct = count / len(df) * 100
    bar = "█" * int(pct / 2)
    print(f"  Level {lvl} ({level_names[lvl]:>14} / {level_names_en[lvl]:>12}): "
          f"{count:>2} students ({pct:>5.1f}%) {bar}")

# --- 4. Per-Student Profile ---
print("\n" + "=" * 70)
print("  4. פרופיל תלמיד / Per-Student Profile")
print("=" * 70)

print(f"\n  {'Student':<16} {'Grade':>5} {'Motiv':>6} {'Growth':>6} {'Init':>5} {'Reg':>5} {'Aware':>6} {'Supp':>5} {'TOTAL':>6} {'Level':>12}")
print("  " + "-" * 80)
for _, row in df.sort_values("overall_proactiveness", ascending=False).iterrows():
    name = row["שם התלמיד/ה"]
    grade = row["כיתה"]
    m = row["dim_motivation"]
    g = row["dim_growth_mindset"]
    i = row["dim_initiative"]
    r = row["dim_self_regulation"]
    a = row["dim_self_awareness"]
    s = row["dim_support"]
    total = row["overall_proactiveness"]
    lvl = level_names_en[row["overall_level"]]
    print(f"  {name:<16} {grade:>5} {m:>6.1f} {g:>6.1f} {i:>5.0f} {r:>5.1f} {a:>6.1f} {s:>5.0f} {total:>6.1f} {lvl:>12}")

# --- 5. Pattern Analysis ---
print("\n" + "=" * 70)
print("  5. ניתוח דפוסים / Pattern Analysis")
print("=" * 70)

# Pattern A: Strong motivation but weak self-regulation
pattern_a = df[(df["dim_motivation"] >= 3.5) & (df["dim_self_regulation"] <= 2.5)]
print(f"\n  דפוס א: מוטיבציה גבוהה + ויסות עצמי נמוך (Pattern A: High Motivation, Low Regulation)")
print(f"  Found: {len(pattern_a)} students")
for _, row in pattern_a.iterrows():
    print(f"    - {row['שם התלמיד/ה']} (motiv={row['dim_motivation']}, reg={row['dim_self_regulation']})")
if len(pattern_a) == 0:
    print("    (No students match this pattern)")

# Pattern B: Strong growth mindset but weak initiative
pattern_b = df[(df["dim_growth_mindset"] >= 3.5) & (df["dim_initiative"] <= 2)]
print(f"\n  דפוס ב: תודעת צמיחה גבוהה + יוזמה נמוכה (Pattern B: High Growth, Low Initiative)")
print(f"  Found: {len(pattern_b)} students")
for _, row in pattern_b.iterrows():
    print(f"    - {row['שם התלמיד/ה']} (growth={row['dim_growth_mindset']}, init={row['dim_initiative']})")
if len(pattern_b) == 0:
    print("    (No students match this pattern)")

# Pattern C: Low self-awareness but high support
pattern_c = df[(df["dim_self_awareness"] <= 2.5) & (df["dim_support"] >= 4)]
print(f"\n  דפוס ג: מודעות עצמית נמוכה + תמיכה גבוהה (Pattern C: Low Awareness, High Support)")
print(f"  Found: {len(pattern_c)} students")
for _, row in pattern_c.iterrows():
    print(f"    - {row['שם התלמיד/ה']} (awareness={row['dim_self_awareness']}, support={row['dim_support']})")
if len(pattern_c) == 0:
    print("    (No students match this pattern)")

# Pattern D: Balanced high (all dimensions >= 3.5)
pattern_d = df[(df[dim_cols] >= 3.5).all(axis=1)]
print(f"\n  דפוס ד: פרופיל מאוזן-גבוה בכל הממדים (Pattern D: Balanced High Profile)")
print(f"  Found: {len(pattern_d)} students")
for _, row in pattern_d.iterrows():
    print(f"    - {row['שם התלמיד/ה']} (overall={row['overall_proactiveness']})")
if len(pattern_d) == 0:
    print("    (No students match this pattern)")

# Pattern E: Balanced low (all dimensions <= 2.5)
pattern_e = df[(df[dim_cols] <= 2.5).all(axis=1)]
print(f"\n  דפוס ה: פרופיל מאוזן-נמוך בכל הממדים (Pattern E: Balanced Low Profile)")
print(f"  Found: {len(pattern_e)} students")
for _, row in pattern_e.iterrows():
    print(f"    - {row['שם התלמיד/ה']} (overall={row['overall_proactiveness']})")
if len(pattern_e) == 0:
    print("    (No students match this pattern)")

# Pattern F: High variation between dimensions (gap >= 3)
df["dim_range"] = df[dim_cols].max(axis=1) - df[dim_cols].min(axis=1)
pattern_f = df[df["dim_range"] >= 3]
print(f"\n  דפוס ו: פער גבוה בין ממדים (Pattern F: High Variation Between Dimensions)")
print(f"  Found: {len(pattern_f)} students")
for _, row in pattern_f.iterrows():
    strongest = max(dim_cols, key=lambda c: row[c])
    weakest = min(dim_cols, key=lambda c: row[c])
    s_name = dimension_names_en[strongest.replace("dim_", "")]
    w_name = dimension_names_en[weakest.replace("dim_", "")]
    print(f"    - {row['שם התלמיד/ה']}: strongest={s_name}({row[strongest]}), weakest={w_name}({row[weakest]}), gap={row['dim_range']}")
if len(pattern_f) == 0:
    print("    (No students match this pattern)")

# --- 6. AI Usage Correlation ---
print("\n" + "=" * 70)
print("  6. קשר בין שימוש ב-AI לפעלנות / AI Usage vs Proactiveness")
print("=" * 70)

ai_users = df[df["כלי AI בשימוש"] != "לא משתמש/ת"]
non_ai = df[df["כלי AI בשימוש"] == "לא משתמש/ת"]

print(f"\n  {'Group':<25} {'N':>3} {'Overall':>8} {'Motiv':>7} {'Growth':>7} {'Init':>6} {'Regul':>6} {'Aware':>6} {'Supp':>6}")
print("  " + "-" * 78)
for label, subset in [("AI Users", ai_users), ("Non-AI Users", non_ai), ("All Students", df)]:
    if len(subset) > 0:
        print(f"  {label:<25} {len(subset):>3} {subset['overall_proactiveness'].mean():>8.2f} "
              f"{subset['dim_motivation'].mean():>7.2f} {subset['dim_growth_mindset'].mean():>7.2f} "
              f"{subset['dim_initiative'].mean():>6.2f} {subset['dim_self_regulation'].mean():>6.2f} "
              f"{subset['dim_self_awareness'].mean():>6.2f} {subset['dim_support'].mean():>6.2f}")

# --- 7. Grade-level Analysis ---
print("\n" + "=" * 70)
print("  7. ניתוח לפי כיתה / Analysis by Grade")
print("=" * 70)

for grade in sorted(df["כיתה"].unique(), key=lambda x: ["ט", "י", "יא", "יב"].index(x)):
    grade_df = df[df["כיתה"] == grade]
    print(f"\n  כיתה {grade} (n={len(grade_df)}):")
    print(f"    Overall: {grade_df['overall_proactiveness'].mean():.2f}")
    for dim_code, dim_en in dimension_names_en.items():
        val = grade_df[f"dim_{dim_code}"].mean()
        print(f"    {dim_en}: {val:.2f}")


# =========================================================================
#  GRAPHS
# =========================================================================
print("\n" + "=" * 70)
print("  יצירת גרפים / Generating Graphs")
print("=" * 70)

colors_5 = ['#E53935', '#FB8C00', '#FDD835', '#43A047', '#1E88E5']
dim_colors = ['#E53935', '#7B1FA2', '#1976D2', '#388E3C', '#F57C00', '#00838F']

# --- Graph 1: Radar/Spider Chart per Dimension (Class Average) ---
fig, ax = plt.subplots(figsize=(8, 8), subplot_kw=dict(polar=True))
dims_order = list(dimension_names_en.keys())
labels = [dimension_names_en[d] for d in dims_order]
values = [df[f"dim_{d}"].mean() for d in dims_order]
values += values[:1]  # close the polygon
angles = np.linspace(0, 2 * np.pi, len(labels), endpoint=False).tolist()
angles += angles[:1]

ax.fill(angles, values, color='#42A5F5', alpha=0.25)
ax.plot(angles, values, 'o-', color='#1565C0', linewidth=2, markersize=8)
ax.set_xticks(angles[:-1])
ax.set_xticklabels(labels, fontsize=10)
ax.set_ylim(0, 5)
ax.set_yticks([1, 2, 3, 4, 5])
ax.set_yticklabels(['1', '2', '3', '4', '5'], fontsize=8)
ax.set_title("Class Average - Proactiveness Dimensions\n(ממוצע כיתתי - ממדי פעלנות)", fontsize=14, pad=20)

# Add level labels on right
for i, (label, val) in enumerate(zip(labels, values[:-1])):
    ax.annotate(f'{val:.1f}', xy=(angles[i], val), fontsize=10, fontweight='bold',
                ha='center', va='bottom', color='#1565C0')

plt.tight_layout()
fig.savefig(f"{OUTPUT_DIR}/01_radar_class_average.png", dpi=150)
print("  Saved: 01_radar_class_average.png")
plt.close()

# --- Graph 2: Level Distribution Heatmap ---
fig, ax = plt.subplots(figsize=(14, 8))

heatmap_data = []
dim_labels_en = []
for dim_code in dims_order:
    level_col = f"level_{dim_code}"
    row = [(df[level_col] == lvl).sum() for lvl in range(1, 6)]
    heatmap_data.append(row)
    dim_labels_en.append(dimension_names_en[dim_code])

heatmap_array = np.array(heatmap_data)
im = ax.imshow(heatmap_array, cmap='YlOrRd', aspect='auto')

ax.set_xticks(range(5))
ax.set_xticklabels([f"Level {i}\n{level_names_en[i]}" for i in range(1, 6)], fontsize=10)
ax.set_yticks(range(len(dim_labels_en)))
ax.set_yticklabels(dim_labels_en, fontsize=11)

# Add text annotations
for i in range(len(heatmap_data)):
    for j in range(5):
        val = heatmap_array[i, j]
        color = 'white' if val > heatmap_array.max() * 0.6 else 'black'
        ax.text(j, i, f'{val}', ha='center', va='center', fontsize=14, fontweight='bold', color=color)

ax.set_title("Level Distribution per Dimension (Student Count)\n(התפלגות רמות לפי ממד - מספר תלמידים)", fontsize=14)
plt.colorbar(im, ax=ax, label='Student Count')
plt.tight_layout()
fig.savefig(f"{OUTPUT_DIR}/02_level_heatmap.png", dpi=150)
print("  Saved: 02_level_heatmap.png")
plt.close()

# --- Graph 3: Overall Level Distribution (Pie) ---
fig, ax = plt.subplots(figsize=(9, 7))
level_counts = df["overall_level"].value_counts().sort_index()
all_levels = pd.Series([0]*5, index=range(1, 6))
all_levels.update(level_counts)
labels_pie = [f"Level {i}: {level_names_en[i]}\n({level_names[i]})" for i in range(1, 6)]
non_zero = all_levels[all_levels > 0]
non_zero_labels = [labels_pie[i-1] for i in non_zero.index]
non_zero_colors = [colors_5[i-1] for i in non_zero.index]

wedges, texts, autotexts = ax.pie(
    non_zero.values, labels=non_zero_labels, autopct='%1.0f%%',
    colors=non_zero_colors, startangle=90, textprops={'fontsize': 10}
)
for at in autotexts:
    at.set_fontsize(12)
    at.set_fontweight('bold')
ax.set_title("Overall Proactiveness Level Distribution\n(התפלגות רמת פעלנות כללית)", fontsize=14)
plt.tight_layout()
fig.savefig(f"{OUTPUT_DIR}/03_overall_level_pie.png", dpi=150)
print("  Saved: 03_overall_level_pie.png")
plt.close()

# --- Graph 4: Per-Student Bar Chart (Overall Score) ---
fig, ax = plt.subplots(figsize=(16, 7))
sorted_df = df.sort_values("overall_proactiveness", ascending=True)
y_pos = range(len(sorted_df))
bar_colors = [colors_5[assign_level(s) - 1] for s in sorted_df["overall_proactiveness"]]
bars = ax.barh(y_pos, sorted_df["overall_proactiveness"].values, color=bar_colors, edgecolor='white')
ax.set_yticks(y_pos)
ax.set_yticklabels(sorted_df["שם התלמיד/ה"].values, fontsize=9)
ax.set_xlabel("Proactiveness Score (1-5)", fontsize=12)
ax.set_title("Overall Proactiveness Score per Student\n(ציון פעלנות כללי לפי תלמיד)", fontsize=14)
ax.set_xlim(0, 5.5)
ax.axvline(x=3, color='gray', linestyle='--', alpha=0.5, label='Midpoint')

for bar, val in zip(bars, sorted_df["overall_proactiveness"].values):
    ax.text(val + 0.05, bar.get_y() + bar.get_height()/2,
            f'{val:.1f}', va='center', fontsize=9, fontweight='bold')

# Legend
from matplotlib.patches import Patch
legend_elements = [Patch(facecolor=colors_5[i], label=f'{level_names_en[i+1]}') for i in range(5)]
ax.legend(handles=legend_elements, loc='lower right', fontsize=9)
plt.tight_layout()
fig.savefig(f"{OUTPUT_DIR}/04_student_overall_scores.png", dpi=150)
print("  Saved: 04_student_overall_scores.png")
plt.close()

# --- Graph 5: Dimension Comparison (Grouped Bar) ---
fig, ax = plt.subplots(figsize=(14, 7))
x = np.arange(len(dims_order))
width = 0.12
grades_list = ["ט", "י", "יא", "יב"]
grade_colors = ['#E53935', '#1976D2', '#388E3C', '#F57C00']

for idx, grade in enumerate(grades_list):
    grade_df = df[df["כיתה"] == grade]
    if len(grade_df) > 0:
        means = [grade_df[f"dim_{d}"].mean() for d in dims_order]
        offset = (idx - len(grades_list)/2 + 0.5) * width
        bars = ax.bar(x + offset, means, width, label=f'Grade {grade}', color=grade_colors[idx], edgecolor='white')
        for bar, val in zip(bars, means):
            ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.03,
                    f'{val:.1f}', ha='center', va='bottom', fontsize=7, fontweight='bold')

ax.set_xticks(x)
ax.set_xticklabels([dimension_names_en[d] for d in dims_order], fontsize=10, rotation=15, ha='right')
ax.set_ylabel("Average Score (1-5)", fontsize=12)
ax.set_title("Dimension Scores by Grade\n(ציוני ממדים לפי כיתה)", fontsize=14)
ax.set_ylim(0, 5.5)
ax.legend(fontsize=10)
ax.axhline(y=3, color='gray', linestyle='--', alpha=0.3)
plt.tight_layout()
fig.savefig(f"{OUTPUT_DIR}/05_dimensions_by_grade.png", dpi=150)
print("  Saved: 05_dimensions_by_grade.png")
plt.close()

# --- Graph 6: Student Dimension Profiles (Selected Students - Top, Middle, Bottom) ---
fig, axes = plt.subplots(1, 3, figsize=(18, 6), subplot_kw=dict(polar=True))

sorted_all = df.sort_values("overall_proactiveness", ascending=False)
selections = [
    ("Top Student", sorted_all.iloc[0]),
    ("Middle Student", sorted_all.iloc[len(sorted_all)//2]),
    ("Bottom Student", sorted_all.iloc[-1])
]

for ax_idx, (title, row) in enumerate(selections):
    ax = axes[ax_idx]
    values = [row[f"dim_{d}"] for d in dims_order]
    values += values[:1]
    angles = np.linspace(0, 2 * np.pi, len(dims_order), endpoint=False).tolist()
    angles += angles[:1]

    ax.fill(angles, values, alpha=0.25, color=dim_colors[ax_idx])
    ax.plot(angles, values, 'o-', color=dim_colors[ax_idx], linewidth=2)
    ax.set_xticks(angles[:-1])
    ax.set_xticklabels([dimension_names_en[d][:8] for d in dims_order], fontsize=8)
    ax.set_ylim(0, 5)
    ax.set_yticks([1, 2, 3, 4, 5])
    ax.set_yticklabels(['1', '2', '3', '4', '5'], fontsize=7)
    ax.set_title(f"{title}\n{row['שם התלמיד/ה']} ({row['overall_proactiveness']:.1f})", fontsize=12)

plt.suptitle("Individual Student Profiles\n(פרופילי תלמידים בודדים)", fontsize=14, y=1.02)
plt.tight_layout()
fig.savefig(f"{OUTPUT_DIR}/06_individual_profiles.png", dpi=150, bbox_inches='tight')
print("  Saved: 06_individual_profiles.png")
plt.close()

# --- Graph 7: Correlation Matrix Between Dimensions ---
fig, ax = plt.subplots(figsize=(10, 8))
corr_data = df[[f"dim_{d}" for d in dims_order]].rename(
    columns={f"dim_{d}": dimension_names_en[d] for d in dims_order}
)
corr_matrix = corr_data.corr()

im = ax.imshow(corr_matrix.values, cmap='RdBu_r', vmin=-1, vmax=1, aspect='auto')
ax.set_xticks(range(len(corr_matrix)))
ax.set_xticklabels(corr_matrix.columns, fontsize=9, rotation=30, ha='right')
ax.set_yticks(range(len(corr_matrix)))
ax.set_yticklabels(corr_matrix.columns, fontsize=9)

for i in range(len(corr_matrix)):
    for j in range(len(corr_matrix)):
        val = corr_matrix.values[i, j]
        color = 'white' if abs(val) > 0.5 else 'black'
        ax.text(j, i, f'{val:.2f}', ha='center', va='center', fontsize=11, fontweight='bold', color=color)

ax.set_title("Correlation Between Dimensions\n(מתאם בין ממדי פעלנות)", fontsize=14)
plt.colorbar(im, ax=ax, label='Correlation')
plt.tight_layout()
fig.savefig(f"{OUTPUT_DIR}/07_dimension_correlations.png", dpi=150)
print("  Saved: 07_dimension_correlations.png")
plt.close()

# --- Graph 8: AI Usage vs Proactiveness ---
fig, ax = plt.subplots(figsize=(10, 7))
for tool in df["כלי AI בשימוש"].unique():
    subset = df[df["כלי AI בשימוש"] == tool]
    ax.scatter(subset["dim_motivation"], subset["overall_proactiveness"],
               s=100, alpha=0.7, label=tool, edgecolors='black', linewidth=0.5)

ax.set_xlabel("Motivation Score", fontsize=12)
ax.set_ylabel("Overall Proactiveness", fontsize=12)
ax.set_title("Motivation vs Overall Proactiveness (by AI Tool)\n(מוטיבציה מול פעלנות כללית לפי כלי AI)", fontsize=14)
ax.legend(fontsize=9, loc='lower right')
ax.set_xlim(0.5, 5.5)
ax.set_ylim(0.5, 5.5)
ax.axhline(y=3, color='gray', linestyle='--', alpha=0.3)
ax.axvline(x=3, color='gray', linestyle='--', alpha=0.3)
plt.tight_layout()
fig.savefig(f"{OUTPUT_DIR}/08_motivation_vs_proactiveness.png", dpi=150)
print("  Saved: 08_motivation_vs_proactiveness.png")
plt.close()

# --- Graph 9: Stacked Bar - Level Distribution per Dimension ---
fig, ax = plt.subplots(figsize=(14, 7))
bottom = np.zeros(len(dims_order))
for lvl in range(1, 6):
    counts = []
    for dim_code in dims_order:
        counts.append((df[f"level_{dim_code}"] == lvl).sum())
    ax.bar(range(len(dims_order)), counts, bottom=bottom,
           label=f'Level {lvl}: {level_names_en[lvl]}', color=colors_5[lvl-1], edgecolor='white')
    bottom += np.array(counts)

ax.set_xticks(range(len(dims_order)))
ax.set_xticklabels([dimension_names_en[d] for d in dims_order], fontsize=10, rotation=15, ha='right')
ax.set_ylabel("Number of Students", fontsize=12)
ax.set_title("Level Distribution Across Dimensions (Stacked)\n(התפלגות רמות על פני הממדים)", fontsize=14)
ax.legend(fontsize=9, loc='upper right')
plt.tight_layout()
fig.savefig(f"{OUTPUT_DIR}/09_stacked_levels.png", dpi=150)
print("  Saved: 09_stacked_levels.png")
plt.close()

# --- Graph 10: Strongest vs Weakest Dimension per Student ---
fig, ax = plt.subplots(figsize=(14, 8))

for i, (_, row) in enumerate(df.iterrows()):
    dim_scores = {d: row[f"dim_{d}"] for d in dims_order}
    strongest = max(dim_scores, key=dim_scores.get)
    weakest = min(dim_scores, key=dim_scores.get)
    gap = dim_scores[strongest] - dim_scores[weakest]

    ax.barh(i, dim_scores[strongest], color='#43A047', alpha=0.7, label='Strongest' if i == 0 else '')
    ax.barh(i, -dim_scores[weakest], color='#E53935', alpha=0.7, label='Weakest' if i == 0 else '')
    ax.text(dim_scores[strongest] + 0.1, i, f'{dimension_names_en[strongest][:6]} ({dim_scores[strongest]:.0f})',
            va='center', fontsize=7, color='#2E7D32')
    ax.text(-dim_scores[weakest] - 0.1, i, f'{dimension_names_en[weakest][:6]} ({dim_scores[weakest]:.0f})',
            va='center', fontsize=7, color='#C62828', ha='right')

ax.set_yticks(range(len(df)))
ax.set_yticklabels(df["שם התלמיד/ה"].values, fontsize=8)
ax.set_xlabel("Score", fontsize=12)
ax.set_title("Strongest vs Weakest Dimension per Student\n(ממד חזק מול חלש ביותר לכל תלמיד)", fontsize=14)
ax.axvline(x=0, color='black', linewidth=1)
ax.legend(fontsize=10, loc='lower right')
plt.tight_layout()
fig.savefig(f"{OUTPUT_DIR}/10_strongest_weakest.png", dpi=150)
print("  Saved: 10_strongest_weakest.png")
plt.close()


# =========================================================================
#  KEY INSIGHTS
# =========================================================================
print("\n" + "=" * 70)
print("  סיכום ותובנות מרכזיות / Summary & Key Insights")
print("=" * 70)

# Strongest dimension
dim_means = {d: df[f"dim_{d}"].mean() for d in dims_order}
strongest_dim = max(dim_means, key=dim_means.get)
weakest_dim = min(dim_means, key=dim_means.get)

print(f"""
  1. STRONGEST DIMENSION (ממד חזק ביותר):
     {dimension_names_en[strongest_dim]} ({dimension_names_he[strongest_dim]})
     Average: {dim_means[strongest_dim]:.2f}/5

  2. WEAKEST DIMENSION (ממד חלש ביותר):
     {dimension_names_en[weakest_dim]} ({dimension_names_he[weakest_dim]})
     Average: {dim_means[weakest_dim]:.2f}/5

  3. OVERALL CLASS AVERAGE (ממוצע כיתתי):
     {df['overall_proactiveness'].mean():.2f}/5 ({level_names_en[assign_level(df['overall_proactiveness'].mean())]})

  4. TOP 3 STUDENTS (3 תלמידים מובילים):""")

top3 = df.nlargest(3, "overall_proactiveness")
for _, row in top3.iterrows():
    print(f"     - {row['שם התלמיד/ה']}: {row['overall_proactiveness']:.1f}/5 ({level_names_en[row['overall_level']]})")

print(f"""
  5. STUDENTS NEEDING SUPPORT (תלמידים הזקוקים לתמיכה):""")
bottom3 = df.nsmallest(3, "overall_proactiveness")
for _, row in bottom3.iterrows():
    weak = min(dims_order, key=lambda d: row[f"dim_{d}"])
    print(f"     - {row['שם התלמיד/ה']}: {row['overall_proactiveness']:.1f}/5 (weakest: {dimension_names_en[weak]})")

print(f"""
  6. DIMENSION GAP (פער בין ממדים):
     Strongest class dimension: {dimension_names_en[strongest_dim]} ({dim_means[strongest_dim]:.2f})
     Weakest class dimension: {dimension_names_en[weakest_dim]} ({dim_means[weakest_dim]:.2f})
     Gap: {dim_means[strongest_dim] - dim_means[weakest_dim]:.2f}

  7. KEY PATTERNS FOUND (דפוסים שנמצאו):
     - Pattern A (High Motivation, Low Regulation): {len(pattern_a)} students
     - Pattern B (High Growth, Low Initiative): {len(pattern_b)} students
     - Pattern C (Low Awareness, High Support): {len(pattern_c)} students
     - Pattern D (Balanced High): {len(pattern_d)} students
     - Pattern E (Balanced Low): {len(pattern_e)} students
     - Pattern F (High Variation): {len(pattern_f)} students

  8. AI USAGE CONNECTION (קשר לשימוש ב-AI):
     AI users average proactiveness: {ai_users['overall_proactiveness'].mean():.2f}
     Non-AI users average proactiveness: {non_ai['overall_proactiveness'].mean():.2f}

  10 GRAPHS SAVED to {OUTPUT_DIR}/
""")

print("=" * 70)
print("  Analysis complete!")
print("=" * 70)
