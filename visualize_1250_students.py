"""
ויזואליזציה - ניתוח 1,250 תלמידים: פעלנות לומד ושימוש ב-AI (Q50-Q59)
Visualization – 1,250-Student Proactiveness & AI Usage Analysis (Q50-Q59)

מבוסס על מחוון 6 ממדים × 5 רמות
Based on the 6-dimension × 5-level rubric
"""

import pandas as pd
import numpy as np
import matplotlib
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import os
import random

matplotlib.rcParams['font.family'] = 'DejaVu Sans'
matplotlib.rcParams['axes.unicode_minus'] = False
plt.rcParams['figure.dpi'] = 150

# ──────────────────────────────────────────────────────────
# RE-GENERATE DATA (same seed as map_q50_59_rubric_1250.py)
# ──────────────────────────────────────────────────────────
random.seed(2024)
np.random.seed(2024)

N = 1250
GRADES = ["י", "יא", "יב"]
GRADE_WEIGHTS = [0.35, 0.35, 0.30]
AI_TOOLS = ["ChatGPT", "Gemini", "Copilot", "Claude", "לא משתמש/ת"]
AI_WEIGHTS = [0.38, 0.22, 0.15, 0.10, 0.15]

DIM_CODES = ["motivation", "growth_mindset", "initiative",
             "self_regulation", "self_awareness", "support"]

DIM_OF_Q = {
    "Q50": "motivation",    "Q51": "motivation",
    "Q52": "growth_mindset","Q53": "growth_mindset",
    "Q54": "initiative",
    "Q55": "self_regulation","Q56": "self_regulation",
    "Q57": "self_awareness", "Q58": "self_awareness",
    "Q59": "support",
}

DIM_DIST = {
    "motivation":       [0.08, 0.18, 0.40, 0.24, 0.10],
    "growth_mindset":   [0.10, 0.20, 0.38, 0.23, 0.09],
    "initiative":       [0.12, 0.25, 0.35, 0.20, 0.08],
    "self_regulation":  [0.10, 0.22, 0.36, 0.22, 0.10],
    "self_awareness":   [0.09, 0.19, 0.38, 0.25, 0.09],
    "support":          [0.11, 0.20, 0.37, 0.22, 0.10],
}

RUBRIC_HE = {
    "motivation":      "א. מוטיבציה ורלוונטיות",
    "growth_mindset":  "ב. תודעת צמיחה",
    "initiative":      "ג. יוזמה ואחריות",
    "self_regulation": "ד. ויסות עצמי",
    "self_awareness":  "ה. מודעות עצמית",
    "support":         "ו. תמיכה וחוויות רגשיות",
}
RUBRIC_EN = {
    "motivation":      "Motivation & Relevance",
    "growth_mindset":  "Growth Mindset",
    "initiative":      "Initiative & Responsibility",
    "self_regulation": "Self-Regulation",
    "self_awareness":  "Self-Awareness",
    "support":         "Support & Emotions",
}
LEVEL_HE = {
    1: "בתחילת הדרך",
    2: "מתפתח/ת",
    3: "מתקדם/ת",
    4: "מיומן/ת",
    5: "מומחה/ית",
}
LEVEL_EN = {1: "Beginning", 2: "Developing", 3: "Advancing",
            4: "Proficient", 5: "Expert"}

LEVEL_COLORS = {
    1: "#EF5350",   # red
    2: "#FFA726",   # orange
    3: "#FFEE58",   # yellow
    4: "#66BB6A",   # green
    5: "#42A5F5",   # blue
}
TOOL_COLORS = {
    "ChatGPT":    "#10A37F",
    "Gemini":     "#4285F4",
    "Copilot":    "#7FBA00",
    "Claude":     "#D97706",
    "לא משתמש/ת": "#9E9E9E",
}


def assign_level(score):
    if score < 1.5: return 1
    elif score < 2.5: return 2
    elif score < 3.5: return 3
    elif score < 4.5: return 4
    return 5


def generate_students(n):
    rows = []
    for i in range(n):
        grade = np.random.choice(GRADES, p=GRADE_WEIGHTS)
        ai_tool = np.random.choice(AI_TOOLS, p=AI_WEIGHTS)
        ai_user = ai_tool != "לא משתמש/ת"
        dim_scores = {}
        for dim in DIM_CODES:
            probs = DIM_DIST[dim].copy()
            if ai_user and dim in ("initiative", "self_regulation"):
                probs = [max(0, p - 0.02) for p in probs]
                probs[2] -= 0.02; probs[3] += 0.03; probs[4] += 0.03
                s = sum(probs); probs = [p / s for p in probs]
            lvl = np.random.choice([1, 2, 3, 4, 5], p=probs)
            lo, hi = lvl - 0.49, lvl + 0.49
            dim_scores[dim] = round(np.clip(np.random.uniform(lo, hi), 1.0, 5.0), 1)
        q_answers = {}
        for q, dim in DIM_OF_Q.items():
            base = dim_scores[dim]
            noise = random.choice([-0.5, 0, 0, 0, 0.5])
            q_answers[q] = int(np.clip(round(base + noise), 1, 5))
        dim_avgs = {}
        for dim in DIM_CODES:
            qs = [q for q, d in DIM_OF_Q.items() if d == dim]
            dim_avgs[dim] = round(np.mean([q_answers[q] for q in qs]), 2)
        overall = round(np.mean(list(dim_avgs.values())), 2)
        row = {"כיתה": grade, "כלי AI": ai_tool}
        for q in ["Q50", "Q51", "Q52", "Q53", "Q54",
                  "Q55", "Q56", "Q57", "Q58", "Q59"]:
            row[q] = q_answers[q]
        for dim in DIM_CODES:
            row[f"avg_{dim}"] = dim_avgs[dim]
            row[f"lvl_{dim}"] = assign_level(dim_avgs[dim])
        row["avg_total"] = overall
        row["lvl_total"] = assign_level(overall)
        row["lvl_name"] = LEVEL_HE[assign_level(overall)]
        rows.append(row)
    return pd.DataFrame(rows)


print("⏳ Generating 1,250 students …")
df = generate_students(N)
print(f"✅ {len(df)} students generated")

OUT = "/home/user/hany/visualize_output_1250"
os.makedirs(OUT, exist_ok=True)

# ══════════════════════════════════════════════════════════
#  GRAPH 01 – Radar: Class Average across 6 Dimensions
# ══════════════════════════════════════════════════════════
print("  📊 01 Radar – class average …")
labels = [RUBRIC_EN[d] for d in DIM_CODES]
values = [df[f"avg_{d}"].mean() for d in DIM_CODES]
angles = np.linspace(0, 2 * np.pi, len(labels), endpoint=False).tolist()
values_loop = values + [values[0]]
angles_loop = angles + [angles[0]]

fig, ax = plt.subplots(figsize=(9, 9), subplot_kw=dict(polar=True))
ax.fill(angles_loop, values_loop, color="#42A5F5", alpha=0.25)
ax.plot(angles_loop, values_loop, "o-", color="#1565C0", linewidth=2.5, markersize=9)
ax.set_xticks(angles)
ax.set_xticklabels(labels, fontsize=11)
ax.set_ylim(0, 5)
ax.set_yticks([1, 2, 3, 4, 5])
ax.set_yticklabels(["1", "2", "3", "4", "5"], fontsize=9, color="gray")
ax.set_title("Class Average — Proactiveness Dimensions\n(ממוצע כיתה – ממדי פעלנות | N=1,250)",
             fontsize=14, pad=28, fontweight="bold")
for i, (ang, val) in enumerate(zip(angles, values)):
    ax.annotate(f"{val:.2f}", xy=(ang, val), fontsize=11, fontweight="bold",
                ha="center", va="bottom", color="#1565C0",
                xytext=(ang, val + 0.22))
plt.tight_layout()
fig.savefig(f"{OUT}/01_radar_class_average.png", dpi=150, bbox_inches="tight")
plt.close()

# ══════════════════════════════════════════════════════════
#  GRAPH 02 – Heatmap: Level Distribution per Dimension
# ══════════════════════════════════════════════════════════
print("  📊 02 Heatmap – level distribution …")
hmap = np.array([[int((df[f"lvl_{d}"] == lvl).sum()) for lvl in range(1, 6)]
                 for d in DIM_CODES])

fig, ax = plt.subplots(figsize=(13, 7))
im = ax.imshow(hmap, cmap="YlOrRd", aspect="auto")
ax.set_xticks(range(5))
ax.set_xticklabels([f"Level {i}\n{LEVEL_EN[i]}" for i in range(1, 6)], fontsize=11)
ax.set_yticks(range(len(DIM_CODES)))
ax.set_yticklabels([RUBRIC_EN[d] for d in DIM_CODES], fontsize=11)
for i in range(len(DIM_CODES)):
    for j in range(5):
        val = hmap[i, j]
        pct = val / N * 100
        color = "white" if val > hmap.max() * 0.6 else "black"
        ax.text(j, i, f"{val}\n({pct:.1f}%)", ha="center", va="center",
                fontsize=10, fontweight="bold", color=color)
ax.set_title("Level Distribution per Dimension — Student Count & %\n"
             "(התפלגות רמות לפי ממד | N=1,250)",
             fontsize=14, fontweight="bold")
plt.colorbar(im, ax=ax, label="Student Count")
plt.tight_layout()
fig.savefig(f"{OUT}/02_level_heatmap.png", dpi=150, bbox_inches="tight")
plt.close()

# ══════════════════════════════════════════════════════════
#  GRAPH 03 – Pie: Overall Level Distribution
# ══════════════════════════════════════════════════════════
print("  📊 03 Pie – overall level …")
level_counts = df["lvl_total"].value_counts().sort_index()
labels_pie = [f"Level {lvl} – {LEVEL_HE[lvl]}\n({level_counts.get(lvl, 0)} students)"
              for lvl in range(1, 6)]
sizes = [level_counts.get(lvl, 0) for lvl in range(1, 6)]
colors_pie = [LEVEL_COLORS[lvl] for lvl in range(1, 6)]
explode = [0.04] * 5

fig, ax = plt.subplots(figsize=(9, 7))
wedges, texts, autotexts = ax.pie(
    sizes, labels=labels_pie, colors=colors_pie, explode=explode,
    autopct="%1.1f%%", startangle=140, textprops={"fontsize": 10},
    pctdistance=0.78,
)
for at in autotexts:
    at.set_fontweight("bold")
ax.set_title("Overall Proactiveness Level Distribution\n"
             "(התפלגות רמת פעלנות כללית | N=1,250)",
             fontsize=14, fontweight="bold")
plt.tight_layout()
fig.savefig(f"{OUT}/03_overall_level_pie.png", dpi=150, bbox_inches="tight")
plt.close()

# ══════════════════════════════════════════════════════════
#  GRAPH 04 – Bar: Dimension Means with Std Dev
# ══════════════════════════════════════════════════════════
print("  📊 04 Bar – dimension means …")
means = [df[f"avg_{d}"].mean() for d in DIM_CODES]
stds  = [df[f"avg_{d}"].std()  for d in DIM_CODES]
bar_colors = ["#42A5F5", "#66BB6A", "#FFA726", "#EF5350", "#AB47BC", "#26C6DA"]

fig, ax = plt.subplots(figsize=(13, 6))
bars = ax.bar([RUBRIC_EN[d] for d in DIM_CODES], means, yerr=stds,
              color=bar_colors, edgecolor="white", linewidth=0.8,
              capsize=5, error_kw={"linewidth": 1.5})
ax.set_ylim(0, 5.6)
ax.axhline(3, color="gray", linestyle="--", alpha=0.5, linewidth=1)
ax.set_ylabel("Average Score (1–5)", fontsize=12)
ax.set_title("Dimension Average Scores with Standard Deviation\n"
             "(ממוצע ציוני ממדים עם סטיית תקן | N=1,250)",
             fontsize=14, fontweight="bold")
ax.set_xticklabels([RUBRIC_EN[d] for d in DIM_CODES], rotation=14, ha="right", fontsize=10)
for bar, mean, std in zip(bars, means, stds):
    ax.text(bar.get_x() + bar.get_width() / 2, mean + std + 0.08,
            f"{mean:.2f}", ha="center", fontsize=11, fontweight="bold")
plt.tight_layout()
fig.savefig(f"{OUT}/04_dimension_means_std.png", dpi=150, bbox_inches="tight")
plt.close()

# ══════════════════════════════════════════════════════════
#  GRAPH 05 – Grouped Bar: Dimensions by Grade
# ══════════════════════════════════════════════════════════
print("  📊 05 Grouped bar – by grade …")
grade_colors = {"י": "#42A5F5", "יא": "#66BB6A", "יב": "#FFA726"}
x = np.arange(len(DIM_CODES))
width = 0.26

fig, ax = plt.subplots(figsize=(14, 7))
for idx, grade in enumerate(GRADES):
    sub = df[df["כיתה"] == grade]
    means_g = [sub[f"avg_{d}"].mean() for d in DIM_CODES]
    offset = (idx - 1) * width
    bars = ax.bar(x + offset, means_g, width, label=f"Grade {grade}",
                  color=grade_colors[grade], edgecolor="white")
    for bar, val in zip(bars, means_g):
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.04,
                f"{val:.2f}", ha="center", fontsize=8.5, fontweight="bold")

ax.set_xticks(x)
ax.set_xticklabels([RUBRIC_EN[d] for d in DIM_CODES], rotation=13, ha="right", fontsize=10)
ax.set_ylabel("Average Score (1–5)", fontsize=12)
ax.set_title("Dimension Scores by Grade\n(ציוני ממדים לפי כיתה | N=1,250)",
             fontsize=14, fontweight="bold")
ax.set_ylim(0, 5.5)
ax.axhline(3, color="gray", linestyle="--", alpha=0.3)
ax.legend(fontsize=11)
plt.tight_layout()
fig.savefig(f"{OUT}/05_dimensions_by_grade.png", dpi=150, bbox_inches="tight")
plt.close()

# ══════════════════════════════════════════════════════════
#  GRAPH 06 – Violin: Score Distribution per Dimension
# ══════════════════════════════════════════════════════════
print("  📊 06 Violin – distribution per dimension …")
data_for_violin = [df[f"avg_{d}"].values for d in DIM_CODES]

fig, ax = plt.subplots(figsize=(14, 7))
parts = ax.violinplot(data_for_violin, positions=range(1, 7),
                      showmedians=True, showextrema=True)
for i, (pc, color) in enumerate(zip(parts["bodies"],
        ["#42A5F5", "#66BB6A", "#FFA726", "#EF5350", "#AB47BC", "#26C6DA"])):
    pc.set_facecolor(color)
    pc.set_alpha(0.7)
parts["cmedians"].set_colors("black")
parts["cbars"].set_colors("gray")
parts["cmaxes"].set_colors("gray")
parts["cmins"].set_colors("gray")

ax.set_xticks(range(1, 7))
ax.set_xticklabels([RUBRIC_EN[d] for d in DIM_CODES], rotation=13, ha="right", fontsize=10)
ax.set_ylabel("Score (1–5)", fontsize=12)
ax.set_ylim(0.5, 5.5)
ax.axhline(3, color="gray", linestyle="--", alpha=0.4)
ax.set_title("Score Distribution per Dimension (Violin Plot)\n"
             "(התפלגות ציונים לפי ממד | N=1,250)",
             fontsize=14, fontweight="bold")
plt.tight_layout()
fig.savefig(f"{OUT}/06_violin_dimensions.png", dpi=150, bbox_inches="tight")
plt.close()

# ══════════════════════════════════════════════════════════
#  GRAPH 07 – AI Tool Usage: Mean Proactiveness per Tool
# ══════════════════════════════════════════════════════════
print("  📊 07 Bar – AI tool vs proactiveness …")
tool_means = (df.groupby("כלי AI")["avg_total"].mean()
                .reindex(AI_TOOLS).dropna())
tool_counts = df["כלי AI"].value_counts().reindex(AI_TOOLS).fillna(0)

fig, ax = plt.subplots(figsize=(11, 6))
bar_colors_tools = [TOOL_COLORS[t] for t in tool_means.index]
bars = ax.bar(tool_means.index, tool_means.values,
              color=bar_colors_tools, edgecolor="white", linewidth=0.8)
ax.axhline(df["avg_total"].mean(), color="#455A64", linestyle="--",
           linewidth=1.5, label=f"Overall mean = {df['avg_total'].mean():.2f}")
ax.set_ylim(0, 5)
ax.set_ylabel("Mean Proactiveness Score (1–5)", fontsize=12)
ax.set_title("Mean Proactiveness Score by AI Tool Used\n"
             "(ממוצע פעלנות לפי כלי AI | N=1,250)",
             fontsize=14, fontweight="bold")
for bar, val, tool in zip(bars, tool_means.values, tool_means.index):
    cnt = int(tool_counts.get(tool, 0))
    ax.text(bar.get_x() + bar.get_width() / 2, val + 0.06,
            f"{val:.2f}\n(n={cnt})", ha="center", fontsize=10, fontweight="bold")
ax.legend(fontsize=11)
plt.tight_layout()
fig.savefig(f"{OUT}/07_ai_tool_vs_proactiveness.png", dpi=150, bbox_inches="tight")
plt.close()

# ══════════════════════════════════════════════════════════
#  GRAPH 08 – Stacked Bar: Level Composition per Dimension
# ══════════════════════════════════════════════════════════
print("  📊 08 Stacked bar – level composition …")
dim_level_pct = {
    d: [float((df[f"lvl_{d}"] == lvl).sum() / N * 100) for lvl in range(1, 6)]
    for d in DIM_CODES
}
x_labels = [RUBRIC_EN[d] for d in DIM_CODES]
bottoms = np.zeros(len(DIM_CODES))

fig, ax = plt.subplots(figsize=(13, 7))
for lvl in range(1, 6):
    vals = [dim_level_pct[d][lvl - 1] for d in DIM_CODES]
    bars = ax.bar(x_labels, vals, bottom=bottoms,
                  color=LEVEL_COLORS[lvl], label=f"Level {lvl} – {LEVEL_EN[lvl]}",
                  edgecolor="white", linewidth=0.5)
    for bar, val, bot in zip(bars, vals, bottoms):
        if val > 5:
            ax.text(bar.get_x() + bar.get_width() / 2, bot + val / 2,
                    f"{val:.0f}%", ha="center", va="center",
                    fontsize=9, fontweight="bold", color="black")
    bottoms += vals

ax.set_ylim(0, 104)
ax.set_ylabel("% of Students", fontsize=12)
ax.set_title("Level Composition per Dimension — % of Students\n"
             "(הרכב רמות לפי ממד | N=1,250)",
             fontsize=14, fontweight="bold")
ax.set_xticklabels(x_labels, rotation=13, ha="right", fontsize=10)
ax.legend(loc="upper right", fontsize=10)
plt.tight_layout()
fig.savefig(f"{OUT}/08_stacked_levels.png", dpi=150, bbox_inches="tight")
plt.close()

# ══════════════════════════════════════════════════════════
#  GRAPH 09 – Scatter: Overall Score vs Initiative (coloured by Grade)
# ══════════════════════════════════════════════════════════
print("  📊 09 Scatter – total vs initiative by grade …")
fig, ax = plt.subplots(figsize=(10, 7))
for grade, color in grade_colors.items():
    sub = df[df["כיתה"] == grade]
    ax.scatter(sub["avg_initiative"], sub["avg_total"],
               c=color, alpha=0.35, s=18, label=f"Grade {grade}",
               edgecolors="none")
ax.set_xlabel("Initiative & Responsibility Score (1–5)", fontsize=12)
ax.set_ylabel("Overall Proactiveness Score (1–5)", fontsize=12)
ax.set_title("Overall Proactiveness vs. Initiative Score by Grade\n"
             "(פעלנות כללית מול יוזמה לפי כיתה | N=1,250)",
             fontsize=14, fontweight="bold")
ax.set_xlim(0.8, 5.2)
ax.set_ylim(0.8, 5.2)
ax.axhline(3, color="gray", linestyle="--", alpha=0.3)
ax.axvline(3, color="gray", linestyle="--", alpha=0.3)
ax.legend(fontsize=11)

# Regression line
m, b = np.polyfit(df["avg_initiative"], df["avg_total"], 1)
xline = np.linspace(1, 5, 100)
ax.plot(xline, m * xline + b, color="#B71C1C", linewidth=2, linestyle="-",
        label=f"Trend  y={m:.2f}x+{b:.2f}")
ax.legend(fontsize=11)
plt.tight_layout()
fig.savefig(f"{OUT}/09_scatter_initiative_vs_total.png", dpi=150, bbox_inches="tight")
plt.close()

# ══════════════════════════════════════════════════════════
#  GRAPH 10 – Box: Score Distribution by AI Tool per Dimension
# ══════════════════════════════════════════════════════════
print("  📊 10 Box – score by AI tool for each dimension …")
fig, axes = plt.subplots(2, 3, figsize=(18, 11))
axes = axes.flatten()

for idx, dim in enumerate(DIM_CODES):
    ax = axes[idx]
    data_by_tool = [df[df["כלי AI"] == t][f"avg_{dim}"].dropna().values
                    for t in AI_TOOLS]
    bp = ax.boxplot(data_by_tool, patch_artist=True, notch=False,
                    medianprops={"color": "black", "linewidth": 2})
    for patch, tool in zip(bp["boxes"], AI_TOOLS):
        patch.set_facecolor(TOOL_COLORS[tool])
        patch.set_alpha(0.75)
    ax.set_xticks(range(1, len(AI_TOOLS) + 1))
    ax.set_xticklabels(AI_TOOLS, fontsize=9, rotation=12)
    ax.set_ylim(0.5, 5.5)
    ax.set_title(f"{RUBRIC_EN[dim]}\n{RUBRIC_HE[dim]}", fontsize=10, fontweight="bold")
    ax.set_ylabel("Score (1–5)", fontsize=9)
    ax.axhline(3, color="gray", linestyle="--", alpha=0.4, linewidth=0.8)

fig.suptitle("Score Distribution per Dimension × AI Tool\n"
             "(התפלגות ציונים לפי ממד וכלי AI | N=1,250)",
             fontsize=15, fontweight="bold", y=1.01)
plt.tight_layout()
fig.savefig(f"{OUT}/10_boxplot_dimension_x_tool.png", dpi=150, bbox_inches="tight")
plt.close()

# ══════════════════════════════════════════════════════════
#  GRAPH 11 – Correlation Heatmap: Dimensions
# ══════════════════════════════════════════════════════════
print("  📊 11 Correlation heatmap …")
dim_cols = [f"avg_{d}" for d in DIM_CODES]
corr = df[dim_cols].corr()
labels_short = [RUBRIC_EN[d] for d in DIM_CODES]

fig, ax = plt.subplots(figsize=(10, 8))
im = ax.imshow(corr.values, cmap="coolwarm", vmin=-1, vmax=1)
ax.set_xticks(range(len(DIM_CODES)))
ax.set_yticks(range(len(DIM_CODES)))
ax.set_xticklabels(labels_short, rotation=25, ha="right", fontsize=10)
ax.set_yticklabels(labels_short, fontsize=10)
for i in range(len(DIM_CODES)):
    for j in range(len(DIM_CODES)):
        val = corr.values[i, j]
        color = "white" if abs(val) > 0.6 else "black"
        ax.text(j, i, f"{val:.2f}", ha="center", va="center",
                fontsize=11, fontweight="bold", color=color)
plt.colorbar(im, ax=ax, label="Pearson r")
ax.set_title("Dimension Correlation Matrix\n"
             "(מטריצת מתאמים בין ממדים | N=1,250)",
             fontsize=14, fontweight="bold")
plt.tight_layout()
fig.savefig(f"{OUT}/11_correlation_heatmap.png", dpi=150, bbox_inches="tight")
plt.close()

# ══════════════════════════════════════════════════════════
#  GRAPH 12 – Histogram: Overall Proactiveness Score Distribution
# ══════════════════════════════════════════════════════════
print("  📊 12 Histogram – overall score …")
fig, ax = plt.subplots(figsize=(11, 6))
n_bins = 30
counts, edges, patches = ax.hist(df["avg_total"], bins=n_bins,
                                  edgecolor="white", linewidth=0.5)
# Colour bars by level
for patch, left_edge in zip(patches, edges[:-1]):
    lvl = assign_level(left_edge + 0.01)
    patch.set_facecolor(LEVEL_COLORS[lvl])
    patch.set_alpha(0.85)

mean_val = df["avg_total"].mean()
median_val = df["avg_total"].median()
ax.axvline(mean_val, color="#1565C0", linewidth=2.5, linestyle="--",
           label=f"Mean = {mean_val:.2f}")
ax.axvline(median_val, color="#6A1B9A", linewidth=2, linestyle=":",
           label=f"Median = {median_val:.2f}")

# Legend for levels
level_patches = [mpatches.Patch(color=LEVEL_COLORS[lvl],
                 label=f"Level {lvl} – {LEVEL_HE[lvl]}") for lvl in range(1, 6)]
handles, _ = ax.get_legend_handles_labels()
ax.legend(handles=handles + level_patches, fontsize=9, loc="upper left")

ax.set_xlabel("Overall Proactiveness Score (1–5)", fontsize=12)
ax.set_ylabel("Number of Students", fontsize=12)
ax.set_title("Distribution of Overall Proactiveness Scores\n"
             "(התפלגות ציון פעלנות כללי | N=1,250)",
             fontsize=14, fontweight="bold")
plt.tight_layout()
fig.savefig(f"{OUT}/12_histogram_overall_score.png", dpi=150, bbox_inches="tight")
plt.close()

# ══════════════════════════════════════════════════════════
#  GRAPH 13 – Radar: AI Users vs Non-Users
# ══════════════════════════════════════════════════════════
print("  📊 13 Radar – AI users vs non-users …")
ai_users  = df[df["כלי AI"] != "לא משתמש/ת"]
non_users = df[df["כלי AI"] == "לא משתמש/ת"]
vals_ai   = [ai_users[f"avg_{d}"].mean() for d in DIM_CODES]
vals_no   = [non_users[f"avg_{d}"].mean() for d in DIM_CODES]
labels_r  = [RUBRIC_EN[d] for d in DIM_CODES]
angles_r  = np.linspace(0, 2 * np.pi, len(labels_r), endpoint=False).tolist()

fig, ax = plt.subplots(figsize=(9, 9), subplot_kw=dict(polar=True))
for vals, color, label in [
    (vals_ai + [vals_ai[0]], "#1565C0", f"AI Users (n={len(ai_users)})"),
    (vals_no + [vals_no[0]], "#B71C1C", f"Non-Users (n={len(non_users)})"),
]:
    ax.plot(angles_r + [angles_r[0]], vals, "o-", linewidth=2.5, label=label, color=color)
    ax.fill(angles_r + [angles_r[0]], vals, alpha=0.15, color=color)

ax.set_xticks(angles_r)
ax.set_xticklabels(labels_r, fontsize=11)
ax.set_ylim(0, 5)
ax.set_yticks([1, 2, 3, 4, 5])
ax.set_yticklabels(["1", "2", "3", "4", "5"], fontsize=9, color="gray")
ax.set_title("Proactiveness Profile: AI Users vs. Non-Users\n"
             "(פרופיל פעלנות: משתמשי AI מול ללא | N=1,250)",
             fontsize=14, pad=28, fontweight="bold")
ax.legend(loc="lower left", bbox_to_anchor=(-0.15, -0.15), fontsize=11)
plt.tight_layout()
fig.savefig(f"{OUT}/13_radar_ai_vs_non.png", dpi=150, bbox_inches="tight")
plt.close()

# ══════════════════════════════════════════════════════════
#  GRAPH 14 – Top/Bottom 10%: Strongest vs Weakest Profiles
# ══════════════════════════════════════════════════════════
print("  📊 14 Bar – top vs bottom 10% profiles …")
top10    = df.nlargest(int(N * 0.10), "avg_total")
bottom10 = df.nsmallest(int(N * 0.10), "avg_total")
means_top = [top10[f"avg_{d}"].mean() for d in DIM_CODES]
means_bot = [bottom10[f"avg_{d}"].mean() for d in DIM_CODES]
x = np.arange(len(DIM_CODES))
w = 0.37

fig, ax = plt.subplots(figsize=(13, 6))
b1 = ax.bar(x - w / 2, means_top, w, color="#1B5E20", alpha=0.82,
            label=f"Top 10% (n={len(top10)})", edgecolor="white")
b2 = ax.bar(x + w / 2, means_bot, w, color="#B71C1C", alpha=0.82,
            label=f"Bottom 10% (n={len(bottom10)})", edgecolor="white")
for bar, val in zip(list(b1) + list(b2), means_top + means_bot):
    ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.06,
            f"{val:.2f}", ha="center", fontsize=10, fontweight="bold")
ax.set_xticks(x)
ax.set_xticklabels([RUBRIC_EN[d] for d in DIM_CODES], rotation=13, ha="right", fontsize=10)
ax.set_ylim(0, 5.8)
ax.axhline(3, color="gray", linestyle="--", alpha=0.4)
ax.set_ylabel("Average Score (1–5)", fontsize=12)
ax.set_title("Dimension Profiles: Top 10% vs Bottom 10% Students\n"
             "(פרופיל ממדים: 10% עליון מול 10% תחתון | N=1,250)",
             fontsize=14, fontweight="bold")
ax.legend(fontsize=11)
plt.tight_layout()
fig.savefig(f"{OUT}/14_top_vs_bottom_10pct.png", dpi=150, bbox_inches="tight")
plt.close()

# ══════════════════════════════════════════════════════════
#  DONE – Summary
# ══════════════════════════════════════════════════════════
graphs = sorted([f for f in os.listdir(OUT) if f.endswith(".png")])
print(f"\n✅ {len(graphs)} graphs saved to {OUT}/")
for g in graphs:
    print(f"   {g}")

print("\n" + "=" * 70)
print("  Summary: N = 1,250 students | 6 dimensions | 14 visualizations")
print("=" * 70)
print(f"  Overall mean proactiveness : {df['avg_total'].mean():.2f}/5")
print(f"  Overall median             : {df['avg_total'].median():.2f}/5")
strongest = max(DIM_CODES, key=lambda d: df[f"avg_{d}"].mean())
weakest   = min(DIM_CODES, key=lambda d: df[f"avg_{d}"].mean())
print(f"  Strongest dimension        : {RUBRIC_EN[strongest]}")
print(f"  Weakest  dimension         : {RUBRIC_EN[weakest]}")
ai_m  = df[df["כלי AI"] != "לא משתמש/ת"]["avg_total"].mean()
no_m  = df[df["כלי AI"] == "לא משתמש/ת"]["avg_total"].mean()
print(f"  AI-user mean               : {ai_m:.2f}/5")
print(f"  Non-AI-user mean           : {no_m:.2f}/5")
print("=" * 70)
