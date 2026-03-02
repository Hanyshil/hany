# CLAUDE.md — AI Assistant Guide for `hany`

This file provides context and conventions for AI assistants (Claude, Copilot, etc.) working in this repository.

---

## Project Overview

An **educational data analysis tool** written in Python that examines:

1. **Student AI Usage Patterns** — How high-school students use AI tools (ChatGPT, Google Bard, Claude, Copilot, Gemini), tracked via survey data.
2. **Student Learner Proactiveness** — A 6-dimension, 5-level rubric-based assessment (Q50–Q59) of student proactivity.

The project produces 20 high-quality PNG visualizations from two Excel datasets of 25 students.

---

## Repository Structure

```
hany/
├── create_sample_data.py            # Generates student_ai_usage.xlsx (25 students, 11 cols)
├── simulate_q50_59.py               # Generates student_proactiveness_q50_59.xlsx
├── analyze_student_ai_usage.py      # Produces 10 PNG graphs → analysis_output/
├── analyze_proactiveness_q50_59.py  # Produces 10 PNG graphs → proactiveness_output/
├── student_ai_usage.xlsx            # Input: AI usage survey data
├── student_proactiveness_q50_59.xlsx# Input: proactiveness rubric responses
├── analysis_output/                 # Generated: AI usage visualizations (10 PNGs)
└── proactiveness_output/            # Generated: proactiveness visualizations (10 PNGs)
```

---

## Tech Stack

| Layer          | Technology                           |
|----------------|--------------------------------------|
| Language       | Python 3.x                           |
| Data I/O       | pandas + openpyxl (Excel read/write) |
| Numerics       | numpy                                |
| Visualization  | matplotlib                           |
| Data Format    | XLSX (input), PNG (output)           |

No package manager manifest (requirements.txt / pyproject.toml) exists. Install dependencies manually:

```bash
pip install pandas numpy matplotlib openpyxl
```

---

## Development Workflow

Scripts must be run in this order when regenerating everything from scratch:

```bash
# Step 1 — Generate AI usage data
python create_sample_data.py

# Step 2 — Generate proactiveness data
python simulate_q50_59.py

# Step 3 — Analyse AI usage → analysis_output/ (10 PNGs)
python analyze_student_ai_usage.py

# Step 4 — Analyse proactiveness → proactiveness_output/ (10 PNGs)
python analyze_proactiveness_q50_59.py
```

Steps 3 and 4 are independent of each other and can be run in any order once the XLSX files exist. Steps 1 and 2 are only needed to regenerate the data files; if the XLSX files already exist (and contain real survey data), skip them.

---

## Data Schema

### `student_ai_usage.xlsx` (11 columns)

| Column             | Description                                      |
|--------------------|--------------------------------------------------|
| name               | Student name (Hebrew)                            |
| grade              | School grade (9–12)                              |
| ai_tool            | Primary AI tool used                             |
| frequency          | Usage frequency (e.g., daily, weekly)            |
| purpose            | Main purpose (homework, research, creativity…)   |
| subject            | Primary subject (math, science, language arts…)  |
| weekly_hours       | Average hours per week                           |
| helpfulness        | Perceived helpfulness rating                     |
| concerns           | Student concerns about AI                        |
| score_before       | Self-assessed score before AI usage (0–100)      |
| score_after        | Self-assessed score after AI usage (0–100)       |

### `student_proactiveness_q50_59.xlsx` (12 columns)

| Column      | Description                                        |
|-------------|----------------------------------------------------|
| name        | Student name (Hebrew)                              |
| grade       | School grade (9–12)                                |
| ai_tool     | Primary AI tool used                               |
| Q50–Q59     | Rubric question responses (1–5 Likert scale)       |

---

## Proactiveness Rubric (6 Dimensions)

| # | Dimension (Hebrew)                   | Questions | English                       |
|---|--------------------------------------|-----------|-------------------------------|
| א | מוטיבציה ורלוונטיות                  | Q50, Q51  | Motivation & Relevance        |
| ב | תודעת צמיחה                          | Q52, Q53  | Growth Mindset                |
| ג | יוזמה ואחריות                        | Q54       | Initiative & Responsibility   |
| ד | ויסות עצמי                           | Q55, Q56  | Self-Regulation               |
| ה | מודעות עצמית                         | Q57, Q58  | Self-Awareness                |
| ו | תמיכה וחוויות רגשיות                 | Q59       | Support & Emotional Experience|

**Proficiency levels (1–5):**

| Level | Hebrew             | English     |
|-------|--------------------|-------------|
| 1     | בתחילת הדרך        | Beginning   |
| 2     | מתפתח/ת            | Developing  |
| 3     | מתקדם/ת            | Advancing   |
| 4     | מיומן/ת            | Proficient  |
| 5     | מומחה/ית           | Expert      |

---

## Code Conventions

### File Paths
- Scripts use **absolute paths** rooted at `/home/user/hany/`. Update these when moving the project to a different machine.
- Output directories are created automatically with `os.makedirs(OUTPUT_DIR, exist_ok=True)` and existing files are overwritten on each run.

### Language / Localisation
- All user-facing strings, comments, and column names are **bilingual (Hebrew + English)**.
- Matplotlib is configured for Hebrew/RTL rendering: `font.family: DejaVu Sans`, `axes.unicode_minus: False`.
- Output PNGs are exported at **150 DPI**.

### Script Structure
Each analysis script follows the same pipeline:

```
Load XLSX → Validate columns → Compute dimension scores
→ Assign proficiency levels → Detect patterns
→ Generate visualisations (10 PNGs) → Print statistical report
```

Section markers use `===== SECTION NAME =====` style headings for easy navigation.

### Naming
- Output PNG files are zero-padded and numbered (e.g., `01_radar_class_average.png`, `10_strongest_weakest.png`).
- Variables use `snake_case` throughout.

---

## Pattern Detection (Proactiveness)

`analyze_proactiveness_q50_59.py` automatically classifies each student into one of six patterns:

| Pattern | Description                                      |
|---------|--------------------------------------------------|
| A       | High Motivation + Low Self-Regulation            |
| B       | High Growth Mindset + Low Initiative             |
| C       | Low Self-Awareness + High Support                |
| D       | Balanced High Profile                            |
| E       | Balanced Low Profile                             |
| F       | High Variation Between Dimensions                |

---

## Testing & Quality

There are currently **no automated tests** and **no linting configuration**.

When adding new analysis logic:
- Validate outputs manually by inspecting printed statistics and generated graphs.
- Ensure that any new XLSX column is added to all relevant scripts (data generation, analysis, and this document).
- Do not add dependencies beyond the existing four (`pandas`, `numpy`, `matplotlib`, `openpyxl`) without updating this file.

---

## Git Conventions

- **Commit messages** should summarise the analysis component changed and list affected files (see existing commit history for style).
- **Branch naming:** feature branches follow the `claude/<description>-<id>` pattern.
- No `.gitignore` is configured; avoid committing large regenerated PNG files unless they represent a final deliverable.

---

## Extending the Project

- **Add more students:** Increase `n_students` in `create_sample_data.py` / `simulate_q50_59.py` and rerun.
- **Add new questions:** Extend the Q-column set in `simulate_q50_59.py`, update the dimension mapping dictionary in `analyze_proactiveness_q50_59.py`, and add visualisations as needed.
- **Real survey data:** Replace the generated XLSX files with actual survey exports; ensure column names match exactly.
- **requirements.txt:** If dependencies grow, add `pip freeze > requirements.txt` to the workflow.
