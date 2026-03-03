# CLAUDE.md — Project Guide for AI Assistants

## Project Overview

This is an **educational research analytics project** that analyzes:
1. **Student AI tool usage patterns** — how 25 high school students (grades 9–12) use AI tools
2. **Learner proactiveness assessment** — evaluating students across 6 pedagogical dimensions

The project is fully **bilingual (Hebrew/English)**. Hebrew is used for column names, chart labels, and print output; English is used for variable names, code logic, and technical comments.

---

## Repository Structure

```
hany/
├── create_sample_data.py               # Step 1: Generate baseline student Excel data
├── simulate_q50_59.py                  # Step 2: Simulate proactiveness survey responses
├── analyze_student_ai_usage.py         # Step 3a: Analyze AI usage patterns → analysis_output/
├── analyze_proactiveness_q50_59.py     # Step 3b: Analyze proactiveness dimensions → proactiveness_output/
│
├── student_ai_usage.xlsx               # Input data: 25 students, AI usage survey
├── student_proactiveness_q50_59.xlsx   # Input data: proactiveness scores (Q50–Q59)
│
├── analysis_output/                    # Output: 10 PNG charts from AI usage analysis
└── proactiveness_output/               # Output: 10 PNG charts from proactiveness analysis
```

---

## Script Execution Order

Scripts are designed to run sequentially. Later scripts depend on files produced by earlier ones:

```bash
python create_sample_data.py           # Creates student_ai_usage.xlsx
python simulate_q50_59.py              # Creates student_proactiveness_q50_59.xlsx
python analyze_student_ai_usage.py     # Reads student_ai_usage.xlsx → analysis_output/*.png
python analyze_proactiveness_q50_59.py # Reads student_proactiveness_q50_59.xlsx → proactiveness_output/*.png
```

Run scripts from the project root (`/home/user/hany/`).

---

## Dependencies

No `requirements.txt` exists. The required packages are:

| Package      | Purpose                        |
|--------------|-------------------------------|
| `pandas`     | Excel I/O, data manipulation   |
| `numpy`      | Numerical operations           |
| `matplotlib` | Chart generation (PNG output)  |
| `openpyxl`   | Excel engine for pandas        |

Install with:
```bash
pip install pandas numpy matplotlib openpyxl
```

---

## Hardcoded Paths

All scripts use **absolute hardcoded paths** pointing to `/home/user/hany/`. There is no path configuration file. When modifying scripts, ensure paths remain consistent:

- Input files: `/home/user/hany/student_ai_usage.xlsx`, `/home/user/hany/student_proactiveness_q50_59.xlsx`
- Output dirs: `/home/user/hany/analysis_output/`, `/home/user/hany/proactiveness_output/`

Output directories are auto-created via `os.makedirs(..., exist_ok=True)`.

---

## Script Descriptions

### `create_sample_data.py`
Generates synthetic baseline data for 25 students (grades 9–12). Uses `random.seed(42)` for reproducibility. Creates `student_ai_usage.xlsx` with columns in Hebrew (כיתה, כלי AI בשימוש, תדירות שימוש, etc.).

### `simulate_q50_59.py`
Reads `student_ai_usage.xlsx` and generates realistic Q50–Q59 proactiveness survey responses based on 5 student personality archetypes (tech_enthusiast, academic_focused, social_learner, struggling_learner, balanced_learner). Uses `random.seed(42)`. Outputs `student_proactiveness_q50_59.xlsx`.

### `analyze_student_ai_usage.py`
Reads `student_ai_usage.xlsx` and produces 10 PNG charts analyzing:
- AI tool distribution, usage frequency, helpfulness ratings
- Impact on self-scores by grade and tool
- Correlation between usage frequency and perceived impact

### `analyze_proactiveness_q50_59.py`
Reads `student_proactiveness_q50_59.xlsx` and analyzes 6 learner dimensions across Q50–Q59:
- **Q50, Q51** — Motivation & Relevance (מוטיבציה ורלוונטיות)
- **Q52, Q53** — Growth Mindset (תודעת צמיחה)
- **Q54**       — Initiative & Responsibility (יוזמה ואחריות)
- **Q55, Q56** — Self-Regulation (ויסות עצמי)
- **Q57, Q58** — Self-Awareness (מודעות עצמית)
- **Q59**       — Support & Emotions (תמיכה וחוויות רגשיות)

Identifies 6 behavioral patterns and generates 10 PNG charts plus statistical reports.

---

## Code Conventions

### Language
- **Variable/function names**: English (snake_case)
- **Print output and chart labels**: Bilingual (Hebrew first, English second where applicable)
- **Comments**: Mix of Hebrew and English; prefer English for logic explanations

### Hebrew/RTL Support
All scripts set matplotlib font config at the top:
```python
matplotlib.rcParams['font.family'] = 'DejaVu Sans'
matplotlib.rcParams['axes.unicode_minus'] = False
```
Do not remove these lines — they are required for Hebrew text rendering in charts.

### Data Access
Excel columns are accessed by their **Hebrew column names** (e.g., `df["כיתה"]`, `df["כלי AI בשימוש"]`). When adding new columns or modifying data access, preserve exact Hebrew column names from the source Excel files.

### Reproducibility
`create_sample_data.py` and `simulate_q50_59.py` use `random.seed(42)`. Do not change seed values unless intentionally regenerating different data.

### Output
- Charts are saved as PNG files with descriptive English filenames (e.g., `ai_tools_distribution.png`)
- Console output uses `"=" * 60` dividers to separate sections
- Both Hebrew and English section headers are printed

---

## No Tests, No CI/CD

This project has no test suite, linting configuration, or CI/CD pipeline. When modifying scripts:
- Verify outputs by running the script and checking console output
- Check the relevant output directory for generated PNG files
- Validate that Excel files are readable and contain expected columns

---

## Git Branch

Active development branch: `claude/claude-md-mmak6gc0k84xk75u-7XKSa`

Branch naming follows the pattern: `claude/<task-slug>-<session-id>`

---

## Common Tasks

**Regenerate all data and charts from scratch:**
```bash
python create_sample_data.py
python simulate_q50_59.py
python analyze_student_ai_usage.py
python analyze_proactiveness_q50_59.py
```

**Add a new analysis chart to AI usage analysis:**
- Add chart generation code to `analyze_student_ai_usage.py`
- Save output to `OUTPUT_DIR` with `plt.savefig(os.path.join(OUTPUT_DIR, "filename.png"), ...)`
- Call `plt.close()` after saving to free memory

**Add a new dimension or question mapping:**
- Update `dimension_map` dict in `analyze_proactiveness_q50_59.py`
- Update `dimension_names_he` and `dimension_names_en` dicts
- Update the `q_cols` filter if the new questions have a different prefix
