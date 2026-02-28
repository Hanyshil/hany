# CLAUDE.md — AI Assistant Guide for This Repository

## Project Overview

This is a **Python data analysis and visualization project** focused on student learning metrics in an educational research context. It analyzes two interrelated datasets:

1. **Student AI Tool Usage** — how students use AI tools (ChatGPT, Claude, Bard, etc.) academically
2. **Learner Proactiveness** — six-dimension rubric-based assessment of self-directed learning behaviors (Q50–Q59)

All data is stored in Excel files; outputs are PNG charts and formatted console reports. There is no web server, database, or API layer.

---

## Repository Structure

```
/
├── create_sample_data.py              # Generate sample AI usage survey data → student_ai_usage.xlsx
├── simulate_q50_59.py                 # Generate proactiveness survey responses → student_proactiveness_q50_59.xlsx
├── analyze_student_ai_usage.py        # Analyze AI usage patterns → analysis_output/
├── analyze_proactiveness_q50_59.py    # Analyze proactiveness dimensions → proactiveness_output/
├── student_ai_usage.xlsx              # AI usage survey data (input)
├── student_proactiveness_q50_59.xlsx  # Proactiveness survey responses (input)
├── analysis_output/                   # Generated PNG charts from AI usage analysis
└── proactiveness_output/              # Generated PNG charts from proactiveness analysis
```

---

## Technology Stack

| Component         | Technology                          |
|-------------------|-------------------------------------|
| Language          | Python 3                            |
| Data processing   | pandas, numpy                       |
| Visualization     | matplotlib                          |
| Spreadsheet I/O   | openpyxl                            |
| Data generation   | random (seeded at 42 for reproducibility) |

There is no `requirements.txt` or `pyproject.toml`. Install dependencies manually:

```bash
pip install pandas numpy matplotlib openpyxl
```

---

## Running the Scripts

Run in this order when starting from scratch or regenerating all outputs:

```bash
# 1. Generate synthetic survey data
python create_sample_data.py         # → student_ai_usage.xlsx
python simulate_q50_59.py            # → student_proactiveness_q50_59.xlsx

# 2. Run analyses (reads the Excel files, writes charts + console output)
python analyze_student_ai_usage.py   # → analysis_output/*.png
python analyze_proactiveness_q50_59.py  # → proactiveness_output/*.png
```

Each analysis script is self-contained. If the Excel input files already exist (real survey data), skip steps 1 and 2 and run only the analysis scripts.

---

## Module Descriptions

### `create_sample_data.py`
- Generates 25 fictional students with realistic Hebrew names and academic attributes
- Randomizes AI tool usage, frequency, purpose, subject, weekly hours, helpfulness rating, and concerns
- Records before/after self-assessment scores to measure perceived improvement
- **Seed:** `random.seed(42)` for reproducibility

### `simulate_q50_59.py`
- Generates proactiveness survey responses for Q50–Q59
- Maps 10 questions to 6 dimensions: Motivation & Relevance, Growth Mindset, Initiative & Responsibility, Self-Regulation, Self-Awareness, Support & Emotions
- Scores on a 1–5 scale (1 = Beginning, 5 = Expert)
- Correlates proactiveness profiles with AI usage patterns from the first dataset
- **Seed:** `np.random.seed(42)`

### `analyze_student_ai_usage.py`
- Reads `student_ai_usage.xlsx`
- Produces 10 charts covering tool popularity, usage frequency, purposes, subjects, helpfulness, hours by tool, before/after scores, concerns, score changes per student, and usage by grade
- Charts saved as `analysis_output/01_*.png` through `analysis_output/10_*.png`

### `analyze_proactiveness_q50_59.py`
- Reads `student_proactiveness_q50_59.xlsx`
- Computes dimension statistics, overall proactiveness levels, per-student profiles, cross-dimension correlations, and grade-level comparisons
- Detects behavioral patterns (e.g., high motivation + low regulation)
- Produces 10 charts in `proactiveness_output/`

---

## Data Conventions

### Hebrew Column Headers
All Excel columns use Hebrew field names matching real survey terminology:

| Hebrew                 | Meaning                         |
|------------------------|---------------------------------|
| `שם התלמיד/ה`          | Student name                    |
| `כיתה`                 | Grade (ט–יב = 9–12)             |
| `כלי AI בשימוש`        | AI tools in use                 |
| `תדירות שימוש`         | Usage frequency                 |
| `מטרת השימוש`          | Purpose of use                  |
| `מקצוע`                | Subject                         |
| `שעות שבועיות`         | Weekly hours                    |
| `רמת עזרה נתפסת`       | Perceived helpfulness (1–5)     |
| `חששות`                | Concerns                        |
| `ציון עצמי לפני/אחרי` | Self-score before/after         |

### Survey Question Mapping (Q50–Q59)
| Questions | Dimension                    |
|-----------|------------------------------|
| Q50, Q51  | Motivation & Relevance       |
| Q52, Q53  | Growth Mindset               |
| Q54       | Initiative & Responsibility  |
| Q55, Q56  | Self-Regulation              |
| Q57, Q58  | Self-Awareness               |
| Q59       | Support & Emotions           |

### Proactiveness Level Scale
| Score Range | Level   |
|-------------|---------|
| 1.0–1.8     | Beginning (מתחיל) |
| 1.8–2.6     | Developing (מתפתח) |
| 2.6–3.4     | Progressing (מתקדם) |
| 3.4–4.2     | Proficient (מיומן) |
| 4.2–5.0     | Expert (מומחה) |

---

## Code Style Conventions

- **Variable names:** English snake_case (e.g., `dim_motivation`, `overall_proactiveness`)
- **Output file naming:** Zero-padded sequential numbering (`01_radar_class_average.png`, `02_level_heatmap.png`, ...)
- **Bilingual output:** All chart titles, labels, and console text use Hebrew and/or English. Match existing language style when editing.
- **Output directories:** Created with `os.makedirs(dir, exist_ok=True)` — do not assume they exist.
- **Console reporting:** Uses formatted print statements with Unicode box-drawing separators. Follow the same style for any new console output.
- **Random seeds:** Always use `random.seed(42)` and `np.random.seed(42)` in data generation scripts to keep results reproducible.
- **DPI:** All charts are saved at 150 DPI (`plt.savefig(..., dpi=150, bbox_inches='tight')`).
- **Font:** `matplotlib.rcParams['font.family'] = 'DejaVu Sans'` is set globally for Hebrew character support.

---

## Visualization Patterns

When adding new charts, follow these conventions:

1. Use predefined color lists (`colors = ['#FF6B6B', '#4ECDC4', ...]`) rather than default matplotlib colors
2. Add value annotations on top of bar charts
3. Use `figsize` appropriate to content (typical: `(12, 8)` for standard charts, `(14, 10)` for complex ones)
4. Set bilingual `title`, `xlabel`, `ylabel` where applicable
5. Save with `plt.savefig(os.path.join(output_dir, 'NN_name.png'), dpi=150, bbox_inches='tight')`
6. Always call `plt.close()` after saving to free memory

---

## Key Design Patterns

### ETL Pipeline
Each analysis script follows: **Extract** (read Excel) → **Transform** (compute/aggregate) → **Load** (write charts + print reports).

### Reproducible Data Generation
Both generation scripts seed Python's `random` module and NumPy's RNG to produce identical output on every run.

### Pattern Detection
`analyze_proactiveness_q50_59.py` identifies named behavioral archetypes (e.g., "High motivation, Low regulation") by comparing dimension scores. When adding new patterns, define a label, a detection condition, and a count threshold.

---

## What This Project Does NOT Have

- No web server or REST API
- No database (Excel files are the data store)
- No unit tests or testing framework
- No CI/CD pipeline
- No `requirements.txt`, `pyproject.toml`, or virtual environment configuration
- No `.gitignore` (Excel files and generated PNGs are committed)

---

## Suggested Improvements (for reference)

If asked to improve this project, consider these additions in order of value:

1. **`requirements.txt`** — list `pandas`, `numpy`, `matplotlib`, `openpyxl` with pinned versions
2. **`.gitignore`** — exclude `__pycache__/`, `*.pyc`, and optionally generated output directories
3. **`README.md`** — high-level overview for human readers
4. **Unit tests** — pytest tests for score computation and pattern detection logic
5. **CLI arguments** — allow passing custom input/output paths instead of hardcoded strings

Do not add any of these unless explicitly requested.
