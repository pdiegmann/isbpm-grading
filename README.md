# ISBPM Grading Support Tool

A Python-based utility to automate and manage the grading process for Information Systems and Business Process Management (ISBPM) examinations. It parses student evaluation data from CSV files and generates a comprehensive, dynamically calculated Excel workbook adhering to the complex German School Grade System logic.

## 🎯 Features

- **Automated Grade Compilation**: Reads multiple CSV files per student (tasks and other criteria) to consolidate all grading data.
- **Native Excel Formulas**: Generates an Excel file where final grades and sub-totals are computed dynamically via embedded native Excel formulas (no static hardcoded values).
- **German Grading System Mapping**: Built-in `GradeMapping` logic to convert percentage scores (0-100%) into standard German university grades (1.0 to 5.0).
- **Structured Outputs**:
  - **Master Overview**: A high-level view showing all students, their total points, final percentages, and ultimate grades.
  - **Individual Student Sheets**: Highly detailed breakdown of scoring per task, contextual reasoning, constraints, and implications, matching the required grading dimensions.

## 📋 Prerequisites

- **Python 3.12** or higher.
- [**uv**](https://github.com/astral-sh/uv) (Extremely fast Python package installer and resolver).

Dependencies managed via `uv` include:
- `openpyxl`
- `pandas`
- `xlsxwriter`

## 🚀 Installation & Setup

1. **Clone the repository** and navigate to the project directory:
   ```bash
   cd /path/to/isbpm-grading
   ```

2. **Sync the environment** to ensure dependencies are installed via `uv`:
   ```bash
   uv sync
   ```

## 🛠️ Usage

The main entry point for the tool is `main.py`. You must provide an input directory containing the generated or exported CSV grading forms.

```bash
uv run main.py --input-dir path/to/csv/dir [--output results.xlsx]
```

### Command Line Arguments

- `--input-dir` (Required): The directory containing the `students.csv` file and individual student grading CSV files.
- `--output` (Optional): The file path for the resulting Excel workbook. Defaults to `grades_output.xlsx`.

### Example

Using the bundled generic grading examples:

```bash
uv run main.py --input-dir docs --output test_output.xlsx
```

## 📁 Expected Input Formats

The tool expects specific naming conventions and structures for the CSV files within the `--input-dir`.

### 1. `students.csv`
A central registry of all examined students. The name must be `students.csv` (or `example-students.csv` fallback). 
**Key Columns Required:** `Username`, `First name`, `Last name`, `Matriculation number`, `E-mail`.

*(Note: The `Username` acts as the primary key for looking up corresponding grading files).*

### 2. `{username}-tasks.csv`
Contains the grading scores for the individual practical tasks submitted by the student.
**Columns Required:** 
- `Task` (Task name/ID)
- Practical details: `practicalTaskCorrect`, `practicalTaskDetails`
- Approach dimension: `approachCorrect`, `approachConvincing`, `approachReferences`
- Situationality dimension: `situationalityCorrect`, `situationalityConvincing`, `situationalityReferences`
- Implications dimension: `implicationsCorrect`, `implicationsConvincing`, `implicationsReferences`

### 3. `{username}-other.csv`
Contains scores for broader qualitative evaluation categories like Formatting, Structure, Style, and overarching Solution Report assessments.
**Columns Required:** 
- `Category` (e.g., "Overall", "Solution Report")
- `Item` (e.g., "Formatting", "Approach - Correctness")
- `Score` (Numerical grade value)
- `Notes` (Textual feedback or reasoning)

## 📊 Grading Logic Overview

The underlying grading model enforces the following distribution:

1. **Overall Formalities (20%)**
   - Formatting (30%)
   - Structure (50%)
   - Style/Language (20%)
   - Assessed on a 0-2 scale.

2. **Practical Tasks (40%)**
   - Correctness (65% weight, scaled 0-2)
   - Level of Detail (35% weight, binary 0-1)

3. **Solution Report (40%)**
   - Represents the theoretical/analytical core.
   - **Approach (50%)**: Methodological choices (Correctness 50%, Convincingness 35%, References 15%).
   - **Situationality (25%)**: Contextualization within the business scenario (Correctness 50%, Convincingness 35%, References 15%).
   - **Implications (25%)**: Constructive conclusions (Correctness 50%, Convincingness 35%, References 15%).

Detailed mappings convert these accumulated multi-level scores into percentage sums, which evaluate against the final 1.0 - 5.0 grade map.

## 🤝 Contributing

When contributing to the codebase:
- Respect the `pyproject.toml` configuration and `uv.lock` dependency trees.
- The heavy lifting for spreadsheet layout is implemented in `excel_generator.py`. Any modifications to styling, formulas, or cell formats should happen there.
- The `parser.py` safely handles different delimiter (`"`, `;`, `,`) quirks inherent in system-exported CSVs.
