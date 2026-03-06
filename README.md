# RMS KPI Program

This project converts raw RMS KPI data into aggregated summaries and interactive HTML dashboards.

## What You Need

- Python 3.10+ (3.11+ recommended)
- Packages:
  - `pandas`
  - `numpy`
  - `openpyxl`
- A desktop environment (scripts use `tkinter` file dialogs)

Install dependencies:

```bash
python3 -m pip install pandas numpy openpyxl
```

## Files and Lifecycle

### Input

- `Raw Data.csv`
  - Raw school-level source data.

### Intermediate 1

- `Tidyed Data.xlsx`
  - Created by `data_transformer.py`
  - Expected sheets:
    - `Sheet 1 - Return_Structure_KPI`
    - `Question Tooltips`

### Intermediate 2

- `split data.xlsx` (or any filename you choose in the save dialog)
  - Created by `Data Splitter (from Return_Structure_KPI).py`
  - Expected sheets:
    - `University_Summary`
    - `Faculty_Summary`
    - `School_Raw_Data`
    - `Question Tooltips`

### Final Outputs

- Faculty dashboards (`*_Faculty_Report.html`)
  - Created by `faculty_dashboard_refactored.py`
  - One HTML file per faculty.
- University dashboard (`University_KPI_Report.html`)
  - Created by `university_dashboard_refactored.py` (or `university_dashboard.py`)

## How to Run (Start to Finish)

Run these scripts in order from this folder:

```bash
python3 data_transformer.py
python3 "Data Splitter (from Return_Structure_KPI).py"
python3 faculty_dashboard_refactored.py
python3 university_dashboard_refactored.py
```

### Step Details

1. `data_transformer.py`
   - Reads `Raw Data.csv`
   - Writes `Tidyed Data.xlsx`

2. `Data Splitter (from Return_Structure_KPI).py`
   - Select `Tidyed Data.xlsx` when prompted
   - Save output as `split data.xlsx` (or your chosen name)

3. `faculty_dashboard_refactored.py`
   - Select the aggregated workbook from step 2
   - Select an output folder
   - Generates one faculty HTML report per faculty

4. `university_dashboard_refactored.py`
   - Select the same aggregated workbook
   - Select output folder
   - Generates `University_KPI_Report.html`

## What Good Output Looks Like

### Aggregated Excel (`split data.xlsx`)

- Contains exactly these summary sheets:
  - `University_Summary` (1 row expected in normal use)
  - `Faculty_Summary` (multiple faculties)
  - `School_Raw_Data` (school-level rows)
  - `Question Tooltips`

### Faculty HTML Reports

- Filenames like:
  - `Arts_Faculty_Report.html`
  - `Engineering_Faculty_Report.html`
- Each report shows:
  - KPI cards
  - Faculty value vs university value
  - Expandable school breakdown bars
  - Tooltip help text (if `Question Tooltips` exists)

### University HTML Report

- Filename:
  - `University_KPI_Report.html`
- Shows:
  - University-level KPI cards
  - Color-coded performance states
  - Sortable KPI order

## Quick Troubleshooting

- `Input file not found`: confirm filename and run from this folder.
- `... sheet not found`: check you selected the correct workbook for that stage.
- Tooltips missing: ensure `Question Tooltips` sheet exists.
- No GUI/file dialog appears: run on a machine/session with desktop GUI support.

