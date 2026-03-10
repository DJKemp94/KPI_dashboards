# RMS KPI Program

This folder contains the scripts, templates, working files, and generated HTML dashboards for the RMS KPI reporting process.

## Current Working Rules

- The faculty HTML files in the project root are the current live files used for GitHub Pages. Do not archive or overwrite them casually.
- The QA build files in `output/qa/html/` are the safe place to iterate on dashboard changes before replacing the live root HTML files.
- Older experiments, screenshots, bytecode, and Playwright scratch output have been moved into `Archive (Mar 2026 cleanup)/`.

## Requirements

- Python 3.10+ (3.11+ recommended)
- Packages:
  - `pandas`
  - `numpy`
  - `openpyxl`
- A desktop session that can open `tkinter` file dialogs

Install dependencies:

```bash
python3 -m pip install pandas numpy openpyxl
```

## Main Files

### Inputs

- `Raw Data.csv`
  - Working filename expected by [data_transformer.py](/Users/davidkemp/Desktop/RMS%20KPI%20Reports/KPI%20Program%20and%20Data%20copy/data_transformer.py)
- `Return_Structure_KPI-2.csv`
  - Example/source export in the same structure as `Raw Data.csv`

### Processing Scripts

- [data_transformer.py](/Users/davidkemp/Desktop/RMS%20KPI%20Reports/KPI%20Program%20and%20Data%20copy/data_transformer.py)
  - Converts the CSV export into `Tidyed Data.xlsx`
- [Data Splitter (from Return_Structure_KPI).py](/Users/davidkemp/Desktop/RMS%20KPI%20Reports/KPI%20Program%20and%20Data%20copy/Data%20Splitter%20(from%20Return_Structure_KPI).py)
  - Builds university, faculty, and school summary sheets
- [faculty_dashboard_refactored.py](/Users/davidkemp/Desktop/RMS%20KPI%20Reports/KPI%20Program%20and%20Data%20copy/faculty_dashboard_refactored.py)
  - Generates one faculty dashboard per faculty from the refactored template
- [university_dashboard_refactored.py](/Users/davidkemp/Desktop/RMS%20KPI%20Reports/KPI%20Program%20and%20Data%20copy/university_dashboard_refactored.py)
  - Generates the university dashboard

### Templates

- [faculty_dashboard_template.html](/Users/davidkemp/Desktop/RMS%20KPI%20Reports/KPI%20Program%20and%20Data%20copy/faculty_dashboard_template.html)
- [university_dashboard_template.html](/Users/davidkemp/Desktop/RMS%20KPI%20Reports/KPI%20Program%20and%20Data%20copy/university_dashboard_template.html)

### Generated Outputs

- Root faculty dashboards such as `Arts_Faculty_Report.html`
  - Current live GitHub Pages versions
- `index.html`
  - Current live university landing/dashboard page
- `output/qa/html/`
  - QA dashboard builds for iteration and review

## End-to-End Process

This is the recommended process from source export through to refreshed dashboards.

### 1. Download the source data from SharePoint

Download the latest RMS export from SharePoint as a CSV.

- If you are following the standard pipeline, save it into this folder as `Raw Data.csv`.
- If your starting file is already named something like `Return_Structure_KPI-2.csv`, either:
  - rename/copy it to `Raw Data.csv`, or
  - keep it as reference and replace `Raw Data.csv` with the same content before running the scripts.

The working assumption in the scripts is that the starting CSV has columns like:

- `School`
- `Return_Start_Date`
- `Return End Date`
- `Date`
- KPI source columns such as `Arrangements1`, `Risk Assessment1`, `H&S Training1`, etc.

`Return_Structure_KPI-2.csv` in this folder is an example of that starting structure.

### 2. Run the CSV-to-workbook transformer

Run:

```bash
python3 data_transformer.py
```

What this does:

- Reads `Raw Data.csv`
- Converts the source columns into the structured KPI workbook format
- Writes `Tidyed Data.xlsx`
- Adds a `Question Tooltips` sheet used by the dashboards

Expected output:

- `Tidyed Data.xlsx`

Key sheet created:

- `Sheet 1 - Return_Structure_KPI`

### 3. Run the splitter/aggregator

Run:

```bash
python3 "Data Splitter (from Return_Structure_KPI).py"
```

When prompted:

1. Select `Tidyed Data.xlsx`
2. Choose where to save the aggregated workbook
3. Save it as `split data.xlsx` or another clearly named workbook

What this does:

- Reads `Sheet 1 - Return_Structure_KPI`
- Aggregates school rows into faculty and university summaries
- Preserves reporting periods and history sheets where available
- Writes a workbook used by the HTML generators

Expected sheets in the output workbook:

- `University_Summary`
- `Faculty_Summary`
- `School_Raw_Data`
- `Question Tooltips`

Possible additional history sheets:

- `University_Summary_History`
- `Faculty_Summary_History`
- `School_Raw_Data_History`

### 4. Generate faculty dashboards

Run:

```bash
python3 faculty_dashboard_refactored.py
```

When prompted:

1. Select the aggregated workbook from step 3
2. Select the output folder

Recommended output folder for safe iteration:

- `output/qa/html/`

What this does:

- Uses [faculty_dashboard_template.html](/Users/davidkemp/Desktop/RMS%20KPI%20Reports/KPI%20Program%20and%20Data%20copy/faculty_dashboard_template.html)
- Creates one HTML dashboard per faculty
- Includes compressed view, trends, tooltips, and school breakdowns

Example outputs:

- `Arts_Faculty_Report.html`
- `Engineering_Faculty_Report.html`
- `Science_Faculty_Report.html`

### 5. Generate the university dashboard

Run:

```bash
python3 university_dashboard_refactored.py
```

When prompted:

1. Select the same aggregated workbook used in step 4
2. Select the output folder

Recommended output folder for safe iteration:

- `output/qa/html/`

Expected output:

- `University_KPI_Report.html`

### 6. Review QA output before replacing live files

Review the generated files in:

- `output/qa/html/`

Check:

- KPI values and counts
- Faculty vs university comparison values
- Compressed view rendering
- Tooltip content
- Expanded school breakdowns
- Any KPIs showing no return / not applicable / exceeded schedule states

### 7. Promote QA files to the live root only when ready

Once you are satisfied with the QA versions:

- copy the reviewed faculty HTML files from `output/qa/html/` into the project root
- replace the existing live root HTML files only at that point

Do not treat the root files as scratch files. They are the current GitHub Pages-backed versions.

## Quick Paths

### Standard full refresh

```bash
python3 data_transformer.py
python3 "Data Splitter (from Return_Structure_KPI).py"
python3 faculty_dashboard_refactored.py
python3 university_dashboard_refactored.py
```

### If your starting point is `Return_Structure_KPI-2.csv`

1. Replace or copy it to `Raw Data.csv`
2. Run the standard full refresh above

## Expected Outputs

### `Tidyed Data.xlsx`

Should contain:

- `Sheet 1 - Return_Structure_KPI`
- `Question Tooltips`

### Aggregated workbook such as `split data.xlsx`

Should contain:

- `University_Summary`
- `Faculty_Summary`
- `School_Raw_Data`
- `Question Tooltips`

### Faculty HTML dashboards

Should show:

- KPI cards for each faculty metric
- Comparison against university values
- Expandable school detail
- Trend charts where history is available
- Tooltip help text

### University HTML dashboard

Should show:

- university-level KPI cards
- sorting controls
- tooltip help text

## Troubleshooting

- `Input file 'Raw Data.csv' not found`
  - Put the latest SharePoint export in this folder and name it `Raw Data.csv`
- `Sheet 1 - Return_Structure_KPI sheet not found`
  - You selected the wrong workbook in the splitter step
- Tooltips missing
  - Confirm the workbook contains `Question Tooltips`
- File dialog does not appear
  - Run the scripts in a desktop session with GUI support
- Dashboard output looks old
  - Confirm you generated into `output/qa/html/` and opened the refreshed file, not an older root HTML file

## Folder Notes

- `Archive (Mar 2026 cleanup)/`
  - old dashboards, screenshots, bytecode, and Playwright test output archived during repo cleanup
- `Archive (Sept 2025`
  - earlier archive retained as-is
- `output/qa/`
  - current QA workspace
