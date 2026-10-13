# RMS KPI dashboards

Generate the University Health & Safety KPI dashboard and one dashboard for each faculty or service group from an RMS CSV export. The current set contains **10 dashboards**, with three reporting dates through **1 August 2026** (the February–July 2026 return).

[Open the current university dashboard](docs/index.html).

## Where things live

| Folder / file | Purpose |
| --- | --- |
| [`docs/`](docs/) | Current dashboards, ready for GitHub Pages. Start with `index.html`. |
| [`data/raw/exports/`](data/raw/exports/) | Dated CSV exports. Keep each new export here so its origin is clear. |
| [`data/raw/latest.csv`](data/raw/latest.csv) | Working CSV used when no input is specified. |
| [`data/processed/`](data/processed/) | Current `tidy.xlsx` and `summary.xlsx` workbooks. |
| [`scripts/`](scripts/) | Data processing and dashboard generation. |
| [`templates/`](templates/) | Dashboard appearance and behaviour. `dashboard_styles.css` and `dashboard_app.js` are shared by every dashboard and inlined at generation; the two HTML templates hold page-specific wording. |
| [`tests/`](tests/) | Checks for missing returns, zero scores and the refresh process. |
| [`archives/`](archives/) | Previous dashboards, data, legacy code and historical reference material. |
| `build/qa/` | Generated review copies, workbooks and logs. Local only; excluded from Git. |
| [`requirements.txt`](requirements.txt) | Python dependencies. |
| `index.html` | Compatibility landing page that redirects to `docs/index.html` when serving the whole repository. |

Edit the templates or scripts, then regenerate. Changes made directly to generated HTML will be overwritten by the next refresh.

Faculty dashboards have **Show results for** buttons for each school or department. Selecting one updates the scores, scope counts, trends, university comparisons and summary in both **Cards** and **Compact list**. The selection stays in place when changing view, sort order or performance-band filters. Select **Faculty total** to restore the combined results and school/department breakdowns.

The University overview generates a **Faculty dashboards** button for every faculty or service report. Faculty pages do not include a back link to the University overview. These controls are built by the generators and templates on every refresh.

## First-time setup

Install Python 3.10 or newer, clone or download this repository, and open a terminal in its main folder.

Create a local Python environment:

```bash
python3 -m venv .venv
```

Activate it on macOS / Linux:

```bash
source .venv/bin/activate
```

On Windows, use `py -m venv .venv` to create the environment, then activate it in PowerShell:

```powershell
.venv\Scripts\Activate.ps1
```

Install the dependencies:

```bash
python -m pip install -r requirements.txt
```

The standard refresh command works without a desktop or file-selection dialogs. No Node.js setup is needed. The dashboards load Chart.js from a CDN, so an internet connection is needed for the trend charts.

## Refresh the dashboards

### 1. Save the new export

Download the RMS CSV export and save it in `data/raw/exports/` with a dated filename, for example `2027-02-01.csv`. Keep the original RMS column names and include historical reporting periods if you want them in the trend charts. The corrected current export, [`2026-08-01.csv`](data/raw/exports/2026-08-01.csv), is an example of the required structure.

Keep a row for every school in each reporting period, including blank returns, so missing submissions can be shown. The pipeline cannot identify an omitted school/date row as a missing return. Dates use UK day/month/year format. School names must match the mapping in `scripts/data_transformer.py`; add new schools there before generating.

### 2. Generate a review copy

To rebuild from the current working CSV:

```bash
python scripts/refresh_dashboards.py
```

To review a new export, substitute its filename:

```bash
python scripts/refresh_dashboards.py --input data/raw/exports/2027-02-01.csv
```

Both commands write to `build/qa/` without changing the current dashboards or working data:

- `html/index.html` and all faculty/service reports;
- `source.csv`, a snapshot of the input used for this build;
- `tidy.xlsx` and `summary.xlsx`;
- `summary.json`, listing reporting dates, dashboard counts and missing returns;
- `pipeline.log`, with processing details.

A successful run replaces the previous review build. An unsuccessful build leaves the previous review build and current dashboards in place. Use `--build-dir` to choose another new review directory if needed.

### 3. Review locally

Start a preview server:

```bash
python -m http.server 8000 --directory build/qa/html
```

Open **http://localhost:8000/** in a browser. Check the reporting period, missing-return notices, faculty/school figures, tooltips and trends. Press `Ctrl+C` in the terminal to stop the server. You can also open `build/qa/html/index.html` directly in your browser.

### 4. Replace the current set

After reviewing, run the same generation command with `--promote`:

```bash
python scripts/refresh_dashboards.py --input data/raw/exports/2027-02-01.csv --promote
```

For the current working CSV, use:

```bash
python scripts/refresh_dashboards.py --promote
```

This rebuilds from the chosen input, then:

1. Backs up the current HTML dashboards and working CSV/workbooks to a new folder under `archives/`, named with a UTC timestamp. A checksum manifest records the backed-up files.
2. Copies the newly generated dashboards into `docs/`.
3. Updates `data/raw/latest.csv` and the two workbooks in `data/processed/` to match the build.

`--promote` updates local files only. It does not commit, push or change GitHub settings. To preview the current set, use `python -m http.server 8000 --directory docs`.

## Publish with GitHub Pages

The publishing folder is now **`docs/`**. After committing and pushing this reorganisation, set the repository's Pages source once:

1. Open the GitHub repository's **Settings → Pages**.
2. Under **Build and deployment**, select **Deploy from a branch**.
3. Select the **main** branch and **/docs** folder, then **Save**.

GitHub Pages serves `docs/index.html` at the site's home address, so the existing dashboard filenames and public paths can remain the same. `docs/.nojekyll` marks this as a plain static site. See [GitHub's publishing-source instructions](https://docs.github.com/en/pages/getting-started-with-github-pages/configuring-a-publishing-source-for-your-github-pages-site).

For later refreshes, review and promote locally, then commit and push the updated `docs/`, `data/` and new archive, together with any script/template changes. GitHub Pages will publish the committed dashboard files; it does not run the Python pipeline. There is no CI generation workflow in this repository.

The `.gitignore` excludes review builds, Python environments and caches, editor settings, macOS metadata and local environment files. Source data, current workbooks and historical archives remain part of the repository.

## How the pipeline works

| Script | Role |
| --- | --- |
| [`refresh_dashboards.py`](scripts/refresh_dashboards.py) | Runs the complete pipeline, creates review output and optionally archives/replaces the current set. |
| [`data_transformer.py`](scripts/data_transformer.py) | Converts RMS CSV columns into school-level KPIs, maps schools to faculties and applies period-specific training interpretation. |
| [`data_splitter.py`](scripts/data_splitter.py) | Produces latest-period summaries, reporting history and dashboard tooltips. |
| [`faculty_dashboard.py`](scripts/faculty_dashboard.py) | Generates each faculty/service dashboard from the summary workbook and faculty template. |
| [`university_dashboard.py`](scripts/university_dashboard.py) | Generates the university dashboard using the same KPI model. |

`tidy.xlsx` contains the school returns and question tooltips. `summary.xlsx` contains seven sheets: `University_Summary`, `Faculty_Summary`, `School_Raw_Data`, the three corresponding `_History` sheets, and `Question Tooltips`.

The individual faculty, university and splitter scripts still support their older interactive workflows, which require Tkinter. Use the refresh command for routine reporting.

## Reporting rules and current data

- The `Health and Safety` administrative row is excluded from reporting.
- Where raw counts are available, aggregate percentages are recalculated from summed counts. The estimate/coverage KPIs use the mean of reported school percentages, including genuine zeros and excluding missing values.
- Blank KPI data remains missing. A completely blank return displays **No return submitted**. An explicit zero is included in calculations; a zero denominator displays **Not Applicable** for the relevant count-based KPI.
- A completion flag does not discard a return that contains KPI data. Blank rows are retained to show missing submissions.
- Training fields are interpreted by reporting date: the latest-period H&S induction and fire-training fields are percentages, while earlier periods use counts.
- In the 1 August 2026 data, Libraries is marked complete. HR, Finance and Infrastructure, and Student & Public Facing Services have no KPI return. The university training-matrix coverage is **90.03%**, including genuine zero scores.
- Engineering's latest inspection totals are **55,865 scheduled** and **32,030 completed**, as confirmed from the source data.

## Verification and troubleshooting

Run the checks after changes to calculations, templates or the refresh process:

```bash
python -m unittest discover -s tests -v
```

If the refresh fails, read the error printed in the terminal. A successful build's detailed processing log is in `build/qa/pipeline.log`. Invalid dates, duplicate school/date rows and unmapped schools must be corrected before promotion. For a missing dependency, activate the environment and reinstall `requirements.txt`.

If charts do not load, check the browser's internet connection. If a GitHub Pages refresh is not visible, check that the updated `docs/` files were committed and pushed, and that Pages uses `main / docs`.

Older material is preserved inside dated archive folders; some historical files retain their original names. Archive manifests created before this reorganisation retain their original path references.
