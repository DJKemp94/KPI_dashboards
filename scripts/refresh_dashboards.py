"""Build a review copy, optionally archive and replace the current dashboards."""

import argparse
import contextlib
from datetime import datetime, timezone
import hashlib
import json
from pathlib import Path
import shutil
import sys
import tempfile

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
if not __package__:
    sys.path.insert(0, str(ROOT))

from scripts import data_splitter as splitter
from scripts.data_transformer import KPIDataTransformer
from scripts.faculty_dashboard import FacultyDashboardGenerator
from scripts.university_dashboard import UniversityDashboardGenerator


def sha256(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def build_dashboards(source, destination):
    """Build into a new directory; leave the published site untouched."""
    destination.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, destination / 'source.csv')
    source = destination / 'source.csv'
    tidy = destination / 'tidy.xlsx'
    summary = destination / 'summary.xlsx'
    html = destination / 'html'
    html.mkdir()
    transformer = KPIDataTransformer()
    transformer.transform_data(source, tidy)
    rows = splitter.load_kpi_sheet(tidy)
    if rows.empty or rows['Date'].isna().any():
        raise ValueError('The export must contain school returns with valid reporting dates.')
    unknown = rows.loc[rows['Faculty'] == 'Unknown', 'School'].unique().tolist()
    if unknown:
        raise ValueError(f'Add these schools to the faculty mapping first: {unknown}')
    if rows.duplicated(['School', 'Date']).any():
        raise ValueError('The export contains duplicate school/reporting-date rows.')

    periods = splitter.get_available_reporting_periods(tidy)
    faculties = [splitter.aggregate_school_to_faculty_data(tidy, period) for period in periods]
    universities = [splitter.aggregate_school_to_university_data(tidy, period) for period in periods]
    schools = [splitter.create_school_raw_data(tidy, period, include_date=True) for period in periods]
    sheets = {
        'University_Summary': splitter.filter_and_order_columns(universities[-1]),
        'Faculty_Summary': splitter.filter_and_order_columns(faculties[-1]),
        'School_Raw_Data': splitter.create_school_raw_data(tidy, periods[-1]),
        'University_Summary_History': splitter.filter_and_order_columns(
            pd.concat(universities, ignore_index=True), include_date=True),
        'Faculty_Summary_History': splitter.filter_and_order_columns(
            pd.concat(faculties, ignore_index=True), include_date=True),
        'School_Raw_Data_History': pd.concat(schools, ignore_index=True),
        'Question Tooltips': splitter.create_tooltips_dataframe(),
    }
    with pd.ExcelWriter(summary, engine='openpyxl') as writer:
        for name, frame in sheets.items():
            frame.to_excel(writer, sheet_name=name, index=False)

    faculty = FacultyDashboardGenerator()
    university = UniversityDashboardGenerator()
    if not faculty.load_excel_data(summary) or not university.load_excel_data(summary):
        raise ValueError('The generated summary workbook could not be loaded.')
    data = faculty.extract_and_process_data()
    university_data = university._extract_university_dashboard_data()
    for name in data['faculties']:
        safe_name = name.replace(' ', '_').replace('&', '').replace(',', '')
        faculty.create_faculty_html_dashboard(data, name, html / f'{safe_name}_Faculty_Report.html')
    university.create_university_html_dashboard(university_data, html / 'index.html')
    (html / '.nojekyll').touch()

    latest = schools[-1]
    report = {
        'source_sha256': sha256(source),
        'reporting_dates': [splitter.format_reporting_date(period) for period in periods],
        'faculty_count': len(data['faculties']),
        'school_count': len(latest),
        'dashboard_count': len(data['faculties']) + 1,
        'missing_returns': latest.loc[~latest['Has KPI data'], 'School'].tolist(),
    }
    (destination / 'summary.json').write_text(json.dumps(report, indent=2) + '\n', encoding='utf-8')
    return report


def promote(build, root=ROOT):
    """Archive the old files before replacing them; restore them if a copy fails."""
    root = Path(root)
    build = Path(build)
    report = json.loads((build / 'summary.json').read_text(encoding='utf-8'))
    if report['source_sha256'] != sha256(build / 'source.csv'):
        raise ValueError('The source snapshot has changed since generation.')
    pages = sorted((build / 'html').glob('*.html'))
    if len(pages) != report['dashboard_count'] or not (build / 'html' / 'index.html').is_file():
        raise ValueError('The dashboard build is incomplete.')

    stamp = datetime.now(timezone.utc).strftime('%Y-%m-%d_%H%M%S_%f-UTC')
    archive = root / 'archives' / stamp
    archive.mkdir(parents=True)
    site = root / 'docs'
    site.mkdir(exist_ok=True)
    replacements = [(page, site / page.name) for page in pages]
    replacements.extend([
        (build / 'source.csv', root / 'data/raw/latest.csv'),
        (build / 'tidy.xlsx', root / 'data/processed/tidy.xlsx'),
        (build / 'summary.xlsx', root / 'data/processed/summary.xlsx'),
    ])
    # Include obsolete faculty pages so a later change of faculty names cannot
    # leave stale dashboards in the current site.
    old_paths = set(site.glob('*.html')) | {target for _, target in replacements}
    existing = {path for path in old_paths if path.is_file()}
    manifest = {}
    for path in sorted(existing):
        relative = path.relative_to(root)
        backup = archive / relative
        backup.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(path, backup)
        manifest[relative.as_posix()] = sha256(backup)
    (archive / 'manifest.json').write_text(json.dumps(manifest, indent=2) + '\n', encoding='utf-8')
    try:
        for source, target in replacements:
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(source, target)
        for obsolete in set(site.glob('*.html')) - {target for _, target in replacements}:
            obsolete.unlink()
        (site / '.nojekyll').touch()
    except Exception:
        for path in old_paths:
            if path in existing:
                shutil.copy2(archive / path.relative_to(root), path)
            elif path.exists():
                path.unlink()
        raise
    return archive


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--input', type=Path, default=ROOT / 'data/raw/latest.csv',
                        help='CSV export to use (default: data/raw/latest.csv).')
    parser.add_argument('--build-dir', type=Path, default=ROOT / 'build/qa',
                        help='Review output directory (default: build/qa).')
    parser.add_argument('--promote', action='store_true',
                        help='Archive and replace docs/ and working data after a successful build.')
    args = parser.parse_args(argv)
    source = args.input.resolve()
    build = args.build_dir.resolve()
    # Limit replacement of an existing build directory to directories created
    # by this command (or the repository's designated local QA directory).
    if build == ROOT / 'build/qa' or (build / '.kpi-build').is_file() or not build.exists():
        pass
    else:
        parser.error('Choose a new --build-dir; the existing directory is not a KPI build.')
    if source == build or build in source.parents:
        parser.error('The input CSV must be outside the build directory.')
    if build == ROOT or build in ROOT.parents or (ROOT in build.parents and build.parts[len(ROOT.parts)] != 'build'):
        parser.error('Use a directory under build/ or a new directory outside the project.')
    if not source.is_file():
        parser.error(f'Input CSV not found: {source}')
    build.parent.mkdir(parents=True, exist_ok=True)
    try:
        with tempfile.TemporaryDirectory(prefix='.kpi-', dir=build.parent) as temp:
            stage = Path(temp) / 'result'
            stage.mkdir()
            with (stage / 'pipeline.log').open('w', encoding='utf-8') as log, contextlib.redirect_stdout(log):
                report = build_dashboards(source, stage)
            (stage / '.kpi-build').touch()
            if build.exists():
                shutil.rmtree(build)
            shutil.move(str(stage), build)
        print(f"Generated {report['dashboard_count']} dashboards for {report['reporting_dates'][-1]}.")
        print(f"Review: {build / 'html/index.html'}")
        print('Missing returns: ' + (', '.join(report['missing_returns']) or 'none'))
        if args.promote:
            archive = promote(build)
            print(f'Previous dashboards and working data archived: {archive}')
            print(f'Current dashboards updated: {ROOT / "docs"}')
        return 0
    except Exception as exc:
        print(f'Refresh failed: {exc}', file=sys.stderr)
        return 1


if __name__ == '__main__':
    raise SystemExit(main())
