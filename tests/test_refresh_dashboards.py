import hashlib
import json
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from scripts.refresh_dashboards import promote


class PromotionTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name) / 'repo'
        self.build = Path(self.temp.name) / 'build'
        self.old = {
            'docs/index.html': 'old university',
            'docs/Obsolete_Faculty_Report.html': 'old faculty',
            'data/raw/latest.csv': 'old CSV',
            'data/processed/tidy.xlsx': 'old tidy',
            'data/processed/summary.xlsx': 'old summary',
        }
        for relative, value in self.old.items():
            path = self.root / relative
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(value)
        for relative, value in {
            'html/index.html': 'new university',
            'html/New_Faculty_Report.html': 'new faculty',
            'source.csv': 'new CSV',
            'tidy.xlsx': 'new tidy',
            'summary.xlsx': 'new summary',
        }.items():
            path = self.build / relative
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(value)
        report = {'source_sha256': hashlib.sha256(b'new CSV').hexdigest(), 'dashboard_count': 2}
        (self.build / 'summary.json').write_text(json.dumps(report))

    def test_promotion_preserves_previous_site_and_data_with_checksums(self):
        archive = promote(self.build, self.root)
        manifest = json.loads((archive / 'manifest.json').read_text())
        for relative, value in self.old.items():
            self.assertEqual((archive / relative).read_text(), value)
            self.assertEqual(manifest[relative], hashlib.sha256(value.encode()).hexdigest())
        self.assertEqual((self.root / 'docs/index.html').read_text(), 'new university')
        self.assertEqual((self.root / 'data/raw/latest.csv').read_text(), 'new CSV')
        self.assertFalse((self.root / 'docs/Obsolete_Faculty_Report.html').exists())
        self.assertTrue((self.root / 'docs/.nojekyll').exists())

    def test_interrupted_promotion_restores_previous_site_and_data(self):
        original_copy = shutil.copy2

        def fail_new_csv(source, target, *args, **kwargs):
            if Path(source) == self.build / 'source.csv':
                raise OSError('Simulated copy failure')
            return original_copy(source, target, *args, **kwargs)

        with patch('scripts.refresh_dashboards.shutil.copy2', side_effect=fail_new_csv):
            with self.assertRaisesRegex(OSError, 'Simulated copy failure'):
                promote(self.build, self.root)
        for relative, value in self.old.items():
            self.assertEqual((self.root / relative).read_text(), value)
        self.assertFalse((self.root / 'docs/New_Faculty_Report.html').exists())

    def test_changed_source_is_rejected_before_current_files_are_touched(self):
        (self.build / 'source.csv').write_text('changed CSV')
        with self.assertRaisesRegex(ValueError, 'source snapshot has changed'):
            promote(self.build, self.root)
        self.assertFalse((self.root / 'archives').exists())
        for relative, value in self.old.items():
            self.assertEqual((self.root / relative).read_text(), value)


class RefreshIntegrationTests(unittest.TestCase):
    def test_headless_refresh_from_another_working_directory(self):
        with tempfile.TemporaryDirectory() as temporary:
            build = Path(temporary) / 'review'
            result = subprocess.run(
                [sys.executable, '-B', str(ROOT / 'scripts/refresh_dashboards.py'),
                 '--build-dir', str(build)],
                cwd=temporary, capture_output=True, text=True,
            )
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            report = json.loads((build / 'summary.json').read_text())
            self.assertEqual(len(list((build / 'html').glob('*.html'))), report['dashboard_count'])
            self.assertEqual((build / 'source.csv').read_bytes(), (ROOT / 'data/raw/latest.csv').read_bytes())
            self.assertTrue((build / 'html/index.html').is_file())
            self.assertTrue((build / 'pipeline.log').is_file())
            with pd.ExcelFile(build / 'summary.xlsx') as workbook:
                self.assertEqual(len(workbook.sheet_names), 7)
                university = pd.read_excel(workbook, sheet_name='University_Summary_History')
            self.assertEqual(len(university), len(report['reporting_dates']))


if __name__ == '__main__':
    unittest.main()
