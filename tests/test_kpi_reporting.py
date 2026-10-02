import contextlib
import io
import sys
import unittest
from pathlib import Path
from unittest.mock import patch

import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from scripts.faculty_dashboard import FacultyDashboardGenerator
from scripts import data_splitter as splitter


class KPIReportingTests(unittest.TestCase):
    def setUp(self):
        self.generator = FacultyDashboardGenerator()
        self.matrix = '% of Training identified in Matrix that is accessible'
        self.source = pd.DataFrame({
            'School': ['CLAS', 'English', 'Humanities', 'HR'],
            'Faculty': ['Arts', 'Arts', 'Arts', 'HR'],
            'Date': ['01/08/2026'] * 4,
            'Number of Staff': [10, 10, np.nan, np.nan],
            'Number of Arrangements': [1, 1, np.nan, np.nan],
            'Number of Arrangements Completed': [0, 1, np.nan, np.nan],
            self.matrix: ['0%', '100%', np.nan, np.nan],
        })

    def aggregate(self, function, source=None):
        with patch.object(pd, 'read_excel', return_value=self.source if source is None else source), contextlib.redirect_stdout(io.StringIO()):
            return function('unused.xlsx', reporting_period='01/08/2026')

    def test_reported_zero_is_included_and_missing_is_excluded_at_both_levels(self):
        faculties = self.aggregate(splitter.aggregate_school_to_faculty_data)
        university = self.aggregate(splitter.aggregate_school_to_university_data)
        self.assertEqual(faculties.set_index('Faculty').loc['Arts', self.matrix], 50)
        self.assertEqual(university.iloc[0][self.matrix], 50)

    def test_entire_missing_faculty_displays_no_return(self):
        faculties = self.aggregate(splitter.aggregate_school_to_faculty_data)
        row = faculties.set_index('Faculty').loc['HR']
        self.assertFalse(row['Has KPI data'])
        data = self.generator._extract_kpi_data(row, 'HR')
        for kpi in data['kpis'].values():
            self.assertIsNone(kpi['percentage'])
            self.assertIsNone(kpi['number'])
            self.assertEqual(kpi['display_text'], 'No return submitted')

    def test_missing_schools_and_partial_fields_remain_missing(self):
        schools = self.aggregate(splitter.create_school_raw_data).set_index('School')
        self.assertFalse(schools.loc['HR', 'Has KPI data'])
        self.assertTrue(pd.isna(schools.loc['HR', self.matrix]))
        partial = self.source.copy()
        partial.loc[0, self.matrix] = np.nan
        schools = self.aggregate(splitter.create_school_raw_data, partial).set_index('School')
        self.assertTrue(schools.loc['CLAS', 'Has KPI data'])
        self.assertTrue(pd.isna(schools.loc['CLAS', self.matrix]))

    def test_genuine_all_zero_return_is_not_missing(self):
        source = self.source.iloc[[0]].copy()
        for col in source.columns[3:]:
            source[col] = 0
        row = self.aggregate(splitter.create_school_raw_data, source).iloc[0]
        self.assertTrue(row['Has KPI data'])
        self.assertFalse(self.generator._is_no_return_row(row))
        self.assertEqual(self.generator._extract_kpi_data(row, 'CLAS')['kpis']['Training Matrix Coverage']['percentage'], 0)

    def test_missing_history_does_not_become_zero(self):
        faculties = self.aggregate(splitter.aggregate_school_to_faculty_data)
        history = self.generator._build_kpi_history(faculties, 'Faculty', 'HR')
        for points in history.values():
            self.assertEqual(len(points), 1)
            self.assertIsNone(points[0]['percentage'])
            self.assertIsNone(points[0]['number'])

    def test_every_kpi_has_a_tooltip(self):
        tooltips = splitter.create_tooltips_dataframe()
        for definition in self.generator.kpi_definitions.values():
            self.assertIn(definition['percentage_col'], tooltips.columns)
            self.assertTrue(tooltips[definition['percentage_col']].iloc[1])


if __name__ == '__main__':
    unittest.main()
