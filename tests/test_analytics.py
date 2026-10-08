"""
Unit Test Suite for Employee Intelligence & Analytics Suite
Verifies data ingestion, statistical calculations, outlier logic, ML pipelines, and exports.
"""

import os
import sys
import unittest
import numpy as np
import pandas as pd

# Add project root to sys.path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from src.data_loader import DataLoader
from src.analytics_engine import AnalyticsEngine
from src.predictive_models import CompensationPredictor
from src.reporter import ReportGenerator


class TestEmployeeAnalytics(unittest.TestCase):
    """Test suite ensuring data integrity and statistical accuracy."""

    def setUp(self):
        # Base fixture with deterministic test data
        self.mock_df = pd.DataFrame({
            "Employee_ID": [101, 102, 103, 104, 105, 106, 107, 108],
            "Name": ["Rahul Sharma", "Anjali Verma", "Vikram Sai", "Priya Patel",
                     "Amit Mishra", "Neha Reddy", "Rohan Das", "Sneha Paul"],
            "Department": ["IT", "HR", "Finance", "IT", "Marketing", "Finance", "HR", "IT"],
            "Salary": [85000, 62000, 95000, 105000, 58000, 92000, 64000, 72000],
            "Experience_Years": [5, 3, 7, 9, 2, 6, 4, 3],
            "Performance_Score": [4.2, 3.5, 4.0, 4.8, 3.1, 4.1, 3.4, 3.8]
        })
        self.engine = AnalyticsEngine(self.mock_df)

    def test_executive_summary_statistics(self):
        """Verifies company-wide math calculations."""
        kpis = self.engine.get_executive_summary()
        self.assertEqual(kpis["headcount"], 8)
        self.assertEqual(kpis["department_count"], 4)

        # Average: (85000+62000+95000+105000+58000+92000+64000+72000)/8 = 79,125
        expected_avg = self.mock_df["Salary"].mean()
        self.assertAlmostEqual(kpis["average_salary"], expected_avg, places=2)
        self.assertEqual(kpis["min_salary"], 58000)
        self.assertEqual(kpis["max_salary"], 105000)

    def test_department_aggregation(self):
        """Verifies group-by calculations and headcount sums."""
        dept_summary = self.engine.get_department_summary()
        self.assertEqual(len(dept_summary), 4)

        # IT should have 3 employees
        it_row = dept_summary[dept_summary["Department"] == "IT"].iloc[0]
        self.assertEqual(it_row["Headcount"], 3)
        # IT average: (85000 + 105000 + 72000) / 3 = 87333.33
        self.assertAlmostEqual(it_row["Average_Salary"], 87333.33, places=1)

    def test_filtering_logic(self):
        """Tests filtering by threshold salary and department."""
        # Filter above 80,000 (should yield Rahul, Vikram, Priya, Neha -> 4 employees)
        high_earners = self.engine.filter_records(min_salary=80000)
        self.assertEqual(len(high_earners), 4)
        self.assertTrue(all(high_earners["Salary"] >= 80000))

        # Filter by department
        it_only = self.engine.filter_records(departments=["IT"])
        self.assertEqual(len(it_only), 3)

        # Filter by experience
        senior_staff = self.engine.filter_records(min_experience=6)
        self.assertEqual(len(senior_staff), 3)

    def test_data_loader_with_aliases(self):
        """Tests that aliases and non-standard column names are mapped accurately."""
        temp_csv = "test_alias_data.csv"
        df_alias = pd.DataFrame({
            "Emp ID": [1, 2],
            "Full Name": ["Alice", "Bob"],
            "Business Unit": ["Finance", "IT"],
            "Annual Salary": [70000, 80000],
            "Tenure": [3, 4]
        })
        df_alias.to_csv(temp_csv, index=False)

        try:
            loader = DataLoader(temp_csv)
            success, msg = loader.load()
            self.assertTrue(success)
            loaded_df = loader.get_dataframe()
            self.assertIn("Salary", loaded_df.columns)
            self.assertIn("Department", loaded_df.columns)
            self.assertIn("Experience_Years", loaded_df.columns)
            self.assertEqual(len(loaded_df), 2)
        finally:
            if os.path.exists(temp_csv):
                os.remove(temp_csv)

    def test_predictive_model_training(self):
        """Tests ML compensation predictor training and inference."""
        predictor = CompensationPredictor()
        success, msg = predictor.train(self.mock_df)
        self.assertTrue(success)
        self.assertTrue(predictor.is_trained)

        # Prediction test
        pred = predictor.predict_salary(experience_years=5, department="IT")
        self.assertIsInstance(pred, float)
        self.assertGreater(pred, 0)

    def test_report_generator_outputs(self):
        """Tests CSV, Excel, and HTML export generation."""
        test_dir = "test_reports"
        reporter = ReportGenerator(output_dir=test_dir)

        # CSV
        csv_file = reporter.export_csv(self.mock_df, "test.csv")
        self.assertTrue(os.path.exists(csv_file))

        # Excel
        excel_file = reporter.export_excel({"Data": self.mock_df}, "test.xlsx")
        self.assertTrue(os.path.exists(excel_file))

        # HTML
        kpis = self.engine.get_executive_summary()
        dept_df = self.engine.get_department_summary()
        html_file = reporter.generate_html_report(kpis, dept_df, self.mock_df, filename="test.html")
        self.assertTrue(os.path.exists(html_file))

        # Cleanup
        for f in [csv_file, excel_file, html_file]:
            if os.path.exists(f):
                os.remove(f)
        if os.path.exists(test_dir):
            os.rmdir(test_dir)


if __name__ == "__main__":
    unittest.main()
