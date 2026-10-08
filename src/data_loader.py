"""
Robust Data Ingestion & Sanitization Engine
Handles multi-format data intake (CSV, Excel, JSON), schema alignment, validation, and derived features.
"""

import os
import logging
from typing import Optional, Dict, Tuple, Any
import numpy as np
import pandas as pd
from src.config import COLUMN_ALIASES, EXPERIENCE_BRACKETS

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger("DataLoader")


class DataLoader:
    """Enterprise-grade data loader with schema normalisation and validation."""

    def __init__(self, file_path: str):
        self.file_path = file_path
        self.raw_df: Optional[pd.DataFrame] = None
        self.df: Optional[pd.DataFrame] = None
        self.metadata: Dict[str, Any] = {}

    def _normalize_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        """Map heterogeneous raw column headers to standard schema names."""
        df = df.copy()
        current_cols = {col.lower().strip().replace(" ", "_"): col for col in df.columns}
        rename_map = {}

        for standard_name, aliases in COLUMN_ALIASES.items():
            for alias in aliases:
                norm_alias = alias.lower().strip().replace(" ", "_")
                if norm_alias in current_cols:
                    rename_map[current_cols[norm_alias]] = standard_name
                    break

        df.rename(columns=rename_map, inplace=True)
        return df

    def _derive_features(self, df: pd.DataFrame) -> pd.DataFrame:
        """Derive analytical features like experience bands and compensation metrics."""
        df = df.copy()

        # Experience Band
        if "Experience_Years" in df.columns:
            def assign_band(exp):
                for low, high, label in EXPERIENCE_BRACKETS:
                    if low <= exp <= high:
                        return label
                return "Unknown"

            df["Experience_Band"] = df["Experience_Years"].apply(assign_band)

        # Bonus & Total Compensation
        if "Salary" in df.columns:
            if "Bonus_Percentage" in df.columns:
                df["Bonus_Amount"] = (df["Salary"] * (df["Bonus_Percentage"] / 100)).round(2)
                df["Total_Compensation"] = df["Salary"] + df["Bonus_Amount"]
            else:
                df["Total_Compensation"] = df["Salary"]

            # Salary Quartiles
            if len(df) >= 4:
                try:
                    df["Salary_Quartile"] = pd.qcut(
                        df["Salary"], q=4, labels=["Q1 (Entry)", "Q2 (Median-)", "Q3 (Median+)", "Q4 (Top)"]
                    )
                except Exception:
                    df["Salary_Quartile"] = "Q1"

        return df

    def load(self) -> Tuple[bool, str]:
        """Loads and validates data from the specified file path."""
        if not os.path.exists(self.file_path):
            msg = f"Target file '{self.file_path}' does not exist on disk."
            logger.error(msg)
            return False, msg

        ext = os.path.splitext(self.file_path)[1].lower()

        try:
            if ext == ".csv":
                self.raw_df = pd.read_csv(self.file_path)
            elif ext in [".xlsx", ".xls"]:
                self.raw_df = pd.read_excel(self.file_path)
            elif ext == ".json":
                self.raw_df = pd.read_json(self.file_path)
            else:
                msg = f"Unsupported file format '{ext}'. Supported formats: .csv, .xlsx, .xls, .json"
                logger.error(msg)
                return False, msg

            logger.info(f"Loaded raw file '{self.file_path}' with {len(self.raw_df)} rows and {len(self.raw_df.columns)} columns.")

            # Normalization
            df = self._normalize_columns(self.raw_df)

            # Minimum required columns check
            required_cols = ["Salary", "Department"]
            missing_required = [c for c in required_cols if c not in df.columns]
            if missing_required:
                msg = f"Schema validation error: Missing required column(s): {missing_required}. Available columns: {list(df.columns)}"
                logger.error(msg)
                return False, msg

            # Clean and sanitize numeric columns
            df["Salary"] = pd.to_numeric(df["Salary"], errors="coerce").fillna(0)
            if "Experience_Years" in df.columns:
                df["Experience_Years"] = pd.to_numeric(df["Experience_Years"], errors="coerce").fillna(0)
            else:
                df["Experience_Years"] = 0

            if "Performance_Score" in df.columns:
                df["Performance_Score"] = pd.to_numeric(df["Performance_Score"], errors="coerce").fillna(3.0)

            if "Employee_ID" not in df.columns:
                df["Employee_ID"] = [f"EMP-{i+1:04d}" for i in range(len(df))]

            if "Name" not in df.columns:
                df["Name"] = [f"Employee {emp_id}" for emp_id in df["Employee_ID"]]

            # Derive enriched features
            self.df = self._derive_features(df)

            self.metadata = {
                "file_path": self.file_path,
                "row_count": len(self.df),
                "columns": list(self.df.columns),
                "departments": self.df["Department"].unique().tolist(),
                "min_salary": float(self.df["Salary"].min()),
                "max_salary": float(self.df["Salary"].max()),
                "avg_salary": float(self.df["Salary"].mean())
            }

            msg = f"Successfully validated and enriched {len(self.df)} employee records."
            logger.info(msg)
            return True, msg

        except Exception as e:
            msg = f"Failed to ingest dataset: {str(e)}"
            logger.exception(msg)
            return False, msg

    def get_dataframe(self) -> pd.DataFrame:
        """Returns the sanitized pandas DataFrame."""
        if self.df is None:
            raise ValueError("No dataset loaded. Call load() first.")
        return self.df
