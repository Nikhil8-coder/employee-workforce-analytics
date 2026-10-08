"""
Advanced Analytics Engine
Computes executive KPIs, departmental breakdowns, statistical dispersion,
pay equity metrics, IQR/Z-score outlier detection, and dynamic multi-criteria filtering.
"""

from typing import Dict, List, Optional, Any, Union
import numpy as np
import pandas as pd


class AnalyticsEngine:
    """Core analytical intelligence suite for employee compensation and workforce data."""

    def __init__(self, df: pd.DataFrame):
        if df is None or df.empty:
            raise ValueError("AnalyticsEngine requires a non-empty DataFrame.")
        self.df = df.copy()

    # -------------------------------------------------------------
    # 1. Executive Summary KPIs
    # -------------------------------------------------------------
    def get_executive_summary(self) -> Dict[str, Any]:
        """Compute top-level corporate workforce metrics."""
        salary = self.df["Salary"]
        headcount = len(self.df)
        total_payroll = float(salary.sum())
        mean_sal = float(salary.mean())
        median_sal = float(salary.median())
        std_sal = float(salary.std()) if headcount > 1 else 0.0
        min_sal = float(salary.min())
        max_sal = float(salary.max())
        q25 = float(salary.quantile(0.25))
        q75 = float(salary.quantile(0.75))
        iqr = q75 - q25

        avg_exp = float(self.df["Experience_Years"].mean()) if "Experience_Years" in self.df.columns else 0.0
        avg_perf = float(self.df["Performance_Score"].mean()) if "Performance_Score" in self.df.columns else 0.0

        num_depts = self.df["Department"].nunique() if "Department" in self.df.columns else 1

        return {
            "headcount": headcount,
            "total_payroll": total_payroll,
            "average_salary": mean_sal,
            "median_salary": median_sal,
            "std_salary": std_sal,
            "min_salary": min_sal,
            "max_salary": max_sal,
            "q25_salary": q25,
            "q75_salary": q75,
            "iqr_salary": iqr,
            "average_experience": avg_exp,
            "average_performance": avg_perf,
            "department_count": num_depts
        }

    # -------------------------------------------------------------
    # 2. Departmental Breakdown
    # -------------------------------------------------------------
    def get_department_summary(self) -> pd.DataFrame:
        """Group by Department and aggregate comprehensive workforce metrics."""
        total_payroll_all = self.df["Salary"].sum()

        agg_dict: Dict[str, Any] = {
            "Employee_ID": "count",
            "Salary": ["mean", "median", "min", "max", "std", "sum"]
        }

        if "Experience_Years" in self.df.columns:
            agg_dict["Experience_Years"] = "mean"

        if "Performance_Score" in self.df.columns:
            agg_dict["Performance_Score"] = "mean"

        grouped = self.df.groupby("Department").agg(agg_dict)

        # Flatten multi-level column names
        summary = pd.DataFrame()
        summary["Department"] = grouped.index
        summary["Headcount"] = grouped[("Employee_ID", "count")].values
        summary["Headcount_Share_Pct"] = (summary["Headcount"] / len(self.df) * 100).round(1)
        summary["Average_Salary"] = grouped[("Salary", "mean")].round(2).values
        summary["Median_Salary"] = grouped[("Salary", "median")].round(2).values
        summary["Min_Salary"] = grouped[("Salary", "min")].round(2).values
        summary["Max_Salary"] = grouped[("Salary", "max")].round(2).values
        summary["Std_Dev_Salary"] = grouped[("Salary", "std")].fillna(0).round(2).values
        summary["Total_Payroll"] = grouped[("Salary", "sum")].round(2).values
        summary["Payroll_Share_Pct"] = ((summary["Total_Payroll"] / max(1, total_payroll_all)) * 100).round(1)

        if "Experience_Years" in self.df.columns:
            summary["Avg_Experience"] = grouped[("Experience_Years", "mean")].round(1).values

        if "Performance_Score" in self.df.columns:
            summary["Avg_Performance"] = grouped[("Performance_Score", "mean")].round(2).values

        # Sort descending by headcount
        summary.sort_values(by="Headcount", ascending=False, inplace=True)
        summary.reset_index(drop=True, inplace=True)
        return summary

    # -------------------------------------------------------------
    # 3. Role / Job Title Breakdown
    # -------------------------------------------------------------
    def get_role_summary(self, min_count: int = 1) -> pd.DataFrame:
        """Breakdown statistics by Job Title."""
        if "Job_Title" not in self.df.columns:
            return pd.DataFrame()

        grouped = self.df.groupby("Job_Title").agg(
            Headcount=("Employee_ID", "count"),
            Average_Salary=("Salary", "mean"),
            Median_Salary=("Salary", "median"),
            Min_Salary=("Salary", "min"),
            Max_Salary=("Salary", "max")
        ).reset_index()

        grouped = grouped[grouped["Headcount"] >= min_count]
        grouped["Average_Salary"] = grouped["Average_Salary"].round(2)
        grouped["Median_Salary"] = grouped["Median_Salary"].round(2)
        grouped.sort_values(by="Average_Salary", ascending=False, inplace=True)
        return grouped

    # -------------------------------------------------------------
    # 4. Outlier & Anomaly Detection (IQR & Z-Score)
    # -------------------------------------------------------------
    def detect_outliers(self, method: str = "iqr", threshold: float = 1.5, by_department: bool = False) -> pd.DataFrame:
        """
        Detect salary anomalies.
        - 'iqr': Standard Tukey boxplot fences (Q1 - k*IQR, Q3 + k*IQR).
        - 'zscore': Standard score (|z| > threshold, e.g. 2.5 or 3.0).
        Can be calculated globally or within each department independently.
        """
        df_out = self.df.copy()

        if by_department and "Department" in df_out.columns:
            def flag_dept_outliers(group):
                sal = group["Salary"]
                if len(group) < 4:
                    group["Is_Outlier"] = False
                    group["Outlier_Type"] = "Normal"
                    group["Outlier_Score"] = 0.0
                    return group

                if method == "iqr":
                    q1 = sal.quantile(0.25)
                    q3 = sal.quantile(0.75)
                    iqr = q3 - q1
                    lower = q1 - threshold * iqr
                    upper = q3 + threshold * iqr
                    is_high = sal > upper
                    is_low = sal < lower
                    group["Is_Outlier"] = is_high | is_low
                    group["Outlier_Type"] = np.where(is_high, "High Anomaly", np.where(is_low, "Low Anomaly", "Normal"))
                    group["Outlier_Score"] = ((sal - group["Salary"].mean()) / max(1, group["Salary"].std())).round(2)
                else:  # zscore
                    std = sal.std()
                    mean = sal.mean()
                    z = (sal - mean) / (std if std > 0 else 1)
                    group["Is_Outlier"] = z.abs() > threshold
                    group["Outlier_Type"] = np.where(z > threshold, "High Anomaly", np.where(z < -threshold, "Low Anomaly", "Normal"))
                    group["Outlier_Score"] = z.round(2)
                return group

            result = df_out.groupby("Department", group_keys=False).apply(flag_dept_outliers)
        else:
            sal = df_out["Salary"]
            if method == "iqr":
                q1 = sal.quantile(0.25)
                q3 = sal.quantile(0.75)
                iqr = q3 - q1
                lower = q1 - threshold * iqr
                upper = q3 + threshold * iqr
                is_high = sal > upper
                is_low = sal < lower
                df_out["Is_Outlier"] = is_high | is_low
                df_out["Outlier_Type"] = np.where(is_high, "High Anomaly", np.where(is_low, "Low Anomaly", "Normal"))
                df_out["Outlier_Score"] = ((sal - sal.mean()) / max(1, sal.std())).round(2)
            else:  # zscore
                std = sal.std()
                mean = sal.mean()
                z = (sal - mean) / (std if std > 0 else 1)
                df_out["Is_Outlier"] = z.abs() > threshold
                df_out["Outlier_Type"] = np.where(z > threshold, "High Anomaly", np.where(z < -threshold, "Low Anomaly", "Normal"))
                df_out["Outlier_Score"] = z.round(2)
            result = df_out

        outliers_only = result[result["Is_Outlier"] == True].copy()
        outliers_only.sort_values(by="Salary", ascending=False, inplace=True)
        return outliers_only

    # -------------------------------------------------------------
    # 5. Pay Equity & Demographic Analysis
    # -------------------------------------------------------------
    def get_pay_equity_summary(self) -> Dict[str, Any]:
        """Calculates demographic pay parity metrics and ratios."""
        equity_data: Dict[str, Any] = {}

        if "Gender" in self.df.columns:
            gender_stats = self.df.groupby("Gender").agg(
                Headcount=("Employee_ID", "count"),
                Average_Salary=("Salary", "mean"),
                Median_Salary=("Salary", "median"),
                Avg_Experience=("Experience_Years", "mean") if "Experience_Years" in self.df.columns else ("Salary", "count")
            ).round(2).reset_index()

            equity_data["gender_stats"] = gender_stats

            # Calculate Female to Male parity ratio
            f_sal = gender_stats.loc[gender_stats["Gender"] == "Female", "Average_Salary"]
            m_sal = gender_stats.loc[gender_stats["Gender"] == "Male", "Average_Salary"]

            if not f_sal.empty and not m_sal.empty and m_sal.values[0] > 0:
                parity_ratio = round((f_sal.values[0] / m_sal.values[0]) * 100, 1)
                equity_data["female_to_male_ratio_pct"] = parity_ratio
                equity_data["pay_gap_pct"] = round(100 - parity_ratio, 1)
            else:
                equity_data["female_to_male_ratio_pct"] = 100.0
                equity_data["pay_gap_pct"] = 0.0

        return equity_data

    # -------------------------------------------------------------
    # 6. Dynamic Filtering Engine
    # -------------------------------------------------------------
    def filter_records(
        self,
        min_salary: Optional[float] = None,
        max_salary: Optional[float] = None,
        departments: Optional[List[str]] = None,
        min_experience: Optional[float] = None,
        max_experience: Optional[float] = None,
        min_performance: Optional[float] = None,
        locations: Optional[List[str]] = None,
        search_query: Optional[str] = None
    ) -> pd.DataFrame:
        """
        Multi-criteria high-performance filtering.
        Expands beyond the original single-threshold filter.
        """
        filtered = self.df.copy()

        if min_salary is not None:
            filtered = filtered[filtered["Salary"] >= min_salary]

        if max_salary is not None:
            filtered = filtered[filtered["Salary"] <= max_salary]

        if departments and len(departments) > 0:
            filtered = filtered[filtered["Department"].isin(departments)]

        if min_experience is not None and "Experience_Years" in filtered.columns:
            filtered = filtered[filtered["Experience_Years"] >= min_experience]

        if max_experience is not None and "Experience_Years" in filtered.columns:
            filtered = filtered[filtered["Experience_Years"] <= max_experience]

        if min_performance is not None and "Performance_Score" in filtered.columns:
            filtered = filtered[filtered["Performance_Score"] >= min_performance]

        if locations and len(locations) > 0 and "Location" in filtered.columns:
            filtered = filtered[filtered["Location"].isin(locations)]

        if search_query and search_query.strip():
            q = search_query.strip().lower()
            mask = filtered["Name"].astype(str).str.lower().str.contains(q)
            if "Employee_ID" in filtered.columns:
                mask = mask | filtered["Employee_ID"].astype(str).str.lower().str.contains(q)
            if "Job_Title" in filtered.columns:
                mask = mask | filtered["Job_Title"].astype(str).str.lower().str.contains(q)
            filtered = filtered[mask]

        return filtered
