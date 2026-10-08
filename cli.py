"""
Enterprise Employee Intelligence CLI
Advanced command-line interface supporting batch processing, arguments, and interactive wizard mode.
"""

import os
import sys
import argparse
from typing import Optional
import pandas as pd

# Ensure project root is in sys.path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

from src.data_loader import DataLoader
from src.analytics_engine import AnalyticsEngine
from src.predictive_models import CompensationPredictor
from src.reporter import ReportGenerator
from src.config import DEFAULT_SALARY_THRESHOLD, CURRENCY_SYMBOL

try:
    from tabulate import tabulate
    HAS_TABULATE = True
except ImportError:
    HAS_TABULATE = False


def print_table(df, title: Optional[str] = None):
    """Print clean formatted tables to terminal."""
    if title:
        print(f"\n{title}")
        print("=" * len(title))

    if HAS_TABULATE:
        print(tabulate(df, headers="keys", tablefmt="fancy_grid", showindex=False))
    else:
        print(df.to_string(index=False))


def display_executive_summary(kpis: dict):
    """Display stylized executive summary cards."""
    print("\n" + "=" * 65)
    print("           GLOBAL WORKFORCE INTELLIGENCE SUMMARY")
    print("=" * 65)
    print(f" Total Headcount            : {kpis['headcount']:,} employees")
    print(f" Active Departments         : {kpis['department_count']}")
    print(f" Total Annual Payroll       : {CURRENCY_SYMBOL}{kpis['total_payroll']:,.2f}")
    print(f" Company-Wide Average Salary: {CURRENCY_SYMBOL}{kpis['average_salary']:,.2f}")
    print(f" Median Base Salary         : {CURRENCY_SYMBOL}{kpis['median_salary']:,.2f}")
    print(f" Compensation Spread (Std)  : {CURRENCY_SYMBOL}{kpis['std_salary']:,.2f}")
    print(f" Salary Range (Min - Max)   : {CURRENCY_SYMBOL}{kpis['min_salary']:,.2f}  to  {CURRENCY_SYMBOL}{kpis['max_salary']:,.2f}")
    print(f" Interquartile Range (IQR)  : {CURRENCY_SYMBOL}{kpis['iqr_salary']:,.2f}")
    if kpis['average_experience'] > 0:
        print(f" Average Experience Tenure  : {kpis['average_experience']:.1f} years")
    if kpis['average_performance'] > 0:
        print(f" Average Performance Score : {kpis['average_performance']:.2f} / 5.0")
    print("=" * 65)


def run_interactive_mode(dataset_path: str):
    """Interactive CLI terminal session."""
    print("\n" + "#" * 65)
    print("  WELCOME TO ENTERPRISE WORKFORCE INTELLIGENCE & ANALYTICS")
    print("#" * 65)

    loader = DataLoader(dataset_path)
    success, msg = loader.load()
    if not success:
        print(f"\n[Fatal Error] {msg}")
        return

    df = loader.get_dataframe()
    engine = AnalyticsEngine(df)
    reporter = ReportGenerator(output_dir="reports")
    ml_model = CompensationPredictor()

    while True:
        print("\n--- ANALYTICS COMMAND CENTER ---")
        print(" [1] View Executive Summary KPIs")
        print(" [2] View Departmental Breakdown & Payroll Share")
        print(" [3] Filter Employees by Custom Salary Threshold (Assignment Task)")
        print(" [4] Multi-Criteria Advanced Filtering (Exp, Perf, Dept)")
        print(" [5] Detect Compensation Outliers & Anomalies (IQR / Z-Score)")
        print(" [6] Pay Equity & Demographic Parity Analysis")
        print(" [7] Train Machine Learning Salary Predictor & Simulator")
        print(" [8] Export Comprehensive Intelligence Package (Excel + HTML Report)")
        print(" [0] Exit Application")
        choice = input("\nEnter operation code [0-8]: ").strip()

        if choice == "1":
            kpis = engine.get_executive_summary()
            display_executive_summary(kpis)

        elif choice == "2":
            dept_df = engine.get_department_summary()
            print_table(dept_df, "--- Department Headcount, Payroll Share & Compensation Metrics ---")

        elif choice == "3":
            try:
                raw_val = input(f"\nEnter baseline threshold salary [{CURRENCY_SYMBOL}80000]: ").strip()
                threshold = float(raw_val) if raw_val else 80000.0
                filtered = engine.filter_records(min_salary=threshold)
                print(f"\n[Result] Found {len(filtered)} employees earning >= {CURRENCY_SYMBOL}{threshold:,.2f} ({len(filtered)/len(df)*100:.1f}% of workforce)")
                if not filtered.empty:
                    preview_cols = [c for c in ["Employee_ID", "Name", "Department", "Job_Title", "Salary", "Experience_Years"] if c in filtered.columns]
                    print_table(filtered[preview_cols].head(15), f"--- Top High Earners Preview (Showing up to 15 of {len(filtered)}) ---")

                    save = input("\nSave cohort to CSV? (y/n) [y]: ").strip().lower()
                    if save != "n":
                        fname = input("Enter output filename [high_earners.csv]: ").strip() or "high_earners.csv"
                        path = reporter.export_csv(filtered, fname)
                        print(f"Successfully saved to: {path}")
            except ValueError:
                print("Invalid numeric value entered.")

        elif choice == "4":
            print("\n-- Dynamic Multi-Filter --")
            try:
                min_s = input("Minimum Salary (blank to skip): ").strip()
                min_s = float(min_s) if min_s else None
                max_s = input("Maximum Salary (blank to skip): ").strip()
                max_s = float(max_s) if max_s else None
                min_exp = input("Minimum Experience Years (blank to skip): ").strip()
                min_exp = float(min_exp) if min_exp else None
                search = input("Search Name or Role (blank to skip): ").strip() or None

                res = engine.filter_records(min_salary=min_s, max_salary=max_s, min_experience=min_exp, search_query=search)
                print(f"\n[Filter Result] Matched {len(res)} employees.")
                if not res.empty:
                    preview_cols = [c for c in ["Employee_ID", "Name", "Department", "Job_Title", "Salary", "Experience_Years", "Performance_Score"] if c in res.columns]
                    print_table(res[preview_cols].head(20), "Filtered Results")
            except ValueError:
                print("Invalid numeric input.")

        elif choice == "5":
            outliers = engine.detect_outliers(method="iqr", threshold=1.5)
            print(f"\n[Outlier Detection] Found {len(outliers)} compensation anomalies via IQR method.")
            if not outliers.empty:
                cols = [c for c in ["Employee_ID", "Name", "Department", "Job_Title", "Salary", "Outlier_Type", "Outlier_Score"] if c in outliers.columns]
                print_table(outliers[cols], "Identified Outlier Records")

        elif choice == "6":
            equity = engine.get_pay_equity_summary()
            if "gender_stats" in equity:
                print_table(equity["gender_stats"], "Demographic Compensation Breakdown")
                print(f"\nFemale-to-Male Salary Ratio : {equity.get('female_to_male_ratio_pct', 'N/A')}%")
                print(f"Unadjusted Gender Pay Gap   : {equity.get('pay_gap_pct', 'N/A')}%")
            else:
                print("No demographic gender data available in this dataset.")

        elif choice == "7":
            print("\nTraining Machine Learning Compensation Model...")
            success, tr_msg = ml_model.train(df)
            print(f"[{'Success' if success else 'Failed'}] {tr_msg}")

            if success:
                print("\n-- Salary Prediction Simulator --")
                try:
                    exp = float(input("Candidate Experience Years: "))
                    depts = df["Department"].unique().tolist()
                    print(f"Available Departments: {', '.join(depts)}")
                    dept = input("Target Department: ").strip()
                    if dept in depts:
                        predicted = ml_model.predict_salary(exp, dept)
                        print(f"\n>>> Recommended Market Salary: {CURRENCY_SYMBOL}{predicted:,.2f} <<<")
                    else:
                        print("Invalid department entered.")
                except ValueError:
                    print("Invalid input.")

        elif choice == "8":
            kpis = engine.get_executive_summary()
            dept_df = engine.get_department_summary()
            high_earners = engine.filter_records(min_salary=DEFAULT_SALARY_THRESHOLD)
            outliers = engine.detect_outliers(method="iqr")

            # Excel
            sheets = {
                "Executive_KPIs": pd.DataFrame([kpis]),
                "Department_Summary": dept_df,
                "High_Earners": high_earners,
                "Outliers": outliers,
                "Full_Dataset": df
            }
            xls_path = reporter.export_excel(sheets, "workforce_intelligence_full.xlsx")
            html_path = reporter.generate_html_report(kpis, dept_df, high_earners, outliers, "executive_brief.html")

            print("\n[Export Complete]")
            print(f" - Multi-Sheet Styled Excel : {xls_path}")
            print(f" - Standalone HTML Brief    : {html_path}")

        elif choice == "0":
            print("\nExiting Enterprise Workforce Intelligence Suite. Goodbye!\n")
            break
        else:
            print("Invalid command. Please enter a number from 0 to 8.")


def main():
    parser = argparse.ArgumentParser(
        description="Enterprise Employee Intelligence & Analytics Suite",
        formatter_class=argparse.RawDescriptionHelpFormatter
    )

    default_data = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "employees.csv")

    parser.add_argument("--dataset", type=str, default=default_data, help="Path to employee CSV, Excel, or JSON dataset")
    parser.add_argument("--threshold", type=float, default=None, help="Filter employees earning above this salary threshold")
    parser.add_argument("--department", type=str, default=None, help="Comma-separated department filter (e.g. IT,Finance)")
    parser.add_argument("--summary", action="store_true", help="Display executive KPIs and department summary")
    parser.add_argument("--outliers", action="store_true", help="Run IQR compensation anomaly detection")
    parser.add_argument("--train-ml", action="store_true", help="Train ML Salary prediction model")
    parser.add_argument("--export-csv", type=str, default=None, help="Export filtered data to CSV file")
    parser.add_argument("--export-excel", type=str, default=None, help="Export comprehensive multi-sheet Excel file")
    parser.add_argument("--export-html", type=str, default=None, help="Export publication-ready HTML executive brief")
    parser.add_argument("--interactive", action="store_true", help="Launch interactive CLI menu")

    args = parser.parse_args()

    # If no flags passed or interactive requested, enter interactive mode
    if args.interactive or len(sys.argv) == 1:
        run_interactive_mode(args.dataset)
        return

    # Batch Mode
    loader = DataLoader(args.dataset)
    success, msg = loader.load()
    if not success:
        print(f"[Error] {msg}")
        sys.exit(1)

    df = loader.get_dataframe()
    engine = AnalyticsEngine(df)
    reporter = ReportGenerator(output_dir="reports")

    if args.summary:
        kpis = engine.get_executive_summary()
        display_executive_summary(kpis)
        dept_df = engine.get_department_summary()
        print_table(dept_df, "--- Department Headcount & Average Compensation ---")

    filtered_df = df
    if args.threshold is not None or args.department is not None:
        depts = [d.strip() for d in args.department.split(",")] if args.department else None
        filtered_df = engine.filter_records(min_salary=args.threshold, departments=depts)
        label = f"Filtered Cohort (Threshold: {args.threshold or 'None'}, Dept: {args.department or 'All'})"
        print(f"\n[Filter] Found {len(filtered_df)} records matching criteria.")
        print_table(filtered_df.head(20), label)

    if args.outliers:
        outliers = engine.detect_outliers(method="iqr")
        print_table(outliers.head(20), "Compensation Outlier Records (IQR Method)")

    if args.train_ml:
        predictor = CompensationPredictor()
        s, m = predictor.train(df)
        print(f"\n[Machine Learning] {m}")

    if args.export_csv:
        path = reporter.export_csv(filtered_df, args.export_csv)
        print(f"\n[Export CSV] Saved to: {path}")

    if args.export_excel:
        kpis = engine.get_executive_summary()
        dept_df = engine.get_department_summary()
        sheets = {
            "Executive_KPIs": pd.DataFrame([kpis]),
            "Department_Summary": dept_df,
            "Filtered_Data": filtered_df
        }
        path = reporter.export_excel(sheets, args.export_excel)
        print(f"\n[Export Excel] Saved to: {path}")

    if args.export_html:
        kpis = engine.get_executive_summary()
        dept_df = engine.get_department_summary()
        path = reporter.generate_html_report(kpis, dept_df, filtered_df, filename=args.export_html)
        print(f"\n[Export HTML] Saved to: {path}")


if __name__ == "__main__":
    main()
