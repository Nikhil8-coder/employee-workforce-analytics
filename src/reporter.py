"""
Multi-Format Export & Executive Reporting Engine
Exports to CSV, Multi-Tab Excel (.xlsx) with formatting, and Standalone Executive HTML Reports.
"""

import os
from typing import Dict, Any, Optional
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


class ReportGenerator:
    """Enterprise report generator for data exports and executive stakeholder briefings."""

    def __init__(self, output_dir: str = "reports"):
        self.output_dir = output_dir
        os.makedirs(self.output_dir, exist_ok=True)

    def export_csv(self, df: pd.DataFrame, filename: str) -> str:
        """Exports DataFrame to CSV document."""
        if not filename.endswith(".csv"):
            filename += ".csv"
        path = os.path.join(self.output_dir, filename)
        df.to_csv(path, index=False)
        return path

    def export_excel(
        self,
        sheets: Dict[str, pd.DataFrame],
        filename: str = "workforce_intelligence_report.xlsx"
    ) -> str:
        """
        Exports multiple DataFrames into a beautifully styled, multi-sheet Excel workbook.
        """
        if not filename.endswith(".xlsx"):
            filename += ".xlsx"
        path = os.path.join(self.output_dir, filename)

        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            for sheet_name, df in sheets.items():
                safe_name = sheet_name[:31]  # Excel 31 char limit
                df.to_excel(writer, sheet_name=safe_name, index=False)

                # Style the worksheet
                ws = writer.sheets[safe_name]
                header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
                header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
                align_center = Alignment(horizontal="center", vertical="center")
                thin_border = Border(
                    left=Side(style='thin', color='E5E7EB'),
                    right=Side(style='thin', color='E5E7EB'),
                    top=Side(style='thin', color='E5E7EB'),
                    bottom=Side(style='thin', color='E5E7EB')
                )

                # Format headers
                for cell in ws[1]:
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = align_center

                # Adjust column widths
                for col in ws.columns:
                    max_len = max(len(str(cell.value or '')) for cell in col)
                    col_letter = get_column_letter(col[0].column)
                    ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

                    for cell in col[1:]:
                        cell.border = thin_border
                        if isinstance(cell.value, float):
                            cell.number_format = '#,##0.00'

        return path

    def generate_html_report(
        self,
        kpis: Dict[str, Any],
        dept_df: pd.DataFrame,
        high_earners_df: pd.DataFrame,
        outliers_df: Optional[pd.DataFrame] = None,
        filename: str = "executive_workforce_brief.html"
    ) -> str:
        """Generates a self-contained, publication-grade HTML executive briefing."""
        if not filename.endswith(".html"):
            filename += ".html"
        path = os.path.join(self.output_dir, filename)

        dept_table_html = dept_df.to_html(classes="min-w-full divide-y divide-gray-200 text-sm", index=False)
        high_earners_preview = high_earners_df.head(15).to_html(classes="min-w-full divide-y divide-gray-200 text-sm", index=False)
        outlier_preview = (outliers_df.head(15).to_html(classes="min-w-full divide-y divide-gray-200 text-sm", index=False)) if outliers_df is not None and not outliers_df.empty else "<p class='text-gray-500 italic p-4'>No statistical compensation outliers detected.</p>"

        html_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Workforce Compensation & Analytics Intelligence Brief</title>
  <script src="https://cdn.tailwindcss.com"></script>
  <style>
    table {{ border-collapse: collapse; width: 100%; }}
    th {{ background-color: #f8fafc; font-weight: 600; text-align: left; padding: 0.75rem 1rem; border-bottom: 2px solid #e2e8f0; color: #334155; }}
    td {{ padding: 0.65rem 1rem; border-bottom: 1px solid #f1f5f9; color: #475569; }}
    tr:hover {{ background-color: #f8fafc; }}
  </style>
</head>
<body class="bg-slate-50 text-slate-800 font-sans p-6 md:p-10">
  <div class="max-w-7xl mx-auto space-y-8">
    <!-- Header -->
    <header class="bg-white p-6 rounded-2xl shadow-sm border border-slate-200 flex flex-col md:flex-row justify-between items-start md:items-center">
      <div>
        <h1 class="text-3xl font-extrabold text-slate-900 tracking-tight">Workforce Compensation & Analytics Brief</h1>
        <p class="text-slate-500 text-sm mt-1">Generated by Enterprise Employee Intelligence System</p>
      </div>
      <div class="mt-4 md:mt-0 flex items-center space-x-3">
        <span class="inline-flex items-center px-3 py-1 rounded-full text-xs font-semibold bg-emerald-100 text-emerald-800">
          ● Live Production Audit
        </span>
      </div>
    </header>

    <!-- Top KPI Cards -->
    <div class="grid grid-cols-1 md:grid-cols-4 gap-6">
      <div class="bg-white p-6 rounded-2xl shadow-sm border border-slate-200">
        <p class="text-xs uppercase tracking-wider font-semibold text-slate-400">Total Headcount</p>
        <p class="text-3xl font-bold text-slate-900 mt-2">{kpis.get('headcount', 0):,}</p>
        <p class="text-xs text-slate-500 mt-1">{kpis.get('department_count', 0)} Active Departments</p>
      </div>
      <div class="bg-white p-6 rounded-2xl shadow-sm border border-slate-200">
        <p class="text-xs uppercase tracking-wider font-semibold text-slate-400">Total Payroll</p>
        <p class="text-3xl font-bold text-indigo-600 mt-2">₹{kpis.get('total_payroll', 0):,.2f}</p>
        <p class="text-xs text-slate-500 mt-1">Annualised Cumulative</p>
      </div>
      <div class="bg-white p-6 rounded-2xl shadow-sm border border-slate-200">
        <p class="text-xs uppercase tracking-wider font-semibold text-slate-400">Average Salary</p>
        <p class="text-3xl font-bold text-slate-900 mt-2">₹{kpis.get('average_salary', 0):,.2f}</p>
        <p class="text-xs text-slate-500 mt-1">Median: ₹{kpis.get('median_salary', 0):,.2f}</p>
      </div>
      <div class="bg-white p-6 rounded-2xl shadow-sm border border-slate-200">
        <p class="text-xs uppercase tracking-wider font-semibold text-slate-400">High Earner Volume</p>
        <p class="text-3xl font-bold text-amber-600 mt-2">{len(high_earners_df):,}</p>
        <p class="text-xs text-slate-500 mt-1">{(len(high_earners_df)/max(1, kpis.get('headcount', 1))*100):.1f}% of Workforce</p>
      </div>
    </div>

    <!-- Department Summary Table -->
    <div class="bg-white p-6 rounded-2xl shadow-sm border border-slate-200">
      <h2 class="text-xl font-bold text-slate-900 mb-4">Departmental Headcount & Compensation Metrics</h2>
      <div class="overflow-x-auto rounded-lg border border-slate-100">
        {dept_table_html}
      </div>
    </div>

    <!-- High Earners Table -->
    <div class="bg-white p-6 rounded-2xl shadow-sm border border-slate-200">
      <div class="flex justify-between items-center mb-4">
        <h2 class="text-xl font-bold text-slate-900">Target Cohort: High Earners (Top 15 Preview)</h2>
        <span class="text-xs font-medium text-slate-500">Total in Cohort: {len(high_earners_df)}</span>
      </div>
      <div class="overflow-x-auto rounded-lg border border-slate-100">
        {high_earners_preview}
      </div>
    </div>

    <!-- Outlier Section -->
    <div class="bg-white p-6 rounded-2xl shadow-sm border border-slate-200">
      <h2 class="text-xl font-bold text-slate-900 mb-2">Compensation Anomaly & Outlier Watchlist</h2>
      <p class="text-sm text-slate-500 mb-4">Statistical anomalies identified via Interquartile Range (IQR) and Z-Score algorithms.</p>
      <div class="overflow-x-auto rounded-lg border border-slate-100">
        {outlier_preview}
      </div>
    </div>

    <!-- Footer -->
    <footer class="text-center text-xs text-slate-400 pt-4">
      Enterprise Employee Intelligence Suite &bull; Confidential Executive Brief &bull; Automated Pipeline Output
    </footer>
  </div>
</body>
</html>
"""
        with open(path, "w", encoding="utf-8") as f:
            f.write(html_content)

        return path
