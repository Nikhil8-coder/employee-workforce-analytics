"""
Configuration and constants for the Employee Intelligence & Analytics Suite.
"""

from typing import Dict, List

import sys
import io

# Ensure UTF-8 output encoding on Windows consoles
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

# Currency display configuration
CURRENCY_SYMBOL = "₹"
CURRENCY_CODE = "INR"

# Column name mappings to support heterogeneous CSV/Excel imports
COLUMN_ALIASES: Dict[str, List[str]] = {
    "Employee_ID": ["employee_id", "emp_id", "id", "employeeid", "emp id"],
    "Name": ["name", "full_name", "employee_name", "fullname", "emp_name"],
    "Department": ["department", "dept", "business_unit", "division"],
    "Salary": ["salary", "annual_salary", "base_salary", "ctc", "compensation"],
    "Experience_Years": ["experience_years", "experience", "exp_years", "exp", "tenure", "years_of_experience"],
    "Job_Title": ["job_title", "title", "role", "designation", "position"],
    "Bonus_Percentage": ["bonus_percentage", "bonus_pct", "bonus", "bonus_%"],
    "Performance_Score": ["performance_score", "performance", "rating", "perf_score", "perf_rating"],
    "Gender": ["gender", "sex"],
    "Age": ["age"],
    "Location": ["location", "city", "office_location", "work_location"],
    "Education": ["education", "education_level", "degree", "qualification"],
    "Joining_Date": ["joining_date", "hire_date", "start_date", "joined_on"],
    "Attrition_Risk": ["attrition_risk", "churn_risk", "flight_risk", "retention_risk"]
}

# Default salary filtering threshold (matching assignment default or configurable)
DEFAULT_SALARY_THRESHOLD = 80000.0

# Experience brackets for grouping
EXPERIENCE_BRACKETS = [
    (0, 2, "Entry (0-2 yrs)"),
    (3, 5, "Junior (3-5 yrs)"),
    (6, 10, "Mid-Level (6-10 yrs)"),
    (11, 15, "Senior (11-15 yrs)"),
    (16, 99, "Lead / Executive (16+ yrs)")
]
