"""
Synthetic Enterprise Dataset Generator
Generates realistic, statistically correlated employee records mimicking enterprise Kaggle datasets.
"""

import os
import random
import numpy as np
import pandas as pd
from datetime import datetime, timedelta

# Fix random seed for reproducibility
np.random.seed(42)
random.seed(42)

FIRST_NAMES = [
    "Aarav", "Aditi", "Amit", "Ananya", "Anjali", "Arjun", "Deepak", "Divya",
    "Gaurav", "Isha", "Karan", "Kavya", "Manish", "Meera", "Neha", "Nikhil",
    "Pooja", "Pranav", "Priya", "Rahul", "Rakesh", "Rhea", "Rohan", "Rohit",
    "Sameer", "Sanjay", "Shreya", "Sneha", "Suresh", "Tanvi", "Varun", "Vikram"
]

LAST_NAMES = [
    "Sharma", "Verma", "Patel", "Mishra", "Reddy", "Das", "Paul", "Sai",
    "Gupta", "Iyer", "Nair", "Mehta", "Bose", "Choudhury", "Joshi", "Kulkarni",
    "Singhania", "Chatterjee", "Banerjee", "Menon", "Kapoor", "Bhat", "Deshmukh"
]

DEPARTMENTS = {
    "Engineering": {
        "titles": ["Junior Software Engineer", "Software Engineer", "Senior Software Engineer", "Staff Engineer", "Engineering Manager", "VP of Engineering"],
        "base_salary": 65000,
        "exp_mult": 11000,
        "bonus_range": (10, 25)
    },
    "Data & AI": {
        "titles": ["Data Analyst", "Data Engineer", "Senior Data Scientist", "Lead ML Engineer", "Head of AI"],
        "base_salary": 70000,
        "exp_mult": 12000,
        "bonus_range": (12, 28)
    },
    "Product": {
        "titles": ["Associate Product Manager", "Product Manager", "Senior Product Manager", "Group Product Manager", "Director of Product"],
        "base_salary": 68000,
        "exp_mult": 11500,
        "bonus_range": (10, 25)
    },
    "Finance": {
        "titles": ["Financial Analyst", "Senior Financial Analyst", "Finance Controller", "Finance Director", "CFO"],
        "base_salary": 60000,
        "exp_mult": 9500,
        "bonus_range": (12, 30)
    },
    "Sales": {
        "titles": ["Sales Development Rep", "Account Executive", "Senior Account Executive", "Regional Sales Director", "Chief Commercial Officer"],
        "base_salary": 50000,
        "exp_mult": 8500,
        "bonus_range": (20, 45)  # Higher sales commissions
    },
    "Marketing": {
        "titles": ["Marketing Associate", "Content Strategist", "Growth Marketing Lead", "Marketing Director", "CMO"],
        "base_salary": 52000,
        "exp_mult": 8000,
        "bonus_range": (8, 20)
    },
    "Human Resources": {
        "titles": ["HR Coordinator", "HR Generalist", "Senior HR Business Partner", "Head of People & Culture"],
        "base_salary": 48000,
        "exp_mult": 7000,
        "bonus_range": (8, 18)
    },
    "Operations": {
        "titles": ["Operations Associate", "Supply Chain Analyst", "Operations Manager", "Director of Operations"],
        "base_salary": 50000,
        "exp_mult": 7500,
        "bonus_range": (7, 18)
    }
}

LOCATIONS = ["Bangalore", "Mumbai", "Hyderabad", "Delhi NCR", "Pune", "Remote"]
EDUCATION_LEVELS = ["Bachelor's", "Master's", "PhD", "Associate / Diploma"]
GENDERS = ["Female", "Male", "Non-Binary"]


def generate_enterprise_dataset(num_records: int = 1000) -> pd.DataFrame:
    """Generate realistic employee dataset with statistical distributions and correlation."""
    records = []
    start_date = datetime(2015, 1, 1)
    end_date = datetime(2026, 1, 1)

    for i in range(1, num_records + 1):
        emp_id = f"EMP-{i:04d}"
        first = random.choice(FIRST_NAMES)
        last = random.choice(LAST_NAMES)
        name = f"{first} {last}"
        gender = random.choices(GENDERS, weights=[0.45, 0.50, 0.05])[0]

        dept = random.choices(
            list(DEPARTMENTS.keys()),
            weights=[0.30, 0.15, 0.10, 0.12, 0.13, 0.08, 0.06, 0.06]
        )[0]
        dept_info = DEPARTMENTS[dept]

        # Experience (skewed towards junior to mid-level with realistic long tail)
        exp_years = int(np.clip(np.random.gamma(shape=2.5, scale=2.8), 0, 25))
        age = 22 + exp_years + random.randint(0, 5)

        # Title matched to experience
        num_titles = len(dept_info["titles"])
        title_idx = min(int(exp_years / 25 * num_titles + random.uniform(-0.5, 0.5)), num_titles - 1)
        title_idx = max(0, title_idx)
        job_title = dept_info["titles"][title_idx]

        # Education
        if dept in ["Data & AI", "Engineering"]:
            edu = random.choices(EDUCATION_LEVELS, weights=[0.45, 0.45, 0.08, 0.02])[0]
        else:
            edu = random.choices(EDUCATION_LEVELS, weights=[0.60, 0.35, 0.02, 0.03])[0]

        edu_bonus = {"Bachelor's": 1.0, "Master's": 1.10, "PhD": 1.25, "Associate / Diploma": 0.90}[edu]

        # Performance score (1.0 to 5.0, bell-curved around 3.6)
        perf_score = round(float(np.clip(np.random.normal(loc=3.6, scale=0.65), 1.0, 5.0)), 2)

        # Salary calculation with realistic noise and performance correlation
        base = dept_info["base_salary"]
        exp_adder = exp_years * dept_info["exp_mult"]
        perf_mult = 1.0 + (perf_score - 3.0) * 0.08  # ±16% adjustment based on performance
        noise = np.random.normal(1.0, 0.07)  # 7% standard deviation noise

        raw_salary = (base + exp_adder) * edu_bonus * perf_mult * noise
        # Round to nearest 1,000
        salary = int(round(raw_salary, -3))

        # Bonus
        bonus_pct = random.randint(*dept_info["bonus_range"])
        bonus_amount = int(salary * (bonus_pct / 100))
        total_comp = salary + bonus_amount

        # Location
        location = random.choice(LOCATIONS)

        # Joining Date
        random_days = random.randint(0, (end_date - start_date).days)
        join_date = (start_date + timedelta(days=random_days)).strftime("%Y-%m-%d")

        # Attrition Risk (Higher if high performer but underpaid, or very low performance)
        expected_salary_benchmark = (base + exp_adder) * edu_bonus
        comp_ratio = salary / max(1, expected_salary_benchmark)

        if comp_ratio < 0.90 and perf_score >= 4.0:
            attrition_risk = "High"  # Underpaid star talent
        elif comp_ratio < 0.95 or perf_score < 2.5:
            attrition_risk = "Medium"
        else:
            attrition_risk = "Low"

        records.append({
            "Employee_ID": emp_id,
            "Name": name,
            "Gender": gender,
            "Department": dept,
            "Job_Title": job_title,
            "Salary": salary,
            "Bonus_Percentage": bonus_pct,
            "Total_Compensation": total_comp,
            "Experience_Years": exp_years,
            "Age": age,
            "Education": edu,
            "Performance_Score": perf_score,
            "Location": location,
            "Joining_Date": join_date,
            "Attrition_Risk": attrition_risk
        })

    return pd.DataFrame(records)


def generate_minimal_dataset() -> pd.DataFrame:
    """Creates the original 8-record sample from the prompt/assignment for backward compatibility."""
    return pd.DataFrame({
        "Employee_ID": [101, 102, 103, 104, 105, 106, 107, 108],
        "Name": [
            "Rahul Sharma", "Anjali Verma", "Vikram Sai", "Priya Patel",
            "Amit Mishra", "Neha Reddy", "Rohan Das", "Sneha Paul"
        ],
        "Department": ["IT", "HR", "Finance", "IT", "Marketing", "Finance", "HR", "IT"],
        "Salary": [85000, 62000, 95000, 105000, 58000, 92000, 64000, 72000],
        "Experience_Years": [5, 3, 7, 9, 2, 6, 4, 3]
    })


if __name__ == "__main__":
    out_dir = os.path.dirname(os.path.abspath(__file__))
    os.makedirs(out_dir, exist_ok=True)

    # 1. Enterprise dataset
    df_enterprise = generate_enterprise_dataset(1200)
    ent_path = os.path.join(out_dir, "employees.csv")
    df_enterprise.to_csv(ent_path, index=False)
    print(f"Enterprise dataset created: {ent_path} ({len(df_enterprise)} records)")

    # 2. Minimal dataset
    df_min = generate_minimal_dataset()
    min_path = os.path.join(out_dir, "sample_minimal.csv")
    df_min.to_csv(min_path, index=False)
    print(f"Minimal sample dataset created: {min_path} ({len(df_min)} records)")
