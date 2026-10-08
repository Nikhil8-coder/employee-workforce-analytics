# 💼 Enterprise Employee Intelligence & Compensation Analytics Suite

> **Transformation from Student Project to Enterprise Analytics Platform**  
> Elevated from a basic 8-record script into a production-grade workforce intelligence platform featuring statistical dispersion analytics, pay equity audits, machine learning compensation prediction, multi-format exports (Excel, HTML, CSV), and an interactive modern web dashboard.

---

## 🚀 Key Advancements Over the Starter Script

| Dimension | Original Project (PDF Assignment) | Enterprise Analytics Platform |
|---|---|---|
| **Dataset Scale** | Hardcoded 8 sample rows | Realistic 1,200+ employee Kaggle-grade dataset across 8 departments |
| **Data Intake** | Rigid CSV with exact headers | Multi-format ingestion (CSV, Excel, JSON) with schema alias mapping & cleaning |
| **Statistical Depth** | Basic mean salary & group-by count | Full executive dispersion (Mean, Median, Std Dev, Min/Max, IQR, 25th/75th percentiles) |
| **Workforce Auditing** | None | Demographic pay equity (gender pay gap), department equity ratios, flight risk |
| **Anomaly Detection** | None | Statistical IQR (Tukey fences) and Z-score outlier detection |
| **Predictive ML** | None | Random Forest Regression for fair-market salary forecasting & Compa-Ratio scoring ($R^2 > 0.95$) |
| **Reporting & Export** | Single plain CSV export | Multi-sheet formatted Excel (`.xlsx`), standalone Tailwind HTML brief, and CSV |
| **User Interfaces** | Basic terminal `input()` prompt | Dual Interface: Interactive Streamlit Web Dashboard + Feature-rich CLI with interactive wizard |
| **Code Quality** | Single script file | Modular Python package, type hints, docstrings, and automated unit test suite |

---

## 📂 Project Architecture

```
employee_analytics/
├── data/
│   ├── employees.csv              # Production Kaggle-scale dataset (1,200 records)
│   ├── sample_minimal.csv        # Backward-compatible 8-record starter dataset
│   └── data_generator.py         # Statistically correlated realistic dataset synthesizer
├── src/
│   ├── __init__.py               # Package exposure
│   ├── config.py                 # Constants, column aliases, UTF-8 console configuration
│   ├── data_loader.py            # Multi-format ingestion, schema normalization & validation
│   ├── analytics_engine.py       # Core analytics (KPIs, dispersion, pay equity, IQR outliers, filters)
│   ├── predictive_models.py      # ML Salary Predictor (Random Forest), Compa-Ratio & risk analyzer
│   └── reporter.py               # Exporters (CSV, styled multi-sheet Excel, standalone HTML report)
├── app.py                        # Streamlit Interactive Web Application
├── index.html                    # Portable standalone Web App (runs in any browser, zero dependencies)
├── cli.py                        # Advanced Command-Line Interface (Interactive & Batch modes)
├── start_app.bat                 # 1-Click Windows desktop launcher
├── Dockerfile                    # Containerization for Cloud VMs / Render / Railway
├── docker-compose.yml            # Docker Compose orchestration
├── export_standalone_html.py     # Generator for standalone portable web dashboard
├── tests/
│   └── test_analytics.py         # Automated unit test suite (100% pass)
├── requirements.txt              # Production dependencies
└── README.md                     # Comprehensive documentation
```

---

## 🛠️ Installation & Setup

1. **Navigate to the project directory:**
   ```bash
   cd C:\Users\nikhi\.gemini\antigravity\scratch\employee_analytics
   ```

2. **Install requirements:**
   ```bash
   pip install -r requirements.txt
   ```

---

## 🖥️ Running the Application

### 1. Interactive Streamlit Web Dashboard
Launch the web dashboard in your browser:
```bash
streamlit run app.py
```
**Dashboard Features:**
- **Executive Overview**: Dynamic KPI cards, salary distribution histogram with mean/median markers, department share donut chart.
- **Department & Role Deep Dive**: Salary spread boxplots, payroll allocation breakdown, and role compensation tables.
- **Pay Equity & Outlier Watchlist**: Demographic pay gap metrics, experience-vs-salary scatterplot, and 1.5×IQR anomaly detection.
- **ML Simulator**: "What-If" interactive compensation predictor for HR hiring decisions.
- **1-Click Export Center**: Instant download buttons for filtered CSV, JSON, styled Excel, and HTML briefs.

---

### 2. Command Line Interface (CLI)

#### A. Interactive Wizard Mode:
```bash
python cli.py --interactive
```
Provides a numbered terminal menu with formatted ASCII/Unicode tables for all operations.

#### B. Batch / Automated Pipeline Mode:
- **Executive Summary & High Earners Filter:**
  ```bash
  python cli.py --summary --threshold 80000 --export-csv high_earners.csv
  ```
- **Detect Compensation Outliers & Train ML:**
  ```bash
  python cli.py --outliers --train-ml
  ```
- **Generate Comprehensive Multi-Sheet Excel & HTML Brief:**
  ```bash
  python cli.py --export-excel workforce_full.xlsx --export-html executive_report.html
  ```

---

## 🧪 Running Unit Tests
To verify all calculations and modules:
```bash
python -m unittest discover tests
```
*Runs 6 comprehensive test cases verifying data loading, mathematical correctness, outlier detection, ML training, and multi-format reporting.*
