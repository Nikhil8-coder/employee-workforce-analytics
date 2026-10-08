"""
Enterprise Workforce Intelligence & Analytics Dashboard
A real-world interactive Streamlit web application for workforce data analysis,
predictive compensation modeling, and executive reporting.
"""

import os
import sys
import io
import pandas as pd
import numpy as np
import altair as alt
import streamlit as st

# Setup path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.data_loader import DataLoader
from src.analytics_engine import AnalyticsEngine
from src.predictive_models import CompensationPredictor
from src.reporter import ReportGenerator
from src.config import DEFAULT_SALARY_THRESHOLD, CURRENCY_SYMBOL

# Streamlit Page Configuration
st.set_page_config(
    page_title="Workforce Intelligence & Compensation Analytics",
    page_icon="💼",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for executive presentation
st.markdown("""
<style>
    .main-header {
        font-size: 2.2rem;
        font-weight: 800;
        color: #1E293B;
        margin-bottom: 0.2rem;
    }
    .sub-header {
        font-size: 1.05rem;
        color: #64748B;
        margin-bottom: 1.5rem;
    }
    .metric-card {
        background-color: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 12px;
        padding: 18px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    .badge-chip {
        display: inline-block;
        padding: 2px 8px;
        border-radius: 12px;
        font-size: 11px;
        font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)


@st.cache_data
def load_dataset(file_path: str):
    """Cached loader for instant performance."""
    loader = DataLoader(file_path)
    success, msg = loader.load()
    if not success:
        return None, msg
    return loader.get_dataframe(), msg


def main():
    # Header Banner
    st.markdown('<div class="main-header">💼 Workforce Intelligence & Analytics Platform</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-header">Enterprise Compensation Analysis &bull; Demographic Equity &bull; Predictive ML Benchmarking</div>', unsafe_allow_html=True)

    # ---------------------------------------------------------
    # Sidebar: Data Source & Filtering Engine
    # ---------------------------------------------------------
    st.sidebar.title("🎛️ Control Center")

    base_dir = os.path.dirname(os.path.abspath(__file__))
    enterprise_data_path = os.path.join(base_dir, "data", "employees.csv")
    minimal_data_path = os.path.join(base_dir, "data", "sample_minimal.csv")

    dataset_source = st.sidebar.radio(
        "Select Dataset Source:",
        ["Enterprise Dataset (1,200 records)", "Original Minimal Sample (8 records)", "Upload Custom CSV/Excel"]
    )

    df_raw = None
    if dataset_source == "Enterprise Dataset (1,200 records)":
        df_raw, msg = load_dataset(enterprise_data_path)
    elif dataset_source == "Original Minimal Sample (8 records)":
        df_raw, msg = load_dataset(minimal_data_path)
    else:
        uploaded_file = st.sidebar.file_uploader("Upload dataset (.csv, .xlsx, .json):", type=["csv", "xlsx", "json"])
        if uploaded_file is not None:
            temp_path = os.path.join(base_dir, "data", f"temp_upload_{uploaded_file.name}")
            with open(temp_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
            df_raw, msg = load_dataset(temp_path)
        else:
            st.info("Please upload a file or switch to Enterprise/Minimal dataset.")
            return

    if df_raw is None:
        st.error(f"Error loading dataset: {msg}")
        return

    st.sidebar.markdown("---")
    st.sidebar.subheader("🔍 Dynamic Filters")

    # Filter: Search
    search_q = st.sidebar.text_input("Search Name, ID, or Role:", placeholder="e.g. Priya or Engineer")

    # Filter: Department
    all_depts = sorted(df_raw["Department"].unique().tolist())
    selected_depts = st.sidebar.multiselect("Department(s):", all_depts, default=all_depts)

    # Filter: Salary Threshold & Range
    sal_min = float(df_raw["Salary"].min())
    sal_max = float(df_raw["Salary"].max())
    default_step = 5000.0 if sal_max > 50000 else 1000.0

    salary_range = st.sidebar.slider(
        f"Base Salary Range ({CURRENCY_SYMBOL}):",
        min_value=sal_min,
        max_value=sal_max,
        value=(sal_min, sal_max),
        step=default_step
    )

    # Filter: Experience
    exp_filter = None
    if "Experience_Years" in df_raw.columns:
        min_exp = float(df_raw["Experience_Years"].min())
        max_exp = float(df_raw["Experience_Years"].max())
        if max_exp > min_exp:
            exp_filter = st.sidebar.slider("Experience Years:", min_value=min_exp, max_value=max_exp, value=(min_exp, max_exp), step=1.0)

    # Filter: Performance
    perf_filter = None
    if "Performance_Score" in df_raw.columns:
        perf_filter = st.sidebar.slider("Minimum Performance Rating:", min_value=1.0, max_value=5.0, value=1.0, step=0.1)

    # Apply filters using AnalyticsEngine
    full_engine = AnalyticsEngine(df_raw)
    filtered_df = full_engine.filter_records(
        min_salary=salary_range[0],
        max_salary=salary_range[1],
        departments=selected_depts,
        min_experience=exp_filter[0] if exp_filter else None,
        max_experience=exp_filter[1] if exp_filter else None,
        min_performance=perf_filter if perf_filter else None,
        search_query=search_q
    )

    st.sidebar.info(f"Showing **{len(filtered_df):,}** of **{len(df_raw):,}** records ({len(filtered_df)/len(df_raw)*100:.1f}%)")

    if filtered_df.empty:
        st.warning("No employee records match the active filter criteria. Adjust the sidebar sliders.")
        return

    active_engine = AnalyticsEngine(filtered_df)
    kpis = active_engine.get_executive_summary()
    reporter = ReportGenerator(output_dir="reports")

    # ---------------------------------------------------------
    # Main Tabs Layout
    # ---------------------------------------------------------
    tab_kpi, tab_dept, tab_equity, tab_ml, tab_data = st.tabs([
        "📊 Executive KPI Overview",
        "🏢 Department & Role Analytics",
        "⚖️ Pay Equity & Outliers",
        "🤖 ML Compensation Simulator",
        "📁 Data Explorer & Exports"
    ])

    # ---------------------------------------------------------
    # TAB 1: EXECUTIVE KPI OVERVIEW
    # ---------------------------------------------------------
    with tab_kpi:
        st.subheader("Global Workforce Summary")

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Workforce Headcount", f"{kpis['headcount']:,}", f"{kpis['department_count']} Departments")
        c2.metric("Total Annual Payroll", f"{CURRENCY_SYMBOL}{kpis['total_payroll']:,.0f}", "Annualised Base")
        c3.metric("Average Base Salary", f"{CURRENCY_SYMBOL}{kpis['average_salary']:,.2f}", f"Median: {CURRENCY_SYMBOL}{kpis['median_salary']:,.0f}")
        c4.metric("Salary Spread (Std Dev)", f"{CURRENCY_SYMBOL}{kpis['std_salary']:,.0f}", f"IQR: {CURRENCY_SYMBOL}{kpis['iqr_salary']:,.0f}")

        st.markdown("<br>", unsafe_allow_html=True)
        col_left, col_right = st.columns([1.2, 1])

        with col_left:
            st.markdown("##### 📈 Salary Distribution & Density")
            hist = alt.Chart(filtered_df).mark_bar(opacity=0.75, color="#3B82F6").encode(
                x=alt.X("Salary:Q", bin=alt.Bin(maxbins=25), title=f"Salary ({CURRENCY_SYMBOL})"),
                y=alt.Y("count():Q", title="Number of Employees"),
                tooltip=["count():Q"]
            ).properties(height=320)

            mean_rule = alt.Chart(pd.DataFrame({'val': [kpis['average_salary']]})).mark_rule(
                color='#EF4444', strokeWidth=2, strokeDash=[4, 4]
            ).encode(x='val:Q')

            median_rule = alt.Chart(pd.DataFrame({'val': [kpis['median_salary']]})).mark_rule(
                color='#10B981', strokeWidth=2
            ).encode(x='val:Q')

            st.altair_chart(hist + mean_rule + median_rule, use_container_width=True)
            st.caption("🔴 Red dashed line = Average Salary | 🟢 Green solid line = Median Salary")

        with col_right:
            st.markdown("##### 💼 Departmental Headcount Allocation")
            dept_counts = filtered_df["Department"].value_counts().reset_index()
            dept_counts.columns = ["Department", "Headcount"]
            donut = alt.Chart(dept_counts).mark_arc(innerRadius=60).encode(
                theta=alt.Theta("Headcount:Q"),
                color=alt.Color("Department:N", scale=alt.Scale(scheme="tableau10")),
                tooltip=["Department:N", "Headcount:Q"]
            ).properties(height=320)
            st.altair_chart(donut, use_container_width=True)

    # ---------------------------------------------------------
    # TAB 2: DEPARTMENT & ROLE ANALYTICS
    # ---------------------------------------------------------
    with tab_dept:
        st.subheader("Departmental Deep Dive")

        dept_summary = active_engine.get_department_summary()
        st.dataframe(
            dept_summary.style.format({
                "Average_Salary": f"{CURRENCY_SYMBOL}{{:,.2f}}",
                "Median_Salary": f"{CURRENCY_SYMBOL}{{:,.2f}}",
                "Min_Salary": f"{CURRENCY_SYMBOL}{{:,.2f}}",
                "Max_Salary": f"{CURRENCY_SYMBOL}{{:,.2f}}",
                "Total_Payroll": f"{CURRENCY_SYMBOL}{{:,.0f}}",
                "Headcount_Share_Pct": "{:.1f}%",
                "Payroll_Share_Pct": "{:.1f}%"
            }),
            use_container_width=True
        )

        col_b1, col_b2 = st.columns(2)
        with col_b1:
            st.markdown("##### 📊 Average Salary by Department")
            bar_dept = alt.Chart(dept_summary).mark_bar(color="#6366F1").encode(
                x=alt.X("Average_Salary:Q", title=f"Average Salary ({CURRENCY_SYMBOL})"),
                y=alt.Y("Department:N", sort="-x", title=""),
                tooltip=["Department:N", "Average_Salary:Q", "Headcount:Q"]
            ).properties(height=300)
            st.altair_chart(bar_dept, use_container_width=True)

        with col_b2:
            st.markdown("##### 📦 Salary Spread Boxplot by Department")
            box_dept = alt.Chart(filtered_df).mark_boxplot(extent="min-max").encode(
                x=alt.X("Department:N", title=""),
                y=alt.Y("Salary:Q", title=f"Salary ({CURRENCY_SYMBOL})"),
                color="Department:N"
            ).properties(height=300)
            st.altair_chart(box_dept, use_container_width=True)

        if "Job_Title" in filtered_df.columns:
            st.markdown("---")
            st.markdown("##### 👔 Top Compensated Roles")
            role_summary = active_engine.get_role_summary()
            if not role_summary.empty:
                st.dataframe(
                    role_summary.head(10).style.format({
                        "Average_Salary": f"{CURRENCY_SYMBOL}{{:,.2f}}",
                        "Median_Salary": f"{CURRENCY_SYMBOL}{{:,.2f}}",
                        "Min_Salary": f"{CURRENCY_SYMBOL}{{:,.2f}}",
                        "Max_Salary": f"{CURRENCY_SYMBOL}{{:,.2f}}"
                    }),
                    use_container_width=True
                )

    # ---------------------------------------------------------
    # TAB 3: PAY EQUITY & OUTLIERS
    # ---------------------------------------------------------
    with tab_equity:
        st.subheader("Workforce Equity & Compensation Anomaly Audits")

        eq_col, sc_col = st.columns([1, 1.2])

        with eq_col:
            st.markdown("##### ⚖️ Demographic Pay Equity")
            equity = active_engine.get_pay_equity_summary()
            if "gender_stats" in equity:
                st.dataframe(
                    equity["gender_stats"].style.format({
                        "Average_Salary": f"{CURRENCY_SYMBOL}{{:,.2f}}",
                        "Median_Salary": f"{CURRENCY_SYMBOL}{{:,.2f}}"
                    }),
                    use_container_width=True
                )
                ratio = equity.get('female_to_male_ratio_pct', 100.0)
                st.metric("Female-to-Male Pay Parity Ratio", f"{ratio}%", f"Gap: {equity.get('pay_gap_pct', 0.0)}%")
            else:
                st.info("No gender data available in current dataset.")

        with sc_col:
            st.markdown("##### ⏳ Experience vs Salary Correlation")
            if "Experience_Years" in filtered_df.columns:
                scatter = alt.Chart(filtered_df).mark_circle(size=60, opacity=0.6).encode(
                    x=alt.X("Experience_Years:Q", title="Years of Experience"),
                    y=alt.Y("Salary:Q", title=f"Salary ({CURRENCY_SYMBOL})"),
                    color="Department:N",
                    tooltip=["Name:N", "Department:N", "Experience_Years:Q", "Salary:Q"]
                ).properties(height=320)
                st.altair_chart(scatter, use_container_width=True)

        st.markdown("---")
        st.markdown("##### 🚨 Statistical Compensation Outlier Detection (Tukey IQR Method)")
        outliers_df = active_engine.detect_outliers(method="iqr", threshold=1.5)
        st.write(f"Detected **{len(outliers_df)}** records outside standard 1.5×IQR boundary thresholds.")

        if not outliers_df.empty:
            outlier_cols = [c for c in ["Employee_ID", "Name", "Department", "Job_Title", "Salary", "Experience_Years", "Outlier_Type", "Outlier_Score"] if c in outliers_df.columns]
            st.dataframe(
                outliers_df[outlier_cols].style.format({"Salary": f"{CURRENCY_SYMBOL}{{:,.2f}}"}),
                use_container_width=True
            )

    # ---------------------------------------------------------
    # TAB 4: ML COMPENSATION SIMULATOR
    # ---------------------------------------------------------
    with tab_ml:
        st.subheader("Machine Learning Compensation Predictor")
        st.markdown("Train an ensemble Random Forest regression model on live workforce data to forecast competitive compensation for new positions or promotions.")

        col_train, col_sim = st.columns([1, 1.3])

        predictor = CompensationPredictor()
        train_success, train_msg = predictor.train(df_raw)

        with col_train:
            st.markdown("##### 🎯 Model Training Performance")
            if train_success:
                st.success(train_msg)
                st.metric("Explained Variance (R²)", f"{predictor.r2:.3f}")
                st.metric("Mean Absolute Error (MAE)", f"{CURRENCY_SYMBOL}{predictor.mae:,.2f}")
            else:
                st.warning(train_msg)

        with col_sim:
            st.markdown("##### 🧮 Interactive 'What-If' Salary Calculator")
            if train_success:
                sim_exp = st.number_input("Candidate Experience (Years):", min_value=0.0, max_value=35.0, value=5.0, step=0.5)
                sim_dept = st.selectbox("Department:", all_depts)
                sim_perf = st.slider("Expected Performance Score (1-5):", min_value=1.0, max_value=5.0, value=3.8, step=0.1)
                sim_edu = st.selectbox("Education Level:", ["Bachelor's", "Master's", "PhD", "Associate / Diploma"])

                if st.button("🚀 Calculate Benchmark Compensation", type="primary"):
                    predicted_sal = predictor.predict_salary(sim_exp, sim_dept, sim_perf, sim_edu)
                    lower_band = round(predicted_sal * 0.90, -2)
                    upper_band = round(predicted_sal * 1.10, -2)

                    st.markdown(f"""
                    <div style="background-color: #EEF2FF; border: 2px solid #6366F1; border-radius: 12px; padding: 20px; text-align: center; margin-top: 15px;">
                        <span style="font-size: 14px; color: #4338CA; font-weight: 600;">RECOMMENDED BASE COMPENSATION</span>
                        <h2 style="color: #1E1B4B; margin: 10px 0; font-size: 2.2rem;">{CURRENCY_SYMBOL}{predicted_sal:,.0f}</h2>
                        <span style="font-size: 13px; color: #6366F1;">Target Market Range (90% - 110%): <b>{CURRENCY_SYMBOL}{lower_band:,.0f}</b> — <b>{CURRENCY_SYMBOL}{upper_band:,.0f}</b></span>
                    </div>
                    """, unsafe_allow_html=True)

    # ---------------------------------------------------------
    # TAB 5: DATA EXPLORER & EXPORTS
    # ---------------------------------------------------------
    with tab_data:
        st.subheader("Data Explorer & Enterprise Export Hub")

        st.dataframe(filtered_df, use_container_width=True)

        st.markdown("---")
        st.markdown("##### 📥 Export Intelligence Packages")

        exp_c1, exp_c2, exp_c3, exp_c4 = st.columns(4)

        # 1. CSV Download
        csv_buffer = io.StringIO()
        filtered_df.to_csv(csv_buffer, index=False)
        exp_c1.download_button(
            label="📄 Download CSV",
            data=csv_buffer.getvalue(),
            file_name="workforce_filtered_export.csv",
            mime="text/csv",
            use_container_width=True
        )

        # 2. JSON Download
        json_buffer = io.StringIO()
        filtered_df.to_json(json_buffer, orient="records", indent=2)
        exp_c2.download_button(
            label="🔤 Download JSON",
            data=json_buffer.getvalue(),
            file_name="workforce_filtered_export.json",
            mime="application/json",
            use_container_width=True
        )

        # 3. Excel Multi-Sheet Generation
        excel_buffer = io.BytesIO()
        sheets = {
            "Executive_KPIs": pd.DataFrame([kpis]),
            "Department_Summary": active_engine.get_department_summary(),
            "Filtered_Workforce": filtered_df,
            "Outlier_Anomalies": active_engine.detect_outliers(method="iqr")
        }
        with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
            for s_name, s_df in sheets.items():
                s_df.to_excel(writer, sheet_name=s_name[:31], index=False)

        exp_c3.download_button(
            label="📊 Download Multi-Sheet Excel",
            data=excel_buffer.getvalue(),
            file_name="workforce_intelligence_suite.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

        # 4. Generate HTML Report
        html_report_path = reporter.generate_html_report(
            kpis,
            active_engine.get_department_summary(),
            active_engine.filter_records(min_salary=DEFAULT_SALARY_THRESHOLD),
            active_engine.detect_outliers(method="iqr")
        )
        with open(html_report_path, "r", encoding="utf-8") as f:
            html_bytes = f.read()

        exp_c4.download_button(
            label="🌐 Download HTML Executive Brief",
            data=html_bytes,
            file_name="executive_workforce_brief.html",
            mime="text/html",
            use_container_width=True
        )


if __name__ == "__main__":
    main()
