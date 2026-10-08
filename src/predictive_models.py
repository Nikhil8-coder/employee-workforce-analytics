"""
Machine Learning Predictive Analytics Suite
Includes ML Salary Estimator, Compa-Ratio Parity Benchmark, and Talent Attrition Risk Analyzer.
"""

from typing import Dict, Any, Tuple, Optional
import numpy as np
import pandas as pd
from sklearn.ensemble import RandomForestRegressor
from sklearn.pipeline import Pipeline
from sklearn.compose import ColumnTransformer
from sklearn.preprocessing import OneHotEncoder, StandardScaler
from sklearn.metrics import r2_score, mean_absolute_error


class CompensationPredictor:
    """Predictive Machine Learning model for market-competitive compensation analysis."""

    def __init__(self):
        self.pipeline: Optional[Pipeline] = None
        self.r2: float = 0.0
        self.mae: float = 0.0
        self.feature_names = []
        self.is_trained: bool = False

    def train(self, df: pd.DataFrame) -> Tuple[bool, str]:
        """Trains an ensemble regression model on historical employee compensation data."""
        if len(df) < 8:
            return False, "Insufficient data rows (minimum 8 required) to train ML predictor."

        # Target
        y = df["Salary"].values

        # Select available predictor features
        numeric_features = []
        categorical_features = []

        if "Experience_Years" in df.columns:
            numeric_features.append("Experience_Years")
        if "Performance_Score" in df.columns:
            numeric_features.append("Performance_Score")

        if "Department" in df.columns:
            categorical_features.append("Department")
        if "Education" in df.columns:
            categorical_features.append("Education")

        if not numeric_features and not categorical_features:
            return False, "No valid predictor features found in dataset."

        X = df[numeric_features + categorical_features]

        transformers = []
        if numeric_features:
            transformers.append(("num", StandardScaler(), numeric_features))
        if categorical_features:
            transformers.append(("cat", OneHotEncoder(handle_unknown="ignore", sparse_output=False), categorical_features))

        preprocessor = ColumnTransformer(transformers=transformers)

        self.pipeline = Pipeline(steps=[
            ("preprocessor", preprocessor),
            ("regressor", RandomForestRegressor(n_estimators=100, max_depth=8, random_state=42))
        ])

        try:
            self.pipeline.fit(X, y)
            preds = self.pipeline.predict(X)
            self.r2 = round(float(r2_score(y, preds)), 3)
            self.mae = round(float(mean_absolute_error(y, preds)), 2)
            self.is_trained = True
            return True, f"ML model trained successfully. R² Score: {self.r2}, MAE: ₹{self.mae:,.2f}"
        except Exception as e:
            return False, f"Failed to train ML model: {str(e)}"

    def predict_salary(
        self,
        experience_years: float,
        department: str,
        performance_score: float = 3.5,
        education: str = "Bachelor's"
    ) -> float:
        """Predicts fair market salary given candidate attributes."""
        if not self.is_trained or self.pipeline is None:
            raise RuntimeError("Model has not been trained yet.")

        input_data = pd.DataFrame([{
            "Experience_Years": experience_years,
            "Performance_Score": performance_score,
            "Department": department,
            "Education": education
        }])

        prediction = self.pipeline.predict(input_data)[0]
        return round(float(prediction), -2)

    def benchmark_workforce(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Calculates Compa-Ratio (Actual Salary / Expected Market Salary) for all employees.
        Flags underpaid and overpaid staff.
        """
        if not self.is_trained or self.pipeline is None:
            self.train(df)

        benchmarked = df.copy()
        features = [col for col in ["Experience_Years", "Performance_Score", "Department", "Education"] if col in benchmarked.columns]
        X = benchmarked[features]

        predicted = self.pipeline.predict(X)
        benchmarked["Benchmark_Predicted_Salary"] = predicted.round(2)
        benchmarked["Compa_Ratio"] = (benchmarked["Salary"] / np.maximum(1, benchmarked["Benchmark_Predicted_Salary"])).round(2)

        conditions = [
            (benchmarked["Compa_Ratio"] < 0.90),
            (benchmarked["Compa_Ratio"] >= 0.90) & (benchmarked["Compa_Ratio"] <= 1.10),
            (benchmarked["Compa_Ratio"] > 1.10)
        ]
        choices = ["Under-Compensated (<90%)", "Market Aligned (90-110%)", "Premium Compensated (>110%)"]
        benchmarked["Equity_Status"] = np.select(conditions, choices, default="Market Aligned")

        return benchmarked
