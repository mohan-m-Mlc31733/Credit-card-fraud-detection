# src/dashboard.py
"""Streamlit dashboard for Credit Card Fraud Detection.
It loads the pre‑trained pipeline (same as FastAPI) and provides:
- Visualization of class distribution, confusion matrices, PR curve, threshold analysis.
- Interactive transaction input form to obtain real‑time prediction via the pipeline.
- Ability to refresh figures from the latest results directory.
"""

import os
import sys

# Ensure the project root (parent of this file's directory) is on sys.path so that
# `src` is importable as a package whether this file is run directly with
# `streamlit run src/dashboard.py` from the project root or imported as a module.
_SRC_DIR = os.path.dirname(os.path.abspath(__file__))
_PROJECT_ROOT = os.path.dirname(_SRC_DIR)
if _PROJECT_ROOT not in sys.path:
    sys.path.insert(0, _PROJECT_ROOT)

import pandas as pd
import streamlit as st
import numpy as np

# Import the pipeline class using an absolute package path
from src.fraud_detection import FraudDetectionPipeline

# Resolve the project root so result/data paths work regardless of cwd
_RESULTS_DIR = os.path.join(_PROJECT_ROOT, "results")

# Initialise pipeline (model trained on startup)
@st.cache_resource
def load_pipeline():
    pipeline = FraudDetectionPipeline()
    data_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "data", "creditcard.csv")
    df = pipeline.load_data(data_path)
    (X_train, y_train), (X_val, y_val), _ = pipeline.split_data(df)
    X_train_clean_scaled, X_train_scaled, X_val_scaled, X_test_scaled = pipeline.preprocess(
        X_train, y_train, X_val, pd.DataFrame()
    )
    pipeline.fit_model(X_train_clean_scaled)
    val_scores = pipeline.compute_anomaly_scores(X_val_scaled)
    pipeline.tune_threshold(val_scores, y_val)
    return pipeline

pipeline = load_pipeline()

st.title("⚠️ Credit Card Fraud Detection Dashboard")
st.sidebar.header("Navigation")
page = st.sidebar.radio("Go to", ["Overview", "Metrics & Plots", "Live Prediction"])

if page == "Overview":
    st.header("Project Overview")
    st.markdown(
        """
        This dashboard visualises the results of the Isolation Forest
        model for credit‑card fraud detection (IBM3201 project).
        *Baseline* uses the default scikit‑learn threshold (0.500).
        *Improved* uses a validation‑tuned threshold (≈0.625).
        """
    )
    # Show class distribution image
    class_img = os.path.join(_RESULTS_DIR, "class_distribution.png")
    if os.path.exists(class_img):
        st.image(class_img, caption="Class distribution", use_container_width=True)

elif page == "Metrics & Plots":
    st.header("Metrics & Visualisations")
    # Load CSV comparison table
    csv_path = os.path.join(_RESULTS_DIR, "metrics_comparison.csv")
    if os.path.exists(csv_path):
        df_metrics = pd.read_csv(csv_path)
        st.table(df_metrics)
    # Plot images
    plot_files = [
        "confusion_matrix_original.png",
        "confusion_matrix_improved.png",
        "precision_recall_curve.png",
        "threshold_analysis.png",
        "anomaly_score_distribution.png",
        "metric_comparison.png",
    ]
    for fname in plot_files:
        fpath = os.path.join(_RESULTS_DIR, fname)
        if os.path.exists(fpath):
            st.image(fpath, caption=fname.replace('_', ' ').replace('.png', ''), use_container_width=True)

else:  # Live Prediction
    st.header("Live Transaction Prediction")
    st.markdown("Enter the feature values for a single transaction (all columns except **Class**).")
    # Build a form dynamically from the pipeline's expected columns
    feature_names = [
        "Time", "V1", "V2", "V3", "V4", "V5", "V6", "V7", "V8", "V9", "V10",
        "V11", "V12", "V13", "V14", "V15", "V16", "V17", "V18", "V19", "V20",
        "V21", "V22", "V23", "V24", "V25", "V26", "V27", "V28", "Amount",
    ]
    input_data = {}
    for col in feature_names:
        input_data[col] = st.number_input(col, value=0.0, format="%f")
    if st.button("Predict"):
        df_input = pd.DataFrame([input_data])
        try:
            X_scaled = pipeline.scaler.transform(df_input)
            score = float(pipeline.compute_anomaly_scores(X_scaled)[0])
            pred = int(pipeline.predict_with_threshold(np.array([score]), pipeline.tuned_threshold)[0])
            st.success(f"**Anomaly score:** {score:.4f}")
            st.info(f"**Prediction (1=Fraud, 0=Genuine):** {pred}")
            st.caption(f"Threshold used: {pipeline.tuned_threshold:.4f}")
        except Exception as e:
            st.error(f"Error during prediction: {e}")

st.sidebar.markdown("---")
st.sidebar.caption("Dashboard built with Streamlit 1.30.0 • Backend model from src/fraud_detection.py")
