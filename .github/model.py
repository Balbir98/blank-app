import streamlit as st
import pandas as pd
import numpy as np
from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestRegressor, RandomForestClassifier
from sklearn.metrics import r2_score, classification_report
from fpdf import FPDF
import matplotlib.pyplot as plt
import tempfile

st.title("Adviser Income Forecast & Performance Flagging (Clean Predictors)")

uploaded_file = st.file_uploader("Upload adviser dataset (CSV)", type="csv")
if uploaded_file:
    # Load data
    df = pd.read_csv(uploaded_file)
    st.subheader("Data Preview")
    st.dataframe(df.head())

    # ---- Define clean predictors only ----
    feature_cols = [
        'Months Since Authorisation',
        'Events Attended Last 12 Months',
        '# CRM Cases All Time',
        'File Review Pass %',
        'Forecasted Revenue'
    ]

    # Ensure numeric types for predictors
    df['File Review Pass %'] = pd.to_numeric(df['File Review Pass %'], errors='coerce')
    df['Forecasted Revenue'] = pd.to_numeric(df['Forecasted Revenue'], errors='coerce')

    # Drop rows missing predictors or the target
    df_clean = df.dropna(subset=feature_cols + ['Annualised Projected Income Current Year'])

    # ---- Regression: forecast income using only true predictors ----
    X = df_clean[feature_cols]
    y_reg = df_clean['Annualised Projected Income Current Year']

    X_train, X_test, y_train, y_test = train_test_split(X, y_reg, test_size=0.2, random_state=42)
    reg = RandomForestRegressor(n_estimators=100, random_state=42)
    reg.fit(X_train, y_train)
    y_pred_reg = reg.predict(X_test)
    r2 = r2_score(y_test, y_pred_reg)
    importances_reg = pd.Series(reg.feature_importances_, index=feature_cols).sort_values(ascending=False)

    st.subheader("Clean Regression Results")
    st.write(f"R² with clean predictors: **{r2:.2f}**")
    st.bar_chart(importances_reg)

    # ---- Classification: underperformer flagging with same predictors ----
    df_clean['Underperformer'] = (
        df_clean['Total Commission Earned Year To Date'] <= df_clean['Total Commission Earned Last Year']
    ).astype(int)
    y_clf = df_clean['Underperformer']
    X2 = df_clean[feature_cols]

    X2_train, X2_test, y2_train, y2_test = train_test_split(X2, y_clf, test_size=0.2, random_state=42)
    clf = RandomForestClassifier(n_estimators=100, random_state=42)
    clf.fit(X2_train, y2_train)
    y2_pred = clf.predict(X2_test)

    report = classification_report(y2_test, y2_pred, output_dict=True)
    report_df = pd.DataFrame(report).transpose()
    importances_clf = pd.Series(clf.feature_importances_, index=feature_cols).sort_values(ascending=False)

    st.subheader("Clean Classification Results")
    st.dataframe(report_df)
    st.bar_chart(importances_clf)

    # ---- Generate charts for PDF ----
    # Regression feature importance chart
    plt.figure(figsize=(8, 4))
    importances_reg.plot.bar()
    plt.title("Feature Importance for Income Forecast")
    plt.ylabel("Importance")
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    reg_chart = tempfile.NamedTemporaryFile(suffix=".png", delete=False).name
    plt.savefig(reg_chart, dpi=150)
    plt.close()

    # Classification feature importance chart
    plt.figure(figsize=(8, 4))
    importances_clf.plot.bar(color='orange')
    plt.title("Feature Importance for Underperformance Flag")
    plt.ylabel("Importance")
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    clf_chart = tempfile.NamedTemporaryFile(suffix=".png", delete=False).name
    plt.savefig(clf_chart, dpi=150)
    plt.close()

    # ---- Build PDF report ----
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(0, 10, "Adviser Performance Insights", ln=True, align='C')

    pdf.set_font("Arial", '', 12)
    pdf.ln(5)
    pdf.cell(0, 8, f"Income Forecast R²: {r2:.2f}", ln=True)
    pdf.image(reg_chart, x=10, w=190)

    pdf.ln(5)
    pdf.cell(0, 8, "Underperformer Classification Report:", ln=True)
    for idx, row in report_df.head(3).iterrows():
        pdf.cell(0, 6, f" {idx}: precision {row['precision']:.2f}, recall {row['recall']:.2f}", ln=True)

    pdf.ln(5)
    pdf.image(clf_chart, x=10, w=190)

    # Full regression feature table on new page
    pdf.add_page()
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(0, 8, "Full Feature Importance (Regression)", ln=True)
    pdf.set_font("Arial", '', 10)
    for feature, coef in importances_reg.items():
        pdf.cell(0, 6, f"{feature:<40} {coef:.4f}", ln=True)

    # Download PDF
    pdf_bytes = pdf.output(dest='S').encode('latin-1')
    st.download_button(
        label="Download Insight Report (PDF)",
        data=pdf_bytes,
        file_name="adviser_insights_clean.pdf",
        mime="application/pdf"
    )
