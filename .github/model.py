import streamlit as st
import pandas as pd
import numpy as np
from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestRegressor, RandomForestClassifier
from sklearn.metrics import r2_score, classification_report
from fpdf import FPDF
from io import BytesIO

st.title("Adviser Income Forecast & Performance Flagging with PDF Export")

uploaded_file = st.file_uploader("Upload adviser dataset (CSV)", type="csv")
if uploaded_file:
    df = pd.read_csv(uploaded_file)
    st.subheader("Data Preview")
    st.dataframe(df.head())

    # Define features and targets
    feature_cols = [
        'Months Since Authorisation',
        'Number of Events Attended All Time',
        'Events Attended Last 12 Months',
        '# CRM Cases All Time',
        'File Review Pass Count',
        'File Review Fail Count',
        'File Review Pass %',
        'Forecasted Revenue',
        'Total Commission Earned Since Authorisation',
        'Total Commission Earned Last Year',
        'Total Commission Earned Year To Date'
    ]
    df_clean = df.dropna(subset=feature_cols + ['Annualised Projected Income Current Year'])

    # Regression for income forecast
    X = df_clean[feature_cols]
    y_reg = df_clean['Annualised Projected Income Current Year']
    X_train, X_test, y_train, y_test = train_test_split(X, y_reg, test_size=0.2, random_state=42)
    reg = RandomForestRegressor(n_estimators=50, random_state=42)
    reg.fit(X_train, y_train)
    y_pred_reg = reg.predict(X_test)
    r2 = r2_score(y_test, y_pred_reg)
    importances_reg = pd.Series(reg.feature_importances_, index=feature_cols).sort_values(ascending=False)

    # Classification for underperformers
    df_clean['Underperformer'] = (df_clean['Total Commission Earned Year To Date'] <= df_clean['Total Commission Earned Last Year']).astype(int)
    y_clf = df_clean['Underperformer']
    X2 = df_clean[feature_cols]
    X2_train, X2_test, y2_train, y2_test = train_test_split(X2, y_clf, test_size=0.2, random_state=42)
    clf = RandomForestClassifier(n_estimators=50, random_state=42)
    clf.fit(X2_train, y2_train)
    y2_pred = clf.predict(X2_test)
    report = classification_report(y2_test, y2_pred, output_dict=True)
    report_df = pd.DataFrame(report).transpose()
    importances_clf = pd.Series(clf.feature_importances_, index=feature_cols).sort_values(ascending=False)

    # Show metrics
    st.subheader("Regression R² Score")
    st.write(f"{r2:.2f}")
    st.subheader("Classification Report")
    st.dataframe(report_df)
    st.subheader("Regression Feature Importance")
    st.bar_chart(importances_reg)
    st.subheader("Classification Feature Importance")
    st.bar_chart(importances_clf)

    # Generate PDF
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(0, 10, "Adviser Performance Insights", ln=True, align='C')

    pdf.set_font("Arial", '', 12)
    pdf.ln(5)
    pdf.cell(0, 8, f"Income Forecast R²: {r2:.2f}", ln=True)
    pdf.ln(3)
    pdf.cell(0, 8, "Top 5 Features for Income Forecast:", ln=True)
    for feature, coef in importances_reg.head(5).items():
        pdf.cell(0, 6, f" - {feature}: {coef:.3f}", ln=True)

    pdf.ln(5)
    pdf.cell(0, 8, "Classification Report (Underperformer Flag):", ln=True)
    for idx, row in report_df.head(5).iterrows():
        pdf.cell(0, 6, f" {idx}: precision {row['precision']:.2f}, recall {row['recall']:.2f}", ln=True)

    pdf.ln(5)
    pdf.cell(0, 8, "Top 5 Features for Underperformance:", ln=True)
    for feature, coef in importances_clf.head(5).items():
        pdf.cell(0, 6, f" - {feature}: {coef:.3f}", ln=True)

    # Create download button
    pdf_bytes = pdf.output(dest='S').encode('latin-1')
    st.download_button(
        label="Download Insight Report (PDF)",
        data=pdf_bytes,
        file_name="adviser_insights.pdf",
        mime="application/pdf"
    )
