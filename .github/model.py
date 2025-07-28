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

    # ---- Ensure commission columns exist & fill missing for classification ----
    for col in ["Total Commission Earned Year To Date", "Total Commission Earned Last Year"]:
        if col not in df.columns:
            df[col] = 0
        else:
            df[col] = df[col].fillna(0)
    
    # ---- Ensure regression target exists & fill missing with 0 ----
    target_col = "Annualised Projected Income Current Year"
    if target_col not in df.columns and "Annualised Projected Income Current Year" not in df.columns:
        # if differently named adjust here...
        pass
    df[target_col] = df.get(target_col, pd.Series(0, index=df.index)).fillna(0)

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

    # ---- Prepare data for regression ----
    # Now only drop if predictors missing, not target
    reg_df = df.dropna(subset=feature_cols)
    n_reg = len(reg_df)
    st.write(f"Records for regression: {n_reg}")

    # ---- Regression Model ----
    if n_reg >= 2:
        X = reg_df[feature_cols]
        y_reg = reg_df[target_col]
        test_size = 0.2 if n_reg * 0.2 >= 1 else 1 / n_reg

        X_train, X_test, y_train, y_test = train_test_split(
            X, y_reg, test_size=test_size, random_state=42
        )
        reg = RandomForestRegressor(n_estimators=100, random_state=42)
        reg.fit(X_train, y_train)
        y_pred_reg = reg.predict(X_test)
        r2 = r2_score(y_test, y_pred_reg)
        importances_reg = pd.Series(reg.feature_importances_, index=feature_cols).sort_values(ascending=False)

        st.subheader("Clean Regression Results")
        st.write(f"R² with clean predictors: **{r2:.2f}**")
        st.bar_chart(importances_reg)
    else:
        st.error("Not enough data for regression.")

    # ---- Prepare data for classification ----
    clf_df = df.copy()
    n_clf = len(clf_df)
    st.write(f"Records for classification: {n_clf}")

    # ---- Classification Model ----
    if n_clf >= 2:
        df_clf = clf_df.copy()
        df_clf['Underperformer'] = (
            df_clf['Total Commission Earned Year To Date'] <= 
            df_clf['Total Commission Earned Last Year']
        ).astype(int)
        X2 = df_clf[feature_cols]
        y_clf = df_clf['Underperformer']
        test_size_clf = 0.2 if n_clf * 0.2 >= 1 else 1 / n_clf

        X2_train, X2_test, y2_train, y2_test = train_test_split(
            X2, y_clf, test_size=test_size_clf, random_state=42
        )
        clf = RandomForestClassifier(n_estimators=100, random_state=42)
        clf.fit(X2_train, y2_train)
        y2_pred = clf.predict(X2_test)

        report = classification_report(y2_test, y2_pred, output_dict=True)
        report_df = pd.DataFrame(report).transpose()
        importances_clf = pd.Series(clf.feature_importances_, index=feature_cols).sort_values(ascending=False)

        st.subheader("Clean Classification Results")
        st.dataframe(report_df)
        st.bar_chart(importances_clf)
    else:
        st.error("Not enough data for classification.")

    # ---- Generate charts & PDF ----
    if 'importances_reg' in locals() and 'importances_clf' in locals():
        # Regression chart
        plt.figure(figsize=(8,4))
        importances_reg.plot.bar()
        plt.title("Feature Importance for Income Forecast")
        plt.ylabel("Importance")
        plt.xticks(rotation=45, ha='right')
        plt.tight_layout()
        reg_chart = tempfile.NamedTemporaryFile(suffix=".png", delete=False).name
        plt.savefig(reg_chart, dpi=150)
        plt.close()

        # Classification chart
        plt.figure(figsize=(8,4))
        importances_clf.plot.bar(color='orange')
        plt.title("Feature Importance for Underperformance Flag")
        plt.ylabel("Importance")
        plt.xticks(rotation=45, ha='right')
        plt.tight_layout()
        clf_chart = tempfile.NamedTemporaryFile(suffix=".png", delete=False).name
        plt.savefig(clf_chart, dpi=150)
        plt.close()

        # Build PDF
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

        pdf.add_page()
        pdf.set_font("Arial", 'B', 14)
        pdf.cell(0, 8, "Full Feature Importance (Regression)", ln=True)
        pdf.set_font("Arial", '', 10)
        for feature, coef in importances_reg.items():
            pdf.cell(0, 6, f"{feature:<40} {coef:.4f}", ln=True)

        pdf_bytes = pdf.output(dest='S').encode('latin-1')
        st.download_button(
            label="Download Insight Report (PDF)",
            data=pdf_bytes,
            file_name="adviser_insights_clean.pdf",
            mime="application/pdf"
        )
