import pandas as pd
import numpy as np
from sklearn.ensemble import IsolationForest
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import os

INPUT_FILE = "Train_Quality 1.xlsx"
OUTPUT_FILE = "Train_output.xlsx"
HIGHLIGHT_COLOR = "FFFF00" 

def infer_column_types(df, threshold=0.8):
    """
    Infer column types: numeric, datetime, categorical, or mixed.
    If no type is dominant (>threshold), mark as 'mixed'.
    """
    types = {}
    for col in df.columns:
        col_data = df[col].dropna()
        n = len(col_data)
        if n == 0:
            types[col] = 'unknown'
            continue
        num_valid = pd.to_numeric(col_data, errors='coerce').notnull().sum()
        dt_valid = pd.to_datetime(col_data, errors='coerce').notnull().sum()
        if num_valid / n >= threshold:
            types[col] = 'numeric'
        elif dt_valid / n >= threshold:
            types[col] = 'datetime'
        elif num_valid / n < (1 - threshold) and dt_valid / n < (1 - threshold):
            types[col] = 'categorical'
        else:
            types[col] = 'mixed'
    return types

def rule_based_anomalies(df, types):
    """
    Rule-based anomaly detection: missing, type mismatch, out-of-range (IQR), and strict length rule for pincode-like columns.
    """
    anomalies = pd.DataFrame('', index=df.index, columns=df.columns)
    anomalies[df.isnull()] = 'missing'
    for col in df.columns:
        if types[col] == 'numeric':
            coerced = pd.to_numeric(df[col], errors='coerce')
            mask = df[col].notnull() & coerced.isnull()
            anomalies.loc[mask, col] = 'type_mismatch'
            # Strict length rule: if 95% of non-missing values have the same length, only those are valid
            valid_mask = (anomalies[col] == '')
            valid_values = df.loc[valid_mask, col].astype(str)
            lengths = valid_values.str.len()
            if not lengths.empty:
                mode_length = lengths.mode()[0]
                mode_count = (lengths == mode_length).sum()
                if mode_count / len(lengths) >= 0.95:
                    # Only mode_length is allowed, others are length_inconsistent
                    for idx, val in valid_values.items():
                        if len(val) != mode_length:
                            anomalies.loc[idx, col] = 'length_inconsistent'
            # Out-of-range (IQR) (little less strict)
            col_data = coerced
            q1 = col_data.quantile(0.25)
            q3 = col_data.quantile(0.75)
            iqr = q3 - q1
            multiplier = 15  # Little less strict
            lower = q1 - multiplier * iqr
            upper = q3 + multiplier * iqr
            out_range = (col_data < lower) | (col_data > upper)
            for idx in df.index:
                if anomalies.loc[idx, col] == '' and out_range.loc[idx]:
                    anomalies.loc[idx, col] = 'out_of_range'
        elif types[col] == 'datetime':
            coerced = pd.to_datetime(df[col], errors='coerce')
            mask = df[col].notnull() & coerced.isnull()
            anomalies.loc[mask, col] = 'type_mismatch'
        elif types[col] == 'mixed':
            mask = df[col].notnull()
            anomalies.loc[mask, col] = 'type_mismatch'
    return anomalies

# def isolation_forest_anomalies(df, types, contamination=0.05):
#     """
#     Isolation Forest-based anomaly detection for numeric columns.
#     Only considers rows with valid numeric data (not missing/type_mismatch).
#     """
#     anomalies = pd.DataFrame('', index=df.index, columns=df.columns)
#     num_cols = [col for col in df.columns if types[col] == 'numeric']
#     if num_cols and len(df) > 10:
#         # Only use rows where all numeric columns are valid numbers
#         valid_mask = pd.DataFrame({
#             col: pd.to_numeric(df[col], errors='coerce').notnull()
#             for col in num_cols
#         }).all(axis=1)
#         X = df.loc[valid_mask, num_cols].apply(pd.to_numeric, errors='coerce')
#         if not X.empty:
#             iso = IsolationForest(contamination=contamination, random_state=42)
#             preds = iso.fit_predict(X)
#             outlier_rows = X.index[preds == -1]
#             for row in outlier_rows:
#                 for col in num_cols:
#                     # Only mark as statistical_outlier if not already missing/type_mismatch/out_of_range
#                     if anomalies.loc[row, col] == '':
#                         anomalies.loc[row, col] = 'statistical_outlier'
#     return anomalies

# def combine_anomalies(rule_anom, iso_anom):
#     """
#     Combine rule-based and isolation forest anomalies, prioritizing rule-based.
#     """
#     combined = rule_anom.copy()
#     for col in iso_anom.columns:
#         for idx in iso_anom.index:
#             if not combined.loc[idx, col] and iso_anom.loc[idx, col]:
#                 combined.loc[idx, col] = iso_anom.loc[idx, col]
#     return combined

def replace_and_highlight(df, anomalies, output_file):
    """
    Replace anomalous cells with error label and highlight them in the Excel output.
    """
    df_out = df.copy()
    for col in df.columns:
        for idx in df.index:
            if anomalies.loc[idx, col]:
                df_out.loc[idx, col] = anomalies.loc[idx, col]
    # Write to Excel
    df_out.to_excel(output_file, index=False)
    # Highlight anomalies
    wb = load_workbook(output_file)
    ws = wb.active
    fill = PatternFill(start_color=HIGHLIGHT_COLOR, end_color=HIGHLIGHT_COLOR, fill_type="solid")
    for i, col in enumerate(df.columns, 1):
        for j, idx in enumerate(df.index, 2):  # +2 for header
            if anomalies.loc[idx, col]:
                ws.cell(row=j, column=i).fill = fill
    wb.save(output_file)

def main():
    if not os.path.exists(INPUT_FILE):
        print(f"Input file not found: {INPUT_FILE}")
        return
    print(f"Reading {INPUT_FILE} ...")
    df = pd.read_excel(INPUT_FILE)
    print("Inferring column types ...")
    types = infer_column_types(df)
    print("Running rule-based anomaly detection ...")
    anomalies = rule_based_anomalies(df, types)
    # print("Running Isolation Forest anomaly detection ...")
    # iso_anom = isolation_forest_anomalies(df, types)
    # print("Combining results ...")
    # anomalies = combine_anomalies(rule_anom, iso_anom)
    print(f"Writing output to {OUTPUT_FILE} ...")
    replace_and_highlight(df, anomalies, OUTPUT_FILE)
    print("Done. Please check the output file for highlighted errors.")

if __name__ == "__main__":
    main()