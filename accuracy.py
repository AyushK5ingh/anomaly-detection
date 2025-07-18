# accuracy.py
# This script provides detailed error analysis, including confusion matrix, F1 score,
# accuracy, precision, and recall for each column and for all columns combined.

import openpyxl
import numpy as np
from sklearn.metrics import confusion_matrix, precision_score, recall_score, f1_score, accuracy_score

def get_highlight_matrix(excel_file, sheet_name=None):
    """
    Extracts a binary matrix from highlighted cells in an Excel file, ignoring the first column.
    Each cell is 1 if highlighted (non-default fill), else 0.
    Args:
        excel_file (str): Path to the Excel file.
        sheet_name (str, optional): Sheet name to read. If None, uses active sheet.
    Returns:
        np.ndarray: 2D array of binary highlight values.
    """
    wb = openpyxl.load_workbook(excel_file, data_only=True)
    ws = wb[sheet_name] if sheet_name else wb.active
    matrix = []
    for row in ws.iter_rows():
        # Ignore the first column
        matrix.append([
            1 if cell.fill and cell.fill.start_color and cell.fill.start_color.rgb != "00000000" and cell.fill.fill_type else 0
            for cell in row[1:]
        ])
    return np.array(matrix)

def compare_excel_highlights(file1, file2, sheet1=None, sheet2=None):
    """
    Compares highlighted cells between two Excel files and prints per-column and overall metrics.
    Args:
        file1 (str): Path to ground truth Excel file.
        file2 (str): Path to model output Excel file.
        sheet1, sheet2 (str, optional): Sheet names for each file.
    Prints:
        Confusion matrix, accuracy, precision, recall, F1 score for each column and overall.
    """
    arr1 = get_highlight_matrix(file1, sheet1)
    arr2 = get_highlight_matrix(file2, sheet2)
    if arr1.shape != arr2.shape:
        raise ValueError("Excel sheets have different shapes after ignoring the first column.")
    flat1 = arr1.flatten()
    flat2 = arr2.flatten()
    if not (set(flat1) <= {0, 1} and set(flat2) <= {0, 1}):
        raise ValueError("Both files must contain only binary highlight values (0 or 1).")

    # Per-column metrics
    print("\n--- Per-Column Metrics ---")
    for col in range(arr1.shape[1]):
        col1 = arr1[:, col]
        col2 = arr2[:, col]
        cm = confusion_matrix(col1, col2, labels=[0,1])
        precision = precision_score(col1, col2, zero_division=0)
        recall = recall_score(col1, col2, zero_division=0)
        f1 = f1_score(col1, col2, zero_division=0)
        accuracy = accuracy_score(col1, col2)
        misclassification = 1 - accuracy
        tn, fp, fn, tp = cm.ravel()
        print(f"\nColumn {col+1}:")
        print("Confusion Matrix (Predicted ↓ / Actual →):")
        print(f"             Actual 0    Actual 1")
        print(f"Pred 0    |   {tn:6}    |   {fn:6}   |  <-- True Neg, False Neg")
        print(f"Pred 1    |   {fp:6}    |   {tp:6}   |  <-- False Pos, True Pos")
        print(f"Accuracy:           {accuracy*100:.2f}%")
        print(f"Misclassification:  {misclassification*100:.2f}%")
        print(f"Precision:          {precision*100:.2f}%")
        print(f"Recall:             {recall*100:.2f}%")
        print(f"F1 Score:           {f1*100:.2f}%")

    # Combined metrics
    print("\n--- Combined Metrics (All Columns) ---")
    cm = confusion_matrix(flat1, flat2, labels=[0,1])
    precision = precision_score(flat1, flat2, zero_division=0)
    recall = recall_score(flat1, flat2, zero_division=0)
    f1 = f1_score(flat1, flat2, zero_division=0)
    accuracy = accuracy_score(flat1, flat2)
    misclassification = 1 - accuracy
    tn, fp, fn, tp = cm.ravel()
    print("Confusion Matrix (Predicted ↓ / Actual →):")
    print(f"             Actual 0    Actual 1")
    print(f"Pred 0    |   {tn:6}    |   {fn:6}   |  <-- True Neg, False Neg")
    print(f"Pred 1    |   {fp:6}    |   {tp:6}   |  <-- False Pos, True Pos")
    print(f"Accuracy:           {accuracy*100:.2f}%")
    print(f"Misclassification:  {misclassification*100:.2f}%")
    print(f"Precision:          {precision*100:.2f}%")
    print(f"Recall:             {recall*100:.2f}%")
    print(f"F1 Score:           {f1*100:.2f}%")

if __name__ == "__main__":
    # Usage: Enter the file name (without extension) for which you want to compare model output
    # Example: If your file is "Train.xlsx", enter "Train"
    input_base = input("Enter file name (without extension) to compare: ").strip()
    gt_file = input_base + ".xlsx"  # Ground truth file
    model_file = input_base + "_output.xlsx"  # Model output file
    compare_excel_highlights(gt_file, model_file)
