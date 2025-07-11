import pandas as pd
from sklearn.metrics import confusion_matrix, precision_score, recall_score, f1_score, accuracy_score

def compare_csv_sheets(file1, file2):
    # Read both CSV files
    df1 = pd.read_csv(file1, header=None)
    df2 = pd.read_csv(file2, header=None)

    # Ensure both DataFrames have the same shape
    if df1.shape != df2.shape:
        raise ValueError("Files have different shapes.")

    # Flatten the DataFrames and compare element-wise
    arr1 = df1.values.flatten()
    arr2 = df2.values.flatten()

    # Check for binary input (0 or 1)
    if not (set(arr1) <= {0, 1} and set(arr2) <= {0, 1}):
        raise ValueError("Both files must contain only binary values (0 or 1).")

    # Metrics
    cm = confusion_matrix(arr1, arr2)
    precision = precision_score(arr1, arr2, zero_division=0)
    recall = recall_score(arr1, arr2, zero_division=0)
    f1 = f1_score(arr1, arr2, zero_division=0)
    accuracy = accuracy_score(arr1, arr2)
    misclassification = 1 - accuracy

    tn, fp, fn, tp = cm.ravel()

    print("Confusion Matrix (Predicted ↓ / Actual →):")
    print(f"             Actual 0    Actual 1")
    print(f"Pred 0    |   {tn:6}    |   {fn:6}   |  <-- True Neg, False Neg")
    print(f"Pred 1    |   {fp:6}    |   {tp:6}   |  <-- False Pos, True Pos")
    # print()
    # print(f"True Negatives (TN):  {tn}")
    # print(f"False Positives (FP): {fp}")
    # print(f"False Negatives (FN): {fn}")
    # print(f"True Positives (TP):  {tp}")
    print()
    print(f"Accuracy:           {accuracy*100:.2f}%")
    print(f"Misclassification:  {misclassification*100:.2f}%")
    print(f"Precision:          {precision*100:.2f}%")
    print(f"Recall:             {recall*100:.2f}%")
    print(f"F1 Score:           {f1*100:.2f}%")

if __name__ == "__main__":
    compare_csv_sheets('Input.csv','Output.csv')