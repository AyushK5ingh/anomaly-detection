# Anomaly Detection Script

## Overview
This script performs anomaly detection on an Excel (.xlsx) file using both rule-based and statistical (Isolation Forest) methods. It highlights detected anomalies in the output Excel file and provides detailed error analysis, including confusion matrix, F1 score, accuracy, precision, and recall for each column and for all columns combined.

## How to Use

### 1. Prepare your input Excel file
- Name your input file without the extension in the variable `INPUT` at the top of the script.
- Example: If your file is "Train.xlsx", set `INPUT = "Train"`.
- The input file must be in the same directory as the script.

### 2. Run the script
- Execute the script using Python 3.
- Example command:
  ```
  python anamoly.py
  ```
- The script will automatically process the file named `INPUT + ".xlsx"` (e.g., "Train.xlsx").

### 3. Output files generated
- `INPUT + "_output.xlsx"` (e.g., "Train_output.xlsx"):
  - Contains the original data with anomalous cells replaced by error labels.
  - Anomalous cells are highlighted in yellow.
- `INPUT + "_missed_and_identified.xlsx"` (e.g., "Train_missed_and_identified.xlsx"):
  - Contains error analysis:
    - Missed errors highlighted in red.
    - Identified errors highlighted in green.
    - Overpredicted errors highlighted in yellow.

### 4. Console Output
- The script prints performance metrics of the model on the dataset.
- It displays confusion matrix, F1 score, accuracy, precision, and recall for each individual column (excluding the first column, which is assumed to be a serial number and is ignored).
- It also prints combined metrics for all columns.

## Notes
- Ensure your input Excel file is properly formatted and located in the correct directory.
- The first column of the input file is ignored during anomaly detection and evaluation, as it is assumed to be a serial number.
- The script requires the following Python packages: `pandas`, `numpy`, `scikit-learn`, `openpyxl`.
- The script also depends on two additional Python files: `accuracy.py` and `errors.py`, which must be present in the same directory.

## Example
If your input file is "Train.xlsx", set:
```
INPUT = "Train"
```
Then run:
```
python anamoly.py
```

Check the generated files "Train_output.xlsx" and "Train_missed_and_identified.xlsx" for results and error analysis.

### Input File Location
If your input file is in a different directory, set `INPUT_FILE` to the full path:
```
INPUT = "Train"
INPUT_FILE = r"D:/your_folder/Train.xlsx"
```
Otherwise, by default, `INPUT_FILE = INPUT + ".xlsx"` (in the current directory).

## Potential Next Steps
To further improve this project:
 - Integration of **Autoencoders** or **LSTM-based anomaly detection** could further improve performance, especially in complex or time-series data.
 - Adding a **real-time API** or **dashboard integration** could enhance usability and accessibility for end users.

