# Anomaly Detection Script

![Python](https://img.shields.io/badge/language-Python-blue)

## Table of Contents
- [Overview](#overview)
- [Features](#features)
- [Installation](#installation)
- [Usage](#usage)
- [How to Use](#how-to-use)
- [Notes](#notes)
- [Input File Location](#input-file-location)
- [Output Files](#output-files)
- [Example](#example)
- [Potential Next Steps](#potential-next-steps)
- [Contributing](#contributing)

## Overview
This script performs anomaly detection on Excel (.xlsx) files using both rule-based and statistical (Isolation Forest) methods. It highlights detected anomalies, provides error analysis (confusion matrix, F1 score, accuracy, precision, recall), and generates files for missed, identified, and overpredicted cells.

## Features
- Rule-based and Isolation Forest anomaly detection
- Excel cell highlighting for anomalies
- Error analysis: confusion matrix, precision, recall, F1, accuracy
- Output files for error analysis and metrics
- Easy integration and extensibility

## Installation
1. Clone the repository:
   ```powershell
   git clone https://github.com/AyushK5ingh/anomaly-detection.git
   ```
2. Install dependencies:
   ```powershell
   pip install pandas numpy scikit-learn openpyxl
   ```

## Usage
1. Place your input Excel file in the project directory.
2. Set the `INPUT` variable in `anamoly.py` to your file name (without extension).
3. Run the script:
   ```powershell
   python anamoly.py
   ```
4. Review the generated output and error analysis files.

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

## Input File Location
If your input file is in a different directory, set `INPUT_FILE` to the full path:
```python
INPUT = "Train"
INPUT_FILE = r"D:/your_folder/Train.xlsx"
```
Otherwise, by default, `INPUT_FILE = INPUT + ".xlsx"` (in the current directory).

## Output Files
- `Train_output.xlsx`: Original data with anomalous cells replaced and highlighted.
- `Train_missed_and_identified.xlsx`: Error analysis with missed (red), identified (green), and overpredicted (yellow) cells.

## Example
If your input file is "Train.xlsx", set:
```python
INPUT = "Train"
```
Then run:
```powershell
python anamoly.py
```
Check the generated files for results and error analysis.

## Potential Next Steps
To further improve this project:
- Integrate Autoencoders or LSTM-based anomaly detection for complex/time-series data.
- Add a real-time API or dashboard for enhanced usability and accessibility for end users.

## Contributing
1. Fork the repository.
2. Create a new branch: `git checkout -b feature-name`
3. Make your changes.
4. Push your branch: `git push origin feature-name`
5. Create a pull request.

---

You can copy and adapt this structure for your project.
