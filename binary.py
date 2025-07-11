import openpyxl
import csv

def excel_highlight_to_binary(input_file, output_file, sheet_name=None):
    wb = openpyxl.load_workbook(input_file, data_only=True)
    ws = wb[sheet_name] if sheet_name else wb.active

    binary_matrix = []
    for row in ws.iter_rows():
        binary_row = []
        for cell in row:
            # Check if the cell has a fill (highlight)
            if cell.fill and cell.fill.start_color and cell.fill.start_color.rgb != "00000000" and cell.fill.fill_type:
                binary_row.append(1)
            else:
                binary_row.append(0)
        binary_matrix.append(binary_row)    # Save as CSV
    with open(output_file, 'w', newline='') as f:
        writer = csv.writer(f)
        writer.writerows(binary_matrix)

if __name__ == "__main__":
    # Example usage
    excel_highlight_to_binary("Train_Quality 1.xlsx", "Input.csv")
    excel_highlight_to_binary("Train_output.xlsx", "Output.csv")
    print("done")