import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
import os

def extract_professional_suffix(input_file):
    wb_in = openpyxl.load_workbook(input_file)
    ws_in = wb_in.active
    if ws_in is None:
        raise ValueError("Input worksheet could not be loaded.")
    input_header_row = [cell.value for cell in ws_in[1]]
    try:
        license_idx = input_header_row.index('License Type')
    except ValueError:
        raise ValueError("'License Type' column not found in input file.")
    suffix_lists = []
    for row in ws_in.iter_rows(min_row=2, values_only=True):
        cell_value = row[license_idx]
        if cell_value is not None:
            split_values = [v.strip() for v in str(cell_value).split(',')]
        else:
            split_values = [""]
        suffix_lists.append(split_values)
    return suffix_lists


def add_professional_suffix_dropdowns(output_file):
    wb = openpyxl.load_workbook(output_file)
    ws = wb['Provider']
    header_row = [cell.value for cell in ws[1]]
    for col_name in [f'Professional Suffix {i}' for i in range(1, 4)]:
        try:
            col_idx = header_row.index(col_name) + 1  # 1-based index for openpyxl
        except ValueError:
            continue  # Skip if the column is not found
        dv = DataValidation(type="list", formula1='=ValidationAndReference!$G$2:$G$511', allow_blank=True)
        max_row = ws.max_row
        col_letter = get_column_letter(col_idx)
        dv_range = f"{col_letter}2:{col_letter}{max_row}"
        dv.add(dv_range)
        ws.add_data_validation(dv)
    wb.save(output_file)

if __name__ == "__main__":
    output_file = os.path.join("Excel Files", "Output.xlsx")
    add_professional_suffix_dropdowns(output_file)
