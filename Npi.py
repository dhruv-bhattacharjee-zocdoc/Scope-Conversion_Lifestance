import openpyxl

def extract_npi(input_file):
    wb_in = openpyxl.load_workbook(input_file)
    ws_in = wb_in.active
    if ws_in is None:
        raise ValueError("Input worksheet could not be loaded.")
    input_header_row = [cell.value for cell in ws_in[1]]
    try:
        npi_idx = input_header_row.index('NPI')
    except ValueError:
        raise ValueError("'NPI' column not found in input file.")
    npi_list = []
    for row in ws_in.iter_rows(min_row=2, values_only=True):
        npi_list.append(row[npi_idx])
    return npi_list
