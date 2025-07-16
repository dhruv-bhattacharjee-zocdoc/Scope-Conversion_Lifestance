import openpyxl

def extract_board_certification(input_file):
    wb_in = openpyxl.load_workbook(input_file)
    ws_in = wb_in.active
    if ws_in is None:
        raise ValueError("Input worksheet could not be loaded.")
    input_header_row = [cell.value for cell in ws_in[1]]
    try:
        cert_idx = input_header_row.index('Board Certification')
    except ValueError:
        raise ValueError("'Board Certification' column not found in input file.")
    cert_list = []
    for row in ws_in.iter_rows(min_row=2, values_only=True):
        cert_list.append(row[cert_idx])
    return cert_list 