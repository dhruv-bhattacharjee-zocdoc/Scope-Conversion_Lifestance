import openpyxl

def extract_headshot(input_file):
    wb_in = openpyxl.load_workbook(input_file)
    ws_in = wb_in.active
    if ws_in is None:
        raise ValueError("Input worksheet could not be loaded.")
    input_header_row = [cell.value for cell in ws_in[1]]
    try:
        headshot_idx = input_header_row.index('Headshot URL')
    except ValueError:
        raise ValueError("'Headshot URL' column not found in input file.")
    headshot_list = []
    for row in ws_in.iter_rows(min_row=2, values_only=True):
        headshot_list.append(row[headshot_idx])
    return headshot_list 