import openpyxl

def extract_education(input_file):
    wb_in = openpyxl.load_workbook(input_file)
    ws_in = wb_in.active
    if ws_in is None:
        raise ValueError("Input worksheet could not be loaded.")
    input_header_row = [cell.value for cell in ws_in[1]]
    try:
        education_idx = input_header_row.index('Highest Level of Education')
    except ValueError:
        raise ValueError("'Highest Level of Education' column not found in input file.")
    try:
        school_idx = input_header_row.index('School')
    except ValueError:
        school_idx = None
    education_list = []
    for row in ws_in.iter_rows(min_row=2, values_only=True):
        education_val = row[education_idx] if education_idx < len(row) else None
        school_val = row[school_idx] if school_idx is not None and school_idx < len(row) else None
        if education_val is not None and str(education_val).strip() != "":
            if school_val is not None and str(school_val).strip() != "":
                combined = f"{education_val}, {school_val}"
            else:
                combined = str(education_val)
        else:
            combined = ""
        education_list.append(combined)
    return education_list
