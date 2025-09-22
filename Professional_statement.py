import openpyxl
import re

def extract_professional_statement(input_file):
    wb_in = openpyxl.load_workbook(input_file)
    ws_in = wb_in.active
    if ws_in is None:
        raise ValueError("Input worksheet could not be loaded.")
    input_header_row = [cell.value for cell in ws_in[1]]
    try:
        bio_idx = input_header_row.index('Bio/Headshot')
    except ValueError:
        raise ValueError("'Bio/Headshot' column not found in input file.")
    bio_list = []
    for row in ws_in.iter_rows(min_row=2, values_only=True):
        bio = row[bio_idx]
        if bio is not None:
            # Remove unwanted symbols except , & ? . and preserve newlines
            bio = re.sub(r"[^\w\s,&?.\n]", "", bio)
            # Replace multiple spaces with a single space (but preserve newlines)
            bio = re.sub(r"[ ]{2,}", " ", bio)
        bio_list.append(bio)
    return bio_list
