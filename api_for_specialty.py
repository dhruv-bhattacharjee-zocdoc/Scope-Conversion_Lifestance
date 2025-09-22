import os
import snowflake.connector
import pandas as pd
import numpy as np
import json
from yaspin import yaspin
from yaspin.spinners import Spinners

# Path to the Excel file with NPI list
excel_path = r"Excel Files/Npi-specialty.xlsx"
df_input = pd.read_excel(excel_path)

# Ensure 'NPI' column is string type and drop missing
npi_list = df_input['NPI'].dropna().astype(str).tolist()
total_npis = len(npi_list)
filled_count = 0
not_found_count = 0

# Connect to Snowflake using SSO (external browser authentication)
conn = snowflake.connector.connect(
    user="dhruv.bhattacharjee@zocdoc.com",
    account="OLIKNSY-ZOCDOC_001",
    warehouse="USER_QUERY_WH",
    database='CISTERN',
    schema='PROVIDER_PREFILL',  # updated schema
    role="PROD_OPS_PUNE_ROLE",
    authenticator='externalbrowser'
)

try:
    cs = conn.cursor()
    # Prepare a DataFrame to collect results
    results_df = pd.DataFrame(columns=['NPI', 'FIRST_NAME', 'LAST_NAME', 'SPECIALTIES'])
    with yaspin(Spinners.dots, text="Processing NPIs...") as spinner:
        for idx, npi in enumerate(npi_list, 1):
            query = f"""
            SELECT * FROM merged_provider
            WHERE NPI:value::string = '{npi}'
            """
            cs.execute(query)
            results = cs.fetchall()
            columns = [desc[0] for desc in cs.description]
            if not results:
                not_found_count += 1
                status = (f"Total Provider's NPIs in the Batch: {total_npis} | "
                          f"NPIs filled: {filled_count} | NPIs not found: {not_found_count}")
                spinner.text = f"{status} (Processing NPI {idx}/{total_npis})"
                continue
            df = pd.DataFrame(results, columns=columns)
            # Remove timezone info from all datetime columns
            for col in df.select_dtypes(include=['datetimetz']).columns:
                df[col] = df[col].dt.tz_localize(None)
            for col in df.columns:
                if df[col].dtype == 'object':
                    if df[col].apply(lambda x: hasattr(x, 'tzinfo') and x.tzinfo is not None).any():
                        df[col] = df[col].apply(lambda x: x.tz_localize(None) if hasattr(x, 'tzinfo') and x.tzinfo is not None else x)
            # Select only the required columns
            selected_columns = ['NPI', 'FIRST_NAME', 'LAST_NAME', 'SPECIALTIES']
            df_selected = df[selected_columns].copy()
            # Extract 'value' from JSON strings in each cell, with special handling for SPECIALTIES
            def extract_value(val, colname):
                if isinstance(val, str):
                    try:
                        parsed = json.loads(val)
                        if colname == 'SPECIALTIES' and isinstance(parsed, list) and len(parsed) > 0:
                            first = parsed[0]
                            if isinstance(first, dict) and 'value' in first:
                                return first['value']
                        if isinstance(parsed, dict) and 'value' in parsed:
                            return parsed['value']
                    except Exception:
                        pass
                return val
            for col in selected_columns:
                df_selected[col] = df_selected[col].apply(lambda x: extract_value(x, col))
            # Drop rows where SPECIALTIES is blank or null
            df_selected = df_selected[df_selected['SPECIALTIES'].notnull() & (df_selected['SPECIALTIES'] != '')]
            # Remove duplicate rows based on NPI and SPECIALTIES
            df_selected = df_selected.drop_duplicates(subset=['NPI', 'SPECIALTIES'], keep='first')
            # Only keep the first match for each NPI
            df_selected = df_selected.groupby('NPI').first().reset_index()
            results_df = pd.concat([results_df, df_selected], ignore_index=True)
            filled_count += 1
            status = (f"Total Provider's NPIs in the Batch: {total_npis} | "
                      f"NPIs filled: {filled_count} | NPIs not found: {not_found_count}")
            spinner.text = f"{status} (Processing NPI {idx}/{total_npis})"
        spinner.ok("✔")
finally:
    cs.close()
    conn.close()

# Update only the relevant columns in the original DataFrame, leaving 'NPI' untouched
if not results_df.empty:
    for col in ['FIRST_NAME', 'LAST_NAME', 'SPECIALTIES']:
        if col not in df_input.columns:
            df_input[col] = None
        # Map the results to the correct NPI rows
        update_map = results_df.set_index('NPI')[col].to_dict()
        df_input[col] = df_input['NPI'].astype(str).map(update_map).combine_first(df_input[col])
    df_input.to_excel(excel_path, index=False)
    print(f"\nResults written to {excel_path}")
    print(f"Total Provider's NPIs in the Batch: {total_npis}")
    print(f"NPIs filled: {filled_count}")
    print(f"NPIs not found: {not_found_count}")
else:
    print("No results found for any NPI in the input file.")
