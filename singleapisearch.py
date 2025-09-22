import os
import snowflake.connector
import pandas as pd
import numpy as np
import json

# Prompt the user for a single NPI number
npi = input("Enter the NPI number to query: ").strip()

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
    query = f"""
    SELECT * FROM merged_provider
    WHERE NPI:value::string = '{npi}'
    """
    cs.execute(query)
    results = cs.fetchall()
    columns = [desc[0] for desc in cs.description]
    if not results:
        print(f"No results found for NPI: {npi}")
    else:
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
        # Print the results to the terminal
        print(df_selected)
finally:
    cs.close()
    conn.close()