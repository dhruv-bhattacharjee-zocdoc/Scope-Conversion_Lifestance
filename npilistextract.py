import pandas as pd
import os

# File paths
def resource_path(relative):
    base = os.path.dirname(__file__)
    return os.path.join(base, relative)

input_path = resource_path(r"Excel Files/Input.xlsx")
snowflake_path = resource_path(r"Excel Files/snowflake.xlsx")

# Read the NPI column from Input.xlsx
input_df = pd.read_excel(input_path)
if 'NPI' in input_df.columns:
    npi_df = input_df[['NPI']].dropna()
    # Write to the default sheet in snowflake.xlsx, overwriting any existing data
    npi_df.to_excel(snowflake_path, index=False)
else:
    print(f"Column 'NPI' not found in {input_path}")
