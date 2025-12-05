import pandas as pd

# ---- CONFIG ----
input_file = "feedback-2025-11-27.csv"
input_excel = "searchstrlegal.xlsx"
output_file = "feedback-legal.csv"
column_name = "userName"
value_to_keep = "luca.bruelhart@unique.ai"
excel_column = "search_string"
# -----------------

df = pd.read_csv(input_file)
filtered_df = df[df[column_name] == value_to_keep]
filtered_df = filtered_df.iloc[:-1]

excel_df = pd.read_excel(input_excel)

if excel_column not in excel_df.columns:
    raise ValueError(f"Column '{excel_column}' not found in Excel file.")

search_strings = excel_df[excel_column]
filtered_df[excel_column] = search_strings.values[:len(filtered_df)]

filtered_df.to_csv(output_file, index=False)

print(f"Filtered CSV saved to: {output_file}")
