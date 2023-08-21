import pandas as pd
file_path = './Untitled spreadsheet (1).xlsx'
df = pd.read_excel(file_path)
duplicate_column = 'Ext User Id'
duplicates = df[df.duplicated(subset=duplicate_column, keep=False)]
last_duplicates = duplicates.duplicated(subset=duplicate_column, keep='last')
last_duplicates_to_remove = ~last_duplicates
cleaned_duplicates = duplicates[last_duplicates_to_remove]
cleaned_df = pd.concat([df, cleaned_duplicates]).drop_duplicates(keep=False)
cleaned_file_path = 'cleanedExcel.xlsx'
cleaned_df.to_excel(cleaned_file_path, index=False)
