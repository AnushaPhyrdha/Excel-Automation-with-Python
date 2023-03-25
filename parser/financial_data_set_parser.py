import numpy as np
import openpyxl as opxl
import pandas as pd

#load workbook
workbook = opxl.load_workbook(filename="F:\\FinancialsSampleData.xlsx", data_only=True)

#select the worksheet
worksheet = workbook['Financials']

#convert the cell references to cell values
worksheet = [[cell.value if cell.value is not None else np.nan for cell in row] for row in worksheet.iter_rows()]

#create the dataframe using the worksheet
financials = pd.DataFrame(worksheet)

#cleaning of working sheet
financials.replace(',', '', inplace=True, regex=True)
financials.replace({r'\s+$': '', r'^\s+': ''}, regex=True).replace(r'\n', ' ', regex=True)

#drop empty columns
financials = financials.dropna(how='all')
financials = financials.dropna(how='all', axis=1)

#replace nan with None
financials = financials.replace({np.nan: None})

#searching for start boundary
search_val = 'Account'
search_val_row, search_val_column = np.where(financials == search_val)
start_row = min(search_val_row)
start_column = min(search_val_column)

#converting merged header to list
list_super_top = financials.values.tolist()[start_row-1]
list_top = financials.values.tolist()[start_row]

#create unique headers
col_headers = []
super_col = list_super_top[0]
key_col = list_top[0]

for i in range(0, len(list_super_top)):
    new_col = ''

    if list_super_top[i] is not None:
       super_col = list_super_top[i].replace(' ', '_')
    
    if list_top[i] is not None:
        key_col = list_top[i].replace(' ', '_')
    
    if super_col is None:
        new_col = key_col
    else:
        new_col = super_col + '_' + key_col

    col_headers.append(new_col)

#drop unwanted header rows
financials = financials.drop(start_row-1)
financials = financials.drop(start_row)

#output dataframe creation
financials_final_df = pd.DataFrame()
for header, placeholder in zip(col_headers, list(financials.columns)):
    financials_final_df[header] = financials[placeholder].tolist()

print(financials_final_df)
