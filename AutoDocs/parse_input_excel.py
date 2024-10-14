import pandas as pd 
import math

"""
grab investor numbers and their indices
for these indices, grab the corresponding rows
"""
def parse_input_excel(file_location, sheet_name, fund_info, inv_info, total_fund_info):
    df = pd.read_excel(file_location, sheet_name=sheet_name)

    #Read in fund info (first row)
    first_row = df.iloc[0]

    for col_name, value in first_row.items():
        if "Unnamed" in col_name: continue
        if "Blank" in col_name: continue
        fund_info[col_name] = str(value)

    fund_info["Re"] = "Re"
    fund_info["Notice Date"] = "Notice Date"
    fund_info["Due Date"] = "Due Date"
    
    #Read in investor info
    row_num = 1
    row = df.iloc[row_num]

    while ("Total Fund" not in row.values):
        index = 0
        inv_num = 0

        for col_name, value in row.items():
            if "Unnamed" in col_name: continue
            if "Blank" in col_name: continue
            if index == 0:
                inv_num = value
                inv_info[inv_num] = dict()
            else:
                inv_info[inv_num][col_name] = str(value)
            index += 1
        
        row_num += 1
        row = df.iloc[row_num]

    #Read in total fund info
    index = 0
    for col_name, value in row.items():
        if "Unnamed" in col_name: continue
        if "Blank" in col_name: continue
        if index < 2:
            index += 1
            continue
        total_fund_info[col_name] = str(value)

if __name__ == "__main__":
    fund_info = dict()
    inv_info = dict()
    total_fund_info = dict()
    parse_input_excel(r"C:\Users\ppark\OneDrive - GP Fund Solutions, LLC\Desktop\doc_gen_files\Allocation.xlsx", "Allocation", fund_info, inv_info, total_fund_info)

    for inv in inv_info:
        for j in inv_info[inv]:
            print(j)




