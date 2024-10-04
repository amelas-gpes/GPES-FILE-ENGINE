import pandas as pd 
from datetime import datetime

def parse_excel(excel):
    allocation = pd.read_excel(excel, sheet_name = "Allocation", skiprows=1)
    total_fund = dict()

    total_fund["Fund Name"] = str(allocation.at[3, "Partner Name"])

    total_fund["Investment"] = int(allocation.at[1, "Investment #1"])
    total_fund["Management Fees"] = int(allocation.at[1, "Gross Mgmt Fee"])
    total_fund["Partnership Expenses"] = int(allocation.at[1, "Pshp Exp"])
    total_fund["Total Amount Due"] = int(allocation.at[1, "Net Amount Due / (Payable)"])

    summary = pd.read_excel(excel, sheet_name = "Summary", skiprows=45)

    total_fund["Capital Commitment"] = int(summary.at[0, "Cumulative"])
    total_fund["Cumulative Capital Contributions"] = int(summary.at[14, "Cumulative"])
    total_fund["Remaining Capital Commitment"] = int(summary.at[18, "Cumulative"])

    summary = pd.read_excel(excel, sheet_name = "Summary", skiprows=0)

    total_fund["Re"] = "Capital Notice " + str(summary.iloc[0,1]) + " - Investment, Management Fees and Partnership Expenses"
    total_fund["Notice Date"] = summary.iloc[1,1].strftime('%m/%d/%Y')
    total_fund["Due Date"] = summary.iloc[3,1].strftime('%m/%d/%Y')

    
    
    print(total_fund)

    
    return total_fund
    """
    1. get fund info, thats hardcoded
    2. go through rows, grab investor info
    3. fill in pdf

    """

