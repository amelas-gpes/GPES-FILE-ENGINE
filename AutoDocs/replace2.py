import pandas as pd 

"""
create map from BI Update
go through each BI, replace.
"""

#get mappings
gp_gc_map = {}
gp_cn_map = {}

df = pd.read_excel("BI Update.xlsx", sheet_name = "LP Update")

key_col = df["Original Company Name"]

index = 0

for i in key_col:
    replacement = df.at[index, "New Company Name"]
    gp_cn_map[i] = replacement 
    index += 1

key_col = df["Original Company Group Code"]

index = 0

for i in key_col:
    replacement = df.at[index, "New Company Group Code"]
    gp_gc_map[i] = replacement 
    index += 1


#replace CRM
df2 = pd.read_excel("AEA LP CRM UPLOAD.xlsx", sheet_name="Investor")

cn = df2["Fund Name"]

new_cn = []

for i in cn:
    new_cn.append(gp_cn_map[i])

"""
inv_short = {
    "ABC Fund VI Offshore" : "ABFuVIOf",
    "ABC Fund VI" : "ABFuVI",
    "ABC Fund V" : "ABFuV",
    "ABC Fund IV" : "ABFuIV",
    "ABC Fund III" : "ABFuII2",
    "ABC Real Estate Fund I Program" : "ABReEsFuIPr",
    "ABC Real Estate Fund II Program" : "ABReEsFuIIPr",
    "ABC Fund II" : "ABFuII1",
    "ABC Fund II Offshore" : "ABFuIIOf",
    "ABC Fund II Aggregator" : "ABFuIIAg",
    "ABC Fund II NAV Facility" : "ABFuIINAFa"
}


gp_inv_shorts = []

index = 0

for i in new_cn:
    name = df2.at[index, "Legal Name"]
    name = name.split()
    fn = name[0][:2]
    ln = name[1][:2]

    gp_inv_shorts.append(fn + ln + inv_short[i])

    index += 1

df2["Investor Code"] = gp_inv_shorts
"""
df2["Fund Name"] = new_cn

df2.to_excel("modified_LP_CRM.xlsx", index=False)
