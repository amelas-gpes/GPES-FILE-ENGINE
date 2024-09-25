import pandas as pd 

"""
create map from BI Update
go through each BI, replace.
"""

#get mappings
gp_gc_map = {}
gp_cn_map = {}

df = pd.read_excel("BI Update.xlsx", sheet_name = "GP Update")

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


#replace GP CRM
df2 = pd.read_excel("AEA GP CRM UPLOAD.xlsx", sheet_name="Investor")

cn = df2["Fund Name"]

new_cn = []

for i in cn:
    new_cn.append(gp_cn_map[i])

"""
inv_short = {
    "Capital Partners I" : "CaPaI",
    "Capital Partners II" : "CaPaII",
    "Managers Fund I" : "MaFuI",
    "Growth Fund V LP" : "GrFuVLP",
    "Growth Fund VI LP" : "GrFuVILP1",
    "Growth Fund VII LP" : "GrFuVILP2",
    "Growth Fund VIII LP" : "GrFuVILP3",
    "Strategic Equity Partners LP" : "StEqPaLP",
    "Funding Program LP" : "FuPrLP",
    "Strategic Partners II LP" : "StPaIILP",
    "Debt Partners II LP" : "DePaIILP1",
    "Debt Partners III LP" : "DePaIILP2",
    "Debt Partners IV LP" : "DePaIVLP",
    "Real Estate III LP" : "ReEsIILP",
    "Real Estate IV LP" : "ReEsIVLP",
    "Strategic Equity Fund II LP" : "StEqFuIILP",
    "Debt Partners V LP" : "DePaVLP",
    "Debt Partners VI LP" : "DePaVILP1",
    "Debt Partners VII LP" : "DePaVILP2"
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

df2.to_excel("modified_GP_CRM.xlsx", index=False)
