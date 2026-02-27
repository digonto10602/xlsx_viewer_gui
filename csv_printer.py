import pandas as pd

path = "RMG Leads USA.xlsx"  # <-- change this

# Load the 3rd sheet (0-based index: 2)
df = pd.read_excel(path, sheet_name=3)

cols = [
    "Full Name",
    "Email",
    "Lead's LinkedIn URL",
    "Job Title",
    "Company Name",
    "Company Website",
    "Company LinkedIn Profile URL",
]

# Print only those columns
#print(df.loc[:, cols].to_string(index=False))

#print(df)

for _, row in df[cols].iterrows():
    printed_any = False
    for c in cols:
        v = row[c]
        if pd.notna(v):
            v = str(v).strip()
            if v:  # also skip empty strings
                print(f"{c}: {v}")
                printed_any = True
    if printed_any:
        print("-----------------------")