import pandas as pd

# Read the CSV with no header
df = pd.read_csv('Customers Master Data for Clean up.csv', encoding='latin1', header=None)

clke_groups = []
fnkp_groups = []
biis_main_groups = []

i = 0
while i < len(df):
    row = df.iloc[i]
    if row[0] == 'B':
        # CLKE group
        if row[1] == 'CLKE':
            group = [row.tolist()]
            j = i + 1
            while j < len(df) and df.iloc[j][0] != 'D':
                group.append(df.iloc[j].tolist())
                j += 1
            if j < len(df) and df.iloc[j][0] == 'D':
                group.append(df.iloc[j].tolist())
                clke_groups.append(group)
                i = j  # move to D
        # FNKP group
        elif row[1] == 'FNKP':
            group = [row.tolist()]
            j = i + 1
            while j < len(df) and df.iloc[j][0] != 'D':
                group.append(df.iloc[j].tolist())
                j += 1
            if j < len(df) and df.iloc[j][0] == 'D':
                group.append(df.iloc[j].tolist())
                fnkp_groups.append(group)
                i = j  # move to D
        # BIIS MAIN group
        elif row[2] == 'BIIS MAIN':
            group = [row.tolist()]
            j = i + 1
            while j < len(df) and df.iloc[j][0] != 'A':
                group.append(df.iloc[j].tolist())
                j += 1
            if j < len(df) and df.iloc[j][0] == 'A':
                group.append(df.iloc[j].tolist())
                biis_main_groups.append(group)
                i = j  # move to A
    i += 1

# Helper to flatten groups for Excel
def flatten_groups(groups):
    flat = []
    for group in groups:
        flat.extend(group)
        flat.append([''] * len(group[0]))  # blank row between groups
    return flat

# Convert to DataFrames
df_clke = pd.DataFrame(flatten_groups(clke_groups))
df_fnkp = pd.DataFrame(flatten_groups(fnkp_groups))
df_biis_main = pd.DataFrame(flatten_groups(biis_main_groups))

# Write to Excel
with pd.ExcelWriter('cleaned_customers_data.xlsx', engine='openpyxl') as writer:
    df_clke.to_excel(writer, sheet_name='B-D_clke', index=False, header=False)
    df_fnkp.to_excel(writer, sheet_name='B-D_fnkp', index=False, header=False)
    df_biis_main.to_excel(writer, sheet_name='B-A_BIIS_MAIN', index=False, header=False)

print("Excel file 'cleaned_customers_data.xlsx' created with grouped records as requested.")