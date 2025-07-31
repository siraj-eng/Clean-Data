import pandas as pd

# List of all categories to group by (column 1)
categories = [
    'CLKE', 'FNKP', 'BGES', 'BGRS', 'BRSD', 'BDIS', 'MAIN',
    'BIIS', 'BKIS', 'BNIS', 'BMIS', 'PAYT'
]

# Read the CSV with no header
df = pd.read_csv('Customers Master Data for Clean up.csv', encoding='latin1', header=None)

# Dictionary to hold groups for each category
category_groups = {cat: [] for cat in categories}

i = 0
while i < len(df):
    row = df.iloc[i]
    if row[0] == 'B' and row[1] in categories:
        cat = row[1]
        group = [row.tolist()]
        j = i + 1
        while j < len(df) and df.iloc[j][0] != 'D':
            group.append(df.iloc[j].tolist())
            j += 1
        if j < len(df) and df.iloc[j][0] == 'D':
            group.append(df.iloc[j].tolist())
            category_groups[cat].append(group)
            i = j  # move to D
    i += 1

# Helper to flatten groups for Excel
def flatten_groups(groups):
    flat = []
    for group in groups:
        flat.extend(group)
        flat.append([''] * len(group[0]))  # blank row between groups
    return flat

# Write to Excel, one sheet per category
with pd.ExcelWriter('cleaned_customers_data.xlsx', engine='openpyxl') as writer:
    for cat, groups in category_groups.items():
        df_cat = pd.DataFrame(flatten_groups(groups))
        df_cat.to_excel(writer, sheet_name=cat, index=False, header=False)

print("Excel file 'cleaned_customers_data.xlsx' created with grouped records for each category.")