import pandas as pd
import re

# --- CONFIGURATION ---
file_path = r'/home/justsumant/Desktop/codeBase/gpf to excel/50_BOCHAHA_T.xlsx'
output_file = '50_BOCHAHA_T_Village_Wise_Sorted.xlsx'
target_column = 'शहर/गांव'
# ---------------------

def clean_sheet_name(name):
    if pd.isna(name) or str(name).strip() == "":
        return "Unknown"
    # Remove invalid characters [] : * ? / \
    clean_name = re.sub(r'[\[\]:*?/\\]', '', str(name))
    return clean_name[:31]  # Excel limit

try:
    print(f"Reading file: {file_path}...")
    df = pd.read_excel(file_path)

    if target_column not in df.columns:
        print(f"❌ Error: Column '{target_column}' not found.")
    else:
        print("Processing data...")

        unique_villages = df[target_column].dropna().unique()

        index_data = []
        sheets_to_write = []

        # Master sheet entry
        master_name = 'Master_Data'
        master_count = len(df)
        index_data.append({
            'Sheet Name': f'=HYPERLINK("#\'{master_name}\'!A1", "{master_name}")',
            'Voter Count': master_count
        })

        # Track sheet names to prevent duplicates
        used_names = {}

        for village in unique_villages:
            village_data = df[df[target_column] == village]
            count = len(village_data)

            # Base sheet name
            base_name = clean_sheet_name(village)

            # Ensure uniqueness (case-insensitive)
            name = base_name
            n = 1
            while name.lower() in used_names:
                name = f"{base_name}_{n}"
                n += 1

            used_names[name.lower()] = True

            # Hyperlink in Index sheet
            hyperlink = f'=HYPERLINK("#\'{name}\'!A1", "{name}")'
            index_data.append({'Sheet Name': hyperlink, 'Voter Count': count})

            sheets_to_write.append((name, village_data))

        index_df = pd.DataFrame(index_data)

        print(f"Writing {len(sheets_to_write) + 2} sheets with Hyperlinks...")

        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            # Index sheet
            index_df.to_excel(writer, sheet_name='Index', index=False)
            worksheet = writer.sheets['Index']
            worksheet.set_column('A:A', 30)
            worksheet.set_column('B:B', 15)

            # Master sheet
            df.to_excel(writer, sheet_name=master_name, index=False)

            # Village sheets
            for sheet_name, data in sheets_to_write:
                data.to_excel(writer, sheet_name=sheet_name, index=False)

        print(f"\n✅ Success! File saved as: {output_file}")
        print("   - Open 'Index' sheet and click any village name to jump to it.")

except FileNotFoundError:
    print(f"\n❌ Error: The file '{file_path}' was not found.")
except Exception as e:
    print(f"\n❌ An error occurred: {e}")
