import pandas as pd
import re

# --- CONFIGURATION ---
INPUT_FILE = "Pos_Matched_Berry_Report.xlsx"  # The file generated from the previous step
OUTPUT_FILE = "Berries_Separated_Sheets_Pos.xlsx"

print("Phase 1: Loading unified master spreadsheet...")
df = pd.read_excel(INPUT_FILE)

# Clean up column headers just in case
df.columns = df.columns.astype(str).str.strip()

if "Berry" not in df.columns:
    print(f"[CRITICAL ERROR] 'Berry' column not found! Available columns: {list(df.columns)}")
    exit()

print("Phase 2: Grouping and sorting into target berry buckets...")

# Define our target sheets
target_sheets = ["Blueberry", "Blackberry", "Others"]

# Initialize a dictionary to hold dataframes for each sheet
sheets_data = {sheet: [] for sheet in target_sheets}

def determine_sheet_bucket(berry_value):
    """
    Scans the berry name prefix to route it to Blueberry, Blackberry, or Others.
    """
    if pd.isna(berry_value):
        return "Others"
        
    val_clean = str(berry_value).strip().lower()
    
    # Check if the matching word starts with 'blue_'
    if val_clean.startswith("blue_"):
        return "Blueberry"
        
    # Check if the matching word starts with 'black_'
    if val_clean.startswith("black_"):
        return "Blackberry"
        
    # Fallback bucket for anything else
    return "Others"

# Process row-by-row and distribute into buckets
for idx, row in df.iterrows():
    bucket = determine_sheet_bucket(row["Berry"])
    sheets_data[bucket].append(row)

print("Phase 3: Formatting sequence indexes and building multi-sheet workbook...")

# Write to the new multi-sheet Excel file
with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
    # Ensure sheets are written in the exact requested order
    ordered_sheet_list = ["Blueberry", "Blackberry", "Others"]
    
    for sheet_name in ordered_sheet_list:
        rows_list = sheets_data[sheet_name]
        
        if len(rows_list) > 0:
            # Convert back to a DataFrame
            sheet_df = pd.DataFrame(rows_list)
            
            # Reset and recalculate the s/n column sequentially from 1 for this specific tab
            if "s/n" in sheet_df.columns:
                sheet_df["s/n"] = range(1, len(sheet_df) + 1)
                
            # Write to its dedicated tab
            sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f" -> Sheet '{sheet_name}' populated successfully with {len(sheet_df)} rows.")
        else:
            # Create an empty template sheet with headers if no rows matched that type
            empty_df = pd.DataFrame(columns=df.columns)
            empty_df.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f" -> Sheet '{sheet_name}' created empty (No matching rows found).")

print(f"\n Success! All data split perfectly into 3 sheets. Saved to: '{OUTPUT_FILE}'")