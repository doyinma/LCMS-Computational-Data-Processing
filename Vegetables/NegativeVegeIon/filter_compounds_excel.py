import pandas as pd
import re

# --- CONFIGURATION ---
INPUT_FILE = "Final_Matched_Vegetables_Report.xlsx"  # The file generated from the previous step
OUTPUT_FILE = "Vegetables_Separated_Sheets.xlsx"

print("Phase 1: Loading unified master spreadsheet...")
df = pd.read_excel(INPUT_FILE)

# Clean up column headers just in case
df.columns = df.columns.astype(str).str.strip()

if "Vegetable" not in df.columns:
    print(f"[CRITICAL ERROR] 'Vegetable' column not found! Available columns: {list(df.columns)}")
    exit()

print("Phase 2: Grouping and sorting into target vegetable buckets...")

# Define our 6 target sheets (excluding 'Others' which handles the rest)
target_sheets = ["Spinach", "O-Spinach", "Broccoli", "O-Broccoli", "Kale", "O-Kale"]

# Initialize a dictionary to hold dataframes for each sheet
sheets_data = {sheet: [] for sheet in target_sheets}
sheets_data["Others"] = []

def determine_sheet_bucket(veg_value):
    """
    Standardizes the vegetable name and routes it to the correct sheet name.
    """
    if pd.isna(veg_value):
        return "Others"
        
    val_clean = str(veg_value).strip().lower()
    
    # Standardize common variations of "O-" prefixes (e.g., 'o_spinach', 'o-spinach', 'o spinach')
    val_clean = re.sub(r'^o\s*[-_\s]\s*', 'o-', val_clean)
    
    # Match Spinach groups
    if "spinach" in val_clean:
        return "O-Spinach" if val_clean.startswith("o-") else "Spinach"
        
    # Match Broccoli groups
    if "broccoli" in val_clean:
        return "O-Broccoli" if val_clean.startswith("o-") else "Broccoli"
        
    # Match Kale groups
    if "kale" in val_clean:
        return "O-Kale" if val_clean.startswith("o-") else "Kale"
        
    return "Others"

# Process row-by-row and distribute into buckets
for idx, row in df.iterrows():
    bucket = determine_sheet_bucket(row["Vegetable"])
    sheets_data[bucket].append(row)

print("Phase 3: Formatting sequence indexes and building multi-sheet workbook...")

# Write to the new multi-sheet Excel file
with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
    # Ensure the sheets are written in your exact requested order
    ordered_sheet_list = ["Spinach", "O-Spinach", "Broccoli", "O-Broccoli", "Kale", "O-Kale", "Others"]
    
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
            print(f" -> Sheet '{sheet_name}' populated successfully with {len(sheet_df)} compounds.")
        else:
            # Create an empty template sheet with headers if no compounds matched that type
            empty_df = pd.DataFrame(columns=df.columns)
            empty_df.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f" -> Sheet '{sheet_name}' created empty (No matching rows found).")

print(f"\n Success! All data split perfectly into 7 sheets. Saved to: '{OUTPUT_FILE}'")