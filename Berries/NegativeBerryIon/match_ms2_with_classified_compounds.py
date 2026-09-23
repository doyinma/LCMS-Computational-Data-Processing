import pandas as pd

# --- CONFIGURATION ---
CLASSIFIED_VEG_FILE = "Berries_With_Dynamic_Classifications_Negative.xlsx"  # Multi-sheet classified workbook
EXTRACTED_MS2_FILE = "Berry_Extracted_MS2_Fragments_Negative.xlsx"          # Flat file from previous script
OUTPUT_FINAL_FILE = "(Negative) Classified_Berries_With_MS2_Negative.xlsx"         # Final annotated output workbook

def merge_ms2_into_classified_sheets():
    # 1. Load the flat MS2 reference data
    print(f"Loading extracted fragments from '{EXTRACTED_MS2_FILE}'...")
    df_ms2 = pd.read_excel(EXTRACTED_MS2_FILE)
    
    # Clean up key columns to prevent mismatching due to trailing spaces or case differences
    df_ms2['Name_Match'] = df_ms2['Name'].astype(str).str.strip().str.upper()
    df_ms2['Formula_Match'] = df_ms2['Formula'].astype(str).str.strip().str.upper()
    
    # Keep only the matching key columns and the final fragment values to avoid duplicate data columns
    df_ms2_lookup = df_ms2[['Name_Match', 'Formula_Match', 'MS2 Fragment Ions']].drop_duplicates(subset=['Name_Match', 'Formula_Match'])

    # 2. Open the multi-sheet vegetable workbook
    print(f"Reading multi-sheet classified workbook '{CLASSIFIED_VEG_FILE}'...")
    xl_veg = pd.ExcelFile(CLASSIFIED_VEG_FILE)
    
    processed_sheets = {}
    
    # 3. Iterate through every sheet and perform the match
    for sheet_name in xl_veg.sheet_names:
        print(f" -> Matching rows in sheet tab: '{sheet_name}'...")
        df_sheet = pd.read_excel(CLASSIFIED_VEG_FILE, sheet_name=sheet_name)
        
        if df_sheet.empty:
            processed_sheets[sheet_name] = df_sheet
            continue
            
        # Create temporary matching keys in the current vegetable sheet
        df_sheet['Name_Match'] = df_sheet['Name'].astype(str).str.strip().str.upper()
        df_sheet['Formula_Match'] = df_sheet['Formula'].astype(str).str.strip().str.upper()
        
        # Left merge joins all original columns with the matching MS2 fragment strings
        df_merged = pd.merge(
            df_sheet, 
            df_ms2_lookup, 
            on=['Name_Match', 'Formula_Match'], 
            how='left'
        )
        
        # Clean up by dropping the temporary string-matching helper columns
        df_merged = df_merged.drop(columns=['Name_Match', 'Formula_Match'])
        
        # If a compound doesn't have matching MS2 fragments in the file, fill it cleanly
        if 'MS2 Fragment Ions' in df_merged.columns:
            df_merged['MS2 Fragment Ions'] = df_merged['MS2 Fragment Ions'].fillna("No Fragments Found")
        
        # Save the updated dataframe back into our sheet memory dictionary
        processed_sheets[sheet_name] = df_merged

    # 4. Write all updated dataframes into a new multi-sheet workbook structure
    print(f"\nWriting matched results safely out to: '{OUTPUT_FINAL_FILE}'...")
    with pd.ExcelWriter(OUTPUT_FINAL_FILE, engine='openpyxl') as writer:
        for sheet, df_final in processed_sheets.items():
            df_final.to_excel(writer, sheet_name=sheet, index=False)
            
    print(f"\n Series Complete! The high-resolution dataset has been generated at: '{OUTPUT_FINAL_FILE}'")

# --- EXECUTE THE MERGE ENGINE ---
try:
    merge_ms2_into_classified_sheets()
except Exception as e:
    print(f"\n Pipeline failed during merge stage due to error: {str(e)}")