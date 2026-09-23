import openpyxl
import pandas as pd
import sys

# --- CONFIGURATION ---
INPUT_FILE = "ExportedCompounds_Features_MS2_Positive.xlsx"  # Change to your actual file name
OUTPUT_FILE = "Extracted_MS2_Fragments_Positive.xlsx"    # Final flat output file

# Exact column numbers based on your layout
NAME_COL = 4      # Column D
FORMULA_COL = 5   # Column E
FRAGMENT_COL = 8  # Column H (Nested m/z values)

def is_blue_row(cell):
    """
    Evaluates cell style properties to safely identify parent compound rows.
    Handles standard hex fills, theme tints, and ARGB format variations.
    """
    if not cell or not cell.fill or not cell.fill.start_color:
        return False
        
    color_obj = cell.fill.start_color
    color_type = color_obj.type
    rgb_str = str(color_obj.rgb).upper().strip()
    
    if color_type == 'rgb' and rgb_str:
        if len(rgb_str) == 8:
            rgb_str = rgb_str[2:]
            
        # Target typical blue spectrums used by Compound Discoverer exports
        if rgb_str.startswith(('33', '4F', '81', 'A0', 'B', 'C', 'D', 'E')) or 'BLUE' in rgb_str:
            return True
            
    if color_type == 'theme':
        if color_obj.theme in [4, 5]:
            return True
            
    return False

def extract_nested_ms2_with_tracking(file_path):
    print("Loading workbook into memory (this can take a moment for large files)...")
    wb = openpyxl.load_workbook(file_path, data_only=True)
    flattened_records = []
    
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        total_rows = ws.max_row
        print(f"\n Started processing sheet tab: '{sheet_name}' ({total_rows} total rows found)")
        
        # State tracking memory variables
        active_name = None
        active_formula = None
        collected_mzs = []
        compounds_counter = 0
        
        # Loop row by row, starting at row 2 to bypass your main table headers
        for row in range(2, total_rows + 1):
            
            # --- PROGRESS TRACKER ---
            if row % 100 == 0 or row == total_rows:
                # Flushes the progress text directly to the console terminal line
                sys.stdout.write(f"\r   -> Progress: Analyzing row {row}/{total_rows}...")
                sys.stdout.flush()
                
            name_cell = ws.cell(row=row, column=NAME_COL)
            formula_cell = ws.cell(row=row, column=FORMULA_COL)
            mz_cell = ws.cell(row=row, column=FRAGMENT_COL)
            
            name_str = str(name_cell.value).strip() if name_cell.value is not None else ""
            formula_str = str(formula_cell.value).strip() if formula_cell.value is not None else ""
            mz_str = str(mz_cell.value).strip() if mz_cell.value is not None else ""
            
            # LOOKAHEAD FILTER: If the NEXT row's column H says exactly "m/z", this row is a parent!
            next_row_has_sub_header = False
            if row < total_rows:
                next_cell_val = str(ws.cell(row=row+1, column=FRAGMENT_COL).value).lower().strip()
                if next_cell_val == "m/z":
                    next_row_has_sub_header = True

            # TRIGGER 1: We hit a MAIN compound row (Identified by Blue fill OR structural lookahead)
            if is_blue_row(name_cell) or next_row_has_sub_header:
                
                # If we already have an active compound group in memory, save it before moving on
                if active_name or active_formula:
                    comma_separated_mzs = ", ".join(collected_mzs) if collected_mzs else "No Fragments"
                    flattened_records.append({
                        "Name": active_name,
                        "Formula": active_formula,
                        "MS2 Fragment Ions": comma_separated_mzs
                    })
                
                # Overwrite memory with the brand new parent compound row details
                active_name = name_str
                active_formula = formula_str
                collected_mzs = []  # Completely reset the bucket for this molecule
                compounds_counter += 1
                
            # TRIGGER 2: We are inside the nested rows under the active parent compound
            else:
                # 1. Skip the row that represents the sub-table text header ("m/z")
                if mz_str.lower() == "m/z":
                    continue
                
                # 2. Extract and append the actual numeric m/z fragment values
                if mz_str != "" and mz_str.lower() != "none":
                    collected_mzs.append(mz_str)
                    
        # Save the final compound block sitting at the very bottom of the workbook sheet
        if active_name or active_formula:
            comma_separated_mzs = ", ".join(collected_mzs) if collected_mzs else "No Fragments"
            flattened_records.append({
                "Name": active_name,
                "Formula": active_formula,
                "MS2 Fragment Ions": comma_separated_mzs
            })
            
        print(f"\n  Finished sheet '{sheet_name}'. Found {compounds_counter} compounds.")

    # Output to dataframe matrix
    df_output = pd.DataFrame(flattened_records)
    return df_output

# --- EXECUTE ENGINE ---
try:
    final_flat_df = extract_nested_ms2_with_tracking(INPUT_FILE)
    
    if final_flat_df.empty:
        print("\n Extraction complete, but no records matched. Ensure your main compound rows contain data.")
    else:
        print(f"\nSaving data matrix to '{OUTPUT_FILE}'...")
        final_flat_df.to_excel(OUTPUT_FILE, index=False)
        print(f" Success! Flat records written to '{OUTPUT_FILE}'")
        print(f"Total Unique Compounds Processed: {len(final_flat_df)}")
        
except Exception as e:
    print(f"\n Script failed due to error: {str(e)}")