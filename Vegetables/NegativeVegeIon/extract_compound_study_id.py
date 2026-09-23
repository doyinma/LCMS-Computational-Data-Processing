import pandas as pd
import numpy as np
import re

# --- CONFIGURATION ---
INPUT_FILE = "ExportedCompounds_RepIons_Filtered.xlsx"  # Change this to your actual file name
OUTPUT_FILE = "Compounds_Perfect_Structured_Output.xlsx"

print("Phase 1: Reading hierarchical excel matrix...")
# Load the raw spreadsheet without headers to treat it as a pure coordinate grid
df = pd.read_excel(INPUT_FILE, header=None)

print("Phase 2: Extracting Master Compounds and scanning downwards for the Study File ID...")

final_records = []
s_n_counter = 1

# Find where the column name definitions live
header_row_idx = None
for idx, row in df.iterrows():
    row_str = [str(cell).strip().lower() for cell in row]
    if "formula" in row_str or "calc. mw" in row_str:
        header_row_idx = idx
        break

if header_row_idx is None:
    print("[ERROR] Could not automatically locate the header row layout. Defaulting to 0.")
    header_row_idx = 0

# Extract headers for mapping lookups
headers = [str(cell).strip().lower() for cell in df.iloc[header_row_idx]]

def get_col_index(possible_names, default_idx):
    for name in possible_names:
        if name in headers:
            return headers.index(name)
    return default_idx

name_col_idx = get_col_index(["name"], 0)
formula_col_idx = get_col_index(["formula"], 1)
mz_col_idx = get_col_index(["m/z", "mz"], 7)
ref_ion_col_idx = get_col_index(["reference ion", "ref ion"], 5)
delta_col_idx = get_col_index(["annot. deltamass [ppm]", "annot. deltamass (ppm)", "deltamass"], 2)
mw_col_idx = get_col_index(["calc. mw", "calc mw"], 3)
rt_col_idx = get_col_index(["rt [min]", "rt (min)", "rt"], 6)

# Column I is position 8 (0-indexed)
STUDY_FILE_COL_IDX = 8 

total_rows = len(df)
idx = header_row_idx + 1

while idx < total_rows:
    row = df.iloc[idx]
    formula_val = str(row.iloc[formula_col_idx]).strip()
    
    # Check if this row is a main compound row
    if pd.notna(row.iloc[formula_col_idx]) and formula_val.upper().startswith("C") and formula_val.lower() != "nan":
        
        # Pull master compound attributes safely
        comp_name = row.iloc[name_col_idx] if pd.notna(row.iloc[name_col_idx]) else ""
        comp_formula = formula_val
        comp_mz = row.iloc[mz_col_idx] if pd.notna(row.iloc[mz_col_idx]) else ""
        comp_ref = row.iloc[ref_ion_col_idx] if pd.notna(row.iloc[ref_ion_col_idx]) else ""
        comp_delta = row.iloc[delta_col_idx] if pd.notna(row.iloc[delta_col_idx]) else ""
        comp_mw = row.iloc[mw_col_idx] if pd.notna(row.iloc[mw_col_idx]) else ""
        comp_rt = row.iloc[rt_col_idx] if pd.notna(row.iloc[rt_col_idx]) else ""
        
        target_study_id = ""
        
        # Scan downwards from the master row to look for the first real Study File ID value
        # It skips the immediate label sub-header text by evaluating with a regex match
        scan_idx = idx + 1
        while scan_idx < total_rows:
            next_row = df.iloc[scan_idx]
            
            # Stop searching if we accidentally hit the NEXT compound master row before finding an ID
            next_row_formula = str(next_row.iloc[formula_col_idx]).strip()
            if pd.notna(next_row.iloc[formula_col_idx]) and next_row_formula.upper().startswith("C") and next_row_formula.lower() != "nan":
                break
                
            file_id_val = str(next_row.iloc[STUDY_FILE_COL_IDX]).strip()
            
            # Regex match check: Looking specifically for a pattern starting with "F" followed by digits (e.g., F14, F25)
            if pd.notna(next_row.iloc[STUDY_FILE_COL_IDX]) and re.match(r'^[fF]\d+', file_id_val):
                target_study_id = file_id_val
                break  # Target found! Stop looking down.
                
            scan_idx += 1
            
        # Append the successfully collected row data points
        final_records.append({
            "s/n": s_n_counter,
            "Name": comp_name,
            "Formula": comp_formula,
            "m/z": comp_mz,
            "Reference Ion": comp_ref,
            "Annot. DeltaMass [ppm]": comp_delta,
            "Calc. MW": comp_mw,
            "RT [min]": comp_rt,
            "Study File ID": target_study_id
        })
        
        s_n_counter += 1
        # Advance the index pointer forward to where we finished scanning this compound's block
        idx = scan_idx if scan_idx > idx else idx + 1
    else:
        idx += 1

# -------------------------------------------------------------------------
# PHASE 3: Compile Report File
# -------------------------------------------------------------------------
print("\nPhase 3: Finalizing reporting worksheet...")
if not final_records:
    print("[ERROR] Zero matches extracted. Please double check the input data values.")
else:
    target_headings = ["s/n", "Name", "Formula", "m/z", "Reference Ion", "Annot. DeltaMass [ppm]", "Calc. MW", "RT [min]", "Study File ID"]
    final_df = pd.DataFrame(final_records, columns=target_headings)
    final_df.to_excel(OUTPUT_FILE, index=False)
    print(f" Success! The hierarchical data has been correctly resolved. Saved report directly to: '{OUTPUT_FILE}'")