import pandas as pd

# --- CONFIGURATION ---
COMPOUNDS_FILE = "Compounds_Perfect_Structured_Output.xlsx"  # The file we just generated
VEGETABLE_MAPPING_FILE = "Study_Information_Filtered.xlsx"       # File containing the mapping of Study File IDs to Vegetable names
OUTPUT_FILE = "Final_Matched_Vegetables_Report.xlsx"

print("Phase 1: Loading data files...")
df_compounds = pd.read_excel(COMPOUNDS_FILE)
df_mapping = pd.read_excel(VEGETABLE_MAPPING_FILE)

# Clean up column headers to ensure exact matching behaves properly
df_compounds.columns = df_compounds.columns.astype(str).str.strip()
df_mapping.columns = df_mapping.columns.astype(str).str.strip()

# Dynamically locate the Study File ID column in the mapping file (e.g., 'Study File ID' or 'Sample')
mapping_study_col = next((c for c in df_mapping.columns if "study file" in c.lower() or "file id" in c.lower()), None)
mapping_veg_col = next((c for c in df_mapping.columns if "vegetable" in c.lower() or "identifier" in c.lower()), None)

if not mapping_study_col or not mapping_veg_col:
    print(f"[CRITICAL ERROR] Mapping file columns not recognized. Found: {list(df_mapping.columns)}")
    exit()

print("Phase 2: Standardizing IDs and merging datasets...")

# Create temporary clean tracking keys to ensure flawless matching regardless of casing or accidental spaces
df_compounds['study_key_temp'] = df_compounds['Study File ID'].astype(str).str.strip().str.upper()
df_mapping['study_key_temp'] = df_mapping[mapping_study_col].astype(str).str.strip().str.upper()

# Merge the datasets using a left join to keep all compounds even if a mapping is missing
merged_df = pd.merge(df_compounds, df_mapping[['study_key_temp', mapping_veg_col]], on='study_key_temp', how='left')

print("Phase 3: Restructuring headers and clean indexing...")

# Rename the matched column to exactly "Vegetable"
merged_df = merged_df.rename(columns={mapping_veg_col: "Vegetable"})

# Recalculate s/n sequentially from 1 to ensure it remains a solid clean sequence
merged_df['s/n'] = range(1, len(merged_df) + 1)

# Organize columns into your exact required sequence
target_sequence = [
    "s/n",
    "Name",
    "Formula",
    "m/z",
    "Reference Ion",
    "Annot. DeltaMass [ppm]",
    "Calc. MW",
    "RT [min]",
    "Study File ID",
    "Vegetable"
]

# Ensure everything matches your layout specification safely
final_df = merged_df[target_sequence].copy()

# Fill any unmapped vegetable entries with "Unknown" just in case
final_df['Vegetable'] = final_df['Vegetable'].fillna("Unknown")

print("Phase 4: Exporting report file...")
final_df.to_excel(OUTPUT_FILE, index=False)
print(f" Success! Final unified report generated and saved to: '{OUTPUT_FILE}'")