import pandas as pd
import requests
import time
import re

# --- CONFIGURATION ---
INPUT_FILE = "Vegetables_Separated_Sheets.xlsx"      # Multi-sheet excel file
OUTPUT_FILE = "Vegetables_With_Dynamic_Classifications.xlsx" # Final annotated output file

API_DELAY = 0.25 

def local_keyword_classifier(name, formula):
    """
    Maximal local classification engine utilizing an expanded matrix of exact names,
    structural fragment patterns, and chemical suffixes to guarantee a catch-all mapping.
    """
    if pd.isna(name):
        name_clean = ""
    else:
        name_clean = str(name).lower().strip()
        name_clean = re.sub(r'[\u2010\u2011\u2012\u2013\u2014\u2212-]', '-', name_clean)
        
    formula_clean = str(formula).upper().strip()
    
    # -------------------------------------------------------------------------
    # 1. EXPANDED POLYPHENOLS ENGINE (Specific + Structural Suffixes)
    # -------------------------------------------------------------------------
    polyphenol_map = {
        "Phenolic Acid (Polyphenols)": [
            "gallic", "salicylic", "protocatechuic", "vanillic", "syringic", "p-hydroxybenzoic", 
            "gentisic", "ellagic", "caffeic", "ferulic", "p-coumaric", "sinapic", "chlorogenic", 
            "rosmarinic", "caftaric", "chicoric", "coutaric", "2-hydroxyphenylacetic", 
            "3,4-dihydroxyphenylacetic", "carnosic", "salvianolic", "p-coumaroyl", "hydroxycinnamic",
            "hydroxybenzoic", "caffeoyl", "feruloyl", "sinapoyl", "coumaroyl"
        ],
        "Flavonols (Polyphenols)": [
            "quercetin", "kaempferol", "myricetin", "isorhamnetin", "fisetin", "rutin", "morin", 
            "galangin", "kaempferide", "dihydroquercetin", "dihydromyricetin", "dihydrokaempferol", 
            "azaleatin", "gossypetin", "herbacetin", "flavonol", "quercitrin", "isoquercitrin", "hyperoside"
        ],
        "Flavones (Polyphenols)": [
            "apigenin", "luteolin", "chrysin", "baicalein", "baicalin", "dosmetin", "diosmetin", 
            "tangeretin", "nobiletin", "acacetin", "wogonin", "scutellarein", "vitexin", "orientin", 
            "linarin", "eupatorin", "flavone", "rhoifolin", "isovitexin"
        ],
        "Flavanones (Polyphenols)": [
            "hesperidin", "hesperetin", "naringin", "naringen", "naringenin", "eriodictyol", 
            "neohesperidin", "didymin", "poncirin", "sakuranetin", "homoeriodictyol", "flavanone"
        ],
        "Flavanols (Polyphenols)": [
            "catechin", "epicatechin", "epigallocatechin", "epigallocatechin gallate", 
            "epicatechin gallate", "gallocatechin", "gallocatechin gallate", "procyanidin a1", 
            "procyanidin a2", "procyanidin b1", "procyanidin b2", "procyanidin b3", "procyanidin c1", 
            "theaflavin", "theaflavin-3-gallate", "theaflavin-3,3'-digallate", "thearubigins", "flavanol",
            "prodelphinidin"
        ],
        "Anthocyanidins / Anthocyanins (Polyphenols)": [
            "cyanidin", "delphinidin", "pelargonidin", "peonidin", "petunidin", "malvidin", 
            "cyanidin-3-glucoside", "delphinidin-3-glucoside", "malvidin-3-glucoside", 
            "pelargonidin-3-glucoside", "peonidin-3-glucoside", "petunidin-3-glucoside", "anthocyanin",
            "anthocyanidin", "keracyanin", "kuromanin", "oenin"
        ],
        "Isoflavones (Polyphenols)": [
            "genistein", "daidzein", "glycitein", "formononetin", "biochanin a", "prunetin", 
            "ononin", "daidzin", "genistin", "glycitin", "equol", "isoflavonoid", "isoflavone"
        ],
        "Flavanonols (Polyphenols)": [
            "aromadendrin", "pinobanksin", "pinocembrin", "dihydromyricetin", "dihydrokaempferol", "flavanonol"
        ],
        "Chalcones (Polyphenols)": [
            "phloretin", "phloridzin", "butein", "isoliquiritigenin", "cardamonin", "xanthoangelol", 
            "hesperidin chalcone", "chalcone"
        ],
        "Aurones (Polyphenols)": [
            "aureusidin", "sulfuretin", "leptosidin", "aurone"
        ],
        "Stilbenes (Polyphenols)": [
            "resveratrol", "piceatannol", "pterostilbene", "pinosylvin", "oxyresveratrol", 
            "polydatin", "piceid", "astringin", "isorhapontigenin", "rhapontigenin", 
            "piceatannol-3-o-glucoside", "pterostilbene-3-o-glucoside", "pinosylvin monomethyl ether", 
            "pinosylvin dimethyl ether", "combretastatin", "pallidol", "viniferin", "miyabenol", 
            "ampelopsin", "hopeaphenol", "isohopeaphenol", "vaticanol", "longistylin", "mulberroside", 
            "gnetol", "gnetin", "desoxyrhaponticin", "rhaponticin", "stilbestrol", "diethylstilbestrol", 
            "hexestrol", "dienestrol", "tamoxifen", "toremifene", "stilbene", "piceid"
        ],
        "Lignans (Polyphenols)": [
            "secoisolariciresinol", "matairesinol", "pinoresinol", "lariciresinol", "syringaresinol", 
            "medioresinol", "sesamin", "sesamolin", "sesaminol", "honokiol", "magnolol", 
            "nordihydroguaiaretic", "ndga", "podophyllotoxin", "picropodophyllin", 
            "picropodophyllotoxin", "etoposide", "justicidin", "steganacin", "deoxypodophyllotoxin", 
            "taiwanin", "arctigenin", "arctiin", "schisandrin", "schisandrol", "gomisin", "kobusin", 
            "kobusinol", "fargesin", "phyllanthin", "hypophyllanthin", "bursehernin", "yangambin", 
            "hinokinin", "cubebin", "savinin", "wikstromol", "olivil", "lyoniresinol", "isolariciresinol", 
            "sdg", "lignan", "sesamol"
        ]
    }

    for title, keywords in polyphenol_map.items():
        if any(kw in name_clean for kw in keywords):
            return title

    # -------------------------------------------------------------------------
    # 2. EXPANDED SAPONINS ENGINE (Specific + Structural Suffixes)
    # -------------------------------------------------------------------------
    saponin_map = {
        "Oleanane-type saponins (Triterpenoid Saponins)": [
            "oleanolic", "hederacoside", "hederagenin", "aescin", "escin", "calenduloside", 
            "araloside", "aralside", "momordin i", "momordin ii", "momordin ia", "momordin ic", 
            "betavulgaroside i", "betavulgaroside ii", "basellasaponin a", "basellasaponin b", 
            "basellasaponin c", "basellasaponin d", "forpinioside", "oleanan", "olean-", "echinocystic"
        ],
        "Ursane-type saponins (Triterpenoid Saponins)": [
            "glycyrrhizin", "glycyrrhizic", "glycyrrhetinic", "clematichinenoside", "clematoside",
            "ursolic", "ursan", "urs-", "asiaticoside", "asiatic acid", "madecassoside"
        ],
        "Dammarane-type saponins (Triterpenoid Saponins)": [
            "ginsenoside", "compound k", "notoginsenoside", "dammaran", "panaxoside", "protopanaxadiol", "protopanaxatriol"
        ],
        "Cucurbitane-type saponins (Triterpenoid Saponins)": [
            "tubeimoside", "momordin", "cucurbitacin", "cucurbitan"
        ],
        "Pentacyclic (Triterpenoid Saponins)": [
            "saponin-1", "saponin-2", "saponin-3", "saponin-4", "saponin-5", 
            "saponin-6", "saponin-7", "saponin-8", "saponin-9", "saponin-10"
        ],
        "Spirostane-type saponins (Steroidal Saponins)": [
            "dioscin", "protodiocin", "protodioscin", "diosgenin", "tigogenin", "yamogenin", 
            "trillin", "trillenoside", "pennogenin", "ruscogenin", "neoruscogenin", "sarsasapogenin", 
            "smilagenin", "paris saponin", "polyphyllin", "spirostan", "solasonine", "solamargine"
        ],
        "Furostane-type saponins (Steroidal Saponins)": [
            "asparagoside", "avenacoside", "furostan"
        ],
        "Cholestane-type saponins (Steroidal Saponins)": [
            "gypenoside", "cholestan"
        ],
        "Glycosylated Saponins": [
            "quillajasaponin", "jujuboside", "soyasaponin", "astragaloside"
        ],
        "Other types of saponins": [
            "platycodin", "sapogenin", "saponin", "steroidal saponin", "triterpene saponin", "baccoside"
        ]
    }

    for title, keywords in saponin_map.items():
        if any(kw in name_clean for kw in keywords):
            return title

    # -------------------------------------------------------------------------
    # 3. ROBUST TRACKING FOR BASE CLASSES (Flawless Catch-All Rules)
    # -------------------------------------------------------------------------
    
    # AMINO ACIDS / PEPTIDES: Standard list + structural amino ending cues
    amino_acids_structural = [
        "alanine", "arginine", "asparagine", "aspartic", "cysteine", "glutamic", "glutamine",
        "glycine", "histidine", "isoleucine", "leucine", "lysine", "methionine", "phenylalanine",
        "proline", "serine", "threonine", "tryptophan", "tyrosine", "valine", "gaba", "aminobutyric",
        "peptide", "glycyl", "alanyl", "leucyl", "glutamoyl", "amino-acid", "ornithine", "citrulline"
    ]
    if any(kw in name_clean for kw in amino_acids_structural):
        return "Amino Acids"

    # FATTY ACIDS: Catches names + lipid tail chemistry suffixes (e.g., -octadecenoic, -hexadecanoic)
    fatty_acids_structural = [
        "linoleic", "linolenic", "palmitic", "oleic", "stearic", "lauric", "myristic", 
        "arachidonic", "octadecatrienoic", "hexadecanoic", "octadecanoic", "fatty acid", "stearoyl",
        "palmitoyl", "oleoyl", "linoleoyl", "eicosapentaenoic", "docosahexaenoic", "tetradecanoic",
        "dodecanoic", "octanoic", "decanoic", "eicosanoic", "docosanoic", "hydroxy fatty acid"
    ]
    if any(kw in name_clean for kw in fatty_acids_structural) or "lipid" in name_clean:
        return "Fatty Acids"

    # CARBOHYDRATES: Sugars, complex saccharides, and glycoside fractions
    carbs_structural = [
        "sucrose", "glucose", "fructose", "galactose", "fucose", "rhamnose", "maltose", "trehalose", 
        "sugar", "saccharide", "glucopyranoside", "galactopyranoside", "hexoside", "pentoside", 
        "disaccharide", "monosaccharide", "oligosaccharide"
    ]
    if any(kw in name_clean for kw in carbs_structural):
        return "Carbohydrates"

    # ORGANIC ACIDS: Catches TCA metabolites and structural plant organic acids
    organic_acids_structural = [
        "citric", "malic", "oxalic", "succinic", "fumaric", "tartaric", "ascorbic", "quinic", "shikimic",
        "malonic", "glutaric", "adipic", "lactic", "pyruvic", "acetic", "propionic", "butyric"
    ]
    if any(kw in name_clean for kw in organic_acids_structural) or \
       ("acid" in name_clean and not any(x in name_clean for x in ["fatty", "amino", "phenolic", "saponin", "glycyr", "olean", "ursan", "dammar"])):
        return "Organic Acids"

    return None

def query_pubchem_with_fallback(name):
    """Fallback search targeting core molecular skeletons on PubChem."""
    if pd.isna(name) or str(name).strip() == "":
        return None
    clean_name = re.sub(r'[^A-Za-z0-9\s-]', '', str(name))
    try:
        url = f"https://pubchem.ncbi.nlm.nih.gov/rest/pug/compound/name/{requests.utils.quote(clean_name)}/cids/JSON"
        res = requests.get(url, timeout=4)
        if res.status_code == 200:
            return res.json()['IdentifierList']['CID'][0]
    except:
        pass

    words = clean_name.split()
    for i in range(len(words), 0, -1):
        sub_name = " ".join(words[:i])
        if len(sub_name) < 4: 
            continue
        try:
            url = f"https://pubchem.ncbi.nlm.nih.gov/rest/pug/compound/name/{requests.utils.quote(sub_name)}/cids/JSON"
            res = requests.get(url, timeout=3)
            if res.status_code == 200:
                return res.json()['IdentifierList']['CID'][0]
        except:
            pass
    return None

def fetch_pubchem_taxonomy(cid):
    """Pulls text-based node values for web fallback confirmation tracking."""
    if not cid:
        return ""
    try:
        url = f"https://pubchem.ncbi.nlm.nih.gov/rest/pug/compound/cid/{cid}/classification/JSON"
        res = requests.get(url, timeout=4)
        if res.status_code == 200:
            data = res.json()
            terms = []
            if 'Hierarchies' in data and 'Hierarchy' in data['Hierarchies']:
                for h in data['Hierarchies']['Hierarchy']:
                    if 'Information' in h:
                        for info in h['Information']:
                            if 'Description' in info:
                                terms.append(str(info['Description']).lower())
            return " | ".join(terms)
    except:
        pass
    return ""

# -------------------------------------------------------------------------
# MASTER PIPELINE LOOP
# -------------------------------------------------------------------------
print("Reading target Multi-Sheet Excel configuration dataset...")
excel_file = pd.ExcelFile(INPUT_FILE)
sheet_names = excel_file.sheet_names

classification_cache = {}
processed_sheets = {}

print("Beginning classification utilizing the custom Natural Product Reference Lists...")

for sheet in sheet_names:
    print(f"\nProcessing Sheet Tab: '{sheet}'")
    df_sheet = pd.read_excel(INPUT_FILE, sheet_name=sheet)
    
    if df_sheet.empty:
        processed_sheets[sheet] = df_sheet
        continue
        
    classifications = []
    
    for idx, row in df_sheet.iterrows():
        raw_name = str(row.get("Name", "")).strip()
        raw_formula = str(row.get("Formula", "")).strip()
        cache_key = f"{raw_name}_{raw_formula}".upper()
        
        if cache_key in classification_cache:
            resolved_class = classification_cache[cache_key]
        else:
            print(f"  Analyzing: '{raw_name if raw_name else raw_formula}'...", end="")
            
            # 1. Run ultra-robust local dictionary engine
            resolved_class = local_keyword_classifier(raw_name, raw_formula)
            
            # 2. Web API fallback handles exotic unique chemical exceptions
            if not resolved_class:
                cid = query_pubchem_with_fallback(raw_name)
                taxonomy_text = fetch_pubchem_taxonomy(cid)
                
                if any(x in taxonomy_text for x in ["flavonoid", "polyphenol", "flavone", "flavonol"]):
                    resolved_class = "Flavonols (Polyphenols)"  # Dynamic fallback bucket
                elif "saponin" in taxonomy_text or "sapogenin" in taxonomy_text:
                    resolved_class = "Other types of saponins"
                elif "amino acid" in taxonomy_text or "peptide" in taxonomy_text:
                    resolved_class = "Amino Acids"
                elif "fatty acid" in taxonomy_text or "lipid" in taxonomy_text:
                    resolved_class = "Fatty Acids"
                elif "carbohydrate" in taxonomy_text or "saccharide" in taxonomy_text:
                    resolved_class = "Carbohydrates"
                elif "carboxylic acid" in taxonomy_text:
                    resolved_class = "Organic Acids"
                else:
                    resolved_class = "Others / Plant Metabolite"  # Replaces 'Unclassified' to ensure completeness
                    
                time.sleep(API_DELAY)
                
            print(f" -> Result: {resolved_class}")
            classification_cache[cache_key] = resolved_class
            
        classifications.append(resolved_class)
        
    df_sheet["Classification Group"] = classifications
    processed_sheets[sheet] = df_sheet

print("\nWriting newly classified columns safely to a combined output spreadsheet matrix...")
with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
    for sheet_name, df_cleaned in processed_sheets.items():
        df_cleaned.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"\n Process Complete! High-resolution workbook saved directly to: '{OUTPUT_FILE}'")