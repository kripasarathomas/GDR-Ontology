import pandas as pd
from rapidfuzz import process, fuzz
from pathlib import Path

# === CONFIG ===
input_folder = Path("finalinput")
output_folder = Path("f1output")
output_folder.mkdir(exist_ok=True)

gdr_file ="GDR.xlsx"   # Reference/master file
FUZZY_THRESHOLD = 85                   # Adjust threshold (0–100)

# === HELPERS ===
def make_unique(cols):
    """Ensure unique column names after Excel import."""
    seen = {}
    result = []
    for col in cols:
        if col not in seen:
            seen[col] = 0
            result.append(col)
        else:
            seen[col] += 1
            result.append(f"{col}_{seen[col]}")
    return result


#mapper function call
   
def main(config=None):
    # --- Get paths from pipeline config ---
    input_folder = Path(config["paths"]["input_folder"]) if config else Path("finalinput")
    output_folder = Path(config["paths"]["mapper_output"]) if config else Path("f1output")
    gdr_file = Path(config["paths"]["gdr_file"]) if config else Path("GDR.xlsx")

    output_folder.mkdir(exist_ok=True)   

# === LOAD GDR FILE (all sheets combined) ===
gdr_sheets = pd.read_excel(gdr_file, sheet_name=None)

gdr_titles = []
for df in gdr_sheets.values():
    df.columns = make_unique(df.columns)
    if "Berufs" in df.columns:
        gdr_titles.extend(df["Berufs"].dropna().astype(str).str.strip().tolist())
gdr_titles = list(set(gdr_titles))  # unique list

# === PROCESS ALL FILES IN INPUT FOLDER ===
for hist_file in input_folder.glob("*.xlsx"):
    if hist_file.name == "GDR.xlsx":
        continue  # skip GDR file itself

    print(f" Processing {hist_file.name} ...")

    # Load all sheets of the historical file
    hist_sheets = pd.read_excel(hist_file, sheet_name=None)

    # Define output file path
    output_file = output_folder / f"comparison_{hist_file.stem}_vs_GDR.xlsx"
    writer = pd.ExcelWriter(output_file, engine="openpyxl")

    # Process each sheet in the historical file
    for sheet_name, df_hist in hist_sheets.items():
        df_hist.columns = make_unique(df_hist.columns)
        if "Berufs" not in df_hist.columns:
            continue

        matches = []
        unmatched = []

        for idx, row in df_hist.iterrows():
            title = str(row["Berufs"]).strip()
            if not title or title.lower() == "nan":
                continue

            match, score, _ = process.extractOne(title, gdr_titles, scorer=fuzz.token_sort_ratio)

            if score >= FUZZY_THRESHOLD:
                row_data = row.to_dict()
                row_data["GDR_Match"] = match
                matches.append(row_data)
            else:
                unmatched.append(row.to_dict())

        # Save results to Excel
        if matches:
            pd.DataFrame(matches).to_excel(writer, sheet_name=f"{sheet_name}_matches", index=False)
        if unmatched:
            pd.DataFrame(unmatched).to_excel(writer, sheet_name=f"{sheet_name}_unmatched", index=False)

    writer.close()
    print(f" Results saved to {output_file}")
