import pandas as pd
from rdflib import Graph, Literal, URIRef, Namespace
from pathlib import Path
from rdflib.namespace import RDF, RDFS, SKOS
import re
import os

# --- CONFIG ---
INPUT_FOLDER = 'f1output'
OUTPUT_FOLDER = 'lastcomparison_ttl'

GLMO = Namespace("http://example.com/glmo/")
ISCO = Namespace("http://data.europa.eu/esco/isco/")
KLDB = Namespace("http://purl.org/lob/kldb/")

def clean_for_uri(text):
    """Cleans a string for safe use in RDF URIs."""
    text = str(text).strip()
    return re.sub(r'[^a-zA-Z0-9_]', '_', text)

def generate_ttl_from_excel_folder(input_folder, output_folder):
    if not os.path.exists(input_folder):
        print(f" ERROR: Input folder '{input_folder}' not found.")
        return
    os.makedirs(output_folder, exist_ok=True)

    for filename in os.listdir(input_folder):
        if not filename.endswith(".xlsx"):
            continue

        excel_file_path = os.path.join(input_folder, filename)
        output_ttl_path = os.path.join(output_folder, f"{os.path.splitext(filename)[0]}.ttl")

        g = Graph()
        g.bind("glmo", GLMO)
        g.bind("skos", SKOS)
        g.bind("rdfs", RDFS)

        try:
            xls = pd.ExcelFile(excel_file_path)
            print(f"\nProcessing: {filename}")
        except Exception as e:
            print(f" ERROR reading {filename}: {e}")
            continue

        for sheet_name in xls.sheet_names:
            df = xls.parse(sheet_name).astype(str)
            print(f"  - Sheet '{sheet_name}', {len(df)} rows")

            for index, row in df.iterrows():
                # Use Berufsnummer if available, else Berufs
                identifier = row.get("berufs_nummer") if pd.notna(row.get("berufs_nummer")) else row.get("Berufs")
                if pd.isna(identifier):
                    identifier = f"row_{index}"

                # Combine sheet name + identifier for URI
                row_uri = URIRef(GLMO[f"{clean_for_uri(sheet_name)}_{clean_for_uri(identifier)}"])

                # Explicit type
                g.add((row_uri, RDF.type, SKOS.Concept))

                # rdfs:label → Berufs
                if "Berufs" in df.columns and pd.notna(row["Berufs"]):
                    g.add((row_uri, RDFS.label, Literal(row["Berufs"])))

                # skos:notation → Berufsnummer
                if "berufs_nummer" in df.columns and pd.notna(row["berufs_nummer"]):
                    g.add((row_uri, SKOS.notation, Literal(row["berufs_nummer"])))

                # All other columns → glmo:<column>
                for col_name, value in row.items():
                    if pd.notna(value) and str(value).strip() != '':
                        if col_name in ["Berufs", "berufs_nummer"]:
                            continue
                        prop_uri = GLMO[clean_for_uri(col_name)]
                        g.add((row_uri, prop_uri, Literal(value)))

        g.serialize(destination=output_ttl_path, format="turtle")
        print(f" Saved {len(g)} triples to {output_ttl_path}")

#ontology function call
def main(config=None):
    input_folder = Path(config["paths"]["mapper_output"]) if config else Path("f1output")
    output_folder = Path(config["paths"]["ontology_output"]) if config else Path("lastcomparison_ttl")
    generate_ttl_from_excel_folder(input_folder, output_folder)

if __name__ == "__main__":
    generate_ttl_from_excel_folder(INPUT_FOLDER, OUTPUT_FOLDER)
