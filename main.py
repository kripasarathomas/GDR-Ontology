import pandas as pd
from rdflib import Graph, URIRef, Literal, Namespace
from rdflib.namespace import RDF, RDFS, SKOS, OWL
import hashlib
from pathlib import Path

# Configuration: Define input/output paths
# Place your source Excel files (e.g., 1971.xlsx, 1985.xlsx) in this folder
INPUT_DIR = Path("C://Users//kripa//Research Project//inputfolder")
# The generated .ttl files will be saved here
OUTPUT_DIR = Path("generated_ttl_files")
# The central mapping file that is compared against
MAPPING_FILE = Path("GDR.xlsx")

# Create directories if they don't exist
INPUT_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

# Global Namespaces (used in all files)
MODERN_NS = Namespace("http://www.bibb.de/modern#")
DBO_NS = Namespace("http://dbpedia.org/ontology/")

# Helper Functions
def sha_uri(base, text):
    return URIRef(base + hashlib.sha1(text.encode("utf-8")).hexdigest())

def add_occupation(graph, uri, label, source_literal):
    graph.add((uri, RDF.type, OWL.Class))
    graph.add((uri, SKOS.prefLabel, Literal(label, lang="de")))
    graph.add((uri, DBO_NS.rightsHolder, source_literal))

def add_mapping(graph, gdr_uri, modern_label):
    modern_uri = sha_uri(str(MODERN_NS), modern_label)
    # Add modern occupation details (only if not already present)
    if (modern_uri, RDF.type, OWL.Class) not in graph:
        graph.add((modern_uri, RDF.type, OWL.Class))
        graph.add((modern_uri, SKOS.prefLabel, Literal(modern_label, lang="de")))
    # Create the link
    graph.add((gdr_uri, SKOS.related, modern_uri))

# Main Processing Logic

# Load the central mapping data ONCE
try:
    df_gdr = pd.read_excel(MAPPING_FILE, sheet_name=0)
    mapping_dict = dict(zip(df_gdr["Berufs"], df_gdr["Zuordnungsmöglichkeit zu anerkannten Ausbildungsberufen"]))
    print(f"Successfully loaded mapping data from '{MAPPING_FILE}'")
except FileNotFoundError:
    print(f"Error: Mapping file not found at '{MAPPING_FILE}'. Please create it and re-run.")
    exit()

# Loop through each Excel file in the input directory
xlsx_files = list(INPUT_DIR.glob("*.xlsx"))
if not xlsx_files:
    print(f"No .xlsx files found in '{INPUT_DIR}'. Please add files to process.")

for file_path in xlsx_files:
    print(f"\n--- Processing file: {file_path.name} ---")

    # Initialize a new, clean graph for EACH file
    g = Graph()
    
    # Create and bind namespaces for this specific graph
    year = file_path.stem  # Gets filename without extension, e.g., "1971"
    gdr_ns = Namespace(f"http://www.ddr-berufe.de/{year}#")
    
    g.bind(f"gdr{year}", gdr_ns)
    g.bind("modern", MODERN_NS)
    g.bind("dbo", DBO_NS)
    g.bind("skos", SKOS)
    g.bind("owl", OWL)
    
    # Process the current Excel file
    xls_file = pd.ExcelFile(file_path)
    seen_labels = set()
    source_literal = Literal(f"GDR_{year}")

    for sheet_name in xls_file.sheet_names:
        df = xls_file.parse(sheet_name)
        print(f"Scanning sheet: {sheet_name}")

        for _, row in df.iterrows():
            base_label = str(row.get("Berufsbezeichnung", "")).strip()
            specialization = str(row.get("Bezeichnung der Spezialisierungsrichtungen", "")).strip()

            if not base_label or base_label.lower() == "nan":
                continue

            full_label = f"{base_label} – {specialization}" if specialization and specialization.lower() != 'nan' else base_label
            
            if full_label in seen_labels:
                continue
            
            seen_labels.add(full_label)
            uri = sha_uri(str(gdr_ns), full_label)
            add_occupation(g, uri, full_label, source_literal)

            # Check for a mapping using the base label
            if base_label in mapping_dict and pd.notna(mapping_dict[base_label]):
                modern_label_raw = str(mapping_dict[base_label])
                # Take the first equivalent occupation if multiple are listed
                modern_label = modern_label_raw.split("/")[0].strip()
                add_mapping(g, uri, modern_label)

    # Save the graph for the current file
    output_path = OUTPUT_DIR / f"{year}.ttl"
    g.serialize(destination=str(output_path), format="turtle")
    print(f"Ontology saved to '{output_path}' with {len(g)} triples.")

print("\n All files processed successfully.")
