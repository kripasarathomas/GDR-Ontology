import os
import sys
import importlib
from pathlib import Path

# --- CONFIG ---
PROJECT_ROOT = Path(__file__).resolve().parent

CONFIG = {
    "paths": {
        "input_folder": PROJECT_ROOT / "finalinput",         # raw Excel input
        "gdr_file": PROJECT_ROOT / "GDR.xlsx",               # reference file
        "mapper_output": PROJECT_ROOT / "f1output",          # mapper results
        "ontology_output": PROJECT_ROOT / "lastcomparison_ttl"  # RDF output
    },
    "modules": {
        "loader": "finalLoader",       # loader.py
        "mapper": "finalMapper",       # mapper.py
        "builder": "finalOntology"     # ontology.py
    }
}

# Ensure project root is in sys.path
sys.path.insert(0, str(PROJECT_ROOT))


def run_module(module_name: str, config: dict):
    """Dynamically import and run the specified module with CONFIG."""
    try:
        module = importlib.import_module(module_name)

        if hasattr(module, "main"):
            print(f"\n Running {module_name}.main() ...")
            module.main(config)  # pass config dictionary
        else:
            print(f"No main(config) function found in {module_name}")
    except Exception as e:
        print(f"Error while running {module_name}: {e}")


def main():
    print("\n Starting GDR Data Processing Pipeline")

    # Step 1: Loader
    if CONFIG["modules"].get("loader"):
        run_module(CONFIG["modules"]["loader"], CONFIG)

    # Step 2: Mapper
    if CONFIG["modules"].get("mapper"):
        run_module(CONFIG["modules"]["mapper"], CONFIG)

    # Step 3: Ontology Builder
    if CONFIG["modules"].get("builder"):
        run_module(CONFIG["modules"]["builder"], CONFIG)

    print("\n Pipeline execution completed!")


if __name__ == "__main__":
    main()
