"""Point the Power BI project's DataFolder parameter at a folder holding the two CSVs.

Usage (close Power BI Desktop first):
    python scripts/set_data_folder.py                 # uses <project>/data/
    python scripts/set_data_folder.py data_private    # relative to the project folder
    python scripts/set_data_folder.py "D:\\vesta\\export"

The folder must contain vesta_Document.csv and vesta_DocumentVersion.csv.
Only VestaDashboard.SemanticModel/definition/expressions.tmdl is changed.
"""
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
TMDL = ROOT / "VestaDashboard.SemanticModel" / "definition" / "expressions.tmdl"

folder = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "data"
if not folder.is_absolute():
    folder = ROOT / folder
folder = folder.resolve()
missing = [f for f in ("vesta_Document.csv", "vesta_DocumentVersion.csv") if not (folder / f).exists()]
if missing:
    sys.exit(f"{folder} is missing: {', '.join(missing)}")

text = TMDL.read_text(encoding="utf-8")
value = str(folder).rstrip("\\") + "\\"
new = re.sub(r'(expression DataFolder = )"[^"]*"', lambda m: m.group(1) + '"' + value + '"', text, count=1)
TMDL.write_text(new, encoding="utf-8", newline="\n")
print(f"DataFolder -> {value}\nOpen VestaDashboard.pbip and click Refresh.")
