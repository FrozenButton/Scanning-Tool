import pandas as pd
from pathlib import Path
from typing import List, Dict, Any

def find_signature_matches(signature_value: int, csv_path: str) -> List[Dict[str, Any]]:
    """
    Find all rows in the scan_signatures.csv file that match the given value.
    Returns a list of dicts for each match.
    """
    try:
        df = pd.read_csv(csv_path)
        matches = df[df['value'] == signature_value]
        return matches.to_dict(orient='records')
    except Exception:
        return []

if __name__ == "__main__":
    # Example usage
    csv_file = Path(__file__).parent.parent.parent / "csv" / "scansig" / "scan_signatures.csv"
    value = int(input("Enter signature value: "))
    results = find_signature_matches(value, str(csv_file))
    if results:
        print(f"Matches for value {value}:")
        for r in results:
            print(r)
    else:
        print(f"No matches found for value {value}.")
