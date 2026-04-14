import csv
from pathlib import Path
from typing import List, Dict, Any

def find_signature_matches(signature_value: int, csv_path: str) -> List[Dict[str, Any]]:
    """
    Find all rows in the scan_signatures.csv file that match the given value.
    Returns a list of dicts for each match.
    """
    matches = []
    with open(csv_path, newline='', encoding='utf-8') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            try:
                if int(row['value']) == signature_value:
                    matches.append(row)
            except (ValueError, KeyError):
                continue
    return matches

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
