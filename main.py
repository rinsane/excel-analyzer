import os
import json
import pandas as pd

def load_first_excel():
    sheets_dir = "sheets"
    for file in os.listdir(sheets_dir):
        if file.endswith(".xlsx"):
            filepath = os.path.join(sheets_dir, file)
            return pd.read_excel(filepath, header=[0, 1, 2])
    raise FileNotFoundError("No .xlsx file found in sheets/")

def print_metadata_section(first_row, df):
    """
    Print the first 5 metadata columns (url, metadata, QA feedback/notes/person).
    """
    url = first_row[[c for c in df.columns if c[-1] == "url"][0]]
    metadata_raw = first_row[[c for c in df.columns if c[-1] == "metadata"][0]]
    dq_feedback = first_row[[c for c in df.columns if c[-1] == "Data QA feedback"][0]]
    dq_notes = first_row[[c for c in df.columns if c[-1] == "Data QA notes"][0]]
    dq_person = first_row[[c for c in df.columns if c[-1] == "Data QA person"][0]]

    print("# Metadata\n")
    print(f"url -> {url}\n")

    # Parse metadata JSON if possible
    print("metadata ->")
    try:
        metadata_parsed = json.loads(metadata_raw)
        print(json.dumps(metadata_parsed, indent=4))
    except Exception:
        print(metadata_raw)
    print()

    print(f"Data QA feedback -> {dq_feedback}")
    print(f"Data QA notes -> {dq_notes}")
    print(f"Data QA person -> {dq_person}\n")

def print_grouped_by_level1(first_row, df):
    """
    Print the first row grouped as:
    # level_1
      ## level_0
         ### level_2 -> value
    """
    # Group columns by level_1 (skipping metadata cols)
    grouped = {}
    for col in df.columns[5:]:  # skip first 5
        lvl0, lvl1, lvl2 = col
        if lvl1 not in grouped:
            grouped[lvl1] = {}
        if lvl0 not in grouped[lvl1]:
            grouped[lvl1][lvl0] = []
        grouped[lvl1][lvl0].append((lvl2, first_row[col]))

    # Pretty print
    for lvl1, trainers in grouped.items():
        print(f"# {lvl1}\n")
        for lvl0, fields in trainers.items():
            print(f"## {lvl0}")
            for lvl2, value in fields:
                print(f"### {lvl2} -> {value}")
            print()
        print()

def main():
    df = load_first_excel()
    first_row = df.iloc[0]

    print_metadata_section(first_row, df)
    print_grouped_by_level1(first_row, df)

if __name__ == "__main__":
    main()
