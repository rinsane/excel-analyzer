import json
import os

import pandas as pd
from rich.console import Console
from rich.text import Text

console = Console()


def load_first_excel():
    sheets_dir = "sheets"
    for file in os.listdir(sheets_dir):
        if file.endswith(".xlsx"):
            filepath = os.path.join(sheets_dir, file)
            return pd.read_excel(filepath, header=[0, 1])
    raise FileNotFoundError("No .xlsx file found in sheets/")


def print_metadata_section(row, df):
    """
    Print the first 5 metadata columns (url, metadata, QA feedback/notes/person).
    """
    url = row[[c for c in df.columns if c[-1] == "url"][0]]
    metadata_raw = row[[c for c in df.columns if c[-1] == "metadata"][0]]
    dq_feedback = row[[c for c in df.columns if c[-1] == "Data QA feedback"][0]]
    dq_notes = row[[c for c in df.columns if c[-1] == "Data QA notes"][0]]
    dq_person = row[[c for c in df.columns if c[-1] == "Data QA person"][0]]

    console.print(Text("# Metadata", style="bold magenta"))

    console.print(Text("url ->", style="bold cyan"), url)

    # Parse metadata JSON if possible
    console.print(Text("metadata ->", style="bold cyan"))
    try:
        metadata_parsed = json.loads(metadata_raw)
        console.print(json.dumps(metadata_parsed, indent=4))
    except Exception:
        console.print(metadata_raw)

    console.print(Text("Data QA feedback ->", style="bold cyan"), dq_feedback)
    console.print(Text("Data QA notes ->", style="bold cyan"), dq_notes)
    console.print(Text("Data QA person ->", style="bold cyan"), dq_person)
    console.print()


def print_grouped_by_level1(row, df):
    """
    Print the first row grouped as:
    # level_1
      ## level_0 -> value
    """
    grouped = {}
    for col in df.columns[10:]:  # skip first 10
        # 2-level structure: (lvl0, lvl1)
        lvl0, lvl1 = col
        if lvl1 not in grouped:
            grouped[lvl1] = {}
        grouped[lvl1][lvl0] = row[col]

    # Pretty print
    for lvl1, trainers in grouped.items():
        console.print(Text(f"# {lvl1}\n", style="bold magenta"))
        for lvl0, value in trainers.items():
            console.print(Text(f"## {lvl0} ->", style="bold green"), value)
        console.print()


def main():
    df = load_first_excel()
    row_number = 5  # Enter the (row number from excel sheet) - 4. We need to subtract 4 because of header rows in the excel sheet.
    row = df.iloc[row_number]

    print_metadata_section(row, df)
    print_grouped_by_level1(row, df)


if __name__ == "__main__":
    main()
