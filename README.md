# Excel Review Analyzer

This script reads the first `.xlsx` file from the `sheets/` directory and prints structured review data.  
It separates metadata (url, metadata JSON, Data QA fields) and grouped trainer assessments for easy comparison.  
Headings are colorized in the terminal for readability.

## Setup

1. Clone the repository and change directory:

   ```bash
   git clone <your-repo-url>
   cd <your-repo-name>
````

2. Create and activate a virtual environment:

   **Linux / macOS**

   ```bash
   python3 -m venv venv
   source venv/bin/activate
   ```

   **Windows (PowerShell)**

   ```powershell
   python -m venv venv
   venv\Scripts\activate
   ```

3. Install dependencies:

   ```bash
   pip install -r requirements.txt
   ```

## Usage

1. Place your Excel file (`.xlsx`) inside the `sheets/` directory.
   The script will automatically pick the first `.xlsx` file it finds.

2. Run the script:

   ```bash
   python main.py
   ```

3. The script will print:

   * Metadata section with url, metadata (formatted as JSON), Data QA feedback, notes, and person.
   * Grouped trainer assessments organized by level 1 -> level 0 -> level 2.

## Notes

* The script assumes the Excel file has a 3-row header structure.
* Only the first row of data is printed by default. You can change the `row_number` variable in `main.py` to print other rows.

