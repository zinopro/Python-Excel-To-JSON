# Excel to JSON Converter
This Python script reads data from an Excel file, processes it, and converts it into JSON format.

## Features
- **Determine Sheet Name**: Automatically detects and uses the first sheet name from the Excel file.
- **Identify Tables**: Identifies tables within the sheet, even if there are multiple tables.
- **Clean Data**: Cleans the data by removing empty rows and ensuring column names are properly formatted.
- **Convert Date to Words**: Converts date strings to a more readable format.
- **Evaluate Formulas**: Evaluates formulas within the sheet and retrieves their computed values.
- **Serialize to JSON**: Serializes the tables into a JSON file.

## Requirements
- Python 3.6+
- `openpyxl` library

## Installation
1. Clone the repository or download the script file.

2. Navigate to the project directory.

3. Create a virtual environment (optional but recommended):

    ```bash
    python3 -m venv .venv
    source .venv/bin/activate
    ```

4. Install the required dependencies:

    ```bash
    pip install openpyxl
    ```

## Usage
1. Place your Excel file (e.g., `example_0.xlsx`) in the project directory.

2. Modify the script if needed to change the input file name, sheet name, and output file name:

    ```python
    if __name__ == "__main__":
        input_excel_file = Path("example_0.xlsx")
        sheet_name = 'Analysis Output'
        output_json_file = 'output.json'
        
        main(input_excel_file, sheet_name, output_json_file)
    ```

3. Run the script:

    ```bash
    python3 main_script.py
    ```

4. The 'output.json' file will be created in the project directory.