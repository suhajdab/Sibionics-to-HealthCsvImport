# Sibionics (XLSX) to Apple Health (CSV) Converter

This tool converts an Excel `.xlsx` exported from the SIBIONICS CGM app into one or more CSV files, with special handling for Continuous Glucose Monitoring (CGM) sensor data to produce per‑day, Health CSV Importer–compliant files. Finally Health CSV Importer can import the blood glucose data into Apple Health.

## Features

- Uses only the Python 3 standard library (no external dependencies).
- Unpacks `.xlsx` files and parses shared strings and worksheet XML directly.
- Reformats timestamps from `MM-DD-YYYY hh:mm AM/PM GMT+X` to ISO 8601 (`YYYY-MM-DD HH:MM:SS ±HH:MM`).
- Renames sensor reading columns to `Blood Glucose (mmol/L)` and outputs per‑day CSV files named `Sensor_Glucose_YYYY-MM-DD.csv`.
- Exports non‑sensor sheets as one CSV per sheet (sheet name sanitized).

## Requirements

- Python 3.6 or later

## Usage

```bash
python3 convert_xlsx_to_csv.py <input.xlsx> [output_dir]
```

- `<input.xlsx>`: Path to the Excel workbook.
- `[output_dir]`: Optional folder for CSV output (defaults to the current directory).

After running, you will see:

- `Sensor_Glucose_YYYY-MM-DD.csv` files for each date in the sensor data.
- Additional `<SheetName>.csv` files for any other sheets.

Each per‑day CSV begins with:

```csv
Timestamp,Blood Glucose (mmol/L)
2025-04-12 14:13:00 +01:00,5.6
...
```

## Example

```bash
$ python3 convert_xlsx_to_csv.py SiSensingCGM.xlsx output_folder/
Wrote output_folder/Sensor_Glucose_2025-04-12.csv
Wrote output_folder/Sensor_Glucose_2025-04-13.csv
...
```

Project created by OpenAI Codex in 4 prompts. For more information, visit [Github](https://https://github.com/openai/codex/).
