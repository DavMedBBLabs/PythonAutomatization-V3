# Excel to XRay JSON Converter

This console application processes Excel files and converts them into the JSON format expected by XRay.

## Usage

Set the following environment variables before running the script:

- `CLIENT_ID`
- `CLIENT_SECRET`
- `AUTH_URL`
- `XRAY_IMPORT_URL` *(not used yet)*
- `PATH_EXCEL` – directory containing the Excel files.
- `PATH_JSON` – directory where the JSON files will be generated.

Install dependencies:

```bash
pip install -r requirements.txt
```

Run the program:

```bash
python excel_to_xray_json.py
```

Follow the menu prompts to provide the project key, feature number and Excel file name.
