# Excel to XRay JSON Converter

This console application processes Excel files and converts them into the JSON format expected by XRay.

## Usage

Set the following environment variables before running the script:

- `CLIENT_ID`
- `CLIENT_SECRET`
- `AUTH_URL`
- `XRAY_IMPORT_URL`
- `PATH_EXCEL` – directory containing the Excel files.
- `PATH_JSON` – directory where the JSON files will be generated.

Install dependencies (a `.env` file can be used to define the variables above):
```bash
python -m venv .venv
```

#### Ejecutar entorno virtual
```bash
.\.venv\Scripts\activate
```

Install dependencies:

```bash
pip install -r requirements.txt
```

Run the program:

```bash
python excel_to_xray_json.py
```

Follow the menu prompts to provide the project key, feature number and Excel file name.
After generating the JSON file, the application will offer the option to send it directly to XRay using the `XRAY_IMPORT_URL` endpoint.
