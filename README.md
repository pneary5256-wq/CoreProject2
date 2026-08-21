# CoreProject2

Applicant Eligibility Analysis Tool for reviewing Aptem applicant exports. Upload a CSV or spreadsheet to see eligibility results, summary metrics, applicant details, and configurable rule checks.

## Run

From this folder, start a local server:

```bash
python3 -m http.server 8000
```

Open `http://localhost:8000` in your browser, then upload the test data set from the `evidence pack`. Supported formats are `.csv`, `.tsv`, `.xls`, `.xlsx`, `.xlsm`, and `.ods`.
