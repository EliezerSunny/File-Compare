# File Compare

A small local web app that compares two files and shows exactly what changed.
Everything runs on your own machine. Files are never uploaded anywhere.

Built with Python, Streamlit and pandas.

## Features

- **Spreadsheets** (`.xlsx`, `.xlsm`, `.xls`, `.csv`, `.tsv`)
  - Pick one or more **key columns** (e.g. `ID`, `Email`) to match rows, or leave empty to compare row by row
  - Shows rows added, rows removed and changed cells (old value → new value)
  - Compares each sheet you choose; reports columns that exist in only one file
  - Downloads a new `comparison_report.xlsx` (Summary, Changed cells, Added in B, Removed from B)
- **Documents** (`.docx`, `.pdf`, `.txt`, `.md`)
  - Side-by-side highlighted text diff with a similarity percentage
  - Downloads `differences.html` (opens in any browser, including phones) and `differences.txt`
  - Only text is compared, not fonts or layout. Scanned PDFs without real text are not supported
- Ignores harmless differences such as `100.0` vs `100`, extra spaces, and blank vs empty cells
- Handles fake Excel files, such as HTML tables or CSV text saved with an `.xls` or `.xlsx` extension (common with web portal exports)
- Keeps long ID numbers and leading zeros (e.g. phone numbers) intact when reading HTML-style files

## Requirements

- Python 3.10 or newer
- The libraries in `requirements.txt`

## Install

### Linux / macOS

Newer Ubuntu and Debian block system-wide `pip install`, so use a virtual environment:

```bash
git clone <your-repo-url>
cd filecompare
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

If `python3 -m venv` fails on Ubuntu/Debian: `sudo apt install python3-full python3-venv`

### Windows

```powershell
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

## Run

```bash
source .venv/bin/activate        # Windows: .venv\Scripts\activate
streamlit run app.py
```

Open http://localhost:8501 in your browser.

## Use it from your phone

Run the app on your computer so it is reachable on your local network:

```bash
streamlit run app.py --server.address 0.0.0.0
```

Find your computer's IP (`hostname -I` on Linux, `ipconfig` on Windows) and open
`http://YOUR-IP:8501` on your phone. Both devices must be on the same Wi-Fi.
If the phone cannot connect, allow the port in your firewall
(e.g. `sudo ufw allow 8501`).

Your files go from your phone to your own computer only.

## How to compare spreadsheets

1. Upload File A (original) and File B (new).
2. Choose the sheet in each file.
3. Choose the key column(s) that uniquely identify a row, such as an ID.
   If you leave it empty, row 1 is compared with row 1, row 2 with row 2, and so on.
4. Review the **Changed cells**, **Added in B** and **Removed from B** tabs.
5. Click **Download Excel report** to save the results as a new file.

If a key value appears more than once, those rows are matched in order of
appearance and labelled `#1`, `#2`. Add more key columns to avoid this.

## Project structure

```
app.py             Streamlit user interface
compare_core.py    Loading and comparison logic (no UI code)
requirements.txt   Python dependencies
```

## Troubleshooting

| Problem | Fix |
|---|---|
| `externally-managed-environment` when running pip | Use a virtual environment (see Install) |
| Excel opens the file with a "format and extension don't match" warning | The file is really HTML/CSV. The app can still read it |
| "This looks like an Excel file but it cannot be opened" | The file is password-protected or damaged. Remove the password or re-save as `.xlsx` |
| "File is empty (0 bytes)" | Re-download or re-upload the file |
| Old error message still shows after updating | Stop the app (`Ctrl+C`), delete `__pycache__`, run `pip install -r requirements.txt`, restart, then hard refresh the browser (`Ctrl+Shift+R`) |
| Installing on Android (Termux) fails | pandas/numpy are hard to build there. Run the app on a PC and open it from the phone instead |

## Limitations

- Word/PDF comparison is text only (no formatting, images or comments)
- Spreadsheet comparison looks at cell values, not formulas, colors or formatting
- Very large files are loaded fully into memory

## License

Add your preferred license here (e.g. MIT).
