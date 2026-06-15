# Excel Sheet Renamer & PDF Exporter

A browser-based tool that splits every sheet in an Excel workbook into a separate, renamed PDF file and downloads them all as a single ZIP.

## Project structure

```
excel-renamer/
├── index.html       ← Open this in your browser
├── css/
│   └── style.css
└── js/
    └── app.js
```

## How to use

1. **Open `index.html`** in any modern browser (Chrome, Edge, Firefox).
2. **Upload your Excel file** (Box 1) — the one that contains all sheets plus a naming sheet.
3. **Configure naming** (Box 2):
   - Select the sheet that holds the output file names.
   - Select the column that contains those names.
   - Set which row is the header (names are read from the row after).
4. **Review the mapping** table — it shows you exactly which sheet → which PDF filename.
5. Click **Generate & Download ZIP** — a ZIP file is downloaded containing one PDF per sheet.

## Notes

- No installation required. No server. All processing happens inside the browser.
- Libraries used (loaded automatically via CDN — internet connection needed):
  - **SheetJS** — reads `.xlsx` / `.xls` files
  - **jsPDF + jspdf-autotable** — converts sheet data to PDF
  - **JSZip** — packages the PDFs into a downloadable ZIP
- For the best PDF rendering, open in **Google Chrome** or **Microsoft Edge**.

## Running in VS Code

You can use the **Live Server** extension for a smooth dev experience:
1. Install the "Live Server" extension in VS Code.
2. Right-click `index.html` → **Open with Live Server**.
3. The app opens automatically in your browser.

Alternatively, just double-click `index.html` — it works without a server.
