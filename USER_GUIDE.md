# DS XML Lineage Extractor — User Guide

A simple desktop tool to extract data lineage from SAP Data Services (DS) XML
export files. **No installation, no Python, no setup required.**

## Getting started (1 minute)

1. Double-click **`DS_Lineage_Extractor.exe`**.
   - The first launch may take 10–20 seconds while the app unpacks itself. This
     is normal and only happens slowly the first time.
   - Windows SmartScreen may show a "Windows protected your PC" warning the
     first time. Click **More info → Run anyway** (this appears for any new
     program that isn't code-signed).
2. **Input Path** — click **Browse...** and pick either:
   - a single `.xml` file exported from Data Services, or
   - a folder that contains several `.xml` files.
3. **Output Path** — where the result files will be saved. It defaults to a
   `DS_Lineage_Output` folder in your Documents; change it if you like.
4. Click **Run Extraction** and watch the progress bar.
5. When it finishes, use the **Export** buttons to save the view you want:
   - **📄 Column-Level** — each source column → target column mapping.
   - **📊 Table-Level** — table-to-table relationships only.
   - **📦 Full Lineage** — everything.

## What you get

Files are written to your Output Path:

| File | What it contains |
|------|------------------|
| `ds_column_lineage.csv` | Every column mapping, ordered Source → Target |
| `ds_table_lineage.csv` | Table-to-table relationships |
| `ds_lineage_summary.csv` | Counts and quality metrics |
| `ds_lineage_results.xlsx` | All of the above as Excel sheets |

Columns are always ordered for easy reading:
`Source_Object, Source_Column, Source_Column_Name, Target_Object, Target_Column, Target_Column_Name, …`

## Tips

- Open the `.csv` files in Excel, or the `.xlsx` directly.
- Large exports (90 MB+) can take several minutes — the progress bar keeps moving.
- If something goes wrong, use **Save Log** and send the log file to your support contact.

---

*For the team that maintains this tool: rebuild the `.exe` after code changes
with `./build_exe.ps1`. End users must use the newly built file — a running copy
does not pick up changes until it is restarted with the new build.*
