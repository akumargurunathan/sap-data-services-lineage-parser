# DOPPIA MACCHINA G2 — User Guide

A desktop tool for extracting end-to-end data lineage from SAP Business Objects Data Services (BODS) XML repository exports. No BODS server connection required.

---

## Getting Started

### Requirements

- Python 3.9 or higher
- Windows 10 / 11

### Install dependencies (first time only)

Open PowerShell and run:

```powershell
pip install lxml openpyxl pandas
```

### Launch the application

```powershell
python ds_ui_launcher.py
```

The application opens with two tabs: **Full Extraction** and **Targeted Lineage**.

---

## Preparing your XML file

Export your BODS repository from the Designer:

> **Tools → Repository Manager → Export → Full Repository**

Save the `.xml` file somewhere accessible on your machine (e.g. `C:\BODS_Exports\repository.xml`). Files over 100 MB are normal and fully supported.

---

## Tab 1 — Full Extraction

Use this tab when you want to extract lineage for **all tables** in the repository at once.

1. Click **Browse** next to *Input XML* and select your `.xml` export file.
2. Click **Browse** next to *Output Folder* and choose where results should be saved.
3. Click **Run Extraction**.
4. Watch the progress log — extraction typically takes 2–10 minutes depending on repository size.
5. When complete, click **Export to Excel** to save the results.

### Output files

| File | Contents |
|------|----------|
| `ds_lineage_results.xlsx` | Full lineage: column mappings, table relationships, summary sheet |
| `ds_column_lineage.csv` | Every source column → target column mapping |
| `ds_table_lineage.csv` | Table-to-table relationships only |
| `ds_lineage_summary.csv` | Counts and quality metrics |

---

## Tab 2 — Targeted Lineage

Use this tab when you want to trace the upstream lineage for **one or more specific target tables**, or for **an entire BODS job**.

### Search by target table

1. In the *Target Tables* field, type one or more table names separated by commas.
   - You can use the bare table name (e.g. `FACT_SALES`) or the schema-qualified name (e.g. `DBO.FACT_SALES`).
2. Set **Max Hops** — how many dataflow steps upstream to follow (default: 6). Increase if your lineage chain is long.
3. Click **Run Search**.
4. Results appear in the tree view on the left — expand any node to see sources, joins, and formula mappings.
5. Click **Export Excel** to save the column-level report.
6. Click **Open Graph** to open the interactive lineage graph in your browser.

### Search by job name

1. Enter the exact BODS job name in the *Job Name* field (e.g. `JOB_LOAD_SALES_DWH`).
2. Click **Run Job Trace**.
3. The engine automatically discovers all target tables written by that job and traces each one upstream.
4. The interactive graph opens with one tab per target table.

---

## Interactive Lineage Graph

The graph opens as a self-contained HTML file in your default browser. No internet connection is needed.

### Controls

| Action | What it does |
|--------|-------------|
| Scroll wheel | Zoom in / out |
| Click and drag | Pan the canvas |
| Click a node | Highlight all connected edges |
| Double-click an SQL node | Expand or collapse the SQL source cluster |
| **Export Tab** button | Download the current target tab as an Excel file |
| **Export All** button | Download all target tabs as a single multi-sheet Excel file |

### Reading the graph

The graph is drawn left-to-right by hop distance:

| Column position | Meaning |
|----------------|---------|
| Rightmost (Hop 0) | The target table you searched for |
| One step left (Hop 1) | Direct source tables or transforms |
| Further left (Hop 2, 3 …) | Upstream staging and source tables |
| Leftmost | Terminal sources — SAP extractors, flat files, Excel files |

### Node colour guide

| Colour | Node type |
|--------|-----------|
| Blue | Target table (Hop 0) |
| Green | Intermediate DWH / staging table |
| Orange | Terminal SAP R/3 source table |
| Purple | Flat file or Excel source |
| Teal | SQL transform cluster |
| Grey | Other transform or lookup |

### SQL clusters

When a dataflow contains an embedded SQL transform, it appears as a **teal SQL node**. The label shows how many physical tables the SQL reads from (e.g. `SQL [▶ 4]`).

- **Double-click** the SQL node to expand it and see the individual source tables.
- **Double-click again** to collapse it back.
- Edges from upstream tables automatically reconnect to the cluster when it is collapsed, so there are no broken connections.

### Multiple target tabs

When you trace a job or multiple tables at once, the graph shows a **tab strip** at the top — one tab per target table. Click any tab to switch views. Each tab shows only the sources relevant to that specific target.

---

## Exporting Results

### From the tree view

- **Export Excel** — saves a column-level lineage report for all searched targets.

### From the graph

- **Export Tab** — exports the records for the currently visible tab to `.xlsx`.
- **Export All** — exports all tabs into one `.xlsx` file with a Summary sheet and one sheet per target table.

---

## Tips

- **Large repositories (200 MB+ XML):** the first run takes longer because the engine indexes the full file. Subsequent searches on the same session reuse the index.
- **Table not found:** check the spelling and schema prefix. Try both `MY_TABLE` and `DBO.MY_TABLE`.
- **Too many hops:** if the graph is very wide, reduce Max Hops to 3–4 to focus on the closest upstream layers.
- **Browser shows old graph:** the engine adds a timestamp to every HTML filename so each run opens a fresh file. If you see a cached page, press `Ctrl + Shift + R` to force-reload.
- **Save the log:** if something goes wrong, copy the log text from the application window and share it with your support contact.

---

## Rebuilding the executable (for maintainers)

After making code changes, rebuild the standalone `.exe` with:

```powershell
.\build_exe.ps1
```

End users must replace their existing `.exe` with the newly built file. A running instance does not pick up code changes until it is closed and restarted.
