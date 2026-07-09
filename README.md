# DOPPIA MACCHINA G2

**End-to-end data lineage extraction engine for SAP Business Objects Data Services (BODS)**

Parses BODS XML repository exports and produces interactive lineage graphs, column-level mapping reports, and AI-generated functional specification documents — without any BODS server connection.

---

## What it does

SAP BODS stores all job/workflow/dataflow metadata in a single exportable XML file. DOPPIA MACCHINA G2 reads that XML and answers:

- Which source tables (SAP, flat files, Excel, staging DWH) ultimately feed a given target table?
- Through which jobs, workflows, and dataflows does data flow?
- What SQL transforms, joins, filters, and formula expressions are applied at each step?
- How many hops separate a target from its terminal sources?

Results are delivered as an interactive HTML canvas graph and as Excel/CSV reports.

---

## Key Features

| Feature | Description |
|---|---|
| **Full Extraction** | Parse the entire XML; export all-table lineage to Excel |
| **Targeted Lineage** | Trace one or more specific target tables upstream up to N hops |
| **Interactive Graph** | Zoomable/pannable canvas with layered hop view, SQL cluster expand/collapse |
| **Job-Level View** | Trace all targets of a BODS job in one run; tabbed graph per target |
| **SQL Cluster** | SQL transforms shown as collapsible clusters; upstream broken-edge fix keeps visual continuity |
| **Intra-dataflow Routing** | DIQuery/DITableSpec/DIInputView topology tracked — each target tab shows only its own SQL sources |
| **Column-Level Lineage** | Per-dataflow mapping rows: source column → formula → target column |
| **AI Functional Doc** | Claude-powered Word document *(work in progress — contributions welcome)* |
| **File/Excel Sources** | Flat-file and Excel datastores resolved to full paths from XML metadata |
| **SAP/ABAP Sources** | R/3 extractor and ABAP sub-flow sources propagated into the lineage graph |

---

## Architecture

```
ds_ui_launcher.py               ← Tkinter GUI + HTML template + all UI logic
│
├── ds_engine/
│   ├── targeted_lineage_runner.py   ← Core BFS engine; walks BODS XML; builds hop maps
│   ├── job_lineage_runner.py        ← Runs targeted engine for every target of a BODS job
│   ├── ds_lineage_engine.py         ← Full-extraction engine (all tables)
│   ├── ds_flat_export_indexer.py    ← Flat-repo index (sibling job/wf/df at root level)
│   ├── ds_expression.py             ← BODS formula/expression parser
│   ├── ds_formula_semantics.py      ← Formula classification (risk, complexity, type)
│   ├── ds_schema_enrichment.py      ← Column-level schema metadata enrichment
│   ├── ds_sql_parser.py             ← Physical table extraction from SQL text
│   ├── ds_context.py                ← Shared context objects
│   └── ds_ai_doc_generator.py       ← Claude-powered Word document generator
│
├── config.py                   ← App-level constants / default paths
├── xml_engine.py               ← Legacy standalone XML engine
└── tests/                      ← pytest suite
```

---

## Installation

```powershell
# Clone or unzip the project
cd "DOPPIA_MACCHINA_G2"

# Create a virtual environment (recommended)
python -m venv .venv
.\.venv\Scripts\Activate.ps1

# Install dependencies
pip install lxml openpyxl pandas

# Optional — for AI Functional Doc generation
pip install anthropic python-docx
```

**Python 3.9+** required. `tkinter` is bundled with the standard Python installer on Windows.

---

## Quick Start

```powershell
python ds_ui_launcher.py
```

### Tab 1 — Full Extraction

1. Browse to your `BODS_EXPORT.xml` file
2. Choose an output folder
3. Click **Run Extraction**
4. Export to Excel when complete

### Tab 2 — Targeted Lineage

1. Enter one or more target table names (comma-separated), e.g. `TBDWFT_INVOICES, DBO.IRI_ANAGCLI_NEW`
2. Set **Max Hops** (default 6)
3. Click **Run Search**
4. Explore results in the tree view
5. Click **Open Graph** to open the interactive HTML canvas in your browser
6. Click **Export Excel** for a detailed column-level report

### Job-Level Trace (advanced)

1. Enter the exact BODS job name, e.g. `SHXIRI01_DATI_OUT`
2. Click **Run Job Trace**
3. The engine finds all target tables for that job and traces each upstream automatically
4. The resulting HTML has one tab per target table

---

## Interactive Graph

The HTML graph is fully self-contained (no server required). Key controls:

| Action | Effect |
|---|---|
| Scroll wheel | Zoom in/out |
| Click + drag | Pan the canvas |
| Click a node | Highlight its edges |
| Double-click SQL node | Expand/collapse SQL cluster |
| **Export Tab** button | Download current tab as XLSX |
| **Export All** button | Download all tabs as multi-sheet XLSX |

### Node colours

| Colour | Type |
|---|---|
| Blue | Target table (Hop 0) |
| Green | Intermediate DWH/staging table |
| Orange | Terminal SAP R/3 source |
| Purple | Flat file / Excel source |
| Teal | SQL transform cluster |
| Grey | Other source/transform |

---

## Output Files

| File | Contents |
|---|---|
| `ds_lineage_results.xlsx` | Full-extraction: all column mappings, table lineage, summary |
| `targeted_lineage_<table>.xlsx` | Per-target column mappings + hop map |
| `job_lineage_<job>_<ts>.html` | Interactive graph for a job run |

---

## Result Dict Shape

`TargetedLineageRunner.run()` returns a Python dict:

```python
{
    "target_table":     str,            # e.g. "DBO.IRI_ANAGCLI_NEW"
    "hop_map":          {table: int},   # distance from target (0 = target itself)
    "upstream_tree":    {table: [parents]},
    "terminal_sources": set,            # leaf nodes (SAP, FILE, EXCEL, bare tables)
    "table_context":    {table: {"Job_Name": str, "Workflow_Name": str, "Dataflow_Name": str}},
    "sql_members":      {"SQL:<ck>": [physical_tables]},  # SQL cluster membership
    "sql_queries":      {"SQL:<ck>": raw_sql_text},       # SQL text for each transform
    "rows":             [lineage_row_dicts],               # column-level mapping rows
}
```

---

## AI Functional Document

> **This section is a work in progress — contributions welcome!**
>
> We are looking for contributors to help design and document the AI-powered functional specification feature. If you have experience with LLM integration, Word document generation (`python-docx`), or SAP BODS data lineage and would like to contribute, please open an issue or start a discussion.
>
> **What we need help with:**
> - Documenting the setup and usage flow end-to-end
> - Reviewing the Claude or any model prompt strategy for formula explanation
> - Testing against different BODS export structures
> - Suggesting the right output format (Word vs PDF vs Markdown)
>
> See [`ds_engine/ds_ai_doc_generator.py`](ds_engine/ds_ai_doc_generator.py) and [`AI_INTEGRATION_NOTES.md`](AI_INTEGRATION_NOTES.md) for the current implementation notes.

---

## XML Source Format

Tested against **SAP BODS 4.2 / 4.3** XML repository exports (File → Export → Repository). Both nested-hierarchy exports and flat-repository exports (where Job/Workflow/Dataflow appear as siblings at root level) are supported.

Key XML elements parsed:

| Element | Purpose |
|---|---|
| `DIDataflow / DIR3Dataflow` | Dataflow boundary |
| `DIDatabaseTableSource / Target` | DB source/target registration |
| `DISAPExtractorSource` | SAP R/3 table source |
| `DIFileSource / DIExcelSource` | Flat-file and Excel sources |
| `DITransformCall` + `sql_text` | SQL transform text extraction |
| `DIQuery` + `DITableSpec` | Query input wiring |
| `DIInputView / DIOutputView` | Intra-dataflow connection topology |
| `DIR3DataflowCall` | ABAP sub-flow call edges |

---

## Running Tests

```powershell
pytest
```

---

## Configuration

Edit `config.py` to set default paths:

```python
INPUT  = {"path": r"C:\BODS_EXPORTS\BODS_FULL_PROD.xml"}
OUTPUT = {"output_dir": r"C:\DS_Output"}
```

---

## Roadmap

- [ ] Natural language query panel (chat over lineage JSON via Claude)
- [ ] Formula explainer column in Excel export (`AI_Explanation` per high-risk formula)
- [ ] Column-level impact analysis (forward lineage: given a source column, which targets are affected?)
- [ ] Delta comparison between two XML exports (what changed between releases?)
- [ ] Offline/local model support via Ollama for air-gapped environments

---

## License

Internal tool — Accenture / Client use only. Not for public redistribution.
