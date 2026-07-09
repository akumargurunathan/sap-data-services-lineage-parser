# AI Integration Notes — DOPPIA MACCHINA G2
**Date:** 2026-07-09  
**Topic:** Connecting the SAP BODS Lineage Engine with AI models (Anthropic Claude)

---

## 1. Context

DOPPIA MACCHINA G2 is a Python desktop app (Tkinter) that parses SAP BODS XML exports and extracts end-to-end data lineage. It already produces rich structured data — hop maps, upstream trees, column-level mapping rows with formulas — but had no AI layer.

The conversation covered two goals:
1. What AI capabilities make sense for this app?
2. How to generate a low-level functional specification document using AI?

---

## 2. AI Integration Ideas (discussed)

### Option 1 — Formula / Expression Explainer *(Easiest, highest value)*
- **File:** `ds_engine/ds_formula_semantics.py`
- `extract_formula_semantics()` currently uses regex to classify formulas.
- Add an AI call to produce a `"AI_Explanation"` field — a plain-English sentence per formula.
- Output appears as an extra column in Excel reports.

### Option 2 — Lineage Summary Narrative *(Medium effort)*
- **File:** `ds_engine/targeted_lineage_runner.py`
- The `result` dict is already structured (target_table, terminal_sources, hop_map, table_context).
- Feed it to an LLM and get a paragraph like:  
  *"Table TBDWFT_SF_WORKORDER is populated by Job JB_SF_WORKORDER via 3 upstream hops. Its primary source is SAP table AUFK..."*

### Option 3 — Natural Language Query Panel *(Most powerful, most work)*
- Add a chat input box in the UI.
- User types: *"Which jobs write to the DWH layer?"*
- LLM gets the full lineage JSON as context (RAG pattern) and answers directly.

### Option 4 — Auto-fill Description Fields *(Quick win)*
- **File:** `ds_engine/ds_metadata_fullpass.py`
- The `"Description": ""` field is always blank.
- AI infers business meaning from column name + formula + source table name.

---

## 3. Functional Document Generation (implemented)

### Goal
Generate a **low-level Word (.docx) functional specification document** from the targeted lineage result — covering every column mapping, all transformation formulas, join/filter conditions, and AI-generated plain-English explanations.

### Document Structure (per target table)

| Section | Content | Source |
|---|---|---|
| 1. Overview | Business purpose, sources, transformation summary | AI (Claude) |
| 2. Data Flow Summary | Job → Workflow → Dataflow chain; terminal sources with hop distance | `hop_map`, `table_context` |
| 3. Column-Level Mapping | Full table per dataflow: Target Column, Target Table, Source Column, Source Table, Transform Type, Formula, Functions Used, Has Agg, Has Cond, Complexity, Nesting, Risk | `result["rows"]` |
| 3.x Join Conditions | Join type, left/right table, condition expression | `JOIN` record rows |
| 3.x Filter Conditions | WHERE clause expressions per dataflow | `FILTER` record rows |
| 4. Complex Transformations Explained | Plain-English AI explanation per complex formula | AI (batched) |
| 5. Transformation Summary | Count by type; aggregation/conditional/high-risk counts | Row stats |
| 6. Assumptions & Notes | Risk flags, testing recommendations, data quality notes | AI |

### Complexity Threshold (what gets AI-explained)
A formula is flagged as "complex" if **any** of:
- `Formula_Complexity >= 3`
- `Nesting_Depth >= 2`
- `Has_Aggregation == "Yes"`
- `Has_Conditional == "Yes"`
- `Formula_Risk in ("HIGH", "VERY_HIGH")`

---

## 4. Files Changed / Created

### New file: `ds_engine/ds_ai_doc_generator.py`
Public API:
```python
from ds_engine.ds_ai_doc_generator import generate_functional_doc

generate_functional_doc(
    results=list_of_result_dicts,   # from TargetedLineageRunner
    output_path="path/to/output.docx",
    api_key="sk-ant-...",
    status_callback=print,          # optional — receives progress messages
)
```

Internal functions:
| Function | Role |
|---|---|
| `_group_rows_by_dataflow(rows)` | Groups row list into `{dataflow_key: {column_mappings, joins, filters}}` |
| `_is_complex(row)` | Returns True if the row's formula meets the complexity threshold |
| `_ai_overview(client, result)` | Calls Claude for a business-English overview paragraph |
| `_ai_formula_explanations(client, rows)` | Batches up to 25 complex rows into one Claude call; returns `{target_column: explanation}` |
| `_ai_assumptions(client, result, stats)` | Calls Claude for bullet-point assumptions & notes |
| `generate_functional_doc(...)` | Main entry point — builds the full Word document |

### Modified: `ds_ui_launcher.py`

**Button added** (Tab 2, export row, line ~2158):
```python
self.ai_doc_btn = ttk.Button(
    exp_frame, text="AI Functional Doc",
    command=self._generate_functional_doc, state=tk.DISABLED,
)
```
- Disabled until a targeted search completes.
- Re-disabled (and text reset) when a new search clears.

**Method added** after `_export_targeted_excel()`:
```python
def _generate_functional_doc(self):
    # Validates ANTHROPIC_API_KEY env var
    # Asks for save path (or uses output directory)
    # Runs ds_ai_doc_generator.generate_functional_doc() in a daemon thread
    # Streams progress to the log panel via root.after()
```

**Other touch points:**
- `_targeted_finished()` — enables `ai_doc_btn` alongside `graph_btn`
- `_clear_targeted()` — disables `ai_doc_btn`

---

## 5. Row Fields Available (for AI context building)

### COLUMN_MAPPING rows
| Field | Description |
|---|---|
| `Record_Type` | COLUMN_MAPPING / UNION_COLUMN_MAPPING / SQL_TRANSFORM / etc. |
| `Source_Object` | Source table name |
| `Source_Column` | Source column name |
| `Target_Object` | Target table name |
| `Target_Column` | Target column name |
| `Formula` | Raw BODS formula |
| `Formula_Clean` | Cleaned/readable formula |
| `Transformation_Type` | DIRECT_SOURCE / DERIVED / UNION / JOIN_KEY |
| `Transformation_Category` | PASS_THROUGH / CALCULATED |
| `Functions_Used` | Comma-separated BODS function names |
| `Formula_Type` | Semantic classification (from ds_formula_semantics) |
| `Formula_Complexity` | Numeric complexity score |
| `Formula_Category` | Transformation category |
| `Formula_Risk` | LOW / MEDIUM / HIGH / VERY_HIGH |
| `Has_Aggregation` | "Yes" / "No" |
| `Has_Conditional` | "Yes" / "No" |
| `Nesting_Depth` | Integer — how deeply nested the expression is |
| `Job_Name` | ETL Job name |
| `Workflow_Name` | Workflow name |
| `Dataflow_Name` | Dataflow name |

### JOIN rows
| Field | Description |
|---|---|
| `Join_Type` | INNER / LEFT / RIGHT / FULL |
| `Join_Left_Object` | Left table |
| `Join_Right_Object` | Right table |
| `Join_Condition` | Full join expression |

### FILTER rows
| Field | Description |
|---|---|
| `Where_Condition` | Full WHERE clause expression |

---

## 6. Setup

```powershell
pip install anthropic python-docx

# Set API key before launching the app
$env:ANTHROPIC_API_KEY = "sk-ant-..."

python ds_ui_launcher.py
```

---

## 7. Usage Flow

1. Launch app → go to **Tab 2 (Targeted Lineage)**
2. Enter target table name(s) → click **Run Search**
3. Wait for results to populate in the tree
4. Click **"AI Functional Doc"** (in the Export row)
5. Choose save path (or it auto-saves to the output directory)
6. Watch progress in the log panel
7. Open the generated `.docx` — landscape A4, sections per target table

---

## 8. Extending the AI Layer

### Add formula explainer to Excel export
In `ds_engine/ds_formula_semantics.py`, after `extract_formula_semantics()`, call the AI for any row where `Formula_Risk == "HIGH"` and append an `AI_Explanation` column to the DataFrame before writing to Excel.

### Natural language query (future)
- Store all `targeted_results` as a JSON blob in memory after each search.
- Add a chat text box in the UI.
- On submit: send the JSON + user question to Claude with `max_tokens=600`.
- Display the answer below the chat box.

### Local / offline model (no internet / sensitive data)
Replace the Anthropic client with Ollama:
```python
import ollama
response = ollama.chat(model="llama3", messages=[{"role": "user", "content": prompt}])
text = response["message"]["content"]
```
Same prompt structure works — just swap the client call.

---

## 9. API Call Strategy

| Call | When | Tokens (approx) |
|---|---|---|
| `_ai_overview` | Once per target table | ~400 out |
| `_ai_formula_explanations` | Once per target table (batches up to 25 complex rows) | ~800 out |
| `_ai_assumptions` | Once per target table | ~500 out |

For a single target table: **3 API calls**, ~1,700 output tokens total.  
For 5 target tables: **15 API calls**, ~8,500 output tokens total.

Model used: `claude-sonnet-5` (fast, cost-effective for structured generation tasks).
