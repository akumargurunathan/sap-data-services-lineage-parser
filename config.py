# config.py

from collections import defaultdict
from pathlib import Path
from datetime import datetime
from copy import deepcopy
import json
import pprint

# ============================================================
# INPUT SETTINGS
# ============================================================
# Where the XML files are located (folder or a single file)
INPUT = {
    'path': 'C:\\Users\\ajith.k.gurunathan\\Downloads\\ENGINE\\SCAN_OUT_FINAL\\CVC_QM_DELIVER\\bundles\\CVC_QM_DELIVERY\\xml',
    'recursive': False,
}

OUTPUT = {
    'output_dir': 'C:\\Users\\ajith.k.gurunathan\\Downloads\\ENGINE\\SCAN_OUT_FINAL\\CVC_QM_DELIVER\\bundles\\CVC_QM_DELIVERY\\eXT',
    'excel_name': 'CV_NIMS_RIPARAZIONI_PIVOT.xlsx',
    'handoff_rows_path': 'C:\\Users\\ajith.k.gurunathan\\Downloads\\ENGINE\\SCAN_OUT_FINAL\\CVC_QM_DELIVER\\bundles\\CVC_QM_DELIVERY\\eXT\\handoff_rows',
    'heavy_excel_name': 'CV_NIMS_RIPARAZIONI_PIVOT.json',
}
PIPELINE = {
    'handoff_enabled': True,
    'handoff_format': 'csv',
    'trigger_heavy_engine': False,
    'heavy_engine_timeout_sec': 0,
}
# Optional: technical->logical name mapping and label normalization
NAME_MAPPING = { "object": {} }
NORMALIZATION = { "case_policy": "upper" }

# (Optional) Smart export behavior for the full XML extract
SMART_EXPORT = {
    'enable': True,
    'prefer': 'xlsx',
    'max_excel_rows': 1000000,
    'split_excel_if_needed': True,
    'excel_sheet_rows': 1000000,
    'csv_delimiter': ',',
}

# ------------------------------------------------------------
# SORTING (HeavyEngine)
#   "normal" → SourceFile ascending (small → large)
#   "hana"   → Bottom-up HANA dependency layering
# ------------------------------------------------------------
SORTING = {
    "mode": "hana" # change to "normal" for simple SourceFile-based sort
}

# ------------------------------------------------------------
# ORDERING / PRESENTATION (used by HANA bottom-up mode)
# ------------------------------------------------------------
ORDERING = {
    "mode": "top_down",             # "hana_bottom_up" or "top_down"
    "include_join_columns": True,
    "include_source_only": False,          # treat rows with only sources as rank 0
    "max_relax_passes": 12                # cycle guard
}

# ------------------------------------------------------------
# TRACE: Recursive column lineage output (HeavyEngine)
# ------------------------------------------------------------

TRACE = {
    "enabled": True,

    # "all"          → trace every column in the handoff (complete column-level DAG)
    # "targets_only" → trace only final output columns (columns not used as source elsewhere)
    # ["COL_A",...]  → trace a specific named column list (for targeted debugging)
    "mode": "all",

    # SAP HANA stack: CV → ANV/DIM → ATV → Source Table = 4-5 hops minimum.
    # Stacked CVs (CV1 → CV2 → CV3 → ANV → ATV → table) can reach 8-9 hops.
    # 10 guarantees physical sources are always reached.
    "max_hops": 10,

    # Per-column row safeguard (prevents one extremely wide column from dominating memory)
    "limit_rows_per_column": 500000,

    # Pull column references from Formula / Filter / Join expressions (R2 per paper)
    "include_expr_tokens": True,

    # Keep pass-throughs and rename chains — required for complete lineage evidence
    "collapse_passthrough": False,
    "coalesce_renames":     False,

    # CRITICAL: True prevents combinatorial hop explosion.
    # False means every previously-seen (xpath, col, hop) path gets re-emitted
    # on each traversal step — causes edge count to grow as N^hops.
    "dedupe_per_hop": True,

    # Inherit node-level filter predicates into traced rows
    "inherit_filters_from_nodes": True,

    # Raised from 500 K — complex multi-view HANA models with 200+ columns
    # × 10 hops can generate several million edges before deduplication.
    # At ~200 bytes/edge, 5 M edges ≈ 1 GB RAM — safe on standard workstations.
    "max_edges": 5000000,

    # Keep full row evidence (all handoff columns) in each emitted DAG edge.
    # Required by the DAG visualiser and the lineage_engine.py graph builder.
    "preserve_target_evidence": True,

    # Stop traversal once a physical source table / DataSource is reached.
    "stop_at_physical_source": True,
}

# ============================================================
# PERFORMANCE / MEMORY HANDLING
# ============================================================
PERF = {
    "use_iterparse": True,            # streaming parse for big XML
    "iterparse_events": ("end",),     # emit nodes on 'end'
    "sqlite_sink_path": r"C:\temp\heavy_engine.db",         # e.g., r"C:\temp\lineage_sink.db"; None -> temp file
    "chunk_rows": 50000,              # batch size when exporting from SQLite
    "enable_diagnostics": False,      # turn on only when debugging
    "build_lineage_graph": True,     # skip combined_graph unless needed
    "node_outputs_emit": "list",     # "none" | "count" | "list" (default keeps memory tiny)
    "node_outputs_max_cols": 2000,    # only when node_outputs_emit == "list"
    "node_outputs_max_chars": 500000, # safety cap
}

# ============================================================
# OFFLINE CATALOG (critical for 100% completeness)
# ============================================================
OFFLINE_CATALOG = {
    "enabled": True,  # ✅ ENABLED - Critical for metadata completeness
    "dir": r"C:\Users\ajith.k.gurunathan\Downloads\ENGINE\DOPPIA_MACCHINA_G2\_catalog",

    # ✅ View-based system (no tables)
    "table_columns": None,
    "view_columns": "SYS_VIEW_COLUMNS.csv",
    "object_dependencies": "SYS_OBJECT_DEPENDENCIES.csv",

    "cols": {
        # Canonical column catalog
        "VIEW_COLUMNS": [
            "SCHEMA_NAME",
            "VIEW_NAME",
            "COLUMN_NAME",
            "POSITION",
            "DATA_TYPE_NAME",
            "LENGTH",
            "SCALE",
            "IS_NULLABLE"
        ],

        #  Runtime dependency resolution
        "OBJECT_DEPENDENCIES": [
            "BASE_SCHEMA_NAME",
            "BASE_OBJECT_NAME",
            "BASE_OBJECT_TYPE",
            "DEPENDENT_SCHEMA_NAME",
            "DEPENDENT_OBJECT_NAME",
            "DEPENDENT_OBJECT_TYPE",
            "DEPENDENCY_TYPE"
        ]
    },

    "csv": {
        "has_header": True,
        "sep": ",",
        "encoding": "utf-8",
        "quotechar": "\"",
        "doublequote": False,
        "escapechar": "\\",
        "engine": "python",
        "on_bad_lines": "skip",  # or "skip" or "error"
        "clean_first": False
    }
}
# ------------------------------------------------------------
# VALIDATION: Input lineage trust checks (HeavyEngine)
# ------------------------------------------------------------
VALIDATION = {
    "enabled": True,

    # Gate behavior after validation:
    # - "fail_fast": stop early if status resolves to FAIL (writes validation-only file)
    # - "warn_continue": continue; attach validation sheets to final Excel
    # - "report_only": ignore status for flow control; attach sheets
    "mode": "warn_continue",

    # Score threshold required to PASS when there are NO hard FAIL issues in summary.
    # (Hard FAIL rules always take precedence unless you override/downgrade them below.)
    "min_quality_score": 85,  # 0..100

    # Limit rows per rule written to VALIDATION_REPORT
    "max_error_rows": 10000,

    # Columns that must exist in the handoff; if any are missing -> FAIL
    "required_columns": [
        "Current_Node","Source_Object","Source_Column","Target_Column",
        "Transformation_Type","Formula","Filter_Expression","Join_Expression",
        "Join_Left_Object","Join_Left_Column","Join_Right_Object","Join_Right_Column"
    ],

    # Duplicate detection key (info/warn)
    "dedupe_key": [
        "SourceFile","XPath","Transformation_Type","Current_Node","Parent_Input_Object",
        "Target_Column","Source_Column","Formula","Filter_Expression","Join_Expression",
        "Join_Left_Column","Join_Right_Column"
    ],

    # Control whether HeavyEngine runs expensive trace when validation FAILs
    "run_trace_on_fail": True,

    # Severity overrides per rule (optional). Example:
    # "join_incomplete_side": "WARN", "mapping_without_source": "WARN", "missing_required_columns": "FAIL"
    "severity_overrides": {
         "join_incomplete_side": "WARN",
         "mapping_without_source": "WARN",
         "missing_required_columns": "FAIL"
    },

    # Ignore entire rules (they won't affect score/status, and won't appear in reports)
    "ignore_rules": [
         "cycle_detected"
    ],

    # Rule weights: we subtract each weight at most once per rule (not per row)
    "weights": {
        "missing_required_columns": 30,
        "mapping_without_source": 20,
        "empty_filter_expression": 10,
        "join_incomplete_side": 15,
        "star_unexpanded": 10,
        "bad_column_name": 5,
        "duplicate_rows": 5,
        "cycle_detected": 5,
        # NEW (readiness checks in heavy / optional):
        "canonical_columns_missing": 5,
        "canonical_columns_incomplete": 5,
        "system_default_used": 5
    }
}

# ============================================================
# WEIGHTS (Paper-Accurate Transformation & Filter Weight System)
# ============================================================
WEIGHTS = {
    # --------------------------------------------------------
    # Positive (light) transformation functions:
    # These reduce the "function count" in Wt/Wf normalization.
    #
    # Derived from the paper: CAST, ROUND, COALESCE, TRIM, UPPER,
    # LOWER, NVL, LTRIM/RTRIM, SUBSTR (light), REPLACE (light).
    # --------------------------------------------------------
    "positive_functions": [
        "CAST", "ROUND", "COALESCE", "NVL",
        "TRIM", "LTRIM", "RTRIM",
        "UPPER", "LOWER",
        "SUBSTR", "REPLACE"
    ],

    # --------------------------------------------------------
    # Wt/Wf clamping
    # Paper: Wt and Wf must be in [0,1], so clamp to this range.
    # --------------------------------------------------------
    "min_weight": 0.0,
    "max_weight": 1.0,

    # --------------------------------------------------------
    # JOIN keys are part of filter evidence (Definition 2, R2)
    # --------------------------------------------------------
    "join_keys_contribute_to_Wf": True,

    # --------------------------------------------------------
    # Additional stability options (Engine-compatible)
    # --------------------------------------------------------

    # Minimum number of columns required to consider meaningful weight
    "min_columns_for_weight": 1,

    # Treat CASE WHEN … THEN … ELSE … END as multiple functions
    # (Matches the paper’s treatment of predicate-heavy expressions)
    "case_expression_penalty": 0,

    # If True, subtract 1 from func-count only if the function
    # wraps a single column (identity-like, paper’s rule)
    "positive_fn_requires_single_column": False
}
# ============================================================
# PROPAGATION (Rule System R1, R2, R3 — Paper-Aligned)
# ============================================================
PROPAGATION = {
    # --------------------------------------------------------
    # Enable or disable propagation rules entirely.
    # --------------------------------------------------------
    "enabled": True,

    # --------------------------------------------------------
    # R2: Filter propagation
    # - Columns used in filters (JOIN / WHERE) propagate as
    #   impact edges with Wf.
    # --------------------------------------------------------
    "propagate_filter_influence": True,

    # Strength of filter influence when propagated
    # (multiplier applied to Weight_Filter during propagation)
    "filter_weight_factor": 1.0,

    # --------------------------------------------------------
    # R3: Parent roll-up (table → schema, column → table)
    # - Summarizes lineage at higher levels.
    # --------------------------------------------------------
    "parent_rollup": {
        "enabled": False, # turn OFF for column-level DAG
        # Levels:
        #   "table" = roll up to table-level
        #   "schema" = roll up to schema-level
        "levels": ["table"],

        # Aggregation rule
        #   max   = keeps strongest weight (default)
        #   sum   = accumulates weights
        #   mean  = average
        "aggregation": "max"
    },

    # --------------------------------------------------------
    # R1: Aggregation of parallel edges
    # - Multiple source→target transformations aggregated.
    # --------------------------------------------------------
    "aggregate_parallel_edges": True,

    # Aggregation mode:
    #   "max"            — use strongest influence
    #   "sum"            — accumulate
    #   "weighted_mean"  — average based on Wt/Wf (paper style)
    "parallel_edge_aggregation": "weighted_mean"
}

# ============================================================
# AUGMENTATION PIPELINE (single pass)
# ============================================================
AUGMENTATION_PIPELINE = {
    "run_offline_catalog_enrichment": True,
    "run_runtime_dependencies": True,      # enable only if you actually need it
}
# To surface Scenario_* columns on each row (CV only)
AUGMENTATION_SCENARIO_ATTRS = [
  "id","dataCategory","calculationScenarioType","visibility",
  "applyPrivilegeType","checkAnalyticPrivileges","defaultClient","defaultLanguage"
]
# ============================================================
# OUTPUT TOGGLES
# ============================================================
OUTPUT_TOGGLES = {
    "write_parquet": False,
    "write_excel": False,   # Excel is memory-heavier; keep opt-in
}

# ============================================================
# XML STRUCTURE CONFIGURATION (CV / ANV / ATV) -- (kept from your file)
# Only structure is shown here in condensed/clean form.
# ============================================================
XML_STRUCTURE_CV = {
    "current_node": {"tag": "calculationView", "node_id_attr": "id", "node_type_attr": "xsi:type", "capture_xpath": True},
    "parent_input_object": {"tag": "input", "reference_attr": "node", "capture_xpath": True},
    "source_target_multi": [
        {"parent_tag": "input", "mapping_tag": "mapping",
         "source_attr_path": "@source", "source_schema_path": None, "source_table_path": None,
         "target_attr": "target", "capture_xpath": True},
        {"parent_tag": "attributes", "mapping_tag": "attribute",
         "source_attr_path": "keyMapping@columnName", "source_schema_path": None,
         "source_table_path": "keyMapping@columnObjectName",
         "target_attr": "id", "capture_xpath": True},
        {"parent_tag": "baseMeasures", "mapping_tag": "measure",
         "source_attr_path": "measureMapping@columnName", "source_schema_path": None,
         "source_table_path": "measureMapping@columnObjectName",
         "target_attr": "id", "transformation_type": "Aggregation",
         "formula_attr_path": "@aggregationType", "capture_xpath": True},
    ],
    "calculated_multi": [
        {"tag": "calculatedViewAttribute", "expression_tag_path": "formula",
         "expression_attr": None, "target_attr": "id",
         "target_from_parent": False, "transformation_type": "Calculated_Column",
         "capture_xpath": True},
        {"tag": "calculatedAttribute", "expression_tag_path": "keyCalculation/formula",
         "expression_attr": None, "target_attr": "id",
         "target_from_parent": False, "transformation_type": "Calculated_Column",
         "capture_xpath": True},
        {"parent_tag": "calculatedMeasures", "mapping_tag": "measure",
         "expression_tag_path": "formula", "expression_attr": None,
         "target_attr": "id", "target_from_parent": False,
         "description_attr_path": "descriptions@defaultDescription",
         "meta_attrs": ["aggregationType","engineAggregation","measureType",
                        "calculatedMeasureType",
                        "datatype","expressionLanguage","length","scale"],
         "transformation_type": "Calculated_Measure", "capture_xpath": True},
    ],
    "node_mapping_multi": [
    {
        # The parent that contains mapping children.
        # In many CVs, mappings live under <viewNode> children (Projection/Aggregation/Union).
        "parent_tag": "viewNode",

        # Only activate when we are inside specific node types (by ancestor local-name()).
        # Tune these to match your actual xsi:type names in the scenario.
        "within_node_tags": ["projectionNode", "aggregationNode", "unionNode"],

        # The element names that hold source→target pairs under that parent.
        # Keep "mapping"; add alternates if your XML uses others (e.g., "attributeMapping").
        "mapping_tags": ["mapping", "attributeMapping", "attribute"],

        # Attribute names for the target and source on those mapping elements.
        # Add/adjust alternates if your XML differs (e.g., "targetName"/"sourceName").
        "target_attr_alts": ["target", "targetName", "id", "name"],
        "source_attr_alts": ["source", "sourceName", "attribute"],

        # How you want them labeled in the handoff
        "transformation_type": "Node_Mapping",
    }
],

# ==============================
# XML STRUCTURE CONFIGURATION – CV (PATCH)
# ==============================

    "filter_multi" : [
        {
            "tag": "filter",
            "condition_tag": None,
            # capture attributes present on <filter .../>
            "condition_expression_attr": ["value", "including", "operator", "xsi:type"],
            # capture the parent viewAttribute id
            "attribute_id_path": "../@id",
            # (optional) compose a readable predicate
            "compose_predicate": {
                "operator_map": {
                    "EQ": "{attr} = '{value}'",
                    "NE": "{attr} <> '{value}'",
                    "NL": "{attr} NOT LIKE '{value}'",  # seen in your file
                    "LE": "{attr} <= '{value}'",
                    "GE": "{attr} >= '{value}'",
                    "LT": "{attr} < '{value}'",
                    "GT": "{attr} > '{value}'"
                },
                "fallback_to_including": True,
                "if_including_true":  "{attr} = '{value}'",
                "if_including_false": "{attr} <> '{value}'"
            },
            "capture_xpath": True
        }
    ],

    "join": {
        "tag": "calculationView",
        "join_name_attr": "id",
        "condition_tag": None,
        "condition_expression_attr": "joinType",
        "left_schema_path": None,
        "left_object_path": "input[1]@node",
        "left_column_path": None,
        "right_schema_path": None,
        "right_object_path": "input[2]@node",
        "right_column_path": None,
        "capture_xpath": True,
        "parse_expression_columns": False,
        "join_attr_tag": "joinAttribute",
    },
    "source_table": {
        "tag": "DataSource",
        "table_name_attr": "id",
        "schema_attr": None,
        "catalog_attr": None,
        "column_tag": None,
        "column_name_attr": None,
        "qualify_columns": False,
        "emit_star_row_if_no_columns": False,
        "capture_xpath": True
    },
    "parent_input_fallback": "__CURRENT_NODE__",
}
# --- Only this block changes ---

XML_STRUCTURE_ANV = {
    "current_node": {"tag": "cube", "node_id_attr": "id", "node_type_attr": None, "capture_xpath": True},
    "parent_input_object": {"tag": None, "reference_attr": None, "capture_xpath": True},
    "source_target_multi": [
        {"parent_tag": "attributes", "mapping_tag": "attribute",
         "source_attr_path": "keyMapping@columnName",
         "source_schema_path": "keyMapping@schemaName",
         "source_table_path": "keyMapping@columnObjectName",
         "target_attr": "id", "capture_xpath": True},
        {"parent_tag": "baseMeasures", "mapping_tag": "measure",
         "source_attr_path": "measureMapping@columnName",
         "source_schema_path": "measureMapping@schemaName",
         "source_table_path": "measureMapping@columnObjectName",
         "target_attr": "id", "transformation_type": "Aggregation",
         "formula_attr_path": "@aggregationType",
         "capture_xpath": True},
    ],
    "calculated_multi": [
        {"tag": "calculatedAttribute", "expression_tag_path": "keyCalculation/formula",
         "target_attr": "id", "transformation_type": "Calculated_Column",
         "capture_xpath": True},
        {"tag": "calculatedViewAttribute", "expression_tag_path": "formula",
         "target_attr": "id", "transformation_type": "Calculated_Column",
         "capture_xpath": True},
        {"parent_tag": "calculatedMeasures", "mapping_tag": "measure",
         "expression_tag_path": "formula", "target_attr": "id",
         "description_attr_path": "descriptions@defaultDescription",
         "meta_attrs": ["aggregationType","engineAggregation","measureType",
                        "calculatedMeasureType",
                        "datatype","expressionLanguage","length","scale"],
         "transformation_type": "Calculated_Measure",
         "capture_xpath": True},
    ],

 # ==============================
# XML STRUCTURE CONFIGURATION – ANV (OPTIONAL PATCH)
# ==============================
# Replace only the "filter_multi" section inside XML_STRUCTURE_ANV with the block below.
# Safe to apply; ignored if those tags do not appear.


# config.py  → CV filter structure
    "filter_multi" : [
        {
            "tag": "filter",
            "condition_tag": None,
            # capture attributes present on <filter .../>
            "condition_expression_attr": ["value", "including", "operator", "xsi:type"],
            # capture the parent viewAttribute id
            "attribute_id_path": "../@id",
            # (optional) compose a readable predicate
            "compose_predicate": {
                "operator_map": {
                    "EQ": "{attr} = '{value}'",
                    "NE": "{attr} <> '{value}'",
                    "NL": "{attr} NOT LIKE '{value}'",  # seen in your file
                    "LE": "{attr} <= '{value}'",
                    "GE": "{attr} >= '{value}'",
                    "LT": "{attr} < '{value}'",
                    "GT": "{attr} > '{value}'"
                },
                "fallback_to_including": True,
                "if_including_true":  "{attr} = '{value}'",
                "if_including_false": "{attr} <> '{value}'"
            },
            "capture_xpath": True
        }
    ],

    "join": {
        "tag": "join",
        "join_name_attr": None,
        "condition_tag": "properties",
        "condition_expression_attr": "joinType",
        "left_schema_path": "leftTable@schemaName",
        "left_object_path": "leftTable@columnObjectName",
        "right_schema_path": "rightTable@schemaName",
        "right_object_path": "rightTable@columnObjectName",
        "capture_xpath": True,
        "parse_expression_columns": False,
    },
    "source_table": {
        "tag": "table",
        "table_name_attr": "columnObjectName",
        "schema_attr": "schemaName",
        "qualify_columns": True,
        "emit_star_row_if_no_columns": False,
        "capture_xpath": True
    },
    "restricted_measures": {
        "parent_tag": "restrictedMeasures",
        "measure_tag": "measure",
        "id_attr": "id",
        "base_measure_attr": "baseMeasure",
        "engine_agg_attr": "engineAggregation",
        "description_attr_path": "descriptions@defaultDescription",
        "restriction_tag": "restriction",
        "logical_op_attr": "logicalOperator",
        "filter_tag": "filter",
        "filter_attr": "attributeName",
        "value_filter_tag": "valueFilter",
        "value_attr": "value",
        "capture_xpath": True
    },
    "parent_input_fallback": "__CURRENT_NODE__",
}

XML_STRUCTURE_ATV = {
    "current_node": {"tag": "ColumnView", "node_id_attr": "name", "node_type_attr": None, "capture_xpath": True},
    "parent_input_object": {"tag": None, "reference_attr": None, "capture_xpath": True},
    "source_target_multi": [
        {"parent_tag": "input", "mapping_tag": "mapping",
         "source_attr_path": "@sourceName",
         "source_table_path": None,
         "target_attr": "targetName",
         "transformation_type": "Mapping",
         "capture_xpath": True},
    ],
    "calculated_multi": [],
    "filter_multi": [],
    "label_binding": {
        "element_tag": "element",
        "element_name_attr": "name",
        "label_tag": "labelElement",
        "capture_xpath": True
    },
    "join": {"tag": None},
    "source_table": {"tag": None},
    "parent_input_fallback": "__CURRENT_NODE__",
}

XML_STRUCTURE_DIM = deepcopy(XML_STRUCTURE_ANV)
XML_STRUCTURE_DIM["current_node"] = {"tag": "dimension", "node_id_attr": "id", "node_type_attr": None, "capture_xpath": True}
XML_STRUCTURE_DIM["parent_input_object"] = {"tag": None, "reference_attr": None, "capture_xpath": True}
XML_STRUCTURE_DIM["parent_input_fallback"] = "__CURRENT_NODE__"

# Combined structure dictionary
XML_STRUCTURE = {
    "CV": XML_STRUCTURE_CV,
    "ANV": XML_STRUCTURE_ANV,
    "ATV": XML_STRUCTURE_ATV,
    "DIM": XML_STRUCTURE_DIM,
}

AUGMENTATION_DIM_ATTRS = [
    "id", "dimensionType", "schemaVersion", "visibility",
    "applyPrivilegeType", "checkAnalyticPrivileges", "defaultClient",
    "defaultLanguage", "hierarchiesSQLEnabled", "translationRelevant",
    "type", "generateConcatAttributes"
]

# =======================
# AUGMENTATION (opt-in)
# =======================
AUGMENTATION = {
    # CV — scenario/datasource
    "enable_scenario_metadata": True,     # attach Scenario_* columns
    "enable_datasource_metadata": True,   # attach DataSource_* via <dataSources>
    # General
    "emit_constant_mappings": True,       # ConstantAttributeMapping
    "emit_logical_bindings": True,        # logicalModel keyMapping
    "log_selected_profile": True,
    "anv_emit_shared_dimensions": True,   # <sharedDimensions>/<logicalJoin>
    "anv_emit_private_joins":    True,    # <privateDataFoundation>/<joins> (Fix 1)
    "anv_emit_source_tables":    True,    # <privateDataFoundation> tableProxies (Fix 1)
    "anv_emit_table_filters":    True,    # <tableProxy>/<columnFilter> (Fix 1)
    "anv_emit_textual_metadata": True,    # descriptions + metadata timestamps (Fix 6)
    "anv_emit_exception_agg":    True,    # <exceptionAggregation> on calc measures (Fix 2)
    # ATV
    "atv_emit_entity_star": True,         # Source_Table star rows from <input><entity>
    "atv_emit_label_binding": True,       # Label_Binding rows from <element><labelElement>
    # DIM (if you add later)
    "dim_enable_metadata": True,
    "dim_enable_textual_metadata": True,
    "dim_emit_source_tables": True,
    "dim_emit_key_mappings": True,
    "dim_emit_attribute_flags": True,
    "dim_emit_joins": True,
    "dim_emit_table_filters": True,
    "dim_emit_calculated_attributes": True,
    "global_anv_pullthrough": True,
}

# ============================================================
# LOGGING
# ============================================================
LOGGING = {
    "level": "INFO"  # or "DEBUG"
}

# ============================================================
# LINEAGE MODES
# ============================================================
LINEAGE_ENGINE = {
    "extract_full": True,
    "extract_subset_only": False,
}

# ============================================================
# VISUALIZATION (optional)
# ============================================================
# VISUALIZATION = {
    # "enabled": False,
    # "target_columns": ["ADDRESS"],  # example
    # "use_traced_subset": False,
    # "show": True,
    # "save_path": None,
    # "layout_mode": "classic",
    # "namespace_targets_by_node": False,
    # "include_xpath": False,
    # "include_node_and_parent": True,
    # "include_null_sources": True,
    # "max_label_len": 50,
# }

# =====================================
# CLOSURE (materialized dependency TC)
# =====================================
CLOSURE = {
    "build_lineage_closure": True,      # precompute closure for fast tracing #“Show me EVERYTHING that could be impacted.”
    "store_pointer_encoding": True,     # pointer-like encoding to reduce memory
    "cycle_handling": {
        "break_cycles_for_build": True, # turn graph to DAG during build
        "restore_cycles_after": True    # reattach broken edges via BFS after build
    },
    "limits": {
        "max_nodes": 999999999,         # safety guard for very large graphs
        "max_edges": 999999999
    }
}

# ==============================
# SEMANTICS (semantic subgraph)
# ==============================
SEMANTICS = {
    "enable_semantic_layer": True,

    # Keep edges whose transform/filter evidence meets thresholds
    "min_transform_weight": 0.35,   # Wt >= 0.30
    "min_filter_weight": 0.25,      # Wf >= 0.20

    # Only keep nodes/edges that share relevant filter context around the focal node
    "require_filter_context_match": False,

    # When muting rather than removing, set transparency factor (if a UI consumes it)
    "mute_low_weight_edges_instead_of_drop": True,
    "mute_alpha": 0.25
}
# ===================================
# SCORING (LLD / LID per component)
# ===================================
SCORING = {
    "enabled": True,
    "compute_LLD": True,
    "compute_LID": True,

    # If a node lacks one side (sources or targets), avoid division-by-zero
    "epsilon": 1e-9,

    # Expose scores into rows (engine can surface these columns in exports)
    "emit_scores_into_rows": True,       # will add LLD_Score, LID_Score per node/row where meaningful
}

# ===============================================
# BUSINESS_SEMANTICS (ontology/terms generation)
# ===============================================
BUSINESS_SEMANTICS = {
        "enabled": True,

        "dictionary": {
                "enable_default_acronyms": True,
                # Add a few SAP/HANA-friendly recodes on top of your current set
                "default_map": {
                    "acct": "account", "agrmnt": "agreement", "nbr": "number", "cd": "code", "amt": "amount", "dt": "date",
                    "mat": "material", "prd": "product", "qty": "quantity", "cust": "customer", "ven": "vendor",
                    "uom": "unit of measure", "yr": "year", "mnth": "month"
                },
        "drop_terms": ["tbl", "vw", "tmp", "stg"]  # ignore purely technical suffixes/prefixes
    },

    # Tokenization & normalization
    "normalize": {
        "split_camel_case": True,
        "split_on_non_alnum": True,
        "lowercase": True,
        "titlecase_output": True,
        "drop_plurals": True,
        "min_token_len": 3
    },

    # Fuzzy matching to deduplicate concept candidates
    "fuzzy_matching": {
        "enabled": True,
        "similarity_threshold": 0.88
    },

    # Use lineage/impact to infer synonyms and relations
    "use_lineage_for_synonyms": True,   # same data content via transformations → semantic relation
    "use_joins_for_fk_relations": True, # joins/filters reveal associations & FK‑like links

    # Outputs (relative to OUTPUT['output_dir'])
    "export": {
        "concepts_path": "business_semantics_concepts.json",
        "ontology_path": "business_semantics_ontology.json"
    }
}

# (Extend your existing PIPELINE block)
PIPELINE.update({
    # New toggles — your xml_engine will check these before launching auxiliary steps
    'trigger_deep_recursive_engine': True,
    'trigger_closure_builder': False,     # set True when a closure_builder.py exists
    'trigger_semantics_builder': False,   # set True when a semantics_builder.py exists

    # Optional: build closure + semantic layer during main run
    'build_closure_during_run': True,     # if True, build materialized closure as part of the pipeline
    'compute_semantic_layer_during_run': True
})

ORDERING.update({
    "final_sort": "evidence",   # or "legacy"
    "evidence_order": "asc"     # "asc" (lower->higher) or "desc"
})


# THE BELOW IS FOR TEST PURPOSE DONT UNCOMMENT IT........AJITH
# ============================================================
# MINIMAL SAFE CONFIG — NO LINEAGE LOSS
# (xml_engine.py behaving like xml_lineage.py)
# ============================================================

# # ----------------------------
# # HARD RULE: KEEP ALL ROWS
# # ----------------------------
# PERF.update({
#     # Must preserve actual column names
#     "node_outputs_emit": "list",
#     "node_outputs_max_cols": 1000000,   # effectively unlimited
#     "node_outputs_max_chars": 10000000,

#     # No pruning / diagnostics / graphs
#     "enable_diagnostics": False,
#     "build_lineage_graph": False,
# })

# # ----------------------------
# # DISABLE ALL LINEAGE FILTERING
# # ----------------------------
# SEMANTICS.update({
#     "enable_semantic_layer": False,
# })

# PROPAGATION.update({
#     "enabled": False,
# })

# SCORING.update({
#     "enabled": False,
# })

# WEIGHTS.update({
#     # Keep weights computation OFF (no pruning logic)
#     "min_weight": 0.0,
# })

# # ----------------------------
# # TURN OFF ENRICHERS (BIGGEST SOURCE OF LOSS)
# # ----------------------------
# OFFLINE_CATALOG.update({
#     "enabled": True,
# })

# AUGMENTATION_PIPELINE.update({
#     "run_offline_catalog_enrichment": True,
#     "run_runtime_dependencies": True,
# })

# # ----------------------------
# # VALIDATION MUST NEVER DROP DATA
# # ----------------------------
# VALIDATION.update({
#     "enabled": False,
# })

# # ----------------------------
# # PIPELINE: STOP HANDOFF CHAINS
# # ----------------------------
# PIPELINE.update({
#     "handoff_enabled": True,
#     "trigger_heavy_engine": False,
#     "trigger_deep_recursive_engine": False,
#     "trigger_closure_builder": False,
#     "trigger_semantics_builder": False,
# })

# # ----------------------------
# # CLOSURE / GRAPH — OFF
# # ----------------------------
# CLOSURE.update({
#     "build_lineage_closure": False,
# })

# # ----------------------------
# # OUTPUT: SIMPLE, LOSSLESS
# # ----------------------------
# LINEAGE_ENGINE.update({
#     "extract_full": True,
#     "extract_subset_only": False,
# })

PROPAGATION.update({
    "parent_rollup": {"enabled": False},
    "aggregate_parallel_edges": False
})

# PERF = {
#     "use_iterparse": True,
#     "iterparse_events": ("end",),

#     # ✅ Use SQLite for large runs
#     "sqlite_sink_path": r"C:\temp\heavy_engine.db",

#     # Chunked export
#     "chunk_rows": 50000,

#     # Do not emit node outputs (huge memory saver)
#     "node_outputs_emit": "none",

#     # Keep graph building on
#     "build_lineage_graph": True,
# }


# OFFLINE_CATALOG = {
#     "enabled": True,
#     "dir": r"C:\Users\ajith.k.gurunathan\Downloads\ENGINE\DOPPIA_MACCHINA_G2\_catalog",
#     "view_columns": "SYS_VIEW_COLUMNS.csv",
#     "object_dependencies": "SYS_OBJECT_DEPENDENCIES.csv",
# }
