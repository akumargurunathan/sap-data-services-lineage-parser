#!/usr/bin/env python3
# -*- coding: utf-8 -*-
#Crash avoiding enigne

"""
XML Lineage Extractor + Deep Recursive Tracer (SAP CV / ANV / ATV / DIM)

- Auto-selects model per file:
    * CV  -> config.XML_STRUCTURE_CV   (root: <calculationView> / <Calculation:scenario>)
    * ANV -> config.XML_STRUCTURE_ANV  (root local-name: cube, e.g., <Cube:cube>)
    * ATV -> config.XML_STRUCTURE_ATV  (root: <View:ColumnView>)
    * DIM -> (root: <Dimension:dimension>)

- Preserves your CV/ANV logic/results.
- Adds ATV (ColumnView) and captures Source_Object from <input><entity> when present.
"""
import os
import re
import glob
import logging
from itertools import zip_longest
from html import unescape as _html_unescape
import json
import offline_catalog as ocmod
from offline_catalog import get_runtime_columns  # (optional helper)
from collections import defaultdict

# Precompiled regexes (performance) #opitonl ho uso questo per scalabiliy--AJITH
RE_SINGLE_QUOTED = re.compile(r"'[^']*'")
RE_DOUBLE_QUOTED = re.compile(r'"([^"]+)"')
RE_IDENTIFIER    = re.compile(r'\b[A-Za-z_][A-Za-z0-9_.]*\b')

import pandas as pd
from lxml import etree
from typing import List, Dict, Tuple, Optional

# ---------------- Load config ----------------
try:
    import config
except ImportError:
    print("ERROR: config.py not found")
    raise SystemExit(1)

# ---------------- Logging ----------------
log_level = getattr(logging, config.LOGGING.get("level", "INFO"), logging.INFO)
logging.basicConfig(
    level=log_level,
    format="%(asctime)s | %(levelname)-8s | %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)
log = logging.getLogger("xml-lineage")

SEMANTICS = getattr(config, "SEMANTICS", {}) if hasattr(config, "SEMANTICS") else {}

def _s(x):
    return (str(x) if x is not None else "").strip()

def _wt_min():
    try:
        return float(SEMANTICS.get("min_transform_weight", 0.30))
    except Exception:
        return 0.30

def _wt_estimate(source: str, target: str) -> float:
    return 1.0 if _s(source) and _s(target) else 0.0



DEBUG_PRINTS = bool(config.LOGGING.get("debug_prints", False))

# ---------------- Constants & Normalizers ----------------
NULL_TOKEN = "<NULL>"  # MUST match your Excel sentinel exactly

SQL_KEYWORDS = {
    "AND","OR","NOT","NULL","CASE","WHEN","THEN","ELSE","END","IN","IS","ON","AS",
    "SUM","COUNT","MAX","MIN","AVG","UPPER","LOWER","TRIM","CAST","LIKE","WHERE",
    "GROUP","BY","ORDER","HAVING","DISTINCT","INNER","LEFT","RIGHT","FULL","JOIN"
}

def _to_str(x):
    return x if isinstance(x, str) else ""

def _norm_single(x: str) -> str:
    s = _to_str(x).strip()
    if not s or s == NULL_TOKEN:
        return ""
    if s.startswith('"') and s.endswith('"') and len(s) >= 2:
        s = s[1:-1]
    return s.split(".")[-1].strip()

def _norm_any(x) -> str:
    if isinstance(x, (list, tuple, set)):
        for item in x:
            n = _norm_single(item)
            if n:
                return n
        return ""
    return _norm_single(x)

def _trace_normalize_token(x) -> str:
    s = _to_str(x).strip()
    if not s or s == NULL_TOKEN:
        return ""
    if s.startswith('"') and s.endswith('"') and len(s) >= 2:
        s = s[1:-1]
    return s.strip()

def _qualify_lineage_column(row, object_keys, column_key):
    col = _to_str(row.get(column_key)).strip()
    if not col or col == NULL_TOKEN:
        return ""
    for object_key in object_keys:
        obj = _to_str(row.get(object_key)).strip()
        if obj and obj != NULL_TOKEN:
            return f"{obj}.{col}"
    return col

# ---------------- Helpers ----------------
def derive_lineage_intent(r):
    tt = (r.get("Transformation_Type") or "").upper()
    lvl = r.get("Lineage_Level")

    if lvl == "NODE":
        return "STRUCTURAL_DEPENDENCY"
    if "CALCULATED" in tt:
        return "DERIVED"
    if "AGGREGATION" in tt:
        return "AGGREGATED"
    if "FILTER" in tt:
        return "FILTERED"
    if "JOIN" in tt:
        return "JOIN_DERIVED"
    if "MAPPING" in tt:
        return "DIRECT_SOURCE"
    if "ROLLUP" in tt:
        return "INHERITED_FROM_VIEW"

    return "UNKNOWN"
def get_attr_or_text(elem, attr_name=None, use_text_if_no_attr=False):
    if elem is None:
        return NULL_TOKEN
    if attr_name:
        value = elem.attrib.get(attr_name, NULL_TOKEN)
        return value.strip() if value and value.strip() else NULL_TOKEN
    if use_text_if_no_attr:
        text_val = elem.text.strip() if elem.text and elem.text.strip() else NULL_TOKEN
        return text_val
    return NULL_TOKEN

def extract_columns_from_expression(expr):
    """Extract identifiers from expressions; ignore literals & SQL keywords."""
    if not expr or expr == NULL_TOKEN:
        return set()
    expr = str(expr)
    found = set()

    # remove string literals first
    expr_wo_literals = re.sub(r"'[^']*'", " ", expr)

    # quoted identifiers: "SCHEMA.OBJ.COL" or "COL"
    for q in re.findall(r'"([^"]+)"', expr_wo_literals):
        n = _norm_single(q)
        if n:
            found.add(n)

    # bare identifiers: A_Z0.. (include dots for qualifier)
    for tok in re.findall(r'\b[A-Za-z_][A-Za-z0-9_.]*\b', expr_wo_literals):
        if tok.upper() in SQL_KEYWORDS:
            continue
        n = _norm_single(tok)
        if n:
            found.add(n)

    return found

def _get_by_path(elem, path: str, *, text_if_no_attr=False, default=NULL_TOKEN):
    """Resolve paths like '@attr', 'child@attr', 'a/b@attr' (tag-name based)."""
    if elem is None or not path:
        return default

    path = str(path).strip()
    if path.startswith("@"):
        attr = path[1:].strip()
        if not attr:
            return default
        val = elem.attrib.get(attr, "").strip()
        return val if val else default

    cur = [elem]
    parts = path.split("/")
    for part in parts:
        if "@" in part:
            tag, attr = part.split("@", 1)
            out = []
            for e in cur:
                for ch in e.findall(tag):
                    val = ch.attrib.get(attr, "").strip()
                    if val:
                        out.append(val)
            return out[0] if out else default
        nxt = []
        for e in cur:
            found = e.findall(part)
            if found:
                nxt.extend(found)
        if not nxt:
            return default
        cur = nxt

    if text_if_no_attr:
        for e in cur:
            txt = (e.text or "").strip()
            if txt:
                return txt
    return default

def _find_central_table(root):
    """namespace-agnostic: tableProxy centralTable='true' -> schema.table"""
    try:
        for tp in root.xpath(".//*[local-name()='tableProxy']"):
            if (tp.attrib.get("centralTable", "false") or "").lower() == "true":
                t = (tp.xpath(".//*[local-name()='table']") or [None])[0]
                if t is not None:
                    sche = (t.attrib.get("schemaName") or "").strip()
                    name = (t.attrib.get("columnObjectName") or "").strip()
                    if sche and name:
                        return f"{sche}.{name}"
    except Exception:
        pass
    return None

def synthesize_from_node_outputs(rows, default_source_obj=None):
    """Create Synth_Mapping rows for visible outputs if central table is known."""
    out = []
    seen = set()
    for r in rows:
        out.append(r)
        node_cols = (r.get("Node_Output_Columns") or "").split(",")
        node_cols = [c.strip() for c in node_cols if c.strip()]
        tgt = r.get("Target_Column")
        src = r.get("Source_Column")
        if src and src != NULL_TOKEN:
            continue
        if tgt and tgt != NULL_TOKEN and tgt in node_cols and default_source_obj:
            synth = dict(r)
            synth["Transformation_Type"] = "Synth_Mapping"
            synth["Source_Object"] = default_source_obj
            synth["Source_Column"] = tgt
            key = (default_source_obj, tgt, synth.get("Target_Column"))
            if key not in seen:
                seen.add(key)
                out.append(synth)
    return out

def _get_text_by_path(node, tag_path: str) -> str:
    """local-name() XPath for 'a/b' or 'a/b@attr'."""
    if not tag_path:
        return ""
    path, attr = (tag_path.split("@", 1) + [None])[:2]
    segments = [s.strip() for s in path.split("/") if s.strip()]
    xp = "."
    for seg in segments:
        xp += f"/*[local-name()='{seg}']"
    try:
        nodes = node.xpath(xp)
    except Exception:
        nodes = []
    if not nodes:
        return ""
    n = nodes[0]
    if attr:
        return (n.attrib.get(attr, "") or "").strip()
    return (n.text or "").strip()

def _detect_profile_for_root(root):
    """Return 'CV' for Calculation:scenario, 'ANV' for Cube:cube, 'ATV' for ColumnView, 'DIM' for Dimension:dimension; else None."""
    local = etree.QName(root.tag).localname
    if local in {"scenario"}:     return "CV"
    if local in {"cube"}:         return "ANV"
    if local in {"ColumnView"}:   return "ATV"
    if local in {"dimension"}:    return "DIM"   # NEW: Dimension support
    return None

def _augment_log(msg):
    try:
        log.info(f"[AUGMENT] {msg}")
    except Exception:
        pass

def _parse_tree_safely(path):
    try:
        parser = etree.XMLParser(recover=True, huge_tree=True)
        tree = etree.parse(path, parser)
        return tree, tree.getroot()
    except Exception:
        return None, None

def _extract_semantic_labels(root) -> Dict[str, Dict[str, str]]:
    """
    Extract semantic labels (defaultDescription) from CV/ANV output attributes.
    Returns: {node_id -> {column_id -> label_text}}
    
    Structure (CV in logicalModel):
    <logicalModel>
      <attributes>
        <attribute id="COLUMN_ID">
          <descriptions defaultDescription="SEMANTIC LABEL"/>
        </attribute>
      </attributes>
    </logicalModel>
    """
    labels_by_node = {}
    try:
        # Find all attributes containers (can be in logicalModel, output, or root)
        for attrs_cont in root.xpath(".//attributes"):
            # iterate over direct children and pick known element names (attribute, measure)
            for attr_elem in list(attrs_cont):
                try:
                    ln = etree.QName(attr_elem.tag).localname
                except Exception:
                    ln = (attr_elem.tag or "").split("}")[-1]
                if ln not in ("attribute", "measure", "calculatedViewAttribute"):
                    continue
                # Try common id/name attributes
                attr_id = (attr_elem.attrib.get("id") or attr_elem.attrib.get("name") or "").strip()
                if not attr_id:
                    continue
                # Find descriptions element (prefer local-name match)
                descs_list = attr_elem.xpath(".//descriptions") or attr_elem.xpath(".//*[local-name()='descriptions']")
                if descs_list:
                    desc_elem = descs_list[0]
                    label = (desc_elem.attrib.get("defaultDescription") or "").strip()
                    if label:
                        # Use generic node key
                        node_key = "__OUTPUT__"
                        if node_key not in labels_by_node:
                            labels_by_node[node_key] = {}
                        labels_by_node[node_key][attr_id] = label
    except Exception as e:
        if DEBUG_PRINTS:
            log.debug(f"[LABELS] Error extracting semantic labels: {e}")
    return labels_by_node

def _make_base_row_for_augment(current_node, path, level=0, null_token=NULL_TOKEN):
    return {
        "Current_Node": current_node or null_token,
        "Hierarchy_Level": level,
        # ✅ NEW: explicit lineage separation
        "Lineage_Level": "COLUMN",      # COLUMN or NODE
        "Target_Node": "",
        "Source_Node": "",
        "SourceFile": path,
        "Parent_Input_Object": null_token,
        "Source_Object": null_token,
        "Source_Column": null_token,
        "Target_Column": null_token,
        "Label": null_token,                # ✅ NEW: semantic label from defaultDescription
        "Node_Output_Columns": null_token,
        "Transformation_Type": null_token,
        "Formula": null_token,
        "Filter_Expression": null_token,
        "Join_Name": null_token,
        "Join_Expression": null_token,
        "Join_Left_Object": null_token,
        "Join_Left_Column": null_token,
        "Join_Right_Object": null_token,
        "Join_Right_Column": null_token,
        "XPath": null_token
    }

# ======= PATCH / CANONICAL HELPERS =======

def _resolve_current_node_fallback_token(struct, default, root):
    """
    If parent_input_fallback == "__CURRENT_NODE__", return root's node id (per profile).
    Else return the configured fallback or default.
    """
    fb = (struct or {}).get("parent_input_fallback", default)
    if fb != "__CURRENT_NODE__":
        return fb
    try:
        rlocal = etree.QName(root.tag).localname
    except Exception:
        rlocal = ""
    if rlocal in {"scenario"}:     # CV
        return (root.attrib.get("id") or "").strip() or default
    if rlocal in {"cube"}:         # ANV
        return (root.attrib.get("id") or "").strip() or default
    if rlocal in {"ColumnView"}:   # ATV
        return (root.attrib.get("name") or "").strip() or default
    if rlocal in {"dimension"}:    # DIM
        return (root.attrib.get("id") or "").strip() or default
    return default

def _normalize_repo_uri_to_viewname(uri: str) -> str:
    """
    "/LAV_DWH_PKG.DEV_PKG.DEV_SAZ/attributeviews/ATV_CUSTOMER" -> "ATV_CUSTOMER"
    "/pkg/analyticviews/ANV_SOMETHING" -> "ANV_SOMETHING"
    If last token missing, return the original URI.
    """
    if not uri:
        return ""
    tok = uri.strip().split("/")[-1]
    return tok or uri.strip()

def _build_view_outputs_registry(all_rows):
    """
    Returns:
      outputs_by_view: {Current_Node -> set of output column names}
      explicit_by_view: {(view, target) -> True} for targets with explicit mapping/source
    """
    outputs_by_view = {}
    explicit_by_view = {}
    for r in all_rows:
        v = (r.get("Current_Node") or "").strip()
        t = (r.get("Target_Column") or "").strip()
        s = (r.get("Source_Column") or "").strip()
        if v and t and t != NULL_TOKEN:
            outputs_by_view.setdefault(v, set()).add(t)
        if v and t and s and s != NULL_TOKEN:
            explicit_by_view[(v, t)] = True
    return outputs_by_view, explicit_by_view
def expand_node_lineage_to_columns(rows):
    """
    Expand NODE-level lineage into COLUMN-level lineage
    using Node_Output_Columns.
    """
    outputs_by_node = {}

    # Collect output columns per node
    for r in rows:
        node = (r.get("Current_Node") or "").strip()
        cols = (r.get("Node_Output_Columns") or "").strip()
        if node and cols:
            outputs_by_node[node] = [
                c.strip() for c in cols.split(",") if c.strip()
            ]

    expanded = []

    for r in rows:
        if r.get("Lineage_Level") != "NODE":
            expanded.append(r)
            continue

        tgt_node = r.get("Target_Node")
        src_node = r.get("Source_Node")

        for col in outputs_by_node.get(tgt_node, []):
            e = dict(r)
            e["Lineage_Level"] = "COLUMN"
            e["Target_Column"] = col
            e["Source_Column"] = col
            e["Transformation_Type"] = "Parent_Rollup_Expanded"
            expanded.append(e)
    return expanded
def expand_node_target_column(rows):
    """
    If Target_Column == Current_Node, expand it
    into per-column lineage using Node_Output_Columns.
    """
    expanded = []

    for r in rows:
        tgt = (r.get("Target_Column") or "").strip()
        cur = (r.get("Current_Node") or "").strip()
        cols = (r.get("Node_Output_Columns") or "").strip()

        # ✅ Normal column row → keep as-is
        if not tgt or tgt != cur or not cols:
            expanded.append(r)
            continue

        # ✅ Expand node-scoped row into column-scoped rows
        for c in [x.strip() for x in cols.split(",") if x.strip()]:
            rr = dict(r)
            rr["Target_Column"] = c
            expanded.append(rr)

    return expanded

def _anv_emit_shared_dimensions(rows, path, root, null_token=NULL_TOKEN):
    """
    Append rows describing semantic bridges from <sharedDimensions>/<logicalJoin>.

    FIXED VERSION:
    - Parent_Input_Object is NORMALIZED (ATV name, not URI)
    - Ensures HeavyEngine can recurse
    """

    extra = []

    try:
        logical_joins = root.xpath(
            ".//*[local-name()='sharedDimensions']/*[local-name()='logicalJoin']"
        )
    except Exception:
        logical_joins = []

    cube_name = (root.attrib.get("id") or "").strip() or null_token

    for lj in logical_joins:
        # Raw URI from XML
        uri = (lj.attrib.get("associatedObjectUri") or "").strip()

        # ✅ NORMALIZE to actual ATV / view name
        src_view = _normalize_repo_uri_to_viewname(uri) or null_token

        # ✅ THIS IS THE CRITICAL FIX
        parent_object = src_view or null_token

        # Join metadata for traceability
        try:
            props = (lj.xpath("./*[local-name()='properties']") or [None])[0]
        except Exception:
            props = None

        join_type = (props.attrib.get("joinType") or "").strip() if props is not None else ""
        join_op   = (props.attrib.get("joinOperator") or "").strip() if props is not None else ""
        card      = (props.attrib.get("cardinality") or "").strip() if props is not None else ""

        meta = ", ".join(
            p for p in [
                f"joinType={join_type}" if join_type else "",
                f"joinOperator={join_op}" if join_op else "",
                f"cardinality={card}" if card else "",
                f"associatedObjectUri={uri}" if uri else "",
            ]
            if p
        ) or null_token

        # --------------------------------------------------
        # (A) attributeRef (#LOCAL) -> associatedAttributeNames
        # --------------------------------------------------
        try:
            attr_refs = lj.xpath("./*[local-name()='attributes']/*[local-name()='attributeRef']")
            assoc_names = lj.xpath("./*[local-name()='associatedAttributeNames']/*[local-name()='attributeName']")
        except Exception:
            attr_refs, assoc_names = [], []

        for i in range(max(len(attr_refs), len(assoc_names))):
            tgt = ((attr_refs[i].text or "").strip().lstrip("#")) if i < len(attr_refs) else ""
            src = ((assoc_names[i].text or "").strip()) if i < len(assoc_names) else ""
            if not tgt or not src:
                continue

            row = _make_base_row_for_augment(cube_name, path, level=1, null_token=null_token)

            # ✅ FIXED HERE
            row["Parent_Input_Object"] = parent_object
            row["Target_Column"] = tgt
            row["Source_Object"] = src_view
            row["Source_Column"] = src
            row["Transformation_Type"] = "Shared_Dimension"
            row["Formula"] = meta

            try:
                tree_obj = etree.ElementTree(root)
                row["XPath"] = tree_obj.getpath(lj)
            except Exception:
                pass

            extra.append(row)

        # --------------------------------------------------
        # (B) associatedAttributeFeatures (alias -> attribute)
        # --------------------------------------------------
        try:
            feats = lj.xpath(
                "./*[local-name()='associatedAttributeFeatures']"
                "/*[local-name()='attributeReference']"
            )
        except Exception:
            feats = []

        for f in feats:
            alias = (f.attrib.get("alias") or "").strip()
            name  = (f.attrib.get("attributeName") or "").strip()
            if not alias or not name:
                continue

            row = _make_base_row_for_augment(cube_name, path, level=1, null_token=null_token)

            # ✅ FIXED HERE
            row["Parent_Input_Object"] = parent_object
            row["Target_Column"] = alias
            row["Source_Object"] = src_view
            row["Source_Column"] = name
            row["Transformation_Type"] = "Shared_Dimension"
            row["Formula"] = meta

            try:
                tree_obj = etree.ElementTree(root)
                row["XPath"] = tree_obj.getpath(f)
            except Exception:
                pass

            extra.append(row)

    rows.extend(extra)
    return rows

def _anv_emit_exception_aggregation(rows, path, root, null_token=NULL_TOKEN):
    """
    Emit one row per <attribute> child of <exceptionAggregation> on calculated measures.
    Each row links the measure (Target_Column) to its exception-aggregation dependency column
    (Source_Column) in the referenced dimension view (Source_Object).
    Transformation_Type = "Exception_Agg_Attribute".
    """
    extra = []
    try:
        cube_name = (root.attrib.get("id") or "").strip() or null_token
    except Exception:
        cube_name = null_token

    try:
        measures = root.xpath(
            ".//*[local-name()='calculatedMeasures']/*[local-name()='measure']"
        )
    except Exception:
        measures = []

    for meas in measures:
        measure_id = (meas.attrib.get("id") or "").strip()
        if not measure_id:
            continue

        try:
            exc_agg_nodes = meas.xpath("./*[local-name()='exceptionAggregation']")
        except Exception:
            exc_agg_nodes = []

        for exc_agg in exc_agg_nodes:
            agg_type = (exc_agg.attrib.get("exceptionAggregationType") or "").strip()
            formula_val = f"exceptionAggregationType={agg_type}" if agg_type else null_token

            try:
                attr_nodes = exc_agg.xpath("./*[local-name()='attribute']")
            except Exception:
                attr_nodes = []

            for attr in attr_nodes:
                attr_name = (attr.attrib.get("attributeName") or "").strip()
                dim_uri   = (attr.attrib.get("dimensionUri") or "").strip()
                if not attr_name:
                    continue

                src_view = _normalize_repo_uri_to_viewname(dim_uri) if dim_uri else null_token

                row = _make_base_row_for_augment(cube_name, path, level=1, null_token=null_token)
                row["Parent_Input_Object"] = cube_name
                row["Target_Column"]       = measure_id
                row["Source_Object"]       = src_view
                row["Source_Column"]       = attr_name
                row["Transformation_Type"] = "Exception_Agg_Attribute"
                row["Formula"]             = formula_val

                try:
                    tree_obj = etree.ElementTree(root)
                    row["XPath"] = tree_obj.getpath(attr)
                except Exception:
                    pass

                extra.append(row)

    rows.extend(extra)
    return rows

def _global_anv_shared_dimension_pullthrough(all_rows):
    """
    For each Shared_Dimension row (ANV -> AssociatedView),
    if a column name C exists in both ANV outputs and AssociatedView outputs but ANV has no explicit source for C,
    append 'Shared_Dimension_Pullthrough' edge: AssociatedView.C -> ANV.C
    """
    out = list(all_rows)
    outputs_by_view, explicit_by_view = _build_view_outputs_registry(all_rows)

    links = set()
    for r in all_rows:
        if r.get("Transformation_Type") in ("Shared_Dimension",):
            anv = (r.get("Current_Node") or "").strip()
            assoc = (r.get("Source_Object") or "").strip()
            if anv and assoc:
                links.add((anv, assoc))

    for (anv, assoc) in sorted(links):
        anv_cols = outputs_by_view.get(anv, set())
        assoc_cols = outputs_by_view.get(assoc, set())
        if not anv_cols or not assoc_cols:
            continue

        for col in sorted(anv_cols.intersection(assoc_cols)):
            if not explicit_by_view.get((anv, col), False):
                row = {
                    "Current_Node": anv,
                    "Hierarchy_Level": 1,
                    "SourceFile": "<GLOBAL_PULLTHROUGH>",
                    "Parent_Input_Object": assoc,
                    "Source_Object": assoc,
                    "Source_Column": col,
                    "Target_Column": col,
                    "Node_Output_Columns": col,
                    "Transformation_Type": "Shared_Dimension_Pullthrough",
                    "Formula": "pullthrough=true",
                    "Filter_Expression": NULL_TOKEN,
                    "Join_Name": NULL_TOKEN,
                    "Join_Expression": NULL_TOKEN,
                    "Join_Left_Object": NULL_TOKEN,
                    "Join_Left_Column": NULL_TOKEN,
                    "Join_Right_Object": NULL_TOKEN,
                    "Join_Right_Column": NULL_TOKEN,
                    "XPath": NULL_TOKEN
                }
                out.append(row)
    return out

# ---------------- NEW: parse <input><entity> into schema.table ----
def _normalize_entity_object(raw: str) -> str:
    """
    Examples:
      '#//\"DWH_LAV_DWH\".SCDWLK_PLANT_T' -> DWH_LAV_DWH.SCDWLK_PLANT_T
      '#//MYSCHEMA.MYTABLE'               -> MYSCHEMA.MYTABLE
      '\"SCHEMA\".\"TABLE\"'              -> SCHEMA.TABLE
    """
    if not raw:
        return ""
    s = raw.strip()
    if s.startswith("#//"):
        s = s[3:].strip()
    parts = s.split(".")
    if len(parts) >= 2:
        schema = parts[-2].strip().strip('"')
        table  = parts[-1].strip().strip('"')
        if schema and table:
            return f"{schema}.{table}"
    return s.strip('"')

def _source_object_from_input_entity(input_elem) -> str:
    """Find <entity> under <input> (namespace-agnostic) and return schema.table"""
    ent_node = None
    try:
        ent_node = (input_elem.xpath("./*[local-name()='entity']") or [None])[0]
    except Exception:
        ent_node = None
    if ent_node is None:
        ent_node = input_elem.find("entity")
    raw = (ent_node.text or "").strip() if ent_node is not None and ent_node.text else ""
    norm = _normalize_entity_object(raw)
    return norm if norm else ""

def _cache_input_context(entity):
    """
    ✅ OPTIMIZED: Cache nearest <input> element search with single traversal.
    
    ELIMINATES:
    - 3x redundant XML tree traversals (before in lines 1103-1110, 1186-1194, 1265-1273)
    - Multiple calls to _source_object_from_input_entity()
    - Code duplication across 3 sections
    
    Returns:
        tuple: (nearest_input_elem or None, source_object_string)
    
    GUARANTEES:
    - Never returns invalid references
    - Returns (None, NULL_TOKEN) on any failure
    - O(n) tree traversal done once per entity
    - Idempotent (repeated calls safe)
    """
    try:
        p = entity
        nearest_input = None
        
        # Single upward traversal
        while p is not None and nearest_input is None:
            try:
                cand = p.xpath("./*[local-name()='input']")
                if cand:
                    nearest_input = cand[0]
                    break
            except Exception:
                pass
            
            p = p.getparent() if hasattr(p, "getparent") else None
        
        # Not found
        if nearest_input is None:
            return (None, NULL_TOKEN)
        
        # Extract source object once
        src_obj = _source_object_from_input_entity(nearest_input)
        return (nearest_input, _ensure_notnull(src_obj))
        
    except Exception as e:
        log.debug(f"[CACHE_INPUT] Failed: {e}")
        return (None, NULL_TOKEN)

def _resolve_source_object_with_fallback(entity, central_table, parent_val, current_node):
    """
    Resolve Source_Object using a clear fallback precedence.

    Precedence:
      1. nearest <input><entity>
      2. central_table
      3. parent_val

    Returns a non-empty string or NULL_TOKEN.
    """
    nearest_input, src_obj = _cache_input_context(entity)
    if _is_null(src_obj):
        src_obj = _ensure_notnull(central_table)
    if _is_null(src_obj):
        src_obj = (_ensure_notnull(parent_val)
                   if parent_val and parent_val != NULL_TOKEN
                   else NULL_TOKEN)

    if _is_null(src_obj):
        log.debug(f"[SOURCE_OBJ] No Source_Object resolved for {current_node}; using NULL_TOKEN")
    else:
        log.debug(f"[SOURCE_OBJ] Resolved for {current_node}: {src_obj}")

    return src_obj
# ---------------- File Discovery ----------------


def discover_xml_files(path, recursive=False):
    r"""
    Discover .xml/.txt under 'path'. If INPUT.respect_numbered_order is True,
    sort by a numeric prefix extracted with INPUT.file_prefix_regex (default: r'^(\d+)').
    Falls back to name sort when the regex doesn't match.
    """
    allowed_ext = (".xml", ".txt")

    # Collect files
    if os.path.isfile(path):
        return [path] if path.lower().endswith(allowed_ext) else []
    if os.path.isdir(path):
        pattern = "**/*" if recursive else "*"
        files = glob.glob(os.path.join(path, pattern), recursive=recursive)
        files = [f for f in files if f.lower().endswith(allowed_ext)]
    else:
        return []

    # Optional: numbered ordering
    try:
        respect = bool(getattr(config, "INPUT", {}).get("respect_numbered_order", False))
        rx = getattr(config, "INPUT", {}).get("file_prefix_regex", r"^(\d+)")
        num_re = re.compile(rx)
    except Exception:
        respect, num_re = False, re.compile(r"^(\d+)")

    if not respect:
        return sorted(files)

    def key_fn(p):
        b = os.path.basename(p)
        m = num_re.match(b)
        if m:
            try:
                return (0, int(m.group(1)), b.lower())
            except Exception:
                return (0, float('inf'), b.lower())
        return (1, float('inf'), b.lower())

    return sorted(files, key=key_fn)

# --------------------------------------------------------------------
# Core Extractor
# --------------------------------------------------------------------
def extract_lineage_from_xml(path):
    """
    Returns:
        results: list[dict]
        lineage_graph: dict[target -> set(source)]
    """
    # -------------- SAFE PARSING --------------
    try:
        parser = etree.XMLParser(recover=True, huge_tree=True)
        tree = etree.parse(path, parser)
        root = tree.getroot()
        tree_obj = etree.ElementTree(root)
    except Exception as e:
        log.exception(f"Failed to parse XML file {path}: {e}")
        return [], {}

    # -------------- PROFILE SELECTION --------------
    root_local = etree.QName(root.tag).localname

    def _tag_eq_local(config_tag: str, actual_local: str) -> bool:
        return bool(config_tag) and actual_local == config_tag.split(":")[-1]

    if root_local in {"calculationView", "scenario"} and hasattr(config, "XML_STRUCTURE_CV"):
        STRUCT = getattr(config, "XML_STRUCTURE_CV")
    elif (_tag_eq_local("Cube:cube", root_local) or root_local == "cube") and hasattr(config, "XML_STRUCTURE_ANV"):
        STRUCT = getattr(config, "XML_STRUCTURE_ANV")
    elif root_local == "dimension" and hasattr(config, "XML_STRUCTURE_DIM"):
        STRUCT = getattr(config, "XML_STRUCTURE_DIM")
    elif root_local == "ColumnView" and hasattr(config, "XML_STRUCTURE_ATV"):
        STRUCT = getattr(config, "XML_STRUCTURE_ATV")
    else:
        STRUCT = getattr(config, "XML_STRUCTURE", {}) or \
                 getattr(config, "XML_STRUCTURE_ANV", {}) or \
                 getattr(config, "XML_STRUCTURE_CV",  {}) or \
                 getattr(config, "XML_STRUCTURE_ATV", {}) or \
                 getattr(config, "XML_STRUCTURE_DIM", {})

    results = []
    lineage_graph = {}
    node_stack = []
    parent_stack = []
    node_targets = {}
    central_table = _find_central_table(root)

    # -------------- HELPERS --------------
    def _get_text_local(node, tag_path):
        if not tag_path:
            return ""
        parts = tag_path.split("@", 1)
        p = parts[0]
        attr = parts[1] if len(parts) > 1 else None
        xp = "."
        for seg in p.split("/"):
            seg = seg.strip()
            if seg:
                xp += f"/*[local-name()='{seg}']"
        try:
            found = node.xpath(xp)
            n = found[0] if found else None
        except Exception:
            n = None
        if n is None:
            return ""
        return (n.attrib.get(attr, "") if attr else (n.text or "").strip())

    def get_safe_node():
        if node_stack:
            return node_stack[-1]
        fb = _resolve_current_node_fallback_token(STRUCT, NULL_TOKEN, root)
        return fb or NULL_TOKEN

    def base_row(current_node):
        return {
            "Current_Node": current_node,
            "Hierarchy_Level": len(node_stack),
            "SourceFile": path,
            "Parent_Input_Object": NULL_TOKEN,
            "Source_Object": NULL_TOKEN,
            "Source_Column": NULL_TOKEN,
            "Target_Column": NULL_TOKEN,
            "Label": NULL_TOKEN,                # ✅ NEW: semantic label from defaultDescription
            "Node_Output_Columns": NULL_TOKEN,
            "Transformation_Type": NULL_TOKEN,
            "Formula": NULL_TOKEN,
            "Filter_Expression": NULL_TOKEN,
            "Join_Name": NULL_TOKEN,
            "Join_Expression": NULL_TOKEN,
            "Join_Left_Object": NULL_TOKEN,
            "Join_Left_Column": NULL_TOKEN,
            "Join_Right_Object": NULL_TOKEN,
            "Join_Right_Column": NULL_TOKEN,
            "XPath": NULL_TOKEN
        }

    # Cache to avoid recomputing/joining the same node outputs many times
    _node_outputs_cache = {}
    def _node_outputs_repr(safe_node, cols):
        """
        Memory-safe representation of Node_Output_Columns:
          - By default, store ONLY the count (constant space).
          - If you need the list, cap it by columns and characters and cache per node.
        """
        mode = getattr(config, "PERF", {}).get("node_outputs_emit", "count")  # "none"|"count"|"list"
        if not cols:
            return NULL_TOKEN
        if mode == "none":
            return NULL_TOKEN
        if mode == "count":
            return str(len(cols))

        # mode == "list" (capped)
        if safe_node in _node_outputs_cache:
            return _node_outputs_cache[safe_node]

        try:
            max_cols  = int(getattr(config, "PERF", {}).get("node_outputs_max_cols", 2000))
            max_chars = int(getattr(config, "PERF", {}).get("node_outputs_max_chars", 200000))
        except Exception:
            max_cols, max_chars = 2000, 200000

        # Cap the number of columns to limit memory and speed
        # Avoid materializing a full sorted list when set is huge
        capped = sorted(cols)[:max_cols] if max_cols else sorted(cols)
        s = ",".join(capped)

        # Truncate overly long strings (as an additional guard)
        if max_chars and len(s) > max_chars:
            s = s[:max_chars] + "...[truncated]"

        # If we truncated by column count, append a tail note
        if max_cols and len(cols) > max_cols:
            s += f",...(+{len(cols) - max_cols} more)"

        _node_outputs_cache[safe_node] = s
        return s


    def add_row(row):
        safe_node = get_safe_node()
        cols = node_targets.get(safe_node)
        # memory-safe, configurable representation
        try:
            row["Node_Output_Columns"] = _node_outputs_repr(safe_node, cols)
        except MemoryError:
            # absolute fallback — never fail the engine on this column
            row["Node_Output_Columns"] = str(len(cols or [])) if cols else NULL_TOKEN
        results.append(row)

    def register_dependency(target, source):
        t = _to_str(target)
        s = _to_str(source)
        if not t or t == NULL_TOKEN or not s or s == NULL_TOKEN:
            return
        lineage_graph.setdefault(t, set()).add(s)

    def _get_text_by_path_localname(node, tag_path):
        if not tag_path:
            return ""
        if "@" in tag_path:
            path, attr = tag_path.split("@", 1)
        else:
            path, attr = tag_path, None
        xp = "."
        for seg in path.split("/"):
            seg = seg.strip()
            if seg:
                xp += f"/*[local-name()='{seg}']"
        try:
            hits = node.xpath(xp)
            if not hits:
                return ""
            n = hits[0]
        except Exception:
            return ""
        return (n.attrib.get(attr, "") if attr else (n.text or "").strip())

    def fq_name(schema, name, catalog=None):
        parts = [p for p in [catalog, schema, name] if p and p != NULL_TOKEN]
        return ".".join(parts) if parts else (name or NULL_TOKEN)
        
    def _smart_choose_format(n_rows, prefer="parquet", max_excel_rows=1_000_000):
        """
        Decide final output format based on size and preference.
        Returns one of: "xlsx", "csv", "parquet".
        """
        if n_rows is None:
            return prefer or "parquet"
        # If more than Excel cap, avoid Excel
        if n_rows > max_excel_rows:
            return "csv" if prefer == "csv" else "parquet"
        # Small datasets -> follow preference (default parquet)
        return prefer or "parquet"
    # heavy_engine.py → helper
    def _select_targets_namespaced(df):
        return sorted(
            (f"{str(r['Current_Node']).strip()}::{str(r['Target_Column']).strip()}")
            for _, r in df[["Current_Node","Target_Column"]].dropna().iterrows()
            if str(r["Current_Node"]).strip() and str(r["Target_Column"]).strip()
        )

    # -------------- MAIN TRAVERSAL --------------
    def traverse(entity):
        tag = etree.QName(entity.tag).localname
        left_obj = None
        right_obj = None
        # ---------- NODE CONTEXT ----------
        node_conf = STRUCT.get("current_node", {}) or {}
        is_node = (tag == node_conf.get("tag", ""))

        if is_node:
            node_name = get_attr_or_text(entity, node_conf.get("node_id_attr"))
            node_stack.append(node_name)
            node_targets.setdefault(node_name, set())

            # CV: pre-seed visible columns
            for va in entity.findall("viewAttributes/viewAttribute"):
                vid = (va.attrib.get("id") or "").strip()
                if vid:
                    node_targets[node_name].add(vid)

            # ATV: pre-seed <element name="...">
            if node_conf.get("tag") == "ColumnView":
                for el in entity.xpath(".//*[local-name()='element']"):
                    en = (el.attrib.get("name") or "").strip()
                    if en:
                        node_targets[node_name].add(en)

        # Logical model (CV) as node context
        is_logical_model = (etree.QName(root.tag).localname == "scenario" and tag == "logicalModel")
        if is_logical_model:
            lm_id = (entity.attrib.get("id") or "").strip()
            if lm_id:
                node_stack.append(lm_id)
                node_targets.setdefault(lm_id, set())

        current_node = get_safe_node()

        # ---------- PARENT INPUT OBJECT ----------
        parent_conf = STRUCT.get("parent_input_object", {}) or {}
        is_parent = (tag == parent_conf.get("tag", ""))
        if is_parent:
            parent_val = get_attr_or_text(entity, parent_conf.get("reference_attr"))
            parent_stack.append(parent_val)

        parent_val = parent_stack[-1] if parent_stack else current_node

        # ============================================================== 
        # 1) SOURCE-TARGET MAPPINGS
        # ==============================================================
        st_multi = [s for s in STRUCT.get("source_target_multi", []) if isinstance(s, dict)]
        if st_multi:
            parent_tags = {s.get("parent_tag") for s in st_multi if s.get("parent_tag")}
            if tag in parent_tags:
                # NEW: capture <input><entity> once per input parent
                input_entity_object = ""
                if tag == "input":
                    input_entity_object = _source_object_from_input_entity(entity)

                for spec in st_multi:
                    if spec.get("parent_tag") != tag:
                        continue
                    for child in entity:
                        ctag = etree.QName(child.tag).localname
                        if ctag != spec.get("mapping_tag"):
                            continue

                        row = base_row(current_node)
                        row["Parent_Input_Object"] = parent_val

                        row["Target_Column"] = get_attr_or_text(child, spec.get("target_attr"))

                        src_col   = _get_by_path(child, spec.get("source_attr_path"),  default=NULL_TOKEN)
                        src_schem = _get_by_path(child, spec.get("source_schema_path"), default=NULL_TOKEN)
                        src_table = _get_by_path(child, spec.get("source_table_path"),  default=NULL_TOKEN)
                        row["Source_Column"] = src_col

                        # A) If schema/table available via config paths
                        if src_schem != NULL_TOKEN and src_table != NULL_TOKEN:
                            row["Source_Object"] = f"{src_schem}.{src_table}"

                        # B) Otherwise, inherit from <input><entity> (ATV/CV)
                        if (not row["Source_Object"] or row["Source_Object"] == NULL_TOKEN) and input_entity_object:
                            row["Source_Object"] = input_entity_object

                        row["Transformation_Type"] = spec.get("transformation_type", "Mapping")
                        # Classify Rename when we have a simple rename without expression
                        tgt_n = (row["Target_Column"] or "").strip()
                        src_n = (row["Source_Column"] or "").strip()
                        if row["Transformation_Type"] == "Mapping" and tgt_n and src_n and tgt_n != src_n:
                            f = (row.get("Formula") or "").strip()
                            if not f or f == NULL_TOKEN:
                                row["Transformation_Type"] = "Rename"

                        # formula override
                        fpath = spec.get("formula_attr_path")
                        if fpath:
                            if fpath.startswith("@"):
                                a = fpath[1:]
                                row["Formula"] = child.attrib.get(a, "") or NULL_TOKEN
                            else:
                                row["Formula"] = _get_by_path(child, fpath, default=NULL_TOKEN)

                        try:
                            row["XPath"] = tree_obj.getpath(child)
                        except Exception:
                            pass

                        safe_node = get_safe_node()
                        node_targets.setdefault(safe_node, set()).add(row["Target_Column"])
                        register_dependency(row["Target_Column"], row["Source_Column"])
                        add_row(row)
    # ==============================================================
    # 1b) NODE-SCOPED MAPPINGS (Projection / Aggregation / Union)
    # ============================================================== 
        nm_specs = [s for s in STRUCT.get("node_mapping_multi", []) if isinstance(s, dict)]
        if nm_specs:
            for spec in nm_specs:
                ptag = spec.get("parent_tag")
                if not ptag or tag != ptag:
                    continue
                # only when we're inside specific node(s)
                if not _has_ancestor_local_in(entity, spec.get("within_node_tags", [])):
                    continue

                # support multiple mapping tag names under the current parent
                mtag_list = spec.get("mapping_tags", ["mapping"])
                mapping_children = []
                for mtag in mtag_list:
                    try:
                        mapping_children.extend(entity.xpath(f"./*[local-name()='{mtag}']"))
                    except Exception:
                        pass

                if not mapping_children:
                    continue

                t_alts = spec.get("target_attr_alts", ["target"])
                s_alts = spec.get("source_attr_alts", ["source"])
                xform  = spec.get("transformation_type", "Node_Mapping")

                for m in mapping_children:
                    tgt = _get_first_attr(m, t_alts, default=NULL_TOKEN)
                    src = _get_first_attr(m, s_alts, default=NULL_TOKEN)
                    if not tgt or tgt == NULL_TOKEN:
                        continue

                    row = base_row(current_node)
                    row["Parent_Input_Object"] = parent_val
                    row["Target_Column"] = tgt
                    row["Source_Column"] = src
                    row["Transformation_Type"] = xform
                     # Try to inherit a concrete Source_Object (nearest <input><entity>), else keep blank
                    # try:
                    #      # Search upwards for an <input> that has <entity> text (CV/ATV)
                    #     p = entity
                    #     nearest_input = None
                    #     while p is not None and nearest_input is None:
                    #         cand = p.xpath("./*[local-name()='input']")
                    #         if cand:
                    #             nearest_input = cand[0]
                    #             break
                    #         p = p.getparent() if hasattr(p, "getparent") else None
                    #     if nearest_input is not None:
                    #         inh = _source_object_from_input_entity(nearest_input)
                    #         if inh:
                    #             row.setdefault("Source_Object", inh)
                    # except Exception:               
                    #     pass
                    src_obj = _resolve_source_object_with_fallback(
                        entity, central_table, parent_val, current_node
                    )
                    if src_obj and src_obj != NULL_TOKEN:
                        row.setdefault("Source_Object", src_obj)

                    # inherit Source_Object from <input><entity> if available (you already parse it earlier)
                    # your earlier logic sets it during st_multi for parent_tag='input'
                    # here we add nothing extra; if your augmenters filled Source_Object, de-dup will keep best row.

                    try:
                        row["XPath"] = tree_obj.getpath(m)
                    except Exception:
                        pass

                    # register dependency when we have a source
                    if src and src != NULL_TOKEN:
                        register_dependency(row["Target_Column"], row["Source_Column"])
                    else:
                        # >>> OFFLINE CATALOG FALLBACK <<<
                        cur = get_safe_node()
                        if cur and "." in cur:
                            schema, name = cur.split(".", 1)
                            for c in get_runtime_columns(ocmod, schema, name):
                                register_dependency(row["Target_Column"], c)

                    # Add visible target to node outputs and store row
                    safe_node = get_safe_node()
                    node_targets.setdefault(safe_node, set()).add(row["Target_Column"])
                    add_row(row)

                # -------- Optional pass-through synthesis inside this node ----------
                if STRUCT.get("emit_passthrough_when_unmapped", True):
                    # Compute which targets have explicit source within THIS node context
                    explicit = set()
                    for m in mapping_children:
                        tv = _get_first_attr(m, t_alts, default="")
                        sv = _get_first_attr(m, s_alts, default="")
                        if tv and sv:
                            explicit.add(tv)

                    # Guard: Check if pass-through is safe (single input or explicitly allowed)
                    limit_to_single = bool(STRUCT.get("limit_passthrough_to_single_input", True))
                    parent = entity.getparent() if hasattr(entity, "getparent") else None
                    
                    single_input_ok = True  # default allow when the guard is disabled
                    if limit_to_single:
                        single_input_ok = False
                        if parent is not None:
                            try:
                                inputs = parent.xpath("./*[local-name()='input']")
                            except Exception:
                                inputs = parent.findall("input") or []
                            single_input_ok = (len(inputs) == 1)
                    
                    # ✅ FIX: Guard properly applied - code moved INSIDE the condition
                    if (not limit_to_single) or single_input_ok:
                        # Try to add pass-through rows for visible outputs with no explicit mapping
                        safe_node = get_safe_node()
                        visible = node_targets.get(safe_node, set()) or set()
                        
                        for tgt in sorted(visible):
                            if not tgt or tgt in explicit or tgt == NULL_TOKEN:
                                continue
                            
                            pr = base_row(current_node)
                            pr["Parent_Input_Object"] = parent_val
                            pr["Target_Column"] = tgt
                            pr["Source_Column"] = tgt  # pass-through
                            pr["Transformation_Type"] = "PassThrough"
                            
                            # Attach Source_Object using the standardized fallback chain
                            src_obj = _resolve_source_object_with_fallback(
                                entity, central_table, parent_val, current_node
                            )
                            if src_obj and src_obj != NULL_TOKEN:
                                pr["Source_Object"] = src_obj
                            
                            try:
                                pr["XPath"] = tree_obj.getpath(entity)
                            except Exception:
                                pass
                            
                            register_dependency(tgt, tgt)
                            add_row(pr)
                    else:
                        # Guard activated: log why pass-throughs not created
                        log.debug(f"[PASSTHROUGH] Skipped for node {current_node}: "
                                f"limit_to_single=True but multiple inputs found")

        # ============================================================== 
        # 2) CALCULATED MEASURES & ATTRIBUTES
        # ==============================================================
        calc_specs = [s for s in STRUCT.get("calculated_multi", []) if isinstance(s, dict)]
        if calc_specs:
            # (A) parent/mapping style (calculatedMeasures/measure)
            for spec in calc_specs:
                ptag = spec.get("parent_tag")
                mtag = spec.get("mapping_tag")
                if ptag and mtag and tag == ptag:
                    items = entity.xpath(f"./*[local-name()='{mtag}']")
                    for meas in items:
                        row = base_row(current_node)
                        row["Parent_Input_Object"] = parent_val

                        target = get_attr_or_text(meas, spec.get("target_attr"))
                        row["Target_Column"] = target or NULL_TOKEN

                        expr_path = spec.get("expression_tag_path")
                        expr = _get_by_path(meas, expr_path, text_if_no_attr=True, default="")
                        if not expr:
                            expr = _get_text_local(meas, expr_path)
                        row["Formula"] = expr.lstrip(" :") if expr else NULL_TOKEN

                        dpath = spec.get("description_attr_path")
                        if dpath:
                            row["Description"] = _get_text_local(meas, dpath) or NULL_TOKEN

                        for k in (spec.get("meta_attrs") or []):
                            row[k] = meas.attrib.get(k, "") or NULL_TOKEN

                        row["Transformation_Type"] = spec.get("transformation_type", "Calculated_Column")
                        # try:
                        #     p = meas   # or `entity` for section B
                        #     nearest_input = None
                        #     while p is not None and nearest_input is None:
                        #         cand = p.xpath("./*[local-name()='input']")
                        #         if cand:
                        #             nearest_input = cand[0]
                        #             break
                        #         p = p.getparent() if hasattr(p, "getparent") else None
                        #     if nearest_input is not None:
                        #         inh = _source_object_from_input_entity(nearest_input)
                        #         if inh:
                        #             row["Source_Object"] = inh
                        # except Exception:
                        #     pass
                        # if not row.get("Source_Object") or row["Source_Object"] == NULL_TOKEN:
                        #     if central_table:
                        #         row["Source_Object"] = central_table
                        # try:
                        #     row["XPath"] = tree_obj.getpath(meas)
                        # except Exception:
                        #     pass
                        
                        # In Calculated Measures section (Line ~1265):
                        src_obj = _resolve_source_object_with_fallback(
                            meas, central_table, parent_val, current_node
                        )
                        if src_obj and src_obj != NULL_TOKEN:
                            row["Source_Object"] = src_obj
                        
                        # for src in extract_columns_from_expression(row["Formula"]):
                        #     register_dependency(row["Target_Column"], src)
                        #     # >>> OFFLINE CATALOG FALLBACK FOR CALCULATED COLUMN <<<
                        #     if not extract_columns_from_expression(row["Formula"]):
                        #         # fallback to runtime columns of current node
                        #         cur = get_safe_node()
                        #         if cur and "." in cur:
                        #             schema, name = cur.split(".", 1)
                        #             for c in get_runtime_columns(ocmod, schema, name):
                        #                 register_dependency(row["Target_Column"], c)
                        extracted_srcs = extract_columns_from_expression(row["Formula"])

                        # Register all extracted sources
                        for src in extracted_srcs:
                            register_dependency(row["Target_Column"], src)

                        # Fallback only if extraction yielded nothing
                        if not extracted_srcs:  # ✅ Use cached result (no re-extraction)
                            # fallback to runtime columns of current node
                            cur = get_safe_node()
                            if cur and "." in cur:
                                try:
                                    schema, name = cur.split(".", 1)
                                    for c in get_runtime_columns(ocmod, schema, name):
                                        register_dependency(row["Target_Column"], c)
                                except Exception as e:
                                    log.debug(f"[FORMULA_FALLBACK] Failed for {cur}: {e}")

                        safe_node = get_safe_node()
                        node_targets.setdefault(safe_node, set()).add(row["Target_Column"])
                        add_row(row)

            # (B) tag-based calculated nodes (calculatedAttribute)
            for spec in calc_specs:
                tag_name = spec.get("tag")
                if tag_name and "/" not in tag_name and tag == tag_name:
                    row = base_row(current_node)
                    row["Parent_Input_Object"] = parent_val
                    row["Target_Column"] = get_attr_or_text(entity, spec.get("target_attr")) or NULL_TOKEN

                    expr_path = spec.get("expression_tag_path")
                    expr = _get_by_path(entity, expr_path, text_if_no_attr=True, default="")
                    if not expr:
                        expr = _get_text_local(entity, expr_path)
                    row["Formula"] = expr.lstrip(" :") if expr else NULL_TOKEN

                    row["Transformation_Type"] = spec.get("transformation_type", "Calculated_Column")

                    dpath = spec.get("description_attr_path")
                    if dpath:
                        row["Description"] = _get_text_local(entity, dpath)

                    for k in (spec.get("meta_attrs") or []):
                        row[k] = entity.attrib.get(k, "") or NULL_TOKEN

                    try:
                        row["XPath"] = tree_obj.getpath(entity)
                    except Exception:
                        pass

                    for src in extract_columns_from_expression(row["Formula"]):
                        register_dependency(row["Target_Column"], src)

                    safe_node = get_safe_node()
                    node_targets.setdefault(safe_node, set()).add(row["Target_Column"])
                    add_row(row)

        # ==============================================================
        # 3) FILTERS  (CV) — robust attribute-based capture + child <valueFilter> fallback
        # ==============================================================
        filter_specs = [s for s in STRUCT.get("filter_multi", []) if isinstance(s, dict)]
        if filter_specs:

            def _attr_dict_from_filter(elem, names):
                """Collect requested attributes and normalized xsi:type; never throw."""
                out = {}
                if elem is None:
                    return out
                # capture requested names
                if names:
                    for n in names:
                        v = (elem.attrib.get(n) or "").strip()
                        if v:
                            out[n] = v
                # normalize xsi:type (namespaced attr)
                xns = "{http://www.w3.org/2001/XMLSchema-instance}type"
                xval = (elem.attrib.get(xns) or elem.attrib.get("xsi:type") or "").strip()
                if xval and "xsi:type" not in out:
                    out["xsi:type"] = xval
                # if caller didn't specify names, grab ALL attributes as a last resort
                if not names:
                    for k, v in elem.attrib.items():
                        v = (v or "").strip()
                        if v:
                            out[k] = v
                return out

            def _parent_attribute_id(elem, path_to_parent_id):
                """Resolve '../@id' from <filter> to its owning <viewAttribute id='...'>."""
                if elem is None or not path_to_parent_id:
                    return ""
                cur = elem
                left, attr = (path_to_parent_id.split("@", 1) + [None])[:2]
                left = (left or "").strip()
                ups = 0
                while left.startswith("../"):
                    ups += 1
                    left = left[3:]
                try:
                    for _ in range(ups):
                        cur = cur.getparent() if hasattr(cur, "getparent") else None
                        if cur is None:
                            return ""
                    return (cur.attrib.get(attr or "") or "").strip()
                except Exception:
                    return ""

            def _compose_filter_predicate(attr_id, adict, composer):
                """SQL-like composition; falls back to a compact attribute bundle."""
                if not attr_id:
                    return ""
                value = adict.get("value", "")
                operator = (adict.get("operator") or "").upper()
                including = (adict.get("including") or "").lower()

                # Composer map first (if provided by config.py)
                if composer and isinstance(composer, dict):
                    opmap = composer.get("operator_map", {}) or {}
                    if operator and operator in opmap and value != "":
                        return opmap[operator].format(attr=attr_id, value=value)
                    if composer.get("fallback_to_including", True) and value != "":
                        if including in ("", "true", "1", "yes"):
                            return composer.get("if_including_true", "{attr} = '{value}'").format(attr=attr_id, value=value)
                        return composer.get("if_including_false", "{attr} <> '{value}'").format(attr=attr_id, value=value)

                # No composer → basic including/value if present
                if value != "" and including != "":
                    return f"{attr_id} = '{value}'" if including in ("", "true", "1", "yes") else f"{attr_id} <> '{value}'"

                return ""  # will bundle later if still empty

            for spec in filter_specs:
                if tag != spec.get("tag"):
                    continue

                attr_names = spec.get("condition_expression_attr")   # e.g. ["value","including","operator","xsi:type"]
                parent_attr_id_path = spec.get("attribute_id_path")  # e.g. "../@id"
                composer = spec.get("compose_predicate")             # optional

                # DEBUG: use existing global flag in your file
                if DEBUG_PRINTS:
                    print("[DEBUG] <filter> encountered")

                # --- read attributes from <filter ...> ---
                adict = _attr_dict_from_filter(entity, attr_names) if not spec.get("condition_tag") else {}
                attr_id = _parent_attribute_id(entity, parent_attr_id_path) if parent_attr_id_path else ""

                # --- optional: read a child <valueFilter ...> if adict didn't bring enough info ---
                if (not adict) or (not adict.get("value") and not adict.get("operator") and not adict.get("including")):
                    try:
                        vchild = None
                        for ch in entity:
                            if hasattr(ch, "tag"):
                                ln = ch.tag.split("}")[-1]
                                if ln == "valueFilter":
                                    vchild = ch; break
                        if vchild is not None:
                            # bring including/value/operator/xsi:type from child
                            xns = "{http://www.w3.org/2001/XMLSchema-instance}type"
                            cad = {}
                            for k in ("value", "including", "operator"):
                                vv = (vchild.attrib.get(k) or "").strip()
                                if vv:
                                    cad[k] = vv
                            xtype = (vchild.attrib.get(xns) or vchild.attrib.get("xsi:type") or "").strip()
                            if xtype:
                                cad["xsi:type"] = xtype
                            for k, v in cad.items():
                                adict.setdefault(k, v)
                    except Exception:
                        pass

                # --- if no parent attr_id (e.g., Privilege:AttributeFilter outside viewAttribute), use @attributeName ---
                if not attr_id:
                    alt_attr = (entity.attrib.get("attributeName") or "").strip()  # Privilege:AttributeFilter case
                    if alt_attr:
                        attr_id = alt_attr

                # --- compose final expression or bundle as a hard fallback ---
                final_expr = _compose_filter_predicate(attr_id, adict, composer)
                if not final_expr or final_expr.strip() == "":
                    if attr_id and adict:
                        bundle = "; ".join([f"{k}={adict[k]}" if k != "value" else f"value='{adict[k]}'" for k in adict])
                        final_expr = f"{attr_id} :: {bundle}"
                    elif adict:
                        final_expr = "; ".join([f"{k}={adict[k]}" if k != "value" else f"value='{adict[k]}'" for k in adict])
                    else:
                        inner = "".join(entity.itertext()).strip()
                        final_expr = inner if inner else "<FILTER_PRESENT_NO_ATTRS>"

                # bind to the actual parent attribute when available (no spraying)
                safe_node = get_safe_node()
                targets = {attr_id} if attr_id else (node_targets.get(safe_node, {NULL_TOKEN}) or {NULL_TOKEN})
                for tgt in targets:
                    row = base_row(current_node)
                    row["Parent_Input_Object"] = parent_val
                    row["Target_Column"] = tgt
                    row["Filter_Expression"] = final_expr
                    row["Transformation_Type"] = "Filter"
                    try:
                        row["XPath"] = tree_obj.getpath(entity)
                    except Exception:
                        pass
                    add_row(row)
                    # >>> OFFLINE CATALOG FALLBACK FOR JOIN OUTPUTS <<<
                    if tgt and tgt != NULL_TOKEN:
                        if left_obj and "." in left_obj:
                            ls, ln = left_obj.split(".", 1)
                            for c in get_runtime_columns(ocmod, ls, ln):
                                register_dependency(tgt, c)
                        if right_obj and "." in right_obj:
                            rs, rn = right_obj.split(".", 1)
                            for c in get_runtime_columns(ocmod, rs, rn):
                                register_dependency(tgt, c)
                if DEBUG_PRINTS:
                    print(f"[DEBUG] Filter attr={attr_id or '<node-wide>'} expr={final_expr}")

        # ==============================================================
        # 4) JOINS (enhanced CV JoinView logic with populated-side analysis)
        # ==============================================================

        join_conf = STRUCT.get("join", {}) if isinstance(STRUCT.get("join", {}), dict) else {}
        is_joinview = (
            tag == join_conf.get("tag", "") and 
            "JoinView" in (entity.get("{http://www.w3.org/2001/XMLSchema-instance}type", "") or "")
        )

        if is_joinview:
            # ----------------------------------------------------------
            # Basic metadata
            # ----------------------------------------------------------
            join_type = (entity.attrib.get("joinType") or "").strip()
            join_order = (entity.attrib.get("joinOrder") or "").strip()
            join_card  = (entity.attrib.get("cardinality") or "").strip()

            # Always initialize to avoid undefined-variable exceptions
            expression = ""

            # ----------------------------------------------------------
            # 1) Join condition (if present)
            # ----------------------------------------------------------
            cond_tag  = join_conf.get("condition_tag")
            cond_attr = join_conf.get("condition_expression_attr", "joinType")

            if cond_tag:
                cond_node = entity.find(cond_tag)
                expression = get_attr_or_text(cond_node, cond_attr, use_text_if_no_attr=True) or ""
            else:
                expression = (entity.get(cond_attr) or "").strip()

            # ----------------------------------------------------------
            # 2) Resolve the <input> nodes
            # ----------------------------------------------------------
            inputs = entity.findall("input")
            left_input  = inputs[0] if len(inputs) >= 1 else None
            right_input = inputs[1] if len(inputs) >= 2 else None

            left_obj  = (left_input.get("node")  if left_input  is not None else NULL_TOKEN) or NULL_TOKEN
            right_obj = (right_input.get("node") if right_input is not None else NULL_TOKEN) or NULL_TOKEN

            # ----------------------------------------------------------
            # 3) Build target->source maps
            # ----------------------------------------------------------
            def _build_target_to_source_map(input_node):
                md = {}
                if input_node is None:
                    return md
                for m in input_node.findall("mapping"):
                    tgt = (m.get("target") or "").strip()
                    src = (m.get("source") or "").strip()
                    if tgt:
                        md[tgt] = src
                return md

            left_map  = _build_target_to_source_map(left_input)
            right_map = _build_target_to_source_map(right_input)

            # ----------------------------------------------------------
            # 4) Declared joinAttribute names
            # ----------------------------------------------------------
            join_attr_tag = join_conf.get("join_attr_tag", "joinAttribute")
            join_names = [(j.get("name") or "").strip() for j in entity.findall(join_attr_tag)]

            # ----------------------------------------------------------
            # 5) Resolve join attributes for each side
            # ----------------------------------------------------------
            def _upstream_has_attr(node_ref, attr):
                if not node_ref or not attr:
                    return False
                ref_id = str(node_ref).lstrip("#")
                try:
                    cand = root.xpath(f".//*[local-name()='calculationView' and @id='{ref_id}']")
                except Exception:
                    cand = []
                for cv in cand:
                    for va in cv.findall("viewAttributes/viewAttribute"):
                        if (va.attrib.get("id") or "").strip() == attr:
                            return True
                return False

            def _resolve_side(attr_name, mapping_dict, node_obj):
                if attr_name in mapping_dict:
                    return mapping_dict[attr_name]
                if _upstream_has_attr(node_obj, attr_name):
                    return attr_name
                return attr_name

            left_cols  = [_resolve_side(n, left_map, left_obj)  for n in join_names]
            right_cols = [_resolve_side(n, right_map, right_obj) for n in join_names]

            # ----------------------------------------------------------
            # 6) Determine visible output targets
            # ----------------------------------------------------------
            safe_node = get_safe_node()
            targets = node_targets.get(safe_node, {NULL_TOKEN}) or {NULL_TOKEN}

            # If node outputs weren't seeded yet, fallback to viewAttributes
            if targets == {NULL_TOKEN} or not targets:
                va_ids = [
                    va.get("id") for va in entity.findall("viewAttributes/viewAttribute")
                    if va.get("id")
                ]
                if va_ids:
                    targets = set(va_ids)

            # ----------------------------------------------------------
            # 7) Build join pairs for expression (ALWAYS DEFINE)
            # ----------------------------------------------------------
            pairs_for_expr = []

            if left_cols and right_cols:
                for lc, rc in zip_longest(left_cols, right_cols, fillvalue=NULL_TOKEN):
                    if (
                        left_obj and lc and lc != NULL_TOKEN and
                        right_obj and rc and rc != NULL_TOKEN
                    ):
                        pairs_for_expr.append(f"{left_obj}.{lc} = {right_obj}.{rc}")

            # Safe default for expression if missing
            expr_base = expression or ""

            # Final join expression (safe)
            join_expr_full = "; ".join(pairs_for_expr) if pairs_for_expr else expr_base
            join_expr_full = join_expr_full if join_expr_full else NULL_TOKEN

            # ----------------------------------------------------------
            # 8) Emit Join rows once per target column
            # ----------------------------------------------------------
            for tgt in targets:
                row = base_row(current_node)
                row["Parent_Input_Object"] = parent_val
                row["Target_Column"]       = tgt
                row["Join_Expression"]     = join_expr_full
                row["Join_Left_Object"]    = left_obj or NULL_TOKEN
                row["Join_Left_Column"]    = NULL_TOKEN
                row["Join_Right_Object"]   = right_obj or NULL_TOKEN
                row["Join_Right_Column"]   = NULL_TOKEN
                row["Transformation_Type"] = "Join"
                row["Join_Node_Id"]        = safe_node
                row["Join_Type"]           = join_type
                row["Join_Operator"]       = ""
                row["Join_Cardinality"]    = (join_card or join_order)

                try:
                    row["XPath"] = tree_obj.getpath(entity)
                except Exception:
                    row["XPath"] = NULL_TOKEN

                add_row(row)

            # ----------------------------------------------------------
            # 9) Join summary row (populated-side analysis)
            # ----------------------------------------------------------
            left_fields  = sorted([t for t, src in left_map.items()  if _wt_estimate(src, t) >= _wt_min() and t])
            right_fields = sorted([t for t, src in right_map.items() if _wt_estimate(src, t) >= _wt_min() and t])

            if left_fields and right_fields:
                pop_side = "BOTH"
            elif right_fields:
                pop_side = "RIGHT"
            elif left_fields:
                pop_side = "LEFT"
            else:
                pop_side = "NONE"

            summary_pairs = [
                f"{left_obj}.{lc} = {right_obj}.{rc}"
                for lc, rc in zip_longest(left_cols, right_cols, fillvalue=NULL_TOKEN)
                if lc and rc and lc != NULL_TOKEN and rc != NULL_TOKEN
            ]

            row = base_row(current_node)
            row["Parent_Input_Object"]          = safe_node
            row["Transformation_Type"]          = "Join"
            row["Join_Node_Id"]                 = safe_node
            row["Join_Type"]                    = join_type
            row["Join_Operator"]                = ""
            row["Join_Cardinality"]             = (join_card or join_order)
            row["Join_Left_Object"]             = left_obj
            row["Join_Right_Object"]            = right_obj
            row["Join_Expression"]              = "; ".join(summary_pairs) if summary_pairs else NULL_TOKEN
            row["Join_Populated_Side"]          = pop_side
            row["Join_Populated_Fields"]        = "; ".join(sorted(set(left_fields) | set(right_fields)))
            row["Join_Populated_Fields_Left"]   = "; ".join(left_fields)
            row["Join_Populated_Fields_Right"]  = "; ".join(right_fields)

            try:
                row["XPath"] = tree_obj.getpath(entity)
            except Exception:
                pass

            add_row(row)

        # 5) RESTRICTED MEASURES (ANV) — attributeName + child <valueFilter>
        # ==============================================================
        rm_conf = STRUCT.get("restricted_measures")
        if isinstance(rm_conf, dict) and entity is root:

            parent_tag  = rm_conf["parent_tag"]
            measure_tag = rm_conf["measure_tag"]

            measures = root.xpath(
                f".//*[local-name()='{parent_tag}']/*[local-name()='{measure_tag}']"
            )

            # Optional config keys (safe defaults if not present)
            val_filter_tag      = rm_conf.get("value_filter_tag", "valueFilter")
            value_attr          = rm_conf.get("value_attr", "value")
            filter_tag          = rm_conf.get("filter_tag", "filter")
            filter_attr         = rm_conf.get("filter_attr", "attributeName")
            value_operator_attr = rm_conf.get("value_operator_attr", "operator")
            value_including_attr= rm_conf.get("value_including_attr", "including")
            operands_tag        = rm_conf.get("operands_tag", "operands")
            operand_value_attr  = rm_conf.get("operand_value_attr", "value")

            for mnode in measures:
                mid = mnode.attrib.get(rm_conf["id_attr"], "").strip()
                base_m = (mnode.attrib.get(rm_conf["base_measure_attr"], "") or "").strip()
                source_col = base_m if base_m else NULL_TOKEN
                source_obj = source_col
                agg = (mnode.attrib.get(rm_conf["engine_agg_attr"], "SUM") or "SUM").upper()
                formula = f"{agg}({source_col})" if source_col != NULL_TOKEN else NULL_TOKEN
                desc = _get_text_by_path_localname(mnode, rm_conf["description_attr_path"]) or NULL_TOKEN

                # ---- Build Filter_Expression from <filter attributeName=...><valueFilter .../></filter> ----
                restriction = mnode.find(rm_conf["restriction_tag"])
                filter_expr = NULL_TOKEN
                if restriction is not None:
                    fnode = restriction.find(filter_tag)  # <filter ... attributeName="X">
                    if fnode is not None:
                        attr_name = (fnode.attrib.get(filter_attr, "") or "").strip()

                        vnode = fnode.find(val_filter_tag)  # <valueFilter ... including/op/value/operands>
                        if vnode is not None and attr_name:
                            # Read attributes safely
                            op   = ((vnode.attrib.get(value_operator_attr) or "").strip().upper())
                            incl = ((vnode.attrib.get(value_including_attr) or "").strip().lower())
                            val  =  (vnode.attrib.get(value_attr) or "")

                            # (Optional) IN/NOT IN via <operands value="..."/>
                            ops = []
                            try:
                                for o in vnode.findall(operands_tag):
                                    ov = (o.attrib.get(operand_value_attr) or "").strip()
                                    if ov != "":
                                        ops.append(ov)
                            except Exception:
                                ops = []

                            # Compose predicate
                            if ops:
                                kw = "IN" if incl in ("", "true", "1", "yes") else "NOT IN"
                                quoted = ", ".join(f"'{x}'" for x in ops)
                                filter_expr = f"{attr_name} {kw} ({quoted})" if quoted else NULL_TOKEN
                            elif op == "NL":
                                # Not-Like; treat empty value as IS NOT NULL (practical HANA rendering)
                                filter_expr = (f"{attr_name} NOT LIKE '{val}'") if val != "" else f"{attr_name} IS NOT NULL"
                            elif val != "":
                                filter_expr = f"{attr_name} = '{val}'" if incl in ("", "true", "1", "yes") else f"{attr_name} <> '{val}'"
                            else:
                                filter_expr = NULL_TOKEN

                # ---- Emit row (unchanged shape/logic) ----
                row = base_row(current_node)
                row["Parent_Input_Object"] = parent_val
                row["Transformation_Type"] = "Aggregate"

                row["Target_Column"] = mid or NULL_TOKEN
                row["Source_Column"] = source_col
                row["Source_Object"] = source_obj
                row["Formula"] = formula
                row["Filter_Expression"] = filter_expr
                row["Description"] = desc

                try:
                    row["XPath"] = tree_obj.getpath(mnode)
                except Exception:
                    pass

                safe_node = get_safe_node()
                node_targets.setdefault(safe_node, set()).add(mid)
                if source_col != NULL_TOKEN:
                    register_dependency(mid, source_col)
                add_row(row)

        # ============================================================== 
        # 6) SOURCE TABLES
        # ============================================================== 
        st_conf = STRUCT.get("source_table")
        if isinstance(st_conf, dict) and tag == st_conf.get("tag", ""):
            raw_name   = get_attr_or_text(entity, st_conf.get("table_name_attr"))
            raw_schema = get_attr_or_text(entity, st_conf.get("schema_attr"))
            raw_cata   = get_attr_or_text(entity, st_conf.get("catalog_attr"))

            full = fq_name(
                raw_schema if raw_schema != NULL_TOKEN else None,
                raw_name   if raw_name   != NULL_TOKEN else None,
                raw_cata   if raw_cata   != NULL_TOKEN else None
            )

            qualify = bool(st_conf.get("qualify_columns", False))
            safe_node = get_safe_node()

            row = base_row(current_node)
            row["Parent_Input_Object"] = parent_val
            row["Source_Object"] = full
            row["Source_Column"] = "*" if not qualify else f"{full}.*"
            row["Target_Column"] = "*"
            row["Transformation_Type"] = "Source_Table"

            try:
                row["XPath"] = tree_obj.getpath(entity)
            except Exception:
                pass

            node_targets.setdefault(safe_node, set()).add("*")
            register_dependency("*", row["Source_Column"])
            add_row(row)

        # -------- Recurse --------
        for child in entity:
            traverse(child)

        if is_node:
            node_stack.pop()
        if is_parent and parent_stack:
            parent_stack.pop()

    # ---------------- EXECUTE ----------------
    traverse(root)

    # ---------------- SYNTHETIC RESOLUTION ----------------
    results = synthesize_from_node_outputs(results, default_source_obj=central_table)

    # ---------------- GUARANTEED RETURN ----------------
    return results, lineage_graph, tree, root

# ---------------- CV helpers ----------------
def _augment_scenario_metadata(rows, path, root, attrs):
    """Attach Scenario_* columns (CV only)."""
    local = etree.QName(root.tag).localname
    if local != "scenario":
        return rows
    kv = {}
    for a in attrs:
        kv[f"Scenario_{a}"] = (root.attrib.get(a, "") or "").strip() or NULL_TOKEN
    if not kv:
        return rows
    for r in rows:
        if r.get("SourceFile") == path:
            r.update({k: v for k, v in kv.items() if k not in r})
    return rows

def _cv_build_datasource_registry(root):
    """
    Build {id: {"type":..., "resourceUri":..., "schema":..., "table":..., "object":...}}
    for <dataSources><DataSource ...>.
    """
    reg = {}
    try:
        ds_nodes = root.xpath(".//*[local-name()='dataSources']/*[local-name()='DataSource']")
    except Exception:
        ds_nodes = []
    for ds in ds_nodes:
        did  = (ds.attrib.get("id") or "").strip()
        dtyp = (ds.attrib.get("type") or "").strip()
        res_node = (ds.xpath("./*[local-name()='resourceUri']") or [None])[0]
        res_uri  = (res_node.text or "").strip() if res_node is not None else ""
        col_node = (ds.xpath("./*[local-name()='columnObject']") or [None])[0]
        schema = (col_node.attrib.get("schemaName") or "").strip() if col_node is not None else ""
        table  = (col_node.attrib.get("columnObjectName") or "").strip() if col_node is not None else ""
        obj = f"{schema}.{table}" if schema and table else ""
        if did:
            reg[did] = {"type": dtyp, "resourceUri": res_uri, "schema": schema, "table": table, "object": obj}
    return reg

def _augment_datasource_metadata(rows, path, root):
    """Attach DataSource_* columns to rows whose Parent_Input_Object starts with '#<DataSourceId>'."""
    reg = _cv_build_datasource_registry(root)
    if not reg:
        return rows
    for r in rows:
        if r.get("SourceFile") != path:
            continue
        p = (r.get("Parent_Input_Object") or "").strip()
        if p.startswith("#"):
            did = p.lstrip("#")
            if did in reg:
                info = reg[did]
                r.setdefault("DataSource_Id", did)
                if info.get("type"):
                    r.setdefault("DataSource_Type", info["type"])
                if info.get("resourceUri"):
                    r.setdefault("DataSource_ResourceUri", info["resourceUri"])
                if info.get("object"):
                    r.setdefault("DataSource_Object", info["object"])
    return rows

def _find_ancestor_calcview_id(node):
    """Walk up to nearest <calculationView ... id='...'> and return id."""
    cur = node
    while cur is not None:
        try:
            if etree.QName(cur.tag).localname == "calculationView":
                cid = (cur.attrib.get("id") or "").strip()
                if cid:
                    return cid
        except Exception:
            pass
        cur = cur.getparent() if hasattr(cur, "getparent") else None
    return ""

def _augment_constant_mappings(rows, path, root, null_token=NULL_TOKEN):
    """Append rows for <mapping xsi:type='...ConstantAttributeMapping' ...> (CV only)."""
    extra = []
    try:
        maps = root.xpath(".//*[local-name()='mapping']")
    except Exception:
        maps = []
    for m in maps:
        xtype = m.attrib.get("{http://www.w3.org/2001/XMLSchema-instance}type", "")
        if "ConstantAttributeMapping" not in xtype:
            continue
        tgt = (m.attrib.get("target") or "").strip()
        val = (m.attrib.get("value") or "").strip()
        nll = (m.attrib.get("null") or "").strip()
        inp = m.getparent()
        while inp is not None and etree.QName(inp.tag).localname != "input":
            inp = inp.getparent()
        parent_node_ref = (inp.attrib.get("node") or "").strip() if inp is not None else ""

        current_node = _find_ancestor_calcview_id(m) or null_token
        row = _make_base_row_for_augment(current_node, path, level=0, null_token=null_token)
        row["Parent_Input_Object"] = parent_node_ref or null_token
        row["Target_Column"] = tgt or null_token
        row["Source_Column"] = "<CONST>"
        parts = []
        if val: parts.append(f"value='{val}'")
        if nll: parts.append(f"null={nll}")
        row["Formula"] = ", ".join(parts) if parts else null_token
        row["Transformation_Type"] = "Constant_Mapping"
        try:
            tree_obj = etree.ElementTree(root)
            row["XPath"] = tree_obj.getpath(m)
        except Exception:
            pass
        extra.append(row)
    rows.extend(extra)
    return rows

def _augment_logical_bindings(rows, path, root, null_token=NULL_TOKEN):
    """Append rows for <logicalModel>/<attributes>/<attribute>/<keyMapping> (CV only)."""
    extra = []
    try:
        attrs = root.xpath(".//*[local-name()='logicalModel']/*[local-name()='attributes']/*[local-name()='attribute']")
    except Exception:
        attrs = []
    for a in attrs:
        aid = (a.attrib.get("id") or "").strip()
        km = (a.xpath("./*[local-name()='keyMapping']") or [None])[0]
        if km is None:
            continue
        obj = (km.attrib.get("columnObjectName") or "").strip()
        col = (km.attrib.get("columnName") or "").strip()
        current_node = ""
        row = _make_base_row_for_augment(current_node or null_token, path, level=0, null_token=null_token)
        row["Target_Column"] = aid or null_token
        row["Source_Object"] = obj or null_token
        row["Source_Column"] = col or null_token
        row["Transformation_Type"] = "Logical_Binding"
        try:
            tree_obj = etree.ElementTree(root)
            row["XPath"] = tree_obj.getpath(a)
        except Exception:
            pass
        extra.append(row)
    rows.extend(extra)
    return rows

# ---------------- ATV helpers ----------------
def _normalize_entity_text(raw):
    """'#//\"SCHEMA\".TABLE' -> SCHEMA.TABLE; '\"A\".\"B\"' -> A.B"""
    if not raw:
        return ""
    s = raw.strip()
    if s.startswith("#//"):
        s = s[3:].strip()
    parts = s.split(".")
    if len(parts) >= 2:
        schema = parts[-2].strip().strip('"')
        table  = parts[-1].strip().strip('"')
        if schema and table:
            return f"{schema}.{table}"
    return s.strip('"')

def _atv_collect_entity_object(input_node):
    ent = (input_node.xpath("./*[local-name()='entity']") or [None])[0]
    if ent is None or not ent.text:
        return ""
    return _normalize_entity_text(ent.text)

def _augment_atv_entity_star(rows, path, root, null_token=NULL_TOKEN):
    """Append Source_Table rows from ATV <input><entity> anchors."""
    extra = []
    try:
        inputs = root.xpath(".//*[local-name()='viewNode']/*[local-name()='input']")
    except Exception:
        inputs = []
    cv_name = (root.attrib.get("name") or "").strip() or null_token
    for inp in inputs:
        obj = _atv_collect_entity_object(inp)
        if not obj:
            continue
        row = _make_base_row_for_augment(cv_name, path, level=1, null_token=null_token)
        row["Parent_Input_Object"] = cv_name
        row["Source_Object"] = obj
        row["Source_Column"] = f"{obj}.*"
        row["Target_Column"] = "*"
        row["Transformation_Type"] = "Source_Table"
        try:
            tree_obj = etree.ElementTree(root)
            row["XPath"] = tree_obj.getpath(inp)
        except Exception:
            pass
        extra.append(row)
    rows.extend(extra)
    return rows

def _augment_atv_label_binding(rows, path, root, null_token=NULL_TOKEN):
    """Append Label_Binding rows: <element name='X'><labelElement>#//.../Y</labelElement></element>."""
    extra = []
    try:
        elements = root.xpath(".//*[local-name()='element']")
    except Exception:
        elements = []
    cv_name = (root.attrib.get("name") or "").strip() or null_token
    for el in elements:
        tgt = (el.attrib.get("name") or "").strip()
        if not tgt:
            continue
        labs = el.xpath("./*[local-name()='labelElement']")
        if len(labs) == 0:
            continue
        label_ref = (labs[0].text or "").strip()
        src = label_ref.split("/")[-1] if label_ref else ""
        row = _make_base_row_for_augment(cv_name, path, level=1, null_token=null_token)
        row["Parent_Input_Object"] = cv_name
        row["Target_Column"] = tgt or null_token
        row["Source_Column"] = src or null_token
        row["Transformation_Type"] = "Label_Binding"
        row["Formula"] = "labelElement"
        try:
            tree_obj = etree.ElementTree(root)
            row["XPath"] = tree_obj.getpath(el)
        except Exception:
            pass
        extra.append(row)
    rows.extend(extra)
    return rows

# ---------------- DIM helpers (NEW) ----------------
def _dim_current_node_name(root, null_token=NULL_TOKEN):
    """Use Dimension id as the Current_Node context."""
    return (root.attrib.get("id") or "").strip() or null_token

def _dim_table_label(schema, obj, alias):
    """Return preferred object label -> alias if present else schema.obj."""
    if alias:
        return alias
    if schema and obj:
        return f"{schema}.{obj}"
    return obj or schema or ""

def _augment_dimension_metadata(rows, path, root, attrs):
    """
    Attach Dimension_* columns (DIM only) to all rows from this SourceFile.
    attrs: list of root attribute names to read (e.g., ['id','dimensionType','applyPrivilegeType',...])
    """
    local = etree.QName(root.tag).localname
    if local != "dimension":
        return rows
    kv = {}
    for a in attrs:
        kv[f"Dimension_{a}"] = (root.attrib.get(a, "") or "").strip() or NULL_TOKEN
    if not kv:
        return rows
    for r in rows:
        if r.get("SourceFile") == path:
            for k, v in kv.items():
                r.setdefault(k, v)
    return rows

def _augment_dimension_textual_metadata(rows, path, root):
    """
    Optional: Attach textual metadata (defaultDescription, comments, metadata timestamps) as Dimension_* columns.
    """
    try:
        desc_node = (root.xpath("./*[local-name()='descriptions']") or [None])[0]
    except Exception:
        desc_node = None
    default_desc = ""
    comments = []
    if desc_node is not None:
        default_desc = (desc_node.attrib.get("defaultDescription") or "").strip()
        try:
            com_nodes = desc_node.xpath("./*[local-name()='comment']")
        except Exception:
            com_nodes = []
        for c in com_nodes:
            txt = (c.attrib.get("text") or "").strip()
            if txt:
                comments.append(_html_unescape(txt).replace("\r", " ").replace("\n", " "))
    # metadata timestamps
    try:
        meta_node = (root.xpath("./*[local-name()='metadata']") or [None])[0]
    except Exception:
        meta_node = None
    activated_at = (meta_node.attrib.get("activatedAt") or "").strip() if meta_node is not None else ""
    changed_at   = (meta_node.attrib.get("changedAt") or "").strip()   if meta_node is not None else ""

    if not (default_desc or comments or activated_at or changed_at):
        return rows

    for r in rows:
        if r.get("SourceFile") == path:
            if default_desc:
                r.setdefault("Dimension_defaultDescription", default_desc)
            if comments:
                r.setdefault("Dimension_Comments", " | ".join(comments))
            if activated_at:
                r.setdefault("Dimension_ActivatedAt", activated_at)
            if changed_at:
                r.setdefault("Dimension_ChangedAt", changed_at)
    return rows

def _dim_emit_source_tables(rows, path, root, null_token=NULL_TOKEN):
    """Append Source_Table rows from privateDataFoundation/tableProxies."""
    extra = []
    try:
        tables = root.xpath(
            ".//*[local-name()='privateDataFoundation']"
            "/*[local-name()='tableProxies']/*[local-name()='tableProxy']/*[local-name()='table']"
        )
    except Exception:
        tables = []
    cur = _dim_current_node_name(root, null_token)
    for t in tables:
        schema = (t.attrib.get("schemaName") or "").strip()
        name   = (t.attrib.get("columnObjectName") or "").strip()
        alias  = (t.attrib.get("alias") or "").strip()
        objlbl = _dim_table_label(schema, name, alias)
        if not objlbl:
            continue
        row = _make_base_row_for_augment(cur, path, level=1, null_token=null_token)
        row["Parent_Input_Object"] = cur
        row["Source_Object"] = objlbl or null_token
        row["Source_Column"] = f"{objlbl}.*"
        row["Target_Column"] = "*"
        row["Transformation_Type"] = "Source_Table"
        try:
            tree_obj = etree.ElementTree(root)
            row["XPath"] = tree_obj.getpath(t)
        except Exception:
            pass
        extra.append(row)
    rows.extend(extra)
    return rows

def _dim_emit_attribute_key_mappings(rows, path, root, null_token=NULL_TOKEN):
    """Append Logical_Binding rows from attributes/attribute/keyMapping. Optionally include attribute flags."""
    extra = []
    try:
        attrs = root.xpath(".//*[local-name()='attributes']/*[local-name()='attribute']")
    except Exception:
        attrs = []
    cur = _dim_current_node_name(root, null_token)
    add_flags = getattr(config, "AUGMENTATION", {}).get("dim_emit_attribute_flags", True)

    for a in attrs:
        aid = (a.attrib.get("id") or "").strip()
        km = (a.xpath("./*[local-name()='keyMapping']") or [None])[0]
        if not aid or km is None:
            continue
        schema = (km.attrib.get("schemaName") or "").strip()
        name   = (km.attrib.get("columnObjectName") or "").strip()
        alias  = (km.attrib.get("alias") or "").strip()
        col    = (km.attrib.get("columnName") or "").strip()
        objlbl = _dim_table_label(schema, name, alias)
        row = _make_base_row_for_augment(cur, path, level=1, null_token=null_token)
        row["Parent_Input_Object"] = cur
        row["Target_Column"] = aid or null_token
        row["Source_Object"] = objlbl or null_token
        row["Source_Column"] = col or null_token
        row["Transformation_Type"] = "Logical_Binding"

        if add_flags:
            key_flag    = (a.attrib.get("key") or "").strip()
            hidden_flag = (a.attrib.get("hidden") or "").strip()
            order_val   = (a.attrib.get("order") or "").strip()
            desc_col    = (a.attrib.get("descriptionColumnName") or "").strip()
            if key_flag:    row.setdefault("Attribute_Key", key_flag)
            if hidden_flag: row.setdefault("Attribute_Hidden", hidden_flag)
            if order_val:   row.setdefault("Attribute_Order", order_val)
            if desc_col:    row.setdefault("Attribute_Description_Column", desc_col)

        try:
            tree_obj = etree.ElementTree(root)
            row["XPath"] = tree_obj.getpath(a)
        except Exception:
            pass
        extra.append(row)
    rows.extend(extra)
    return rows

def _dim_emit_joins(rows, path, root, null_token=NULL_TOKEN):
    """Append Join rows from privateDataFoundation/joins/join (pairs by position)."""
    extra = []
    try:
        joins = root.xpath(".//*[local-name()='privateDataFoundation']/*[local-name()='joins']/*[local-name()='join']")
    except Exception:
        joins = []
    cur = _dim_current_node_name(root, null_token)
    for j in joins:
        lt = (j.xpath("./*[local-name()='leftTable']") or [None])[0]
        rt = (j.xpath("./*[local-name()='rightTable']") or [None])[0]
        if lt is None or rt is None:
            continue
        ls, ln, la = (lt.attrib.get("schemaName","").strip(),
                      lt.attrib.get("columnObjectName","").strip(),
                      lt.attrib.get("alias","").strip())
        rs, rn, ra = (rt.attrib.get("schemaName","").strip(),
                      rt.attrib.get("columnObjectName","").strip(),
                      rt.attrib.get("alias","").strip())
        lobj = _dim_table_label(ls, ln, la) or null_token
        robj = _dim_table_label(rs, rn, ra) or null_token

        prop = (j.xpath("./*[local-name()='properties']") or [None])[0]
        join_type = (prop.attrib.get("joinType") or "").strip() if prop is not None else ""
        join_op   = (prop.attrib.get("joinOperator") or "").strip() if prop is not None else ""
        card      = (prop.attrib.get("cardinality") or "").strip() if prop is not None else ""
        lang_col  = (j.attrib.get("languageColumn") or "").strip()

        left_cols  = [ (c.text or "").strip() for c in (j.xpath("./*[local-name()='leftColumns']/*[local-name()='columnName']") or []) ]
        right_cols = [ (c.text or "").strip() for c in (j.xpath("./*[local-name()='rightColumns']/*[local-name()='columnName']") or []) ]
        pair_count = min(len(left_cols), len(right_cols)) if (left_cols and right_cols) else 0
        pair_count = max(pair_count, 1) if (left_cols or right_cols) else 0

        for idx in range(pair_count):
            lc = left_cols[idx]  if idx < len(left_cols)  else "<MISSING>"
            rc = right_cols[idx] if idx < len(right_cols) else "<MISSING>"
            row = _make_base_row_for_augment(cur, path, level=1, null_token=null_token)
            row["Parent_Input_Object"] = cur
            row["Transformation_Type"] = "Join"
            row["Join_Left_Object"] = lobj
            row["Join_Left_Column"] = lc or null_token
            row["Join_Right_Object"] = robj
            row["Join_Right_Column"] = rc or null_token

            expr_parts = []
            if join_op:
                expr_parts.append(join_op)
            expr = f"{lobj}.{lc} = {robj}.{rc}" if (lc and rc and lc != "<MISSING>" and rc != "<MISSING>") else ""
            if expr:
                expr_parts.append(expr)
            row["Join_Expression"] = " | ".join(expr_parts) if expr_parts else null_token

            row["Join_Language_Column"] = lang_col if lang_col else null_token

            fparts = []
            if join_type: fparts.append(f"type={join_type}")
            if card:      fparts.append(f"cardinality={card}")
            if lang_col:  fparts.append(f"languageColumn={lang_col}")
            row["Formula"] = ", ".join(fparts) if fparts else null_token
            try:
                tree_obj = etree.ElementTree(root)
                row["XPath"] = tree_obj.getpath(j)
            except Exception:
                pass
            extra.append(row)
    rows.extend(extra)
    return rows

def _dim_emit_table_filters(rows, path, root, null_token=NULL_TOKEN):
    """Append Filter rows from tableProxy/columnFilter/valueFilter.
       Supports AccessControl:SingleValueFilter and ListValueFilter (operands),
       renders IN / NOT IN based on 'including' flag when operator=IN."""
    extra = []
    try:
        proxies = root.xpath(".//*[local-name()='privateDataFoundation']/*[local-name()='tableProxies']/*[local-name()='tableProxy']")
    except Exception:
        proxies = []
    cur = _dim_current_node_name(root, null_token)
    for px in proxies:
        t = (px.xpath("./*[local-name()='table']") or [None])[0]
        if t is None:
            continue
        schema = (t.attrib.get("schemaName") or "").strip()
        name   = (t.attrib.get("columnObjectName") or "").strip()
        alias  = (t.attrib.get("alias") or "").strip()
        objlbl = _dim_table_label(schema, name, alias) or null_token

        cfs = px.xpath("./*[local-name()='columnFilter']")
        for cf in cfs:
            col = (cf.attrib.get("columnName") or "").strip()
            vf  = (cf.xpath("./*[local-name()='valueFilter']") or [None])[0]
            parts = []
            filter_expr = f"{objlbl}.{col}" if col else null_token

            if vf is not None:
                xtype = vf.attrib.get("{http://www.w3.org/2001/XMLSchema-instance}type", "")
                incl  = (vf.attrib.get("including") or "").strip().lower()
                op    = (vf.attrib.get("operator") or "").strip().upper()
                val   = (vf.attrib.get("value") or "").strip()
                ops = [ (o.attrib.get("value") or "").strip()
                        for o in (vf.xpath("./*[local-name()='operands']") or []) ]
                if xtype: parts.append(f"xsi:type={xtype}")
                if op:    parts.append(f"operator={op}")
                if incl:  parts.append(f"including={incl}")
                if val:   parts.append(f"value='{val}'")
                if ops:
                    parts.append("operands=[" + ", ".join([f"'{x}'" for x in ops if x]) + "]")
                    if col and objlbl and op == "IN":
                        in_kw = "IN" if (incl in ("", "true")) else "NOT IN"
                        filter_expr = f"{objlbl}.{col} {in_kw} (" + ", ".join([f"'{x}'" for x in ops if x]) + ")"

            row = _make_base_row_for_augment(cur, path, level=1, null_token=null_token)
            row["Parent_Input_Object"] = cur
            row["Source_Object"] = objlbl
            row["Filter_Expression"] = filter_expr or null_token
            row["Transformation_Type"] = "Filter"
            row["Formula"] = ", ".join(parts) if parts else null_token
            try:
                tree_obj = etree.ElementTree(root)
                row["XPath"] = tree_obj.getpath(cf)
            except Exception:
                pass
            extra.append(row)
    rows.extend(extra)
    return rows

def _dim_emit_calculated_attributes(rows, path, root, null_token=NULL_TOKEN):
    """Append Calculated_Attribute rows from calculatedAttributes/.../formula."""
    extra = []
    try:
        cattrs = root.xpath(".//*[local-name()='calculatedAttributes']/*[local-name()='calculatedAttribute']")
    except Exception:
        cattrs = []
    cur = _dim_current_node_name(root, null_token)
    for ca in cattrs:
        aid = (ca.attrib.get("id") or "").strip()
        if not aid:
            continue
        kc = (ca.xpath("./*[local-name()='keyCalculation']") or [None])[0]
        fm = (kc.xpath("./*[local-name()='formula']") or [None])[0] if kc is not None else None
        expr_raw = (fm.text or "").strip() if fm is not None else ""
        expr = _html_unescape(expr_raw) if expr_raw else ""
        row = _make_base_row_for_augment(cur, path, level=1, null_token=null_token)
        row["Parent_Input_Object"] = cur
        row["Target_Column"] = aid or null_token
        row["Transformation_Type"] = "Calculated_Attribute"
        row["Formula"] = expr or null_token
        try:
            tree_obj = etree.ElementTree(root)
            row["XPath"] = tree_obj.getpath(ca)
        except Exception:
            pass
        extra.append(row)
    rows.extend(extra)
    return rows
    
def _has_ancestor_local_in(node, allowed_locals: list) -> bool:
    """Return True if the current node has any ancestor whose local-name() is in allowed_locals."""
    if not allowed_locals:
        return True
    cur = node.getparent() if hasattr(node, "getparent") else None
    allowed = set([s.strip() for s in allowed_locals if s and str(s).strip()])
    while cur is not None:
        try:
            if etree.QName(cur.tag).localname in allowed:
                return True
        except Exception:
            pass
        cur = cur.getparent() if hasattr(cur, "getparent") else None
    return False

def _get_first_attr(elem, names, default=NULL_TOKEN):
    """Return the first non-empty attribute among names."""
    if not elem or not names:
        return default
    for n in names:
        v = (elem.attrib.get(n) or "").strip()
        if v:
            return v
    return default

def _is_null(val):
    """
    ✅ STANDARDIZED: Check if value is null-like.
    
    Returns:
        True if val is None, empty string, or NULL_TOKEN
        False otherwise
    
    CANONICAL TRUTH: Only NULL_TOKEN represents "actual NULL" in lineage.
    """
    if val is None:
        return True
    if isinstance(val, str):
        s = val.strip()
        return s == "" or s == NULL_TOKEN
    return False

def _ensure_notnull(val, default=None):
    """
    ✅ STANDARDIZED: Guarantee output is never empty string or None.
    
    Args:
        val: Value to check
        default: Fallback value (defaults to NULL_TOKEN)
    
    Returns:
        default (or NULL_TOKEN) if val is null-like, else val stripped
    
    SAFETY: Always returns valid string, never None or empty string
    """
    if default is None:
        default = NULL_TOKEN
    
    if _is_null(val):
        return default
    
    return str(val).strip() if isinstance(val, str) else str(val)

    
def add_offline_runtime_dependencies(rows, offline_catalog_instance):
    """
    Append object-level Runtime_Dependency rows using offline SYS.OBJECT_DEPENDENCIES.

    Architectural guarantees:
    - OfflineCatalog MUST be instantiated once by xml_engine_plus.py
    - This function must NOT instantiate OfflineCatalog
    - SINGLE-RUN safe (no duplicate runtime deps)
    - Uses canonical key normalization
    - Fails fast on catalog contract violations
    """

    import logging
    log = logging.getLogger("xml-lineage")

    # Always work on a copy
    out = list(rows)
    seen = set()

    # --------------------------------------------------
    # SINGLE-RUN GUARD
    # --------------------------------------------------
    if any(r.get("Transformation_Type") == "Runtime_Dependency" for r in out):
        log.info("[CATALOG_OFFLINE] Runtime dependencies already added; skipping.")
        return out

    if offline_catalog_instance is None:
        log.info("[CATALOG_OFFLINE] Offline catalog not available; skipping runtime deps.")
        return out

    oc = offline_catalog_instance

    # --------------------------------------------------
    # Collect candidate objects (normalized)
    # --------------------------------------------------
    candidates = set()
    for r in out:
        for k in ("Current_Node", "Source_Object"):
            v = (r.get(k) or "").strip()
            if v and "." in v:
                schema, name = v.split(".", 1)
                candidates.add((
                    schema.strip().upper(),
                    name.strip().upper()
                ))
    # --------------------------------------------------
    # Resolve dependencies from catalog
    # --------------------------------------------------
    try:
        for (schema, name) in sorted(candidates):
            deps = oc.get_dependencies(schema, name)
            if not deps:
                continue

            for (bs, bo, bt, dep) in deps:
                key = (schema, name, bs, bo, bt, dep)
                if key in seen:
                    continue
                seen.add(key)

                out.append({
                    "Current_Node": f"{schema}.{name}",
                    "Hierarchy_Level": 1,
                    "SourceFile": "<RUNTIME_OFFLINE>",
                    "Parent_Input_Object": f"{schema}.{name}",
                    "Source_Object": f"{bs}.{bo}",
                    "Source_Column": "<NULL>",
                    "Target_Column": "*",
                    "Transformation_Type": "Runtime_Dependency",
                    "Formula": f"type={bt}; dependency={dep}",
                    "Filter_Expression": "<NULL>",
                    "Join_Name": "<NULL>",
                    "Join_Expression": "<NULL>",
                    "Join_Left_Object": "<NULL>",
                    "Join_Left_Column": "<NULL>",
                    "Join_Right_Object": "<NULL>",
                    "Join_Right_Column": "<NULL>",
                    "XPath": "<NULL>",
                })

        return out

    # --------------------------------------------------
    # Error semantics
    # --------------------------------------------------
    except RuntimeError:
        # Catalog contract violations must be fatal
        raise
    except Exception as e:
        # Optional runtime deps must not break pipeline
        log.warning(f"[CATALOG_OFFLINE] add_offline_runtime_dependencies failed: {e}")
        return out
def enrich_with_offline_catalog(rows, offline_catalog_instance):
    """
    Enhanced offline-catalog-based enrichment with performance optimizations.

    NEW CAPABILITIES:
    - Source_Object inference from dependencies
    - Target_Object recovery for incomplete lineage
    - Intelligent NULL classification with catalog context
    - Performance caching for repeated lookups
    - Comprehensive fallback chains

    Responsibilities:
    - Expand '*' Source_Table rows to explicit columns
    - Populate Node_Output_Columns from catalog
    - Recover metadata for NULL lineage gaps (catalog fallback)
    - Infer missing Source_Object/Target_Object from dependencies
    - Guaranteed SINGLE-RUN behavior
    - Guaranteed NO catalog re-instantiation
    - Safe for all XML shapes (joins, optimized CVs, partial graphs)

    This function MUST be called exactly once per pipeline execution,
    AFTER XML extraction + augmentation.
    """

    import logging
    log = logging.getLogger("xml-lineage")

    # Always work on a copy
    out = list(rows)

    # ------------------------------------------------------------------
    # SINGLE-RUN GUARD (hard stop)
    # ------------------------------------------------------------------
    if any("|CatalogFallback" in (r.get("Transformation_Type") or "") for r in out):
        log.info("[CATALOG] Enrichment already applied; skipping.")
        return out

    if offline_catalog_instance is None:
        log.info("[CATALOG] Offline catalog not available; skipping enrichment.")
        return out

    oc = offline_catalog_instance

    # ------------------------------------------------------------------
    # PERFORMANCE: Cache for repeated lookups
    # ------------------------------------------------------------------
    column_cache = {}  # (schema, name) -> list of columns
    dependency_cache = {}  # (schema, name) -> list of dependencies

    def get_cached_columns(schema, name):
        key = (schema.upper(), name.upper())
        if key not in column_cache:
            column_cache[key] = oc.get_view_columns(schema, name)
        return column_cache[key]

    def get_cached_dependencies(schema, name):
        key = (schema.upper(), name.upper())
        if key not in dependency_cache:
            dependency_cache[key] = oc.get_dependencies(schema, name)
        return dependency_cache[key]

    def normalize_parent_object(parent):
        if not parent:
            return ""
        candidate = str(parent).strip()
        if candidate.startswith("#"):
            candidate = candidate[1:].strip()
        # Use only the first identifier if a comma-separated or tokenized parent list is present.
        candidate = candidate.split(",")[0].split()[0].strip()
        return candidate

    def split_node_name(node):
        if node and "." in node:
            return node.split(".", 1)
        return "", node or ""

    # ------------------------------------------------------------------
    # TIER-1: Aggressive STAR Expansion (safe, catalog-backed)
    # ------------------------------------------------------------------
    expanded = []
    to_remove = []

    for r in out:
        src_col = (r.get("Source_Column") or "").strip()
        tgt_col = (r.get("Target_Column") or "").strip()
        node = (r.get("Current_Node") or "").strip()

        if src_col not in ("*", "") and tgt_col not in ("*", ""):
            continue

        schema, name = split_node_name(node)
        cols = get_cached_columns(schema, name)
        if not cols:
            continue

        for c in cols:
            nr = dict(r)
            if src_col == "*":
                nr["Source_Column"] = c
            if tgt_col == "*":
                nr["Target_Column"] = c
            nr["Transformation_Type"] = (
                (nr.get("Transformation_Type") or "") + "|StarExpanded"
            )
            expanded.append(nr)

        to_remove.append(r)

    # replace star rows with expanded rows
    for r in to_remove:
        out.remove(r)

    out.extend(expanded)

    # ------------------------------------------------------------------
    # TIER-2: Source_Object Inference from Parent Inputs + Dependencies
    # NEW: Fill missing Source_Object using parent input nodes first,
    #      then dependency relationships when available.
    # ------------------------------------------------------------------
    for r in out:
        src_obj = (r.get("Source_Object") or "").strip()
        node = (r.get("Current_Node") or "").strip()

        if src_obj and src_obj not in ("", "<NULL>"):
            continue  # Already has source object

        parent_candidate = normalize_parent_object(r.get("Parent_Input_Object"))
        if parent_candidate and parent_candidate not in ("", "<NULL>"):
            if parent_candidate != node:
                r["Source_Object"] = parent_candidate
                r["Transformation_Type"] = (
                    (r.get("Transformation_Type") or "") + "|ParentInputSource"
                )
                continue

            # Self-referential node inputs (e.g. Filter/Join inside same node)
            r["Source_Object"] = node
            r["Transformation_Type"] = (
                (r.get("Transformation_Type") or "") + "|SelfInputSource"
            )
            continue

        schema, name = split_node_name(node)
        deps = get_cached_dependencies(schema, name)

        if deps:
            # Use first dependency as source object
            base_schema, base_obj, base_type, dep_kind = deps[0]
            inferred_src = f"{base_schema}.{base_obj}"
            r["Source_Object"] = inferred_src
            r["Transformation_Type"] = (
                (r.get("Transformation_Type") or "") + "|SourceInferred"
            )

    # ------------------------------------------------------------------
    # TIER-3: Target_Object Recovery
    # NEW: Fill missing Target_Object for incomplete lineage
    # ------------------------------------------------------------------
    for r in out:
        tgt_obj = (r.get("Target_Object") or "").strip()
        node = (r.get("Current_Node") or "").strip()

        if tgt_obj and tgt_obj not in ("", "<NULL>"):
            continue  # Already has target object

        # Use current node as target object
        if node:
            r["Target_Object"] = node
            r["Transformation_Type"] = (
                (r.get("Transformation_Type") or "") + "|TargetRecovered"
            )

    # ------------------------------------------------------------------
    # TIER-4: Column Carry-Forward (Identity propagation)
    # Safe when exactly one source object exists
    # ------------------------------------------------------------------
    by_node = {}
    for r in out:
        node = (r.get("Current_Node") or "").strip()
        if node:
            by_node.setdefault(node, []).append(r)

    for node, rowset in by_node.items():
        sources = {
            (r.get("Source_Object") or "").strip()
            for r in rowset
            if (r.get("Source_Object") or "").strip()
            not in ("", "<NULL>")
        }

        if len(sources) != 1:
            continue  # ambiguous → do nothing

        src_obj = next(iter(sources))
        for r in rowset:
            tgt = (r.get("Target_Column") or "").strip()
            src = (r.get("Source_Column") or "").strip()

            if tgt and src in ("", "<NULL>"):
                r["Source_Column"] = tgt
                r["Transformation_Type"] = (
                    (r.get("Transformation_Type") or "") + "|CarryForward"
                )

    # ------------------------------------------------------------------
    # TIER-5: Intelligent NULL Classification with Context
    # ------------------------------------------------------------------
    for r in out:
        src = (r.get("Source_Column") or "").strip()
        tgt = (r.get("Target_Column") or "").strip()
        src_obj = (r.get("Source_Object") or "").strip()
        node = (r.get("Current_Node") or "").strip()

        # Context-aware NULL classification
        if src in ("", "<NULL>"):
            if src_obj and "." in src_obj:
                r["Source_Column"] = "<NULL_XML>"  # XML didn't specify
            else:
                r["Source_Column"] = "<NULL_INFERRED>"  # Could infer but no source object

        if tgt in ("", "<NULL>"):
            if "." in node:
                r["Target_Column"] = "<NULL_DERIVED>"  # Derived from node context
            else:
                r["Target_Column"] = "<NULL_UNKNOWN>"  # No context available

    # ------------------------------------------------------------------
    # LEGACY: STAR EXPANSION (Source_Table.*) - Keep for compatibility
    # ------------------------------------------------------------------
    expanded = []

    for r in out:
        tt = (r.get("Transformation_Type") or "").strip()
        if not tt.startswith("Source_Table"):
            continue

        src_obj = (r.get("Source_Object") or "").strip()
        src_col = (r.get("Source_Column") or "").strip()

        if not src_obj or "." not in src_obj:
            continue
        if src_col not in ("*", f"{src_obj}.*"):
            continue

        schema, name = src_obj.split(".", 1)
        cols = get_cached_columns(schema, name)

        if not cols:
            continue

        for c in cols:
            nr = dict(r)
            nr["Source_Column"] = c
            if nr.get("Target_Column") in (None, "", "*"):
                nr["Target_Column"] = c
            nr["Transformation_Type"] = tt + "_Expanded"
            expanded.append(nr)

    out.extend(expanded)

    # ------------------------------------------------------------------
    # LEGACY: NORMAL NODE OUTPUT POPULATION (catalog-backed)
    # ------------------------------------------------------------------
    by_node = {}
    for r in out:
        node = (r.get("Current_Node") or "").strip()
        if node:
            by_node.setdefault(node, []).append(r)

    for node, group in by_node.items():
        schema, name = split_node_name(node)
        cols = get_cached_columns(schema, name)

        if not cols:
            continue

        joined = ",".join(cols)
        for r in group:
            if not r.get("Node_Output_Columns") or r["Node_Output_Columns"] in ("", "<NULL>"):
                r["Node_Output_Columns"] = joined

    # ------------------------------------------------------------------
    # LEGACY: NULL → CATALOG FALLBACK (critical recovery rule)
    # ------------------------------------------------------------------
    for r in out:
        tgt = (r.get("Target_Column") or "").strip()
        node = (r.get("Current_Node") or "").strip()
        out_cols = (r.get("Node_Output_Columns") or "").strip()

        # Trigger strictly when lineage is broken
        if not tgt or out_cols not in ("", "<NULL>"):
            continue

        try:
            schema, name = split_node_name(node)
            cols = get_cached_columns(schema, name)

            if cols:
                r["Node_Output_Columns"] = ",".join(cols)
                r["Transformation_Type"] = (
                    (r.get("Transformation_Type") or "") + "|CatalogFallback"
                )

        except Exception as e:
            log.debug(f"[CATALOG_FALLBACK] failed node={node}: {e}")

    return out
# # ---------------- Orchestrator ----------------
# def augment_lineage(xml_path, tree, root, rows):
#     """
#     Post-processing that only APPENDS rows or ADDS metadata columns.
#     We now receive tree/root from the caller to avoid a second parse.
#     """
#     aug_cfg = getattr(config, "AUGMENTATION", {}) or {}
#     if root is None:
#         return rows
    
#     profile = _detect_profile_for_root(root)  # 'CV'/'ANV'/'ATV'/'DIM'/None
#     if aug_cfg.get("log_selected_profile", True):
#         _augment_log(f"Profile detected for {os.path.basename(xml_path)}: {profile or 'Unknown'}")

#     # CV-only augmentations
#     if profile == "CV":
#         if aug_cfg.get("enable_scenario_metadata", False):
#             attrs = getattr(config, "AUGMENTATION_SCENARIO_ATTRS", []) or []
#             rows = _augment_scenario_metadata(rows, xml_path, root, attrs)
#         if aug_cfg.get("enable_datasource_metadata", False):
#             rows = _augment_datasource_metadata(rows, xml_path, root)
#         if aug_cfg.get("emit_constant_mappings", False):
#             rows = _augment_constant_mappings(rows, xml_path, root, null_token=NULL_TOKEN)
#         if aug_cfg.get("emit_logical_bindings", False):
#             rows = _augment_logical_bindings(rows, xml_path, root, null_token=NULL_TOKEN)

#     # ATV
#     if profile == "ATV":
#         if aug_cfg.get("atv_emit_entity_star", False):
#             rows = _augment_atv_entity_star(rows, xml_path, root, null_token=NULL_TOKEN)
#         if aug_cfg.get("atv_emit_label_binding", False):
#             rows = _augment_atv_label_binding(rows, xml_path, root, null_token=NULL_TOKEN)

#     # ANV
#     if profile == "ANV":
#         if aug_cfg.get("anv_emit_shared_dimensions", True):
#             rows = _anv_emit_shared_dimensions(rows, xml_path, root, null_token=NULL_TOKEN)

#     # DIM
#     if profile == "DIM":
#         if aug_cfg.get("dim_enable_metadata", True):
#             attrs = getattr(config, "AUGMENTATION_DIM_ATTRS", []) or []
#             rows = _augment_dimension_metadata(rows, xml_path, root, attrs)
#         if aug_cfg.get("dim_enable_textual_metadata", True):
#             rows = _augment_dimension_textual_metadata(rows, xml_path, root)
#         if aug_cfg.get("dim_emit_source_tables", True):
#             rows = _dim_emit_source_tables(rows, xml_path, root, null_token=NULL_TOKEN)
#         if aug_cfg.get("dim_emit_key_mappings", True):
#             rows = _dim_emit_attribute_key_mappings(rows, xml_path, root, null_token=NULL_TOKEN)
#         if aug_cfg.get("dim_emit_joins", True):
#             rows = _dim_emit_joins(rows, xml_path, root, null_token=NULL_TOKEN)
#         if aug_cfg.get("dim_emit_table_filters", True):
#             rows = _dim_emit_table_filters(rows, xml_path, root, null_token=NULL_TOKEN)
#         if aug_cfg.get("dim_emit_calculated_attributes", True):
#             rows = _dim_emit_calculated_attributes(rows, xml_path, root, null_token=NULL_TOKEN)

#     return rows
def augment_lineage(xml_path, tree, root, rows, offline_catalog_instance=None):
    """
    Post-processing that only APPENDS rows or ADDS metadata columns.
    We now receive tree/root from the caller to avoid a second parse.
    """
    aug_cfg = getattr(config, "AUGMENTATION", {}) or {}

    # Safety guard
    if root is None:
        return rows

    # ----------------------------------------------------------
    # Detect object profile
    # ----------------------------------------------------------
    profile = _detect_profile_for_root(root)  # 'CV'/'ANV'/'ATV'/'DIM'/None
    if aug_cfg.get("log_selected_profile", True):
        _augment_log(
            f"Profile detected for {os.path.basename(xml_path)}: {profile or 'Unknown'}"
        )

    # ✅ NEW: Extract and apply semantic labels from CV/ANV output attributes
    semantic_labels = _extract_semantic_labels(root)
    label_map = semantic_labels.get("__OUTPUT__", {})
    if label_map and aug_cfg.get("capture_semantic_labels", True):
        for r in rows:
            tgt_col = (r.get("Target_Column") or "").strip()
            if tgt_col and tgt_col in label_map and (r.get("Label") or NULL_TOKEN) == NULL_TOKEN:
                r["Label"] = label_map[tgt_col]
        _augment_log(f"Applied {len(label_map)} semantic labels to rows")

    # ----------------------------------------------------------
    # CV-only augmentations
    # ----------------------------------------------------------
    if profile == "CV":
        if aug_cfg.get("enable_scenario_metadata", False):
            attrs = getattr(config, "AUGMENTATION_SCENARIO_ATTRS", []) or []
            rows = _augment_scenario_metadata(rows, xml_path, root, attrs)

        if aug_cfg.get("enable_datasource_metadata", False):
            rows = _augment_datasource_metadata(rows, xml_path, root)

        if aug_cfg.get("emit_constant_mappings", False):
            rows = _augment_constant_mappings(
                rows, xml_path, root, null_token=NULL_TOKEN
            )

        if aug_cfg.get("emit_logical_bindings", False):
            rows = _augment_logical_bindings(
                rows, xml_path, root, null_token=NULL_TOKEN
            )

    # ----------------------------------------------------------
    # ATV augmentations
    # ----------------------------------------------------------
    if profile == "ATV":
        if aug_cfg.get("atv_emit_entity_star", False):
            rows = _augment_atv_entity_star(
                rows, xml_path, root, null_token=NULL_TOKEN
            )

        if aug_cfg.get("atv_emit_label_binding", False):
            rows = _augment_atv_label_binding(
                rows, xml_path, root, null_token=NULL_TOKEN
            )

    # ----------------------------------------------------------
    # ANV augmentations
    # ----------------------------------------------------------
    if profile == "ANV":
        if aug_cfg.get("anv_emit_shared_dimensions", True):
            rows = _anv_emit_shared_dimensions(
                rows, xml_path, root, null_token=NULL_TOKEN
            )
        if aug_cfg.get("anv_emit_private_joins", True):
            rows = _dim_emit_joins(
                rows, xml_path, root, null_token=NULL_TOKEN
            )
        if aug_cfg.get("anv_emit_source_tables", True):
            rows = _dim_emit_source_tables(
                rows, xml_path, root, null_token=NULL_TOKEN
            )
        if aug_cfg.get("anv_emit_table_filters", True):
            rows = _dim_emit_table_filters(
                rows, xml_path, root, null_token=NULL_TOKEN
            )
        if aug_cfg.get("anv_emit_textual_metadata", True):
            rows = _augment_dimension_textual_metadata(rows, xml_path, root)
        if aug_cfg.get("anv_emit_exception_agg", True):
            rows = _anv_emit_exception_aggregation(
                rows, xml_path, root, null_token=NULL_TOKEN
            )

    # ----------------------------------------------------------
    # DIM augmentations
    # ----------------------------------------------------------
    if profile == "DIM":
        if aug_cfg.get("dim_enable_metadata", True):
            attrs = getattr(config, "AUGMENTATION_DIM_ATTRS", []) or []
            rows = _augment_dimension_metadata(rows, xml_path, root, attrs)

        if aug_cfg.get("dim_enable_textual_metadata", True):
            rows = _augment_dimension_textual_metadata(rows, xml_path, root)

        if aug_cfg.get("dim_emit_source_tables", True):
            rows = _dim_emit_source_tables(
                rows, xml_path, root, null_token=NULL_TOKEN
            )

        if aug_cfg.get("dim_emit_key_mappings", True):
            rows = _dim_emit_attribute_key_mappings(
                rows, xml_path, root, null_token=NULL_TOKEN
            )

        if aug_cfg.get("dim_emit_joins", True):
            rows = _dim_emit_joins(
                rows, xml_path, root, null_token=NULL_TOKEN
            )

        if aug_cfg.get("dim_emit_table_filters", True):
            rows = _dim_emit_table_filters(
                rows, xml_path, root, null_token=NULL_TOKEN
            )

        if aug_cfg.get("dim_emit_calculated_attributes", True):
            rows = _dim_emit_calculated_attributes(
                rows, xml_path, root, null_token=NULL_TOKEN
            )

    # ----------------------------------------------------------
    # ✅ OFFLINE CATALOG ENRICHMENT
    # ----------------------------------------------------------
    try:
        if offline_catalog_instance is not None:
            log.info("[augment] Applying offline catalog enrichment")
            rows = enrich_with_offline_catalog(rows, offline_catalog_instance)
    except Exception as e:
        log.warning(f"[augment] Offline catalog enrichment failed: {e}")

    return rows
# ---------------- Export ----------------
# def export_lineage(rows, path=None):
#     df = pd.DataFrame(rows)
#     if not df.empty:
#         df.fillna(NULL_TOKEN, inplace=True)
#     output_dir = config.OUTPUT.get("output_dir", ".")
#     os.makedirs(output_dir, exist_ok=True)
#     excel_name = config.OUTPUT.get("excel_name", "column_lineage.xlsx")
#     full_path = path or os.path.join(output_dir, excel_name)
#     try:
#         df.to_excel(full_path, index=False, engine="openpyxl")
#         log.info(f"[EXPORT] Full extract: {full_path} (rows={len(df)})")
#     except Exception as e:
#         log.exception(f"[EXPORT] Failed: {e}")
def export_lineage(rows, path=None):
    # ------------------------------
    # Build RAW lineage
    # ------------------------------
    df_raw = pd.DataFrame(rows)
    if not df_raw.empty:
        df_raw.fillna(NULL_TOKEN, inplace=True)

    output_dir = config.OUTPUT.get("output_dir", ".")
    os.makedirs(output_dir, exist_ok=True)
    excel_name = config.OUTPUT.get("excel_name", "column_lineage.xlsx")
    full_path = path or os.path.join(output_dir, excel_name)

    try:
        # ------------------------------
        # Build ANV_LOOKUP (derived)
        # ------------------------------
        if not df_raw.empty:
            anv_mask = df_raw["Current_Node"].astype(str).str.startswith("ANV_")
            df_anv = df_raw[anv_mask].copy()

            if not df_anv.empty:
                anv_lookup = (
                    df_anv
                    .groupby("Current_Node")
                    .agg(
                        Inferred_Parents=(
                            "Source_Object",
                            lambda x: ",".join(sorted(set(v for v in x if v and v != NULL_TOKEN)))
                        ),
                        Column_Count=("Target_Column", "nunique"),
                        Star_Join_Flag=(
                            "Source_Object",
                            lambda x: "Y" if len(set(v for v in x if v and v != NULL_TOKEN)) > 1 else "N"
                        ),
                        Shared_Dimension_Flag=(
                            "Transformation_Type",
                            lambda x: "Y" if "Shared_Dimension" in set(x) else "N"
                        ),
                    )
                    .reset_index()
                    .rename(columns={"Current_Node": "ANV_Name"})
                )
            else:
                anv_lookup = pd.DataFrame(
                    columns=[
                        "ANV_Name",
                        "Inferred_Parents",
                        "Column_Count",
                        "Star_Join_Flag",
                        "Shared_Dimension_Flag",
                    ]
                )

        else:
            anv_lookup = pd.DataFrame()

        # ------------------------------
        # Build ANV_COMPLETED_LINEAGE
        # (derived view, RAW unchanged)
        # ------------------------------
        if not df_raw.empty and not anv_lookup.empty:
            df_completed = df_raw.merge(
                anv_lookup[["ANV_Name", "Inferred_Parents", "Star_Join_Flag"]],
                how="left",
                left_on="Current_Node",
                right_on="ANV_Name",
            )

            # Fill missing Parent_Input_Object only for ANVs
            mask = (
                df_completed["Current_Node"].astype(str).str.startswith("ANV_")
                & (df_completed["Parent_Input_Object"] == NULL_TOKEN)
            )

            df_completed.loc[mask, "Parent_Input_Object"] = (
                df_completed.loc[mask, "Inferred_Parents"]
            )
        else:
            df_completed = df_raw.copy()

        # ------------------------------
        # Write Excel with multiple sheets
        # ------------------------------
        with pd.ExcelWriter(full_path, engine="openpyxl") as writer:
            df_raw.to_excel(writer, sheet_name="LINEAGE_RAW", index=False)
            anv_lookup.to_excel(writer, sheet_name="ANV_LOOKUP", index=False)
            df_completed.to_excel(writer, sheet_name="ANV_COMPLETED_LINEAGE", index=False)

        log.info(
            f"[EXPORT] Excel written: {full_path} "
            f"(rows={len(df_raw)}, anvs={len(anv_lookup)})"
        )

    except Exception as e:
        log.exception(f"[EXPORT] Failed: {e}")
def lint_lineage(rows):
    """
    Find problematic targets:
    - Target_Column is set and not NULL
    - Source_Column is missing/NULL
    - And the target is NOT present in Node_Output_Columns for that row's node
    Returns a list of dicts (rows) for reporting.
    """
    issues = []
    for r in rows:
        t = (r.get("Target_Column") or "").strip()
        s = (r.get("Source_Column") or "").strip()
        tt = (r.get("Transformation_Type") or "").strip()

        if not t or t == NULL_TOKEN:
            continue
        if tt in ("Filter", "Join", "Runtime_Dependency", "Source_Table"):
            # these legitimately may not have Source_Column
            continue
        if s and s != NULL_TOKEN:
            continue

        outputs = (r.get("Node_Output_Columns") or "").strip()
        outs = [c.strip() for c in outputs.split(",")] if outputs else []
        if t not in outs:
            issues.append({
                "Current_Node": r.get("Current_Node"),
                "Target_Column": t,
                "Source_Column": s or NULL_TOKEN,
                "Transformation_Type": tt or NULL_TOKEN,
                "Parent_Input_Object": r.get("Parent_Input_Object"),
                "Node_Output_Columns": outputs or NULL_TOKEN,
                "SourceFile": r.get("SourceFile"),
                "XPath": r.get("XPath"),
                "Hint": "No explicit source and target not found in node outputs. Check Projection/Union/Aggregation mappings."
            })
    return issues

def export_lint_report(issues, path=None):
    if not issues:
        log.info("[LINTER] No lineage gaps detected.")
        return None
    df = pd.DataFrame(issues).fillna(NULL_TOKEN)
    output_dir = config.OUTPUT.get("output_dir", ".")
    os.makedirs(output_dir, exist_ok=True)
    out_path = path or os.path.join(output_dir, "column_lineage_gaps.xlsx")
    try:
        df.to_excel(out_path, index=False, engine="openpyxl")
        log.info(f"[LINTER] Gaps report written: {out_path} (rows={len(df)})")
        return out_path
    except Exception as e:
        log.exception(f"[LINTER] Failed to write report: {e}")
        return None        

def _dedupe_rows_in_memory(rows, key_fields):
    seen = set()
    out = []
    for r in rows:
        key = tuple((r.get(k) or "") for k in key_fields)
        if key in seen:
            continue
        seen.add(key)
        out.append(r)
    return out

def export_traced_subset(rows, targets):
    if isinstance(targets, str):
        targets = [targets]
    targets = [t for t in targets if _trace_normalize_token(t)]
    seen = set()
    subset = []
    for t in targets:
        try:
            traced = trace_full_lineage(t, rows, include_expression_tokens=True)
        except Exception as e:
            log.exception(f"[TRACE] Failed for '{t}': {e}")
            traced = []
        for r in traced:
            key = (
                r.get("SourceFile"), r.get("XPath"), r.get("Transformation_Type"),
                r.get("Current_Node"), r.get("Parent_Input_Object"),
                r.get("Target_Column"), r.get("Source_Column"),
                r.get("Formula"), r.get("Filter_Expression"), r.get("Join_Expression"),
                r.get("Join_Left_Column"), r.get("Join_Right_Column"),
            )
            if key not in seen:
                seen.add(key)
                subset.append(r)
    if not subset:
        log.warning("[TRACE] No traced rows. Check target names or expressions/columns.")
        return None
    output_dir = config.OUTPUT.get("output_dir", ".")
    os.makedirs(output_dir, exist_ok=True)
    stem = os.path.splitext(config.OUTPUT.get("excel_name", "column_lineage.xlsx"))[0]
    out_path = os.path.join(output_dir, f"{stem}_traced_subset.xlsx")
    try:
        df = pd.DataFrame(subset).fillna(NULL_TOKEN)
        trace_summary = []
        if "Trace_Target" in df.columns:
            for target, group in df.groupby("Trace_Target"):
                trace_summary.append({
                    "Trace_Target": target,
                    "Requested_Targets": ", ".join(sorted({str(t) for t in targets if _trace_normalize_token(t)})),
                    "Row_Count": len(group),
                    "Unique_Current_Nodes": group["Current_Node"].nunique() if "Current_Node" in group else 0,
                    "Unique_Source_Columns": group["Source_Column"].nunique() if "Source_Column" in group else 0,
                    "Max_Trace_Depth": int(group["Trace_Depth"].max()) if "Trace_Depth" in group else 0,
                })
        else:
            trace_summary.append({
                "Requested_Targets": ", ".join(sorted({str(t) for t in targets if _trace_normalize_token(t)})),
                "Row_Count": len(df),
            })

        with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="TRACE_ROWS", index=False)
            pd.DataFrame(trace_summary).to_excel(writer, sheet_name="TRACE_SUMMARY", index=False)

        log.info(f"[TRACE] Traced subset exported: {out_path} (rows={len(df)})")
        return out_path
    except Exception as e:
        log.exception(f"[TRACE] Failed to export subset: {e}")
        return None

# ---------------- Deep Recursive Trace ----------------
def trace_full_lineage(target_column, rows, *, include_expression_tokens=True, max_hops=9999):
    """
    Deep upstream tracer:
    - Index by normalized Target_Column
    - Parents from Source_Column, Join columns, Formula/Filter/Join expr tokens
    """
    import pandas as pd, re
    df = pd.DataFrame(rows)
    if not df.empty:
        df = df.fillna("")

    def _trace_norm_base(value: str) -> str:
        if not value:
            return ""
        return _trace_normalize_token(value).split(".")[-1].strip()

    def resolve_matches(term: str):
        if not term:
            return []
        if term in by_target:
            return [term]
        return sorted(target_by_base.get(_trace_norm_base(term), []))

    df["_SRC_N"] = df.apply(
        lambda r: _qualify_lineage_column(
            r,
            ["Parent_Input_Object", "Source_Object", "Source_Node"],
            "Source_Column",
        ),
        axis=1,
    )
    df["_TGT_N"] = df.apply(
        lambda r: _qualify_lineage_column(
            r,
            ["Current_Node", "Target_Object", "Target_Node"],
            "Target_Column",
        ),
        axis=1,
    )
    df["_JLC_N"] = df.apply(
        lambda r: _qualify_lineage_column(
            r,
            ["Join_Left_Object"],
            "Join_Left_Column",
        ),
        axis=1,
    )
    df["_JRC_N"] = df.apply(
        lambda r: _qualify_lineage_column(
            r,
            ["Join_Right_Object"],
            "Join_Right_Column",
        ),
        axis=1,
    )

    target_by_base = defaultdict(set)
    known_exact = set()

    def add_token(token: str):
        if not token:
            return
        known_exact.add(token)
        target_by_base[_trace_norm_base(token)].add(token)

    for col in ["_SRC_N", "_TGT_N", "_JLC_N", "_JRC_N"]:
        if col in df.columns:
            for v in df[col].tolist():
                add_token(_trace_normalize_token(v))

    if include_expression_tokens:
        for col in ["Formula", "Filter_Expression", "Join_Expression"]:
            if col not in df.columns:
                continue
            for expr in df[col].tolist():
                if not expr:
                    continue
                expr_text = re.sub(r"'[^']*'", " ", str(expr))
                for quoted in re.findall(r'"([^"]+)"', expr_text):
                    token = _trace_normalize_token(quoted)
                    if token in target_by_base:
                        for full in target_by_base[_trace_norm_base(token)]:
                            add_token(full)
                for tok in re.findall(r"[A-Za-z_][A-Za-z0-9_.]*", expr_text):
                    if tok.upper() in SQL_KEYWORDS:
                        continue
                    token = _trace_normalize_token(tok)
                    if token in target_by_base:
                        for full in target_by_base[_trace_norm_base(token)]:
                            add_token(full)

    def extract_cols(expr):
        if not expr:
            return []
        expr_text = re.sub(r"'[^']*'", " ", str(expr))
        out = []
        for quoted in re.findall(r'"([^"]+)"', expr_text):
            token = _trace_normalize_token(quoted)
            if not token:
                continue
            if token in known_exact:
                out.append(token)
            else:
                out.extend(resolve_matches(token))
        for tok in re.findall(r"[A-Za-z_][A-Za-z0-9_.]*", expr_text):
            if tok.upper() in SQL_KEYWORDS:
                continue
            token = _trace_normalize_token(tok)
            if not token:
                continue
            if token in known_exact:
                out.append(token)
            else:
                out.extend(resolve_matches(token))
        return list(dict.fromkeys(out))

    by_target = defaultdict(list)
    for _, r in df.iterrows():
        tgt = _trace_normalize_token(r.get("_TGT_N", ""))
        if tgt:
            by_target[tgt].append(r.to_dict())

    visited_cols, visited_rows = set(), set()
    chain = []

    def rid(r):
        return (
            r.get("SourceFile"), r.get("XPath"), r.get("Transformation_Type"),
            r.get("Current_Node"), r.get("Parent_Input_Object"),
            r.get("Target_Column"), r.get("Source_Column"),
            r.get("Formula"), r.get("Filter_Expression"), r.get("Join_Expression"),
            r.get("Join_Left_Column"), r.get("Join_Right_Column"),
        )

    def recurse(col, depth=0):
        term = _trace_normalize_token(col)
        if not term or term in visited_cols or depth > max_hops:
            return
        visited_cols.add(term)
        for match in resolve_matches(term):
            rows_for_target = by_target.get(match, [])
            if not rows_for_target:
                continue
            for r in rows_for_target:
                key = rid(r)
                if key in visited_rows:
                    continue
                visited_rows.add(key)
                traced_row = dict(r)
                traced_row["Trace_Target"] = match
                traced_row["Trace_Depth"] = depth
                traced_row["Trace_Parents"] = ";".join(
                    [p for p in (r.get("_SRC_N") or "", r.get("_JLC_N") or "", r.get("_JRC_N") or "") if p]
                )
                traced_row["Trace_Reason"] = "STRUCTURAL"
                if r.get("Formula"):
                    traced_row["Trace_Reason"] = "FORMULA"
                elif r.get("Filter_Expression"):
                    traced_row["Trace_Reason"] = "FILTER"
                elif r.get("Join_Expression"):
                    traced_row["Trace_Reason"] = "JOIN"
                chain.append(traced_row)
                parents = []
                if r.get("_SRC_N"):
                    parents.append(r["_SRC_N"])
                parents.extend(extract_cols(r.get("Formula", "")))
                parents.extend(extract_cols(r.get("Filter_Expression", "")))
                parents.extend(extract_cols(r.get("Join_Expression", "")))
                if r.get("_JLC_N"):
                    parents.append(r["_JLC_N"])
                if r.get("_JRC_N"):
                    parents.append(r["_JRC_N"])
                for p in parents:
                    recurse(p, depth + 1)

    recurse(target_column, 0)
    return chain

# -------- Diagnostics: NULL profiling & file utilization --------
def _print_section(title: str):
    print("\n" + "=" * 90)
    print(title)
    print("=" * 90)

def _pct(n, d):
    return (100.0 * n / d) if (d and d > 0) else 0.0

def print_null_profile(rows, *, key_columns=None, top_n=40):
    """
    Prints NULL/NOT-NULL counts and percentages for each column over the full dataset.
    key_columns: optional prioritized column list to print first (e.g., lineage fields).
    top_n: when DataFrame is very wide, limit to first N 'other' columns for brevity.
    """
    import pandas as pd
    if not rows:
        print("[diag] No rows to profile.")
        return
    df = pd.DataFrame(rows).copy()
    # Normalize empties
    for c in df.columns:
        df[c] = df[c].fillna("").astype(str)
    total = len(df)

    # Prefer to show key columns first if provided
    if key_columns:
        ordered = [c for c in key_columns if c in df.columns] + \
                  [c for c in df.columns if c not in (key_columns or [])]
    else:
        ordered = list(df.columns)

    if not key_columns and len(ordered) > top_n:
        ordered = ordered[:top_n]

    _print_section(f"[NULL PROFILE] Columns = {len(df.columns)} | Rows = {total}")
    for col in ordered:
        if col not in df.columns:
            continue
        nulls = (df[col].str.strip() == "").sum()
        not_nulls = total - nulls
        print(f"[col] {col:30s}  null={nulls:8d} ({_pct(nulls, total):6.2f}%)"
              f" | not-null={not_nulls:8d} ({_pct(not_nulls, total):6.2f}%)")

def print_file_utilization(rows, *, info_columns=None, min_show=1, top_n_files=100):
    """
    'Utilization' per SourceFile:
      - total rows per file
      - % informative rows (rows where at least one of info_columns is non-empty)
      - per-file NULL% for each info column

    info_columns: columns we treat as informative for lineage (defaults provided below).
    min_show: only print files with at least this many rows.
    """
    import pandas as pd
    if not rows:
        print("[diag] No rows to profile by file.")
        return
    df = pd.DataFrame(rows).copy()
    for c in df.columns:
        df[c] = df[c].fillna("").astype(str)

    if "SourceFile" not in df.columns:
        print("[diag] No 'SourceFile' column—skipping file utilization.")
        return

    default_info_cols = [
        "Transformation_Type", "Current_Node", "Parent_Input_Object",
        "Source_Object", "Source_Column", "Target_Column",
        "Formula", "Filter_Expression", "Join_Expression"
    ]
    info_columns = info_columns or [c for c in default_info_cols if c in df.columns]
    if not info_columns:
        print("[diag] No informative columns present—skipping file utilization.")
        return

    # A row is "informative" if any info column is non-empty
    any_info = (df[info_columns].apply(lambda r: any(str(v).strip() != "" for v in r), axis=1))
    df["_is_informative"] = any_info

    g = df.groupby("SourceFile", dropna=False)
    counts = g.size().rename("rows").reset_index()
    counts = counts.sort_values("rows", ascending=False).head(top_n_files)

    _print_section(f"[FILE UTILIZATION] Files = {g.ngroups} (showing up to {top_n_files})")
    for _, rec in counts.iterrows():
        src = rec["SourceFile"]
        total = int(rec["rows"])
        if total < min_show:
            continue
        sub = df[df["SourceFile"] == src]
        informative = int(sub["_is_informative"].sum())
        print(f"[file] {src}  rows={total}  informative={informative} ({_pct(informative, total):6.2f}%)")

        # Per-file null profile on the informative columns only
        for col in info_columns:
            nulls = (sub[col].str.strip() == "").sum()
            print(f"       - {col:22s} null={nulls:6d} ({_pct(nulls, total):6.2f}%)  not-null={total - nulls:6d} ({_pct(total - nulls, total):6.2f}%)")
       
# --------------------------------------------------------------------
# Smart export helpers (prevents Excel oversize errors)
# --------------------------------------------------------------------
def _smart_choose_format(n_rows, prefer="parquet", max_excel_rows=1_000_000):
    """
    Decide final output format based on size and preference.
    Returns one of: "xlsx", "csv", "parquet".
    - If rows > max_excel_rows: avoid Excel (return csv/parquet)
    - Else follow preference (default parquet)
    """
    if n_rows is None:
        return prefer or "parquet"
    if n_rows > max_excel_rows:
        # prefer Parquet for big datasets unless user insists on CSV
        return "csv" if (prefer or "").lower() == "csv" else "parquet"
    return (prefer or "parquet").lower()

def _export_dataframe_smart(df, base_stem, output_dir, config_module):
    """
    Export df using SMART_EXPORT rules from config.py:
      - parquet if preferred and safe
      - csv if too large for Excel
      - xlsx if small enough OR explicitly preferred; supports sheet splitting
    Returns the path(s) written as a list.
    """
    import os, math
    import pandas as pd

    os.makedirs(output_dir, exist_ok=True)

    smart = getattr(config_module, "SMART_EXPORT", {}) or {}
    prefer       = (smart.get("prefer", "parquet") or "parquet").lower()
    max_xlsx_rows= int(smart.get("max_excel_rows", 1_000_000))
    split_ok     = bool(smart.get("split_excel_if_needed", True))
    per_sheet    = int(smart.get("excel_sheet_rows", 1_000_000))
    delimiter    = smart.get("csv_delimiter", ",")

    n_rows = len(df)
    choice = _smart_choose_format(n_rows, prefer, max_xlsx_rows)

    # 1) Parquet
    if choice == "parquet":
        path = os.path.join(output_dir, f"{base_stem}.parquet")
        df.to_parquet(path, index=False)
        return [path]

    # 2) CSV
    if choice == "csv":
        path = os.path.join(output_dir, f"{base_stem}.csv")
        df.to_csv(path, index=False, encoding="utf-8", sep=delimiter)
        return [path]

    # 3) XLSX — only if small enough
    if n_rows <= max_xlsx_rows:
        path = os.path.join(output_dir, f"{base_stem}.xlsx")
        with pd.ExcelWriter(path, engine="openpyxl") as xw:
            df.to_excel(xw, index=False, sheet_name="Sheet1")
        return [path]

    # If Excel was chosen but rows exceed threshold
    if not split_ok:
        # fall back to CSV
        path = os.path.join(output_dir, f"{base_stem}.csv")
        df.to_csv(path, index=False, encoding="utf-8", sep=delimiter)
        return [path]

    # Multi-sheet split
    path = os.path.join(output_dir, f"{base_stem}.xlsx")
    sheets = math.ceil(n_rows / per_sheet)
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        for i in range(sheets):
            start = i * per_sheet
            stop = min(start + per_sheet, n_rows)
            df.iloc[start:stop].to_excel(xw, index=False, sheet_name=f"Part_{i+1}")
    return [path]

# xml_engine.py → REPLACEMENT FOR export_rows_handoff()
def export_rows_handoff(rows, *, path=None, fmt="xlsx"):
    """
    Write handoff rows for HeavyEngine or downstream processors.
    Guaranteed safe behavior:
      - Handles directory or file path
      - Generates timestamped names if directory provided
      - Supports CSV / XLSX / Parquet
      - Ensures string columns are high-quality dtype ("string")
      - Returns the full output path
    """
    import os, time
    import pandas as pd

    df = pd.DataFrame(rows)

    # ----------------------
    # 1) Normalize DataFrame
    # ----------------------
    if df.empty:
        log.warning("[HANDOFF] DataFrame is empty — nothing written.")
        return None

    df = df.fillna(NULL_TOKEN)

    # Convert object columns to Pandas "string" dtype for safety
    for col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].astype("string")

    # ----------------------
    # 2) Determine base path
    # ----------------------
    # Priority:
    #   1) explicit path from caller
    #   2) config.OUTPUT['handoff_rows_path']
    #   3) output_dir/handoff_rows
    base = path or config.OUTPUT.get("handoff_rows_path")

    if not base:
        base = os.path.join(
            config.OUTPUT.get("output_dir", "."),
            "handoff_rows"
        )

    # ----------------------
    # 3) Resolve directory vs file behavior
    # ----------------------
    # CASE A: base is a directory → create timestamped file inside it
    # CASE B: base is a file → use its directory + name
    # CASE C: base has no extension = treat as directory
    is_directory = (
        os.path.isdir(base) or 
        os.path.splitext(base)[1] == ""
    )

    if is_directory:
        os.makedirs(base, exist_ok=True)
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        stem = f"handoff_rows_{timestamp}"
        dir_path = base
    else:
        # base is a file path (e.g. ".../handoff_rows.csv")
        dir_path = os.path.dirname(base)
        os.makedirs(dir_path, exist_ok=True)
        stem = os.path.splitext(os.path.basename(base))[0]

    # ----------------------
    # 4) Export according to fmt
    # ----------------------
    fmt = (fmt or "csv").lower()

    if fmt == "parquet":
        out_path = os.path.join(dir_path, f"{stem}.parquet")
        df.to_parquet(out_path, index=False)

    elif fmt == "xlsx":
        out_path = os.path.join(dir_path, f"{stem}.xlsx")
        df.to_excel(out_path, index=False, engine="openpyxl")

    else:   # CSV is default
        out_path = os.path.join(dir_path, f"{stem}.csv")
        df.to_csv(out_path, index=False, encoding="utf-8", sep=",")

    # ----------------------
    # 5) Log + return
    # ----------------------
    log.info(f"[HANDOFF] Rows written for HeavyEngine: {out_path} (rows={len(df)})")
    return out_path
# ---------------- Main ----------------


# ========================
# Post-processing helpers
# ========================

def _tokenize_expr_components(expr: str):
    """Return counts for columns, constants, functions, predicates for a simple Wt/Wf model."""
    if not expr:
        return 0, 0, 0, 0
    s = str(expr)

    cols   = len(extract_columns_from_expression(s))
    # string literal OR number (int/float)
    consts = len(re.findall(r"(?:'[^']*'|\b\d+(?:\.\d+)?\b)", s))
    # function call tokens like NAME(...)
    funcs  = len(re.findall(r"\b[A-Za-z_][A-Za-z0-9_]*\s*\(", s))
    # simple predicate tokens
    preds  = len(re.findall(r"=|<>|>=|<=|\bIN\b|\bLIKE\b|\bAND\b|\bOR\b", s, flags=re.IGNORECASE))
    return cols, consts, funcs, preds

def _compute_weight(expr: str, positive_functions=None, clamp=(0.0,1.0)):
    cols, consts, funcs, preds = _tokenize_expr_components(expr)
    if positive_functions:
        for fn in positive_functions:
            funcs -= len(re.findall(rf"\b{re.escape(fn)}\b\s*\(", str(expr), flags=re.IGNORECASE))
        funcs = max(funcs, 0)
    denom = max(cols + consts + funcs + preds, 0)
    w = (cols / denom) if denom else 0.0
    lo, hi = clamp
    return max(lo, min(hi, w))

def compute_weights(rows, cfg):
    if not cfg.get('enabled', True):
        return rows
    pos_fns = cfg.get('positive_functions', [])
    clamp = (cfg.get('min_weight', 0.0), cfg.get('max_weight', 1.0))
    join_contrib = bool(cfg.get('join_keys_contribute_to_Wf', True))
    out = []
    for r in rows:
        rr = dict(r)
        rr['Weight_Transform'] = _compute_weight(rr.get('Formula') or '', pos_fns, clamp)
        wf = _compute_weight(rr.get('Filter_Expression') or '', pos_fns, clamp)
        if join_contrib:
            jl = rr.get('Join_Left_Column') or ''
            jr = rr.get('Join_Right_Column') or ''
            if jl or jr:
                wf = max(wf, 0.5)
        rr['Weight_Filter'] = wf
        out.append(rr)
    return out

def propagate_rules(rows, cfg):
    if not cfg.get('enabled', True):
        return rows
    out = list(rows)
    if cfg.get('propagate_filter_influence', True):
        factor = float(cfg.get('filter_weight_factor', 1.0))
        extra = []
        for r in rows:
            t = (r.get('Target_Column') or '').strip()
            if not t:
                continue
            fexpr = r.get('Filter_Expression') or ''
            toks = extract_columns_from_expression(fexpr)
            for c in toks:
                extra.append({
                    **r,
                    'Transformation_Type': 'Filter_Propagation',
                    'Source_Column': c,
                    'Weight_Filter': max(r.get('Weight_Filter', 0.0) * factor, 0.0)
                })
        out.extend(extra)
    roll = cfg.get('parent_rollup', {})
    if roll.get('enabled', True):
        levels = roll.get('levels', ['table'])
        agg = roll.get('aggregation', 'max')
        extra = []
        for r in rows:
            tgt_parent = (r.get('Current_Node') or '').strip()
            src_parent = (r.get('Source_Object') or '').strip()
            if not tgt_parent or not src_parent:
                continue
            e = dict(r)

            e['Transformation_Type'] = 'Parent_Rollup'
            # ✅ NODE-level lineage (never pollute column fields)
            e['Lineage_Level'] = 'NODE'
            e['Target_Node'] = tgt_parent
            e['Source_Node'] = src_parent

            e['Target_Column'] = ""   # ✅ keep column lineage clean
            e['Source_Column'] = ""

            e['Rollup_Levels'] = ','.join(levels)
            e['Rollup_Aggregation'] = agg
            extra.append(e)
        out.extend(extra)
    if cfg.get('aggregate_parallel_edges', True):
        mode = cfg.get('parallel_edge_aggregation', 'weighted_mean')
        key = lambda r: (r.get('Current_Node'), r.get('Target_Column'), r.get('Source_Column'))
        buckets = {}
        for r in out:
            buckets.setdefault(key(r), []).append(r)
        merged = []
        for k, lst in buckets.items():
            if len(lst) == 1:
                merged.append(lst[0])
                continue
            base = dict(lst[0])
            wts = [x.get('Weight_Transform', 0.0) for x in lst]
            wfs = [x.get('Weight_Filter', 0.0) for x in lst]
            if mode == 'sum':
                base['Weight_Transform'] = sum(wts)
                base['Weight_Filter'] = sum(wfs)
            elif mode == 'max':
                base['Weight_Transform'] = max(wts)
                base['Weight_Filter'] = max(wfs)
            else:
                base['Weight_Transform'] = sum(wts) / max(len(wts),1)
                base['Weight_Filter'] = sum(wfs) / max(len(wfs),1)
            merged.append(base)
        out = merged
    return out

def build_closure(rows, cfg):
    if not cfg.get('build_lineage_closure', False):
        return {}
    adj = {}
    for r in rows:
        t = (r.get('Target_Column') or '').strip()
        s = (r.get('Source_Column') or '').strip()
        if not t or not s or t == '<NULL>' or s == '<NULL>':
            continue
        adj.setdefault(t, set()).add(s)
    closure = {}
    temp_mark = set(); perm_mark = set(); order = []
    def visit(n):
        if n in perm_mark or n in temp_mark:
            return
        temp_mark.add(n)
        for m in adj.get(n, []):
            visit(m)
        temp_mark.remove(n)
        perm_mark.add(n)
        order.append(n)
    for node in list(adj.keys()):
        visit(node)
    for n in order:
        deps = set(adj.get(n, set()))
        acc = set()
        for m in deps:
            acc.add(m)
            acc |= closure.get(m, set())
        closure[n] = acc
    return closure

def apply_semantic_layer(rows, cfg):
    if not cfg.get('enable_semantic_layer', False):
        return rows
    wt_min = float(cfg.get('min_transform_weight', 0.0))
    wf_min = float(cfg.get('min_filter_weight', 0.0))
    keep = []
    for r in rows:
        if r.get('Weight_Transform', 0.0) >= wt_min or r.get('Weight_Filter', 0.0) >= wf_min:
            keep.append(r)
    return keep

def compute_scores(rows, cfg):
    if not cfg.get('enabled', True):
        return rows
    epsilon = float(cfg.get('epsilon', 1e-9))
    emit = bool(cfg.get('emit_scores_into_rows', True))
    by_node = {}
    for r in rows:
        n = (r.get('Current_Node') or '').strip()
        if not n:
            continue
        by_node.setdefault(n, []).append(r)
    node_scores = {}
    for n, lst in by_node.items():
        source_sum = sum(x.get('Weight_Transform', 0.0) + x.get('Weight_Filter', 0.0) for x in lst if x.get('Source_Column'))
        target_sum = sum(x.get('Weight_Transform', 0.0) + x.get('Weight_Filter', 0.0) for x in lst if x.get('Target_Column'))
        denom = max(source_sum + target_sum, epsilon)
        LLD = source_sum / denom
        LID = source_sum / denom
        node_scores[n] = (LLD, LID)
    if emit:
        out = []
        for r in rows:
            n = (r.get('Current_Node') or '').strip()
            LLD, LID = node_scores.get(n, (0.0, 0.0))
            rr = dict(r)
            rr['LLD_Score'] = LLD
            rr['LID_Score'] = LID
            out.append(rr)
        return out
    return rows

def generate_business_semantics(rows, cfg, out_dir):
    if not cfg.get('enabled', True):
        return
    norm_cfg = cfg.get('normalize', {})
    dict_cfg = cfg.get('dictionary', {})
    drop_terms = set(t.lower() for t in dict_cfg.get('drop_terms', []))
    amap = {k.lower(): v for k, v in (dict_cfg.get('default_map') or {}).items()}
    def normalize(name):
        if not name:
            return ''
        s = str(name)
        if norm_cfg.get('split_camel_case', True):
            s = re.sub(r'(?<!^)(?=[A-Z])', ' ', s)
        if norm_cfg.get('split_on_non_alnum', True):
            s = re.sub(r'[^A-Za-z0-9]+', ' ', s)
        toks = [t for t in s.split() if len(t) >= int(norm_cfg.get('min_token_len', 2))]
        norm = []
        for t in toks:
            low = t.lower()
            if low in drop_terms:
                continue
            low = amap.get(low, low)
            norm.append(low)
        if norm_cfg.get('titlecase_output', True):
            return ' '.join(w.capitalize() for w in norm)
        return ' '.join(norm)
    concepts = {}
    for r in rows:
        for candidate in [(r.get('Current_Node') or ''), (r.get('Target_Column') or ''), (r.get('Source_Column') or '')]:
            c = normalize(candidate)
            if c:
                concepts.setdefault(c, set()).add(candidate)
    os.makedirs(out_dir, exist_ok=True)
    conc_path = os.path.join(out_dir, cfg.get('export', {}).get('concepts_path', 'business_semantics_concepts.json'))
    onto_path = os.path.join(out_dir, cfg.get('export', {}).get('ontology_path', 'business_semantics_ontology.json'))
    with open(conc_path, 'w', encoding='utf-8') as f:
        json.dump({k: sorted(list(v)) for k, v in sorted(concepts.items())}, f, indent=2, ensure_ascii=False)
    ontology = {'concepts': sorted(list(concepts.keys())), 'relations': []}
    with open(onto_path, 'w', encoding='utf-8') as f:
        json.dump(ontology, f, indent=2, ensure_ascii=False)

if __name__ == "__main__":

    # ============================================================
    # DISCOVER XML FILES
    # ============================================================
    input_path = config.INPUT.get("path", ".")
    recursive = config.INPUT.get("recursive", False)
    xml_files = discover_xml_files(input_path, recursive)

    if not xml_files:
        log.error("No XML files found.")
        raise SystemExit(1)

    log.info(f"[ENGINE] Discovered {len(xml_files)} XML files under: {input_path}")

    all_rows = []

    # Only allocate the graph if enabled
    build_graph = getattr(config, "PERF", {}).get("build_lineage_graph", False)
    combined_graph = {} if build_graph else None

    # ============================================================
    # EXTRACT & AUGMENT XML FILES
    # ============================================================
    for idx, file in enumerate(xml_files, 1):
        log.info(f"[ENGINE] Processing XML ({idx}/{len(xml_files)}): {file}")

        # Extract lineage
        rows, graph, tree, root = extract_lineage_from_xml(file)
        if not rows:
            log.warning(f"[WARN] No rows produced from {file}")

        # Per-file augmentation (no re-parse)
        try:
            aug_rows = augment_lineage(file, tree, root, rows)
        except Exception as _e:
            log.warning(f"[AUGMENT] Skipped for {file}: {_e}")
            aug_rows = []

        # Collect (if not using SQLite sink)
        all_rows.extend(rows)
        all_rows.extend(aug_rows)

        # Combine lineage graph only if enabled
        if build_graph and graph:
            for k, v in graph.items():
                combined_graph.setdefault(k, set()).update(v)
    # ============================================================
    # DEDUPE BASE ROWS
    # ============================================================
    dedupe_keys = [
        "SourceFile", "XPath", "Transformation_Type", "Current_Node", "Parent_Input_Object",
        "Target_Column", "Source_Column", "Formula", "Filter_Expression", "Join_Expression",
        "Join_Left_Column", "Join_Right_Column"
    ]
    all_rows = _dedupe_rows_in_memory(all_rows, dedupe_keys)
    log.info(f"[DEDUPE] Base deduped rows = {len(all_rows)}")
    # === Research Enhancements wiring ===
    all_rows = compute_weights(all_rows, getattr(config, 'WEIGHTS', {}) or {})
    all_rows = propagate_rules(all_rows, getattr(config, 'PROPAGATION', {}) or {})
    all_rows = apply_semantic_layer(all_rows, getattr(config, 'SEMANTICS', {}) or {})
    closure = build_closure(all_rows, getattr(config, 'CLOSURE', {}) or {})
    all_rows = compute_scores(all_rows, getattr(config, 'SCORING', {}) or {})


    # ============================================================
    # GLOBAL PULL-THROUGH
    # ============================================================
    try:
        before = len(all_rows)
        if getattr(config, "AUGMENTATION", {}).get("global_anv_pullthrough", False):
            all_rows = _global_anv_shared_dimension_pullthrough(all_rows)
        all_rows = _dedupe_rows_in_memory(all_rows, dedupe_keys)
        log.info(f"[GLOBAL] Pull-through added {len(all_rows)-before} rows")
    except Exception as e:
        log.warning(f"[GLOBAL] Pull-through augmentation skipped: {e}")


    # =======================
    # OFFLINE CATALOG ENRICHMENT (SINGLE PASS)
    # =======================
    if config.AUGMENTATION_PIPELINE.get("run_offline_catalog_enrichment", True):
        before_rows = len(all_rows)
        before_expanded = sum(1 for r in all_rows if (r.get("Transformation_Type") or "").endswith("_Expanded"))
        try:
            all_rows = enrich_with_offline_catalog(all_rows)  # << keep only ONE call in the entire script
            all_rows = _dedupe_rows_in_memory(all_rows, dedupe_keys)
            after_rows = len(all_rows)
            after_expanded = sum(1 for r in all_rows if (r.get("Transformation_Type") or "").endswith("_Expanded"))
            log.info(f"[CATALOG_OFFLINE] Star-expansion added {after_rows - before_rows} rows; expanded cols increased {after_expanded - before_expanded}")
        except Exception as e:
            log.warning(f"[CATALOG_OFFLINE] Enrichment skipped: {e}")

    if config.AUGMENTATION_PIPELINE.get("run_runtime_dependencies", False):
        dep_before = len(all_rows)
        all_rows = add_offline_runtime_dependencies(all_rows)
        all_rows = _dedupe_rows_in_memory(all_rows, dedupe_keys)
        log.info(f"[CATALOG_OFFLINE] Runtime dependencies added {len(all_rows) - dep_before} rows")
    try:
        bs_cfg = getattr(config, 'BUSINESS_SEMANTICS', {}) or {}
        out_dir = config.OUTPUT.get('output_dir', '.')
        generate_business_semantics(all_rows, bs_cfg, out_dir)
    except Exception as e:
        log.warning(f"[SEMANTICS] Generation skipped: {e}")


    # ============================================================
    # LINTER: COLUMN GAP CHECK
    # ============================================================
    try:
        issues = lint_lineage(all_rows)
        if issues:
            log.warning(f"[LINTER] Found {len(issues)} lineage gaps (incomplete mappings)")
        export_lint_report(issues)
    except Exception as e:
        log.warning(f"[LINTER] Skipped: {e}")

    if not all_rows:
        log.error("No lineage rows extracted after enrichment.")
        raise SystemExit(1)

    # >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    # NEW: Diagnostics — NULL / NOT-NULL and file utilization
    if getattr(config, "PERF", {}).get("enable_diagnostics", False) and log.level <= logging.DEBUG:
        try:
            key_cols = [
                "SourceFile","Transformation_Type","Current_Node","Parent_Input_Object",
                "Source_Object","Source_Column","Target_Column",
                "Node_Output_Columns","Formula","Filter_Expression","Join_Expression"
            ]
            print_null_profile(all_rows, key_columns=key_cols, top_n=80)
            # print_file_utilization(all_rows, info_columns=key_cols, min_show=1, top_n_files=200)
        except Exception as e:
            log.debug("[diag] diagnostics failed: %s", e)
    # ------------------------------------------------------------
    # HANDOFF to HeavyEngine (Model B) — controlled by config.PIPELINE
    # ------------------------------------------------------------
    pipe_cfg = getattr(config, "PIPELINE", {})
    if pipe_cfg.get("handoff_enabled", True):
        try:
            # 1) Write rows to the chosen handoff format
            fmt = pipe_cfg.get("handoff_format", "csv")
            handoff_path = export_rows_handoff(
                all_rows,
                path=config.OUTPUT.get("handoff_rows_path"),
                fmt=fmt
            )

            # 2) Optional post-handoff jobs driven by PIPELINE toggles
            #    (deep tracer, closure_builder, semantics_builder)
            try:
                # Re-read in case defaults differ here
                pipe_cfg = getattr(config, "PIPELINE", {}) or {}

                # 2.1) deep_recursive_engine (single-sheet tracer)
                if pipe_cfg.get("trigger_deep_recursive_engine", False):
                    import subprocess, sys, os
                    print("[HANDOFF] Launching deep_recursive_engine (single-sheet)...")
                    cmd = [
                        sys.executable,
                        os.path.join(os.path.dirname(__file__), "deep_recursive_engine.py"),
                        "--reuse-file", handoff_path,
                        "--excel-name", "deep_recursion_single.xlsx"
                    ]
                    proc = subprocess.Popen(
                        cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                        bufsize=1, universal_newlines=True
                    )
                    for line in proc.stdout:
                        print("[deep/sub] " + line.rstrip())
                    ret = proc.wait()
                    if ret != 0:
                        raise RuntimeError(f"deep_recursive_engine exited with code {ret}")
                    print("[HANDOFF] deep_recursive_engine completed successfully.")
                else:
                    log.info("[HANDOFF] trigger_deep_recursive_engine=False; skipping deep tracer.")

                # 2.2) Optional closure_builder step
                if pipe_cfg.get("trigger_closure_builder", False):
                    try:
                        import subprocess, sys, os
                        closure_path = os.path.join(os.path.dirname(__file__), "closure_builder.py")
                        if not os.path.isfile(closure_path):
                            raise FileNotFoundError(f"Missing: {closure_path}")
                        print("[HANDOFF] Launching closure_builder ...")
                        cmd = [sys.executable, closure_path, "--reuse-file", handoff_path]
                        subprocess.check_call(cmd)
                        print("[HANDOFF] closure_builder completed successfully.")
                    except Exception as ce:
                        print(f"[HANDOFF] closure_builder flow failed: {ce}")
                else:
                    log.info("[HANDOFF] trigger_closure_builder=False; skipping.")

                # 2.3) Optional semantics_builder step
                if pipe_cfg.get("trigger_semantics_builder", False):
                    try:
                        import subprocess, sys, os
                        sem_path = os.path.join(os.path.dirname(__file__), "semantics_builder.py")
                        if not os.path.isfile(sem_path):
                            raise FileNotFoundError(f"Missing: {sem_path}")
                        print("[HANDOFF] Launching semantics_builder ...")
                        cmd = [sys.executable, sem_path, "--reuse-file", handoff_path]
                        subprocess.check_call(cmd)
                        print("[HANDOFF] semantics_builder completed successfully.")
                    except Exception as se:
                        print(f"[HANDOFF] semantics_builder flow failed: {se}")
                else:
                    log.info("[HANDOFF] trigger_semantics_builder=False; skipping.")

            except Exception as de:
                print(f"[HANDOFF] deep/closure/semantics flow failed: {de}")

            # 3) Trigger heavy_engine as a separate process (optional)
            if pipe_cfg.get("trigger_heavy_engine", True):
                import subprocess
                import sys
                import shlex  # (kept if you later want to join/print the command)

                print("[HANDOFF] Launching HeavyEngine (streaming)...")

                cmd = [
                    sys.executable,
                    os.path.join(os.path.dirname(__file__), "heavy_engine.py"),
                    "--reuse-file", handoff_path,
                    "--excel-name", config.OUTPUT.get("heavy_excel_name", "column_lineage_heavy.xlsx"),
                ]

                timeout = int(getattr(config, "PIPELINE", {}).get("heavy_engine_timeout_sec", 0) or 0)

                try:
                    proc = subprocess.Popen(
                        cmd,
                        stdout=subprocess.PIPE,
                        stderr=subprocess.STDOUT,
                        bufsize=1,
                        universal_newlines=True
                    )
                    for line in proc.stdout:
                        print("[heavy/sub] " + line.rstrip())

                    ret = proc.wait(timeout=timeout if timeout > 0 else None)
                    if ret != 0:
                        raise RuntimeError(f"HeavyEngine exited with code {ret}")

                    print("[HANDOFF] HeavyEngine completed successfully.")

                except subprocess.TimeoutExpired:
                    proc.kill()
                    raise RuntimeError(f"HeavyEngine timed out after {timeout}s")

            else:
                log.info("[HANDOFF] trigger_heavy_engine=False; skipping HeavyEngine launch.")

        except Exception as he:
            log.warning(f"[HANDOFF] HeavyEngine flow failed: {he}")
    else:
        log.info("[HANDOFF] handoff_enabled=False; skipping handoff file and HeavyEngine.")

    # ============================================================
    # EXPORT FULL LINEAGE
    # ============================================================
    engine_cfg = getattr(config, "LINEAGE_ENGINE", {})
    extract_full = engine_cfg.get("extract_full", True)
    subset_only = engine_cfg.get("extract_subset_only", False)

    if extract_full and not subset_only:
        # Use the smart exporter instead of always writing Excel
        import pandas as pd
        df = pd.DataFrame(all_rows).fillna(NULL_TOKEN)
        out_dir = config.OUTPUT.get("output_dir", ".")
        stem = os.path.splitext(config.OUTPUT.get("excel_name", "column_lineage.xlsx"))[0]
        written = _export_dataframe_smart(df, stem, out_dir, config)
        log.info(f"[EXPORT] Written: {written}")
    else:
        log.info("[ENGINE] Full export skipped (subset-only mode or disabled).")

    # ============================================================
    # VISUALIZATION (IF NEEDED)
    # ============================================================
    viz = getattr(config, "VISUALIZATION", {}) or {}
    if viz.get("enabled", True) and viz.get("use_traced_subset", True):
        targets = viz.get("target_columns") or []
        if isinstance(targets, str):
            targets = [targets]
        if targets:
            subset_path = export_traced_subset(all_rows, targets)
            if subset_only:
                log.info("[ENGINE] Subset-only mode: stopping after subset export.")
                raise SystemExit(0)
        else:
            log.info("[TRACE] No target_columns specified for traced subset.")