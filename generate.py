"""Generate ICH mutation CSVs from the Maori Heritage Project spreadsheet.

Reads the structured spreadsheet, parses CRM field semantics, and emits
alizarin-compatible CSV files that:
  - Create new reusable branches (Appellative Status, Whakapapa, etc.)
  - Create new resource models (Tikanga, Event)
  - Add ICH fields to existing HER models (Modify)
  - Add Reference Mātauranga links to unchanged models (Same)

Usage:
    python ich/generate.py [--xlsx path] [--pkg-dir path] [--output-dir ich/]
"""

import argparse
import csv
import io
import json
import re
import uuid
from collections import OrderedDict, defaultdict
from pathlib import Path

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

SCRIPT_DIR = Path(__file__).parent
DEFAULT_XLSX = "Maori Heritage Project.xlsx"

# UUID namespace for deterministic generation of new branch/model UUIDs.
ICH_NAMESPACE = uuid.UUID("a1b2c3d4-e5f6-7890-abcd-ef1234567890")

CSV_HEADER = [
    "action", "subject", "object",
    "params.name", "params.datatype", "params.cardinality",
    "params.ontology_class", "params.parent_property",
    "params.description", "params.config",
    "params.ontology_property", "params.publication_id",
]

# Prefix → full URI mapping for CRM path parsing.
PREFIX_MAP = {
    "crm": "http://www.cidoc-crm.org/cidoc-crm/",
    "aaao": "https://ontology.swissartresearch.net/aaao/",
    "rdf": "http://www.w3.org/1999/02/22-rdf-syntax-ns#",
    "rdfs": "http://www.w3.org/2000/01/rdf-schema#",
    "xsd": "http://www.w3.org/2001/XMLSchema#",
    "xsl": "http://www.w3.org/2001/XMLSchema#",
    "crmdig": "http://www.ics.forth.gr/isl/CRMdig/",
    "crmsci": "http://www.cidoc-crm.org/extensions/crmsci/",
    "crmgeo": "http://www.cidoc-crm.org/extensions/crmgeo/",
    "mahero": "https://ontology.swissartresearch.net/mahero/",
}

# Known discrepancies between spreadsheet HER model names and actual HER names.
# Spreadsheet name → actual HER name.
HER_MODEL_NAME_ALIASES = {
    "Group": "Organization",
    "Site": "Area",
    "Consltation": "Consultation",  # typo in HER data
}

# Spreadsheet Field Type → alizarin datatype.
FIELD_TYPE_MAP = {
    "string": "string",
    "concept": "concept",
    "date": "date",
    "reference model": "resource-instance-list",
    "collection": "semantic",  # subgraph attachment
    "bngcentrepoint": "bngcentrepoint",
}

# Sheet name → HER model name mapping (from the Models sheet).
# Populated at runtime from the spreadsheet.
MODEL_SHEET_MAP = {}

# ---------------------------------------------------------------------------
# Spreadsheet parsing helpers
# ---------------------------------------------------------------------------

def open_workbook(xlsx_path):
    """Open the spreadsheet with openpyxl."""
    import openpyxl
    return openpyxl.load_workbook(str(xlsx_path), data_only=True)


def find_column_indices(header_row):
    """Dynamically map column names → indices from the header row.

    Returns a dict like {"order": 0, "collection": 1, "field_name": 17, ...}.
    Column names vary per sheet so we search for known keywords.
    """
    cols = {}
    for idx, cell_val in enumerate(header_row):
        if cell_val is None:
            continue
        val = str(cell_val).strip().lower()
        if val.startswith("order"):
            cols["order"] = idx
        elif val == "part of collection":
            cols["collection"] = idx
        elif val == "model specific collection name":
            cols["collection_name"] = idx
        elif val == "field name":
            cols["field_name"] = idx
        elif val == "field semantics":
            cols["field_semantics"] = idx
        elif val == "field type":
            cols["field_type"] = idx
        elif val.startswith("reference model") or val == "reference model / collection":
            cols["reference_model"] = idx
        elif val == "ontomap":
            cols["ontomap"] = idx
        elif val.startswith("new field"):
            cols["new_field"] = idx
        elif val == "cardinality":
            cols["cardinality"] = idx
        elif val == "field":
            cols["field_display"] = idx
        elif val == "category":
            cols["category"] = idx
    return cols


def parse_models_sheet(wb):
    """Parse the 'Models' sheet → list of model dicts.

    Each dict: {id, name, her_model, ontology, action, sheet_name}
    """
    ws = wb["Models"]
    models = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        model_id = row[0]
        if model_id is None:
            continue
        name = (row[1] or "").strip()
        her_model = (row[2] or "").strip() if row[2] else None
        ontology = (row[3] or "").strip() if row[3] else ""
        action = (row[4] or "").strip().lower() if row[4] else ""
        models.append({
            "id": str(model_id).strip(),
            "name": name,
            "her_model": her_model,
            "ontology": ontology,
            "action": action,  # "new", "modify", "same"
        })
    return models


def find_sheet_for_model(wb, model):
    """Find which sheet corresponds to a model, matching by model name or ontology.

    Uses a priority system: exact start match > HER model start match >
    substring match, to avoid e.g. 'Person' matching 'Whenua...Persons'.
    Also handles Excel's 31-char sheet name truncation.
    """
    model_name = model["name"].lower().strip()
    her_model = (model.get("her_model") or "").lower().strip()
    candidates = [
        sn for sn in wb.sheetnames
        if sn.lower().strip() not in ("models", "models descriptions")
    ]

    # Pass 1: sheet name starts with model name (best match)
    for sn in candidates:
        sl = sn.lower().strip()
        if model_name and sl.startswith(model_name):
            return sn

    # Pass 2: sheet name starts with HER model name
    for sn in candidates:
        sl = sn.lower().strip()
        if her_model and sl.startswith(her_model):
            return sn

    # Pass 3: model name starts with the sheet name (truncated sheet names)
    for sn in candidates:
        sl = sn.lower().strip()
        if model_name and model_name.startswith(sl):
            return sn
        if her_model and her_model.startswith(sl):
            return sn

    # Pass 4: substring match (model name in sheet name), prefer shortest sheet
    matches = []
    for sn in candidates:
        sl = sn.lower().strip()
        if model_name and model_name in sl:
            matches.append(sn)
        elif her_model and her_model in sl:
            matches.append(sn)
    if matches:
        return min(matches, key=len)

    return None


def parse_model_sheet(wb, sheet_name):
    """Parse a model sheet → list of field dicts.

    Each field dict:
      order, collection, collection_name, field_display, field_name,
      field_semantics, field_type, reference_model, ontomap,
      is_new, cardinality
    """
    ws = wb[sheet_name]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []

    cols = find_column_indices(rows[0])
    fields = []
    for row_idx, row in enumerate(rows[1:], start=2):
        vals = list(row)

        # Skip header-like / "Overall" rows
        order_val = vals[cols["order"]] if "order" in cols and cols["order"] < len(vals) else None
        if order_val is None:
            continue
        try:
            order = float(order_val)
        except (ValueError, TypeError):
            continue

        def g(key):
            if key in cols and cols[key] < len(vals):
                v = vals[cols[key]]
                return str(v).strip() if v is not None else None
            return None

        field_type_raw = g("field_type")
        field_type = field_type_raw.lower().strip() if field_type_raw else None
        is_new_raw = g("new_field")
        is_new = is_new_raw and is_new_raw.lower().strip() == "yes"

        cardinality_raw = g("cardinality")
        cardinality = None
        if cardinality_raw:
            cardinality_raw = cardinality_raw.strip()
            if "n" in cardinality_raw:
                cardinality = "n"
            else:
                cardinality = "1"

        fields.append({
            "order": order,
            "collection": g("collection"),
            "collection_name": g("collection_name"),
            "field_display": g("field_display"),
            "field_name": g("field_name"),
            "field_semantics": g("field_semantics"),
            "field_type": field_type,
            "reference_model": g("reference_model"),
            "ontomap": g("ontomap"),
            "is_new": is_new,
            "cardinality": cardinality,
        })
    return fields


# ---------------------------------------------------------------------------
# CRM path parsing
# ---------------------------------------------------------------------------

def expand_prefix(prefixed_name):
    """Expand 'crm:E55_Type' → full URI."""
    if ":" not in prefixed_name:
        return prefixed_name
    prefix, local = prefixed_name.split(":", 1)
    base = PREFIX_MAP.get(prefix.lower())
    if base:
        return base + local
    return prefixed_name


def parse_crm_path(semantics_str):
    """Parse a Field Semantics string into a list of path segments.

    Input format: ->prefix:Property->prefix:Class[alias]->...
    Also handles paths without leading -> (e.g. "crm:P53...->...")

    Returns list of dicts:
      [{"property": "http://...", "cls": "http://...", "alias": "5_1"}, ...]
    """
    if not semantics_str:
        return []

    # Normalise: strip leading whitespace, handle paths that don't start with ->
    s = semantics_str.strip()

    # The path is a series of ->property->class[alias] pairs.
    # Some paths have spaces around ->, normalise them.
    s = s.replace("-> ", "->").replace(" ->", "->")

    # Split on -> to get tokens
    parts = [p.strip() for p in s.split("->") if p.strip()]

    # Group into (property, class) pairs. Even indices = property, odd = class.
    # But the final node might be rdf:literal or xsl:date (a leaf type, not a class).
    segments = []
    i = 0
    while i < len(parts):
        prop_token = parts[i]
        cls_token = parts[i + 1] if i + 1 < len(parts) else None
        i += 2

        # Extract alias from class token: E55_Type[5_1] → cls=E55_Type, alias=5_1
        alias = None
        cls_raw = cls_token
        if cls_token and "[" in cls_token:
            m = re.match(r"(.+?)\[(.+?)\]", cls_token)
            if m:
                cls_raw = m.group(1)
                alias = normalise_alias(m.group(2))

        # Handle leaf types (rdf:literal, xsl:date, etc.) — these represent the
        # terminal data node, not a CRM class node
        if cls_raw and cls_raw.lower() in (
            "rdf:literal", "xsd:date", "xsl:date",
            "rdfs:literal",
        ):
            # This is a leaf — the property points to a literal
            segments.append({
                "property": expand_prefix(prop_token),
                "cls": expand_prefix(cls_raw),
                "alias": alias,
                "is_literal_leaf": True,
            })
        else:
            # Handle class tokens that contain "/" for multiple types
            # e.g. "crm:E22 Human-Made Object/aaao:ZE36 Persons"
            # Just take the first class for ontology purposes
            if cls_raw and "/" in cls_raw:
                # Split and take first
                first_cls = cls_raw.split("/")[0].strip()
                # Fix spaces in class names: "crm:E22 Human-Made Object" → "crm:E22_Human-Made_Object"
                first_cls = first_cls.replace(" ", "_")
                cls_raw = first_cls

            # Fix spaces in class names
            if cls_raw:
                # e.g. "crm:E22 Human-Made Object" → "crm:E22_Human-Made_Object"
                if ":" in cls_raw:
                    prefix_part, local_part = cls_raw.split(":", 1)
                    local_part = local_part.replace(" ", "_")
                    cls_raw = f"{prefix_part}:{local_part}"

            segments.append({
                "property": expand_prefix(prop_token),
                "cls": expand_prefix(cls_raw) if cls_raw else None,
                "alias": alias,
                "is_literal_leaf": False,
            })

    return segments


# ---------------------------------------------------------------------------
# UUID generation
# ---------------------------------------------------------------------------

def make_branch_uuid(collection_id):
    """Deterministic UUID for a new ICH branch, keyed by collection ID."""
    return str(uuid.uuid5(ICH_NAMESPACE, f"branch:{collection_id}"))


def make_model_uuid(model_id):
    """Deterministic UUID for a new ICH resource model."""
    return str(uuid.uuid5(ICH_NAMESPACE, f"model:{model_id}"))


def make_coppice_uuid(parent_context, node_name, branch_uuid):
    """Deterministic UUID for a coppiced (copied) subgraph instance."""
    return str(uuid.uuid5(ICH_NAMESPACE,
                          f"coppice:{parent_context}:{node_name}:{branch_uuid}"))


def resolve_collection_to_branch(ref_str, new_branch_map, existing_branch_map):
    """Resolve a collection reference like 'MORC.3_name' or 'HERC.79_simple_timespan'
    to a branch UUID.

    Checks new ICH branches first, then existing HER branches.
    """
    if not ref_str:
        return None
    ref = ref_str.strip()
    # Try direct match in new branches
    if ref in new_branch_map:
        return new_branch_map[ref]
    # Try in existing branches
    if ref in existing_branch_map:
        return existing_branch_map[ref]
    # Try the part after the underscore (e.g. HERC.79_simple_timespan → simple_timespan)
    parts = ref.split("_", 1)
    if len(parts) > 1:
        suffix = parts[1]
        # Prefer exact suffix match, then shortest key match
        for map_ in (new_branch_map, existing_branch_map):
            exact = [v for k, v in map_.items() if k.endswith(f"_{suffix}")]
            if exact:
                return exact[0]
    return f"TODO:branch_uuid_for_{ref}"


# ---------------------------------------------------------------------------
# CSV writing
# ---------------------------------------------------------------------------

# Known ontology URI typos from the spreadsheet → canonical form
_ONTOLOGY_URI_FIXES = {
    "http://www.cidoc-crm.org/cidoc-crm/E52_Time-span":
        "http://www.cidoc-crm.org/cidoc-crm/E52_Time-Span",
    "http://www.cidoc-crm.org/cidoc-crm/E22_Human-Made_Object":
        "http://www.cidoc-crm.org/cidoc-crm/E22_Human-Made_Object",
}


def csv_row(action, subject="", obj="", name="", datatype="",
            cardinality="", ontology_class="", parent_property="",
            description="", config="", ontology_property="",
            publication_id=""):
    """Build a single CSV row as a list."""
    ontology_class = _ONTOLOGY_URI_FIXES.get(ontology_class, ontology_class)
    return [
        action, subject, obj,
        name, datatype, cardinality,
        ontology_class, parent_property,
        description, config,
        ontology_property, publication_id,
    ]


def write_csv(filepath, rows):
    """Write rows (list of lists) to a CSV file with our standard header."""
    with open(filepath, "w", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(CSV_HEADER)
        for row in rows:
            writer.writerow(row)
    print(f"  Wrote {filepath.name} ({len(rows)} rows)")


# ---------------------------------------------------------------------------
# Node name generation
# ---------------------------------------------------------------------------

def sanitise_node_name(name):
    """Turn a display name into a valid alizarin node name (snake_case, lowercase)."""
    s = name.strip().lower()
    s = re.sub(r"[^a-z0-9]+", "_", s)
    s = s.strip("_")
    return s


def normalise_alias(alias):
    """Normalise CRM path aliases for consistent node naming.

    Handles spreadsheet inconsistencies like SRDF518_1 vs SRDF.518_1
    by ensuring a consistent separator pattern.
    """
    if not alias:
        return alias
    # Insert dot after letter prefix if missing:
    # SRDF518_1 → SRDF.518_1, but leave SRDF.518_1 alone
    normalised = re.sub(r"^([A-Za-z]+)(\d)", r"\1.\2", alias)
    return normalised


def make_node_name_from_alias(alias, prefix=""):
    """Generate a node name from an alias hint like SRDF.524_1."""
    if alias:
        s = alias.lower().replace(".", "_")
        if prefix:
            return f"{prefix}_{s}"
        return s
    return None


# ---------------------------------------------------------------------------
# Reference model config generation
# ---------------------------------------------------------------------------

# Known model UUIDs from HER — populated from --pkg-dir at runtime.
HER_MODEL_UUIDS = {}

# MORM model UUIDs — populated at generation time for new models.
MORM_MODEL_UUIDS = {}


def resolve_model_uuid(ref_str):
    """Resolve a reference model string like 'MORM.9_Person' to its UUID.

    Returns a list of (graphid, model_name) tuples.
    """
    if not ref_str:
        return []
    results = []
    # Can be comma-separated or newline-separated
    refs = re.split(r"[,\n]+", ref_str)
    for ref in refs:
        ref = ref.strip()
        if not ref:
            continue
        # Try MORM reference first
        morm_match = re.match(r"(MORM\.\d+)(?:_(.+))?", ref)
        if morm_match:
            morm_id = morm_match.group(1)
            model_name = (morm_match.group(2) or "").strip()
            if morm_id in MORM_MODEL_UUIDS:
                results.append((MORM_MODEL_UUIDS[morm_id], model_name))
            else:
                # Fallback — try to find by name in HER
                if model_name and model_name in HER_MODEL_UUIDS:
                    results.append((HER_MODEL_UUIDS[model_name], model_name))
                else:
                    results.append((f"TODO:{morm_id}", model_name))
        else:
            # Plain model name — try exact, alias, then fuzzy matching
            if ref in HER_MODEL_UUIDS:
                results.append((HER_MODEL_UUIDS[ref], ref))
            elif ref in HER_MODEL_NAME_ALIASES and HER_MODEL_NAME_ALIASES[ref] in HER_MODEL_UUIDS:
                results.append((HER_MODEL_UUIDS[HER_MODEL_NAME_ALIASES[ref]], ref))
            else:
                # Fuzzy: ref might include ontology scope
                # e.g. "Mātauranga E33 Linguistic Object" → starts with "Mātauranga"
                found = False
                for name, uid in HER_MODEL_UUIDS.items():
                    if ref.startswith(name + " ") or name.startswith(ref + " "):
                        results.append((uid, ref))
                        found = True
                        break
                if not found:
                    results.append((f"TODO:{ref}", ref))
    return results


def build_resource_instance_config(ref_str):
    """Build params.config JSON for resource-instance-list fields."""
    models = resolve_model_uuid(ref_str)
    if not models:
        return ""
    graphs = []
    for graphid, _name in models:
        graphs.append({
            "graphid": graphid,
            "ontologyProperty": "",
            "inverseOntologyProperty": "",
            "useOntologyRelationship": False,
        })
    return json.dumps({"graphs": graphs, "searchDsl": "", "searchString": ""})


# ---------------------------------------------------------------------------
# Branch building
# ---------------------------------------------------------------------------

def build_branch_rows(collection_id, collection_name, fields, branch_uuid,
                      new_branch_map=None, existing_branch_map=None):
    """Build CSV rows for a new branch from a set of fields sharing a collection.

    For Collection-type fields, emits add_subgraph (or coppice_subgraph for
    repeated use of the same sub-branch) to attach the referenced branch.

    Returns list of CSV row lists.
    """
    new_branch_map = new_branch_map or {}
    existing_branch_map = existing_branch_map or {}
    rows = []
    safe_name = sanitise_node_name(collection_name or collection_id)
    rows.append(csv_row("create_branch", safe_name, branch_uuid))

    # Track emitted nodes and subgraph attachments within this branch.
    emitted_nodes = set()
    # branch_uuid → first node name where it was attached (for coppice tracking)
    attached_subgraphs = {}

    sorted_fields = sorted(fields, key=lambda f: f["order"])

    for field in sorted_fields:
        semantics = field.get("field_semantics")
        if not semantics:
            continue
        segments = parse_crm_path(semantics)
        if not segments:
            continue

        field_type = field.get("field_type", "")
        ref_model = field.get("reference_model")
        cardinality = field.get("cardinality")
        display_name = field.get("field_name") or field.get("field_display") or ""

        # Walk path segments. All but the last are intermediate (semantic) nodes.
        parent = safe_name  # branch root
        for seg_idx, seg in enumerate(segments):
            is_last = (seg_idx == len(segments) - 1)
            is_literal_leaf = seg.get("is_literal_leaf", False)

            alias = seg.get("alias")
            if alias:
                node_name = sanitise_node_name(alias)
            elif seg.get("cls"):
                cls_local = seg["cls"].rsplit("/", 1)[-1] if "/" in seg["cls"] else seg["cls"].rsplit("#", 1)[-1]
                node_name = sanitise_node_name(cls_local)
            else:
                node_name = f"node_{seg_idx}"

            full_node_key = f"{parent}_{node_name}"

            if is_last and not is_literal_leaf:
                if full_node_key not in emitted_nodes:
                    emitted_nodes.add(full_node_key)
                    dt = FIELD_TYPE_MAP.get(field_type, field_type or "string")
                    if field_type == "collection":
                        dt = "semantic"

                    card = cardinality or ("n" if dt == "semantic" else "1")
                    config = ""
                    if dt == "resource-instance-list" and ref_model:
                        config = build_resource_instance_config(ref_model)

                    rows.append(csv_row(
                        "add_node", parent, node_name,
                        name=display_name,
                        datatype=dt,
                        cardinality=card,
                        ontology_class=seg.get("cls", ""),
                        parent_property=seg.get("property", ""),
                        config=config,
                    ))

                    # Collection fields: attach the referenced branch as a subgraph
                    if field_type == "collection" and ref_model:
                        sub_uuid = resolve_collection_to_branch(
                            ref_model, new_branch_map, existing_branch_map)
                        if sub_uuid in attached_subgraphs:
                            # Same branch already attached elsewhere in this graph
                            # → coppice the first attachment to create a new copy
                            first_node = attached_subgraphs[sub_uuid]
                            new_uuid = make_coppice_uuid(
                                collection_id, node_name, sub_uuid)
                            rows.append(csv_row(
                                "coppice_subgraph", first_node,
                                publication_id=new_uuid,
                            ))
                        else:
                            rows.append(csv_row(
                                "add_subgraph", node_name, sub_uuid))
                            attached_subgraphs[sub_uuid] = node_name

                parent = node_name

            elif is_literal_leaf:
                if full_node_key not in emitted_nodes:
                    emitted_nodes.add(full_node_key)
                    dt = FIELD_TYPE_MAP.get(field_type, "string")
                    card = cardinality or "1"
                    rows.append(csv_row(
                        "add_node", parent, node_name,
                        name=display_name,
                        datatype=dt,
                        cardinality=card,
                        ontology_class=seg.get("cls", ""),
                        parent_property=seg.get("property", ""),
                    ))
                parent = node_name

            else:
                # Intermediate semantic node
                if full_node_key not in emitted_nodes:
                    emitted_nodes.add(full_node_key)
                    rows.append(csv_row(
                        "add_node", parent, node_name,
                        name=display_name if is_last else node_name.replace("_", " ").title(),
                        datatype="semantic",
                        cardinality="n",
                        ontology_class=seg.get("cls", ""),
                        parent_property=seg.get("property", ""),
                    ))
                parent = node_name

    return rows


# ---------------------------------------------------------------------------
# Model modification building
# ---------------------------------------------------------------------------

def build_model_modification_rows(model, fields, new_branch_map,
                                   existing_branch_map=None, her_graph_uuid=None):
    """Build CSV rows for modifying an existing HER model.

    - load_graph with HER UUID
    - add_subgraph for each new ICH branch
    - add_node for standalone new fields (not in a collection)
    """
    existing_branch_map = existing_branch_map or {}
    rows = []

    if her_graph_uuid:
        rows.append(csv_row("load_graph", her_graph_uuid))
    else:
        rows.append(csv_row("load_graph", f"TODO:uuid_for_{model['her_model'] or model['name']}"))

    # Determine root node name for the model (apply aliases for known typos).
    her_name = model.get("her_model") or model["name"]
    her_name = HER_MODEL_NAME_ALIASES.get(her_name, her_name)
    root_name = sanitise_node_name(her_name)

    # Track which new branches to attach (by collection_id).
    attached_branches = set()
    standalone_new_fields = []

    for field in fields:
        if not field.get("is_new"):
            continue
        coll = field.get("collection")
        if coll and coll in new_branch_map:
            if coll not in attached_branches:
                attached_branches.add(coll)
                branch_uuid = new_branch_map[coll]
                rows.append(csv_row("add_subgraph", root_name, branch_uuid))
        else:
            standalone_new_fields.append(field)

    # Standalone new fields — add directly to model root
    for field in standalone_new_fields:
        semantics = field.get("field_semantics")
        if not semantics:
            continue
        segments = parse_crm_path(semantics)
        if not segments:
            continue

        _add_path_nodes(rows, root_name, segments, field,
                        new_branch_map=new_branch_map,
                        existing_branch_map=existing_branch_map)

    return rows


def build_new_model_rows(model, fields, new_branch_map, existing_branch_map):
    """Build CSV rows for a new resource model (Tikanga, Event).

    - create_model
    - add_subgraph for existing HER branches
    - add_subgraph for new ICH branches
    - add_node for standalone fields
    """
    rows = []
    model_uuid = MORM_MODEL_UUIDS.get(model["id"], make_model_uuid(model["id"]))

    safe_name = sanitise_node_name(model["name"])
    # Determine ontology class for the root
    ontology_scope = model.get("ontology", "")
    root_cls = ""
    if ontology_scope:
        # e.g. "ZE13 Speech Act" → try to expand
        # Look for known prefix matches
        if any(ontology_scope.startswith(p) for p in ("E", "D", "Z")):
            # Guess CRM class
            cls_name = ontology_scope.replace(" ", "_").replace("/", "_")
            if cls_name.startswith("Z"):
                root_cls = expand_prefix(f"aaao:{cls_name}")
            elif cls_name.startswith("D"):
                root_cls = expand_prefix(f"crmdig:{cls_name}")
            else:
                root_cls = expand_prefix(f"crm:{cls_name}")

    rows.append(csv_row("create_model", safe_name, model_uuid,
                         name=model["name"], ontology_class=root_cls))

    # Group fields by collection
    coll_fields = defaultdict(list)
    standalone_fields = []
    for field in sorted(fields, key=lambda f: f["order"]):
        coll = field.get("collection")
        if coll:
            coll_fields[coll].append(field)
        else:
            standalone_fields.append(field)

    # Attach existing HER branches
    attached_branches = set()
    for coll_id, coll_fields_list in coll_fields.items():
        # Check if this is an existing HER collection (not new)
        any_new = any(f.get("is_new") for f in coll_fields_list)
        if not any_new and coll_id in existing_branch_map:
            if coll_id not in attached_branches:
                attached_branches.add(coll_id)
                rows.append(csv_row("add_subgraph", safe_name,
                                    existing_branch_map[coll_id]))
        elif coll_id in new_branch_map:
            if coll_id not in attached_branches:
                attached_branches.add(coll_id)
                rows.append(csv_row("add_subgraph", safe_name,
                                    new_branch_map[coll_id]))
        elif coll_id in existing_branch_map:
            if coll_id not in attached_branches:
                attached_branches.add(coll_id)
                rows.append(csv_row("add_subgraph", safe_name,
                                    existing_branch_map[coll_id]))
        else:
            # Unknown collection — emit as branch attachment with TODO
            if coll_id not in attached_branches:
                attached_branches.add(coll_id)
                rows.append(csv_row("add_subgraph", safe_name,
                                    f"TODO:branch_uuid_for_{coll_id}"))

    # Standalone fields
    for field in standalone_fields:
        semantics = field.get("field_semantics")
        if not semantics:
            continue
        segments = parse_crm_path(semantics)
        if not segments:
            continue
        _add_path_nodes(rows, safe_name, segments, field,
                        new_branch_map=new_branch_map,
                        existing_branch_map=existing_branch_map)

    return rows


def _add_path_nodes(rows, parent, segments, field,
                    new_branch_map=None, existing_branch_map=None):
    """Add CRM path nodes to rows for a standalone field.

    For Collection-type fields, emits add_subgraph after the semantic node.
    """
    new_branch_map = new_branch_map or {}
    existing_branch_map = existing_branch_map or {}
    current_parent = parent
    for seg_idx, seg in enumerate(segments):
        is_last = seg_idx == len(segments) - 1
        is_literal_leaf = seg.get("is_literal_leaf", False)

        alias = seg.get("alias")
        if alias:
            node_name = sanitise_node_name(alias)
        elif seg.get("cls"):
            cls_local = seg["cls"].rsplit("/", 1)[-1] if "/" in seg["cls"] else seg["cls"].rsplit("#", 1)[-1]
            node_name = sanitise_node_name(cls_local)
        else:
            node_name = f"node_{seg_idx}"

        if is_last and not is_literal_leaf:
            field_type = field.get("field_type", "")
            dt = FIELD_TYPE_MAP.get(field_type, field_type or "string")
            if field_type == "collection":
                dt = "semantic"
            card = field.get("cardinality") or ("n" if dt == "semantic" else "1")
            config = ""
            if dt == "resource-instance-list" and field.get("reference_model"):
                config = build_resource_instance_config(field["reference_model"])
            display_name = field.get("field_name") or field.get("field_display") or ""
            rows.append(csv_row(
                "add_node", current_parent, node_name,
                name=display_name, datatype=dt, cardinality=card,
                ontology_class=seg.get("cls", ""),
                parent_property=seg.get("property", ""),
                config=config,
            ))
            # Collection fields: attach referenced branch as subgraph
            if field_type == "collection" and field.get("reference_model"):
                sub_uuid = resolve_collection_to_branch(
                    field["reference_model"], new_branch_map, existing_branch_map)
                if sub_uuid:
                    rows.append(csv_row("add_subgraph", node_name, sub_uuid))
        elif is_literal_leaf:
            dt = FIELD_TYPE_MAP.get(field.get("field_type", ""), "string")
            card = field.get("cardinality") or "1"
            display_name = field.get("field_name") or field.get("field_display") or ""
            rows.append(csv_row(
                "add_node", current_parent, node_name,
                name=display_name, datatype=dt, cardinality=card,
                ontology_class=seg.get("cls", ""),
                parent_property=seg.get("property", ""),
            ))
        else:
            # Intermediate semantic node
            rows.append(csv_row(
                "add_node", current_parent, node_name,
                name=node_name.replace("_", " ").title(),
                datatype="semantic", cardinality="n",
                ontology_class=seg.get("cls", ""),
                parent_property=seg.get("property", ""),
            ))
        current_parent = node_name


# ---------------------------------------------------------------------------
# HER package scanning
# ---------------------------------------------------------------------------

def scan_pkg_graphs(pkg_dir):
    """Scan existing HER graph JSONs to build UUID lookup tables.

    Returns:
      her_models: dict of model_name → graph_uuid
      her_branches: dict of branch_name → graph_uuid
      branch_by_collection_hint: dict of collection_id_guess → graph_uuid
    """
    her_models = {}
    her_branches = {}
    branch_by_node = {}

    if not pkg_dir or not pkg_dir.exists():
        return her_models, her_branches, branch_by_node

    graphs_dir = pkg_dir / "graphs"
    for section in ("resource_models", "branches"):
        section_dir = graphs_dir / section
        if not section_dir.exists():
            continue
        for fp in sorted(section_dir.glob("*.json")):
            try:
                data = json.loads(fp.read_text())
                graph = data.get("graph", [{}])[0] if "graph" in data else data
                graph_id = graph.get("graphid", "")
                name = graph.get("name", "")
                if isinstance(name, dict):
                    name = name.get("en", next(iter(name.values()), ""))

                if section == "resource_models":
                    her_models[name] = graph_id
                    her_models[name.replace(" ", "_").lower()] = graph_id
                else:
                    her_branches[name] = graph_id
                    her_branches[name.lower()] = graph_id
                    her_branches[name.replace(" ", "_").lower()] = graph_id
                    # Try to extract a root node name as collection hint
                    nodes = graph.get("nodes", [])
                    root_node = graph.get("root", {})
                    if root_node:
                        root_name = root_node.get("name", "")
                        if root_name:
                            branch_by_node[root_name.lower()] = graph_id
            except (json.JSONDecodeError, IndexError, KeyError):
                continue

    return her_models, her_branches, branch_by_node


# ---------------------------------------------------------------------------
# Reference Mātauranga link
# ---------------------------------------------------------------------------

MATAURANGA_SEMANTICS = "->crm:P129i_is_subject_of->crm:E33_Linguistic_Object[292_1]"
MATAURANGA_FIELD_NAME = "Source Reference Work"
MATAURANGA_REF = "MORM.1_Mātauranga"


def build_matauranga_ref_row(parent_name):
    """Build an add_node row for the Reference Mātauranga link."""
    segments = parse_crm_path(MATAURANGA_SEMANTICS)
    if not segments:
        return None
    seg = segments[-1]
    node_name = sanitise_node_name(seg.get("alias", "292_1"))
    config = build_resource_instance_config(MATAURANGA_REF)
    return csv_row(
        "add_node", parent_name, node_name,
        name="Reference Mātauranga",
        datatype="resource-instance-list",
        cardinality="1",
        ontology_class=seg.get("cls", ""),
        parent_property=seg.get("property", ""),
        config=config,
    )


# ---------------------------------------------------------------------------
# Main generation logic
# ---------------------------------------------------------------------------

def generate(xlsx_path, pkg_dir=None, output_dir=None):
    """Main entry point: parse spreadsheet, generate all CSVs."""
    if output_dir is None:
        output_dir = SCRIPT_DIR
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    wb = open_workbook(xlsx_path)

    # --- 1. Parse Models sheet ---
    models = parse_models_sheet(wb)
    print(f"Found {len(models)} models")

    new_models = [m for m in models if m["action"] == "new"]
    modify_models = [m for m in models if m["action"] == "modify"]
    same_models = [m for m in models if m["action"] == "same"]
    print(f"  New: {len(new_models)}, Modify: {len(modify_models)}, Same: {len(same_models)}")

    # --- 2. Scan HER package for existing UUIDs ---
    her_models, her_branches, branch_by_node = scan_pkg_graphs(
        Path(pkg_dir) if pkg_dir else None
    )
    HER_MODEL_UUIDS.update(her_models)
    print(f"  HER models: {len(her_models)}, HER branches: {len(her_branches)}")

    # --- 3. Assign UUIDs to new MORM models ---
    for model in models:
        morm_id = model["id"]  # e.g. "MORM.1"
        if model["action"] == "new":
            MORM_MODEL_UUIDS[morm_id] = make_model_uuid(morm_id)
        else:
            # Try HER model name, then aliases, then MORM model name
            resolved = False
            for candidate in [
                model.get("her_model"),
                HER_MODEL_NAME_ALIASES.get(model.get("her_model", "")),
                model["name"],
                HER_MODEL_NAME_ALIASES.get(model["name"]),
            ]:
                if candidate and candidate in her_models:
                    MORM_MODEL_UUIDS[morm_id] = her_models[candidate]
                    resolved = True
                    break
            if not resolved and model.get("her_model"):
                print(f"  WARNING: Could not resolve HER model for {morm_id} "
                      f"({model['her_model']})")

    # Register MORM model names (Māori names and HER aliases) in HER_MODEL_UUIDS
    # so cross-model references like "Group" or "Mātauranga" resolve correctly.
    for model in models:
        morm_id = model["id"]
        if morm_id not in MORM_MODEL_UUIDS:
            continue
        uid = MORM_MODEL_UUIDS[morm_id]
        HER_MODEL_UUIDS[model["name"]] = uid
        if model.get("her_model"):
            HER_MODEL_UUIDS[model["her_model"]] = uid

    # --- 4. Parse all model sheets and collect new-field collections ---
    all_model_fields = {}  # model_id → [fields]
    new_collection_fields = defaultdict(list)  # collection_id → [fields from all sheets]
    existing_collections_used = set()  # collection IDs used by existing (non-new) fields
    # Map collection_id → display name(s) from the spreadsheet for branch matching.
    collection_display_names = {}  # coll_id → set of display names

    for model in models:
        sheet_name = find_sheet_for_model(wb, model)
        if not sheet_name:
            print(f"  WARNING: No sheet found for {model['id']} ({model['name']})")
            continue
        fields = parse_model_sheet(wb, sheet_name)
        all_model_fields[model["id"]] = fields
        print(f"  {model['id']} ({model['name']}): {sheet_name} → {len(fields)} fields")

        for field in fields:
            coll = field.get("collection")
            if not coll:
                continue
            if field.get("is_new"):
                new_collection_fields[coll].append(field)
            else:
                existing_collections_used.add(coll)
            # Track display names for this collection
            coll_name = field.get("collection_name")
            if coll_name:
                collection_display_names.setdefault(coll, set()).add(coll_name)

    # --- 5. Assign UUIDs to new branches ---
    # Deduplicate: same collection_id across sheets means same branch.
    new_branch_map = {}  # collection_id → branch_uuid
    sorted_collections = sorted(
        new_collection_fields.items(),
        key=lambda kv: min(f["order"] for f in kv[1]) if kv[1] else 0,
    )
    for coll_id, _fields in sorted_collections:
        new_branch_map[coll_id] = make_branch_uuid(coll_id)

    # --- 6. Build existing branch map (collection_id → UUID) ---
    # Match spreadsheet collection IDs to HER branch UUIDs by display name.
    # Build a reverse lookup: lowered branch name → UUID for fuzzy matching.
    _branch_lower = {k.lower(): v for k, v in her_branches.items()}
    existing_branch_map = {}
    for coll_id in existing_collections_used:
        matched = False
        display_names = collection_display_names.get(coll_id, set())

        # Pass 1: exact display name match
        for dname in display_names:
            if dname in her_branches:
                existing_branch_map[coll_id] = her_branches[dname]
                matched = True
                break
            if dname.lower() in _branch_lower:
                existing_branch_map[coll_id] = _branch_lower[dname.lower()]
                matched = True
                break
        if matched:
            continue

        # Pass 2: HER branch name contains display name or vice versa
        # e.g. "Event Names" → find branch "Names"
        for dname in display_names:
            dl = dname.lower()
            for bname_lower, buuid in _branch_lower.items():
                if bname_lower in dl or dl in bname_lower:
                    existing_branch_map[coll_id] = buuid
                    matched = True
                    break
            if matched:
                break
        if matched:
            continue

        # Pass 3: collection ID suffix → branch name
        parts = coll_id.split("_", 1)
        if len(parts) > 1:
            name_part = parts[1]
            title_name = name_part.replace("_", " ").title()
            if title_name in her_branches:
                existing_branch_map[coll_id] = her_branches[title_name]
                continue
            if name_part.lower() in _branch_lower:
                existing_branch_map[coll_id] = _branch_lower[name_part.lower()]
                continue
            # Partial match on suffix — prefer shortest branch name
            np_lower = name_part.lower()
            candidates = [
                (bn, bu) for bn, bu in _branch_lower.items()
                if np_lower in bn
            ]
            if candidates:
                best = min(candidates, key=lambda x: len(x[0]))
                existing_branch_map[coll_id] = best[1]
                continue

        existing_branch_map[coll_id] = f"TODO:branch_uuid_for_{coll_id}"

    # --- 7. Generate branch CSV files ---
    branch_file_num = 1
    for coll_id, fields in sorted_collections:
        branch_uuid = new_branch_map[coll_id]

        # Use the collection name from the first field
        coll_name = None
        for f in fields:
            if f.get("collection_name"):
                coll_name = f["collection_name"]
                break
        if not coll_name:
            coll_name = coll_id

        # Deduplicate fields by field_semantics (same path = same node)
        seen_semantics = set()
        unique_fields = []
        for f in sorted(fields, key=lambda x: x["order"]):
            key = f.get("field_semantics", "")
            if key and key not in seen_semantics:
                seen_semantics.add(key)
                unique_fields.append(f)
            elif not key:
                unique_fields.append(f)

        branch_rows = build_branch_rows(
            coll_id, coll_name, unique_fields, branch_uuid,
            new_branch_map=new_branch_map,
            existing_branch_map=existing_branch_map,
        )
        filename = f"{branch_file_num:02d}_{sanitise_node_name(coll_id)}.csv"
        write_csv(output_dir / filename, branch_rows)
        branch_file_num += 1

    # --- 8. Generate new model CSVs ---
    model_file_num = 20
    for model in new_models:
        model_id = model["id"]
        fields = all_model_fields.get(model_id, [])
        if not fields:
            print(f"  WARNING: No fields for new model {model_id}")
            continue
        model_rows = build_new_model_rows(model, fields, new_branch_map, existing_branch_map)
        safe_name = sanitise_node_name(model["name"])
        filename = f"{model_file_num}_{safe_name}_model.csv"
        write_csv(output_dir / filename, model_rows)
        model_file_num += 1

    # --- 9. Generate modification CSVs for Modify models ---
    mod_file_num = 30
    for model in modify_models:
        model_id = model["id"]
        fields = all_model_fields.get(model_id, [])
        new_fields = [f for f in fields if f.get("is_new")]
        if not new_fields:
            print(f"  INFO: No new fields for modify model {model_id} ({model['name']})")
            continue

        her_uuid = MORM_MODEL_UUIDS.get(model_id)
        mod_rows = build_model_modification_rows(
            model, fields, new_branch_map,
            existing_branch_map=existing_branch_map,
            her_graph_uuid=her_uuid,
        )
        safe_name = sanitise_node_name(model.get("her_model") or model["name"])
        filename = f"{mod_file_num}_{safe_name}_modifications.csv"
        write_csv(output_dir / filename, mod_rows)
        mod_file_num += 1

    # --- 10. Generate Reference Mātauranga CSV for "Same" models only ---
    # Modify/New models already include the Mātauranga ref in their own CSVs.
    mat_file_num = 40

    for model in same_models:
        model_id = model["id"]

        her_uuid = MORM_MODEL_UUIDS.get(model_id)
        if not her_uuid:
            her_uuid = f"TODO:uuid_for_{model.get('her_model') or model['name']}"

        her_name = model.get("her_model") or model["name"]
        # Apply aliases to fix known typos (e.g. Consltation → Consultation)
        her_name = HER_MODEL_NAME_ALIASES.get(her_name, her_name)
        root_name = sanitise_node_name(her_name)

        rows = []
        rows.append(csv_row("load_graph", her_uuid))
        mat_row = build_matauranga_ref_row(root_name)
        if mat_row:
            rows.append(mat_row)
        safe_name = sanitise_node_name(her_name)
        filename = f"{mat_file_num}_{safe_name}_matauranga_ref.csv"
        write_csv(output_dir / filename, rows)
        mat_file_num += 1

    # --- 11. Generate collection assignment CSV ---
    # Placeholder for clm.reference_change_collection rows.
    # These require knowing the node UUIDs after graph building,
    # so we emit stubs with TODO markers for now.
    coll_rows = []
    coll_rows.append(csv_row(
        "# Collection assignments for ICH concept/reference fields",
    ))
    coll_rows.append(csv_row(
        "# These rows require node UUIDs from the built graphs.",
    ))
    coll_rows.append(csv_row(
        "# Run once after initial graph build and fill in UUIDs.",
    ))
    write_csv(output_dir / "60_collections.csv", coll_rows)

    print(f"\nDone. Generated CSVs in {output_dir}/")


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Generate ICH mutation CSVs from Maori Heritage Project spreadsheet"
    )
    parser.add_argument(
        "--xlsx", default=str(DEFAULT_XLSX),
        help="Path to Maori Heritage Project.xlsx",
    )
    parser.add_argument(
        "--pkg-dir", default=None,
        help="Path to arches_her/pkg directory for existing graph UUIDs",
    )
    parser.add_argument(
        "--output-dir", default=str(SCRIPT_DIR),
        help="Output directory for generated CSVs (default: ich/)",
    )
    args = parser.parse_args()

    generate(
        xlsx_path=args.xlsx,
        pkg_dir=args.pkg_dir,
        output_dir=args.output_dir,
    )


if __name__ == "__main__":
    main()
