"""Build ICH-extended Arches HER graphs from CSV mutations.

Loads the arches-her package (branches + resource models), then applies
ICH mutation CSVs on top to produce extended/new graphs.

Usage:
    python run.py --pkg-dir path/to/arches_her/pkg \
        [--output-dir output/]
        [--no-ontology-validation]
"""

import argparse
import json
import os
import re
import urllib.request
from pathlib import Path

import alizarin
import alizarin_clm  # noqa: F401 — registers reference datatype + widgets
from alizarin import RustRdmCache

SCRIPT_DIR = Path(__file__).parent
BASE_URI = "https://doc.govt.nz/"
ONTOLOGY_CACHE_DIR = SCRIPT_DIR / ".cache" / "ontologies"

# Remote ontology sources: (cache_subdir, {filename: raw_url, ...})
_GH_RAW = "https://raw.githubusercontent.com"
REMOTE_ONTOLOGIES = [
    (
        "aaao",
        {
            "AAAo_v2.0.1.rdf": (
                f"{_GH_RAW}/swiss-art-research-net/aaao/"
                "refs/heads/main/Serialization/AAAo_v2.0.1.rdf"
            ),
        },
    ),
    (
        "linkedart",
        {
            "ontology_config.json": (
                f"{_GH_RAW}/thegetty/linkedart_ontology/"
                "refs/heads/master/linked_art/ontology_config.json"
            ),
            "cidoc.xml": (
                f"{_GH_RAW}/thegetty/linkedart_ontology/"
                "refs/heads/master/linked_art/cidoc.xml"
            ),
            "linkedart.xml": (
                f"{_GH_RAW}/thegetty/linkedart_ontology/"
                "refs/heads/master/linked_art/linkedart.xml"
            ),
            "linkedart_crm_enhancements.xml": (
                f"{_GH_RAW}/thegetty/linkedart_ontology/"
                "refs/heads/master/linked_art/linkedart_crm_enhancements.xml"
            ),
        },
    ),
]


def fetch_remote_ontologies():
    """Download remote ontology files to local cache if not already present."""
    for subdir, files in REMOTE_ONTOLOGIES:
        cache_dir = ONTOLOGY_CACHE_DIR / subdir
        for filename, url in files.items():
            cached = cache_dir / filename
            if not cached.exists():
                print(f"  Downloading {subdir}/{filename}")
                cached.parent.mkdir(parents=True, exist_ok=True)
                urllib.request.urlretrieve(url, cached)


def load_rdm_cache(base_dir, subdirs=None):
    """Load SKOS XML from concept/collection directories."""
    if subdirs is None:
        subdirs = ("reference_data/concepts", "reference_data/controlled_lists")
    cache = alizarin.get_global_rdm_cache()
    for subdir in subdirs:
        folder = base_dir / subdir
        if folder.is_dir():
            for fp in sorted(folder.glob("*.xml")):
                cache.add_from_skos_xml(fp.read_text(), BASE_URI)
    alizarin.set_global_rdm_cache(cache)


def load_pkg_graphs(pkg_dir):
    """Register all branch and resource model graphs from pkg directory.

    Auto-publishes branches that lack a publication.publicationid so they
    can be referenced via add_subgraph.
    """
    graph_ids = []
    graphs_dir = pkg_dir / "graphs"
    for section in ("branches", "resource_models"):
        section_dir = graphs_dir / section
        if not section_dir.exists():
            continue
        for fp in sorted(section_dir.glob("*.json")):
            raw = fp.read_text()
            data = json.loads(raw)
            graph = data.get("graph", [data])[0] if isinstance(data.get("graph"), list) else data
            # Auto-publish branches missing publication.publicationid
            pub = graph.get("publication") or {}
            if section == "branches" and not pub.get("publicationid"):
                graph["publication"] = {**pub, "publicationid": graph.get("graphid", "")}
                raw = json.dumps(data)
            graph_id = alizarin.register_graph(raw)
            print(f"  Registered {section}: {fp.name}")
            graph_ids.append(graph_id)
    return graph_ids


def export_graphs(graph_ids, out_dir):
    """Export all tracked graphs to output directory (last registration wins)."""
    seen = set()
    unique_ids = []
    for graph_id in reversed(graph_ids):
        if graph_id not in seen:
            seen.add(graph_id)
            unique_ids.append(graph_id)
    unique_ids.reverse()

    for graph_id in unique_ids:
        graph_json = alizarin.get_graph_json(graph_id)
        graph = json.loads(graph_json)
        name = graph.get("name", graph_id)
        if isinstance(name, dict):
            name = name.get("en", name.get(next(iter(name)), graph_id))
        subtype = "resource_models" if graph.get("isresource") else "branches"
        graph_dir = out_dir / "graphs" / subtype
        graph_dir.mkdir(parents=True, exist_ok=True)
        safe_name = name.replace("/", "_")
        graph_path = graph_dir / f"{safe_name}.json"
        graph_path.write_text(json.dumps({"graph": [graph]}, indent=2))
        print(f"  {graph_path}")


def export_collections(out_dir):
    """Export all collections from cache as SKOS XML."""
    skos_dir = out_dir / "reference_data" / "collections"
    skos_dir.mkdir(parents=True, exist_ok=True)
    cache = alizarin.get_global_rdm_cache()
    cids = cache.get_collection_ids()
    for cid in cids:
        col = cache.get_collection(cid)
        xml = col.to_skos_xml(BASE_URI)
        (skos_dir / f"{cid}.xml").write_text(xml)
    print(f"  Wrote {len(cids)} SKOS collections to {skos_dir}/")


def main():
    parser = argparse.ArgumentParser(
        description="Build ICH-extended graphs from arches-her + ICH mutations"
    )
    parser.add_argument("--pkg-dir", required=True,
                        help="Path to arches_her/pkg directory")
    parser.add_argument("--output-dir", default="output")
    parser.add_argument("--extra-ontology-dir", action="append", default=[],
                        help="Additional ontology directory (with ontology_config.json). "
                             "Can be repeated. E.g. arches-for-science Linked Art ontology.")
    parser.add_argument("--no-ontology-validation", action="store_true",
                        help="Disable ontology class/property validation")
    args = parser.parse_args()

    pkg_dir = Path(args.pkg_dir)
    out_dir = Path(args.output_dir)

    # --- Setup ---
    alizarin.set_global_rdm_cache(RustRdmCache())
    alizarin.set_rdm_namespace(BASE_URI)

    # --- Reference Data ---
    print("--- HER Reference Data ---")
    load_rdm_cache(pkg_dir)

    # --- Package Graphs (arches-her base) ---
    print("--- HER Package Graphs ---")
    graphs_dir = pkg_dir / "graphs"
    if not graphs_dir.exists():
        parser.error(f"Graphs directory not found: {graphs_dir}\n"
                     f"Check that --pkg-dir points to the arches_her/pkg directory.")
    graph_ids = load_pkg_graphs(pkg_dir)

    # --- MHP models (authoritative hand-built resource models) ---
    # Convert concept nodes to reference type with controlledList.
    mhp_dir = SCRIPT_DIR / "mhp"
    if mhp_dir.is_dir():
        print("--- MHP Models ---")
        for fp in sorted(mhp_dir.glob("*.json")):
            data = json.loads(fp.read_text())
            graph = data.get("graph", [{}])[0] if "graph" in data else data
            for node in graph.get("nodes", []):
                if node.get("datatype") in ("concept", "concept-list"):
                    node["datatype"] = "reference"
                    config = node.get("config") or {}
                    if not config.get("controlledList"):
                        name = node.get("name", "")
                        if isinstance(name, dict):
                            name = name.get("en", "")
                        config["controlledList"] = name.replace("-", " ").title()
                        node["config"] = config
            graph_id = alizarin.register_graph(json.dumps(data))
            print(f"  Registered: {fp.name}")
            graph_ids.append(graph_id)

    # --- Ontology ---
    ontology_validator = None
    if not args.no_ontology_validation:
        from alizarin import OntologyValidator
        ontology_files = []

        fetch_remote_ontologies()

        # Load ontology dirs: arches-her cidoc_crm, cached remote ontologies, extras
        ontology_dirs = [pkg_dir / "ontologies" / "cidoc_crm"]
        for subdir, _ in REMOTE_ONTOLOGIES:
            ontology_dirs.append(ONTOLOGY_CACHE_DIR / subdir)
        for extra in args.extra_ontology_dir:
            ontology_dirs.append(Path(extra))

        for ont_dir in ontology_dirs:
            config_path = ont_dir / "ontology_config.json"
            if config_path.exists():
                with open(config_path) as f:
                    config = json.load(f)
                ontology_files.append(str(ont_dir / config["base"]))
                for ext in config.get("extensions", []):
                    ext_path = ont_dir / ext
                    if ext_path.exists():
                        ontology_files.append(str(ext_path))
            else:
                # Directory without config — add all XML/RDF files directly
                for fp in sorted(ont_dir.glob("*.xml")) + sorted(ont_dir.glob("*.rdf")):
                    ontology_files.append(str(fp))

        if ontology_files:
            ontology_validator = OntologyValidator(ontology_files)
            print(f"  Loaded {ontology_validator.class_count} classes")

    # --- Register widget types ---
    for dt in ("reference", "reference-list"):
        alizarin.register_widget(
            "reference-select-widget",
            "19e56148-82b8-47eb-b66e-f6243639a1a8",
            default_config_json="{}",
            datatype=dt,
        )

    # --- Process numbered ICH CSV files ---
    # 60_collections.csv is excluded: it's reference-only; collection
    # assignments are applied at MHP load time above and inline in branch
    # add_node configs.
    print("\n--- ICH Mutations ---")
    graph_ids.extend(
        alizarin.process_mutation_csvs(
            str(SCRIPT_DIR), ontology_validator=ontology_validator,
            exclude_pattern=r"60_.*\.csv",
        )
    )

    # --- Export ---
    print("\n--- Export ---")
    export_graphs(graph_ids, out_dir)
    export_collections(out_dir)


if __name__ == "__main__":
    main()
