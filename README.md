# ICH — Intangible Cultural Heritage Graphs

Extensions to Arches HER resource models for Intangible Cultural Heritage.

Credit and thanks to [Takin Solutions](https://www.takin.solutions/) for establishing these.

## Structure

- `generate.py` — reads `Maori Heritage Project.xlsx` and produces numbered mutation CSVs
- `run.py` — loads the arches-her package, then applies the mutation CSVs to build extended graphs
- `01_*` – `17_*` — new ICH branch definitions
- `20_*` – `21_*` — new resource models (Event, Tikanga)
- `30_*` – `36_*` — modifications to existing HER models
- `40_*` – `48_*` — matauranga reference additions to unmodified HER models
- `60_collections.csv` — collection assignments (placeholder)

## Prerequisites

- Python 3.10+
- [alizarin](https://github.com/flaxandteal/alizarin) installed (`pip install alizarin`)
- A checkout of [arches-her](https://github.com/archesproject/arches-her) (the `pkg` directory is needed)

## Generating CSVs from the spreadsheet

```bash
python generate.py --xlsx "../Maori Heritage Project.xlsx" \
    --pkg-dir ../catalina-graphs/arches-her/arches_her/pkg
```

This reads the spreadsheet and writes numbered CSV files into the current directory.

## Building the graphs

```bash
python run.py --pkg-dir ../catalina-graphs/arches-her/arches_her/pkg \
    [--output-dir output/] \
    [--no-ontology-validation]
```

This:

1. Loads HER reference data (concepts and controlled lists)
2. Registers all HER branches and resource models (auto-publishing unpublished branches)
3. Processes each numbered ICH mutation CSV in order
4. Exports all graphs and SKOS collections to the output directory

Output lands in `output/` by default:

```
output/
  graphs/
    branches/       # HER + ICH branches
    resource_models/ # HER + ICH resource models (modified and new)
  reference_data/
    collections/    # SKOS XML collections
```

### Options

| Flag | Description |
|------|-------------|
| `--pkg-dir` | **(required)** Path to the arches-her `pkg` directory |
| `--output-dir` | Output directory (default: `output/`) |
| `--no-ontology-validation` | Skip CRM ontology class/property validation |
