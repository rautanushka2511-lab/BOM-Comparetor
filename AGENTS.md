# AGENTS: How to be immediately productive in this repository

This document gives concise, actionable guidance for AI coding agents working on this BOM Comparison project.
Follow the checklist below, then consult the referenced files for examples and patterns.

Checklist (quick-start)
- Understand the two-stage pipeline: comparator -> hierarchy -> outputs (Excel + JSON). See `app.py`, `bom_comp_levelcs.py`, `generate_json_hierarchy.py`.
- Run the service locally: `python app.py` (dev mode runs `uvicorn("app:app", reload=True)`).
- Reproduce offline: run comparator alone (`bom_comp_levelcs.py`) and hierarchy (`generate_json_hierarchy.py`) with the CLI flags shown below.
- Use `config91112.json` and `submodule_structure.json` for lookups and submodule mapping; the SBERT model bundle is under `models/all-MiniLM-L6-v2`.

Big picture (components & flow)
- FastAPI front-end/service: `app.py` — accepts uploads and streams NDJSON progress. It spawns two worker subprocesses (same python bundle) with `--worker comparator` and `--worker hierarchy`.
- Comparator: `bom_comp_levelcs.py` — normalises inputs, builds trees, runs embedding-based fuzzy matching (SBERT), classifies Carry-Over/Modified/Replaced/New/Deleted and exports an Excel report and `summary.json`.
- Hierarchy generator: `generate_json_hierarchy.py` — reads comparator Excel, builds JSON hierarchy `bom_hierarchy.json`, writes submodule summaries and injects SubmoduleSummary into `summary.json` and the Excel file.

Important command examples
- Run the FastAPI app (dev):
  - `python app.py`  # starts uvicorn with reload when not frozen
- Run comparator standalone:
  - python bom_comp_levelcs.py --old path/to/old.xlsx --new path/to/new.xlsx --output out.xlsx --config config91112.json --skip_choice Y --dynamic_pairs '[{"old":"OldCol","new":"NewCol"}]' --map_old '{"Component number":"Material"}' --map_new '{"Component number":"Item ID"}'
  - Notes: `--map_old` and `--map_new` are JSON maps from canonical column name -> actual column name in the file (e.g. {"Component number":"Material"}).
- Run hierarchy standalone:
  - python generate_json_hierarchy.py --input out.xlsx --output bom_hierarchy.json --submodules submodule_structure.json --config_path config91112.json --summary_json_path summary.json

Key runtime conventions & patterns
- Column canonical names used throughout: See top of `bom_comp_levelcs.py` (CAN_COL_LEVEL = "Level", CAN_COL_COMP = "Component number", CAN_COL_DESC = "Object description", CAN_COL_CS = "Special procurement", CAN_COL_QTY = "Component quantity"). Agents should use these canonical names when reasoning about columns.
- Column mapping: The code auto-detects common aliases via `_UNIFIED_COL_ALIAS_MAP` / `_SAP_COL_ALIAS_MAP` / `_PLM_COL_ALIAS_MAP`. If you pass `--map_old`/`--map_new`, provide mappings canonical->user column.
- Dynamic comparisons: `dynamic_pairs` is a JSON list of objects {"old": "<oldCol>", "new": "<newCol>"}. These values should be original/raw column names from each input file.
- Skip rule: `--skip_choice` is either `Y` (default special-procurement skipping rules apply) or `X` (no skipping). The FastAPI UI sends `Y`/`X`.
- Worker mode: `app.py` invokes itself in worker mode: `python app.py --worker comparator ...` and `--worker hierarchy ...`. When writing tooling that launches the comparator/hierarchy, replicate this pattern (use sys.executable when possible).

Data and model integration
- SBERT is expected to be available locally in `models/all-MiniLM-L6-v2`. `bom_comp_levelcs.py` will try to load it from that directory; otherwise it attempts to load by name and may fail without internet.
- Embedding cache: `.emb_cache` directory is used to cache SBERT embeddings (file names hashed by DataFrame bytes). Keep this between runs for speed.
- Optional acceleration: `faiss-cpu` and `rapidfuzz` are optional. The comparator falls back to NumPy-based NN search and non-rapidfuzz fuzzy behaviour if these are missing. See top of `bom_comp_levelcs.py`.

Files/directories you will use most
- `app.py` — FastAPI service entry-point and worker launcher
- `bom_comp_levelcs.py` — comparator and Excel exporter (primary heavy logic)
- `generate_json_hierarchy.py` — hierarchy generation and Excel injection
- `config91112.json` — segment rules, lookups, thresholds
- `submodule_structure.json` — submodule-to-location mapping used by the hierarchy
- `models/` — bundled sentence-transformer model
- `session_files/` — runtime per-request working directories created by `app.py`

Debugging tips specific to this repo
- If FastAPI stream ends with an error, check the last ~50 lines buffered by `app.py` — comparator/hierarchy print to stdout and `app.py` returns these on failure.
- Look for the comparator marker: `Exporting results to: <path>` — `app.py` uses this string to resolve the produced Excel path when outputs are written in a non-default name.
- For Cross-Over BOMs (root levels 0 vs 1), comparator implements pruning: see `_detect_root_level` and `prune_root_level` in `bom_comp_levelcs.py`. Tests or agents that craft sample inputs should vary 'Level' to exercise this path.

Build & packaging notes
- Desktop bundle: `npm run dist` (see `package.json`) runs PyInstaller to create a Python backend under `dist/app` then runs `electron-builder` to assemble Windows deployment artifacts into `Deployment/`.
- PyInstaller invocation is in `package.json` (script `build:py`) — it bundles `app.py`, hidden imports for `bom_comp_levelcs` and `generate_json_hierarchy`, and adds `static`, `models`, and `config91112.json` as data files.
- Use `do_build.py` / `do_build.bat` or `build_app.bat` if you need the scripted clean/build steps used by maintainers.

Examples of project-specific invariants to preserve when editing
- Never assume numeric component IDs are floats — code forces string dtype for ID columns and writes component columns with text format in Excel (see `write_sheet` in `bom_comp_levelcs.py`).
- Many functions rely on the `SourceIndex` mapping between cleaned rows and original raw input rows — preserve `SourceIndex` when refactoring export logic.
- `ParentPath` strings and immediate parent id logic are used heavily; try to use the helper `compare_parent_paths` and `node.parent` access rather than brittle string parsing.

If you are writing tests or automated fixes
- Use small sample Excel files exercising: (a) root level 0 vs 1, (b) dynamic_pairs mapping, (c) missing columns (map via `--map_old`), (d) submodule 14-char component IDs.
- Validate outputs: comparator produces an XLSX and `summary.json`; hierarchy produces `bom_hierarchy.json` and updates `summary.json`. Confirm these three artefacts after runs.

Quick pointers to inspect
- CLI flags: see `argparse` in `bom_comp_levelcs.py` (lines around the `main()` parser) and `generate_json_hierarchy.py`.
- Build scripts: `package.json`, `do_build.py`, `do_build.bat`, `build_app.bat`.

Last updated: 2026-05-09

