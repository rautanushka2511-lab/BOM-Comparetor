#!/usr/bin/env python3
"""
generate_json_hierarchy.py
Build a BOM JSON tree using a unified hierarchy approach.
- MODIFIED: Corrected duplicate detection to find components with multiple parents.
- FIXED: Repaired subtree statistics calculation to ensure accurate counts.
- ADDED: Metadata filtering to reduce the size of the 'raw' data field based on primary status.
- MODIFIED: This script now handles the final "Submodule Summary" generation.
            It reads the output from bom_comp_levelcs.py, generates the hierarchy,
            calculates subtree stats for submodules, and then *modifies*
            the summary.json and Excel files to inject this new data.
- FIX: Corrected logic to find submodules for "Replaced" and "Deleted" nodes.
- FIX: Submodule summary now uses the node's actual `id` (e.g., "X --> Y")
         and `name` (for description) instead of generic metadata.
- ADDED: Text-based Design Group summary to the top of the Submodule Summary sheet.
- FIX: Corrected openpyxl method from `merge_range` to `merge_cells`.
- MODIFIED (Nov 2025): Submodule logic now uses a composite key (Design Group, Submodule Address)
                       and extracts both from the component ID.
- MODIFIED (Nov 2025): Submodule logic now only triggers for component IDs of *exactly* 14 digits.
- MODIFIED (Nov 2025): Removed "depth == 2" check.
- MODIFIED (Nov 2025): Added logic to consolidate flat root-level submodules under one true root.
- MODIFIED (Nov 2025): Added cleaning for component numbers with '_' suffix (e.g., _01).
- MODIFIED (Nov 2025): Fixed typo in `update_summary_json` (summary_path -> summary_json_path).
- MODIFIED (Nov 2025): Added --config_path arg. Submodule map is now built by *merging*
                       config91112.json (for descriptions) and submodule_structure.json (for locations).
- MODIFIED (Nov 2025): Removed incorrect underscore cleaning from PASS 1.
- MODIFIED (Nov 2025): Added depth check (<= 3) for submodule identification.
- MODIFIED (Nov 2025): Submodule logic now *reads* the Design Group from the raw data
                       and *slices* the Submodule Address from the component number.
- MODIFIED (Nov 2025): Corrected submodule address slicing to 5 digits [4:9].
- MODIFIED (Nov 2025): Reads "Parent Path Changed" from Excel and adds it as a
                       top-level key to each node in the JSON for use by the visualizer.
- MODIFIED (Nov 2025): Reads `root_level` from summary.json to implement dynamic
                       submodule depth detection (root_level + 1).
- MODIFIED (Nov 2025): Detects `is_dummy` flag for submodules with '_' in original component number.
- FIX (Nov 2025): Submodule depth check is now correctly `depth == 2` regardless of `root_level`.
- FIX (Nov 2025): Submodule logic is now robust. It tries to read Design Group
                  from the Excel column (for PLM) and deduces it from the
                  component number if the column is missing (for SAP).
- FIX (Nov 2025): Submodule logic now based on component number structure, not depth.
                  - 14-char = Submodule
                  - 14-char + '_' + text = Dummy Submodule
- FIX (Nov 2025): Updated to read Space-separated Excel headers (e.g., "Old Component number")
                  instead of Underscores to match updated bom_comp_levelcs.py.
- MODIFIED (Dec 2025): Updated Metadata Filtering to PRESERVE 'Old Level' and 'New Level'
                       so Cross-Over comparison (e.g., 0 vs 1) is visible in Inspector.
"""

import argparse
import json
import math
import re
import sys
from typing import Any, Dict, List, Optional
import pandas as pd
from collections import defaultdict, Counter

try:
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    print("Error: `openpyxl` library not found. Please install it with `pip install openpyxl`", file=sys.stderr)
    sys.exit(1)

# ---------------------------
# Helpers
# ---------------------------

def _clean_value(val: Any) -> str:
    if val is None or (isinstance(val, float) and math.isnan(val)):
        return ""
    return str(val).strip()

_path_cache = {}
def normalize_path(p: str) -> str:
    if not p: return ""
    if p in _path_cache: return _path_cache[p]
    normalized = str(p).strip().strip('/')
    _path_cache[p] = normalized
    return normalized

def natural_key(s: str):
    s = s or ""; parts = re.split(r'(\d+)', s); key = [int(p) if p.isdigit() else p.lower() for p in parts]; return key

def parse_change_type(change_str: str, raw: Dict[str, Any]) -> List[str]:
    if not change_str: return []
    items = [item.strip() for part in change_str.split(';') for item in part.split(',')]; return sorted([item for item in items if item and item in raw and ':' in raw[item]])

def split_path(path: str) -> List[str]:
    p = normalize_path(path); return [seg for seg in p.split('/') if seg] if p else []

def join_path(parts: List[str]) -> str:
    return '/'.join([p for p in parts if p])

def pick_first_existing(raw: Dict[str, Any], keys: List[str]) -> Optional[str]:
    for k in keys:
        v = raw.get(k)
        if v not in (None, ''): return str(v)
    return None

def load_data_from_sheets(excel_path: str, sheet_spec: str, debug: bool = True) -> pd.DataFrame:
    def dbg(*args, **kwargs):
        if debug: print(*args, **kwargs)
    try: xls = pd.ExcelFile(excel_path)
    except FileNotFoundError: raise FileNotFoundError(f"Input file not found: {excel_path}")
    
    all_sheet_names = xls.sheet_names
    filtered_sheet_names = [name for name in all_sheet_names if not re.search(r'dynamic\s*compare', name, re.IGNORECASE)]
    
    if sheet_spec.lower() == 'all': sheets_to_load = filtered_sheet_names
    else:
        requested_sheets = [s.strip() for s in sheet_spec.split(',')]; sheets_to_load = [s for s in requested_sheets if s in filtered_sheet_names]
    
    if not sheets_to_load: raise ValueError(f"No valid sheets could be loaded from '{excel_path}' with spec '{sheet_spec}'.")
    
    df_list = [pd.read_excel(xls, sheet_name=s, dtype=str).fillna("") for s in sheets_to_load]
    combined_df = pd.concat(df_list, ignore_index=True)
    dbg(f"Total rows loaded: {len(combined_df)}")
    return combined_df

# ---------------------------
# NEW: Configuration for Metadata Filtering
# ---------------------------
# UPDATED (Dec 2025): Removed "Old Level" and "New Level" from removal lists 
# to support Cross-Over visualization.
COLS_TO_REMOVE = {
    "Carry-Over": {
        "Change Type", "Change Detail", "Level", "Component number",
        "Object description", "Special procurement", "Component quantity", "ParentPath",
        "SkippedChildren"
    },
    "Modified": {
        "Level", "Component number", "Object description", "Special procurement",
        "Component quantity", "ParentPath", "SkippedChildren"
    },
    "Replaced": {
        "Level", "Component number", "Object description", "Special procurement",
        "Component quantity", "ParentPath", "SkippedChildren"
    },
    "Newly Added": {
        "Old Component number", "Old Object description",
        "Old Special procurement", "Old Component quantity", "Old ParentPath",
        "Old SkippedChildren", "Old Bulk Mat. (Item Mat)", 
        "Change Type", "Change Detail", "Level", "Component number",
        "Object description", "Special procurement", "Component quantity",
        "ParentPath", "SkippedChildren"
    },
    "Deleted": {
        "New Component number", "New Object description",
        "New Special procurement", "New Component quantity", "New ParentPath",
        "New SkippedChildren", "New Bulk Mat. (Item Mat)", 
        "Change Type", "Change Detail", "Level", "Component number",
        "Object description", "Special procurement", "Component quantity",
        "ParentPath", "SkippedChildren"
    }
}


# ---------------------------
# Core build logic
# ---------------------------

def build_merged_tree(excel_path: str, submodule_path: str, config_path: str, sheet_name: str = "All", debug: bool = True, root_level: int = 1) -> List[Dict]:
    def dbg(*args, **kwargs):
        if debug: print(*args, **kwargs)

    submodule_map = {}
    
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            config_data = json.load(f)
        lookups = config_data.get("lookups", {}).get("design_subgroups", {})
        for dg, addresses in lookups.items():
            for addr, desc in addresses.items():
                submodule_map[(dg, addr)] = {
                    "DESIGN GROUP": dg,
                    "SUBMODULE ADDRESS": addr,
                    "SUBMODULE": desc, 
                    "Location": "" 
                }
        dbg(f"Loaded {len(submodule_map)} submodule descriptions from {config_path}.")
    except Exception as e:
        raise ValueError(f"Error parsing config JSON file {config_path}: {e}")

    try:
        with open(submodule_path, 'r', encoding='utf-8') as f:
            submodule_data = json.load(f)
        location_count = 0
        for item in submodule_data:
            key = (item.get("DESIGN GROUP"), item.get("SUBMODULE ADDRESS"))
            if key in submodule_map:
                submodule_map[key]["Location"] = item.get("Location", "")
                location_count += 1
            elif item.get("DESIGN GROUP") and item.get("SUBMODULE ADDRESS"):
                submodule_map[key] = {
                    "DESIGN GROUP": item.get("DESIGN GROUP"),
                    "SUBMODULE ADDRESS": item.get("SUBMODULE ADDRESS"),
                    "SUBMODULE": item.get("SUBMODULE", ""),
                    "Location": item.get("Location", "")
                }
        dbg(f"Updated/Added {location_count} submodules with locations from {submodule_path}.")
    except Exception as e:
        raise ValueError(f"Error parsing submodule JSON file {submodule_path}: {e}")
    
    addr_to_dg_map = { v['SUBMODULE ADDRESS']: v['DESIGN GROUP'] for k, v in submodule_map.items() }
    dbg(f"Built inverted map for {len(addr_to_dg_map)} addresses.")

    df = load_data_from_sheets(excel_path, sheet_name, debug)
    
    required_cols = ["Primary Status", "Hierarchical Status", "Old Component number", "New Component number", "Old ParentPath", "New ParentPath", "Parent Path Changed"]
    if any(col not in df.columns for col in required_cols): 
        missing = [col for col in required_cols if col not in df.columns]
        print(f"Warning: Missing required columns {missing} in input data.", file=sys.stderr)

    node_map: Dict[str, Dict] = {}
    canonical_redirect: Dict[str, str] = {}
    replaced_olds = set()
    valid_primary_statuses = {"Newly Added", "Replaced", "Deleted", "Carry-Over", "Modified"}

    def _name(id_: str, desc: str) -> str: return f"{id_} — {desc}" if desc else id_

    for idx, r in df.iterrows():
        row = {k: _clean_value(v) for k, v in r.to_dict().items()}
        primary_status = (row.get("Primary Status") or "Carry-Over").strip().title()
        if primary_status not in valid_primary_statuses: primary_status = "Carry-Over"

        new_id = row.get("New Component number") or ""
        old_id = row.get("Old Component number") or ""

        if not new_id and not old_id: continue

        new_ppath = normalize_path(row.get("New ParentPath", ""))
        old_ppath = normalize_path(row.get("Old ParentPath", ""))
        path_new = f"{new_ppath}/{new_id}" if new_id and new_ppath else new_id
        path_old = f"{old_ppath}/{old_id}" if old_id and old_ppath else old_id
        
        node_id = new_id or old_id
        desc = row.get("New Object description") or row.get("Old Object description") or ""
        hierarchical_status = row.get("Hierarchical Status", "Unchanged")
        presence = "B1/B2" if new_id and old_id else ("B2" if new_id else "B1")
        
        parent_path_changed = row.get("Parent Path Changed", "N/A")

        if primary_status == "Replaced":
            if not new_id or not old_id: continue
            replaced_olds.add(old_id)
            canonical_path = normalize_path(path_new) or normalize_path(path_old) or new_id
            merged_id = f"{old_id} --> {new_id}"
            
            merged_node = { 
                "id": merged_id, 
                "name": _name(merged_id, desc), 
                "primary_status": "Replaced", 
                "hierarchical_status": hierarchical_status, 
                "children": [], 
                "raw": row.copy(), 
                "presence": "B1/B2", 
                "is_placeholder": False,
                "parent_path_changed": parent_path_changed 
            }
            key_choice = canonical_path or merged_id
            node_map.setdefault(key_choice, merged_node)
            if normalize_path(path_old): canonical_redirect[normalize_path(path_old)] = key_choice
            if normalize_path(path_new) and normalize_path(path_new) != normalize_path(path_old): canonical_redirect[normalize_path(path_new)] = key_choice
        else:
            if node_id in replaced_olds and primary_status in ("Deleted", "Carry-Over"): continue
            
            if primary_status == "Deleted": canonical_path = normalize_path(path_old)
            elif primary_status == "Newly Added": canonical_path = normalize_path(path_new)
            else: canonical_path = normalize_path(path_new) or normalize_path(path_old)
            if not canonical_path: canonical_path = node_id

            node_map.setdefault(canonical_path, { 
                "id": node_id, 
                "name": _name(node_id, desc), 
                "primary_status": primary_status, 
                "hierarchical_status": hierarchical_status, 
                "children": [], 
                "raw": row.copy(), 
                "presence": presence, 
                "is_placeholder": False,
                "parent_path_changed": parent_path_changed 
            })

    for path in list(node_map.keys()):
        parts = split_path(path)
        for i in range(1, len(parts)):
            parent_path = join_path(parts[:i])
            if parent_path not in node_map:
                node_map[parent_path] = { 
                    "id": parts[i - 1], 
                    "name": parts[i - 1], 
                    "primary_status": "Carry-Over", 
                    "hierarchical_status": "Unchanged", 
                    "children": [], 
                    "raw": {}, 
                    "presence": "", 
                    "is_placeholder": True,
                    "parent_path_changed": "No" 
                }

    def resolve_key(key: str) -> str:
        key_n = normalize_path(key); visited = {key_n}
        while key_n in canonical_redirect:
            key_n = canonical_redirect[key_n]
            if key_n in visited: break
            visited.add(key_n)
        return key_n
    
    parent_child_map = defaultdict(list)
    for key, node in node_map.items():
        if node.get("is_placeholder"): continue
        raw = node.get("raw", {})
        
        parent_path_str = raw.get("New ParentPath") if node['primary_status'] not in ['Deleted'] else raw.get("Old ParentPath")
        if not parent_path_str: parent_path_str = raw.get("Old ParentPath")
        
        parts = split_path(key)
        structural_parent_path = join_path(parts[:-1]) if len(parts) > 1 else ""
        
        parent_key = resolve_key(normalize_path(parent_path_str))
        if parent_key not in node_map and structural_parent_path:
             parent_key = resolve_key(structural_parent_path)

        if parent_key and parent_key in node_map:
            parent_child_map[parent_key].append(key)

    for p_key, c_keys in parent_child_map.items():
        if p_key in node_map:
            node_map[p_key]["children"].extend([node_map[c_key] for c_key in c_keys if c_key in node_map])

    top_level_nodes = [node for key, node in node_map.items() if key not in {c for p in parent_child_map.values() for c in p} and not node.get("is_placeholder")]

    real_roots = []
    submodule_roots = []
    for node in top_level_nodes:
        primary_status = node.get("primary_status", "Carry-Over")
        check_id = None
        
        if primary_status == "Replaced":
            check_id = node.get("raw", {}).get("New Component number")
        elif primary_status == "Deleted":
            check_id = node.get("raw", {}).get("Old Component number")
        else:
            check_id = pick_first_existing(node.get("raw", {}), ["New Component number", "Old Component number"]) or node.get("id")
        
        if check_id and "_" in check_id:
            check_id = check_id.split("_", 1)[0]
            
        if check_id and len(check_id) == 14:
            submodule_roots.append(node)
        else:
            real_roots.append(node)

    if real_roots and submodule_roots:
        best_root = next((r for r in real_roots if r.get("primary_status") == "Replaced"), None)
        if not best_root:
             best_root = next((r for r in real_roots if "B2" in r.get("presence", "")), None)
        if not best_root and real_roots:
             best_root = real_roots[0]
        
        if best_root:
            dbg(f"Consolidating {len(submodule_roots)} flat submodules under root '{best_root.get('id')}'...")
            best_root.setdefault("children", []).extend(submodule_roots)
            top_level_nodes = real_roots 
        else:
             dbg("Warning: No best root found, but real roots and submodule roots exist. Re-combining.")
             top_level_nodes = real_roots + submodule_roots
    
    elif not real_roots and submodule_roots:
        dbg("Warning: No non-submodule roots found. Tree will be flat.")
        top_level_nodes = submodule_roots 
    
    else:
        dbg("No flat submodules found to consolidate or no real root found.")
        top_level_nodes = real_roots + submodule_roots 
    
    primary_b2_root = next((n for n in sorted(top_level_nodes, key=lambda n: natural_key(n.get("id",""))) if "B2" in n.get("presence", "")), None)
    
    deleted_tops = [n for n in top_level_nodes if n.get("primary_status") == "Deleted" and n not in submodule_roots] 
    
    if deleted_tops and primary_b2_root:
        primary_b2_root["children"].extend(deleted_tops)
        top_level_nodes = [n for n in top_level_nodes if n not in deleted_tops]


    component_parents = defaultdict(set)
    q = list(top_level_nodes)
    visited_nodes = set()
    while q:
        node = q.pop(0)
        uid = id(node)
        if uid in visited_nodes:
            continue
        visited_nodes.add(uid)

        node_id = node.get("id")
        parent_path = node.get('raw', {}).get('New ParentPath') or node.get('raw', {}).get('Old ParentPath')
        
        if "-->" in node_id:
            old_part, new_part = [p.strip() for p in node_id.split("-->")]
            component_parents[old_part].add(node.get('raw', {}).get('Old ParentPath'))
            component_parents[new_part].add(node.get('raw', {}).get('New ParentPath'))
        else:
            component_parents[node_id].add(parent_path)
        
        for child in node.get("children", []):
            q.append(child)
            
    duplicate_ids = {cid for cid, parents in component_parents.items() if len(parents) > 1}
    dbg(f"Found {len(duplicate_ids)} components with multiple parents (duplicates).")
            
    def final_pass(node: Dict, depth: int, parent_path_str: str):
        raw = node.setdefault("raw", {})
        node_id = node.get("id", "")
        primary_status = node.get("primary_status", "Carry-Over")
        
        is_dup = False
        if "-->" in node_id:
            old_part, new_part = [p.strip() for p in node_id.split("-->")]
            is_dup = old_part in duplicate_ids or new_part in duplicate_ids
        else:
            is_dup = node_id in duplicate_ids
        node['is_duplicate'] = is_dup

        node['is_bulk_material'] = (raw.get("Old Bulk Mat. (Item Mat)", "").strip().upper() == 'X' or raw.get("New Bulk Mat. (Item Mat)", "").strip().upper() == 'X')
        
        if "parent_path_changed" not in node:
             node["parent_path_changed"] = raw.get("Parent Path Changed", "N/A")

        node['submodule_data'] = None
        node['is_dummy'] = False
        check_id = None

        original_comp = pick_first_existing(raw, ["New Component number", "Old Component number"]) or node_id

        if original_comp:
            if "_" in original_comp:
                parts = original_comp.split("_", 1)
                if len(parts[0]) == 14:
                    check_id = parts[0] 
                    node['is_dummy'] = True
            
            elif len(original_comp) == 14:
                check_id = original_comp
                node['is_dummy'] = False

            if check_id:
                submodule_address = check_id[4:9]
                design_group = None
                
                design_group_raw = pick_first_existing(raw, ["New Design Group", "Old Design Group"])
                if design_group_raw:
                    design_group = str(design_group_raw).strip().split('.')[0].zfill(2)
                    
                if not design_group:
                    design_group = addr_to_dg_map.get(submodule_address)
                    
                # Always ensure it is classified as a submodule by attaching submodule_data
                fallback_data = {
                    "DESIGN GROUP": design_group or "",
                    "SUBMODULE ADDRESS": submodule_address,
                    "SUBMODULE": node.get("name", check_id)
                }
                    
                if design_group:
                    lookup_key = (design_group, submodule_address)
                    node['submodule_data'] = submodule_map.get(lookup_key, fallback_data)
                else:
                    node['submodule_data'] = fallback_data

        raw['Component number'] = pick_first_existing(raw, ["New Component number", "Old Component number"]) or node_id
        raw['Parent path'] = parent_path_str
        raw['Json depth'] = str(depth) 
        
        if primary_status in COLS_TO_REMOVE:
            keys_to_remove = COLS_TO_REMOVE[primary_status]
            node['raw'] = {k: v for k, v in raw.items() if k not in keys_to_remove}

        current_path = f"{parent_path_str}/{node_id}" if parent_path_str else node_id
        for ch in node.get("children", []):
            final_pass(ch, depth + 1, current_path)
            
        node["children"] = sorted(node.get("children", []), key=lambda c: natural_key(c.get("id", "")))

        change_counts = defaultdict(int)
        
        change_counts[primary_status] += 1
            
        total_desc = len(node.get("children", []))
        for ch in node.get("children", []):
            ch_stats = ch.get("subtree_stats", {})
            total_desc += ch_stats.get("total_descendants", 0)
            for status, count in ch_stats.get("change_counts", {}).items():
                change_counts[status] += count
        
        node["subtree_stats"] = {
            "total_direct_children": len(node.get("children", [])),
            "total_descendants": total_desc,
            "change_counts": dict(change_counts) 
        }

    sorted_roots = sorted(top_level_nodes, key=lambda n: natural_key(n.get("id", "")))
    for root in sorted_roots:
        final_pass(root, 1, "")
    
    return sorted_roots

# ---------------------------
# NEW: Post-processing to update Excel/JSON
# ---------------------------

def _find_all_submodules(nodes: List[Dict]) -> List[Dict]:
    submodules = []
    stack = list(nodes)
    visited = set()
    while stack:
        node = stack.pop()
        node_id = node.get("id")
        if node_id in visited:
            continue
        visited.add(node_id)
        
        if node.get("submodule_data"):
            submodules.append(node)
        
        stack.extend(node.get("children", []))
    return submodules

def generate_submodule_summary_data(root_nodes: List[Dict], manual_overrides: Dict[str, str] = None) -> List[Dict]:
    submodule_nodes = _find_all_submodules(root_nodes)
    summary_list = []
    
    def _norm_status(s):
        if not s: return 'Carry-Over'
        s = str(s).strip()
        if s.lower() == 'unchanged': return 'Carry-Over'
        if 'added' in s.lower(): return 'Newly Added'
        if 'deleted' in s.lower(): return 'Deleted'
        if 'replaced' in s.lower(): return 'Replaced'
        if 'modified' in s.lower(): return 'Modified'
        return s

    for node in submodule_nodes:
        meta = node.get("submodule_data", {})
        stats = node.get("subtree_stats", {})
        counts = stats.get("change_counts", {})
        
        node_id = node.get("id", "")
        node_name = node.get("name", "")
        
        desc = meta.get("SUBMODULE", "") 
        if not desc:
            if " — " in node_name:
                desc = node_name.split(" — ", 1)[1]
            elif node_name != node_id:
                desc = node_name

        is_dummy = node.get("is_dummy", False) 
        raw_status = node.get("primary_status", "Carry-Over")
        
        split_nodes = []
        if " --> " in node_id and manual_overrides:
            old_id, new_id = node_id.split(" --> ")
            old_id, new_id = old_id.strip(), new_id.strip()
            if old_id in manual_overrides or new_id in manual_overrides:
                split_nodes = [
                    (old_id, manual_overrides.get(old_id, "Deleted")),
                    (new_id, manual_overrides.get(new_id, "Newly Added"))
                ]
                
        if not split_nodes:
            if manual_overrides and node_id in manual_overrides:
                ps = manual_overrides[node_id]
            else:
                ps = _norm_status(raw_status)
            
            # General algorithmic path using SBERT natural outputs
            split_nodes = [(node_id, ps)]

        for part_id, ps in split_nodes:
            exclude = False
            
            # Compute final normalized status early so it is consistent
            final_ps = ps
            if final_ps == "":
                final_ps = "Carry-Over"

            summary_list.append({
                "Design Group": meta.get("DESIGN GROUP", ""),
                "Submodule Address": meta.get("SUBMODULE ADDRESS", ""),
                "Component Number": part_id, 
                "Description": desc, 
                "Location": meta.get("Location", ""),
                "Primary Status": ps,
                "Computed_Status": final_ps,
                "Hierarchical Status": node.get("hierarchical_status", "Unchanged"),
                "Total Added": counts.get("Newly Added", 0),
                "Total Deleted": counts.get("Deleted", 0),
                "Total Replaced": counts.get("Replaced", 0),
                "Total Modified": counts.get("Modified", 0),
                "Total Carry-Over": counts.get("Carry-Over", 0),
                "is_dummy": is_dummy,
                "exclude_from_counts": exclude
            })
        
    summary_list.sort(key=lambda x: (x["Design Group"], x["Submodule Address"], x["Component Number"]))
    return summary_list

def load_manual_overrides(excel_path: str) -> Dict[str, str]:
    """Reads manual classifications from Excel 'Groups' sheets if they exist."""
    overrides = {}
    try:
        # Load workbook with data_only=True to get values instead of formulas
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        sheet_map = {
            'Newly Added Groups': 'Newly Added',
            'Deleted Groups': 'Deleted',
            'Replaced Groups': 'Replaced',
            'Modified Groups': 'Modified',
            'Carryover Groups': 'Carry-Over'
        }
        sub_cols = [3, 7, 11, 15, 19, 23] # Submodule ID columns in group sheets
        for sheet_name, status in sheet_map.items():
            if sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                # IDs start from row 3
                for row in ws.iter_rows(min_row=3, values_only=True):
                    for ci in sub_cols:
                        if ci < len(row) and row[ci] is not None:
                            val = str(row[ci]).strip()
                            if val and len(val) > 5: # Minimal sanity check for part ID
                                overrides[val] = status
        print(f"Loaded {len(overrides)} manual submodule overrides from {excel_path}.")
    except Exception as e:
        print(f"Note: Could not load manual overrides from {excel_path}: {e}")
    return overrides

def update_summary_json(summary_json_path: str, submodule_data: List[Dict]):
    try:
        with open(summary_json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except Exception as e:
        print(f"Warning: Could not read summary.json at {summary_json_path}. {e}", file=sys.stderr)
        data = {}
        
    data["SubmoduleSummary"] = submodule_data
    
    # Update Status Breakdown to match SubmoduleSummary counts (for Dashboard KPIs)
    # This ensures the dashboard matches the manual summary numbers.
    counts = Counter([item.get("Computed_Status", item["Primary Status"]) for item in submodule_data if not item.get("exclude_from_counts")])
    status_breakdown = {}
    for s in ["Newly Added", "Deleted", "Replaced", "Modified", "Carry-Over"]:
        status_breakdown[s] = {"count": counts.get(s, 0), "percentage": 0.0}
    
    data["Status Breakdown"] = status_breakdown
    
    try:
        with open(summary_json_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        print(f"Successfully updated {summary_json_path} with SubmoduleSummary and Status Breakdown.")
    except Exception as e:
        print(f"Error: Could not write updated summary.json. {e}", file=sys.stderr)

def update_excel_report(excel_path: str, submodule_data: List[Dict]):
    try:
        wb = openpyxl.load_workbook(excel_path)
    except Exception as e:
        print(f"Error: Could not load Excel file {excel_path} for modification. {e}", file=sys.stderr)
        return

    # Keep output deterministic: recreate summary sheets on each run.
    for sheet_name in ("Submodule Summary", "Summary", "summery", "Submodule Summery", "Carryover Groups", "Exclusive Groups", "Modified Groups", "Newly Added Groups", "Replaced Groups", "Deleted Groups"):
        if sheet_name in wb.sheetnames:
            del wb[sheet_name]

    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    title_font = Font(bold=True, size=14, color="1F497D")
    text_font = Font(bold=True)
    text_align = Alignment(wrap_text=True, vertical="top")
    thin_side = Side(style="thin", color="1F1F1F")
    thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    fill_colors = {
        "Added": "C6EFCE",
        "Deleted": "FFC7CE",
        "Replaced": "FFEB9C",
        "Modified": "FCD5B4",
        "Carry-Over": "D9E1F2",
    }
    count_fills = {k: PatternFill(start_color=v, end_color=v, fill_type="solid") for k, v in fill_colors.items()}

    def _norm_status(v: Any) -> str:
        s = str(v or "").strip().lower()
        if s in {"added", "newly added"}:
            return "Newly Added"
        if s == "deleted":
            return "Deleted"
        if s == "replaced":
            return "Replaced"
        if s == "modified":
            return "Modified"
        return "Carry-Over"

    def _module_count(items: List[Dict], statuses: List[str]) -> int:
        dgs = {
            str(i.get("Design Group", "")).strip()
            for i in items
            if i.get("Computed_Status", _norm_status(i.get("Primary Status"))) in statuses and str(i.get("Design Group", "")).strip()
        }
        return len(dgs)

    def _module_list(items: List[Dict], statuses: List[str]) -> str:
        dgs = sorted(list({
            str(i.get("Design Group", "")).strip()
            for i in items
            if i.get("Computed_Status", _norm_status(i.get("Primary Status"))) in statuses and str(i.get("Design Group", "")).strip()
        }))
        return ", ".join(dgs)

    def _write_submodule_table(ws, table_start_row: int):
        headers = [
            "Design Group", "Submodule Address", "Component Number", "Description", "Location",
            "Primary Status", "Hierarchical Status",
            "Total Added", "Total Deleted", "Total Replaced", "Total Modified", "Total Carry-Over",
        ]
        col_widths = [14, 20, 30, 45, 20, 15, 20, 15, 15, 15, 15, 18]

        for c, (header, width) in enumerate(zip(headers, col_widths), 1):
            cell = ws.cell(row=table_start_row, column=c, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_align
            cell.border = thin_border
            ws.column_dimensions[get_column_letter(c)].width = width

        row_num = table_start_row + 1
        for item in submodule_data:
            ws.cell(row=row_num, column=1, value=item.get("Design Group"))
            ws.cell(row=row_num, column=2, value=item.get("Submodule Address"))
            ws.cell(row=row_num, column=3, value=item.get("Component Number"))
            ws.cell(row=row_num, column=4, value=item.get("Description"))
            ws.cell(row=row_num, column=5, value=item.get("Location"))
            ws.cell(row=row_num, column=6, value=item.get("Primary Status"))
            ws.cell(row=row_num, column=7, value=item.get("Hierarchical Status"))
            ws.cell(row=row_num, column=8, value=item.get("Total Added")).fill = count_fills["Added"]
            ws.cell(row=row_num, column=9, value=item.get("Total Deleted")).fill = count_fills["Deleted"]
            ws.cell(row=row_num, column=10, value=item.get("Total Replaced")).fill = count_fills["Replaced"]
            ws.cell(row=row_num, column=11, value=item.get("Total Modified")).fill = count_fills["Modified"]
            ws.cell(row=row_num, column=12, value=item.get("Total Carry-Over")).fill = count_fills["Carry-Over"]
            for c in range(1, 13):
                ws.cell(row=row_num, column=c).border = thin_border
            row_num += 1

    dg_counts = defaultdict(lambda: {
        "Newly Added": 0,
        "Deleted": 0,
        "Replaced": 0,
        "Modified": 0,
        "Carry-Over": 0,
        "total": 0
    })
    status_counts = {
        "Carry-Over": 0,
        "Modified": 0,
        "Newly Added": 0,
        "Replaced": 0,
        "Deleted": 0,
    }

    for item in submodule_data:
        dg = str(item.get("Design Group", "")).strip()
        if dg:
            dg_counts[dg]["total"] += 1
            # Determine the submodule's status for DG categorization
            node_ps = item.get("Computed_Status", item.get("Primary Status", ""))
            if node_ps in dg_counts[dg]:
                dg_counts[dg][node_ps] += 1
            else:
                dg_counts[dg]["Carry-Over"] += 1

        if not item.get("exclude_from_counts"):
            node_ps = item.get("Computed_Status", item.get("Primary Status", ""))
            status_counts[node_ps] += 1

    # Detailed status-wise DG lists for the Submodule Summary sheet
    added_dgs = []
    deleted_dgs = []
    replaced_dgs = []
    modified_dgs = []
    carry_over_dgs = []

    for dg in sorted(dg_counts.keys(), key=natural_key):
        c = dg_counts[dg]
        total = c["total"]
        
        if c["Newly Added"] > 0:
            added_dgs.append(f"{dg} ({c['Newly Added']}/{total})")
        if c["Deleted"] > 0:
            deleted_dgs.append(f"{dg} ({c['Deleted']}/{total})")
        if c["Replaced"] > 0:
            replaced_dgs.append(f"{dg} ({c['Replaced']}/{total})")
        if c["Modified"] > 0:
            modified_dgs.append(f"{dg} ({c['Modified']}/{total})")
        if c["Carry-Over"] > 0 and (c["Newly Added"] + c["Deleted"] + c["Replaced"] + c["Modified"]) == 0:
            carry_over_dgs.append(f"{dg} ({c['Carry-Over']}/{total})")

    def _write_summary_layout(ws):
        ws.cell(row=1, column=1, value="SUBMODULE CHANGE SUMMARY").font = title_font

        ws.cell(row=3, column=1, value="CHANGED DESIGN GROUPS (ADDED/DELETED/REPLACED/MODIFIED):").font = text_font
        ws.merge_cells(start_row=4, start_column=1, end_row=4, end_column=12)
        
        # Consolidate all changed categories for this specific view
        all_changed = added_dgs + deleted_dgs + replaced_dgs + modified_dgs
        ws.cell(row=4, column=1, value=", ".join(dict.fromkeys(all_changed)) or "None")
        ws.cell(row=4, column=1).alignment = text_align
        ws.row_dimensions[4].height = 40

        ws.cell(row=6, column=1, value="CARRY-OVER DESIGN GROUPS:").font = text_font
        ws.merge_cells(start_row=7, start_column=1, end_row=7, end_column=12)
        ws.cell(row=7, column=1, value=", ".join(carry_over_dgs) or "None")
        ws.cell(row=7, column=1).alignment = text_align
        ws.row_dimensions[7].height = 40

        # Note: We use the improved table function in the main scope


    def _write_exact_summary_layout(ws):
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        import datetime

        bold_fmt = Font(bold=True)
        center_fmt = Alignment(horizontal="center", vertical="center")
        border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
        
        fill_blue = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
        fill_cyan = PatternFill(start_color="00B0F0", end_color="00B0F0", fill_type="solid")
        fill_green = PatternFill(start_color="E2F0D9", end_color="E2F0D9", fill_type="solid")
        fill_orange = PatternFill(start_color="FCD5B4", end_color="FCD5B4", fill_type="solid")
        fill_grey = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")
        fill_yellow = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
        
        c_carry_over = status_counts.get("Carry-Over", 0)
        c_modified = status_counts.get("Modified", 0)
        c_added = status_counts.get("Newly Added", 0)
        c_replaced = status_counts.get("Replaced", 0)
        c_deleted = status_counts.get("Deleted", 0)
        c_exclusive = c_modified + c_added + c_replaced
        c_total = c_carry_over + c_exclusive
        
        m_carry_over = _module_count(submodule_data, ["Carry-Over"])
        m_exclusive = _module_count(submodule_data, ["Modified", "Newly Added", "Replaced"])
        m_modified = _module_count(submodule_data, ["Modified"])
        m_added = _module_count(submodule_data, ["Newly Added"])
        m_replaced = _module_count(submodule_data, ["Replaced"])
        m_deleted = _module_count(submodule_data, ["Deleted"])
        
        l_carry_over = _module_list(submodule_data, ["Carry-Over"])
        l_exclusive = _module_list(submodule_data, ["Modified", "Newly Added", "Replaced"])
        l_modified = _module_list(submodule_data, ["Modified"])
        l_added = _module_list(submodule_data, ["Newly Added"])
        l_replaced = _module_list(submodule_data, ["Replaced"])
        l_deleted = _module_list(submodule_data, ["Deleted"])
        
        ws.column_dimensions["C"].width = 30
        ws.column_dimensions["D"].width = 12
        ws.column_dimensions["F"].width = 30
        ws.column_dimensions["G"].width = 12
        ws.column_dimensions["H"].width = 12
        ws.column_dimensions["I"].width = 12
        
        now_str = datetime.datetime.now().strftime("Date- %d %B %Y")
        
        # Format the top header
        ws.merge_cells("C3:F3")
        cell_zero = ws.cell(row=3, column=3, value=0)
        cell_zero.fill = fill_blue
        cell_zero.font = Font(bold=True, color="FFFFFF")
        cell_zero.alignment = center_fmt
        for col in [4, 5, 6]:
            ws.cell(row=3, column=col).fill = fill_blue

        ws.merge_cells("G3:I3")
        cell_date = ws.cell(row=3, column=7, value=now_str)
        cell_date.fill = fill_blue
        cell_date.font = Font(bold=True, color="FFFFFF", italic=True)
        cell_date.alignment = center_fmt
        for col in [8, 9]:
            ws.cell(row=3, column=col).fill = fill_blue
            
        ws.cell(row=5, column=3, value="Total Sub module's").font = bold_fmt; ws.cell(row=5, column=3).fill = fill_cyan
        ws.cell(row=5, column=4, value=c_total).font = bold_fmt; ws.cell(row=5, column=4).fill = fill_cyan
        ws.cell(row=5, column=6, value="Carry-Over Sub module's").font = bold_fmt; ws.cell(row=5, column=6).fill = fill_green
        ws.cell(row=5, column=7, value=c_carry_over).font = bold_fmt; ws.cell(row=5, column=7).fill = fill_green
        ws.cell(row=5, column=8, value="Module's").font = bold_fmt; ws.cell(row=5, column=8).fill = fill_green
        ws.cell(row=5, column=9, value=m_carry_over).font = bold_fmt; ws.cell(row=5, column=9).fill = fill_green
        
        ws.cell(row=6, column=3, value="Carry-Over Sub module's").fill = fill_green
        ws.cell(row=6, column=4, value=c_carry_over).font = bold_fmt; ws.cell(row=6, column=4).fill = fill_green
        ws.merge_cells("F6:I6")
        ws.cell(row=6, column=6, value=l_carry_over).alignment = Alignment(wrap_text=True)
        ws.row_dimensions[6].height = 30
        
        ws.cell(row=7, column=3, value="Exclusive Sub module's (A+B+C)").font = bold_fmt; ws.cell(row=7, column=3).fill = fill_orange
        ws.cell(row=7, column=4, value=c_exclusive).font = bold_fmt; ws.cell(row=7, column=4).fill = fill_orange
        
        ws.cell(row=9, column=3, value="A. Modified Sub module's").fill = fill_grey
        ws.cell(row=9, column=4, value=c_modified).font = bold_fmt; ws.cell(row=9, column=4).fill = fill_grey
        ws.cell(row=9, column=6, value="Exclusive Sub module's").font = bold_fmt; ws.cell(row=9, column=6).fill = fill_orange
        ws.cell(row=9, column=7, value=c_exclusive).font = bold_fmt; ws.cell(row=9, column=7).fill = fill_orange
        ws.cell(row=9, column=8, value="Module's *").font = bold_fmt; ws.cell(row=9, column=8).fill = fill_orange
        ws.cell(row=9, column=9, value=m_exclusive).font = bold_fmt; ws.cell(row=9, column=9).fill = fill_orange
        
        ws.cell(row=10, column=3, value="B. Newly Added Sub module's").fill = fill_grey
        ws.cell(row=10, column=4, value=c_added).font = bold_fmt; ws.cell(row=10, column=4).fill = fill_grey
        ws.merge_cells("F10:I10")
        ws.cell(row=10, column=6, value=l_exclusive).alignment = Alignment(wrap_text=True)
        ws.row_dimensions[10].height = 30
        
        ws.cell(row=11, column=3, value="C. Replaced Sub module's").fill = fill_grey
        ws.cell(row=11, column=4, value=c_replaced).font = bold_fmt; ws.cell(row=11, column=4).fill = fill_grey
        
        ws.cell(row=12, column=6, value="A. Modified Sub module's").font = bold_fmt; ws.cell(row=12, column=6).fill = fill_grey
        ws.cell(row=12, column=7, value=c_modified).font = bold_fmt; ws.cell(row=12, column=7).fill = fill_grey
        ws.cell(row=12, column=8, value="Module's").font = bold_fmt; ws.cell(row=12, column=8).fill = fill_grey
        ws.cell(row=12, column=9, value=m_modified).font = bold_fmt; ws.cell(row=12, column=9).fill = fill_grey
        ws.merge_cells("F13:I13")
        ws.cell(row=13, column=6, value=l_modified).alignment = Alignment(wrap_text=True)
        
        ws.cell(row=13, column=3, value="Note- \n* The exclusive module count is subject to change, given that some modules are present across A, B, and C categories").font = Font(italic=True, color="0070C0")
        ws.merge_cells("C13:E15")
        ws.cell(row=13, column=3).alignment = Alignment(wrap_text=True, vertical="top")
        
        ws.cell(row=15, column=6, value="B. Newly Added Sub module's").font = bold_fmt; ws.cell(row=15, column=6).fill = fill_grey
        ws.cell(row=15, column=7, value=c_added).font = bold_fmt; ws.cell(row=15, column=7).fill = fill_grey
        ws.cell(row=15, column=8, value="Module's").font = bold_fmt; ws.cell(row=15, column=8).fill = fill_grey
        ws.cell(row=15, column=9, value=m_added).font = bold_fmt; ws.cell(row=15, column=9).fill = fill_grey
        ws.merge_cells("F16:I16")
        ws.cell(row=16, column=6, value=l_added).alignment = Alignment(wrap_text=True)
        ws.row_dimensions[16].height = 30
        
        ws.cell(row=18, column=6, value="C. Replaced Sub module's").font = bold_fmt; ws.cell(row=18, column=6).fill = fill_grey
        ws.cell(row=18, column=7, value=c_replaced).font = bold_fmt; ws.cell(row=18, column=7).fill = fill_grey
        ws.cell(row=18, column=8, value="Module's").font = bold_fmt; ws.cell(row=18, column=8).fill = fill_grey
        ws.cell(row=18, column=9, value=m_replaced).font = bold_fmt; ws.cell(row=18, column=9).fill = fill_grey
        ws.merge_cells("F19:I19")
        ws.cell(row=19, column=6, value=l_replaced).alignment = Alignment(wrap_text=True)
        
        ws.merge_cells("C21:I21")
        cell = ws.cell(row=21, column=3, value="Base VC Submodules deleted in New VC")
        cell.font = bold_fmt
        cell.fill = fill_yellow
        cell.alignment = center_fmt
        
        ws.cell(row=23, column=3, value="Deleted Sub module's").fill = fill_yellow
        ws.cell(row=23, column=4, value=c_deleted).font = bold_fmt; ws.cell(row=23, column=4).fill = fill_yellow
        ws.cell(row=23, column=6, value="Deleted Sub module's").font = bold_fmt; ws.cell(row=23, column=6).fill = fill_yellow
        ws.cell(row=23, column=7, value=c_deleted).font = bold_fmt; ws.cell(row=23, column=7).fill = fill_yellow
        ws.cell(row=23, column=8, value="Module's").font = bold_fmt; ws.cell(row=23, column=8).fill = fill_yellow
        ws.cell(row=23, column=9, value=m_deleted).font = bold_fmt; ws.cell(row=23, column=9).fill = fill_yellow
        ws.merge_cells("F24:I24")
        ws.cell(row=24, column=6, value=l_deleted).alignment = Alignment(wrap_text=True)
        ws.row_dimensions[24].height = 30
        
        for row in range(3, 26):
            for col in range(3, 10):
                ws.cell(row=row, column=col).border = border
        
        ws.sheet_view.showGridLines = False

    # --- START REORGANIZED SUMMARY SHEETS ---
    # 1. Summary (Detailed Box Layout)
    ws_detail = wb.create_sheet(title="Summary", index=0)
    _write_exact_summary_layout(ws_detail)

    # 2. Submodule Summary (Table Layout)
    ws = wb.create_sheet(title="Submodule Summary", index=1)
    # Improve the submodule table by adding a "Total Changes" column
    def _write_improved_submodule_table(ws, table_start_row: int):
        headers = [
            "Design Group", "Submodule Address", "Component Number", "Description", "Location",
            "Primary Status", "Total Changes", 
            "Total Added", "Total Deleted", "Total Replaced", "Total Modified", "Total Carry-Over",
        ]
        col_widths = [14, 20, 25, 40, 20, 15, 15, 12, 12, 12, 12, 15]

        for c, (header, width) in enumerate(zip(headers, col_widths), 1):
            cell = ws.cell(row=table_start_row, column=c, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_align
            cell.border = thin_border
            ws.column_dimensions[get_column_letter(c)].width = width

        row_num = table_start_row + 1
        for item in submodule_data:
            ws.cell(row=row_num, column=1, value=item.get("Design Group"))
            ws.cell(row=row_num, column=2, value=item.get("Submodule Address"))
            ws.cell(row=row_num, column=3, value=item.get("Component Number"))
            ws.cell(row=row_num, column=4, value=item.get("Description"))
            ws.cell(row=row_num, column=5, value=item.get("Location"))
            ws.cell(row=row_num, column=6, value=item.get("Primary Status"))
            
            # Total Changes = Added + Deleted + Replaced + Modified
            tc = (item.get("Total Added", 0) + item.get("Total Deleted", 0) + 
                  item.get("Total Replaced", 0) + item.get("Total Modified", 0))
            ws.cell(row=row_num, column=7, value=tc).font = Font(bold=True)
            
            ws.cell(row=row_num, column=8, value=item.get("Total Added")).fill = count_fills["Added"]
            ws.cell(row=row_num, column=9, value=item.get("Total Deleted")).fill = count_fills["Deleted"]
            ws.cell(row=row_num, column=10, value=item.get("Total Replaced")).fill = count_fills["Replaced"]
            ws.cell(row=row_num, column=11, value=item.get("Total Modified")).fill = count_fills["Modified"]
            ws.cell(row=row_num, column=12, value=item.get("Total Carry-Over")).fill = count_fills["Carry-Over"]
            
            for c in range(1, 13):
                ws.cell(row=row_num, column=c).border = thin_border
            row_num += 1

    # Overriding _write_summary_layout logic slightly for the "Submodule Summary" sheet
    ws.cell(row=1, column=1, value="SUBMODULE CHANGE SUMMARY").font = title_font
    
    # Improved breakdown section at the top of Submodule Summary
    ws.cell(row=3, column=1, value="CARRY-OVER SUBMODULES:").font = text_font
    ws.merge_cells(start_row=4, start_column=1, end_row=4, end_column=12)
    ws.cell(row=4, column=1, value=", ".join(carry_over_dgs) or "None").alignment = text_align
    ws.cell(row=4, column=1).border = thin_border
    ws.row_dimensions[4].height = 40

    ws.cell(row=6, column=1, value="EXCLUSIVE SUBMODULES:").font = text_font
    ws.merge_cells(start_row=7, start_column=1, end_row=7, end_column=12)
    all_changed = added_dgs + deleted_dgs + replaced_dgs + modified_dgs
    ws.cell(row=7, column=1, value=", ".join(dict.fromkeys(all_changed)) or "None").alignment = text_align
    ws.cell(row=7, column=1).border = thin_border
    ws.row_dimensions[7].height = 40

    ws.cell(row=9, column=1, value="MODIFIED SUBMODULES:").font = text_font
    ws.merge_cells(start_row=10, start_column=1, end_row=10, end_column=12)
    ws.cell(row=10, column=1, value=", ".join(modified_dgs) or "None").alignment = text_align
    ws.cell(row=10, column=1).border = thin_border
    ws.row_dimensions[10].height = 40

    ws.cell(row=12, column=1, value="NEWLY ADDED SUBMODULES:").font = text_font
    ws.merge_cells(start_row=13, start_column=1, end_row=13, end_column=12)
    ws.cell(row=13, column=1, value=", ".join(added_dgs) or "None").alignment = text_align
    ws.cell(row=13, column=1).border = thin_border
    ws.row_dimensions[13].height = 40

    ws.cell(row=15, column=1, value="REPLACED SUBMODULES:").font = text_font
    ws.merge_cells(start_row=16, start_column=1, end_row=16, end_column=12)
    ws.cell(row=16, column=1, value=", ".join(replaced_dgs) or "None").alignment = text_align
    ws.cell(row=16, column=1).border = thin_border
    ws.row_dimensions[16].height = 40

    ws.cell(row=18, column=1, value="DELETED SUBMODULES:").font = text_font
    ws.merge_cells(start_row=19, start_column=1, end_row=19, end_column=12)
    ws.cell(row=19, column=1, value=", ".join(deleted_dgs) or "None").alignment = text_align
    ws.cell(row=19, column=1).border = thin_border
    ws.row_dimensions[19].height = 40
    
    _write_improved_submodule_table(ws, table_start_row=21)


    def _write_groups_sheets(wb):
        from openpyxl.styles import Font, Alignment
        from openpyxl.utils import get_column_letter
        header_font = Font(bold=True)
        
        def _create_group_sheet(sheet_title, status_filters):
            ws = wb.create_sheet(title=sheet_title)
            # Filter items for this status, excluding those marked for exclusion
            items = [i for i in submodule_data if not i.get("exclude_from_counts") and i.get("Computed_Status", i.get("Primary Status", "")) in status_filters]
            
            # Write headers for columns 1 to 40 (Sr.No, Design Group, Exclusive Submodule)
            # We repeat the pattern (Blank, Sr.No, Design Group, Exclusive Submodule)
            col_offset = 1
            for i in range(10):
                ws.cell(row=2, column=col_offset + 1, value="Sr.No").font = header_font
                ws.cell(row=2, column=col_offset + 2, value="Design Group").font = header_font
                ws.cell(row=2, column=col_offset + 3, value="Exclusive Submodule").font = header_font
                ws.column_dimensions[get_column_letter(col_offset + 1)].width = 6
                ws.column_dimensions[get_column_letter(col_offset + 2)].width = 15
                ws.column_dimensions[get_column_letter(col_offset + 3)].width = 25
                col_offset += 4
                
            # Write items in rows, filling horizontally first, or vertically first?
            # Legacy layout fills vertically by block. Let's do a simple vertical fill per column block.
            rows_per_col = 40
            for idx, item in enumerate(items):
                block = idx // rows_per_col
                row_in_block = idx % rows_per_col
                c_base = block * 4 + 1
                r = row_in_block + 3
                
                ws.cell(row=r, column=c_base + 1, value=idx + 1)
                ws.cell(row=r, column=c_base + 2, value=item.get("Design Group", ""))
                ws.cell(row=r, column=c_base + 3, value=item.get("Component Number", ""))
                
        _create_group_sheet("Carryover Groups", ["Carry-Over"])
        _create_group_sheet("Exclusive Groups", ["Modified", "Newly Added", "Replaced"])
        _create_group_sheet("Modified Groups", ["Modified"])
        _create_group_sheet("Newly Added Groups", ["Newly Added"])
        _create_group_sheet("Replaced Groups", ["Replaced"])
        _create_group_sheet("Deleted Groups", ["Deleted"])

    _write_groups_sheets(wb)

    try:
        wb.save(excel_path)
        print(f"Successfully updated {excel_path} with Submodule Summary and Summary sheets.")
    except Exception as e:
        print(f"Error: Could not save updated Excel file {excel_path}. {e}", file=sys.stderr)


# ---------------------------
# CLI / Main
# ---------------------------
def main():
    parser = argparse.ArgumentParser(description="Generate a BOM JSON tree from a comparator file.")
    parser.add_argument("--input", "-i", required=True, help="Path to the comparator Excel file (e.g., BOM_Comparison_Result.xlsx).")
    parser.add_argument("--output", "-o", required=True, help="Path for the output JSON hierarchy file (e.g., bom_hierarchy.json).")
    parser.add_argument("--submodules", required=True, help="Path to the submodule structure JSON file.")
    parser.add_argument("--config_path", required=True, help="Path to the config.json file (e.g., config91112.json) for lookups.")
    parser.add_argument("--sheet", default="All", help="Sheet to read from the input Excel file.")
    parser.add_argument("--no-debug", action="store_true", help="Disable debug printing.")
    
    parser.add_argument("--summary_json_path", default="summary.json", help="Path to the summary.json file to update.")
    parser.add_argument("--excel_report_path", default=None, help="Path to the Excel report file to update. Defaults to --input path.")
    parser.add_argument("--manual_overrides_path", default=None, help="Path to the original Excel file containing manual override sheets.")
    
    args = parser.parse_args()

    excel_report_path = args.excel_report_path if args.excel_report_path else args.input
    summary_json_path = args.summary_json_path
    
    try:
        with open(summary_json_path, 'r', encoding='utf-8') as f:
            summary_data = json.load(f)
        root_level = summary_data.get("Overview", {}).get("Root Level", 1)
        print(f"Read Root Level: {root_level} from {summary_json_path}")
    except Exception as e:
        print(f"Warning: Could not read root_level from {summary_json_path}. Defaulting to 1. Error: {e}", file=sys.stderr)
        root_level = 1

    try:
        merged_tree = build_merged_tree(args.input, args.submodules, args.config_path, sheet_name=args.sheet, debug=not args.no_debug, root_level=root_level)
        
        print(f"\nSuccessfully generated JSON tree with {len(merged_tree)} root(s).")
        with open(args.output, "w", encoding="utf-8") as f:
            json.dump(merged_tree, f, indent=2, ensure_ascii=False)
        print(f"Output saved to: {args.output}")

        print("Loading manual overrides from Excel (if any)...")
        overrides_path = args.manual_overrides_path if args.manual_overrides_path else excel_report_path
        manual_overrides = load_manual_overrides(overrides_path)

        print("Generating submodule summary from hierarchy...")
        submodule_summary = generate_submodule_summary_data(merged_tree, manual_overrides)
        
        if submodule_summary:
            print(f"Found {len(submodule_summary)} submodules. Updating summary and Excel files...")
            update_summary_json(summary_json_path, submodule_summary)
            update_excel_report(excel_report_path, submodule_summary)
        else:
            print("No submodules found in hierarchy, skipping summary/Excel updates.")

    except (FileNotFoundError, ValueError) as e:
        print(f"\nError: {e}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"\nAn unexpected error occurred: {e}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    main()

