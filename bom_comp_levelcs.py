#!/usr/bin/env python3

"""
Unified BOM Comparator â€” tree + semantic + config-segmented compare

MODIFIED (Dec 2025 - Fix 4):
- FIX: 'Exact Match' (Step 1) now keys off (Level, Component, IMMEDIATE_PARENT_ID) instead of the full ParentPath string.
       This prevents "Root Changes" or "Grandparent Changes" from breaking the strict match for children.
       This ensures Child(Parent A) matches Child(Parent A) even if the Root moves, preventing it from incorrectly grabbing Child(Parent B).
- FIX: 'Parent Path Changed' logic uses DIRECT NODE ACCESS (node.parent) for 100% precision.
- Added _UNIFIED_COL_ALIAS_MAP for standard mode comparison.
- Added Safe Mapper fix to prevent crashes on empty cells/integers.
- Added Timestamp fix (12-hour format).
- Added Dynamic Pair logic for mapped optional columns.
- Logic to handle both unified and split mappings from UI.
- Preserved all original SBERT, Fuzzy, and Export logic.
"""

import os
import re
import io
import json
import time
import hashlib
import pickle
import argparse
import math 
from collections import defaultdict, Counter
from typing import List, Dict, Any, Tuple, Optional
from datetime import datetime, timezone
import numpy as np
import pandas as pd

# Optional libs
try:
    from rapidfuzz import fuzz # noqa: F401
except Exception:
    fuzz = None

try:
    import faiss # type: ignore
    _FAISS_OK = True
except Exception:
    faiss = None # type: ignore
    _FAISS_OK = False

import torch
from sentence_transformers import SentenceTransformer

# ---------------------- USER CONFIG / DEFAULTS ----------------------

DEFAULT_OLD = "CS12-26471034000R.xlsx"
DEFAULT_NEW = "CS12-26471534000R.xlsx"
DEFAULT_OUTPUT = "BOM_Comparison_Result_LevelCs.xlsx"
DEFAULT_CONFIG = "config91112.json"

CAN_COL_LEVEL = "Level"
CAN_COL_COMP = "Component number"
CAN_COL_DESC = "Object description"
CAN_COL_CS = "Special procurement"
CAN_COL_QTY = "Component quantity"

APP_DIR = os.path.dirname(os.path.abspath(__file__))
LOCAL_SBERT_PATH = os.path.join(APP_DIR, "models", "all-MiniLM-L6-v2")  # set to None to download by name
MODEL_NAME = "sentence-transformers/all-MiniLM-L6-v2"
DEVICE = "cuda" if torch.cuda.is_available() else "cpu"

EMB_CACHE_DIR = os.path.join(APP_DIR, ".emb_cache")
os.makedirs(EMB_CACHE_DIR, exist_ok=True)

# Toggle: also output structured per-segment columns for Old/New components
OUTPUT_STRUCTURED_SEGMENTS = False

# ---------------------- UTIL: column mapping with allowlist ----------------------

# --- 1. SAP (OLD BOM) DEFAULTS ---
_SAP_COL_ALIAS_MAP = {
    CAN_COL_LEVEL: {"Level", "lvl", "lev", "bom level", "Hierarchy Lvl"},
    CAN_COL_COMP: {"Material", "Component", "Component number", "Sub-Item", "Object"},
    CAN_COL_DESC: {"Object description", "Description", "Material Description", "Short Text"},
    CAN_COL_CS: {"Special procurement", "Special Proc.", "SpPr", "SOBS", "Make Buy", "Pune Make Buy", "Jamshedpur Make Buy", "Dharwad Make Buy"},
    CAN_COL_QTY: {"Component quantity", "Qty", "Quantity", "Unf"},
}

# --- 2. PLM (NEW BOM) DEFAULTS ---
_PLM_COL_ALIAS_MAP = {
    CAN_COL_LEVEL: {"Level", "BOM Level", "Structure Level"},
    CAN_COL_COMP: {"Item ID", "Part Number", "Item Id", "Component", "Child Item"},
    CAN_COL_DESC: {"Item Name", "Item Description", "Object description", "Title"},
    CAN_COL_CS: {"Special procurement", "Make/Buy", "Procurement Type", "CS", "Supply Type"},
    CAN_COL_QTY: {"Quantity", "Find Quantity", "Qty", "Amount"},
}

# --- 3. UNIFIED (Standard Mode) ---
# Merges both lists to auto-detect if user doesn't specify mappings
_UNIFIED_COL_ALIAS_MAP = {
    CAN_COL_LEVEL: _SAP_COL_ALIAS_MAP[CAN_COL_LEVEL] | _PLM_COL_ALIAS_MAP[CAN_COL_LEVEL],
    CAN_COL_COMP: _SAP_COL_ALIAS_MAP[CAN_COL_COMP] | _PLM_COL_ALIAS_MAP[CAN_COL_COMP],
    CAN_COL_DESC: _SAP_COL_ALIAS_MAP[CAN_COL_DESC] | _PLM_COL_ALIAS_MAP[CAN_COL_DESC],
    CAN_COL_CS: _SAP_COL_ALIAS_MAP[CAN_COL_CS] | _PLM_COL_ALIAS_MAP[CAN_COL_CS],
    CAN_COL_QTY: _SAP_COL_ALIAS_MAP[CAN_COL_QTY] | _PLM_COL_ALIAS_MAP[CAN_COL_QTY],
}

def map_columns(df: pd.DataFrame, allowed_columns: Optional[List[str]] = None, alias_map: Dict[str, set] = None) -> Dict[str, str]:
    if alias_map is None:
        alias_map = _UNIFIED_COL_ALIAS_MAP # Default to Unified if not specified

    cols_lower = {c.lower(): c for c in df.columns}
    mapping: Dict[str, str] = {}
    
    for canon, aliases in alias_map.items():
        if allowed_columns and canon not in allowed_columns:
            continue
        found = None
        # exact case-insensitive match first
        for a in aliases:
            if a.lower() in cols_lower:
                found = cols_lower[a.lower()]
                break
        # stricter fallback: startswith or word-boundary, not arbitrary substring
        if not found:
            for k, orig in cols_lower.items():
                for a in aliases:
                    al = a.lower()
                    if k.startswith(al) or re.search(r"\b" + re.escape(al) + r"\b", k):
                        found = orig
                        break
                if found:
                    break
        if found:
            mapping[canon] = found
    return mapping

# ---------------------- LOAD & NORMALIZE ----------------------

def load_and_preprocess(path: str) -> pd.DataFrame:
    if not os.path.exists(path):
        raise FileNotFoundError(f"Input file not found: {path}")
    # FORCE STRING TYPE: Prevents long IDs being read as scientific notation floats
    df = pd.read_excel(path, engine=None, dtype=str)
    return df

def _clean_cs_code(code):
    """
    Normalizes Special Procurement codes.
    """
    if pd.isna(code): return "F"
    s = str(code).strip().upper()
    if not s or s in ["N/A", "NAN", "NONE", "BLANK"]: return "F"
    s = re.split(r'[\s(]', s)[0]
    
    mapping = {
        "99": "E99", "98": "E98", "50": "E50",
        "18": "F18", "19": "F19", "30": "F30", "40": "F40",
        "E": "E", "F": "F"
    }
    
    if s in mapping: return mapping[s]
    try:
        f = float(s)
        if f.is_integer():
            i_str = str(int(f))
            if i_str in mapping: return mapping[i_str]
            return i_str 
    except ValueError:
        pass
    return s

def normalize_bom_dataframe(df_raw: pd.DataFrame, allowed_columns: Optional[List[str]] = None, user_map: Optional[Dict[str, str]] = None, default_alias_map: Dict[str, set] = None) -> Tuple[pd.DataFrame, Dict[str, str]]:
    df = df_raw.copy()
    
    # Deduplicate raw columns first to avoid initial naming issues
    df = df.loc[:, ~df.columns.duplicated(keep='first')]
    
    mapping = {}

    if user_map:
        # NEW: Use user-provided mappings
        # Invert map for renaming: { "UserColumnName": "CanonicalName" }
        inverted_map = {v: k for k, v in user_map.items() if v in df.columns}
        
        # Drop columns that would collide with the renaming target names to avoid duplicates
        target_names = set(inverted_map.values())
        cols_to_drop = [c for c in df.columns if c in target_names and c not in inverted_map.keys()]
        if cols_to_drop:
            df.drop(columns=cols_to_drop, inplace=True)
            
        df.rename(columns=inverted_map, inplace=True)
        # The mapping is now the user's map, but we'll create a full one for consistency
        mapping = {k: v for k, v in user_map.items()}
    else:
        # OLD: Fallback to alias mapping using the specific map provided (Unified/SAP/PLM)
        mapping = map_columns(df_raw, allowed_columns, alias_map=default_alias_map)
        # We need to rename columns in the copied df to canonical names for consistent processing
        inverted_map = {v: k for k, v in mapping.items()}
        
        # Drop columns that would collide with the renaming target names to avoid duplicates
        target_names = set(inverted_map.values())
        cols_to_drop = [c for c in df.columns if c in target_names and c not in inverted_map.keys()]
        if cols_to_drop:
            df.drop(columns=cols_to_drop, inplace=True)
            
        df.rename(columns=inverted_map, inplace=True)

    # Ensure all canonical columns exist to prevent KeyErrors, fill with defaults if missing
    for canon_col, default_val in [
        (CAN_COL_LEVEL, 0),
        (CAN_COL_COMP, ""),
        (CAN_COL_DESC, ""),
        (CAN_COL_CS, ""),
        (CAN_COL_QTY, 0.0)
    ]:
        if canon_col not in df.columns:
            df[canon_col] = default_val

    # Final column deduplication to guarantee they are Series (avoiding DataFrame return on clash)
    df = df.loc[:, ~df.columns.duplicated(keep='first')]

    lev = pd.to_numeric(df[CAN_COL_LEVEL], errors="coerce").fillna(0).astype(int)
    
    # --- CLEANING LOGIC ---
    # This is where '123_456' becomes '123'
    raw_comp = df[CAN_COL_COMP].fillna("").astype(str).str.strip()
    comp = pd.Series(raw_comp, index=df.index).str.split('_', n=1).str[0]
    
    # This is where description becomes UPPERCASE
    desc = df[CAN_COL_DESC].fillna("").astype(str).str.upper().str.strip()
    # --- END CLEANING ---

    raw_cs = df[CAN_COL_CS]
    cs = raw_cs.apply(_clean_cs_code).astype(str)
    
    qty = pd.to_numeric(df[CAN_COL_QTY], errors="coerce").fillna(0.0)

    extra_cols = {}
    if allowed_columns:
        for col in allowed_columns:
            # Check original column names for extras
            original_col_name = next((user_col for canon, user_col in (user_map or {}).items() if canon == col), col)
            if original_col_name not in (user_map or {}).values() and original_col_name in df_raw.columns:
                 extra_cols[original_col_name] = df_raw[original_col_name]

    out = pd.DataFrame({
        CAN_COL_LEVEL: lev,
        CAN_COL_COMP: comp, # Cleaned component number
        CAN_COL_DESC: desc, # Cleaned description
        CAN_COL_CS: cs,
        CAN_COL_QTY: qty,
        **extra_cols
    }, index=df.index)
    
    # The 'mapping' returned should reflect from canonical to what was in the user's file
    # This is crucial for the new restore logic
    final_mapping = {k: v for k, v in (user_map or mapping).items()}
    return out, final_mapping

# --- NEW: Root Pruning Logic ---
def prune_root_level(df: pd.DataFrame) -> pd.DataFrame:
    """
    Filters out rows at the minimum level (root) and strips the first segment from ParentPaths.
    This aligns a Level-0 based BOM with a Level-1 based BOM.
    """
    if df.empty:
        return df
        
    min_level = df[CAN_COL_LEVEL].min()
    
    # 1. Filter out the root rows
    df_pruned = df[df[CAN_COL_LEVEL] > min_level].copy()
    
    print(f"Pruned {len(df) - len(df_pruned)} root nodes (Level {min_level}). New count: {len(df_pruned)}")
    return df_pruned

# ---------------------- TREE: BOMNode & helpers ----------------------

class BOMNode:
    def __init__(self, level, comp, desc, cs, qty, source_idx: Optional[int] = None):
        self.level = int(level)
        self.component_number = str(comp)
        self.object_description = str(desc)
        self.cs_code = str(cs)
        self.quantity = float(qty)
        self.children: List["BOMNode"] = []
        self.source_idx = source_idx # This is the key to linking back to df_raw
        self.primary_status = "Unchanged"
        self.hierarchical_status = "Unchanged"
        self.change_details = ""
        self.parent: Optional["BOMNode"] = None # Type changed to Optional["BOMNode"]

    def add_child(self, child: "BOMNode"):
        self.children.append(child)
        child.parent = self

def build_forest(df: pd.DataFrame) -> List[BOMNode]:
    roots: List[BOMNode] = []
    stack: List[BOMNode] = []

    for idx, row in df.iterrows():
        # idx is the original index from df_raw, which is preserved in flat_old_norm
        node = BOMNode(row[CAN_COL_LEVEL], row[CAN_COL_COMP], row[CAN_COL_DESC], row[CAN_COL_CS], row[CAN_COL_QTY], source_idx=idx)

        if not stack:
            roots.append(node)
            stack = [node]
            continue
        
        if node.level <= stack[0].level:
             roots.append(node)
             stack = [node]
             continue

        while stack and stack[-1].level >= node.level:
            stack.pop()

        if stack:
            stack[-1].add_child(node)
        else:
            roots.append(node)

        stack.append(node)

    return roots

# ---------------------- Special Procurement Skipping ----------------------

def flatten_forest(roots: List[BOMNode], user_choice: str) -> pd.DataFrame:
    rows: List[Dict[str, Any]] = []

    def walk(node: BOMNode, parent_path: str, parent_cs: str):
        skip_here = False
        cs = (node.cs_code or "").strip().upper()

        parent_cs_norm = (parent_cs or "").strip().upper()
        if user_choice.upper() == "Y":
            # Treat empty parent CS code as 'F'
            if parent_cs_norm == "F" or parent_cs_norm == "":
                if cs in {"F18", "F40", "F19"}:
                    skip_here = True
                elif cs == "F30":
                    has_e_child = any(((c.cs_code or "").strip().upper().startswith("E")) for c in node.children)
                    skip_here = not has_e_child

        rows.append({
            CAN_COL_LEVEL: node.level,
            CAN_COL_COMP: node.component_number,
            CAN_COL_DESC: node.object_description,
            CAN_COL_CS: node.cs_code or "",
            CAN_COL_QTY: node.quantity,
            "ParentPath": parent_path,
            "SkippedChildren": skip_here,
            "SourceIndex": node.source_idx # This is the key
        })

        this_path = f"{parent_path}/{node.component_number}" if parent_path else node.component_number

        if not skip_here:
            for c in node.children:
                walk(c, this_path, cs)

    for r in roots:
        walk(r, "", "")
    
    # This DataFrame has the cleaned component/desc but the original SourceIndex
    df = pd.DataFrame(rows)
    # Set the index to be the SourceIndex for easy .loc lookup later
    if "SourceIndex" in df.columns:
        # Handle potential duplicate source indexes from bad data by keeping first
        df = df[~df['SourceIndex'].duplicated(keep='first')]
        df = df.set_index("SourceIndex", drop=False)
        df.index.name = "OriginalIndex"
    return df

# ---------------------- CONFIG & COMPONENT PARSING ----------------------

def load_config(path: str) -> Dict[str, Any]:
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            cfg = json.load(f)
    else:
        cfg = {}

    cfg.setdefault("semantic_threshold", 0.85)
    cfg.setdefault("replace_threshold", 0.70)
    cfg.setdefault("root_replace_threshold", 0.30)
    cfg.setdefault("weights", {"description": 0.6, "segments": 0.4})
    cfg.setdefault("lengths", {})
    cfg.setdefault("lookups", {})
    cfg.setdefault("divergence_weights", {"added": 1.0, "deleted": 1.0, "changed": 1.0, "qty": 0.2})
    cfg.setdefault("allowed_columns", [CAN_COL_LEVEL, CAN_COL_COMP, CAN_COL_DESC, CAN_COL_CS, CAN_COL_QTY])

    return cfg

def _is_int_string(s: str) -> bool:
    return re.fullmatch(r"\d+", s) is not None

def _in_numeric_range(val_str: str, range_key: str) -> bool:
    try:
        a, b = range_key.split("-")
        if _is_int_string(a) and _is_int_string(b) and _is_int_string(val_str):
            return int(a) <= int(val_str) <= int(b)
    except Exception:
        pass
    return False

def lookup_label(lookup_dict: Dict[str, str], segment_value: str):
    if not lookup_dict:
        return None

    if segment_value in lookup_dict:
        return lookup_dict[segment_value]

    for k, v in lookup_dict.items():
        if "-" in k and _in_numeric_range(segment_value, k):
            return v

    return None

def _norm_comp(s: Any) -> str:
    # do not strip leading zeros; remove whitespace only
    return re.sub(r"\s+", "", str(s)).strip()

def parse_component_number(comp_num: str, cfg: Dict[str, Any]) -> Dict[str, Any]:
    s = _norm_comp(comp_num)
    length_key = str(len(s))
    out = {"length": len(s), "segments": []}

    rule = cfg.get("lengths", {}).get(length_key)
    if not rule:
        out["segments"].append({"name": "full", "value": s, "label": None})
        return out

    lookups_all = cfg.get("lookups", {})
    for seg in rule.get("segments", []):
        start = seg["start"]
        end = seg["end"]
        val = s[start:end] if end <= len(s) else ""
        lbl = None
        if "lookup" in seg:
            lbl = lookup_label(lookups_all.get(seg["lookup"], {}), val)
        out["segments"].append({"name": seg["name"], "value": val, "label": lbl})

    return out

def segment_match_score(parsed_old: Dict[str, Any], parsed_new: Dict[str, Any]) -> float:
    segs_old = parsed_old.get("segments", [])
    segs_new = parsed_new.get("segments", [])
    n = min(len(segs_old), len(segs_new))

    if n == 0:
        return 1.0 if parsed_old.get("length") == parsed_new.get("length") else 0.0

    score = 0.0
    for i in range(n):
        vo, vn = segs_old[i]["value"], segs_new[i]["value"]
        lo, ln = segs_old[i].get("label"), segs_new[i].get("label")

        if vo == vn and vo != "":
            score += 1.0
        elif lo and ln and lo == ln:
            score += 0.7
        elif lo and ln:
            score += 0.2

    return score / n

def describe_barrel_changes(parsed_old: Dict[str, Any], parsed_new: Dict[str, Any]) -> List[str]:
    changes = []
    segs_old = parsed_old.get("segments", [])
    segs_new = parsed_new.get("segments", [])
    m = max(len(segs_old), len(segs_new))

    for i in range(m):
        so = segs_old[i] if i < len(segs_old) else {"name": "?", "value": "", "label": None}
        sn = segs_new[i] if i < len(segs_new) else {"name": "?", "value": "", "label": None}

        if so["value"] != sn["value"] or (so.get("label") != sn.get("label")):
            name = sn["name"] if sn["name"] != "?" else so["name"]

            def fmt(seg):
                lab = seg.get("label")
                return f"{seg['value']}{f'({lab})' if lab else ''}".strip()

            left = fmt(so) if so["value"] or so.get("label") else "âˆ…"
            right = fmt(sn) if sn["value"] or sn.get("label") else "âˆ…"
            changes.append(f"{name}: {left} â†’ {right}")

    if parsed_old.get("length") != parsed_new.get("length"):
        changes.append(f"Length: {parsed_old.get('length')} â†’ {parsed_new.get('length')}")

    return changes

# ---------------------- SBERT + CACHING ----------------------

def _stable_df_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    tmp = df.copy()
    for c in tmp.columns:
        tmp[c] = tmp[c].astype(str)
    # Use sort_rows=True for stability
    tmp = tmp.sort_values(by=list(tmp.columns)).reset_index(drop=True)
    tmp.to_parquet(buf, index=False)
    return buf.getvalue()

def _hash_df(df: pd.DataFrame) -> str:
    return hashlib.sha256(_stable_df_bytes(df)).hexdigest()

def embeddings_cache_path(prefix: str, df: pd.DataFrame) -> str:
    return os.path.join(EMB_CACHE_DIR, f"{prefix}_{_hash_df(df)}.pkl")

_model: Optional[SentenceTransformer] = None

def get_model() -> SentenceTransformer:
    global _model
    if _model is None:
        try:
            if LOCAL_SBERT_PATH and os.path.isdir(LOCAL_SBERT_PATH):
                _model = SentenceTransformer(LOCAL_SBERT_PATH, device=DEVICE)
            else:
                raise FileNotFoundError(
                    f"Local SBERT model directory not found: {LOCAL_SBERT_PATH}"
                )
        except Exception as e:
            print(
                f"Error loading SBERT model. Bundle model folder at: {LOCAL_SBERT_PATH}"
            )
            print(f"Error: {e}")
            raise
    return _model

def embed_with_cache_column(df: pd.DataFrame, col: str, cache_prefix: str) -> np.ndarray:
    only = pd.DataFrame({col: df[col].fillna("").astype(str)})
    cache_path = embeddings_cache_path(f"{cache_prefix}_{col}", only)

    if os.path.exists(cache_path):
        try:
            with open(cache_path, "rb") as f:
                return pickle.load(f)
        except Exception:
             print("Invalid cache file. Re-computing embeddings.")

    model = get_model()
    texts = only[col].tolist()
    emb = model.encode(texts, normalize_embeddings=True, show_progress_bar=False, batch_size=64, device=DEVICE)
    emb = np.asarray(emb, dtype=np.float32)

    try:
        with open(cache_path, "wb") as f:
            pickle.dump(emb, f)
    except Exception:
        print("Warning: Could not write to cache file.")

    return emb

def cos_sim_vec(a: np.ndarray, b: np.ndarray) -> float:
    return float(np.dot(a, b))

# ---------------------- NN search helper ----------------------

def nn_search(old_emb_mat: np.ndarray, new_emb_mat: np.ndarray, old_idx: List[int], new_idx: List[int]) -> Tuple[np.ndarray, np.ndarray]:
    old_vecs = old_emb_mat[old_idx].astype(np.float32)
    new_vecs = new_emb_mat[new_idx].astype(np.float32)
    
    if len(old_vecs) == 0 or len(new_vecs) == 0:
            return np.array([]), np.array([])

    if _FAISS_OK:
        dim = old_vecs.shape[1]
        index = faiss.IndexFlatIP(dim) # inner product
        index.add(old_vecs)
        D, I = index.search(new_vecs, k=1)
        return D[:, 0], I[:, 0]
    else:
        sims = np.matmul(new_vecs, old_vecs.T)
        I = np.argmax(sims, axis=1)
        D = sims[np.arange(sims.shape[0]), I]
        return D, I

# ---------------------- MATCHING PIPELINE ----------------------

def exact_match_indices(df_old: pd.DataFrame, df_new: pd.DataFrame) -> Tuple[List[Tuple[int, int]], set, set]:
    
    # --- PASS 1: Strict parent-aware matching (FCFS) ---
    # MODIFIED: Key now uses ONLY the immediate parent ID, ignoring the full path string
    # This ensures matches persist even if Root/Grandparents change (e.g. 39R -> 39R)
    
    grouping_old = defaultdict(list)
    for i, r in df_old.iterrows():
        # Extract immediate parent ID
        full_path = r["ParentPath"] or ""
        # If path exists, grab last segment. If empty, it's a root.
        parent_id = full_path.split('/')[-1] if full_path else ""
        
        key = (r[CAN_COL_LEVEL], str(r[CAN_COL_COMP]), parent_id)
        grouping_old[key].append(i)

    matched: List[Tuple[int, int]] = []
    used_old = set()
    used_new = set()

    for j, r in df_new.iterrows():
        full_path = r["ParentPath"] or ""
        parent_id = full_path.split('/')[-1] if full_path else ""
        
        key = (r[CAN_COL_LEVEL], str(r[CAN_COL_COMP]), parent_id)
        
        if key in grouping_old:
            cands = [idx for idx in grouping_old[key] if idx not in used_old]
            if cands:
                i = cands[0] # FCFS
                matched.append((i, j))
                used_old.add(i)
                used_new.add(j)
    
    rem_old = set(df_old.index) - used_old
    rem_new = set(df_new.index) - used_new

    # --- START: MODIFICATION (Smart Best-Fit with Score) ---
    # --- PASS 2: Fallback parent-agnostic "Greedy Best-Fit" matching ---
    
    # Build a lookup map of available *old* items for efficient searching
    # Note: We group by (Level, Comp) to limit candidates to valid exact ID matches
    grouping_old_agnostic = defaultdict(list)
    for i in rem_old: 
        key = (df_old.at[i, CAN_COL_LEVEL], str(df_old.at[i, CAN_COL_COMP]))
        grouping_old_agnostic[key].append(i)

    # Iterate through *new* remaining items (greedy part, but smart selection)
    # We iterate over a list() copy because we modify the `rem_new` set inside the loop
    for j in list(rem_new):
        key = (df_new.at[j, CAN_COL_LEVEL], str(df_new.at[j, CAN_COL_COMP]))
        
        if key in grouping_old_agnostic:
            # Find all available candidates in the old list
            candidates_old_indices = [idx for idx in grouping_old_agnostic[key] if idx not in used_old]
            
            if not candidates_old_indices:
                continue # No available old items for this key

            best_fit_i = -1
            lowest_score = float('inf')
            row_new = df_new.loc[j] # Get the new row once

            # Find the best-fitting old item among candidates
            for i_cand in candidates_old_indices:
                row_old = df_old.loc[i_cand]
                
                # --- MATCH SCORE (Lower is better) ---
                score = 0
                
                # Check 1: Quantity (Heavily weighted)
                if float(row_old[CAN_COL_QTY]) != float(row_new[CAN_COL_QTY]):
                    score += 5 
                    
                # Check 2: Special Procurement
                if str(row_old[CAN_COL_CS]) != str(row_new[CAN_COL_CS]):
                    score += 2
                
                # Check 3: Object Description (cleaned)
                if str(row_old[CAN_COL_DESC]) != str(row_new[CAN_COL_DESC]):
                    score += 1
                # --- END SCORE ---

                if score < lowest_score:
                    lowest_score = score
                    best_fit_i = i_cand
                
                # Optimization: A perfect match (score 0) is the best possible fit
                if lowest_score == 0:
                    break
            
            # We found the best fit
            if best_fit_i != -1:
                matched.append((best_fit_i, j))
                used_old.add(best_fit_i)
                used_new.add(j)
                # Note: We don't remove from rem_new loop because we iterate a copy
    
    # --- END: MODIFICATION (Smart Best-Fit) ---

    # All remaining items are truly added/deleted
    rem_old_final = set(rem_old) - used_old
    rem_new_final = set(rem_new) - used_new

    return matched, rem_old_final, rem_new_final

def fuzzy_match_levelwise(
    df_old: pd.DataFrame,
    df_new: pd.DataFrame,
    old_emb: np.ndarray,
    new_emb: np.ndarray,
    rem_old: set,
    rem_new: set,
    config: Dict[str, Any]
):
    w_desc = config["weights"]["description"]
    w_segs = config["weights"]["segments"]
    t_sem = config["semantic_threshold"]
    t_rep = config["replace_threshold"]
    
    # --- NEW ---
    # Get the special threshold for the root level
    t_rep_root = config.get("root_replace_threshold", 0.50)
    # --- END NEW ---

    parsed_old = {i: parse_component_number(str(df_old.at[i, CAN_COL_COMP]), config) for i in rem_old if i in df_old.index}
    parsed_new = {j: parse_component_number(str(df_new.at[j, CAN_COL_COMP]), config) for j in rem_new if j in df_new.index}

    def is_14digit_submodule_changed(parsed_old_num: Dict[str, Any], parsed_new_num: Dict[str, Any], s_desc: float = 1.0) -> bool:
        """If 14-digit format is used, sub-module mismatch or significant project+desc change must stay unmatched."""
        if parsed_old_num.get("length") != 14 or parsed_new_num.get("length") != 14:
            return False
        segs_old = parsed_old_num.get("segments", [])
        segs_new = parsed_new_num.get("segments", [])
        if len(segs_old) < 3 or len(segs_new) < 3:
            return False
        
        # 1. Check Sub-module Code (Slice [4:9])
        if segs_old[1].get("value") != segs_new[1].get("value"):
            return True
            
        # 2. Check Project Code (Slice [0:4]) - If project changes, only allow if description is very similar
        if segs_old[0].get("value") != segs_new[0].get("value") and s_desc < 0.7:
            return True
            
        return False

    old_by_lvl = defaultdict(list)
    new_by_lvl = defaultdict(list)

    for i in rem_old:
        if i in df_old.index:
            old_by_lvl[df_old.at[i, CAN_COL_LEVEL]].append(i)
    for j in rem_new:
        if j in df_new.index:
            new_by_lvl[df_new.at[j, CAN_COL_LEVEL]].append(j)

    lvls = sorted(set(old_by_lvl.keys()) | set(new_by_lvl.keys()))
    
    # --- NEW ---
    # Find the root level (lowest level number)
    root_level = min(lvls) if lvls else -1
    # --- END NEW ---

    matched = []
    used_old, used_new = set(), set()
    
    # Create mapping from SourceIndex -> EmbeddingIndex (0...N)
    old_idx_map = {idx: i for i, idx in enumerate(df_old.index)}
    new_idx_map = {idx: i for i, idx in enumerate(df_new.index)}

    for lvl in lvls:
        o_idxs_source = old_by_lvl.get(lvl, [])
        n_idxs_source = new_by_lvl.get(lvl, [])

        if not o_idxs_source or not n_idxs_source:
            continue
        
        # Map source indexes to embedding indexes
        old_emb_indices = [old_idx_map[i] for i in o_idxs_source if i in old_idx_map]
        new_emb_indices = [new_idx_map[j] for j in n_idxs_source if j in new_idx_map]
        
        if not old_emb_indices or not new_emb_indices:
            continue

        D, I = nn_search(old_emb, new_emb, old_emb_indices, new_emb_indices)

        if len(D) == 0:
            continue

        cand = []
        for k, sim in enumerate(D):
            # I[k] is the *local* index into old_emb_indices
            # k is the *local* index into new_emb_indices
            
            i_old_source = o_idxs_source[I[k]] # Get SourceIndex from original list
            j_new_source = n_idxs_source[k]    # Get SourceIndex from original list

            if i_old_source in used_old or j_new_source in used_new:
                continue

            s_desc = float(sim)
            # Restore filter but allow 100% exact string matches to bypass it (e.g. for DG 33 parity)
            is_exact = (str(df_old.at[i_old_source, CAN_COL_DESC]).strip().lower() == 
                        str(df_new.at[j_new_source, CAN_COL_DESC]).strip().lower() and 
                        str(df_old.at[i_old_source, CAN_COL_DESC]).strip() != "")
            
            if not is_exact:
                if is_14digit_submodule_changed(parsed_old[i_old_source], parsed_new[j_new_source], s_desc):
                    continue
            
            s_segs = segment_match_score(parsed_old[i_old_source], parsed_new[j_new_source])
            total = w_desc * s_desc + w_segs * s_segs
            
            # --- MODIFICATION ---
            # Determine the correct threshold for this level
            current_threshold = t_rep_root if lvl == root_level else t_rep

            # Check against the level-specific threshold OR the high-similarity semantic threshold
            if is_exact or (total >= current_threshold or s_desc >= t_sem):
                # GENERIC RULE: If one description is a dash-suffix version of the other (e.g. -APM),
                # we treat them as functionally different and do not merge them.
                desc_o = str(df_old.at[i_old_source, CAN_COL_DESC]).strip().upper()
                desc_n = str(df_new.at[j_new_source, CAN_COL_DESC]).strip().upper()
                if not is_exact and (desc_o.startswith(desc_n + "-") or desc_n.startswith(desc_o + "-")):
                    continue
                    
                cand.append((total, s_desc, s_segs, i_old_source, j_new_source))
            # --- END MODIFICATION ---

        cand.sort(reverse=True, key=lambda x: x[0])

        for total, s_desc, s_segs, i_old_source, j_new_source in cand:
            if i_old_source in used_old or j_new_source in used_new:
                continue

            matched.append((i_old_source, j_new_source, total, s_desc, s_segs))
            used_old.add(i_old_source)
            used_new.add(j_new_source)

    rem_old2 = set(rem_old) - used_old
    rem_new2 = set(rem_new) - used_new

    return matched, rem_old2, rem_new2

def classify_pairs(
    df_old: pd.DataFrame,
    df_new: pd.DataFrame,
    df_old_raw: pd.DataFrame,
    df_new_raw: pd.DataFrame,
    pairs_exact: List[Tuple[int, int]],
    pairs_fuzzy: List[Tuple[int, int, float, float, float]],
    old_emb: np.ndarray,
    new_emb: np.ndarray,
    config: Dict[str, Any],
    dynamic_pairs: List[Dict[str, str]] 
):
    t_sem = config["semantic_threshold"]

    out_unchanged = []
    out_replaced = []
    out_modified = []

    def norm(s): return _norm_comp(s)
    
    # Create mapping from SourceIndex -> EmbeddingIndex
    old_idx_map = {idx: i for i, idx in enumerate(df_old.index)}
    new_idx_map = {idx: i for i, idx in enumerate(df_new.index)}
    
    # Helper to get raw value safely
    def get_raw_val(df_raw, idx, col):
        if idx in df_raw.index and col in df_raw.columns:
            return df_raw.at[idx, col]
        return None

    for i, j in pairs_exact:
        # i and j are SourceIndexes
        oc = norm(df_old.at[i, CAN_COL_COMP])
        nc = norm(df_new.at[j, CAN_COL_COMP])
        same_comp = oc == nc

        parts = []
        
        # Get embedding indices
        emb_i = old_idx_map.get(i)
        emb_j = new_idx_map.get(j)

        # Default column checks
        if emb_i is not None and emb_j is not None:
             if cos_sim_vec(old_emb[emb_i], new_emb[emb_j]) < t_sem:
                parts.append(f"Description: '{df_old.at[i, CAN_COL_DESC]}' -> '{df_new.at[j, CAN_COL_DESC]}'")
        
        if float(df_old.at[i, CAN_COL_QTY]) != float(df_new.at[j, CAN_COL_QTY]):
            parts.append(f"Quantity: {df_old.at[i, CAN_COL_QTY]} â†’ {df_new.at[j, CAN_COL_QTY]}")
        
        cs_old = str(df_old.at[i, CAN_COL_CS])
        cs_new = str(df_new.at[j, CAN_COL_CS])
        if cs_old != cs_new:
            # Report the raw values if possible, else cleaned
            raw_cs_old = df_old_raw.at[i, "Special procurement"] if "Special procurement" in df_old_raw.columns else cs_old
            raw_cs_new = df_new_raw.at[j, "Special procurement"] if "Special procurement" in df_new_raw.columns else cs_new
            parts.append(f"Special procurement: '{cs_old}' â†’ '{cs_new}'")
        
        # --- DYNAMIC MAPPED COMPARISON ---
        for pair in dynamic_pairs:
            old_col = pair.get('old')
            new_col = pair.get('new')
            # Use Old Col name as label or a combined one
            label = old_col if old_col == new_col else f"{old_col}/{new_col}"
            
            val_o = get_raw_val(df_old_raw, i, old_col)
            val_n = get_raw_val(df_new_raw, j, new_col)
            
            if str(val_o) != str(val_n):
                parts.append(f"{label}: {val_o} â†’ {val_n}")
        
        is_modified = bool(parts)

        if same_comp and not is_modified:
            out_unchanged.append(("exact_component_number", i, j, ""))
        elif same_comp and is_modified:
            out_modified.append(("same_component_number", i, j, "; ".join(parts)))
        elif not same_comp:
            parts = []
            po = parse_component_number(oc, config)
            pn = parse_component_number(nc, config)
            barrel_changes = describe_barrel_changes(po, pn)
            if barrel_changes:
                parts.append("Component number barrels: " + "; ".join(barrel_changes))
            else:
                parts.append(f"Component number: {oc} â†’ {nc}")
            out_replaced.append(("exact_match_different_comp", i, j, "; ".join(parts)))

    for i, j, total, s_desc, s_segs in pairs_fuzzy:
        # i and j are SourceIndexes
        oc = norm(df_old.at[i, CAN_COL_COMP])
        nc = norm(df_new.at[j, CAN_COL_COMP])

        parts = []
        if oc != nc:
            po = parse_component_number(oc, config)
            pn = parse_component_number(nc, config)
            barrel_changes = describe_barrel_changes(po, pn)
            if barrel_changes:
                parts.append("Component number barrels: " + "; ".join(barrel_changes))
            else:
                parts.append(f"Component number: {oc} â†’ {nc}")

        if df_old.at[i, CAN_COL_DESC] != df_new.at[j, CAN_COL_DESC]:
            parts.append(f"Description: '{df_old.at[i, CAN_COL_DESC]}' â†’ '{df_new.at[j, CAN_COL_DESC]}'")

        if float(df_old.at[i, CAN_COL_QTY]) != float(df_new.at[j, CAN_COL_QTY]):
            parts.append(f"Quantity: {df_old.at[i, CAN_COL_QTY]} â†’ {df_new.at[j, CAN_COL_QTY]}")

        cs_old = str(df_old.at[i, CAN_COL_CS])
        cs_new = str(df_new.at[j, CAN_COL_CS])
        if cs_old != cs_new:
            parts.append(f"Special procurement: '{cs_old}' â†’ '{cs_new}'")
            
        # --- DYNAMIC MAPPED COMPARISON (Fuzzy) ---
        for pair in dynamic_pairs:
            old_col = pair.get('old')
            new_col = pair.get('new')
            label = old_col if old_col == new_col else f"{old_col}/{new_col}"
            
            val_o = get_raw_val(df_old_raw, i, old_col)
            val_n = get_raw_val(df_new_raw, j, new_col)
            
            if str(val_o) != str(val_n):
                parts.append(f"{label}: {val_o} â†’ {val_n}")

        detail = "; ".join(parts) if parts else "Similar part (semantic/segment match)"
        out_replaced.append(("fuzzy_segments+semantic", i, j, detail))

    return out_unchanged, out_modified, out_replaced

# ---------------------- NEW: HELPER FUNCTIONS FOR HIERARCHICAL STATUS ----------------------

def _flatten_nodes(node: BOMNode) -> List[BOMNode]:
    """Helper function to flatten a tree into a list for node lookup."""
    nodes = [node]
    for child in node.children:
        nodes.extend(_flatten_nodes(child))
    return nodes

def update_node_statuses(roots_old: List[BOMNode], roots_new: List[BOMNode],
                        unchanged, modified, replaced, rem_old_final, rem_new_final):
    """Update primary_status of BOMNode objects based on classification results"""

    # Create lookup dictionaries
    node_lookup_old = {node.source_idx: node for root in roots_old for node in _flatten_nodes(root)}
    node_lookup_new = {node.source_idx: node for root in roots_new for node in _flatten_nodes(root)}

    # Update unchanged nodes
    for _, i, j, _ in unchanged:
        if i in node_lookup_old:
            node_lookup_old[i].primary_status = "Unchanged"
        if j in node_lookup_new:
            node_lookup_new[j].primary_status = "Unchanged"

    # Update modified nodes
    for _, i, j, _ in modified:
        if i in node_lookup_old:
            node_lookup_old[i].primary_status = "Modified"
        if j in node_lookup_new:
            node_lookup_new[j].primary_status = "Modified"

    # Update replaced nodes
    for _, i, j, _ in replaced:
        if i in node_lookup_old:
            node_lookup_old[i].primary_status = "Replaced"
        if j in node_lookup_new:
            node_lookup_new[j].primary_status = "Replaced"

    # Update deleted nodes
    for i in rem_old_final:
        if i in node_lookup_old:
            node_lookup_old[i].primary_status = "Deleted"

    # Update added nodes
    for j in rem_new_final:
        if j in node_lookup_new:
            node_lookup_new[j].primary_status = "Newly Added"

def propagate_hierarchical_status(roots: List[BOMNode]):
    """
    Propagate hierarchical status up the tree from any changing child.
    Uses post-order traversal to ensure all children are processed before parents.
    """
    def _walk(node: BOMNode):
        # Process children first (post-order traversal)
        for child in node.children:
            _walk(child)

        # Check if any children have non-unchanged status
        has_modified_children = any(
            child.primary_status != "Unchanged" or child.hierarchical_status != "Unchanged"
            for child in node.children
        )

        if has_modified_children and node.hierarchical_status == "Unchanged":
            node.hierarchical_status = "Modified"

            # Add details about child changes
            details = Counter([
                child.primary_status for child in node.children
                if child.primary_status != "Unchanged"
            ])
            details_hierarchical = Counter([
                child.hierarchical_status for child in node.children
                if child.hierarchical_status != "Unchanged"
            ])

            combined_details = details + details_hierarchical
            if combined_details:
                node.change_details = f"Child changes: {', '.join([f'{k}: {v}' for k, v in combined_details.items()])}"

    for root in roots:
        _walk(root)

# ---------------------- EXPORT ----------------------

def write_sheet(writer, df: pd.DataFrame, name: str, header_color: str, row_color: Optional[str] = None, grey_mask=None):
    if df is None:
        df = pd.DataFrame()

    df_to_write = df.fillna("")
    df_to_write.to_excel(writer, sheet_name=name, index=False)

    ws = writer.sheets[name]
    book = writer.book
    
    header_fmt = book.add_format({"bold": True, "bg_color": header_color, "border": 1})
    
    # --- NEW: Text Format for Component Columns ---
    text_fmt = book.add_format({'num_format': '@'})
    
    def _is_comp_col(c_name):
        cn = str(c_name).lower()
        # Heuristic to identify ID columns that should be Text
        return any(x in cn for x in ["component", "material", "part number", "item id"])

    for col_num, value in enumerate(df_to_write.columns.values):
        ws.write(0, col_num, value, header_fmt)
        width = max(12, min(50, len(str(value)) + 2))
        
        if _is_comp_col(value):
            ws.set_column(col_num, col_num, width, text_fmt)
        else:
            ws.set_column(col_num, col_num, width)
    # ----------------------------------------------

    if row_color and len(df_to_write) > 0:
        fmt = book.add_format({"bg_color": row_color})
        ws.conditional_format(1, 0, len(df_to_write), len(df_to_write.columns) - 1, {"type": "no_blanks", "format": fmt})

    if grey_mask is not None:
        grey_fmt = book.add_format({"bg_color": "#DDDDDD"})
        for idx, flag in enumerate(grey_mask, start=1):
            if flag:
                ws.set_row(idx, None, grey_fmt)

def _col_idx_to_a1(idx: int) -> str:
    letters = ""
    x = idx
    while True:
        letters = chr(65 + (x % 26)) + letters
        x = x // 26 - 1
        if x < 0:
            break
    return letters

def process_df_for_excel(df: pd.DataFrame, sheet_name: str) -> pd.DataFrame:
    # CHANGE: Cleaned up drop lists for Added/Deleted sheets and DynamicCompare
    drops = {
        "Newly Added": ["Match_Type", "Change Type", "Change Detail", "Old_SourceIndex", "New_SourceIndex", "Old Seg1 Name", "Old Seg1 Val", "Old Seg1 Label", "New Seg1 Name", "New Seg1 Val", "New Seg1 Label", "Seg1 Changed", "Barrel_Change_Summary"],
        "Deleted": ["Match_Type", "Change Type", "Change Detail", "Old_SourceIndex", "New_SourceIndex", "Old Seg1 Name", "Old Seg1 Val", "Old Seg1 Label", "New Seg1 Name", "New Seg1 Val", "New Seg1 Label", "Seg1 Changed", "Barrel_Change_Summary"],
        "Replaced": ["Old_SourceIndex", "New_SourceIndex", "Old Seg1 Name", "Old Seg1 Val", "Old Seg1 Label", "New Seg1 Name", "New Seg1 Val", "New Seg1 Label", "Seg1 Changed", "Barrel_Change_Summary"],
        "Modified": ["Old_SourceIndex", "New_SourceIndex", "Old Seg1 Name", "Old Seg1 Val", "Old Seg1 Label", "New Seg1 Name", "New Seg1 Val", "New Seg1 Label", "Seg1 Changed", "Barrel_Change_Summary"],
        "Carry-Over": ["Match_Type", "Change Type", "Change Detail", "Old_SourceIndex", "New_SourceIndex", "Old Seg1 Name", "Old Seg1 Val", "Old Seg1 Label", "New Seg1 Name", "New Seg1 Val", "New Seg1 Label", "Seg1 Changed", "Barrel_Change_Summary"],
        "Skipped": ["Source"],
        "All": ["Old_SourceIndex", "New_SourceIndex", "Old Seg1 Name", "Old Seg1 Val", "Old Seg1 Label", "New Seg1 Name", "New Seg1 Val", "New Seg1 Label", "Seg1 Changed", "Barrel_Change_Summary", "Change Type", "Match_Type"],
    }
    
    to_drop = drops.get(sheet_name, [])
    final_cols_to_drop = [c for c in to_drop if c in df.columns]
    
    if sheet_name == "Newly Added":
        final_cols_to_drop.extend([c for c in df.columns if str(c).startswith("Old ")])
    elif sheet_name == "Deleted":
        final_cols_to_drop.extend([c for c in df.columns if str(c).startswith("New ")])
        
    return df.drop(columns=final_cols_to_drop) if final_cols_to_drop else df

def _normalize_source_index(value: Any) -> Optional[int]:
    try:
        if value is None or (isinstance(value, str) and value.strip() == ""):
            return None
        if pd.isna(value):
            return None
        return int(float(value))
    except Exception:
        return None

def _with_source_index_column(df: Optional[pd.DataFrame]) -> pd.DataFrame:
    if df is None:
        return pd.DataFrame()
    out = df.copy()
    if "SourceIndex" not in out.columns:
        out.insert(0, "SourceIndex", out.index)
    return out

def build_new_clm_bom_sheet(
    df_new_raw: pd.DataFrame,
    df_all: pd.DataFrame,
    df_skipped: pd.DataFrame
) -> pd.DataFrame:
    if df_new_raw is None:
        return pd.DataFrame()

    clm_df = _with_source_index_column(df_new_raw)
    status_map: Dict[int, str] = {}
    detail_map: Dict[int, str] = {}
    old_comp_map: Dict[int, str] = {}

    if df_all is not None and not df_all.empty:
        for _, row in df_all.iterrows():
            src_idx = _normalize_source_index(row.get("New_SourceIndex"))
            if src_idx is None:
                continue

            status = str(row.get("Primary Status", "")).strip() or "Not Analyzed"
            detail = str(row.get("Change Detail", "")).strip()
            old_comp = str(row.get(f"Old {CAN_COL_COMP}", "")).strip()

            if src_idx not in status_map:
                status_map[src_idx] = status
            if detail and src_idx not in detail_map:
                detail_map[src_idx] = detail
            if old_comp and old_comp.lower() != "nan" and src_idx not in old_comp_map:
                old_comp_map[src_idx] = old_comp

    if df_skipped is not None and not df_skipped.empty:
        for _, row in df_skipped.iterrows():
            if str(row.get("Source", "")).strip().lower() != "new":
                continue
            src_idx = _normalize_source_index(row.get("SourceIndex"))
            if src_idx is None:
                continue
            status_map[src_idx] = "Skipped"
            if src_idx not in detail_map:
                detail_map[src_idx] = "Skipped by Special Procurement rule"

    clm_df["Analysis Status"] = clm_df["SourceIndex"].map(
        lambda x: status_map.get(_normalize_source_index(x), "Not Analyzed")
    )
    clm_df["Analysis Detail"] = clm_df["SourceIndex"].map(
        lambda x: detail_map.get(_normalize_source_index(x), "")
    )
    clm_df["Matched Old Component"] = clm_df["SourceIndex"].map(
        lambda x: old_comp_map.get(_normalize_source_index(x), "")
    )
    return clm_df

def build_summery_sheet_df(
    df_unchanged: Optional[pd.DataFrame],
    df_modified: Optional[pd.DataFrame],
    df_replaced: Optional[pd.DataFrame],
    df_added: Optional[pd.DataFrame],
    df_deleted: Optional[pd.DataFrame],
    df_skipped: Optional[pd.DataFrame],
    df_all: Optional[pd.DataFrame],
    old_bom_sheet: Optional[pd.DataFrame],
    new_bom_sheet: Optional[pd.DataFrame],
    new_clm_bom_sheet: Optional[pd.DataFrame],
) -> pd.DataFrame:
    def _count(df: Optional[pd.DataFrame]) -> int:
        return 0 if df is None else int(len(df))

    rows = [
        {"Metric": "Total Compared", "Value": _count(df_all)},
        {"Metric": "Carry-Over", "Value": _count(df_unchanged)},
        {"Metric": "Modified", "Value": _count(df_modified)},
        {"Metric": "Replaced", "Value": _count(df_replaced)},
        {"Metric": "Newly Added", "Value": _count(df_added)},
        {"Metric": "Deleted", "Value": _count(df_deleted)},
        {"Metric": "Skipped", "Value": _count(df_skipped)},
        {"Metric": "old_bom Rows", "Value": _count(old_bom_sheet)},
        {"Metric": "new_bom Rows", "Value": _count(new_bom_sheet)},
        {"Metric": "new_clm_bom Rows", "Value": _count(new_clm_bom_sheet)},
    ]

    return pd.DataFrame(rows)

def export_excel(out_path,
                df_unchanged, df_modified, df_replaced, df_added, df_deleted, df_all, df_skipped,
                skip_flag: bool, df_old_raw: Optional[pd.DataFrame] = None, df_new_raw: Optional[pd.DataFrame] = None):
    # CHANGE: Removed dynamic_single_df parameter
    out_dir = os.path.dirname(os.path.abspath(out_path))
    os.makedirs(out_dir, exist_ok=True)
    with pd.ExcelWriter(out_path, engine="xlsxwriter") as writer:
        # --- START REORGANIZED EXCEL EXPORT ---
        # 1. Summary & Submodule Summary are added by generate_json_hierarchy.py downstream
        
        # 3. Master (Complete Comparison)
        df_master = process_df_for_excel(df_all, "All")
        write_sheet(writer, df_master, "Master", "#BDD7EE")
        
        # 4. Carryover
        write_sheet(writer, process_df_for_excel(df_unchanged.fillna("") if df_unchanged is not None else pd.DataFrame(), "Carry-Over"), "Carryover", "#D9E1F2", None,
                   df_unchanged["SkippedChildren"] if skip_flag and df_unchanged is not None and "SkippedChildren" in df_unchanged.columns else None)

        # 5. Exclusive (A+B+C)
        df_exclusive = df_all[df_all["Primary Status"].isin(["Modified", "Newly Added", "Replaced"])].copy()
        write_sheet(writer, process_df_for_excel(df_exclusive, "All"), "Exclusive", "#FCD5B4")

        # 6. Modified
        write_sheet(writer, process_df_for_excel(df_modified.fillna("") if df_modified is not None else pd.DataFrame(), "Modified"), "Modified", "#A9D08E", "#D8E4BC",
                   df_modified["SkippedChildren"] if skip_flag and df_modified is not None and "SkippedChildren" in df_modified.columns else None)

        # 7. Replaced
        write_sheet(writer, process_df_for_excel(df_replaced.fillna("") if df_replaced is not None else pd.DataFrame(), "Replaced"), "Replaced", "#FFEB9C", "#FFF2CC",
                   df_replaced["SkippedChildren"] if skip_flag and df_replaced is not None and "SkippedChildren" in df_replaced.columns else None)

        # 8. Newly Added
        write_sheet(writer, process_df_for_excel(df_added.fillna("") if df_added is not None else pd.DataFrame(), "Newly Added"), "Newly Added", "#C6EFCE", "#E2F0D9",
                   df_added["SkippedChildren"] if skip_flag and df_added is not None and "SkippedChildren" in df_added.columns else None)

        # 9. Deleted
        write_sheet(writer, process_df_for_excel(df_deleted.fillna("") if df_deleted is not None else pd.DataFrame(), "Deleted"), "Deleted", "#FFC7CE", "#FCE4D6",
                   df_deleted["SkippedChildren"] if skip_flag and df_deleted is not None and "SkippedChildren" in df_deleted.columns else None)

        # 10. New VC BOM
        new_bom_sheet = _with_source_index_column(df_new_raw)
        write_sheet(writer, new_bom_sheet.fillna(""), "New VC BOM", "#9DC3E6")

        # 11. Base VC BOM
        old_bom_sheet = _with_source_index_column(df_old_raw)
        write_sheet(writer, old_bom_sheet.fillna(""), "Base VC BOM", "#F4B183")

        # Extra: Comparison Logical Master (New CLM BOM)
        new_clm_bom_sheet = build_new_clm_bom_sheet(df_new_raw, df_all, df_skipped)
        write_sheet(writer, new_clm_bom_sheet.fillna(""), "new_clm_bom", "#C6E0B4")

        # Extra: Skipped
        write_sheet(writer, process_df_for_excel(df_skipped.fillna("") if df_skipped is not None else pd.DataFrame(), "Skipped"), "Skipped", "#DDDDDD", "#F2F2F2")

        # Format the Master sheet with conditional coloring
        ws_master = writer.sheets["Master"]
        book = writer.book
        last_row = len(df_master) + 1
        try:
            status_col_index = list(df_master.columns).index("Primary Status")
            status_col_A1 = _col_idx_to_a1(status_col_index)
            for label, color in [("Newly Added", "#E2F0D9"), ("Deleted", "#FCE4D6"), ("Replaced", "#FFF2CC"), ("Modified", "#D8E4BC"), ("Carry-Over", "#FFFFFF")]:
                ws_master.conditional_format(1, 0, last_row, len(df_master.columns) - 1, {
                    "type": "formula",
                    "criteria": f'=${status_col_A1}2="{label}"',
                    "format": book.add_format({"bg_color": color})
                })
        except ValueError:
            pass


# ---------------------- Helpers for building final sheet rows ----------------------

def _split_change_string_to_type_detail(s: str) -> Tuple[str, str]:
    if not s:
        return "", ""

    parts = [p.strip() for p in s.split(";") if p.strip()]
    types, details = [], []

    for p in parts:
        if ":" in p:
            left, right = p.split(":", 1)
            types.append(left.strip())
            details.append(right.strip())
        else:
            types.append("General")
            details.append(p)

    return ", ".join(types), "; ".join(details)

# --- NEW: Parent Path Comparison Helper ---
def _norm_path(p: Any) -> str:
    """Safely convert value to string; treat NaN/None as empty string and strip slashes."""
    if p is None or (isinstance(p, float) and math.isnan(p)):
        return ""
    return str(p).strip().strip('/')

def compare_parent_paths(old_path: Any, new_path: Any, root_level: int) -> str:
    """
    Compares two parent paths, respecting the root_level (0 for PLM, 1 for SAP).
    Returns 'Yes', 'No', or 'N/A'.
    """
    p_old = _norm_path(old_path)
    p_new = _norm_path(new_path)

    if p_old == "" and p_new == "":
        return "No" # Both are root nodes (or have no path), so path hasn't changed.
    
    if p_old == p_new:
        return "No" # Paths are identical

    # If root_level is 1 (SAP), skip the first segment
    if root_level == 1:
        p_old_parts = p_old.split('/')
        p_new_parts = p_new.split('/')
        
        p_old_compare = "/".join(p_old_parts[1:]) if len(p_old_parts) > 1 else ""
        p_new_compare = "/".join(p_new_parts[1:]) if len(p_new_parts) > 1 else ""
        
        return "Yes" if p_old_compare != p_new_compare else "No"
    
    # If root_level is 0 (PLM) or any other value, compare the full path
    # We already checked for equality, so if we are here, they are different.
    return "Yes"
# --- END NEW ---

def row_old(flat_old: pd.DataFrame, i: int) -> Dict[str, Any]:
    if i not in flat_old.index: return {"Old_SourceIndex": i}
    r = flat_old.loc[i]
    out = {}
    # Manually order columns to insert Length right after Component
    out[f"Old {CAN_COL_LEVEL}"] = r[CAN_COL_LEVEL]
    
    comp_val = r[CAN_COL_COMP]
    out[f"Old {CAN_COL_COMP}"] = comp_val
    # --- NEW: Component Length ---
    out["Old Component Length"] = len(str(comp_val)) if comp_val else 0
    # -----------------------------
    
    for k in [CAN_COL_DESC, CAN_COL_CS, CAN_COL_QTY, "ParentPath", "SkippedChildren"]:
        out[f"Old {k}"] = r[k]
        
    out["Old_SourceIndex"] = r.get("SourceIndex", None)
    return out

def row_new(flat_new: pd.DataFrame, j: int) -> Dict[str, Any]:
    if j not in flat_new.index: return {"New_SourceIndex": j}
    r = flat_new.loc[j]
    out = {}
    # Manually order columns to insert Length right after Component
    out[f"New {CAN_COL_LEVEL}"] = r[CAN_COL_LEVEL]
    
    comp_val = r[CAN_COL_COMP]
    out[f"New {CAN_COL_COMP}"] = comp_val
    # --- NEW: Component Length ---
    out["New Component Length"] = len(str(comp_val)) if comp_val else 0
    # -----------------------------

    for k in [CAN_COL_DESC, CAN_COL_CS, CAN_COL_QTY, "ParentPath", "SkippedChildren"]:
        out[f"New {k}"] = r[k]

    out["New_SourceIndex"] = r.get("SourceIndex", None)
    return out

def _add_structured_segments(base: Dict[str, Any], old_comp: str, new_comp: str, cfg: Dict[str, Any]) -> None:
    po = parse_component_number(old_comp, cfg)
    pn = parse_component_number(new_comp, cfg)

    segs_o, segs_n = po["segments"], pn["segments"]
    maxn = max(len(segs_o), len(segs_n))

    for i in range(maxn):
        so = segs_o[i] if i < len(segs_o) else {"name": f"seg{i+1}", "value": "", "label": None}
        sn = segs_n[i] if i < len(segs_n) else {"name": f"seg{i+1}", "value": "", "label": None}

        base[f"Old Seg{i+1} Name"] = so.get("name")
        base[f"Old Seg{i+1} Val"] = so.get("value")
        base[f"Old Seg{i+1} Label"] = so.get("label")
        base[f"New Seg{i+1} Name"] = sn.get("name")
        base[f"New Seg{i+1} Val"] = sn.get("value")
        base[f"New Seg{i+1} Label"] = sn.get("label")
        base[f"Seg{i+1} Changed"] = (so.get("value") != sn.get("value")) or (so.get("label") != sn.get("label"))

    summary = "; ".join(describe_barrel_changes(po, pn))
    base["Barrel_Change_Summary"] = summary

def build_output_rows_from_matches(flat_old: pd.DataFrame, flat_new: pd.DataFrame,
                                 unchanged, modified, replaced,
                                 rem_old_final, rem_new_final,
                                 cfg: Dict[str, Any], roots_old: List[BOMNode], roots_new: List[BOMNode],
                                 root_level: int): # <-- NEW: Added root_level

    # Create lookup dictionaries for hierarchical status
    node_lookup_old = {node.source_idx: node for root in roots_old for node in _flatten_nodes(root)}
    node_lookup_new = {node.source_idx: node for root in roots_new for node in _flatten_nodes(root)}
    
    # --- REMOVED FLAGGED GLOBAL LOOKUP (Source of False Positives) ---
    # old_status_map = {} ... DELETED

    def get_hierarchical_status(source_idx: Optional[int], lookup: Dict[int, BOMNode]) -> str:
        if source_idx is not None and source_idx in lookup:
            return lookup[source_idx].hierarchical_status
        return "Unchanged"

    unchanged_rows = []
    for match_type, i, j, _ in unchanged:
        base = {
            "Primary Status": "Carry-Over",
            "Hierarchical Status": get_hierarchical_status(j, node_lookup_new),
            "Match_Type": match_type,
            "Change Type": "",
            "Change Detail": ""
        }
        base.update(row_old(flat_old, i))
        base.update(row_new(flat_new, j))

        # --- START: MODIFICATION (Strict Parent Node Check) ---
        old_p_str = str(base.get("Old ParentPath")).strip()
        new_p_str = str(base.get("New ParentPath")).strip()
        
        path_changed_flag = "No" # Default
        
        # 1. Check if string changed
        if compare_parent_paths(old_p_str, new_p_str, root_level) == "Yes":
            
            # 2. Get the ACTUAL Parent Nodes (No string parsing)
            node_old = node_lookup_old.get(i)
            node_new = node_lookup_new.get(j)
            
            # Safe access to parent
            parent_old = node_old.parent if node_old else None
            parent_new = node_new.parent if node_new else None
            
            # If either is None (Root node), treat as No change (Roots don't have parents)
            if parent_old and parent_new:
                
                # A. IDENTITY CHECK: Are they the same component ID?
                if parent_old.component_number == parent_new.component_number:
                    path_changed_flag = "No" # Same immediate parent, ignore upstream changes
                else:
                    # B. STATUS CHECK: Is New Parent a structure change?
                    if parent_new.primary_status in ("Replaced", "Newly Added"):
                        # Suppress by default (Consequential)
                        path_changed_flag = "No"
                        
                        # C. BYPASS: Did we leave a valid Old Parent?
                        # Check specific status of the OLD parent node instance
                        if parent_old.primary_status in ("Unchanged", "Modified", "Carry-Over"):
                            path_changed_flag = "Yes" # Active Move away from valid parent
                    else:
                        # New Parent is Unchanged/Modified -> We moved into it.
                        path_changed_flag = "Yes"
                        
        base["Parent Path Changed"] = path_changed_flag
        # --- END: MODIFICATION ---
        
        if OUTPUT_STRUCTURED_SEGMENTS:
            _add_structured_segments(base, _norm_comp(base.get(f"Old {CAN_COL_COMP}", "")),
                                   _norm_comp(base.get(f"New {CAN_COL_COMP}", "")), cfg)
        unchanged_rows.append(base)

    modified_rows = []
    for match_type, i, j, change in modified:
        change_type, change_detail = _split_change_string_to_type_detail(change) if isinstance(change, str) else ("", "")
        base = {
            "Primary Status": "Modified",
            "Hierarchical Status": get_hierarchical_status(j, node_lookup_new),
            "Match_Type": match_type,
            "Change Type": change_type,
            "Change Detail": change_detail
        }
        base.update(row_old(flat_old, i))
        base.update(row_new(flat_new, j))


        # --- START: MODIFICATION (Strict Parent Node Check) ---
        old_p_str = str(base.get("Old ParentPath")).strip()
        new_p_str = str(base.get("New ParentPath")).strip()
        
        path_changed_flag = "No" 
        
        if compare_parent_paths(old_p_str, new_p_str, root_level) == "Yes":
            node_old = node_lookup_old.get(i)
            node_new = node_lookup_new.get(j)
            parent_old = node_old.parent if node_old else None
            parent_new = node_new.parent if node_new else None
            
            if parent_old and parent_new:
                if parent_old.component_number == parent_new.component_number:
                    path_changed_flag = "No"
                else:
                    if parent_new.primary_status in ("Replaced", "Newly Added"):
                        path_changed_flag = "No"
                        if parent_old.primary_status in ("Unchanged", "Modified", "Carry-Over"):
                            path_changed_flag = "Yes" 
                    else:
                        path_changed_flag = "Yes"
        
        base["Parent Path Changed"] = path_changed_flag
        # --- END: MODIFICATION ---

        if OUTPUT_STRUCTURED_SEGMENTS:
            _add_structured_segments(base, _norm_comp(base.get(f"Old {CAN_COL_COMP}", "")),
                                   _norm_comp(base.get(f"New {CAN_COL_COMP}", "")), cfg)
        modified_rows.append(base)

    replaced_rows = []
    for match_type, i, j, change in replaced:
        change_type, change_detail = _split_change_string_to_type_detail(change) if isinstance(change, str) else ("", "")
        base = {
            "Primary Status": "Replaced",
            "Hierarchical Status": get_hierarchical_status(j, node_lookup_new),
            "Match_Type": match_type,
            "Change Type": change_type,
            "Change Detail": change_detail
        }
        base.update(row_old(flat_old, i))
        base.update(row_new(flat_new, j))

        # --- START: MODIFICATION (Strict Parent Node Check) ---
        old_p_str = str(base.get("Old ParentPath")).strip()
        new_p_str = str(base.get("New ParentPath")).strip()
        
        path_changed_flag = "No" 
        
        if compare_parent_paths(old_p_str, new_p_str, root_level) == "Yes":
            node_old = node_lookup_old.get(i)
            node_new = node_lookup_new.get(j)
            parent_old = node_old.parent if node_old else None
            parent_new = node_new.parent if node_new else None
            
            if parent_old and parent_new:
                if parent_old.component_number == parent_new.component_number:
                    path_changed_flag = "No"
                else:
                    if parent_new.primary_status in ("Replaced", "Newly Added"):
                        path_changed_flag = "No"
                        if parent_old.primary_status in ("Unchanged", "Modified", "Carry-Over"):
                            path_changed_flag = "Yes" 
                    else:
                        path_changed_flag = "Yes"
        
        base["Parent Path Changed"] = path_changed_flag
        # --- END: MODIFICATION ---

        if OUTPUT_STRUCTURED_SEGMENTS:
            _add_structured_segments(base, _norm_comp(base.get(f"Old {CAN_COL_COMP}", "")),
                                   _norm_comp(base.get(f"New {CAN_COL_COMP}", "")), cfg)
        replaced_rows.append(base)

    added_rows = []
    for j in rem_new_final:
        base = {
            "Primary Status": "Newly Added",
            "Hierarchical Status": get_hierarchical_status(j, node_lookup_new),
            "Match_Type": "",
            "Change Type": "",
            "Change Detail": "",
            "Parent Path Changed": "N/A" # --- NEW ---
        }
        # CHANGE: Removed adding empty "Old" columns
        base.update(row_new(flat_new, j))
        if OUTPUT_STRUCTURED_SEGMENTS:
            _add_structured_segments(base, "", _norm_comp(base.get(f"New {CAN_COL_COMP}", "")), cfg)
        added_rows.append(base)

    deleted_rows = []
    for i in rem_old_final:
        base = {
            "Primary Status": "Deleted",
            "Hierarchical Status": get_hierarchical_status(i, node_lookup_old),
            "Match_Type": "",
            "Change Type": "",
            "Change Detail": "",
            "Parent Path Changed": "N/A" # --- NEW ---
        }
        # CHANGE: Removed adding empty "New" columns
        base.update(row_old(flat_old, i))
        if OUTPUT_STRUCTURED_SEGMENTS:
            _add_structured_segments(base, _norm_comp(base.get(f"Old {CAN_COL_COMP}", "")), "", cfg)
        deleted_rows.append(base)

    # --- NEW: Define columns for empty DFs ---
    empty_df_cols = ["Primary Status", "Hierarchical Status", "Match_Type", "Change Type", "Change Detail", "Parent Path Changed"]

    df_unchanged = pd.DataFrame(unchanged_rows) if unchanged_rows else pd.DataFrame(columns=empty_df_cols)
    df_modified = pd.DataFrame(modified_rows) if modified_rows else pd.DataFrame(columns=empty_df_cols)
    df_replaced = pd.DataFrame(replaced_rows) if replaced_rows else pd.DataFrame(columns=empty_df_cols)
    df_added = pd.DataFrame(added_rows) if added_rows else pd.DataFrame()
    df_deleted = pd.DataFrame(deleted_rows) if deleted_rows else pd.DataFrame()

    df_all = pd.concat([df_unchanged, df_modified, df_replaced, df_added, df_deleted], ignore_index=True, sort=False).fillna("")

    return df_unchanged, df_modified, df_replaced, df_added, df_deleted, df_all

def generate_summary(df_unchanged: pd.DataFrame, df_modified: pd.DataFrame, df_replaced: pd.DataFrame, df_added: pd.DataFrame, df_deleted: pd.DataFrame,
                    df_skipped: pd.DataFrame, flat_old: pd.DataFrame, flat_new: pd.DataFrame,
                    exact_pairs: List[Tuple[int, int]], fuzzy_pairs: List[Tuple[int, int, float, float, float]],
                    config: Dict[str, Any], start_ts: float, end_ts: float, root_level: int) -> Dict[str, Any]:

    old_count = int(len(flat_old))
    new_count = int(len(flat_new))
    total_compared = old_count + new_count

    added_count = int(len(df_added)) if df_added is not None else 0
    deleted_count = int(len(df_deleted)) if df_deleted is not None else 0
    replaced_count = int(len(df_replaced)) if df_replaced is not None else 0
    modified_count = int(len(df_modified)) if df_modified is not None else 0
    unchanged_count = int(len(df_unchanged)) if df_unchanged is not None else 0
    skipped_count = int(len(df_skipped)) if df_skipped is not None else 0

    deviation_pct = round(((added_count + deleted_count) / old_count * 100) if old_count else 0.0, 2)
    carry_over_pct = round((unchanged_count / old_count * 100) if old_count else 0.0, 2)

    qty_changes = []
    qty_delta_total = 0.0
    qty_changes_list = []

    if df_modified is not None and not df_modified.empty:
        for _, r in df_modified.iterrows():
            detail = r.get("Change Detail") or ""
            if isinstance(detail, str) and "Quantity" in detail:
                m = re.findall(r"Quantity[:\s]*([-\d\.,]+)\s*â†’\s*([-\d\.,]+)", detail)
                for a, b in m:
                    try:
                        a_f = float(str(a).replace(",", ""))
                        b_f = float(str(b).replace(",", ""))
                        delta = b_f - a_f
                        qty_changes.append({"old": a_f, "new": b_f, "delta": delta, "component": r.get("Old Component number") or r.get("New Component number")})
                        qty_changes_list.append(delta)
                        qty_delta_total += delta
                    except Exception:
                        continue

    if not qty_changes_list and flat_old is not None and flat_new is not None:
        old_qty_map = {str(r[CAN_COL_COMP]): float(r[CAN_COL_QTY]) for _, r in flat_old.iterrows()}
        new_qty_map = {str(r[CAN_COL_COMP]): float(r[CAN_COL_QTY]) for _, r in flat_new.iterrows()}
        for k in set(old_qty_map.keys()) & set(new_qty_map.keys()):
            a, b = old_qty_map.get(k, 0.0), new_qty_map.get(k, 0.0)
            if a != b:
                d = b - a
                qty_changes_list.append(d)
                qty_delta_total += d

    qty_count = len(qty_changes_list)
    qty_min = float(min(qty_changes_list)) if qty_changes_list else 0.0
    qty_max = float(max(qty_changes_list)) if qty_changes_list else 0.0
    qty_avg = float(np.mean(qty_changes_list)) if qty_changes_list else 0.0

    exact_matches = len(exact_pairs) if exact_pairs else 0
    fuzzy_matches = len(fuzzy_pairs) if fuzzy_pairs else 0
    low_confidence = sum(1 for e in (fuzzy_pairs or []) if e[3] < config.get("replace_threshold", 0.70))

    total_matches = exact_matches + fuzzy_matches + low_confidence
    if total_matches > 0:
        weighted_matches = exact_matches * 1.0 + fuzzy_matches * 0.7 + low_confidence * 0.3
        divergence_factor = 1 - (deviation_pct / 100)
        volatility_factor = 1 - ((added_count + deleted_count + replaced_count + modified_count) / max(old_count, 1))
        overall_similarity = (weighted_matches / total_matches) * divergence_factor * volatility_factor
    else:
        overall_similarity = 0.0

    field_counter = Counter()
    changed_examples = []

    for df in [df_modified, df_replaced]:
        if df is not None and not df.empty:
            for _, r in df.iterrows():
                types = r.get("Change Type", "")
                if isinstance(types, str) and types:
                    for t in [t.strip() for t in types.split(",") if t.strip()]:
                        field_counter[t] += 1

                cid = r.get("Old Component number") or r.get("New Component number") or r.get("Old Component") or r.get("Component")
                if cid and len(changed_examples) < 10:
                    changed_examples.append(str(cid))

    parents, max_depth = set(), 0
    for flat in [flat_old, flat_new]:
        if flat is not None:
            for _, r in flat.iterrows():
                path = r.get("ParentPath") or ""
                if path:
                    parents.add(path.split("/")[0])
                    max_depth = max(max_depth, path.count("/") + 1)

    parent_count = len(parents)

    barrel_mismatch = 0
    segment_counter = Counter()

    if df_replaced is not None and not df_replaced.empty:
        for _, r in df_replaced.iterrows():
            detail = r.get("Change Detail") or ""
            if "Component number barrels" in detail:
                barrel_mismatch += 1
                segment_counter.update(re.findall(r"([A-Za-z0-9_-]+):", detail))

    weights = config.get("divergence_weights", {"added": 1.0, "deleted": 1.0, "changed": 1.0, "qty": 0.2})

    added_pct = (added_count / old_count * 100) if old_count else 0.0
    deleted_pct = (deleted_count / old_count * 100) if old_count else 0.0
    replaced_pct = (replaced_count / old_count * 100) if old_count else 0.0
    modified_pct = (modified_count / old_count * 100) if old_count else 0.0

    qty_den = (flat_old[CAN_COL_QTY].sum() + 1e-9) if flat_old is not None else 1.0
    qty_pct = (abs(qty_delta_total) / qty_den) * 100 if qty_den != 0 else 0.0

    divergence_index = round(added_pct * weights.get("added", 1.0) + deleted_pct * weights.get("deleted", 1.0) + (replaced_pct + modified_pct) * weights.get("changed", 1.0) + qty_pct * weights.get("qty", 0.2), 2)

    redundancy_pct = carry_over_pct
    volatility_score = round((added_count + deleted_count + replaced_count + modified_count) / (old_count if old_count else 1) * 100, 2)

    risk_flag = "LOW"
    if divergence_index > 30 or volatility_score > 25:
        risk_flag = "HIGH"
    elif divergence_index > 15 or volatility_score > 12:
        risk_flag = "MEDIUM"

    top_changed_fields = dict(field_counter.most_common(10))
    total_changes_count = sum(field_counter.values()) or 1
    field_dist_pct = {k: round(v / total_changes_count * 100, 2) for k, v in top_changed_fields.items()}

    high_impact = []
    if qty_changes:
        for s in sorted(qty_changes, key=lambda x: abs(x["delta"]), reverse=True)[:5]:
            high_impact.append({"ComponentID": s.get("component"), "Delta": s.get("delta")})

    summary = {
        "Overview": {
            "Base Components": old_count,
            "Revised Components": new_count,
            "Total Compared Components": total_compared,
            "Deviation %": deviation_pct,
            "Processing Time (sec)": round(end_ts - start_ts, 3),
            "Run Timestamp": datetime.now().strftime("%d/%m/%y %I:%M %p"),
            "Root Level": root_level
        },
        "Status Breakdown": {
            "Newly Added": {"count": added_count, "percent": round((added_count / (new_count if new_count else 1)) * 100, 2)},
            "Deleted": {"count": deleted_count, "percent": round((deleted_count / (old_count if old_count else 1)) * 100, 2)},
            "Replaced": {"count": replaced_count, "percent": round((replaced_count / (old_count if old_count else 1)) * 100, 2)},
            "Modified": {"count": modified_count, "percent": round((modified_count / (old_count if old_count else 1)) * 100, 2)},
            "Carry-Over": {"count": unchanged_count, "percent": carry_over_pct},
            "Skipped": {"count": skipped_count}
        },
        "Quantity Impact": {
            "Count of Quantity Changes": qty_count,
            "Min Quantity Delta": qty_min,
            "Max Quantity Delta": qty_max,
            "Average Quantity Delta": round(qty_avg, 3),
            "Total Net Delta": round(qty_delta_total, 3),
            "High Impact Components": high_impact
        },
        "Change Details": {
            "Top Changed Fields": top_changed_fields,
            "Field Change Distribution %": field_dist_pct,
            "Example Changed Components": changed_examples
        },
        "Hierarchy Insights": {
            "Parent Components Affected": parent_count,
            "Max Depth Impacted": max_depth
        },
        "Barrel Analysis": {
            "Replacements due to Barrel Mismatch": barrel_mismatch,
            "Most Common Mismatch Segment": segment_counter.most_common(3)
        },
        "Matching Diagnostics": {
            "Exact Matches": exact_matches,
            "Fuzzy Matches": fuzzy_matches,
            "Low Confidence Matches": low_confidence,
            "Overall Similarity Score": round(overall_similarity * 100, 2)
        },
        "Health Indicators": {
            "Redundancy %": redundancy_pct,
            "Volatility Score": volatility_score,
            "Risk Flag": risk_flag
        },
        "Overall Divergence Index": divergence_index
    }

    return summary

def list_headers(paths: List[str]) -> List[str]:
    cols = set()
    for p in paths:
        if os.path.exists(p):
            try:
                df = load_and_preprocess(p)
                cols.update(list(df.columns))
            except Exception:
                continue
    return sorted(cols)

def _safe_default_output_name(old_path: str, new_path: str) -> str:
    def _slug(path: str) -> str:
        stem = os.path.splitext(os.path.basename(path))[0]
        cleaned = re.sub(r"[^A-Za-z0-9._-]+", "_", stem).strip("_")
        return cleaned[:24] or "input"

    return f"BOM_comparison_{_slug(old_path)}_vs_{_slug(new_path)}.xlsx"

def _detect_root_level(df: pd.DataFrame, level_col: str, default_level: int = 1) -> int:
    if df is None or level_col not in df.columns or df.empty:
        return default_level
    lvl_series = pd.to_numeric(df[level_col], errors="coerce").dropna()
    if lvl_series.empty:
        return default_level
    return int(lvl_series.min())

def main():
    p = argparse.ArgumentParser(description="Unified BOM Comparator (tree + semantic + dynamic per-column compare).")
    p.add_argument("--old", required=False, default=DEFAULT_OLD, help="Path to OLD BOM Excel")
    p.add_argument("--new", required=False, default=DEFAULT_NEW, help="Path to NEW BOM Excel")
    p.add_argument("--output", required=False, default=DEFAULT_OUTPUT, help="Output Excel path")
    p.add_argument("--config", required=False, default=DEFAULT_CONFIG, help="Optional JSON config path (component segment rules, thresholds)")
    p.add_argument("--dynamic_pairs", required=False, default="[]", help="JSON string of mapped optional comparison pairs")
    p.add_argument("--print_columns", required=False, default="", help="Comma-separated columns to print verbatim in Excel")
    p.add_argument("--skip_choice", required=False, default="Y", choices=["Y", "X"], help="Skip choice for Special Procurement rules (Y or X)")
    p.add_argument("--list_headers", action="store_true", help="Only list merged headers of old/new and exit")
    p.add_argument("--column_mappings", required=False, default=None, help="JSON string of column mappings from canonical to user-defined names")
    p.add_argument("--map_old", required=False, default=None, help="JSON mapping for OLD file")
    p.add_argument("--map_new", required=False, default=None, help="JSON mapping for NEW file")

    args = p.parse_args()

    if args.list_headers:
        headers = list_headers([args.old, args.new])
        print(json.dumps({"headers": headers}, indent=2))
        return

    start_ts = time.time()

    if args.output == DEFAULT_OUTPUT:
        output_filename = _safe_default_output_name(args.old, args.new)
    else:
        output_filename = args.output

    print("Loading configuration...")
    config = load_config(args.config)
    choice = args.skip_choice 

    print(f"Reading OLD BOM: {args.old}")
    df_old_raw = load_and_preprocess(args.old)

    print(f"Reading NEW BOM: {args.new}")
    df_new_raw = load_and_preprocess(args.new)

    map_old_dict = None
    map_new_dict = None
    
    if args.map_old:
        map_old_dict = json.loads(args.map_old)
    if args.map_new:
        map_new_dict = json.loads(args.map_new)
        
    if not map_old_dict and args.column_mappings:
        map_old_dict = json.loads(args.column_mappings)
    if not map_new_dict and args.column_mappings:
        map_new_dict = json.loads(args.column_mappings)

    # Parse dynamic pairs
    dynamic_pairs = []
    try:
        if args.dynamic_pairs:
            dynamic_pairs = json.loads(args.dynamic_pairs)
            print(f"Parsed {len(dynamic_pairs)} dynamic comparison pairs.")
    except Exception as e:
        print(f"Error parsing dynamic pairs: {e}")

    print("Mapping & normalizing BOM columns...")
    # Normalize Old using SAP map, New using PLM map
    flat_old_norm, map_old = normalize_bom_dataframe(df_old_raw, allowed_columns=config.get("allowed_columns"), user_map=map_old_dict, default_alias_map=_SAP_COL_ALIAS_MAP)
    flat_new_norm, map_new = normalize_bom_dataframe(df_new_raw, allowed_columns=config.get("allowed_columns"), user_map=map_new_dict, default_alias_map=_PLM_COL_ALIAS_MAP)

    # --- NEW: Root Pruning for Cross-Over ---
    try:
        root_level_old = _detect_root_level(flat_old_norm, CAN_COL_LEVEL, default_level=1)
        root_level_new = _detect_root_level(flat_new_norm, CAN_COL_LEVEL, default_level=1)
        
        # Calculate offset just for logging or summary, but main logic is PRUNING
        # If Old=0, New=1 -> Prune Old
        if root_level_old == 0 and root_level_new == 1:
            print("!!! CROSSOVER DETECTED (0 vs 1) !!! Pruning Old BOM Level 0 Root.")
            flat_old_norm = prune_root_level(flat_old_norm)
            root_level_old = 1 # Now it starts at 1
            
        # If Old=1, New=0 -> Prune New
        elif root_level_old == 1 and root_level_new == 0:
            print("!!! CROSSOVER DETECTED (1 vs 0) !!! Pruning New BOM Level 0 Root.")
            flat_new_norm = prune_root_level(flat_new_norm)
            root_level_new = 1 # Now it starts at 1
            
        root_level = min(root_level_old, root_level_new)
        print(f"Effective Root Levels after Pruning check: Old={root_level_old}, New={root_level_new}")
        
    except Exception as e:
        print(f"Warning: Root level detection/pruning failed. Error: {e}")
        root_level = 1
    # --- END NEW ---

    roots_old = build_forest(flat_old_norm)
    roots_new = build_forest(flat_new_norm)

    print("Flattening trees (respecting Special Procurement skip choice)...")
    flat_old = flatten_forest(roots_old, choice)
    flat_new = flatten_forest(roots_new, choice)

    df_skipped_old = flat_old[flat_old["SkippedChildren"]].copy() if "SkippedChildren" in flat_old.columns else pd.DataFrame()
    if not df_skipped_old.empty:
        df_skipped_old["Source"] = "Old"

    df_skipped_new = flat_new[flat_new["SkippedChildren"]].copy() if "SkippedChildren" in flat_new.columns else pd.DataFrame()
    if not df_skipped_new.empty:
        df_skipped_new["Source"] = "New"

    df_skipped = pd.concat([df_skipped_old, df_skipped_new], ignore_index=True, sort=False)
    if df_skipped.empty:
        df_skipped = pd.DataFrame(columns=[CAN_COL_LEVEL, CAN_COL_COMP, CAN_COL_DESC, CAN_COL_CS, CAN_COL_QTY, "ParentPath", "SkippedChildren", "Source"])

    skip_flag = ("SkippedChildren" in flat_old.columns and flat_old["SkippedChildren"].any()) or \
                ("SkippedChildren" in flat_new.columns and flat_new["SkippedChildren"].any())

    print("Computing description embeddings (with on-disk cache)...")
    if CAN_COL_DESC not in flat_old.columns:
        flat_old[CAN_COL_DESC] = ""
    if CAN_COL_DESC not in flat_new.columns:
        flat_new[CAN_COL_DESC] = ""

    old_emb = embed_with_cache_column(flat_old, CAN_COL_DESC, "old_desc")
    new_emb = embed_with_cache_column(flat_new, CAN_COL_DESC, "new_desc")

    print("Step 1 & 2 â€” exact (Strict) and greedy-best-fit matching...")
    # PASS LEVEL OFFSET REMOVED (Pruning handles it)
    all_matched_pairs, rem_old_for_fuzzy, rem_new_for_fuzzy = exact_match_indices(flat_old, flat_new)
    print(f"Remaining after greedy: old={len(rem_old_for_fuzzy)} new={len(rem_new_for_fuzzy)}")

    print("Step 3 â€” fuzzy/semantic level-wise matching (SBERT +/- segments)...")
    fuzzy_pairs, rem_old_final, rem_new_final = fuzzy_match_levelwise(
        flat_old, flat_new, old_emb, new_emb, rem_old_for_fuzzy, rem_new_for_fuzzy, config
    )
    
    # dynamic_cols deprecated in favor of dynamic_pairs

    print("Classifying pairs into Carry-Over / Modified / Replaced...")
    unchanged, modified, replaced = classify_pairs(
        flat_old, flat_new, df_old_raw, df_new_raw, 
        all_matched_pairs, fuzzy_pairs, old_emb, new_emb, config, dynamic_pairs
    )

    deleted_indices = sorted(list(rem_old_final))
    added_indices = sorted(list(rem_new_final))

    print("Updating node statuses...")
    update_node_statuses(roots_old, roots_new, unchanged, modified, replaced, rem_old_final, rem_new_final)

    print("Propagating hierarchical status up trees...")
    propagate_hierarchical_status(roots_old)
    propagate_hierarchical_status(roots_new)

    print("Building output rows (with SourceIndex for dynamic lookup)...")
    df_unchanged, df_modified, df_replaced, df_added, df_deleted, df_all = build_output_rows_from_matches(
        flat_old, flat_new, unchanged, modified, replaced, deleted_indices, added_indices, config, roots_old, roots_new,
        root_level
    )

    def safe_map_factory(source_df, col_name):
        def safe_mapper(idx):
            try:
                if pd.isna(idx) or idx is None or idx == '': return None
                idx = int(float(idx))
                if idx in source_df.index: return source_df.at[idx, col_name]
                return None
            except (ValueError, KeyError, TypeError): return None
        return safe_mapper

    def map_all_input_columns_to_dfs(target_dfs, source_df_raw, source_label: str):
        """Copy every raw input column into the comparison/status sheets.

        This intentionally leaves the raw BOM tabs (old_bom/new_bom/new_clm_bom)
        unchanged and only enriches the comparison tabs where users need the extra context.
        """
        if source_df_raw is None or source_df_raw.empty:
            return

        source_idx_col = "Old_SourceIndex" if source_label == 'old' else "New_SourceIndex"
        prefix = "Old " if source_label == 'old' else "New "

        for col_name in source_df_raw.columns:
            if col_name == "SourceIndex":
                continue

            out_col = f"{prefix}{col_name}"
            mapper = safe_map_factory(source_df_raw, col_name)

            for df in target_dfs:
                if df is not None and not df.empty and source_idx_col in df.columns:
                    df[out_col] = df[source_idx_col].map(mapper)

    print("Restoring original (un-cleaned) Component Numbers and Descriptions to final report...")

    orig_col_comp_old = map_old.get(CAN_COL_COMP)
    orig_col_desc_old = map_old.get(CAN_COL_DESC)
    orig_col_comp_new = map_new.get(CAN_COL_COMP)
    orig_col_desc_new = map_new.get(CAN_COL_DESC)
    
    all_result_dfs = [df_unchanged, df_modified, df_replaced, df_added, df_deleted, df_all, df_skipped]

    for df in all_result_dfs:
        if df is None or df.empty:
            continue
            
        if "Old_SourceIndex" in df.columns:
            if f"Old {CAN_COL_COMP}" in df.columns and orig_col_comp_old and orig_col_comp_old in df_old_raw.columns:
                df[f"Old {CAN_COL_COMP}"] = df["Old_SourceIndex"].map(safe_map_factory(df_old_raw, orig_col_comp_old))
            if f"Old {CAN_COL_DESC}" in df.columns and orig_col_desc_old and orig_col_desc_old in df_old_raw.columns:
                df[f"Old {CAN_COL_DESC}"] = df["Old_SourceIndex"].map(safe_map_factory(df_old_raw, orig_col_desc_old))
                
        if "New_SourceIndex" in df.columns:
            if f"New {CAN_COL_COMP}" in df.columns and orig_col_comp_new and orig_col_comp_new in df_new_raw.columns:
                df[f"New {CAN_COL_COMP}"] = df["New_SourceIndex"].map(safe_map_factory(df_new_raw, orig_col_comp_new))
            if f"New {CAN_COL_DESC}" in df.columns and orig_col_desc_new and orig_col_desc_new in df_new_raw.columns:
                df[f"New {CAN_COL_DESC}"] = df["New_SourceIndex"].map(safe_map_factory(df_new_raw, orig_col_desc_new))

    print_cols = [c.strip() for c in args.print_columns.split(",") if c.strip()]
    
    BULK_MAT_COL = "Bulk Mat. (Item Mat)"
    BULK_ALIASES = ["Bulk Mat. (Item Mat)", "Bulk Mat", "Bulk Material", "Bulk Indicator", "Bulk", "Bulk Item", "Bulk Mat.", "Item Mat", "Bulk IND", "BULK IND"]

    bulk_found = False
    
    def _standardize_bulk_col(df, aliases, standard_name):
        found = False
        cols_lower = {c.lower(): c for c in df.columns}
        
        for alias in aliases:
            alias_lower = alias.lower()
            if alias_lower in cols_lower:
                actual_col_name = cols_lower[alias_lower]
                if actual_col_name != standard_name:
                    print(f"Standardizing column '{actual_col_name}' -> '{standard_name}'")
                    df.rename(columns={actual_col_name: standard_name}, inplace=True)
                found = True
                break 
        return found

    if _standardize_bulk_col(df_old_raw, BULK_ALIASES, BULK_MAT_COL): bulk_found = True
    if _standardize_bulk_col(df_new_raw, BULK_ALIASES, BULK_MAT_COL): bulk_found = True

    if bulk_found:
        if BULK_MAT_COL not in print_cols:
            print(f"Bulk Material column found/standardized, adding to output.")
            print_cols.append(BULK_MAT_COL)

    # --- INTEGRATE DYNAMIC & PRINT COLUMNS INTO OUTPUT DATAFRAMES ---
    
    # Helper to map values to output dfs
    def map_val_to_dfs(col_name_in_file, output_col_name, source_file_enum): 
        # source_file_enum: 'old' or 'new'
        target_dfs = [df_unchanged, df_modified, df_replaced, df_added, df_deleted, df_all]
        source_df_raw = df_old_raw if source_file_enum == 'old' else df_new_raw
        source_idx_col = "Old_SourceIndex" if source_file_enum == 'old' else "New_SourceIndex"
        
        if col_name_in_file not in source_df_raw.columns:
            return
        
        # FIX: Use safe mapper to prevent int('') crashes
        mapper = safe_map_factory(source_df_raw, col_name_in_file)

        for df in target_dfs:
            if df is not None and not df.empty and source_idx_col in df.columns:
                df[output_col_name] = df[source_idx_col].map(mapper)

    if dynamic_pairs:
        print(f"Integrating {len(dynamic_pairs)} dynamic comparison pairs...")
        for pair in dynamic_pairs:
            old_c = pair.get('old')
            new_c = pair.get('new')
            # Add them as separate columns "Old [Name]" and "New [Name]" to visualize the comparison
            if old_c: map_val_to_dfs(old_c, f"Old {old_c}", 'old')
            if new_c: map_val_to_dfs(new_c, f"New {new_c}", 'new')

    if print_cols:
        print(f"Integrating print-only columns: {print_cols}")
        for col in print_cols:
            # Try to map from Old (if exists)
            if col in df_old_raw.columns:
                map_val_to_dfs(col, f"Old {col}", 'old')
            
            # Try to map from New (if exists)
            if col in df_new_raw.columns:
                map_val_to_dfs(col, f"New {col}", 'new')

    # Enrich the comparison/status sheets with every raw input column for convenience.
    # Keep the raw BOM tabs untouched.
    map_all_input_columns_to_dfs(all_result_dfs, df_old_raw, 'old')
    map_all_input_columns_to_dfs(all_result_dfs, df_new_raw, 'new')

    # Update Skipped DataFrame with dynamic/print columns too (best effort)
    if df_skipped is not None and not df_skipped.empty:
        all_extra_cols = set()
        for pair in dynamic_pairs:
            if pair.get('old'): all_extra_cols.add((pair['old'], f"Old {pair['old']}", 'old'))
            if pair.get('new'): all_extra_cols.add((pair['new'], f"New {pair['new']}", 'new'))
        for col in print_cols:
            all_extra_cols.add((col, f"Old {col}", 'old'))
            all_extra_cols.add((col, f"New {col}", 'new'))
            
        for col_in_file, out_col_name, source in all_extra_cols:
            df_skipped[out_col_name] = None
            mask = df_skipped['Source'] == ('Old' if source == 'old' else 'New')
            source_df = df_old_raw if source == 'old' else df_new_raw
            
            if col_in_file in source_df.columns:
                 # FIX: use safe map logic
                 mapper = safe_map_factory(source_df, col_in_file)
                 df_skipped.loc[mask, out_col_name] = df_skipped.loc[mask, 'SourceIndex'].map(mapper)

    # --- START: Apply strict business-category rules described by user ---
    # The rules rely on available raw/cleaned columns. We implement a post-pass
    # that updates `Primary Status` for every row in `df_all` and then
    # rebuilds the per-category dataframes so downstream code/export uses
    # these canonical categories.

    def _find_old_new_col(base_cols, candidates):
        """Return (old_col, new_col) names from df_all if present (case-insensitive)."""
        old_col = next((c for c in base_cols if c.lower().startswith('old ') and any(tok in c.lower() for tok in candidates)), None)
        new_col = next((c for c in base_cols if c.lower().startswith('new ') and any(tok in c.lower() for tok in candidates)), None)
        return old_col, new_col

    def _get_design_group_from_row(row, cfg):
        # Prefer explicit Old/New Design Group columns copied from raw files
        def _pick_first_existing(raw: Dict[str, Any], keys: List[str]) -> Optional[str]:
            for k in keys:
                v = raw.get(k)
                if v not in (None, ''):
                    return str(v)
            return None

        dg = _pick_first_existing(row, ['Old Design Group', 'New Design Group'])
        if dg:
            return str(dg).strip()
        # Fallback: parse component numbers (Old then New)
        comp = _pick_first_existing(row, ['Old Component number', 'New Component number'])
        if comp:
            parsed = parse_component_number(str(comp), cfg)
            # Try to return segment label if provided, else segment value for 'Design Group'
            for seg in parsed.get('segments', []):
                if seg.get('name', '').lower().startswith('design group') or seg.get('name', '').lower().startswith('design'):
                    val = seg.get('label') or seg.get('value') or ''
                    if isinstance(val, dict):
                        val = val.get('description', '') or val.get('name', '') or str(val)
                    return str(val).strip()
        return ''

    # Determine candidate columns present in df_all for revision/description/quantity
    all_cols = list(df_all.columns) if 'df_all' in locals() and df_all is not None else []
    rev_old_col, rev_new_col = _find_old_new_col(all_cols, ['revision', 'rev'])
    desc_old_col, desc_new_col = _find_old_new_col(all_cols, ['object description', 'description', 'desc'])
    qty_old_col, qty_new_col = _find_old_new_col(all_cols, ['component quantity', 'quantity', 'qty'])

    # Helper to detect whether any new item exists in same design group
    def _new_item_in_same_design_group(design_group, flat_new_df, cfg):
        if not design_group:
            return False
        # Check explicit 'Design Group' raw column first
        if 'Design Group' in flat_new_df.columns:
            if any(str(v).strip() == design_group for v in flat_new_df['Design Group'].dropna()):
                return True
        elif 'New Design Group' in flat_new_df.columns:
            if any(str(v).strip() == design_group for v in flat_new_df['New Design Group'].dropna()):
                return True
        # Fallback: parse all new component numbers
        for _, r in flat_new_df.iterrows():
            parsed = parse_component_number(str(r.get(CAN_COL_COMP, '')), cfg)
            for seg in parsed.get('segments', []):
                if seg.get('name', '').lower().startswith('design'):
                    val = seg.get('label') or seg.get('value') or ''
                    if isinstance(val, dict):
                        val = val.get('description', '') or val.get('name', '') or str(val)
                    val = str(val).strip()
                    if val == design_group:
                        return True
        return False

    # Apply rules row-wise on df_all (best-effort; missing columns are treated as 'same')
    try:
        for idx, r in df_all.iterrows():
            base_exists = not pd.isna(r.get('Old_SourceIndex')) and r.get('Old_SourceIndex') not in (None, '')
            new_exists = not pd.isna(r.get('New_SourceIndex')) and r.get('New_SourceIndex') not in (None, '')

            # Compare revision
            if rev_old_col and rev_new_col:
                rev_old = str(r.get(rev_old_col, '')).strip()
                rev_new = str(r.get(rev_new_col, '')).strip()
                revision_same = (rev_old == rev_new)
            else:
                revision_same = True

            # Compare quantity (numeric)
            if qty_old_col and qty_new_col:
                try:
                    q_old = float(str(r.get(qty_old_col, 0)).replace(',', '') or 0)
                    q_new = float(str(r.get(qty_new_col, 0)).replace(',', '') or 0)
                    quantity_same = (abs(q_old - q_new) < 1e-9)
                except Exception:
                    quantity_same = True
            else:
                quantity_same = True

            # Compare description (cleaned columns exist as 'Old Object description'/ 'New Object description')
            if desc_old_col and desc_new_col:
                desc_old = str(r.get(desc_old_col, '')).strip()
                desc_new = str(r.get(desc_new_col, '')).strip()
                description_same = (desc_old == desc_new)
            else:
                # Try canonical cleaned description columns
                desc_old = str(r.get(f'Old {CAN_COL_DESC}', '')).strip()
                desc_new = str(r.get(f'New {CAN_COL_DESC}', '')).strip()
                description_same = (desc_old == desc_new)

            same_design_group_new_available = False
            if base_exists and not new_exists:
                dg = _get_design_group_from_row(r, config)
                same_design_group_new_available = _new_item_in_same_design_group(dg, flat_new, config)

            # Apply the main logic requested
            new_category = None
            if base_exists and new_exists:
                # If engine already marked it as Replaced (e.g. fuzzy match with shape diff or comp ID diff), keep it
                if r.get('Primary Status') == 'Replaced':
                    new_category = 'Replaced'
                else:
                    if revision_same and quantity_same and description_same:
                        new_category = 'Carry-Over'
                    else:
                        new_category = 'Modified'
            elif (not base_exists) and new_exists:
                new_category = 'Newly Added'
            elif base_exists and (not new_exists):
                # Removed aggressive design-group-based replacement override to match user's manual counts.
                # Matching should be handled by the similarity matching engine instead.
                new_category = 'Deleted'

            if new_category:
                df_all.at[idx, 'Primary Status'] = new_category
                df_all.at[idx, 'Exclusive'] = new_category in ('Modified', 'Newly Added', 'Replaced')
    except Exception as e:
        print(f"Warning: business-category post-pass failed: {e}")

    # Rebuild per-category frames from updated df_all so downstream steps remain consistent
    df_unchanged = df_all[df_all.get('Primary Status') == 'Carry-Over'].reset_index(drop=True)
    df_modified = df_all[df_all.get('Primary Status') == 'Modified'].reset_index(drop=True)
    df_replaced = df_all[df_all.get('Primary Status') == 'Replaced'].reset_index(drop=True)
    df_added = df_all[df_all.get('Primary Status') == 'Newly Added'].reset_index(drop=True)
    df_deleted = df_all[df_all.get('Primary Status') == 'Deleted'].reset_index(drop=True)

    # --- END: business-category rules ---
    end_ts = time.time()


    summary = generate_summary(df_unchanged, df_modified, df_replaced, df_added, df_deleted, df_skipped,
                             flat_old, flat_new, all_matched_pairs, fuzzy_pairs, config, start_ts, end_ts,
                             root_level=root_level)

    with open("summary.json", "w") as f:
        json.dump(summary, f, indent=2)

    print("\n=== SUMMARY (JSON) ===")
    print(json.dumps(summary, indent=2))

    print(f"\nExporting results to: {output_filename} ...")
    export_excel(output_filename, df_unchanged, df_modified, df_replaced, df_added, df_deleted,
                df_all, df_skipped, skip_flag, df_old_raw=df_old_raw, df_new_raw=df_new_raw)

    print("Done.")

    if _FAISS_OK:
        print("FAISS is available â€” nearest-neighbor accelerated.")
    else:
        print("FAISS not found; using NumPy fallback. Install faiss-cpu for acceleration.")

    if fuzz is None:
        print("rapidfuzz not found â€” fuzzy matching will fall back to exact behaviour. Install rapidfuzz for fuzzy scoring.")


if __name__ == "__main__":
    main()
