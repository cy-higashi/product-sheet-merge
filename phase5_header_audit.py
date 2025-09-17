import os
import sys
import re
from typing import List, Tuple, Dict, Any, Optional
import pandas as pd

# Utility: simple normalization for header comparison
_def_ws_re = re.compile(r"\s+")

def normalize_header(text: Any) -> str:
    if pd.isna(text):
        return ""
    s = str(text)
    s = s.replace("\n", " ")
    s = _def_ws_re.sub(" ", s).strip()
    return s


def find_header_row(df: pd.DataFrame) -> Optional[int]:
    # B列が「項目」の行
    for i in range(len(df)):
        try:
            if str(df.iat[i, 1]).strip() == "項目":
                return i
        except Exception:
            continue
    return None


def build_effective_headers(df: pd.DataFrame, header_row: int, lookahead_rows: int = 2) -> List[str]:
    """
    For each column >= 2 (C列以降), pick the last non-empty candidate among:
    - row header_row (main header)
    - header_row+1, header_row+2 (sub headers)
    Return a list aligned to df columns (length == df.shape[1]) with raw strings (not normalized),
    keeping original strings for later reporting while also providing normalized variants separately.
    """
    num_cols = df.shape[1]
    effective: List[str] = []
    for j in range(num_cols):
        candidates: List[str] = []
        for k in range(lookahead_rows + 1):
            r = header_row + k
            if r < len(df):
                candidates.append(df.iat[r, j])
        # choose last non-empty
        chosen = ""
        for val in reversed(candidates):
            if pd.notna(val) and str(val).strip() != "":
                chosen = str(val)
                break
        effective.append(chosen)
    return effective


def first_data_row_index(df: pd.DataFrame, header_row: int) -> Optional[int]:
    for i in range(header_row + 1, len(df)):
        row = df.iloc[i, 2:]
        if row.notna().any():
            return i
    return None


def audit_file(file_path: str) -> Dict[str, Any]:
    try:
        df = pd.read_excel(file_path, header=None)
    except Exception as e:
        return {"file": file_path, "error": f"read_error: {e}"}

    header_row = find_header_row(df)
    if header_row is None:
        return {"file": file_path, "error": "header_row_not_found"}

    row_main = [df.iat[header_row, j] if j < df.shape[1] else None for j in range(df.shape[1])]
    row_sub1 = [df.iat[header_row+1, j] if header_row+1 < len(df) else None for j in range(df.shape[1])]
    row_sub2 = [df.iat[header_row+2, j] if header_row+2 < len(df) else None for j in range(df.shape[1])]

    eff = build_effective_headers(df, header_row, lookahead_rows=2)
    eff_norm = [normalize_header(x) for x in eff]
    main_norm = [normalize_header(x) for x in row_main]
    sub1_norm = [normalize_header(x) for x in row_sub1]

    data_idx = first_data_row_index(df, header_row)
    sample_values: List[Any] = []
    if data_idx is not None:
        sample_values = [df.iat[data_idx, j] for j in range(df.shape[1])]

    # stats: columns where main header is empty but subheaders exist
    missing_main_with_sub = []
    for j in range(2, df.shape[1]):
        if main_norm[j] == "" and (sub1_norm[j] != "" or normalize_header(row_sub2[j]) != ""):
            missing_main_with_sub.append(j)

    # collect duplicates in effective headers (normalized)
    counts: Dict[str, int] = {}
    for name in eff_norm[2:]:
        counts[name] = counts.get(name, 0) + 1
    duplicates = sorted([name for name, c in counts.items() if name and c > 1])

    return {
        "file": file_path,
        "header_row": header_row,
        "num_cols": df.shape[1],
        "main_headers": row_main,
        "sub_headers": row_sub1,
        "effective_headers": eff,
        "effective_headers_norm": eff_norm,
        "sample_row_index": data_idx,
        "sample_values": sample_values,
        "missing_main_with_sub_cols": missing_main_with_sub,
        "duplicate_effective_headers": duplicates,
    }


def load_master_headers(all_collect_path: str) -> List[str]:
    try:
        df = pd.read_excel(all_collect_path, header=None)
    except Exception:
        return []
    if len(df) == 0:
        return []
    headers = [normalize_header(v) for v in df.iloc[0, 2:].tolist()]
    return headers


def main():
    municipality = os.environ.get("MUNICIPALITY_NAME")
    if not municipality:
        print("[ERROR] Set MUNICIPALITY_NAME env var.")
        sys.exit(1)

    base_dir = os.path.join(r"G:\共有ドライブ\★OD\99_商品管理\DATA\Phase3\HARV", municipality)

    # target files: prefer normalized
    files = [f for f in os.listdir(base_dir) if f.startswith("PAT") and f.endswith("_normalized.xlsx")]
    if not files:
        files = [f for f in os.listdir(base_dir) if f.startswith("PAT") and f.endswith(".xlsx")]
    files = sorted(files)

    all_collect_path = os.path.join(base_dir, "all_collect.xlsx")
    master_headers_norm = load_master_headers(all_collect_path)

    print(f"Municipality: {municipality}")
    print(f"Base dir: {base_dir}")
    print(f"all_collect exists: {os.path.exists(all_collect_path)}  (headers={len(master_headers_norm)})")

    records = []
    for f in files:
        fp = os.path.join(base_dir, f)
        info = audit_file(fp)
        records.append(info)

    # Summary: how many columns rely on subheaders across files
    total_missing_main = sum(len(r.get("missing_main_with_sub_cols", [])) for r in records if "error" not in r)
    print(f"\n[SUMMARY] Files analyzed: {len(records)}  missing-main-with-sub columns (total): {total_missing_main}")

    # If master headers exist, compute per-file coverage and mismatches
    if master_headers_norm:
        print("\n[MASTER vs EFFECTIVE] Coverage by file:")
        for r in records:
            if "error" in r:
                print(f" - {os.path.basename(r['file'])}: ERROR: {r['error']}")
                continue
            eff = r["effective_headers_norm"][2:]
            eff_set = set([e for e in eff if e])
            master_set = set([m for m in master_headers_norm if m])
            covered = len(eff_set & master_set)
            print(f" - {os.path.basename(r['file'])}: effective={len(eff_set)} covered_in_master={covered} duplicates={len(r['duplicate_effective_headers'])}")

    # Emit a CSV for detailed inspection
    out_rows = []
    for r in records:
        if "error" in r:
            out_rows.append({"file": r["file"], "error": r["error"]})
            continue
        num_cols = r["num_cols"]
        for j in range(2, num_cols):
            out_rows.append({
                "file": os.path.basename(r["file"]),
                "col_index": j,
                "main": normalize_header(r["main_headers"][j]),
                "sub1": normalize_header(r["sub_headers"][j]),
                "effective": normalize_header(r["effective_headers"][j]),
                "sample": r["sample_values"][j] if r["sample_values"] else None,
                "is_missing_main_but_has_sub": j in r["missing_main_with_sub_cols"],
            })
    out_df = pd.DataFrame(out_rows)
    out_dir = os.path.join(base_dir, "diagnostics")
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, "phase5_header_audit.csv")
    out_df.to_csv(out_path, index=False, encoding="utf-8-sig")
    print(f"\n[OUTPUT] Detailed per-column audit -> {out_path}")

    # Print top problematic columns per file (where main empty but sub present)
    print("\n[Details] Columns relying on subheaders (main empty):")
    for r in records:
        if "error" in r:
            continue
        cols = r["missing_main_with_sub_cols"]
        if cols:
            print(f" - {os.path.basename(r['file'])}: {len(cols)} columns -> {cols[:10]}{'...' if len(cols) > 10 else ''}")

if __name__ == "__main__":
    main()
