import os
import sys
import re
from typing import List, Dict, Any, Optional
import pandas as pd

TARGET_HEADERS = [
    "商品名",
    "商品名（伝票記載用）",
    "返礼品説明",
    "ご担当者様",
    "発送元名称",
    "住所",
    "TEL",
]

_ws_re = re.compile(r"\s+")

def norm(s: Any) -> str:
    if pd.isna(s):
        return ""
    s = str(s).replace("\n", " ")
    s = _ws_re.sub(" ", s).strip()
    return s


def find_header_row(df: pd.DataFrame) -> Optional[int]:
    for i in range(len(df)):
        try:
            if str(df.iat[i, 1]).strip() == "項目":
                return i
        except Exception:
            pass
    return None


def first_data_row(df: pd.DataFrame, header_row: int) -> Optional[int]:
    for i in range(header_row + 1, len(df)):
        if df.iloc[i, 2:].notna().any():
            return i
    return None


def load_master_headers(base_dir: str) -> List[str]:
    ac = os.path.join(base_dir, "all_collect.xlsx")
    if not os.path.exists(ac):
        return []
    try:
        df = pd.read_excel(ac, header=None)
    except Exception:
        return []
    if len(df) == 0:
        return []
    return [norm(x) for x in df.iloc[0, 2:].tolist()]


def analyze_file(fp: str, master_headers_norm: List[str], targets_norm: List[str]) -> Dict[str, Any]:
    try:
        df = pd.read_excel(fp, header=None)
    except Exception as e:
        return {"file": fp, "error": f"read_error: {e}"}

    hr = find_header_row(df)
    if hr is None:
        return {"file": fp, "error": "header_row_not_found"}

    sh_raw = df.iloc[hr, 2:].tolist()
    sh_norm = [norm(x) for x in sh_raw]
    di = first_data_row(df, hr)
    sample = df.iloc[di].tolist() if di is not None else []

    result_rows = []
    for th in targets_norm:
        info: Dict[str, Any] = {"target": th}
        if th in sh_norm:
            idx = sh_norm.index(th)
            src_col = 2 + idx
            info.update({
                "found": True,
                "src_col": src_col,
                "sample_value": sample[src_col] if sample and src_col < len(sample) else None,
            })
        else:
            # suggest close matches by substring
            candidates = []
            for j, h in enumerate(sh_norm):
                if h and (th in h or h in th):
                    candidates.append((j, h))
            info.update({
                "found": False,
                "candidates": "; ".join([f"{2+j}:{h}" for j, h in candidates][:5])
            })
        result_rows.append(info)

    # intersect with master headers
    covered = len(set(sh_norm) & set(master_headers_norm))

    return {
        "file": os.path.basename(fp),
        "header_row": hr,
        "data_row": di,
        "source_headers_count": len([h for h in sh_norm if h]),
        "covered_in_master": covered,
        "trace": result_rows,
    }


def main():
    municipality = os.environ.get("MUNICIPALITY_NAME")
    if not municipality:
        print("[ERROR] Set MUNICIPALITY_NAME env var.")
        sys.exit(1)

    base_dir = os.path.join(r"G:\共有ドライブ\★OD\99_商品管理\DATA\Phase3\HARV", municipality)

    files = [f for f in os.listdir(base_dir) if f.startswith("PAT") and f.endswith("_normalized.xlsx")]
    files = sorted(files)

    # classify double-normalized for awareness
    double_norm = [f for f in files if "_normalized_normalized.xlsx" in f]
    single_norm = [f for f in files if f not in double_norm]

    print(f"Municipality: {municipality}")
    print(f"Base dir: {base_dir}")
    print(f"Files (single_norm={len(single_norm)}  double_norm={len(double_norm)}):")
    for f in single_norm[:10]:
        print(f"  - {f}")
    if double_norm:
        print("Double-normalized present (should be excluded in production):")
        for f in double_norm:
            print(f"  - {f}")

    master_headers_norm = load_master_headers(base_dir)
    print(f"Master headers (all_collect) count: {len(master_headers_norm)}")

    targets_norm = [norm(x) for x in TARGET_HEADERS]
    rows = []
    for f in single_norm:
        fp = os.path.join(base_dir, f)
        info = analyze_file(fp, master_headers_norm, targets_norm)
        # console summary
        if "error" in info:
            print(f" - {info['file']}: ERROR {info['error']}")
            continue
        print(f" - {info['file']}: headers={info['source_headers_count']} covered={info['covered_in_master']}")
        for t in info["trace"]:
            if t["found"]:
                print(f"    * {t['target']} -> src_col={t['src_col']} sample={t['sample_value']}")
            else:
                print(f"    * {t['target']} -> NOT FOUND; candidates={t.get('candidates','')}")
        # for CSV
        for t in info["trace"]:
            rows.append({
                "file": info["file"],
                "target": t["target"],
                "found": t["found"],
                "src_col": t.get("src_col"),
                "sample_value": t.get("sample_value"),
                "candidates": t.get("candidates"),
            })

    out_dir = os.path.join(base_dir, "diagnostics")
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, "phase5_row_trace.csv")
    pd.DataFrame(rows).to_csv(out_path, index=False, encoding="utf-8-sig")
    print(f"\n[OUTPUT] Row trace CSV -> {out_path}")

if __name__ == "__main__":
    main()
