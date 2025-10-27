# csv_worker.py
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import argparse, csv, os, sys, json, base64
from pathlib import Path
import pandas as pd

COMMON_ENCODINGS = ["utf-8", "utf-8-sig", "cp1251", "windows-1251", "latin-1"]
COMMON_SEPS = [",", ";", "\t", "|", ":"]
CHUNKSIZE = 200_000  # розмір чанка

def detect_encoding_and_sep(path: str):
    head = ""
    enc_used = "utf-8"
    for enc in COMMON_ENCODINGS:
        try:
            with open(path, "r", encoding=enc, errors="strict") as f:
                head = f.read(8192)
            enc_used = enc
            break
        except Exception:
            continue
    if not head:
        with open(path, "r", encoding="utf-8", errors="ignore") as f:
            head = f.read(8192)
        enc_used = "utf-8"
    try:
        dialect = csv.Sniffer().sniff(head, delimiters=COMMON_SEPS)
        sep = dialect.delimiter
    except Exception:
        counts = {d: head.count(d) for d in COMMON_SEPS}
        sep = max(counts, key=counts.get) if counts else ","
    return enc_used, sep

def build_filter_fn(spec: dict):
    mode = spec.get("mode", "6")
    raw_col = spec.get("column", "")

    def resolve_col(df: pd.DataFrame):
        if raw_col in df.columns: return raw_col
        low = {c.lower().strip(): c for c in df.columns}
        return low.get(raw_col.lower().strip())

    if mode == "1":  # текст ==
        val = spec["value"]
        def _fn(df):
            col = resolve_col(df)
            return df.iloc[0:0] if not col else df[df[col].astype(str) == val]
        return _fn
    if mode == "2":  # contains
        sub = spec["value"].lower()
        def _fn(df):
            col = resolve_col(df)
            return df.iloc[0:0] if not col else df[df[col].astype(str).str.lower().str.contains(sub, na=False)]
        return _fn
    if mode == "3":  # isin
        values = set(x.lower() for x in spec["values"])
        def _fn(df):
            col = resolve_col(df)
            return df.iloc[0:0] if not col else df[df[col].astype(str).str.lower().isin(values)]
        return _fn
    if mode == "4":  # число ==
        try: target = float(spec["value"])
        except Exception: target = None
        def _fn(df):
            if target is None: return df.iloc[0:0]
            col = resolve_col(df)
            if not col: return df.iloc[0:0]
            s = pd.to_numeric(df[col], errors="coerce")
            return df[s == target]
        return _fn
    if mode == "5":  # число у діапазоні
        try:
            vmin = float(spec["min"]); vmax = float(spec["max"])
        except Exception:
            vmin = float("-inf"); vmax = float("inf")
        def _fn(df):
            col = resolve_col(df)
            if not col: return df.iloc[0:0]
            s = pd.to_numeric(df[col], errors="coerce")
            return df[(s >= vmin) & (s <= vmax)]
        return _fn
    return lambda df: df  # без фільтра

def main():
    ap = argparse.ArgumentParser(description="Worker: обробка одного CSV → тимчасові CSV по роках")
    ap.add_argument("--input", required=True)
    ap.add_argument("--year", required=True, type=int)
    ap.add_argument("--tmp-dir", required=True)
    ap.add_argument("--filter-b64", required=True)
    args = ap.parse_args()

    spec = json.loads(base64.urlsafe_b64decode(args.filter_b64.encode("utf-8")).decode("utf-8"))
    filter_fn = build_filter_fn(spec)
    force_text_col = spec.get("column") if spec.get("mode") in ("1","2","3") and spec.get("force_text", False) else None

    enc, sep = detect_encoding_and_sep(args.input)
    print(f"[worker] file={args.input} year={args.year} enc={enc} sep='{sep}'", flush=True)

    year_dir = Path(args.tmp_dir) / str(args.year)
    year_dir.mkdir(parents=True, exist_ok=True)
    out_path = year_dir / f"part_{os.getpid()}.csv"

    header_written = False
    total_read = 0
    total_kept = 0

    try:
        for chunk in pd.read_csv(args.input, encoding=enc, sep=sep, engine="python",
                                 chunksize=CHUNKSIZE, on_bad_lines="warn", dtype=str):
            total_read += len(chunk)
            keep = filter_fn(chunk)
            if force_text_col and (force_text_col in keep.columns):
                keep[force_text_col] = keep[force_text_col].astype(str)
            if len(keep):
                keep.to_csv(out_path, mode="a", index=False, header=(not header_written), encoding="utf-8")
                header_written = True
                total_kept += len(keep)
            print(f"[worker] {Path(args.input).name}: read {total_read:,} kept {total_kept:,}", flush=True)
    except Exception as e:
        print(f"[worker][ERROR] {e}", flush=True)
        sys.exit(2)

    print(f"[worker] DONE {Path(args.input).name}: kept {total_kept:,} → {out_path}", flush=True)
    sys.exit(0)

if __name__ == "__main__":
    main()
