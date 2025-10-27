#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import argparse, base64, json, os, re, shutil, subprocess, sys, tempfile
from pathlib import Path
from typing import Dict, List, Optional
import pandas as pd

MAX_EXCEL_ROWS = 1_048_576

EXCEL_ENGINE = None
try:
    import xlsxwriter  # noqa
    EXCEL_ENGINE = "xlsxwriter"
except Exception:
    try:
        import openpyxl  # noqa
        EXCEL_ENGINE = "openpyxl"
    except Exception:
        EXCEL_ENGINE = None

def print_about():
    print("""
UI режими:
  --ui wt       : одне WT-вікно з панелями.
  --ui wt-win   : окремі WT-вікна (надійний варіант).
  --ui consoles : окремі класичні консолі.

Роки: беруться з назви файла (19/20xx). Якщо року в назві нема — скрипт запитає YYYY.
""")

def prompt_filter_params() -> Dict:
    print("\n=== Налаштування фільтра ===")
    col = input("Назва колонки для фільтра (як у заголовку CSV): ").strip()
    print("Тип значення:\n 1) Текст: точний збіг (==)\n 2) Текст: містить (contains)\n 3) Текст: перелік (isin, через кому)\n 4) Число: ==\n 5) Число: діапазон [min; max]\n 6) Без фільтра")
    mode = (input("Оберіть 1/2/3/4/5/6: ").strip() or "6")
    params = {"column": col, "mode": mode, "force_text": False}
    if mode == "1":
        v = input("Значення (текст): ").strip()
        params["value"] = v
        params["force_text"] = (input("Форматувати колонку як ТЕКСТ? [y/N]: ").strip().lower() == "y")
    elif mode == "2":
        v = input("Підрядок (текст): ").strip()
        params["value"] = v
        params["force_text"] = (input("Форматувати колонку як ТЕКСТ? [y/N]: ").strip().lower() == "y")
    elif mode == "3":
        v = input("Перелік значень через кому: ").strip()
        params["values"] = [x.strip() for x in v.split(",") if x.strip()]
        params["force_text"] = (input("Форматувати колонку як ТЕКСТ? [y/N]: ").strip().lower() == "y")
    elif mode == "4":
        params["value"] = input("Числове значення (==): ").strip()
    elif mode == "5":
        params["min"] = input("Мінімум: ").strip()
        params["max"] = input("Максимум: ").strip()
    else:
        print("Фільтр вимкнено.")
    return params

def infer_year_from_filename(p: str) -> Optional[int]:
    m = re.search(r"(19|20)\d{2}", Path(p).stem)
    if m:
        y = int(m.group(0))
        if 1900 <= y <= 2100:
            return y
    return None

def resolve_years_for_files(files: List[str]) -> Dict[str, int]:
    print("\n=== Визначення року для кожного файла ===")
    years: Dict[str, int] = {}
    for p in files:
        y = infer_year_from_filename(p)
        if y is None:
            s = input(f"Рік для {Path(p).name} (YYYY): ").strip()
            while not re.fullmatch(r"(19|20)\d{2}", s):
                s = input("  Коректний рік (YYYY): ").strip()
            y = int(s)
        years[p] = y
        print(f"  {Path(p).name} → {y}")
    return years

def open_windows_wt_win(files, years, worker, tmpdir, filter_b64):
    import shutil as _sh
    wt = _sh.which("wt.exe")
    if not wt:
        return None
    py = sys.executable or "python"
    procs = []
    for p in files:
        procs.append(subprocess.Popen([
            wt, "new-window",
            "cmd", "/k",
            py, worker,
            "--input", p,
            "--year", str(years[p]),
            "--tmp-dir", tmpdir,
            "--filter-b64", filter_b64
        ]))
    return procs

def open_windows_wt(files, years, worker, tmpdir, filter_b64):
    import shutil as _sh
    wt = _sh.which("wt.exe")
    if not wt:
        return None
    py = sys.executable or "python"
    args = [
        wt, "-w", "0",
        "new-tab", "cmd", "/k",
        py, worker,
        "--input", files[0],
        "--year", str(years[files[0]]),
        "--tmp-dir", tmpdir,
        "--filter-b64", filter_b64
    ]
    split_h = True
    for p in files[1:]:
        args += [
            ";", "split-pane", ("-H" if split_h else "-V"),
            "cmd", "/k",
            py, worker,
            "--input", p,
            "--year", str(years[p]),
            "--tmp-dir", tmpdir,
            "--filter-b64", filter_b64
        ]
        split_h = not split_h
    subprocess.Popen(args)
    return True

def open_windows_consoles(files, years, worker, tmpdir, filter_b64):
    procs = []
    py = sys.executable or "python"
    CREATE_NEW_CONSOLE = 0x00000010
    for p in files:
        procs.append(subprocess.Popen(
            [py, worker, "--input", p, "--year", str(years[p]), "--tmp-dir", tmpdir, "--filter-b64", filter_b64],
            creationflags=CREATE_NEW_CONSOLE))
    return procs

def write_excel_from_temp(tmpdir: str, out_xlsx: str, force_text_col: Optional[str]):
    if not EXCEL_ENGINE:
        raise RuntimeError("Встановіть xlsxwriter або openpyxl.")
    years = sorted([int(p.name) for p in Path(tmpdir).iterdir() if p.is_dir() and p.name.isdigit()])
    with pd.ExcelWriter(out_xlsx, engine=EXCEL_ENGINE) as writer:
        for y in years:
            ydir = Path(tmpdir) / str(y)
            parts = sorted(ydir.glob("*.csv"))
            if not parts: continue
            sheet_idx, rows_on_sheet, header_written = 1, 0, False
            cols = list(pd.read_csv(parts[0], nrows=0, dtype=str).columns)
            def sheet_name(i): return f"{y}" if i == 1 else f"{y} ({i})"
            for part in parts:
                for chunk in pd.read_csv(part, chunksize=250_000, dtype=str):
                    available = MAX_EXCEL_ROWS - 1 - rows_on_sheet
                    if available <= 0:
                        sheet_idx, rows_on_sheet, header_written = sheet_idx+1, 0, False
                    if len(chunk) > available > 0:
                        write_df, rest = chunk.iloc[:available], chunk.iloc[available:]
                    else:
                        write_df, rest = chunk, None
                    sh = sheet_name(sheet_idx)
                    write_df = write_df.reindex(columns=cols)
                    write_df.to_excel(writer, sheet_name=sh, index=False,
                                      header=(not header_written),
                                      startrow=(0 if not header_written else rows_on_sheet + 1))
                    ws = writer.sheets[sh]
                    if not header_written:
                        try:
                            if EXCEL_ENGINE == "xlsxwriter":
                                ws.freeze_panes(1, 0); ws.autofilter(0, 0, 0, len(cols)-1)
                                if force_text_col and force_text_col in cols:
                                    j = cols.index(force_text_col)
                                    ws.set_column(j, j, None, writer.book.add_format({"num_format": "@"}))
                            else:
                                ws.freeze_panes = "A2"
                        except Exception: pass
                        header_written = True
                    rows_on_sheet += len(write_df)
                    if rest is not None and len(rest) > 0:
                        sheet_idx, rows_on_sheet, header_written = sheet_idx+1, 0, False
                        sh = sheet_name(sheet_idx)
                        rest = rest.reindex(columns=cols)
                        rest.to_excel(writer, sheet_name=sh, index=False, header=True, startrow=0)
                        ws2 = writer.sheets[sh]
                        try:
                            if EXCEL_ENGINE == "xlsxwriter":
                                ws2.freeze_panes(1, 0); ws2.autofilter(0, 0, 0, len(cols)-1)
                                if force_text_col and force_text_col in cols:
                                    j = cols.index(force_text_col)
                                    ws2.set_column(j, j, None, writer.book.add_format({"num_format": "@"}))
                            else:
                                ws2.freeze_panes = "A2"
                        except Exception: pass
                        rows_on_sheet = len(rest); header_written = True
    print(f"[orchestrator] XLSX готово: {out_xlsx}")

def main():
    ap = argparse.ArgumentParser(description="Паралельна обробка CSV у вікнах/панелях + збірка у XLSX")
    ap.add_argument("-o", "--output", required=True, help="Вихідний XLSX")
    ap.add_argument("--ui", choices=["wt", "wt-win", "consoles"], default="wt",
                    help="wt=панелі в одному WT; wt-win=окремі WT-вікна; consoles=окремі консолі")
    ap.add_argument("--about", action="store_true", help="Пояснення та вихід")
    ap.add_argument("files", nargs="*", help="Шляхи до CSV")
    args = ap.parse_args()

    if args.about: print_about(); return

    files = args.files or [s for s in re.split(r"\s+", input("Шляхи до CSV (через пробіл): ").strip()) if s]
    filt = prompt_filter_params()
    filter_b64 = base64.urlsafe_b64encode(json.dumps(filt, ensure_ascii=False).encode("utf-8")).decode("utf-8")
    force_text_col = filt.get("column") if filt.get("mode") in ("1","2","3") and filt.get("force_text", False) else None
    years = resolve_years_for_files(files)

    tmpdir = tempfile.mkdtemp(prefix="csv_parallel_")
    print(f"[orchestrator] temp-dir: {tmpdir}")

    worker = str(Path(__file__).with_name("csv_worker.py"))
    if not Path(worker).exists():
        print(f"[ERROR] Не знайдено {worker} поруч із оркестратором."); sys.exit(2)

    print("[orchestrator] Старт воркерів...")
    procs = None
    if args.ui == "wt":
        ok = open_windows_wt(files, years, worker, tmpdir, filter_b64)
        if not ok: print("[orchestrator] wt.exe не знайдено — режим consoles."); procs = open_windows_consoles(files, years, worker, tmpdir, filter_b64)
    elif args.ui == "wt-win":
        procs = open_windows_wt_win(files, years, worker, tmpdir, filter_b64)
        if procs is None: print("[orchestrator] wt.exe не знайдено — режим consoles."); procs = open_windows_consoles(files, years, worker, tmpdir, filter_b64)
    else:
        procs = open_windows_consoles(files, years, worker, tmpdir, filter_b64)

    if isinstance(procs, list) and procs:
        for i, p in enumerate(procs, start=1):
            p.wait(); print(f"[orchestrator] worker {i} завершився (pid={p.pid}).")
    else:
        input("[orchestrator] Натисніть Enter, коли всі панелі/вікна завершаться...")

    print("[orchestrator] Збірка XLSX...")
    write_excel_from_temp(tmpdir, args.output, force_text_col=force_text_col)

    try: shutil.rmtree(tmpdir)
    except Exception: print(f"[orchestrator] Тимчасові файли лишилися тут: {tmpdir}")

if __name__ == "__main__":
    main()
