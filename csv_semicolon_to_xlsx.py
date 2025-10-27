# csv_semicolon_to_xlsx.py
import sys, re
from pathlib import Path

SEP = ";"  # ЖОРСТКО фіксуємо роздільник — крапка з комою

QUOTE_MAP = str.maketrans({
    "“": '"', "”": '"', "„": '"', "‟": '"', "«": '"', "»": '"',
    "‚": "'", "‘": "'", "’": "'", "‹": "'", "›": "'", "´": "'", "`": "'",
})

def norm_quotes(s: str) -> str:
    if not s: return s
    s = s.translate(QUOTE_MAP)          # типографські → ASCII
    s = s.replace('\\"', '""')          # \\" → ""
    s = s.replace("'", '"')             # одинарні → подвійні (на випадок «кривої» розмітки)
    return s

def smart_split(line: str) -> list[str]:
    out, buf, inq = [], [], False
    i, n = 0, len(line)
    while i < n:
        ch = line[i]
        if ch == '"':
            if inq and i+1 < n and line[i+1] == '"':  # "" → "
                buf.append('"'); i += 2; continue
            inq = not inq; i += 1; continue
        if ch == SEP and not inq:
            out.append("".join(buf)); buf = []; i += 1; continue
        buf.append(ch); i += 1
    out.append("".join(buf))
    return out

def is_complete(buf: str) -> bool:
    inq = False; i = 0; n = len(buf)
    while i < n:
        ch = buf[i]
        if ch == '"':
            if inq and i+1 < n and buf[i+1] == '"': i += 2; continue
            inq = not inq
        i += 1
    return not inq

def iter_records(fin):
    chunk = []
    for raw in fin:
        chunk.append(raw)
        buf = "".join(chunk)
        if is_complete(norm_quotes(buf)):
            yield buf.rstrip("\r\n")
            chunk = []
    if chunk:
        yield "".join(chunk).rstrip("\r\n")

def safe_number(s: str):
    t = "" if s is None else str(s).strip()
    if t == "": return ""
    if re.fullmatch(r"[0-9]{11,}", t): return t
    t2 = t.replace(",", ".")
    if re.fullmatch(r"[+-]?[0-9]+(\.[0-9]+)?", t2) and len(t2.replace(".","").lstrip("+-")) <= 15:
        try: return float(t2) if "." in t2 else int(t2)
        except: return t
    return t

def main(inp: Path, out: Path, encoding_hint: str | None):
    try:
        import xlsxwriter
    except Exception:
        sys.exit("Встановіть пакет:  python -m pip install xlsxwriter")

    encs = [encoding_hint] if encoding_hint else ["utf-8-sig","utf-8","cp1251","cp1252","latin1"]
    for enc in encs:
        try:
            f = open(inp, "r", encoding=enc, errors="replace", newline="")
            break
        except Exception:
            continue
    else:
        sys.exit("Не вдалося відкрити файл у жодному з відомих кодувань.")

    with f, xlsxwriter.Workbook(str(out)) as wb:
        recs = iter_records(f)
        try:
            header_raw = next(recs)
        except StopIteration:
            sys.exit("Порожній файл.")
        header = smart_split(norm_quotes(header_raw))

        ws = wb.add_worksheet("Data")
        fmt_h   = wb.add_format({"bold": True, "text_wrap": True, "valign": "top", "bg_color": "#D7E4BC", "border": 1})
        fmt_txt = wb.add_format({"num_format": "@"})
        fmt_num = wb.add_format({"border": 1, "num_format": "0.############"})
        fmt     = wb.add_format({"border": 1})

        for c, h in enumerate(header): ws.write(0, c, h, fmt_h)
        ws.freeze_panes(1, 0)
        widths = [max(5, min(50, len(str(h)))) for h in header]

        name2idx = {str(h).strip().lower(): i for i, h in enumerate(header)}
        force = {k for k in ["reg_addr_koatuu","n_reg_new"] if k in name2idx}
        r = 1; total = 0
        for raw in recs:
            row = smart_split(norm_quotes(raw))
            for c, v in enumerate(row):
                if c in [name2idx[k] for k in force]:
                    ws.write_string(r, c, "" if v is None else str(v), fmt_txt)
                else:
                    vv = safe_number(v)
                    (ws.write_number if isinstance(vv,(int,float)) else ws.write)(r, c, vv, fmt_num if isinstance(vv,(int,float)) else fmt)
                w = min(50, max(5, len(str(v))))
                if c >= len(widths): widths.extend([5]*(c+1-len(widths)))
                if w > widths[c]: widths[c] = w
            r += 1; total += 1
            if total % 200000 == 0: print(f"... {total:,}")
        for c, w in enumerate(widths): ws.set_column(c, c, w)
        try: ws.autofilter(0,0,r-1,len(widths)-1)
        except: pass
    print(f"[OK] XLSX: {out} | рядків: {total:,}")

if __name__ == "__main__":
    # Виклик: python csv_semicolon_to_xlsx.py <input.csv> <output.xlsx> [encoding]
    if len(sys.argv) < 3:
        sys.exit("Використання: python csv_semicolon_to_xlsx.py in.csv out.xlsx [cp1251|utf-8-sig|...]")
    inp = Path(sys.argv[1]); out = Path(sys.argv[2]); enc = sys.argv[3] if len(sys.argv) >= 4 else None
    main(inp, out, enc)
