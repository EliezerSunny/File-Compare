"""Core logic for comparing two files. No UI code here, so it can be tested or reused."""
import difflib
import io
from datetime import date, datetime

import pandas as pd

TABULAR = {"xlsx", "xlsm", "xls", "csv", "tsv"}
TEXTUAL = {"docx", "txt", "md", "pdf"}


def ext(name: str) -> str:
    return name.rsplit(".", 1)[-1].lower() if "." in name else ""


def norm_cell(v) -> str:
    """Turn any cell value into a clean comparable string."""
    try:
        if pd.isna(v):
            return ""
    except (TypeError, ValueError):
        pass
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, float):
        return str(int(v)) if v.is_integer() else repr(round(v, 10))
    if isinstance(v, (datetime, pd.Timestamp)):
        return v.isoformat(sep=" ")
    if isinstance(v, date):
        return v.isoformat()
    return str(v).strip()


def normalize(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [norm_cell(c) or f"(blank col {i + 1})" for i, c in enumerate(df.columns)]
    # make duplicate header names unique
    seen, cols = {}, []
    for c in df.columns:
        seen[c] = seen.get(c, 0) + 1
        cols.append(c if seen[c] == 1 else f"{c}.{seen[c]}")
    df.columns = cols
    return df.astype(object).apply(lambda col: col.map(norm_cell))


# ----------------------------------------------------------------- loading
def _read_csv(data: bytes) -> pd.DataFrame:
    last_err = None
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            return pd.read_csv(io.BytesIO(data), sep=None, engine="python",
                               dtype=object, encoding=enc)
        except UnicodeDecodeError as err:
            last_err = err
    raise last_err


def _decode(data: bytes) -> str:
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode("latin-1", errors="replace")


def _read_unknown_excel(data: bytes) -> dict:
    """File has an Excel extension but pandas can't read it as Excel.
    Detect what it really is (renamed CSV, HTML table export...)."""
    if data[:2] == b"PK" or data[:4] == b"\xd0\xcf\x11\xe0":
        raise ValueError(
            "This looks like an Excel file but it cannot be opened. It may be "
            "password-protected or damaged. Remove the password (or re-save it from "
            "Excel as .xlsx) and try again.")
    if b"\x00" in data[:4096]:  # binary data, not a text export
        raise ValueError("The file is not a valid Excel file (unrecognised binary "
                         "content). Open it in Excel and save it again as .xlsx.")
    text = _decode(data)
    if text.lstrip()[:1] == "<":  # HTML table exported with an .xls name
        try:
            from collections import defaultdict
            # converters keep every cell as text (no lost leading zeros / long IDs)
            tables = pd.read_html(io.StringIO(text), flavor=None,
                                  converters=defaultdict(lambda: str))
        except Exception as err:
            raise ValueError("The file is a web page, not a real Excel file, and no "
                             f"table could be read from it ({err}).")
        return {f"Table {i + 1}": t.astype(object) for i, t in enumerate(tables)}
    try:
        return {"Data": _read_csv(data)}  # text file saved with an Excel extension
    except Exception:
        raise ValueError("The file is not a valid Excel file. Open it in Excel and "
                         "save it again as .xlsx, then upload it.")


def load_tables(name: str, data: bytes) -> dict:
    """Return {sheet_name: DataFrame} for Excel/CSV/TSV files."""
    e = ext(name)
    if not data:
        raise ValueError(f"'{name}' is empty (0 bytes). Re-upload the file.")
    if e in ("xlsx", "xlsm", "xls"):
        try:
            sheets = pd.read_excel(io.BytesIO(data), sheet_name=None, dtype=object)
        except Exception:
            sheets = _read_unknown_excel(data)
    elif e in ("csv", "tsv"):
        sheets = {"Data": _read_csv(data)}
    else:
        raise ValueError(f"Not a table file: {name}")
    if not sheets:
        raise ValueError(f"No readable sheets found in '{name}'.")
    return {str(k): normalize(v) for k, v in sheets.items()}


def load_text(name: str, data: bytes) -> list:
    """Return a list of text lines for Word/TXT/MD/PDF files."""
    e = ext(name)
    if e == "docx":
        from docx import Document
        from docx.table import Table
        from docx.text.paragraph import Paragraph

        doc = Document(io.BytesIO(data))
        lines = []
        for child in doc.element.body.iterchildren():
            if child.tag.endswith("}p"):
                lines.append(Paragraph(child, doc).text.strip())
            elif child.tag.endswith("}tbl"):
                for row in Table(child, doc).rows:
                    lines.append(" | ".join(c.text.strip().replace("\n", " ") for c in row.cells))
        return [l for l in lines if l]
    if e == "pdf":
        from pypdf import PdfReader

        reader = PdfReader(io.BytesIO(data))
        text = "\n".join((p.extract_text() or "") for p in reader.pages)
        return [l.strip() for l in text.splitlines() if l.strip()]
    if e in ("txt", "md"):
        for enc in ("utf-8-sig", "cp1252", "latin-1"):
            try:
                return [l.rstrip() for l in data.decode(enc).splitlines()]
            except UnicodeDecodeError:
                continue
    raise ValueError(f"Not a text/document file: {name}")


# --------------------------------------------------------------- comparing
def compare_tables(a: pd.DataFrame, b: pd.DataFrame, key_cols=None) -> dict:
    """Compare two normalized tables.

    key_cols: list of column names that identify a row (e.g. ["ID"]).
              If empty/None, rows are compared by position (row 1 vs row 1).
    """
    key_cols = list(key_cols or [])
    common = [c for c in a.columns if c in b.columns]
    only_a = [c for c in a.columns if c not in b.columns]
    only_b = [c for c in b.columns if c not in a.columns]
    for k in key_cols:
        if k not in common:
            raise ValueError(f"Key column '{k}' must exist in both files.")
    value_cols = [c for c in common if c not in key_cols]

    def keyed(df):
        df = df.copy()
        if key_cols:
            base = [" | ".join(map(str, row)) for row in df[key_cols].itertuples(index=False)]
        else:
            base = [str(i + 1) for i in range(len(df))]
        totals, seen, keys = {}, {}, []
        for k in base:
            totals[k] = totals.get(k, 0) + 1
        for k in base:  # duplicates get #1, #2... so no row is silently lost
            seen[k] = seen.get(k, 0) + 1
            keys.append(k if totals[k] == 1 else f"{k} (#{seen[k]})")
        df.index = pd.Index(keys, dtype=object)
        return df

    ka, kb = keyed(a), keyed(b)
    dup_a = sum("(#" in k for k in ka.index) if key_cols else 0
    dup_b = sum("(#" in k for k in kb.index) if key_cols else 0

    removed_keys = [k for k in ka.index if k not in kb.index]
    added_keys = [k for k in kb.index if k not in ka.index]
    both = [k for k in ka.index if k in kb.index]

    changes = []
    for k in both:
        for c in value_cols:
            old, new = ka.at[k, c], kb.at[k, c]
            if old != new:
                changes.append({"Row/Key": k, "Column": c, "File A": old, "File B": new})

    label = "Row" if not key_cols else "Key"
    changed_df = pd.DataFrame(changes, columns=["Row/Key", "Column", "File A", "File B"])
    removed_df = ka.loc[removed_keys].reset_index(names=label)
    added_df = kb.loc[added_keys].reset_index(names=label)

    return {
        "changed": changed_df,
        "removed": removed_df,
        "added": added_df,
        "only_in_a_cols": only_a,
        "only_in_b_cols": only_b,
        "rows_a": len(a),
        "rows_b": len(b),
        "rows_changed": changed_df["Row/Key"].nunique() if len(changed_df) else 0,
        "dup_a": dup_a,
        "dup_b": dup_b,
        "identical": not (len(changed_df) or len(removed_df) or len(added_df)
                          or only_a or only_b),
    }


def compare_text(a_lines: list, b_lines: list) -> dict:
    ratio = difflib.SequenceMatcher(None, a_lines, b_lines, autojunk=False).ratio()
    html = difflib.HtmlDiff(wrapcolumn=70).make_table(
        a_lines, b_lines, "File A", "File B", context=True, numlines=2)
    added = removed = 0
    for l in difflib.ndiff(a_lines, b_lines):
        if l.startswith("+ "):
            added += 1
        elif l.startswith("- "):
            removed += 1
    return {"similarity": ratio, "html": html, "added": added, "removed": removed,
            "identical": a_lines == b_lines}


def build_report(result: dict) -> bytes:
    """Excel report of a table comparison."""
    out = io.BytesIO()
    summary = pd.DataFrame({
        "Item": ["Rows in A", "Rows in B", "Rows added in B", "Rows removed from B",
                 "Rows with changed cells", "Changed cells",
                 "Columns only in A", "Columns only in B"],
        "Value": [result["rows_a"], result["rows_b"], len(result["added"]),
                  len(result["removed"]), result["rows_changed"], len(result["changed"]),
                  ", ".join(result["only_in_a_cols"]) or "-",
                  ", ".join(result["only_in_b_cols"]) or "-"],
    })
    with pd.ExcelWriter(out, engine="openpyxl") as xw:
        summary.to_excel(xw, sheet_name="Summary", index=False)
        result["changed"].to_excel(xw, sheet_name="Changed cells", index=False)
        result["added"].to_excel(xw, sheet_name="Added in B", index=False)
        result["removed"].to_excel(xw, sheet_name="Removed from B", index=False)
    return out.getvalue()


def build_text_reports(a_lines, b_lines, html_table, name_a="File A", name_b="File B"):
    """Return (standalone HTML bytes, plain-text unified diff bytes)."""
    page = ("<!doctype html><html><head><meta charset='utf-8'>"
            "<meta name='viewport' content='width=device-width, initial-scale=1'>"
            "<title>Comparison</title><style>body{font-family:sans-serif;font-size:13px}"
            "table.diff{width:100%}.diff_add{background:#aaffaa}.diff_chg{background:#ffff77}"
            ".diff_sub{background:#ffaaaa}td{padding:2px 4px}</style></head><body>"
            "<h2>Comparison: " + name_a + " vs " + name_b + "</h2>"
            "<p>Green = added, red = removed, yellow = changed.</p>"
            + html_table + "</body></html>")
    txt = "\n".join(difflib.unified_diff(a_lines, b_lines, name_a, name_b, lineterm=""))
    return page.encode("utf-8"), (txt or "No differences.").encode("utf-8")
