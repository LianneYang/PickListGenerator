"""
WebApp_PickListGenerator.py
---------------------------
Streamlit single-file app: pick-list generator.

Upload a stock file + a demand file, get back a formatted Excel workbook.

The app auto-detects the demand structure:

  • No duplicate materials  →  appends Pick Location / Pick QTY / Cumulative /
    Batch to the flat demand sheet (renamed to 'WO').
  • Duplicates found        →  builds WO, Summary PN, Pick List, and DN.

Run locally:   streamlit run WebApp_PickListGenerator.py
"""

from io import BytesIO

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ==================================================================
# CONFIG
# ==================================================================

SHEET_SUMMARY = "Summary PN"
SHEET_OUT = "Pick List"
SHEET_DN = "DN"
SHEET_FLAT = "WO"
SHEET_PROBLEMS = "Problems"

# 1-based positions in the user's mind; converted to 0-based when used.
SHEET_SUMMARY_POSITION = 2   # → index 1
SHEET_PICKLIST_POSITION = 3  # → index 2
SHEET_DN_POSITION = 4        # → index 3

# ------------------------------------------------------------------
#  Column width settings
#  • The "*_WIDTH" presets act as a MINIMUM floor for the matching
#    column — the column can still grow wider if a cell contains a
#    longer value, so no content gets hidden.
#  • AUTOFIT_MAX_WIDTH caps how far a column may expand (keeps very
#    long Descriptions from making the sheet unusably wide).
#  • AUTOFIT_MIN_WIDTH is the absolute floor for any column without
#    an explicit preset.
# ------------------------------------------------------------------
SUM_OF_QUANTITY_WIDTH = 16.78
SCHED_AGREEMENT_WIDTH = 17
PICK_LOCATION_WIDTH = 16
PICK_QTY_WIDTH = 12
CUMULATIVE_WIDTH = 12
BATCH_WIDTH = 14
UNRESTRICTED_WIDTH = 20
COMMENTS_WIDTH = 30

AUTOFIT_MIN_WIDTH = 10
AUTOFIT_MAX_WIDTH = 60
AUTOFIT_PADDING = 3

# ------------------------------------------------------------------
#  WO sheet columns.
#
#  "Unrestricted-Use Stock" and "Comments" are OPTIONAL — if the
#  demand file doesn't contain them, blank values are written so the
#  WO sheet always shows the same column layout.
# ------------------------------------------------------------------
FLAT_COLUMNS = [
    "Schedule line date",
    "Discharge Location",
    "Material",
    "Description",
    "Quantity",
    "Unrestricted-Use Stock",
    "Comments",
    "Customer Material Number",
    "Sched.agreemnt",
]

PICK_COLUMNS = ["Pick Location", "Pick QTY", "Cumulative", "Batch"]

DN_COLUMNS = [
    "Schedule line date",
    "Discharge Location",
    "Material",
    "Description",
    "Quantity",
    "Customer Material Number",
    "Sched.agreemnt",
    "Remarks",
    "DN",
]

STOCK_MATERIAL = "Material"
STOCK_BIN = "Storage Bin"
STOCK_BATCH = "Batch"
STOCK_AVAILABLE = "Available stock"

# Columns written as integers (drop the ".0")
INT_COLUMNS = {"Quantity", "Pick QTY", "Cumulative",
               "Sum of Quantity", "Unrestricted-Use Stock"}


# ==================================================================
# STOCK
# ==================================================================

def load_stock_from_frame(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    missing = [c for c in (STOCK_MATERIAL, STOCK_BIN, STOCK_BATCH, STOCK_AVAILABLE)
               if c not in df.columns]
    if missing:
        raise KeyError(f"Missing stock columns: {missing} — found: {list(df.columns)}")

    df = df[[STOCK_MATERIAL, STOCK_BIN, STOCK_BATCH, STOCK_AVAILABLE]].rename(
        columns={STOCK_MATERIAL: "material",
                 STOCK_BIN: "bin",
                 STOCK_BATCH: "batch",
                 STOCK_AVAILABLE: "available"})

    df["material"] = df["material"].astype(str).str.strip().str.lstrip("0")
    df["bin"] = df["bin"].astype(str).str.strip()
    df["batch"] = df["batch"].astype(str).str.strip()
    df["available"] = pd.to_numeric(df["available"], errors="coerce").fillna(0)

    df = df[(df["bin"] != "") & (df["available"] > 0)]
    # Batches are consumed in FIFO order. Within a batch the bin order
    # here only serves as a deterministic tie-breaker — the actual
    # selection is done by `allocate` using the closest-fit rules.
    df = df.sort_values(["batch", "bin"], kind="stable").reset_index(drop=True)
    return df


def load_stock_from_bytes(data: bytes) -> pd.DataFrame:
    df = pd.read_excel(BytesIO(data), sheet_name=0, dtype=str)
    return load_stock_from_frame(df)


# ==================================================================
# DETECT DEMAND STRUCTURE
# ==================================================================

def detect_demand_from_bytes(data: bytes):
    wb = load_workbook(BytesIO(data), read_only=True)
    sheetnames = wb.sheetnames
    wb.close()

    base_cols = {"Schedule line date", "Discharge Location", "Material",
                 "Description", "Quantity"}

    flat_sheet = None
    for sn in sheetnames:
        try:
            head = pd.read_excel(BytesIO(data), sheet_name=sn, dtype=str, nrows=0)
            head.columns = [str(c).strip() for c in head.columns]
            if base_cols.issubset(set(head.columns)):
                flat_sheet = sn
                break
        except Exception:
            continue

    if flat_sheet is None:
        raise KeyError("No flat sheet with Material/Description/Quantity found.")

    df = pd.read_excel(BytesIO(data), sheet_name=flat_sheet, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]

    mats = (df["Material"].dropna()
            .astype(str).str.strip().str.lstrip("0"))
    mats = mats[mats != ""]
    has_dupes = mats.duplicated().any()

    return {"sheet": flat_sheet, "duplicates": bool(has_dupes)}


# ==================================================================
# DEMAND LOADERS  (duplicate mode)
# ==================================================================

def _read_flat_sheet(data: bytes, sheet: str) -> pd.DataFrame:
    df = pd.read_excel(BytesIO(data), sheet_name=sheet, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    required = {"Material", "Description", "Quantity"}
    missing = required - set(df.columns)
    if missing:
        raise KeyError(f"Missing columns in demand sheet '{sheet}': {sorted(missing)}")
    return df


def load_demand_from_bytes(data: bytes) -> pd.DataFrame:
    buf = BytesIO(data)
    wb = load_workbook(buf, read_only=True)
    sheetnames = wb.sheetnames
    wb.close()

    if SHEET_SUMMARY in sheetnames:
        raw = pd.read_excel(BytesIO(data), sheet_name=SHEET_SUMMARY,
                            dtype=str, header=None)
        header_row = None
        for i in range(min(30, len(raw))):
            row_vals = [str(v).strip().lower() for v in raw.iloc[i].tolist()
                        if pd.notna(v)]
            if "material" in " | ".join(row_vals) and \
                    ("quantity" in " | ".join(row_vals) or "qty" in " | ".join(row_vals)):
                header_row = i
                break
        if header_row is None:
            raise KeyError(f"No header row found in '{SHEET_SUMMARY}'.")

        df = pd.read_excel(BytesIO(data), sheet_name=SHEET_SUMMARY,
                           dtype=str, header=header_row)
        df.columns = [str(c).strip() for c in df.columns]

        def find_col(keywords):
            for c in df.columns:
                cl = c.lower()
                if any(k in cl for k in keywords):
                    return c
            return None

        col_mat = find_col(["material"])
        col_desc = find_col(["description", "row label", "desc"])
        col_qty = find_col(["sum of quantity", "quantity", "qty"])

        df = df.rename(columns={col_mat: "material",
                                col_desc: "description",
                                col_qty: "quantity"})
        df = df[["material", "description", "quantity"]]
        df = df[df["material"].notna()]
        df = df[~df["material"].astype(str).str.strip()
        .str.lower().str.startswith("grand")]
        df = df[~df["material"].astype(str).str.strip().str.startswith("=")]
    else:
        chosen = None
        for sn in sheetnames:
            try:
                df = _read_flat_sheet(data, sn)
                chosen = sn
                break
            except KeyError:
                continue
        if chosen is None:
            raise KeyError("No flat sheet with Material/Description/Quantity found.")
        df = df[["Material", "Description", "Quantity"]].rename(
            columns={"Material": "material",
                     "Description": "description",
                     "Quantity": "quantity"})

    df["material"] = df["material"].astype(str).str.strip().str.lstrip("0")
    df["description"] = df["description"].astype(str).str.strip()
    df["quantity"] = pd.to_numeric(df["quantity"], errors="coerce").fillna(0)
    df = df[df["quantity"] > 0].reset_index(drop=True)
    return df


def load_dn_source_from_bytes(data: bytes) -> pd.DataFrame:
    """
    Loads the DN source and preserves the ORIGINAL ROW ORDER of the
    demand file. We do not re-sort by (Location, Sched.agreemnt, Material)
    anymore — we add a hidden `_row_order` column capturing the position
    in the file, so build_dn can reproduce the same sequence as the WO
    sheet for each Discharge Location.
    """
    buf = BytesIO(data)
    wb = load_workbook(buf, read_only=True)
    sheetnames = wb.sheetnames
    wb.close()

    base_cols = {"Schedule line date", "Discharge Location", "Material",
                 "Description", "Quantity", "Customer Material Number",
                 "Sched.agreemnt"}

    df = None
    for sn in sheetnames:
        try:
            candidate = pd.read_excel(BytesIO(data), sheet_name=sn, dtype=str)
            candidate.columns = [str(c).strip() for c in candidate.columns]
            if base_cols.issubset(set(candidate.columns)):
                df = candidate
                break
        except Exception:
            continue

    if df is None:
        raise KeyError("No sheet found with the DN base columns.")

    # Capture original row order BEFORE any processing
    df = df.reset_index(drop=False).rename(columns={"index": "_row_order"})

    for col in DN_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    keep_cols = DN_COLUMNS + ["_row_order"]
    df = df[keep_cols].copy()
    df = df[df["Material"].notna()]
    df["Discharge Location"] = df["Discharge Location"].astype(str).str.strip()
    df["Material"] = df["Material"].astype(str).str.strip().str.lstrip("0")
    df["Description"] = df["Description"].astype(str).str.strip()
    df["Sched.agreemnt"] = df["Sched.agreemnt"].astype(str).str.strip()
    df["Remarks"] = df["Remarks"].fillna("").astype(str)
    df["DN"] = df["DN"].fillna("").astype(str)
    df["Quantity"] = pd.to_numeric(df["Quantity"],
                                   errors="coerce").fillna(0)

    parsed = pd.to_datetime(df["Schedule line date"], errors="coerce")
    df["Schedule line date"] = parsed.dt.strftime("%d/%m/%Y").where(
        parsed.notna(), df["Schedule line date"])

    df = (df
    .groupby(["Discharge Location", "Sched.agreemnt", "Material"],
             as_index=False)
    .agg({
        "Schedule line date": "first",
        "Description": "first",
        "Quantity": "sum",
        "Customer Material Number": "first",
        "Remarks": "first",
        "DN": "first",
        "_row_order": "min",
    }))

    # Order strictly by the earliest appearance in the source file
    df = df.sort_values(["_row_order", "Discharge Location", "Material"],
                        kind="stable").reset_index(drop=True)
    df = df.drop(columns=["_row_order"])
    return df


def load_flat_sheet_keep_columns_from_bytes(data: bytes):
    buf = BytesIO(data)
    wb = load_workbook(buf, read_only=True)
    sheetnames = wb.sheetnames
    wb.close()

    required = {"Material", "Description", "Quantity"}
    chosen, df = None, None
    for sn in sheetnames:
        try:
            candidate = pd.read_excel(BytesIO(data), sheet_name=sn, dtype=str)
            candidate.columns = [str(c).strip() for c in candidate.columns]
            if required.issubset(set(candidate.columns)):
                df = candidate
                chosen = sn
                break
        except Exception:
            continue

    if df is None:
        raise KeyError("No flat sheet with Material/Description/Quantity found.")

    # Ensure every wanted WO column exists, then keep only those columns.
    # Optional ones ("Unrestricted-Use Stock", "Comments") are filled
    # with "" when the demand file doesn't provide them.
    for col in FLAT_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    df = df[FLAT_COLUMNS].copy()

    # Text columns: strip whitespace only.
    for col in ("Discharge Location", "Description", "Comments",
                "Unrestricted-Use Stock", "Customer Material Number",
                "Sched.agreemnt"):
        df[col] = df[col].fillna("").astype(str).str.strip()

    df["Material"] = df["Material"].astype(str).str.strip().str.lstrip("0")
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce").fillna(0)

    parsed = pd.to_datetime(df["Schedule line date"], errors="coerce")
    df["Schedule line date"] = parsed.dt.strftime("%d/%m/%Y").where(
        parsed.notna(), df["Schedule line date"].astype(str))

    df = df[df["Material"] != ""].reset_index(drop=True)
    return df, chosen


# ==================================================================
# SUMMARY / ALLOCATION / BUILDERS
# ==================================================================

def summarize(demand: pd.DataFrame) -> pd.DataFrame:
    return (demand
            .groupby("material", as_index=False)
            .agg({"description": "first", "quantity": "sum"})
            .sort_values("material")
            .reset_index(drop=True))


def allocate(stock: pd.DataFrame, material: str, need: float):
    """
    Allocate stock for a material.

    Rules
    -----
    • Batches are consumed in FIFO order (as sorted in `load_stock`).
    • Within a batch:
        - The FIRST pick chooses the bin whose available qty is
          closest to the total demand.
        - Every FOLLOWING pick chooses the bin that makes the
          running cumulative closest to the total demand.
    • If total available < demand, a final {"short": qty} entry is
      appended so callers can flag the shortfall.
    """
    rows = stock[stock["material"] == material]
    if rows.empty:
        return [{"short": need}]

    need = float(need)
    result = []
    remaining = need

    # Batch FIFO order — preserve the order established in load_stock
    batches = list(dict.fromkeys(rows["batch"].tolist()))

    for batch in batches:
        if remaining <= 0:
            break

        batch_rows = rows[rows["batch"] == batch]
        pool = list(zip(batch_rows["bin"].tolist(),
                        batch_rows["available"].astype(float).tolist()))

        while remaining > 0 and pool:
            cumulative_so_far = need - remaining

            if cumulative_so_far == 0:
                # First pick: bin whose qty is closest to the demand
                best_idx, best_key = None, None
                for i, (_, avail) in enumerate(pool):
                    key = (abs(avail - need), avail, i)
                    if best_key is None or key < best_key:
                        best_key, best_idx = key, i
            else:
                # Following picks: bin whose resulting cumulative is
                # closest to the demand
                best_idx, best_key = None, None
                for i, (_, avail) in enumerate(pool):
                    take = min(avail, remaining)
                    new_cum = cumulative_so_far + take
                    key = (abs(new_cum - need), take, i)
                    if best_key is None or key < best_key:
                        best_key, best_idx = key, i

            bin_code, avail = pool.pop(best_idx)
            take = min(avail, remaining)
            result.append({"bin": bin_code, "batch": batch, "qty": take})
            remaining -= take

        if remaining <= 0:
            break

    if remaining > 0:
        result.append({"short": remaining})

    return result


def build_picklist(stock: pd.DataFrame, summary: pd.DataFrame):
    """
    Builds the Pick List.
    Rule:
      • Normal pick rows carry bin / batch / qty / cumulative.
      • If total available < demand → add a 'Miss Stock' row below
        the material's pick rows, showing the missed qty in Pick QTY.
      • If no stock at all → single 'Miss Stock' row for the material.
    """
    out_rows, problems = [], []

    for idx, d in summary.iterrows():
        mat = d["material"]
        desc = d["description"]
        need = float(d["quantity"])

        if idx > 0:
            out_rows.append({c: "" for c in [
                "Material", "Description", "Sum of Quantity",
                "Pick Location", "Pick QTY", "Cumulative", "Batch"]})

        # Per-material safety net: one bad row can't kill the whole file.
        try:
            alloc = allocate(stock, mat, need)
        except Exception as e:
            problems.append(f"{mat} ({desc}) — internal error: {e}")
            alloc = [{"short": need}]

        # ----- No usable stock at all -----
        if not alloc or (len(alloc) == 1 and "short" in alloc[0]):
            short_qty = alloc[0]["short"] if alloc else need
            out_rows.append({
                "Material": mat,
                "Description": desc,
                "Sum of Quantity": need,
                "Pick Location": "Miss Stock",
                "Pick QTY": short_qty,
                "Cumulative": "",
                "Batch": "",
            })
            problems.append(f"{mat} ({desc}) — NO STOCK, short by {short_qty:.0f}")
            continue

        # ----- Normal picks + possible Miss Stock row -----
        cumulative, line_count, short_qty = 0.0, 0, 0

        for item in alloc:
            if "short" in item:
                short_qty = item["short"]
                continue
            cumulative += item["qty"]
            line_count += 1
            out_rows.append({
                "Material": mat if line_count == 1 else "",
                "Description": desc if line_count == 1 else "",
                "Sum of Quantity": need if line_count == 1 else "",
                "Pick Location": item["bin"],
                "Pick QTY": item["qty"],
                "Cumulative": cumulative,
                "Batch": item["batch"],
            })

        if short_qty:
            out_rows.append({
                "Material": "", "Description": "", "Sum of Quantity": "",
                "Pick Location": "Miss Stock", "Pick QTY": short_qty,
                "Cumulative": "", "Batch": "",
            })
            problems.append(f"{mat} ({desc}) — SHORT by {short_qty:.0f}")

    return pd.DataFrame(out_rows, columns=[
        "Material", "Description", "Sum of Quantity",
        "Pick Location", "Pick QTY", "Cumulative", "Batch"]), problems


def build_dn(dn_source: pd.DataFrame) -> pd.DataFrame:
    """
    Builds the DN sheet.
    Rows are grouped by Discharge Location, and within each location
    they appear in the SAME ORDER as the WO sheet (i.e. the original
    file order). A blank row separates each Discharge Location block.
    """
    if dn_source.empty:
        return dn_source

    df = dn_source.copy()

    # Build a location-order index based on first appearance
    loc_first_seen = {}
    for i, loc in enumerate(df["Discharge Location"]):
        if loc not in loc_first_seen:
            loc_first_seen[loc] = i

    df["_loc_order"] = df["Discharge Location"].map(loc_first_seen)
    df["_orig_order"] = range(len(df))

    df = df.sort_values(["_loc_order", "_orig_order"], kind="stable").reset_index(drop=True)

    out_rows = []
    prev_loc = None
    for _, row in df.iterrows():
        loc = row["Discharge Location"]
        if prev_loc is not None and loc != prev_loc:
            out_rows.append({c: "" for c in DN_COLUMNS})
        out_rows.append({c: row[c] for c in DN_COLUMNS})
        prev_loc = loc

    return pd.DataFrame(out_rows, columns=DN_COLUMNS)


# ==================================================================
# NO-DUPLICATE MODE
# ==================================================================

def _flat_values_from_first(first, mat_key, desc, need, date_str):
    """
    Pull every WO base-column value from the first row of a material
    group. Optional columns fall back to "" when absent.
    """
    values = {c: first.get(c, "") for c in FLAT_COLUMNS}
    values.update({
        "Schedule line date": date_str,     # normalised date
        "Material":           mat_key,      # leading zeros stripped
        "Description":        desc,
        "Quantity":           need,         # summed across the group
    })
    return values


def build_flat_layout_from_bytes(data: bytes, sheet: str,
                                 stock: pd.DataFrame):
    """
    Groups the flat sheet by Material, allocates stock, and returns
    a DataFrame with the base + pick columns.
    Adds a bold 'Miss Stock' row when short, with missed qty in Pick QTY.
    """
    df = pd.read_excel(BytesIO(data), sheet_name=sheet, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    for col in FLAT_COLUMNS:           # ensure the two optional columns exist
        if col not in df.columns:
            df[col] = ""

    new_cols = FLAT_COLUMNS + PICK_COLUMNS

    out_rows, problems = [], []
    prev_mat = None

    for mat_raw, group in df.groupby("Material", sort=False):
        mat_key = str(mat_raw).strip().lstrip("0")
        if mat_key == "" or mat_key.lower() == "nan":
            continue

        first = group.iloc[0]
        desc = str(first.get("Description", "")).strip()
        need = pd.to_numeric(group["Quantity"], errors="coerce").fillna(0).sum()

        raw_date = first.get("Schedule line date", "")
        parsed = pd.to_datetime(raw_date, errors="coerce")
        date_str = parsed.strftime("%d/%m/%Y") if pd.notna(parsed) else raw_date

        base = _flat_values_from_first(first, mat_key, desc, need, date_str)

        if prev_mat is not None:
            out_rows.append({c: "" for c in new_cols})

        # Per-material safety net.
        try:
            alloc = allocate(stock, mat_key, need)
        except Exception as e:
            problems.append(f"{mat_key}  ({desc})  – internal error: {e}")
            alloc = [{"short": need}]

        # ----- No usable stock at all -----
        if not alloc or (len(alloc) == 1 and "short" in alloc[0]):
            short_qty = alloc[0]["short"] if alloc else need
            problems.append(f"{mat_key}  ({desc})  – NO STOCK, "
                            f"short by {short_qty:.0f}")
            row = {c: "" for c in new_cols}
            row.update(base)
            row.update({
                "Pick Location": "Miss Stock",
                "Pick QTY":      short_qty,
            })
            out_rows.append(row)
            prev_mat = mat_key
            continue

        # ----- Normal picks + possible Miss Stock row -----
        cumulative = 0.0
        line_count = 0
        short_qty = 0

        for item in alloc:
            if "short" in item:
                short_qty = item["short"]
                continue
            cumulative += item["qty"]
            line_count += 1

            row = {c: "" for c in new_cols}
            if line_count == 1:
                row.update(base)
            row["Pick Location"] = item["bin"]
            row["Pick QTY"] = item["qty"]
            row["Cumulative"] = cumulative
            row["Batch"] = item["batch"]
            out_rows.append(row)

        if short_qty:
            miss_row = {c: "" for c in new_cols}
            miss_row["Pick Location"] = "Miss Stock"
            miss_row["Pick QTY"] = short_qty
            out_rows.append(miss_row)
            problems.append(f"{mat_key}  ({desc})  – SHORT by {short_qty:.0f}")

        prev_mat = mat_key

    return new_cols, out_rows, problems


# ==================================================================
# STYLING
# ==================================================================

HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
HEADER_FILL = PatternFill("solid", fgColor="00B0F0")

HEADER_ALIGN_WRAP = Alignment(horizontal="center", vertical="center", wrap_text=True)
HEADER_ALIGN_NOWRAP = Alignment(horizontal="center", vertical="center", wrap_text=False)
LEFT_ALIGN = Alignment(horizontal="left", vertical="center")
RIGHT_ALIGN = Alignment(horizontal="right", vertical="center")
CENTER_ALIGN = Alignment(horizontal="center", vertical="center")

DARK_SIDE = Side(style="thin", color="404040")
DARK_BORDER = Border(left=DARK_SIDE, right=DARK_SIDE,
                     top=DARK_SIDE, bottom=DARK_SIDE)

NO_WRAP_HEADERS = {"Sum of Quantity", "Sched.agreemnt"}

# Missed-quantity styling (red + bold)
MISS_QTY_FONT = Font(bold=True, color="FF0000")


def _style_header(ws, headers):
    for c in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.border = DARK_BORDER
        name = headers[c - 1] if headers else None
        cell.alignment = (HEADER_ALIGN_NOWRAP
                          if name in NO_WRAP_HEADERS else HEADER_ALIGN_WRAP)
    ws.row_dimensions[1].height = 32


def _clean_value(col_name, value):
    """Return a cleaned cell value: blank→None, numeric→int when integral."""
    if value is None:
        return None
    if isinstance(value, float) and pd.isna(value):
        return None
    if value == "":
        return None
    if col_name in INT_COLUMNS:
        try:
            f = float(value)
            if f.is_integer():
                return int(f)
        except (TypeError, ValueError):
            pass
    return value


def _write_df(ws, df: pd.DataFrame):
    for c_idx, col_name in enumerate(df.columns, start=1):
        ws.cell(row=1, column=c_idx, value=col_name)
    for r_idx, row in enumerate(df.itertuples(index=False, name=None), start=2):
        for c_idx, value in enumerate(row, start=1):
            col_name = df.columns[c_idx - 1]
            ws.cell(row=r_idx, column=c_idx,
                    value=_clean_value(col_name, value))


# -------------------- FLEXIBLE AUTOFIT --------------------

def _display_width(value) -> int:
    """Approximate how many characters a value needs."""
    if value is None:
        return 0
    text = str(value)
    return len(text) if text else 0


def autofit_all(ws, headers, preferred: dict = None):
    """
    Autofit EVERY column from header + cell content.
    `preferred` provides a floor per column name (never shrinks below it),
    but the actual width can grow to fit longer content, capped at
    AUTOFIT_MAX_WIDTH. Presets don't hide content.
    """
    preferred = preferred or {}
    for c_idx, name in enumerate(headers, start=1):
        longest = _display_width(name)
        for r in range(2, ws.max_row + 1):
            v = ws.cell(r, c_idx).value
            longest = max(longest, _display_width(v))
        floor = preferred.get(name, AUTOFIT_MIN_WIDTH)
        width = max(floor, longest + AUTOFIT_PADDING)
        width = min(width, AUTOFIT_MAX_WIDTH)
        ws.column_dimensions[get_column_letter(c_idx)].width = width


# ==================================================================
# PROBLEMS SHEET
# ==================================================================

def _write_problems_sheet(wb, problems):
    """Add a 'Problems' sheet listing short materials (only if any)."""
    if not problems:
        return
    if SHEET_PROBLEMS in wb.sheetnames:
        del wb[SHEET_PROBLEMS]

    ws = wb.create_sheet(SHEET_PROBLEMS)
    ws.cell(row=1, column=1, value="Issues")
    ws.cell(row=1, column=1).font = HEADER_FONT
    ws.cell(row=1, column=1).fill = HEADER_FILL
    ws.cell(row=1, column=1).border = DARK_BORDER
    ws.cell(row=1, column=1).alignment = HEADER_ALIGN_NOWRAP
    ws.row_dimensions[1].height = 24

    for r, p in enumerate(problems, start=2):
        cell = ws.cell(row=r, column=1, value=p)
        cell.alignment = LEFT_ALIGN

    longest = max((_display_width(p) for p in problems), default=20)
    ws.column_dimensions["A"].width = min(max(longest + 3, 40), 120)


# ==================================================================
# NO-DUPLICATE WRITER
# ==================================================================

def _write_flat_output(wb, sheet_name: str, headers: list, rows: list):
    """Replace the flat sheet with the pick-list-extended version,
    renamed to 'WO', and style it. Bolds Miss Stock rows and shows the
    missed quantity in red."""
    if sheet_name in wb.sheetnames:
        idx = wb.sheetnames.index(sheet_name)
        del wb[sheet_name]
    else:
        idx = 0

    ws = wb.create_sheet(SHEET_FLAT, idx)

    for c_idx, col in enumerate(headers, start=1):
        ws.cell(row=1, column=c_idx, value=col)
    for r_idx, row in enumerate(rows, start=2):
        for c_idx, col in enumerate(headers, start=1):
            ws.cell(row=r_idx, column=c_idx,
                    value=_clean_value(col, row.get(col, "")))

    n_cols = len(headers)

    for c_idx, cell in enumerate(ws[1], start=1):
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.border = DARK_BORDER
        name = headers[c_idx - 1]
        cell.alignment = (HEADER_ALIGN_NOWRAP
                          if name in NO_WRAP_HEADERS else HEADER_ALIGN_WRAP)
    ws.row_dimensions[1].height = 32

    pick_loc_idx = headers.index("Pick Location") + 1
    pick_qty_idx = headers.index("Pick QTY") + 1

    for r in range(2, ws.max_row + 1):
        is_separator = all(
            ws.cell(r, c).value in (None, "")
            for c in range(1, n_cols + 1))
        if is_separator:
            continue
        is_material_start = ws.cell(r, 3).value not in (None, "")
        is_miss = str(ws.cell(r, pick_loc_idx).value).strip() == "Miss Stock"

        for c in range(1, n_cols + 1):
            cell = ws.cell(r, c)
            header = headers[c - 1]

            if header == "Batch":
                cell.alignment = RIGHT_ALIGN
            elif header in ("Sum of Quantity", "Quantity",
                            "Unrestricted-Use Stock"):
                cell.alignment = CENTER_ALIGN
            else:
                cell.alignment = LEFT_ALIGN

            if header in PICK_COLUMNS:
                cell.border = DARK_BORDER
            elif header in FLAT_COLUMNS and is_material_start:
                cell.border = DARK_BORDER
            else:
                cell.border = None

        if is_material_start:
            for c, header in enumerate(headers, start=1):
                if header in FLAT_COLUMNS:
                    ws.cell(r, c).font = Font(bold=True)

        if is_miss:
            for c in range(1, n_cols + 1):
                ws.cell(r, c).font = Font(bold=True)
            ws.cell(r, pick_qty_idx).font = MISS_QTY_FONT

    preferred = {
        "Schedule line date": 14,
        "Discharge Location": 14,
        "Material": 16,
        "Quantity": 12,
        "Unrestricted-Use Stock": UNRESTRICTED_WIDTH,
        "Comments": COMMENTS_WIDTH,
        "Customer Material Number": 28,
        "Sched.agreemnt": SCHED_AGREEMENT_WIDTH,
        "Pick Location": PICK_LOCATION_WIDTH,
        "Pick QTY": PICK_QTY_WIDTH,
        "Cumulative": CUMULATIVE_WIDTH,
        "Batch": BATCH_WIDTH,
        "Sum of Quantity": SUM_OF_QUANTITY_WIDTH,
    }
    autofit_all(ws, headers, preferred=preferred)


# ==================================================================
# MERGE HELPERS  (Pick List readability)
# ==================================================================

def _merge_block(ws, start_row, end_row, cols):
    """Merge a rectangular block if it spans more than one row."""
    if start_row is None or end_row is None or end_row <= start_row:
        return
    for col in cols:
        ws.merge_cells(start_row=start_row, start_column=col,
                       end_row=end_row, end_column=col)


def _merge_material_blocks(ws, col_start, col_end):
    """
    Merge the Material/Description/Sum-of-Quantity cells (columns
    col_start..col_end) across consecutive pick rows of one material.
    A new material block starts whenever column 1 has a value.
    """
    block_start, prev_mat = None, None
    for r in range(2, ws.max_row + 1):
        mat = ws.cell(r, 1).value
        is_blank = mat in (None, "")
        if is_blank:
            _merge_block(ws, block_start, r - 1,
                         range(col_start, col_end + 1))
            block_start, prev_mat = None, None
            continue
        if mat != prev_mat:
            _merge_block(ws, block_start, r - 1,
                         range(col_start, col_end + 1))
            block_start, prev_mat = r, mat
    _merge_block(ws, block_start, ws.max_row,
                 range(col_start, col_end + 1))


# ==================================================================
# PUBLIC BUILDER
# ==================================================================

def build_picklist_workbook(stock_bytes: bytes,
                            demand_bytes: bytes) -> BytesIO:
    stock = load_stock_from_bytes(stock_bytes)
    layout = detect_demand_from_bytes(demand_bytes)

    wb = load_workbook(BytesIO(demand_bytes))

    # ---------------- NO-DUPLICATE MODE ----------------
    if not layout["duplicates"]:
        headers, rows, problems = build_flat_layout_from_bytes(
            demand_bytes, layout["sheet"], stock
        )
        _write_flat_output(wb, layout["sheet"], headers, rows)
        _write_problems_sheet(wb, problems)

        out = BytesIO()
        wb.save(out)
        out.seek(0)
        return out

    # ---------------- DUPLICATE MODE ----------------
    demand = load_demand_from_bytes(demand_bytes)
    dn_source = load_dn_source_from_bytes(demand_bytes)
    flat_df, flat_sheet_name = load_flat_sheet_keep_columns_from_bytes(demand_bytes)

    summary = summarize(demand)
    picklist, problems = build_picklist(stock, summary)
    dn = build_dn(dn_source)

    for s in (SHEET_SUMMARY, SHEET_OUT, SHEET_DN, SHEET_PROBLEMS):
        if s in wb.sheetnames:
            del wb[s]
    if flat_sheet_name in wb.sheetnames:
        idx = wb.sheetnames.index(flat_sheet_name)
        del wb[flat_sheet_name]
    else:
        idx = 0

    # ---------- WO ----------
    ws_flat = wb.create_sheet(SHEET_FLAT, idx)
    _write_df(ws_flat, flat_df)
    _style_header(ws_flat, list(flat_df.columns))
    n_flat = len(flat_df.columns)
    flat_headers = list(flat_df.columns)
    flat_qty_idx = flat_headers.index("Quantity") + 1
    for r in range(2, ws_flat.max_row + 1):
        if all(ws_flat.cell(r, c).value in (None, "") for c in range(1, n_flat + 1)):
            continue
        for c in range(1, n_flat + 1):
            cell = ws_flat.cell(r, c)
            cell.alignment = CENTER_ALIGN if c == flat_qty_idx else LEFT_ALIGN
            cell.border = DARK_BORDER
    autofit_all(ws_flat, flat_headers, preferred={
        "Schedule line date": 14,
        "Discharge Location": 14,
        "Material": 16,
        "Quantity": 12,
        "Unrestricted-Use Stock": UNRESTRICTED_WIDTH,
        "Comments": COMMENTS_WIDTH,
        "Customer Material Number": 28,
        "Sched.agreemnt": SCHED_AGREEMENT_WIDTH,
        "Sum of Quantity": SUM_OF_QUANTITY_WIDTH,
    })

    # ---------- Summary PN ----------
    ws_sum = wb.create_sheet(SHEET_SUMMARY)
    summary_out = summary.rename(columns={
        "material": "Material", "description": "Description",
        "quantity": "Sum of Quantity"})
    _write_df(ws_sum, summary_out)
    _style_header(ws_sum, list(summary_out.columns))
    sum_headers = list(summary_out.columns)
    sum_qty_idx = sum_headers.index("Sum of Quantity") + 1
    for r in range(2, ws_sum.max_row + 1):
        for c in range(1, 4):
            cell = ws_sum.cell(r, c)
            cell.alignment = CENTER_ALIGN if c == sum_qty_idx else LEFT_ALIGN
            cell.border = DARK_BORDER
    autofit_all(ws_sum, sum_headers, preferred={
        "Material": 16,
        "Sum of Quantity": SUM_OF_QUANTITY_WIDTH,
    })

    # ---------- Pick List ----------
    ws_pick = wb.create_sheet(SHEET_OUT)
    _write_df(ws_pick, picklist)
    _style_header(ws_pick, list(picklist.columns))
    pick_headers = list(picklist.columns)
    pick_loc_idx = pick_headers.index("Pick Location") + 1
    pick_qty_idx = pick_headers.index("Pick QTY") + 1
    for row in range(2, ws_pick.max_row + 1):
        if all(ws_pick.cell(row, c).value in (None, "") for c in range(1, 8)):
            continue
        is_start = ws_pick.cell(row, 1).value not in (None, "")
        is_miss = str(ws_pick.cell(row, pick_loc_idx).value).strip() == "Miss Stock"
        for col in range(1, 8):
            c = ws_pick.cell(row, col)
            header = pick_headers[col - 1]
            if header == "Batch":
                c.alignment = RIGHT_ALIGN
            elif header == "Sum of Quantity":
                c.alignment = CENTER_ALIGN
            else:
                c.alignment = LEFT_ALIGN
            if 4 <= col <= 7:
                c.border = DARK_BORDER
            elif col in (1, 2, 3) and is_start:
                c.border = DARK_BORDER
            else:
                c.border = None
        if is_start:
            for c in (1, 2, 3):
                ws_pick.cell(row, c).font = Font(bold=True)
        if is_miss:
            for c in range(1, 8):
                ws_pick.cell(row, c).font = Font(bold=True)
            ws_pick.cell(row, pick_qty_idx).font = MISS_QTY_FONT

    # Merge Material / Description / Sum-of-Quantity across pick rows
    _merge_material_blocks(ws_pick, 1, 3)

    autofit_all(ws_pick, pick_headers, preferred={
        "Material": 16,
        "Sum of Quantity": SUM_OF_QUANTITY_WIDTH,
        "Pick Location": PICK_LOCATION_WIDTH,
        "Pick QTY": PICK_QTY_WIDTH,
        "Cumulative": CUMULATIVE_WIDTH,
        "Batch": BATCH_WIDTH,
    })

    # ---------- DN ----------
    ws_dn = wb.create_sheet(SHEET_DN)
    _write_df(ws_dn, dn)
    _style_header(ws_dn, list(dn.columns))
    n_dn = len(dn.columns)
    dn_headers = list(dn.columns)
    dn_qty_idx = dn_headers.index("Quantity") + 1
    dn_col_idx = dn_headers.index("DN") + 1
    sched_col_idx = dn_headers.index("Sched.agreemnt") + 1
    loc_col_idx = dn_headers.index("Discharge Location") + 1
    for row in range(2, ws_dn.max_row + 1):
        if all(ws_dn.cell(row, c).value in (None, "") for c in range(1, n_dn + 1)):
            continue
        for col in range(1, n_dn + 1):
            cell = ws_dn.cell(row, col)
            cell.alignment = CENTER_ALIGN if col == dn_qty_idx else LEFT_ALIGN
            cell.border = DARK_BORDER
    autofit_all(ws_dn, dn_headers, preferred={
        "Schedule line date": 14,
        "Discharge Location": 14,
        "Material": 16,
        "Quantity": 12,
        "Customer Material Number": 28,
        "Sched.agreemnt": SCHED_AGREEMENT_WIDTH,
        "Sum of Quantity": SUM_OF_QUANTITY_WIDTH,
    })

    # Merge DN column per (Discharge Location, Sched.agreemnt) block
    block_start, prev_key = None, None
    for r in range(2, ws_dn.max_row + 1):
        is_blank = all(
            ws_dn.cell(r, c).value in (None, "")
            for c in range(1, n_dn + 1))
        if is_blank:
            if block_start is not None and r - 1 > block_start:
                ws_dn.merge_cells(
                    start_row=block_start, start_column=dn_col_idx,
                    end_row=r - 1, end_column=dn_col_idx)
            block_start, prev_key = None, None
            continue
        loc = ws_dn.cell(r, loc_col_idx).value
        sched = ws_dn.cell(r, sched_col_idx).value
        key = (loc, sched)
        if key != prev_key:
            if block_start is not None and r - 1 > block_start:
                ws_dn.merge_cells(
                    start_row=block_start, start_column=dn_col_idx,
                    end_row=r - 1, end_column=dn_col_idx)
            block_start, prev_key = r, key
    if block_start is not None and ws_dn.max_row > block_start:
        ws_dn.merge_cells(
            start_row=block_start, start_column=dn_col_idx,
            end_row=ws_dn.max_row, end_column=dn_col_idx)

    for r in range(2, ws_dn.max_row + 1):
        cell = ws_dn.cell(r, dn_col_idx)
        cell.alignment = Alignment(horizontal="left", vertical="center")

    # ---------- Problems ----------
    _write_problems_sheet(wb, problems)

    # ---------- Sheet order ----------
    for name, pos in ((SHEET_SUMMARY, SHEET_SUMMARY_POSITION),
                      (SHEET_OUT, SHEET_PICKLIST_POSITION),
                      (SHEET_DN, SHEET_DN_POSITION)):
        cur = wb.sheetnames.index(name)
        tgt = min(pos - 1, len(wb.sheetnames) - 1)
        wb.move_sheet(name, offset=tgt - cur)

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out


# ==================================================================
# CACHED WRAPPERS  (avoid re-parsing on every click)
# ==================================================================

@st.cache_data(show_spinner=False)
def cached_stock(data: bytes) -> pd.DataFrame:
    return load_stock_from_bytes(data)


@st.cache_data(show_spinner=False)
def cached_layout(data: bytes) -> dict:
    return detect_demand_from_bytes(data)


# ==================================================================
# STREAMLIT UI
# ==================================================================

st.set_page_config(page_title="Pick List Generator", page_icon="📦", layout="wide")

st.title("📦 Pick List Generator")
st.caption("Upload the **stock** file and the **demand** file to generate "
           "the pick list. The app auto-detects whether the demand has "
           "duplicate materials and produces the appropriate layout.")

col1, col2 = st.columns(2)

with col1:
    stock_file = st.file_uploader(
        "1️⃣ Stock file",
        type=["xlsx"],
        accept_multiple_files=False,
        key="stock",
    )

with col2:
    demand_file = st.file_uploader(
        "2️⃣ Demand file",
        type=["xlsx"],
        accept_multiple_files=False,
        key="demand",
    )

if stock_file and demand_file:
    st.success("Both files received. Ready to generate the pick list.")

    # Clear stale output whenever the uploaded files change
    current = (stock_file.name, demand_file.name)
    if st.session_state.get("last_files") != current:
        st.session_state.pop("output", None)
        st.session_state.pop("output_name", None)
        st.session_state["last_files"] = current

    # ---------- Preview ----------
    with st.expander("🔍 Preview parsed inputs", expanded=False):
        try:
            stock_prev = cached_stock(stock_file.getvalue())
            layout_prev = cached_layout(demand_file.getvalue())
            st.write(f"**Mode:** "
                     f"{'duplicate' if layout_prev['duplicates'] else 'no-duplicate'} "
                     f"(sheet `{layout_prev['sheet']}`)")
            st.write(f"**Stock rows:** {len(stock_prev)}")
            st.dataframe(stock_prev.head(50), use_container_width=True)
        except Exception as e:
            st.error(f"Preview failed: {e}")

    # ---------- Generate ----------
    if st.button("🚀 Generate Pick List", type="primary"):
        with st.spinner("Allocating stock and building the workbook…"):
            try:
                output = build_picklist_workbook(
                    stock_bytes=stock_file.getvalue(),
                    demand_bytes=demand_file.getvalue(),
                )
                st.session_state["output"] = output
                st.session_state["output_name"] = f"PickList_{demand_file.name}"
                st.success("✅ Workbook generated successfully.")
            except Exception as e:
                st.error(f"❌ Error while generating the workbook:\n\n{e}")

    if "output" in st.session_state:
        st.download_button(
            label="📥 Download pick list",
            data=st.session_state["output"],
            file_name=st.session_state["output_name"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
else:
    st.info("Please upload **both** files to continue.")

st.divider()
st.caption(
    "**Notes** — The output workbook depends on the demand file:\n\n"
    "- **No duplicate materials** → the flat sheet is renamed to `WO` and the "
    "columns `Pick Location`, `Pick QTY`, `Cumulative`, `Batch` are appended "
    "to the right of the original columns. If a material is short on stock, a "
    "bold `Miss Stock` row is added below its pick lines, with the missed "
    "quantity shown in red in the `Pick QTY` column.\n"
    "- **Duplicate materials** → the workbook contains `WO`, `Summary PN`, "
    "`Pick List`, and `DN` sheets. Short materials show a bold `Miss Stock` row "
    "with the missed quantity in red.\n"
    "- If any material was short, a **Problems** sheet is added with a summary "
    "of every shortfall.\n\n"
    "**Allocation rules** — Batches are consumed FIFO. Within a batch, the "
    "first pick chooses the bin whose available qty is closest to the demand; "
    "each following pick chooses the bin that makes the running cumulative "
    "closest to the demand."
)