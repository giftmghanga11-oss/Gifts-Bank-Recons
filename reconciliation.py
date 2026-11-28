"""Bank vs ERP cashbook reconciliation engine.

Layout of the report (one sheet):
    [ all bank columns ] | Variance | [ all ERP columns ]

Every bank/ERP amount column that the user pairs in the app is used for matching,
for the variance formulas and for number formatting, so banks with a different number
or naming of columns work without code changes.
"""
import datetime
from io import BytesIO

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

EXCEL_EPOCH = datetime.datetime(1899, 12, 30)
NONE_CHOICE = "(none)"

IN, OUT = "IN", "OUT"
REVERSAL_STATUSES = ("ERP Reversal", "Bank Reversal")
SECTION_LABELS = {
    IN: "Money in: bank credits vs ERP receipts",
    OUT: "Money out: bank debits vs ERP payments",
}


# --------------------------------------------------------------------------- cleaning
def _clean_amount(value):
    """Converts numbers, currency strings, or Excel serial-date values back to float."""
    if value is None:
        return None
    if isinstance(value, float) and np.isnan(value):
        return None
    if pd.isna(value):
        return None
    if isinstance(value, (datetime.datetime, datetime.date, pd.Timestamp)):
        dt = pd.Timestamp(value).to_pydatetime().replace(tzinfo=None)
        delta = dt - EXCEL_EPOCH
        return float(delta.days + delta.seconds / 86400.0)
    if isinstance(value, (int, float, np.number)):
        return float(value)
    if isinstance(value, str):
        s = value.strip()
        if not s or s.upper() in {"#VALUE!", "#REF!", "#NAME?", "#N/A", "NAN", "NONE"}:
            return None
        cleaned = s.replace(",", "").replace("KSh", "").replace("KES", "").strip()
        if cleaned.startswith("(") and cleaned.endswith(")"):
            cleaned = "-" + cleaned[1:-1].strip()
        try:
            return float(cleaned)
        except ValueError:
            parsed_dt = pd.to_datetime(s, errors="coerce")
            if pd.notna(parsed_dt) and (parsed_dt.year < 1990 or parsed_dt.year > 2035):
                dt = parsed_dt.to_pydatetime().replace(tzinfo=None)
                delta = dt - EXCEL_EPOCH
                return float(delta.days + delta.seconds / 86400.0)
            return None
    return None


def _clean_value(value):
    """Makes a cell safe to write to Excel (no NaN/NaT/numpy types)."""
    if value is None:
        return None
    if isinstance(value, pd.Timestamp):
        return None if pd.isna(value) else value.to_pydatetime()
    if isinstance(value, (np.integer,)):
        return int(value)
    if isinstance(value, (np.floating,)):
        return None if np.isnan(value) else float(value)
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    return value


def repair_cashbook_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Fixes date-formatted amounts and #VALUE! cells in ERP exports (Payment amount vs Paid amount)."""
    fixed = df.copy()
    if "Payment amount" in fixed.columns:
        coerced_pay = fixed["Payment amount"].apply(_clean_amount)
        if "Paid amount" in fixed.columns:
            coerced_paid = fixed["Paid amount"].apply(_clean_amount)
            coerced_pay = coerced_pay.fillna(coerced_paid)
        fixed["Payment amount"] = coerced_pay
    if "Paid amount" in fixed.columns:
        fixed["Paid amount"] = fixed["Paid amount"].apply(_clean_amount)
    return fixed


def _parse_keywords(text) -> list[str]:
    return [k.strip().lower() for k in str(text or "").split(",") if k.strip()]


def _is_money_out(value, keywords: list[str]) -> bool:
    if value is None or pd.isna(value):
        return False
    text = str(value).strip().lower()
    return any(k in text for k in keywords)


# --------------------------------------------------------------------------- normalising
def normalize_transactions(
    df: pd.DataFrame,
    in_col,
    out_col,
    source: str,
    direction_col=None,
    out_keywords="",
) -> pd.DataFrame:
    """
    Turns a statement into a list of (row, absolute amount, IN/OUT flow, raw row).

    in_col / out_col are the user-selected "money in" and "money out" amount columns:
      * two different columns  -> each non-zero value goes to its own side
      * the same column        -> one amount column; direction comes from `direction_col`
                                  (rows containing any of `out_keywords` are money out)
                                  or, with no direction column, from the sign (negative = out)
      * None                   -> that side does not exist in this statement
    Rows that cannot be classified are never dropped silently: they are listed in
    out.attrs["skipped"].
    """
    in_col = in_col if in_col in df.columns else None
    out_col = out_col if out_col in df.columns else None
    direction_col = direction_col if direction_col in df.columns else None
    keywords = _parse_keywords(out_keywords)
    amount_cols = {c for c in (in_col, out_col) if c}

    rows, skipped = [], []
    for idx, row in df.iterrows():
        src_row = int(idx) + 2  # Excel row number (header is row 1)

        raw = {c: _clean_value(row[c]) for c in df.columns}
        for c in amount_cols:
            raw[c] = _clean_amount(row[c])  # numeric so the variance formulas work

        if in_col and in_col == out_col:
            value = raw[in_col]
            if not value:
                continue
            if direction_col:
                flow = OUT if _is_money_out(row[direction_col], keywords) else IN
            else:
                flow = IN if value > 0 else OUT
            amount = abs(value)
        else:
            in_v = abs(raw[in_col] or 0.0) if in_col else 0.0
            out_v = abs(raw[out_col] or 0.0) if out_col else 0.0
            if in_v > 0 and out_v > 0:
                skipped.append((src_row, f"{source}: both money-in and money-out amounts on one row"))
                continue
            if in_v > 0:
                flow, amount = IN, in_v
            elif out_v > 0:
                flow, amount = OUT, out_v
            else:
                continue

        rows.append({"source_row": src_row, "abs_amount": amount, "flow": flow, "_raw": raw})

    cols = ["source_row", "abs_amount", "flow", "_raw"]
    out = pd.DataFrame(rows, columns=cols) if rows else pd.DataFrame(columns=cols)
    out.attrs["raw_columns"] = list(df.columns)
    out.attrs["skipped"] = skipped
    return out


# --------------------------------------------------------------------------- matching
def _align_side(bank_side: pd.DataFrame, erp_side: pd.DataFrame, section: str, tolerance: float) -> list[dict]:
    """
    Sort both sides smallest -> largest, match one-to-one (|difference| < tolerance), then lay the
    rows out so every matched pair shares a row. Unmatched rows on one side leave the other side
    blank, which pushes the later entries of the other side downwards.
    """
    b = bank_side.sort_values(["abs_amount", "source_row"]).reset_index(drop=True)
    e = erp_side.sort_values(["abs_amount", "source_row"]).reset_index(drop=True)
    b_amts = b["abs_amount"].to_numpy(dtype=float)
    e_amts = e["abs_amount"].to_numpy(dtype=float)
    n_b, n_e = len(b), len(e)

    # 1) match: for each bank entry take the closest unused ERP entry strictly within tolerance
    b_to_e: dict[int, int] = {}
    used: set[int] = set()
    for bi in range(n_b):
        lo = int(np.searchsorted(e_amts, b_amts[bi] - tolerance, side="right"))
        hi = int(np.searchsorted(e_amts, b_amts[bi] + tolerance, side="left"))
        best, best_diff = None, tolerance
        for ei in range(lo, hi):
            if ei in used:
                continue
            diff = abs(b_amts[bi] - e_amts[ei])
            if diff < best_diff:
                best, best_diff = ei, diff
                if diff == 0.0:
                    break
        if best is not None:
            b_to_e[bi] = best
            used.add(best)
    matched_e = set(b_to_e.values())

    # 2) lay out in ascending order
    records: list[dict] = []

    def erp_only(ei):
        records.append({
            "status": "ERP Only", "section": section, "variance_val": -float(e_amts[ei]),
            "bank_raw": None, "erp_raw": e.at[ei, "_raw"],
        })

    e_ptr = 0
    for bi in range(n_b):
        if bi in b_to_e:
            ei = b_to_e[bi]
            while e_ptr < ei:                      # smaller ERP-only entries come first
                if e_ptr not in matched_e:
                    erp_only(e_ptr)
                e_ptr += 1
            records.append({
                "status": "Matched", "section": section,
                "variance_val": float(b_amts[bi] - e_amts[ei]),
                "bank_raw": b.at[bi, "_raw"], "erp_raw": e.at[ei, "_raw"],
            })
            e_ptr = max(e_ptr, ei + 1)
        else:
            while e_ptr < n_e and e_ptr not in matched_e and e_amts[e_ptr] < b_amts[bi]:
                erp_only(e_ptr)
                e_ptr += 1
            records.append({
                "status": "Bank Only", "section": section, "variance_val": float(b_amts[bi]),
                "bank_raw": b.at[bi, "_raw"], "erp_raw": None,
            })
    while e_ptr < n_e:
        if e_ptr not in matched_e:
            erp_only(e_ptr)
        e_ptr += 1
    return records


def _pair_reversals(records: list[dict], tolerance: float) -> list[dict]:
    """
    An unmatched entry on the money-in side and an unmatched entry of the same amount on the money-out
    side, both on the SAME statement, are an entry and its reversal:
      * ERP Only (in)  + ERP Only (out)  -> "ERP Reversal"   (posted, then reversed in the cashbook)
      * Bank Only (in) + Bank Only (out) -> "Bank Reversal"  (credited, then debited back by the bank)
    The two legs are moved to a block at the bottom, side by side in adjacent rows. The second leg's
    variance is flipped, so each pair nets to nil and no longer shows as an unreconciled difference.
    One-to-one: each entry is used at most once; the difference must be strictly less than tolerance.
    """
    in_label, out_label = SECTION_LABELS[IN], SECTION_LABELS[OUT]
    amount = lambda i: abs(records[i]["variance_val"])
    pairs = []  # (new_status, in_index, out_index)
    for status, new_status in (("ERP Only", "ERP Reversal"), ("Bank Only", "Bank Reversal")):
        ins = sorted((i for i, r in enumerate(records) if r["status"] == status and r["section"] == in_label), key=amount)
        outs = [i for i, r in enumerate(records) if r["status"] == status and r["section"] == out_label]
        used = set()
        for i in ins:
            best, best_diff = None, tolerance
            for j in outs:
                if j in used:
                    continue
                diff = abs(amount(i) - amount(j))
                if diff < best_diff:
                    best, best_diff = j, diff
            if best is not None:
                used.add(best)
                pairs.append((new_status, i, best))

    if not pairs:
        return records
    moved = {i for _, i, j in pairs} | {j for _, i, j in pairs}
    kept = [r for k, r in enumerate(records) if k not in moved]
    block = []
    for new_status, i, j in sorted(pairs, key=lambda p: (amount(p[1]), p[0])):
        first = dict(records[i], status=new_status, reversal_leg="in")
        second = dict(records[j], status=new_status, reversal_leg="out", variance_val=-records[j]["variance_val"])
        block += [first, second]
    return kept + block


def reconcile(
    bank: pd.DataFrame,
    cashbook: pd.DataFrame,
    tolerance: float = 1.0,
    mapping: dict | None = None,
    pair_reversals: bool = True,
) -> dict:
    """`mapping` = {"bank_in","bank_out","erp_in","erp_out"} (column names or None); used for the Excel formulas."""
    if tolerance <= 0:
        raise ValueError("Tolerance must be greater than zero.")

    bank_cols = bank.attrs.get("raw_columns", [])
    erp_cols = cashbook.attrs.get("raw_columns", [])

    records = []
    for flow in (IN, OUT):  # money in first, money out at the bottom
        records += _align_side(
            bank[bank["flow"] == flow], cashbook[cashbook["flow"] == flow], SECTION_LABELS[flow], tolerance
        )

    if pair_reversals:
        records = _pair_reversals(records, tolerance)

    preview_rows, statuses = [], []
    for rec in records:
        b_raw = rec["bank_raw"] or {}
        e_raw = rec["erp_raw"] or {}
        row = {c: b_raw.get(c) for c in bank_cols}
        var_name = "Variance"
        while var_name in row:
            var_name += " (calc)"
        row[var_name] = rec["variance_val"]
        for c in erp_cols:
            name = c
            while name in row:          # keep every report column name unique (the styler needs this)
                name += "_ERP"
            row[name] = e_raw.get(c)
        preview_rows.append(row)
        statuses.append(rec["status"])

    skipped = list(bank.attrs.get("skipped", [])) + list(cashbook.attrs.get("skipped", []))
    return {
        "report_df": pd.DataFrame(preview_rows),
        "statuses": statuses,
        "records": records,
        "bank_cols": bank_cols,
        "erp_cols": erp_cols,
        "mapping": mapping or {},
        "skipped": skipped,
        "summary": {
            "matched": statuses.count("Matched"),
            "bank_only": statuses.count("Bank Only"),
            "erp_only": statuses.count("ERP Only"),
            "erp_reversals": statuses.count("ERP Reversal") // 2,
            "bank_reversals": statuses.count("Bank Reversal") // 2,
            "net_variance": float(sum(r["variance_val"] for r in records)),
            "bank_rows": len(bank),
            "erp_rows": len(cashbook),
        },
    }


# --------------------------------------------------------------------------- excel export
def build_reconciliation_excel(results: dict) -> bytes:
    """
    One sheet: bank columns | Variance (live formula) | ERP columns.
    Green rows = bank only, red rows = ERP only. Variance = bank amount - ERP amount, using the
    columns the user paired in the app.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    green = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")
    red = PatternFill(start_color="F4CCCC", end_color="F4CCCC", fill_type="solid")
    grey = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    yellow = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    blue = PatternFill(start_color="CFE2F3", end_color="CFE2F3", fill_type="solid")
    bold = Font(bold=True)
    var_font = Font(color="FF0000")

    bank_cols, erp_cols, records = results["bank_cols"], results["erp_cols"], results["records"]
    mp = results.get("mapping", {})
    var_idx = len(bank_cols) + 1
    erp_start = var_idx + 1

    headers = list(bank_cols) + [None] + list(erp_cols)
    ws.append(headers)
    for c in range(1, len(bank_cols) + 1):
        ws.cell(row=1, column=c).fill = grey
        ws.cell(row=1, column=c).font = bold
    for c in range(erp_start, len(headers) + 1):
        ws.cell(row=1, column=c).font = bold

    def bank_letter(col):
        return get_column_letter(bank_cols.index(col) + 1) if col in bank_cols else None

    def erp_letter(col):
        return get_column_letter(erp_start + erp_cols.index(col)) if col in erp_cols else None

    # does the ERP amount column hold negatives on a side? then wrap it in ABS() in the formula
    def erp_has_negative(col):
        return any(
            isinstance(r["erp_raw"].get(col), (int, float)) and r["erp_raw"].get(col) < 0
            for r in records if r["erp_raw"]
        ) if col else False

    side_cfg = {}
    for flow, b_key, e_key in ((IN, "bank_in", "erp_in"), (OUT, "bank_out", "erp_out")):
        b_col, e_col = mp.get(b_key), mp.get(e_key)
        side_cfg[SECTION_LABELS[flow]] = {
            "b_letter": bank_letter(b_col),
            "e_letter": erp_letter(e_col),
            "b_abs": b_col is not None and b_col == mp.get("bank_out" if flow == IN else "bank_in"),
            "e_abs": erp_has_negative(e_col),
        }

    for rec in records:
        b_raw, e_raw = rec["bank_raw"] or {}, rec["erp_raw"] or {}
        ws.append([b_raw.get(c) for c in bank_cols] + [None] + [e_raw.get(c) for c in erp_cols])
        r = ws.max_row

        cfg = side_cfg[rec["section"]]
        cell = ws.cell(row=r, column=var_idx)
        if cfg["b_letter"] and cfg["e_letter"]:
            b_ref = f"ABS({cfg['b_letter']}{r})" if cfg["b_abs"] else f"{cfg['b_letter']}{r}"
            e_ref = f"ABS({cfg['e_letter']}{r})" if cfg["e_abs"] else f"{cfg['e_letter']}{r}"
            cell.value = f"={b_ref}-{e_ref}"
            if rec.get("reversal_leg") == "out":      # reversal leg: flipped so the pair nets to nil
                cell.value = f"=-({b_ref}-{e_ref})"
        else:
            cell.value = rec["variance_val"]
        cell.number_format = "#,##0.00"
        cell.font = var_font

        fill = {"Bank Only": green, "ERP Only": red, "ERP Reversal": yellow, "Bank Reversal": blue}.get(rec["status"])
        if fill:
            for c in range(1, len(headers) + 1):
                ws.cell(row=r, column=c).fill = fill

    amount_cols = {mp.get(k) for k in ("bank_in", "bank_out", "erp_in", "erp_out")} - {None}
    for c_idx, h in enumerate(headers, start=1):
        for r_idx in range(2, ws.max_row + 1):
            cell = ws.cell(row=r_idx, column=c_idx)
            if isinstance(cell.value, (datetime.datetime, datetime.date)):
                cell.number_format = "dd-mmm-yy"
            elif isinstance(cell.value, (int, float)) and (h in amount_cols or h == "Running Balance "):
                cell.number_format = "#,##0.00"
        ws.column_dimensions[get_column_letter(c_idx)].width = 16

    ws.freeze_panes = "A2"
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()
