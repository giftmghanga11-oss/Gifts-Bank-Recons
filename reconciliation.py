import datetime
from io import BytesIO
import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

EXCEL_EPOCH = datetime.datetime(1899, 12, 30)


def _clean_amount(value):
    """Converts numbers, currency strings, or Excel serial-date values back to float."""
    if value is None or (isinstance(value, float) and np.isnan(value)):
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


def _direction_multiplier(value):
    if pd.isna(value):
        return 1.0
    text = str(value).strip().lower()
    # Outgoing bank transfers / payments (ZATRF, ZABTR, Debits) vs incoming receipts (ZARWT, POSR, Credits)
    if any(x in text for x in ["debit", "withdrawal", "payment", "dr", "zatrf", "zabtr"]):
        return -1.0
    if any(x in text for x in ["credit", "deposit", "receipt", "cr", "zarwt", "posr", "zawtr", "zacsh"]):
        return 1.0
    return 1.0


def repair_cashbook_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Fixes date-formatted amounts and #VALUE! cells in ERP exports (e.g. Payment amount vs Paid amount)."""
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


def normalize_transactions(df: pd.DataFrame, mapping: dict, source: str) -> pd.DataFrame:
    working_df = repair_cashbook_dataframe(df) if source == "ERP" else df.copy()

    rows = []
    debit_col = mapping.get("debit")
    credit_col = mapping.get("credit")
    amount_col = mapping.get("amount")
    direction_col = mapping.get("direction")

    if not direction_col and source == "ERP":
        for cand in ["Payment no.", "Payment method"]:
            if cand in working_df.columns:
                direction_col = cand
                break

    for idx, row in working_df.iterrows():
        src_row = int(idx) + 2
        raw_dict = {col: row[col] for col in working_df.columns}

        if debit_col or credit_col:
            dr_val = abs(_clean_amount(row[debit_col]) or 0.0) if (debit_col and debit_col in working_df.columns) else 0.0
            cr_val = abs(_clean_amount(row[credit_col]) or 0.0) if (credit_col and credit_col in working_df.columns) else 0.0

            if dr_val > 0.0 and cr_val == 0.0:
                rows.append({
                    "source_row": src_row,
                    "abs_amount": dr_val,
                    "flow": "DEBIT",
                    "source": source,
                    "_raw": raw_dict,
                })
            elif cr_val > 0.0 and dr_val == 0.0:
                rows.append({
                    "source_row": src_row,
                    "abs_amount": cr_val,
                    "flow": "CREDIT",
                    "source": source,
                    "_raw": raw_dict,
                })
        elif amount_col and amount_col in working_df.columns:
            raw = _clean_amount(row[amount_col])
            if raw is None or raw == 0.0:
                continue
            if direction_col and direction_col in working_df.columns:
                mult = _direction_multiplier(row[direction_col])
                signed_amt = abs(raw) * mult
            else:
                signed_amt = raw

            if source == "ERP":
                flow = "DEBIT" if signed_amt > 0 else "CREDIT"
            else:
                flow = "CREDIT" if signed_amt > 0 else "DEBIT"

            rows.append({
                "source_row": src_row,
                "abs_amount": abs(signed_amt),
                "flow": flow,
                "source": source,
                "_raw": raw_dict,
            })
        else:
            raise ValueError(f"{source}: select either Debit/Credit columns or an Amount column.")

    cols = ["source_row", "abs_amount", "flow", "source", "_raw"]
    out = pd.DataFrame(rows, columns=cols) if rows else pd.DataFrame(columns=cols)
    out.attrs["raw_columns"] = list(working_df.columns)
    return out


def _align_side(
    bank_side_df: pd.DataFrame,
    erp_side_df: pd.DataFrame,
    pair_label: str,
    tolerance: float,
) -> list[dict]:
    b_sorted = bank_side_df.sort_values(by=["abs_amount", "source_row"], ascending=[True, True]).reset_index(drop=True)
    e_sorted = erp_side_df.sort_values(by=["abs_amount", "source_row"], ascending=[True, True]).reset_index(drop=True)

    n_b, n_e = len(b_sorted), len(e_sorted)
    b_to_e: dict[int, int] = {}
    e_used: set[int] = set()

    e_start = 0
    for bi in range(n_b):
        b_amt = float(b_sorted.at[bi, "abs_amount"])
        while e_start < n_e and (b_amt - float(e_sorted.at[e_start, "abs_amount"])) >= tolerance:
            e_start += 1

        best_ei = None
        best_diff = tolerance
        for ei in range(e_start, n_e):
            e_amt = float(e_sorted.at[ei, "abs_amount"])
            if (e_amt - b_amt) >= tolerance:
                break
            if ei in e_used:
                continue
            diff = abs(b_amt - e_amt)
            if diff < best_diff:
                best_diff = diff
                best_ei = ei
                if best_diff == 0.0:
                    break

        if best_ei is not None:
            b_to_e[bi] = best_ei
            e_used.add(best_ei)

    records = []
    b_ptr, e_ptr = 0, 0

    while b_ptr < n_b or e_ptr < n_e:
        while e_ptr < n_e and e_ptr in e_used:
            e_ptr += 1

        has_b = b_ptr < n_b
        has_e = e_ptr < n_e

        if has_b and b_ptr in b_to_e:
            ei = b_to_e[b_ptr]
            b_amt = float(b_sorted.at[b_ptr, "abs_amount"])
            e_amt = float(e_sorted.at[ei, "abs_amount"])
            records.append({
                "status": "Matched",
                "pair": pair_label,
                "variance_val": b_amt - e_amt,
                "bank_raw": b_sorted.at[b_ptr, "_raw"],
                "erp_raw": e_sorted.at[ei, "_raw"],
            })
            b_ptr += 1
        elif has_b and (not has_e or float(b_sorted.at[b_ptr, "abs_amount"]) <= float(e_sorted.at[e_ptr, "abs_amount"])):
            b_amt = float(b_sorted.at[b_ptr, "abs_amount"])
            records.append({
                "status": "Bank Only",
                "pair": pair_label,
                "variance_val": b_amt,
                "bank_raw": b_sorted.at[b_ptr, "_raw"],
                "erp_raw": None,
            })
            b_ptr += 1
        elif has_e:
            e_amt = float(e_sorted.at[e_ptr, "abs_amount"])
            records.append({
                "status": "ERP Only",
                "pair": pair_label,
                "variance_val": -e_amt,
                "bank_raw": None,
                "erp_raw": e_sorted.at[e_ptr, "_raw"],
            })
            e_ptr += 1

    return records


def reconcile(bank: pd.DataFrame, cashbook: pd.DataFrame, tolerance: float = 1.0) -> dict:
    if tolerance <= 0:
        raise ValueError("Tolerance must be greater than zero.")

    bank_cols = bank.attrs.get("raw_columns", [])
    erp_cols = cashbook.attrs.get("raw_columns", [])

    bank_dr = bank[bank["flow"] == "DEBIT"]
    bank_cr = bank[bank["flow"] == "CREDIT"]
    erp_dr = cashbook[cashbook["flow"] == "DEBIT"]
    erp_cr = cashbook[cashbook["flow"] == "CREDIT"]

    # Bank Credits vs ERP Debits first, then Bank Debits vs ERP Credits at the bottom
    cr_vs_dr = _align_side(bank_cr, erp_dr, "Bank Credits vs ERP Debits", tolerance)
    dr_vs_cr = _align_side(bank_dr, erp_cr, "Bank Debits vs ERP Credits", tolerance)
    records = cr_vs_dr + dr_vs_cr

    # Build preview DataFrame matching exact Excel column layout: [Bank cols | Variance | ERP cols]
    preview_rows = []
    statuses = []
    for rec in records:
        b_raw = rec["bank_raw"] or {}
        e_raw = rec["erp_raw"] or {}
        row_data = {}
        for c in bank_cols:
            row_data[c] = b_raw.get(c, None)
        row_data["Variance"] = rec["variance_val"]
        for c in erp_cols:
            out_col = f"{c}_ERP" if c in row_data else c
            row_data[out_col] = e_raw.get(c, None)
        preview_rows.append(row_data)
        statuses.append(rec["status"])

    report_df = pd.DataFrame(preview_rows)

    return {
        "report_df": report_df,
        "statuses": statuses,
        "records": records,
        "bank_cols": bank_cols,
        "erp_cols": erp_cols,
        "summary": {
            "matched": sum(1 for s in statuses if s == "Matched"),
            "bank_only": sum(1 for s in statuses if s == "Bank Only"),
            "erp_only": sum(1 for s in statuses if s == "ERP Only"),
        },
    }


def build_reconciliation_excel(results: dict) -> bytes:
    """
    Exports a single-sheet Excel workbook matching the exact NCBA Bank rec format:
    Bank columns on the left, Variance formula column in the middle (blank header),
    Cashbook/ERP columns on the right, green fill for Bank-Only rows, red fill for ERP-Only rows.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    green_fill = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")
    red_fill = PatternFill(start_color="F4CCCC", end_color="F4CCCC", fill_type="solid")
    bank_header_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    bold_font = Font(bold=True)
    red_var_font = Font(color="FF0000")

    bank_cols = results["bank_cols"]
    erp_cols = results["erp_cols"]
    records = results["records"]

    # Row 1: Bank headers | None (Variance column) | ERP headers
    headers = list(bank_cols) + [None] + list(erp_cols)
    ws.append(headers)

    for c_idx in range(1, len(bank_cols) + 1):
        cell = ws.cell(row=1, column=c_idx)
        cell.fill = bank_header_fill
        cell.font = bold_font
    for c_idx in range(len(bank_cols) + 2, len(headers) + 1):
        ws.cell(row=1, column=c_idx).font = bold_font

    var_col_idx = len(bank_cols) + 1

    bank_dr_letter = get_column_letter(bank_cols.index("Debit Amount") + 1) if "Debit Amount" in bank_cols else None
    bank_cr_letter = get_column_letter(bank_cols.index("Credit Amount") + 1) if "Credit Amount" in bank_cols else None
    erp_amt_col = "Payment amount" if "Payment amount" in erp_cols else ("Paid amount" if "Paid amount" in erp_cols else None)
    erp_amt_letter = get_column_letter(var_col_idx + 1 + erp_cols.index(erp_amt_col)) if erp_amt_col else None

    for rec in records:
        b_raw = rec["bank_raw"] or {}
        e_raw = rec["erp_raw"] or {}
        row_vals = [b_raw.get(c) for c in bank_cols] + [None] + [e_raw.get(c) for c in erp_cols]
        ws.append(row_vals)
        cur_row = ws.max_row

        var_cell = ws.cell(row=cur_row, column=var_col_idx)
        if rec["pair"] == "Bank Credits vs ERP Debits" and bank_cr_letter and erp_amt_letter:
            var_cell.value = f"={bank_cr_letter}{cur_row}-{erp_amt_letter}{cur_row}"
        elif rec["pair"] == "Bank Debits vs ERP Credits" and bank_dr_letter and erp_amt_letter:
            var_cell.value = f"={bank_dr_letter}{cur_row}-{erp_amt_letter}{cur_row}"
        else:
            var_cell.value = rec["variance_val"]

        var_cell.number_format = "#,##0.00"
        var_cell.font = red_var_font

        fill = green_fill if rec["status"] == "Bank Only" else (red_fill if rec["status"] == "ERP Only" else None)
        if fill:
            for c_idx in range(1, len(headers) + 1):
                ws.cell(row=cur_row, column=c_idx).fill = fill

    money_cols = {"Debit Amount", "Credit Amount", "Running Balance ", "Payment amount", "Paid amount"}
    date_cols = {"Posting Date", "Value Date", "Payment date", "Invoice due date", "Payment due date", "Invoice date"}

    for c_idx, h in enumerate(headers, start=1):
        if h in money_cols:
            for r_idx in range(2, ws.max_row + 1):
                cell = ws.cell(row=r_idx, column=c_idx)
                if isinstance(cell.value, (int, float)):
                    cell.number_format = "#,##0.00"
        elif h in date_cols:
            for r_idx in range(2, ws.max_row + 1):
                cell = ws.cell(row=r_idx, column=c_idx)
                if isinstance(cell.value, (datetime.datetime, datetime.date, pd.Timestamp)):
                    cell.number_format = "mm-dd-yy"

        ws.column_dimensions[get_column_letter(c_idx)].width = 16

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()
