from io import BytesIO
import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter


def _clean_amount(value):
    if pd.isna(value):
        return None
    if isinstance(value, (int, float, np.number)):
        val = float(value)
        return None if np.isnan(val) else val
    if isinstance(value, str):
        cleaned = value.replace(",", "").replace("KSh", "").replace("KES", "").strip()
        if cleaned.startswith("(") and cleaned.endswith(")"):
            cleaned = "-" + cleaned[1:-1].strip()
        try:
            val = float(cleaned)
            return None if np.isnan(val) else val
        except ValueError:
            return None
    return None


def _direction_multiplier(value):
    if pd.isna(value):
        return 1.0
    text = str(value).strip().lower()
    # In Sage X3 NCBA Payment enquiry: ZATRF / ZABTR are outgoing bank transfers (ERP Credits),
    # while ZARWT / POSR are customer receipts (ERP Debits).
    if any(x in text for x in ["debit", "withdrawal", "payment", "dr", "zatrf", "zabtr"]):
        return -1.0
    if any(x in text for x in ["credit", "deposit", "receipt", "cr", "zarwt", "posr", "zawtr", "zacsh"]):
        return 1.0
    return 1.0


def normalize_transactions(df: pd.DataFrame, mapping: dict, source: str) -> pd.DataFrame:
    """
    Normalizes Bank or ERP rows while keeping all original columns intact for side-by-side export.
    For Bank: Debit Amount -> flow='DEBIT', Credit Amount -> flow='CREDIT'.
    For ERP: Receipts (ZARWT/POSR or positive amounts) -> flow='DEBIT' (matches Bank Credit);
             Transfers/Payments (ZATRF/ZABTR or negative amounts) -> flow='CREDIT' (matches Bank Debit).
    """
    rows = []
    date_col = mapping.get("date")
    ref_col = mapping.get("reference")
    desc_col = mapping.get("description")
    debit_col = mapping.get("debit")
    credit_col = mapping.get("credit")
    amount_col = mapping.get("amount")
    direction_col = mapping.get("direction")

    # Auto-detect Payment no. / Payment method as direction fallback on Sage X3 exports
    if not direction_col and source == "ERP":
        for cand in ["Payment no.", "Payment method"]:
            if cand in df.columns:
                direction_col = cand
                break

    for idx, row in df.iterrows():
        src_row = int(idx) + 2
        dt = pd.to_datetime(row[date_col], errors="coerce") if (date_col and date_col in df.columns) else pd.NaT
        ref = "" if not (ref_col and ref_col in df.columns) or pd.isna(row[ref_col]) else str(row[ref_col]).strip()
        desc = "" if not (desc_col and desc_col in df.columns) or pd.isna(row[desc_col]) else str(row[desc_col]).strip()

        raw_dict = {col: row[col] for col in df.columns}
        # Ensure clean numeric Paid amount is mirrored into Payment amount if Payment amount has #VALUE!/dates
        if source == "ERP" and "Paid amount" in raw_dict and "Payment amount" in raw_dict:
            clean_paid = _clean_amount(raw_dict["Paid amount"])
            if clean_paid is not None:
                raw_dict["Payment amount"] = clean_paid

        if debit_col or credit_col:
            dr_val = abs(_clean_amount(row[debit_col]) or 0.0) if (debit_col and debit_col in df.columns) else 0.0
            cr_val = abs(_clean_amount(row[credit_col]) or 0.0) if (credit_col and credit_col in df.columns) else 0.0

            if dr_val > 0.0 and cr_val == 0.0:
                rows.append({
                    "source_row": src_row, "date": dt, "reference": ref, "description": desc,
                    "amount": -dr_val, "abs_amount": dr_val, "flow": "DEBIT", "source": source,
                    "_raw": raw_dict,
                })
            elif cr_val > 0.0 and dr_val == 0.0:
                rows.append({
                    "source_row": src_row, "date": dt, "reference": ref, "description": desc,
                    "amount": cr_val, "abs_amount": cr_val, "flow": "CREDIT", "source": source,
                    "_raw": raw_dict,
                })
        elif amount_col and amount_col in df.columns:
            raw = _clean_amount(row[amount_col])
            if raw is None or raw == 0.0:
                continue
            if direction_col and direction_col in df.columns:
                mult = _direction_multiplier(row[direction_col])
                signed_amt = abs(raw) * mult
            else:
                signed_amt = raw

            # On ERP side, receipts (positive) are ERP Debits; transfers/payments (negative) are ERP Credits
            if source == "ERP":
                flow = "DEBIT" if signed_amt > 0 else "CREDIT"
            else:
                flow = "CREDIT" if signed_amt > 0 else "DEBIT"

            rows.append({
                "source_row": src_row, "date": dt, "reference": ref, "description": desc,
                "amount": signed_amt, "abs_amount": abs(signed_amt), "flow": flow, "source": source,
                "_raw": raw_dict,
            })
        else:
            raise ValueError(f"{source}: valid amount or debit/credit column mapping is required.")

    cols = ["source_row", "date", "reference", "description", "amount", "abs_amount", "flow", "source", "_raw"]
    df_out = pd.DataFrame(rows, columns=cols) if rows else pd.DataFrame(columns=cols)
    df_out.attrs["raw_columns"] = list(df.columns)
    return df_out


def _align_opposite_side(
    bank_side_df: pd.DataFrame,
    erp_side_df: pd.DataFrame,
    pair_label: str,
    tolerance: float,
) -> tuple[pd.DataFrame, list[dict], set, set]:
    b_sorted = bank_side_df.sort_values(by=["abs_amount", "source_row"], ascending=[True, True]).reset_index(drop=True)
    e_sorted = erp_side_df.sort_values(by=["abs_amount", "source_row"], ascending=[True, True]).reset_index(drop=True)

    n_b, n_e = len(b_sorted), len(e_sorted)
    b_to_e: dict[int, int] = {}
    e_used: set[int] = set()
    matched_b_rows: set[int] = set()
    matched_e_rows: set[int] = set()

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
            matched_b_rows.add(int(b_sorted.at[bi, "source_row"]))
            matched_e_rows.add(int(e_sorted.at[best_ei, "source_row"]))

    summary_rows = []
    template_records = []
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
            variance = b_amt - e_amt
            summary_rows.append({
                "Pair": pair_label, "Status": "Matched",
                "Bank Ref": b_sorted.at[b_ptr, "reference"], "Bank Date": b_sorted.at[b_ptr, "date"],
                "Bank Description": b_sorted.at[b_ptr, "description"], "Bank Amount": b_amt,
                "Variance": variance, "ERP Amount": e_amt,
                "ERP Description": e_sorted.at[ei, "description"], "ERP Date": e_sorted.at[ei, "date"],
                "ERP Ref": e_sorted.at[ei, "reference"],
            })
            template_records.append({
                "status": "Matched", "pair": pair_label,
                "bank_raw": b_sorted.at[b_ptr, "_raw"],
                "erp_raw": e_sorted.at[ei, "_raw"],
            })
            b_ptr += 1
        elif has_b and (not has_e or float(b_sorted.at[b_ptr, "abs_amount"]) <= float(e_sorted.at[e_ptr, "abs_amount"])):
            b_amt = float(b_sorted.at[b_ptr, "abs_amount"])
            summary_rows.append({
                "Pair": pair_label, "Status": "Bank Only",
                "Bank Ref": b_sorted.at[b_ptr, "reference"], "Bank Date": b_sorted.at[b_ptr, "date"],
                "Bank Description": b_sorted.at[b_ptr, "description"], "Bank Amount": b_amt,
                "Variance": b_amt, "ERP Amount": np.nan,
                "ERP Description": "", "ERP Date": pd.NaT, "ERP Ref": "",
            })
            template_records.append({
                "status": "Bank Only", "pair": pair_label,
                "bank_raw": b_sorted.at[b_ptr, "_raw"],
                "erp_raw": None,
            })
            b_ptr += 1
        elif has_e:
            e_amt = float(e_sorted.at[e_ptr, "abs_amount"])
            summary_rows.append({
                "Pair": pair_label, "Status": "ERP Only",
                "Bank Ref": "", "Bank Date": pd.NaT, "Bank Description": "", "Bank Amount": np.nan,
                "Variance": -e_amt, "ERP Amount": e_amt,
                "ERP Description": e_sorted.at[e_ptr, "description"], "ERP Date": e_sorted.at[e_ptr, "date"],
                "ERP Ref": e_sorted.at[e_ptr, "reference"],
            })
            template_records.append({
                "status": "ERP Only", "pair": pair_label,
                "bank_raw": None,
                "erp_raw": e_sorted.at[e_ptr, "_raw"],
            })
            e_ptr += 1

    cols = [
        "Pair", "Status", "Bank Ref", "Bank Date", "Bank Description",
        "Bank Amount", "Variance", "ERP Amount", "ERP Description", "ERP Date", "ERP Ref"
    ]
    df_out = pd.DataFrame(summary_rows, columns=cols) if summary_rows else pd.DataFrame(columns=cols)
    return df_out, template_records, matched_b_rows, matched_e_rows


def _build_investigation_df(side_df: pd.DataFrame, matched_rows: set, source_label: str) -> pd.DataFrame:
    cols = ["Source", "Flow", "Matched", "Reason", "Amount", "Date", "Reference", "Description", "Source Row"]
    if side_df.empty:
        return pd.DataFrame(columns=cols)
    counts = side_df["abs_amount"].value_counts().to_dict()
    flagged = []
    for _, r in side_df.iterrows():
        srow = int(r["source_row"])
        amt = float(r["abs_amount"])
        is_matched = srow in matched_rows
        dup_cnt = int(counts.get(amt, 1))
        if not is_matched or dup_cnt > 1:
            reasons = []
            if not is_matched:
                reasons.append("Unmatched")
            if dup_cnt > 1:
                reasons.append(f"Duplicate amount ({dup_cnt}x)")
            flagged.append({
                "Source": source_label, "Flow": r["flow"],
                "Matched": "Yes" if is_matched else "No",
                "Reason": " & ".join(reasons), "Amount": amt,
                "Date": r["date"], "Reference": r["reference"],
                "Description": r["description"], "Source Row": srow,
            })
    return pd.DataFrame(flagged, columns=cols) if flagged else pd.DataFrame(columns=cols)


def reconcile(bank: pd.DataFrame, cashbook: pd.DataFrame, tolerance: float = 1.0) -> dict:
    if tolerance <= 0:
        raise ValueError("Tolerance must be greater than zero.")

    bank_dr = bank[bank["flow"] == "DEBIT"]
    bank_cr = bank[bank["flow"] == "CREDIT"]
    erp_dr = cashbook[cashbook["flow"] == "DEBIT"]
    erp_cr = cashbook[cashbook["flow"] == "CREDIT"]

    # Bank Credits <-> ERP Debits (Receipts first, matching NCBA Bank rec.xlsx order)
    cr_vs_dr, rec_cr_dr, m_b_cr, m_e_dr = _align_opposite_side(bank_cr, erp_dr, "Bank Credits vs ERP Debits", tolerance)
    # Bank Debits <-> ERP Credits (Payments/Sweeps at bottom, matching NCBA Bank rec.xlsx order)
    dr_vs_cr, rec_dr_cr, m_b_dr, m_e_cr = _align_opposite_side(bank_dr, erp_cr, "Bank Debits vs ERP Credits", tolerance)

    all_df = pd.concat([cr_vs_dr, dr_vs_cr], ignore_index=True)
    matched_df = all_df[all_df["Status"] == "Matched"].copy()
    bank_only_df = all_df[all_df["Status"] == "Bank Only"].copy()
    erp_only_df = all_df[all_df["Status"] == "ERP Only"].copy()

    inv_parts = [
        _build_investigation_df(bank_cr, m_b_cr, "BANK"),
        _build_investigation_df(bank_dr, m_b_dr, "BANK"),
        _build_investigation_df(erp_dr, m_e_dr, "ERP"),
        _build_investigation_df(erp_cr, m_e_cr, "ERP"),
    ]
    non_empty = [p for p in inv_parts if not p.empty]
    investigate_df = pd.concat(non_empty, ignore_index=True) if non_empty else pd.DataFrame(
        columns=["Source", "Flow", "Matched", "Reason", "Amount", "Date", "Reference", "Description", "Source Row"]
    )

    total_bank_amt = float(bank["abs_amount"].sum()) if not bank.empty else 0.0
    total_erp_amt = float(cashbook["abs_amount"].sum()) if not cashbook.empty else 0.0
    matched_bank_amt = float(matched_df["Bank Amount"].sum()) if not matched_df.empty else 0.0
    matched_erp_amt = float(matched_df["ERP Amount"].sum()) if not matched_df.empty else 0.0
    outstanding_bank_amt = float(bank_only_df["Bank Amount"].sum()) if not bank_only_df.empty else 0.0
    outstanding_erp_amt = float(erp_only_df["ERP Amount"].sum()) if not erp_only_df.empty else 0.0
    net_variance = float(all_df["Variance"].sum()) if not all_df.empty else 0.0

    summary_table = pd.DataFrame([
        {"Metric": "Total Bank Transactions", "Value": total_bank_amt},
        {"Metric": "Total ERP Transactions", "Value": total_erp_amt},
        {"Metric": "Matched Bank Amount", "Value": matched_bank_amt},
        {"Metric": "Matched ERP Amount", "Value": matched_erp_amt},
        {"Metric": "Outstanding Bank Items (Green)", "Value": outstanding_bank_amt},
        {"Metric": "Outstanding ERP Items (Red)", "Value": outstanding_erp_amt},
        {"Metric": "Net Variance (Bank - ERP)", "Value": net_variance},
        {"Metric": "Matched Pairs Count", "Value": int(len(matched_df))},
        {"Metric": "Bank-Only Count", "Value": int(len(bank_only_df))},
        {"Metric": "ERP-Only Count", "Value": int(len(erp_only_df))},
        {"Metric": "Items Flagged for Investigation", "Value": int(len(investigate_df))},
    ])

    return {
        "summary": {
            "matched": int(len(matched_df)),
            "bank_only": int(len(bank_only_df)),
            "erp_only": int(len(erp_only_df)),
            "potential_duplicates": int(len(investigate_df)),
            "net_variance": net_variance,
        },
        "summary_table": summary_table,
        "debits_vs_credits": dr_vs_cr,
        "credits_vs_debits": cr_vs_dr,
        "matched": matched_df,
        "bank_only": bank_only_df,
        "erp_only": erp_only_df,
        "investigate": investigate_df,
        "all": all_df,
        "_template_records": rec_cr_dr + rec_dr_cr,
        "_bank_cols": bank.attrs.get("raw_columns", []),
        "_erp_cols": cashbook.attrs.get("raw_columns", []),
    }


def build_reconciliation_excel(results: dict) -> bytes:
    """
    Generates an Excel workbook whose primary sheet ('Reconciliation Report') matches
    NCBA Bank rec.xlsx layout: Bank columns (A..I), Variance formula column (J),
    ERP columns (K..), green fill for Bank-Only rows, red fill for ERP-Only rows,
    and reconciliation totals at the bottom, followed by detail tabs.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Reconciliation Report"

    green_fill = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")
    red_fill = PatternFill(start_color="F4CCCC", end_color="F4CCCC", fill_type="solid")
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    bold_font = Font(bold=True)

    bank_cols = results.get("_bank_cols", [])
    erp_cols = results.get("_erp_cols", [])
    records = results.get("_template_records", [])

    # Header row: Bank cols | Variance | ERP cols
    headers = list(bank_cols) + ["Variance"] + list(erp_cols)
    ws.append(headers)
    for c_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=c_idx)
        cell.fill = header_fill
        cell.font = header_font

    var_col_idx = len(bank_cols) + 1
    var_col_letter = get_column_letter(var_col_idx)

    # Locate Bank Debit/Credit and ERP Payment amount columns for live Excel variance formulas
    bank_dr_letter = get_column_letter(bank_cols.index("Debit Amount") + 1) if "Debit Amount" in bank_cols else None
    bank_cr_letter = get_column_letter(bank_cols.index("Credit Amount") + 1) if "Credit Amount" in bank_cols else None
    erp_amt_col = "Payment amount" if "Payment amount" in erp_cols else ("Paid amount" if "Paid amount" in erp_cols else None)
    erp_amt_letter = get_column_letter(var_col_idx + 1 + erp_cols.index(erp_amt_col)) if erp_amt_col else None

    for rec in records:
        b_raw = rec.get("bank_raw") or {}
        e_raw = rec.get("erp_raw") or {}
        row_vals = [b_raw.get(c) for c in bank_cols] + [None] + [e_raw.get(c) for c in erp_cols]
        ws.append(row_vals)
        cur_row = ws.max_row

        # Live Excel Variance formula (e.g. =H2-R2 for Credits, =G68-R68 for Debits)
        if rec["pair"] == "Bank Credits vs ERP Debits" and bank_cr_letter and erp_amt_letter:
            ws.cell(row=cur_row, column=var_col_idx).value = f"={bank_cr_letter}{cur_row}-{erp_amt_letter}{cur_row}"
        elif rec["pair"] == "Bank Debits vs ERP Credits" and bank_dr_letter and erp_amt_letter:
            ws.cell(row=cur_row, column=var_col_idx).value = f"={bank_dr_letter}{cur_row}-{erp_amt_letter}{cur_row}"

        ws.cell(row=cur_row, column=var_col_idx).number_format = "#,##0.00"

        # Highlight Bank-Only green and ERP-Only red
        fill = green_fill if rec["status"] == "Bank Only" else (red_fill if rec["status"] == "ERP Only" else None)
        if fill:
            for c_idx in range(1, len(headers) + 1):
                ws.cell(row=cur_row, column=c_idx).fill = fill

    last_data_row = ws.max_row

    # Reconciliation Totals & Summary block at the bottom
    ws.append([])
    summary_df = results.get("summary_table", pd.DataFrame())
    for _, s_row in summary_df.iterrows():
        ws.append([s_row["Metric"], s_row["Value"]])
        r_idx = ws.max_row
        ws.cell(row=r_idx, column=1).font = bold_font
        val_cell = ws.cell(row=r_idx, column=2)
        val_cell.font = bold_font
        if isinstance(s_row["Value"], float):
            val_cell.number_format = "#,##0.00"

    # Format currency and date columns cleanly
    money_names = {"Debit Amount", "Credit Amount", "Running Balance ", "Variance", "Payment amount", "Paid amount"}
    for c_idx, h in enumerate(headers, start=1):
        if h in money_names:
            for r_idx in range(2, last_data_row + 1):
                cell = ws.cell(row=r_idx, column=c_idx)
                if isinstance(cell.value, (int, float)):
                    cell.number_format = "#,##0.00"
        ws.column_dimensions[get_column_letter(c_idx)].width = 18

    # Secondary breakdown tabs
    extra_tabs = [
        ("Bank Credits vs ERP Debits", results.get("credits_vs_debits")),
        ("Bank Debits vs ERP Credits", results.get("debits_vs_credits")),
        ("Matched", results.get("matched")),
        ("Bank Only", results.get("bank_only")),
        ("ERP Only", results.get("erp_only")),
        ("Needs Investigation", results.get("investigate")),
    ]
    for title, df in extra_tabs:
        sub_ws = wb.create_sheet(title=title[:31])
        if df is None or df.empty:
            sub_ws.append(["No records"])
            continue
        sub_cols = list(df.columns)
        sub_ws.append(sub_cols)
        for c_i in range(1, len(sub_cols) + 1):
            sub_ws.cell(row=1, column=c_i).fill = header_fill
            sub_ws.cell(row=1, column=c_i).font = header_font
        for _, r in df.iterrows():
            sub_ws.append([None if pd.isna(r[c]) else r[c] for c in sub_cols])

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()
