from dataclasses import dataclass
from io import BytesIO
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

TOLERANCE = 1.0  # Strictly < KSh 1.00 is accepted as matched


@dataclass
class SideItem:
    row_id: int
    side: str          # "BANK" or "ERP"
    flow: str          # "DEBIT" or "CREDIT"
    amount: float      # Positive magnitude
    date: str
    reference: str
    description: str
    matched: bool = False


def _clean_str(val: object) -> str:
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return ""
    return str(val).strip()


def _to_float(val: object) -> float:
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return 0.0
    if isinstance(val, (int, float, np.number)):
        return float(val)
    cleaned = str(val).replace(",", "").replace("KSh", "").replace("KES", "").strip()
    if cleaned.startswith("(") and cleaned.endswith(")"):
        cleaned = "-" + cleaned[1:-1].strip()
    try:
        return float(cleaned)
    except ValueError:
        return 0.0


def _find_col(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    lower_map = {str(c).strip().lower(): c for c in df.columns}
    for cand in candidates:
        if cand.lower() in lower_map:
            return lower_map[cand.lower()]
    for col in df.columns:
        c_low = str(col).strip().lower()
        if any(cand.lower() in c_low for cand in candidates):
            return col
    return None


def extract_items(df: pd.DataFrame, source_label: str) -> Tuple[List[SideItem], List[SideItem]]:
    """
    Splits a Bank or ERP DataFrame into positive-magnitude DEBIT and CREDIT SideItems.
    Supports either separate Debit/Credit columns or a single Amount column.
    """
    date_col = _find_col(df, ["date", "posting date", "tran date", "value date"])
    ref_col = _find_col(df, ["reference", "ref", "doc no", "document", "cheque", "voucher", "narration ref"])
    desc_col = _find_col(df, ["description", "narration", "details", "particulars", "memo"])

    debit_col = _find_col(df, ["debit", "dr", "withdrawal", "withdrawals", "payment", "payments"])
    credit_col = _find_col(df, ["credit", "cr", "deposit", "deposits", "receipt", "receipts"])
    amt_col = _find_col(df, ["amount", "value", "net amount", "signed amount"])

    debits: List[SideItem] = []
    credits: List[SideItem] = []

    for idx, row in df.iterrows():
        dt = _clean_str(row[date_col]) if date_col else ""
        ref = _clean_str(row[ref_col]) if ref_col else ""
        desc = _clean_str(row[desc_col]) if desc_col else ""

        if debit_col or credit_col:
            dr_val = abs(_to_float(row[debit_col])) if debit_col else 0.0
            cr_val = abs(_to_float(row[credit_col])) if credit_col else 0.0

            if dr_val > 0 and cr_val == 0:
                debits.append(SideItem(int(idx), source_label, "DEBIT", dr_val, dt, ref, desc))
            elif cr_val > 0 and dr_val == 0:
                credits.append(SideItem(int(idx), source_label, "CREDIT", cr_val, dt, ref, desc))
            elif dr_val > 0 and cr_val > 0:
                # If both populated on one row, treat net difference
                diff = cr_val - dr_val
                if diff > 0:
                    credits.append(SideItem(int(idx), source_label, "CREDIT", diff, dt, ref, desc))
                elif diff < 0:
                    debits.append(SideItem(int(idx), source_label, "DEBIT", abs(diff), dt, ref, desc))
        elif amt_col:
            val = _to_float(row[amt_col])
            if val < 0:
                debits.append(SideItem(int(idx), source_label, "DEBIT", abs(val), dt, ref, desc))
            elif val > 0:
                credits.append(SideItem(int(idx), source_label, "CREDIT", val, dt, ref, desc))

    debits.sort(key=lambda x: (x.amount, x.reference, x.date, x.row_id))
    credits.sort(key=lambda x: (x.amount, x.reference, x.date, x.row_id))
    return debits, credits


def _match_sorted_lists(
    bank_items: List[SideItem],
    erp_items: List[SideItem],
    pair_label: str,
) -> pd.DataFrame:
    """
    One-to-one matching on sorted lists using strict abs(bank.amount - erp.amount) < TOLERANCE.
    Ignores dates and descriptions when matching.
    Preserves unmatched items with blank opposite columns and Variance = Bank - ERP.
    """
    b_sorted = sorted(bank_items, key=lambda x: (x.amount, x.row_id))
    e_sorted = sorted(erp_items, key=lambda x: (x.amount, x.row_id))

    b_matched_to: Dict[int, int] = {}  # b_idx -> e_idx
    e_used: set = set()

    # Greedy nearest-within-tolerance pass in ascending bank order
    e_start = 0
    n_erp = len(e_sorted)

    for b_idx, b_item in enumerate(b_sorted):
        # Advance start pointer past items far smaller than current bank amount
        while e_start < n_erp and (b_item.amount - e_sorted[e_start].amount) >= TOLERANCE:
            e_start += 1

        best_e_idx: Optional[int] = None
        best_diff: float = TOLERANCE

        for e_idx in range(e_start, n_erp):
            e_item = e_sorted[e_idx]
            # Since e_sorted is ascending, once e_item.amount - b_item.amount >= TOLERANCE, no later ERP item fits
            if (e_item.amount - b_item.amount) >= TOLERANCE:
                break
            if e_idx in e_used:
                continue
            diff = abs(b_item.amount - e_item.amount)
            if diff < best_diff:
                best_diff = diff
                best_e_idx = e_idx
                if best_diff == 0.0:
                    break

        if best_e_idx is not None:
            b_matched_to[b_idx] = best_e_idx
            e_used.add(best_e_idx)

    # Build aligned rows preserving ascending order on each side without dropping unmatched entries
    rows = []
    b_ptr = 0
    e_ptr = 0

    while b_ptr < len(b_sorted) or e_ptr < len(e_sorted):
        b_item = b_sorted[b_ptr] if b_ptr < len(b_sorted) else None
        e_item = e_sorted[e_ptr] if e_ptr < len(e_sorted) else None

        # Skip already-consumed ERP items when advancing ERP pointer
        while e_item is not None and e_ptr in e_used:
            e_ptr += 1
            e_item = e_sorted[e_ptr] if e_ptr < len(e_sorted) else None

        if b_item is not None and b_ptr in b_matched_to:
            paired_e_idx = b_matched_to[b_ptr]
            paired_e = e_sorted[paired_e_idx]
            b_item.matched = True
            paired_e.matched = True
            variance = round(b_item.amount - paired_e.amount, 6)
            rows.append({
                "Pair": pair_label,
                "Status": "Matched",
                "Bank Ref": b_item.reference,
                "Bank Date": b_item.date,
                "Bank Description": b_item.description,
                "Bank Amount": b_item.amount,
                "Variance": variance,
                "ERP Amount": paired_e.amount,
                "ERP Description": paired_e.description,
                "ERP Date": paired_e.date,
                "ERP Ref": paired_e.reference,
            })
            b_ptr += 1

        elif b_item is not None and (e_item is None or b_item.amount <= e_item.amount):
            # Unmatched Bank item (Bank-only -> positive Bank Amount, Variance = Bank - 0 = +Bank)
            rows.append({
                "Pair": pair_label,
                "Status": "Bank Only",
                "Bank Ref": b_item.reference,
                "Bank Date": b_item.date,
                "Bank Description": b_item.description,
                "Bank Amount": b_item.amount,
                "Variance": b_item.amount,
                "ERP Amount": np.nan,
                "ERP Description": "",
                "ERP Date": "",
                "ERP Ref": "",
            })
            b_ptr += 1

        elif e_item is not None:
            # Unmatched ERP item (ERP-only -> ERP Amount shown, Variance = 0 - ERP = -ERP)
            rows.append({
                "Pair": pair_label,
                "Status": "ERP Only",
                "Bank Ref": "",
                "Bank Date": "",
                "Bank Description": "",
                "Bank Amount": np.nan,
                "Variance": -e_item.amount,
                "ERP Amount": e_item.amount,
                "ERP Description": e_item.description,
                "ERP Date": e_item.date,
                "ERP Ref": e_item.reference,
            })
            e_ptr += 1

    return pd.DataFrame(rows)


def _flag_investigation_items(items: List[SideItem], side_name: str, flow_name: str) -> pd.DataFrame:
    """
    Flags unmatched items and duplicate-amount clusters where at least one item is unmatched
    or multiple identical amounts exist.
    """
    amount_counts: Dict[float, int] = {}
    for it in items:
        amount_counts[it.amount] = amount_counts.get(it.amount, 0) + 1

    flagged = []
    for it in items:
        is_dup_amt = amount_counts[it.amount] > 1
        if not it.matched or is_dup_amt:
            reasons = []
            if not it.matched:
                reasons.append("Unmatched")
            if is_dup_amt:
                reasons.append(f"Duplicate amount ({amount_counts[it.amount]}x)")
            flagged.append({
                "Source": side_name,
                "Side": flow_name,
                "Matched": "Yes" if it.matched else "No",
                "Reason": " & ".join(reasons),
                "Amount": it.amount,
                "Date": it.date,
                "Reference": it.reference,
                "Description": it.description,
            })
    return pd.DataFrame(flagged)


def run_reconciliation(bank_df: pd.DataFrame, erp_df: pd.DataFrame) -> Dict[str, pd.DataFrame]:
    bank_dr, bank_cr = extract_items(bank_df, "BANK")
    erp_dr, erp_cr = extract_items(erp_df, "ERP")

    # Opposite accounting sides:
    # Bank Debits <-> ERP Credits
    # Bank Credits <-> ERP Debits
    dr_vs_cr_df = _match_sorted_lists(bank_dr, erp_cr, "Bank Debits vs ERP Credits")
    cr_vs_dr_df = _match_sorted_lists(bank_cr, erp_dr, "Bank Credits vs ERP Debits")

    combined_df = pd.concat([dr_vs_cr_df, cr_vs_dr_df], ignore_index=True)

    matched_df = combined_df[combined_df["Status"] == "Matched"].copy()
    bank_only_df = combined_df[combined_df["Status"] == "Bank Only"].copy()
    erp_only_df = combined_df[combined_df["Status"] == "ERP Only"].copy()

    all_side_items = bank_dr + bank_cr + erp_dr + erp_cr
    inv_parts = [
        _flag_investigation_items(bank_dr, "Bank", "Debit"),
        _flag_investigation_items(bank_cr, "Bank", "Credit"),
        _flag_investigation_items(erp_dr, "ERP", "Debit"),
        _flag_investigation_items(erp_cr, "ERP", "Credit"),
    ]
    investigation_df = pd.concat([p for p in inv_parts if not p.empty], ignore_index=True) if any(not p.empty for p in inv_parts) else pd.DataFrame(
        columns=["Source", "Side", "Matched", "Reason", "Amount", "Date", "Reference", "Description"]
    )

    total_bank = sum(x.amount for x in bank_dr) + sum(x.amount for x in bank_cr)
    total_erp = sum(x.amount for x in erp_dr) + sum(x.amount for x in erp_cr)
    matched_bank_amt = float(matched_df["Bank Amount"].sum()) if not matched_df.empty else 0.0
    matched_erp_amt = float(matched_df["ERP Amount"].sum()) if not matched_df.empty else 0.0
    outstanding_bank_amt = float(bank_only_df["Bank Amount"].sum()) if not bank_only_df.empty else 0.0
    outstanding_erp_amt = float(erp_only_df["ERP Amount"].sum()) if not erp_only_df.empty else 0.0
    total_variance = float(combined_df["Variance"].sum()) if not combined_df.empty else 0.0

    summary_df = pd.DataFrame([
        {"Metric": "Total Bank Transactions (Amount)", "Value": round(total_bank, 2)},
        {"Metric": "Total ERP Transactions (Amount)", "Value": round(total_erp, 2)},
        {"Metric": "Matched Bank Amount", "Value": round(matched_bank_amt, 2)},
        {"Metric": "Matched ERP Amount", "Value": round(matched_erp_amt, 2)},
        {"Metric": "Outstanding Bank Items (Green)", "Value": round(outstanding_bank_amt, 2)},
        {"Metric": "Outstanding ERP Items (Red)", "Value": round(outstanding_erp_amt, 2)},
        {"Metric": "Net Variance (Bank - ERP)", "Value": round(total_variance, 2)},
        {"Metric": "Matched Transaction Count", "Value": int(len(matched_df))},
        {"Metric": "Bank-Only Transaction Count", "Value": int(len(bank_only_df))},
        {"Metric": "ERP-Only Transaction Count", "Value": int(len(erp_only_df))},
        {"Metric": "Items Flagged for Investigation", "Value": int(len(investigation_df))},
    ])

    return {
        "Bank Debits vs ERP Credits": dr_vs_cr_df,
        "Bank Credits vs ERP Debits": cr_vs_dr_df,
        "Matched Transactions": matched_df,
        "Bank-Only Transactions": bank_only_df,
        "ERP-Only Transactions": erp_only_df,
        "Investigation & Duplicates": investigation_df,
        "Reconciliation Summary": summary_df,
        "All Side Items": pd.DataFrame([vars(x) for x in all_side_items]),
    }


def _write_styled_sheet(ws, df: pd.DataFrame, highlight_status: bool = False) -> None:
    green_fill = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")
    red_fill = PatternFill(start_color="F4CCCC", end_color="F4CCCC", fill_type="solid")
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)

    if df.empty:
        ws.append(["No records in this section."])
        return

    cols = list(df.columns)
    ws.append(cols)
    for c_idx in range(1, len(cols) + 1):
        cell = ws.cell(row=1, column=c_idx)
        cell.fill = header_fill
        cell.font = header_font

    status_col_idx = cols.index("Status") + 1 if "Status" in cols else None

    for _, r in df.iterrows():
        row_vals = [None if (isinstance(r[c], float) and np.isnan(r[c])) else r[c] for c in cols]
        ws.append(row_vals)
        current_row = ws.max_row

        row_fill = None
        if highlight_status and status_col_idx:
            st = str(r["Status"])
            if st == "Bank Only":
                row_fill = green_fill
            elif st == "ERP Only":
                row_fill = red_fill

        if row_fill:
            for c_idx in range(1, len(cols) + 1):
                ws.cell(row=current_row, column=c_idx).fill = row_fill

    # Format numeric columns without destroying exact precision
    num_headers = {"Bank Amount", "Variance", "ERP Amount", "Amount", "Value"}
    for c_idx, col_name in enumerate(cols, start=1):
        if col_name in num_headers:
            for r_idx in range(2, ws.max_row + 1):
                cell = ws.cell(row=r_idx, column=c_idx)
                if isinstance(cell.value, (int, float)):
                    cell.number_format = "#,##0.00"

    # Auto-fit column widths with a sensible cap
    for c_idx, col_name in enumerate(cols, start=1):
        max_len = len(str(col_name))
        for r_idx in range(2, min(ws.max_row + 1, 200)):
            val_str = str(ws.cell(row=r_idx, column=c_idx).value or "")
            max_len = max(max_len, len(val_str))
        ws.column_dimensions[get_column_letter(c_idx)].width = min(max_len + 4, 42)


def export_to_excel(results: Dict[str, pd.DataFrame]) -> bytes:
    wb = Workbook()
    # Remove default sheet
    wb.remove(wb.active)

    sheet_order = [
        ("Bank Debits vs ERP Credits", True),
        ("Bank Credits vs ERP Debits", True),
        ("Matched Transactions", False),
        ("Bank-Only Transactions", True),
        ("ERP-Only Transactions", True),
        ("Investigation & Duplicates", False),
        ("Reconciliation Summary", False),
    ]

    for sheet_name, do_highlight in sheet_order:
        ws = wb.create_sheet(title=sheet_name[:31])
        df = results.get(sheet_name, pd.DataFrame())
        _write_styled_sheet(ws, df, highlight_status=do_highlight)

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()
