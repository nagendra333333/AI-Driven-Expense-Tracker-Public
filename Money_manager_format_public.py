"""
Money Manager Format Converter
================================
Reads the latest  Full_statement_*.xlsx  from  OUTPUT_FOLDER/
Converts it to Money Manager .tsv format
Saves to  Money_Manager_output/  in the same folder as this script

TSV Column Order (exact Money Manager format):
  Date | Account | Category | Subcategory | Note | INR |
  Income/Expense | Description | Amount | Currency | Account

Field Mapping:
  Date           → Full_statement.Date  (formatted MM/DD/YYYY)
  Account        → always "Card"
  Category       → Expenses Category  (Debit rows)
                   Income Category    (Credit rows)
  Subcategory    → always empty
  Note           → Expenses sub-category  (Debit rows)
                   Income sub-category    (Credit rows)
  INR            → Debit or Credit amount
  Income/Expense → "Expense" (Debit) | "Income" (Credit)
  Description    → always empty
  Amount         → same as INR
  Currency       → INR
  Account (last) → same as INR amount.
"""

import sys
import re
from pathlib import Path
from datetime import datetime

import pandas as pd

# ─── FOLDERS ─────────────────────────────────────────────────────────────────
SCRIPT_DIR          = Path(__file__).resolve().parent
OUTPUT_FOLDER       = SCRIPT_DIR / "OUTPUT_FOLDER"
MONEY_MANAGER_OUT   = SCRIPT_DIR / "Money_Manager_output"

# ─── HELPERS ─────────────────────────────────────────────────────────────────

def find_latest_statement(folder: Path) -> Path | None:
    """Return the most recent Full_statement_*.xlsx by filename date."""
    files = sorted(folder.glob("Full_statement_*.xlsx"))
    return files[-1] if files else None


def parse_date(raw) -> datetime | None:
    """Parse any date format found in Full_statement Date column."""
    s = str(raw).strip()
    if not s or s.lower() in ("nan", "nat", "none", ""):
        return None
    for fmt in (
        "%d-%m-%Y",   # 01-03-2026
        "%d/%m/%Y",   # 01/03/2026
        "%Y-%m-%d",   # 2026-03-01
        "%d/%m/%y",   # 01/03/26
        "%d-%m-%y",   # 01-03-26
        "%m/%d/%Y",   # 03/01/2026
        "%d %b %Y",   # 01 Mar 2026
    ):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None


def fmt_date(raw) -> str:
    """Return date as MM/DD/YYYY for Money Manager."""
    dt = parse_date(raw)
    return dt.strftime("%m/%d/%Y") if dt else str(raw).strip()


def clean(val) -> str:
    """Return empty string for NaN/None/nan, else stripped string."""
    s = str(val).strip()
    return "" if s.lower() in ("nan", "none", "nat", "") else s


def to_num(val) -> float | None:
    """Convert amount string to float."""
    s = str(val).replace(",", "").strip()
    try:
        v = float(s)
        return v if v > 0 else None
    except (ValueError, TypeError):
        return None


def fmt_amount(v: float) -> str:
    """Format as integer string if whole number, else 2 decimal places."""
    return str(int(v)) if v == int(v) else f"{v:.2f}"


# ─── MAIN ─────────────────────────────────────────────────────────────────────

def main():
    print("=" * 60)
    print("  Money Manager Format Converter")
    print(f"  Source : {OUTPUT_FOLDER}")
    print(f"  Output : {MONEY_MANAGER_OUT}")
    print("=" * 60)

    # ── Find latest Full_statement file ──────────────────────────────────────
    latest = find_latest_statement(OUTPUT_FOLDER)
    if latest is None:
        print(f"\n[ERROR] No Full_statement_*.xlsx found in:\n        {OUTPUT_FOLDER}")
        sys.exit(1)
    print(f"\n  Reading: {latest.name}")

    # ── Load Transactions sheet ───────────────────────────────────────────────
    try:
        df = pd.read_excel(str(latest), sheet_name="Transactions", dtype=str)
    except Exception as e:
        print(f"[ERROR] Could not read {latest.name}: {e}")
        sys.exit(1)

    print(f"  Rows loaded: {len(df)}")

    # ── Normalise columns (strip whitespace, handle missing) ──────────────────
    df.columns = [c.strip() for c in df.columns]
    for col in ["Debit", "Credit", "Expenses Category", "Expenses sub-category",
                "Income Category", "Income sub-category", "Date"]:
        if col not in df.columns:
            df[col] = ""

    # ── Build TSV rows ────────────────────────────────────────────────────────
    tsv_rows = []

    for _, row in df.iterrows():
        debit  = to_num(row.get("Debit",  ""))
        credit = to_num(row.get("Credit", ""))

        # Skip rows with no monetary value at all
        if debit is None and credit is None:
            continue

        date_str = fmt_date(row["Date"])

        # ── Expense row (Debit) ───────────────────────────────────────────────
        if debit is not None and debit > 0:
            category = clean(row.get("Expenses Category", ""))
            note     = clean(row.get("Expenses sub-category", ""))
            tsv_rows.append({
                "Date"           : date_str,
                "Account"        : "Card",
                "Category"       : category,
                "Subcategory"    : "",
                "Note"           : note,
                "INR"            : fmt_amount(debit),
                "Income/Expense" : "Expense",
                "Description"    : "",
                "Amount"         : fmt_amount(debit),
                "Currency"       : "INR",
                "Account2"       : fmt_amount(debit),
            })

        # ── Income row (Credit) ───────────────────────────────────────────────
        if credit is not None and credit > 0:
            category = clean(row.get("Income Category", ""))
            note     = clean(row.get("Income sub-category", ""))
            tsv_rows.append({
                "Date"           : date_str,
                "Account"        : "Card",
                "Category"       : category,
                "Subcategory"    : "",
                "Note"           : note,
                "INR"            : fmt_amount(credit),
                "Income/Expense" : "Income",
                "Description"    : "",
                "Amount"         : fmt_amount(credit),
                "Currency"       : "INR",
                "Account2"       : fmt_amount(credit),
            })

    if not tsv_rows:
        print("\n[ERROR] No valid rows to export.")
        sys.exit(1)

    out_df = pd.DataFrame(tsv_rows, columns=[
        "Date", "Account", "Category", "Subcategory", "Note",
        "INR", "Income/Expense", "Description", "Amount", "Currency", "Account2"
    ])

    # Rename Account2 → Account (same header name as original, duplicate col)
    out_df.rename(columns={"Account2": "Account"}, inplace=True)

    # ── Determine output filename from latest statement date ──────────────────
    stmt_date = latest.stem.replace("Full_statement_", "")   # e.g. 2026-03-23
    MONEY_MANAGER_OUT.mkdir(exist_ok=True)
    out_path = MONEY_MANAGER_OUT / f"Money_Manager_{stmt_date}.tsv"

    # ── Write TSV ─────────────────────────────────────────────────────────────
    out_df.to_csv(str(out_path), sep="\t", index=False)

    print(f"\n  Expense rows : {len(out_df[out_df['Income/Expense'] == 'Expense'])}")
    print(f"  Income  rows : {len(out_df[out_df['Income/Expense'] == 'Income'])}")
    print(f"  Total   rows : {len(out_df)}")
    print(f"\n  ✅  Saved → {out_path.name}")
    print("=" * 60)


if __name__ == "__main__":
    main()
