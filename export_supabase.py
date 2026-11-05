
import os
import json
from pathlib import Path

from dotenv import load_dotenv
from supabase import create_client
from openpyxl import load_workbook


# ============================================================
# CONFIGURATION
# ============================================================

load_dotenv()

SUPABASE_URL = os.getenv("SUPABASE_URL")
SUPABASE_KEY = os.getenv("SUPABASE_KEY")

TABLE_NAME = "part_numbers"

EXCEL_FILE = Path("Control db_BOM.xlsx")
SHEET_NAME = "Part_Numbers"

BATCH_SIZE = 1000


# ============================================================
# SUPABASE CONNECTION
# ============================================================

def connect_supabase():
    if not SUPABASE_URL or not SUPABASE_KEY:
        raise ValueError("Missing Supabase credentials.")

    return create_client(SUPABASE_URL, SUPABASE_KEY)


# ============================================================
# DATA EXTRACTION
# ============================================================

def fetch_table_data(supabase):
    records = []
    offset = 0

    while True:
        response = (
            supabase
            .table(TABLE_NAME)
            .select("*")
            .range(offset, offset + BATCH_SIZE - 1)
            .execute()
        )

        batch = response.data

        if not batch:
            break

        records.extend(batch)
        offset += len(batch)

        if len(batch) < BATCH_SIZE:
            break

    return records


# ============================================================
# EXCEL EXPORT
# ============================================================

def export_to_excel(records):

    if not records:
        print("No records found.")
        return

    workbook = load_workbook(EXCEL_FILE)

    if SHEET_NAME not in workbook.sheetnames:
        raise ValueError(f"Sheet '{SHEET_NAME}' not found.")

    worksheet = workbook[SHEET_NAME]

    # Clear existing cell values
    for row in worksheet.iter_rows():
        for cell in row:
            cell.value = None

    # Get column names
    headers = list(records[0].keys())

    # Write headers
    for col, header in enumerate(headers, start=1):
        worksheet.cell(row=1, column=col, value=header)

    # Write records
    for row_idx, record in enumerate(records, start=2):
        for col_idx, header in enumerate(headers, start=1):

            value = record.get(header)

            if isinstance(value, (dict, list)):
                value = json.dumps(value, ensure_ascii=False)

            worksheet.cell(
                row=row_idx,
                column=col_idx,
                value=value
            )

    workbook.save(EXCEL_FILE)

    print(f"Export completed: {len(records)} records.")


# ============================================================
# MAIN
# ============================================================

def main():

    supabase = connect_supabase()

    records = fetch_table_data(supabase)

    export_to_excel(records)


if __name__ == "__main__":
    main()
