from datetime import date
from decimal import Decimal
from pathlib import Path

from invoice_splitter.config import get_settings
from invoice_splitter.excel.writer import apply_transaction


def run():
    s = get_settings()  # usa EXCEL_PATH del .env [3](https://stackoverflow.com/questions/63182578/renaming-excel-sheets-name-exceeds-31-characters-error)
    excel_path = s.excel_path

    # Tabla que NO existe aún (para forzar creación)
    table_name = "Surti_table"

    vendor_id = 9999999
    bill = "000000001"  # 9 dígitos tipo texto (como tu Excel) [2](https://support.microsoft.com/en-us/office/excel-doesn-t-fully-support-some-special-characters-in-the-filename-or-folder-path-20728217-f08a-4d63-a741-821a14cec380)[2](https://support.microsoft.com/en-us/office/excel-doesn-t-fully-support-some-special-characters-in-the-filename-or-folder-path-20728217-f08a-4d63-a741-821a14cec380)

    # Debe coincidir con headers base del writer [2](https://support.microsoft.com/en-us/office/excel-doesn-t-fully-support-some-special-characters-in-the-filename-or-folder-path-20728217-f08a-4d63-a741-821a14cec380)[2](https://support.microsoft.com/en-us/office/excel-doesn-t-fully-support-some-special-characters-in-the-filename-or-folder-path-20728217-f08a-4d63-a741-821a14cec380)
    row = {
        "Date": date.today(),
        "Bill number": bill,
        "ID": vendor_id,
        "Vendor": "SURTI",
        "Service/ concept": "Test auto-create table",
        "CC": 1100036,
        "GL account": 7418000000,
        "Subtotal assigned by CC": Decimal("100.00"),
        "% IVA": Decimal("0.15"),
        "IVA assigned by CC": Decimal("15.00"),
        "Total assigned by CC": Decimal("115.00"),
    }

    backup_dir = excel_path.parent / "invoice_splitter_backups"

    backup_path, deleted_by_table, backup_created = apply_transaction(
        excel_path=excel_path,
        backup_dir=backup_dir,
        vendor_id=vendor_id,
        bill_number=bill,
        table_to_rows={table_name: [row]},
        backup_path=None,
        retention_keep_last_n=30,
        retention_keep_days=30,
    )  # lógica de backup/retención/transacción [2](https://support.microsoft.com/en-us/office/excel-doesn-t-fully-support-some-special-characters-in-the-filename-or-folder-path-20728217-f08a-4d63-a741-821a14cec380)

    print("OK")
    print("Backup:", backup_path)
    print("Deleted:", deleted_by_table)
    print("Backup created now?:", backup_created)


if __name__ == "__main__":
    run()
