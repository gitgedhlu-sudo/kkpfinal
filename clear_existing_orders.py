"""Safely clear all existing KKP Fireworks orders while preserving stock.

Run from the directory containing app.py and products.json:
    python clear_existing_orders.py --confirm-clear-orders

The script backs up the SQLite database first, deletes rows from the orders
 table, removes user/shared order receipt files, and regenerates an empty
 all_orders.xlsx report. Inventory and order_sequence tables are preserved.
"""
from datetime import datetime
from pathlib import Path
import shutil
import sys

if "--confirm-clear-orders" not in sys.argv:
    print("No changes made. To confirm, run:")
    print("  python clear_existing_orders.py --confirm-clear-orders")
    raise SystemExit(2)

import app  # Uses the same ORDER_DATA_DIR and database settings as the website.

backup_stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
backup_path = app.DATA_DIR / f"orders_backup_before_clear_{backup_stamp}.sqlite3"

if not app.DB_FILE.exists():
    print(f"Database not found: {app.DB_FILE}")
    raise SystemExit(1)

# Snapshot the database before making any destructive change.
shutil.copy2(app.DB_FILE, backup_path)

with app.db_connection() as db:
    old_count = db.execute("SELECT COUNT(*) FROM orders").fetchone()[0]
    db.execute("BEGIN IMMEDIATE")
    db.execute("DELETE FROM orders")
    db.commit()

# Remove every per-user receipt folder/file, but leave the orders directory itself.
removed_user_files = 0
for item in app.ORDERS_DIR.iterdir():
    if item.is_dir():
        removed_user_files += sum(1 for p in item.rglob("*") if p.is_file())
        shutil.rmtree(item)
    elif item.is_file():
        item.unlink()
        removed_user_files += 1

# Remove shared order PDFs/Excels and the old aggregate workbook; preserve unrelated directories.
removed_shared_files = 0
for item in app.ALL_ORDERS_DIR.iterdir():
    if item.is_file() and (item.suffix.lower() in {".pdf", ".xlsx", ".xlsm"}):
        item.unlink()
        removed_shared_files += 1

# Rebuild an empty workbook using the app's existing report format.
app.rebuild_all_orders_workbook()

print("Order cleanup completed successfully.")
print(f"Orders deleted: {old_count}")
print(f"Per-user receipt files removed: {removed_user_files}")
print(f"Shared receipt/report files removed: {removed_shared_files}")
print(f"Database backup: {backup_path}")
print("Inventory stock and order-number sequence were not changed.")
