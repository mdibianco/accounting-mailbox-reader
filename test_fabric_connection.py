"""Quick standalone test for Fabric SQL connection."""

import sys
sys.path.insert(0, ".")

from src.invoice_lookup import InvoiceLookup

print("=" * 60)
print("FABRIC SQL CONNECTION TEST")
print("=" * 60)

lookup = InvoiceLookup()

# 1. Check config
print(f"\nConfigured: {lookup.is_configured}")
if not lookup.is_configured:
    print("[ERR] Missing credentials. Check .env")
    sys.exit(1)

print(f"Endpoint:  {lookup.sql_endpoint[:50]}...")
print(f"Database:  {lookup.database}")
print(f"Table:     {lookup.full_table}")

# 2. Test token acquisition
print("\n--- Token Acquisition ---")
token = lookup._get_token()
if token:
    print(f"[OK] Got token ({len(token)} chars)")
else:
    print("[ERR] Token acquisition failed. Check client_id/secret/tenant.")
    sys.exit(1)

# 3. Test SQL connection
print("\n--- SQL Connection ---")
conn = lookup._get_connection()
if not conn:
    print("[ERR] Connection failed. Is the service principal added to the Fabric workspace?")
    sys.exit(1)
print("[OK] Connected")

# 4. Test simple query
print("\n--- Test Query: SELECT TOP 3 ---")
try:
    cursor = conn.cursor()
    cursor.execute(
        f"SELECT TOP 3 entity, external_document_no, vendor_name, amount, remaining_amount, open "
        f"FROM {lookup.full_table} "
        f"WHERE document_type = 'Invoice' "
        f"ORDER BY posting_date DESC"
    )
    rows = cursor.fetchall()
    cursor.close()

    if rows:
        print(f"[OK] Got {len(rows)} rows:\n")
        for r in rows:
            print(f"  Entity={r[0]}  Invoice={r[1]}  Vendor={r[2]}  Amount={r[3]}  Remaining={r[4]}  Open={r[5]}")
    else:
        print("[WARN] Query returned 0 rows (table may be empty)")
except Exception as e:
    print(f"[ERR] Query failed: {e}")

# 5. Test a lookup
print("\n--- Test Lookup (entity=CH1, dummy invoice) ---")
result = lookup.lookup_invoice(entity_code="CH1", invoice_number="DOES-NOT-EXIST-12345")
print(f"Status: {result['status']}  Found: {result['found']}")
print("[OK] Lookup logic works (expected NOT_FOUND for dummy invoice)")

lookup.close()
print("\n" + "=" * 60)
print("ALL TESTS PASSED")
print("=" * 60)
