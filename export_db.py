"""
Database Data Export Utility
Exports all current data from PostgreSQL as SQL INSERT statements.
Usage: python export_db.py [output_file.sql]
"""

import sys
import os
import datetime
from decimal import Decimal

sys.path.insert(0, os.path.dirname(__file__))
from app import get_db

def sql_quote(val):
    if val is None:
        return 'NULL'
    if isinstance(val, (int, float)):
        return str(val)
    if isinstance(val, Decimal):
        return str(val)
    if isinstance(val, (datetime.date, datetime.datetime)):
        return f"'{val.isoformat()}'"
    if isinstance(val, bool):
        return 'TRUE' if val else 'FALSE'
    s = str(val).replace("'", "''")
    return f"'{s}'"

def export():
    output_filename = sys.argv[1] if len(sys.argv) > 1 else "backup_data.sql"
    output_path = os.path.join(os.path.dirname(__file__), output_filename)
    
    tables = ['customers', 'products', 'product_lots', 'quotes', 'quote_items', 'orders', 'order_items']
    
    conn = get_db()
    cur = conn.cursor()
    
    lines = [
        "-- =============================================================================",
        f"-- Database Data Dump: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        "-- =============================================================================",
        "BEGIN;\n"
    ]
    
    for table in tables:
        cur.execute(f"SELECT * FROM {table};")
        rows = cur.fetchall()
        col_names = [desc[0] for desc in cur.description]
        
        lines.append(f"-- Table: {table} ({len(rows)} records)")
        for row in rows:
            cols_str = ", ".join(col_names)
            vals_str = ", ".join(sql_quote(v) for v in row)
            lines.append(f"INSERT INTO {table} ({cols_str}) VALUES ({vals_str}) ON CONFLICT DO NOTHING;")
        
        # Reset serial sequences to max id + 1
        pk_col = f"{table[:-1] if table.endswith('s') and not table.endswith('ss') else table}_id"
        if table == 'product_lots':
            pk_col = 'lot_id'
        seq_name = f"{table}_{pk_col}_seq"
        lines.append(f"SELECT setval('{seq_name}', COALESCE((SELECT MAX({pk_col}) FROM {table}), 1), true);\n")
        
    lines.append("COMMIT;\n")
    
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write("\n".join(lines))
        
    print(f"Successfully exported {len(tables)} tables to {output_path}")
    cur.close()
    conn.close()

if __name__ == '__main__':
    export()

