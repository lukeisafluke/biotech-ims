"""
Database & File System Backup & Synchronization Framework
=============================================================================
Provides automated, incremental and archived synchronization of:
1. PostgreSQL database (tables, schema, data) to SQL restore script
2. Generated document folders (invoices, quotes, packing slips, shipping labels)
3. Excel templates and system configuration
4. External drive auto-detection and retention pruning
"""

import os
import sys
import json
import shutil
import string
import ctypes
import argparse
from datetime import datetime
import psycopg2
from psycopg2.extras import RealDictCursor

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DOCS_DIR = os.path.join(BASE_DIR, 'generated_docs')
TEMPLATES_DIR = os.path.join(BASE_DIR, 'excel_templates')
HISTORY_FILE = os.path.join(BASE_DIR, '.backup_history.json')

# DB configuration (inherits from environment or defaults matching app.py)
DB_HOST = os.environ.get("DB_HOST", "localhost")
DB_PORT = os.environ.get("DB_PORT", "5432")
DB_NAME = os.environ.get("DB_NAME", "inventory_db")
DB_USER = os.environ.get("DB_USER", "postgres")
DB_PASS = os.environ.get("DB_PASS", "postgres")

TABLES_ORDERED = [
    'customers',
    'products',
    'product_lots',
    'quotes',
    'quote_items',
    'orders',
    'order_items'
]

def get_db_connection():
    return psycopg2.connect(
        dbname=DB_NAME,
        user=DB_USER,
        password=DB_PASS,
        host=DB_HOST,
        port=DB_PORT
    )

def detect_available_drives():
    """Scans Windows drive letters for removable and external drives."""
    available = []
    if sys.platform == 'win32':
        kernel32 = ctypes.windll.kernel32
        type_names = {
            2: 'Removable Drive (USB)',
            3: 'Fixed Drive (External/Secondary)',
            4: 'Network Share Drive'
        }
        for letter in string.ascii_uppercase:
            root = f"{letter}:\\"
            dtype = kernel32.GetDriveTypeW(root)
            if dtype in type_names and letter != 'C':  # Skip primary system drive C:
                # Check write permissions
                test_file = os.path.join(root, '.backup_write_test')
                writable = False
                try:
                    with open(test_file, 'w') as f:
                        f.write('ok')
                    os.remove(test_file)
                    writable = True
                except Exception:
                    writable = False

                available.append({
                    'letter': letter,
                    'root': root,
                    'type': type_names[dtype],
                    'writable': writable,
                    'suggested_path': os.path.join(root, 'System_Backups')
                })
    return available

def dump_database_sql(out_filepath):
    """Generates a complete standalone SQL dump file with all schema and data."""
    conn = get_db_connection()
    cur = conn.cursor(cursor_factory=RealDictCursor)

    stats = {}
    lines = [
        "-- =============================================================================",
        f"-- Database PostgreSQL Backup Dump",
        f"-- Generated At: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        f"-- Database: {DB_NAME} on {DB_HOST}:{DB_PORT}",
        "-- =============================================================================\n",
        "BEGIN;\n",
        "SET client_encoding = 'UTF8';\n",
        "SET check_function_bodies = false;\n",
    ]

    for table in TABLES_ORDERED:
        # Check if table exists
        cur.execute("""
            SELECT EXISTS (
                SELECT FROM information_schema.tables 
                WHERE table_schema = 'public' AND table_name = %s
            )
        """, (table,))
        exists = cur.fetchone()['exists']
        if not exists:
            continue

        # Get columns
        cur.execute("""
            SELECT column_name, data_type 
            FROM information_schema.columns 
            WHERE table_schema = 'public' AND table_name = %s
            ORDER BY ordinal_position
        """, (table,))
        cols_info = cur.fetchall()
        cols = [c['column_name'] for c in cols_info]

        # Fetch all rows
        cur.execute(f"SELECT * FROM {table}")
        rows = cur.fetchall()
        stats[table] = len(rows)

        lines.append(f"-- Table: {table} ({len(rows)} records)")
        lines.append(f"TRUNCATE TABLE {table} RESTART IDENTITY CASCADE;\n")

        for r in rows:
            val_strs = []
            for col in cols:
                v = r[col]
                if v is None:
                    val_strs.append("NULL")
                elif isinstance(v, (int, float)):
                    val_strs.append(str(v))
                elif isinstance(v, bool):
                    val_strs.append("TRUE" if v else "FALSE")
                else:
                    # Escape single quotes and backslashes
                    s = str(v).replace("'", "''")
                    val_strs.append(f"'{s}'")

            col_list_str = ", ".join(f'"{c}"' for c in cols)
            val_list_str = ", ".join(val_strs)
            lines.append(f'INSERT INTO "{table}" ({col_list_str}) VALUES ({val_list_str});')

        lines.append("")

    lines.append("COMMIT;\n")
    lines.append(f"-- Dump completed: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")

    cur.close()
    conn.close()

    os.makedirs(os.path.dirname(os.path.abspath(out_filepath)), exist_ok=True)
    with open(out_filepath, 'w', encoding='utf-8') as f:
        f.write("\n".join(lines))

    return stats

def sync_directory_tree(src, dst):
    """Mirror copies files from src to dst incrementally, returning count and bytes."""
    if not os.path.exists(src):
        return 0, 0

    os.makedirs(dst, exist_ok=True)
    count = 0
    total_bytes = 0

    for root, dirs, files in os.walk(src):
        rel = os.path.relpath(root, src)
        dest_root = os.path.join(dst, rel) if rel != '.' else dst
        os.makedirs(dest_root, exist_ok=True)

        for f in files:
            src_file = os.path.join(root, f)
            dst_file = os.path.join(dest_root, f)

            copy_needed = True
            if os.path.exists(dst_file):
                src_stat = os.stat(src_file)
                dst_stat = os.stat(dst_file)
                if src_stat.st_size == dst_stat.st_size and abs(src_stat.st_mtime - dst_stat.st_mtime) < 1.0:
                    copy_needed = False

            if copy_needed:
                shutil.copy2(src_file, dst_file)

            stat = os.stat(src_file)
            count += 1
            total_bytes += stat.st_size

    return count, total_bytes

def prune_old_archives(archives_dir, max_keep=10):
    """Keeps the latest max_keep archives in the archives directory, removing older ones."""
    if not os.path.exists(archives_dir):
        return []
    entries = [os.path.join(archives_dir, d) for d in os.listdir(archives_dir) if os.path.isdir(os.path.join(archives_dir, d))]
    entries.sort(key=lambda p: os.path.getmtime(p), reverse=True)
    
    removed = []
    if len(entries) > max_keep:
        for old_dir in entries[max_keep:]:
            try:
                shutil.rmtree(old_dir)
                removed.append(os.path.basename(old_dir))
            except Exception:
                pass
    return removed

def perform_backup(target_root=None, max_archives=10):
    """
    Executes a complete server synchronization and backup to target_root.
    If target_root is None, attempts to detect a connected external drive.
    """
    start_time = datetime.now()

    if not target_root:
        # Check environment variable
        target_root = os.environ.get("BACKUP_TARGET") or os.environ.get("TACT_BACKUP_TARGET")
    if not target_root:
        drives = detect_available_drives()
        writable_drives = [d for d in drives if d['writable']]
        if writable_drives:
            target_root = writable_drives[0]['suggested_path']
        else:
            # Fallback to local backup folder if no external drive is plugged in
            target_root = os.path.join(BASE_DIR, 'local_backups')

    target_root = os.path.abspath(target_root)
    os.makedirs(target_root, exist_ok=True)

    timestamp_str = start_time.strftime('%Y%m%d_%H%M%S')
    archive_dir = os.path.join(target_root, 'archives', f'backup_{timestamp_str}')
    latest_dir = os.path.join(target_root, 'latest')

    os.makedirs(archive_dir, exist_ok=True)
    os.makedirs(latest_dir, exist_ok=True)

    # 1. Dump database to SQL
    db_dump_archive = os.path.join(archive_dir, 'database_dump.sql')
    db_dump_latest = os.path.join(latest_dir, 'database_dump.sql')
    db_stats = dump_database_sql(db_dump_archive)
    shutil.copy2(db_dump_archive, db_dump_latest)

    # 2. Sync generated documents (preserving folder structure)
    docs_archive = os.path.join(archive_dir, 'generated_docs')
    docs_latest = os.path.join(latest_dir, 'generated_docs')
    docs_count, docs_bytes = sync_directory_tree(DOCS_DIR, docs_latest)
    sync_directory_tree(DOCS_DIR, docs_archive)

    # 3. Sync Excel templates
    templates_archive = os.path.join(archive_dir, 'excel_templates')
    templates_latest = os.path.join(latest_dir, 'excel_templates')
    tmpl_count, tmpl_bytes = sync_directory_tree(TEMPLATES_DIR, templates_latest)
    sync_directory_tree(TEMPLATES_DIR, templates_archive)

    # 4. Copy schema.sql and app configuration
    schema_file = os.path.join(BASE_DIR, 'schema.sql')
    if os.path.exists(schema_file):
        shutil.copy2(schema_file, os.path.join(archive_dir, 'schema.sql'))
        shutil.copy2(schema_file, os.path.join(latest_dir, 'schema.sql'))

    # 5. Build manifest
    end_time = datetime.now()
    duration = (end_time - start_time).total_seconds()
    total_size_mb = round((os.path.getsize(db_dump_archive) + docs_bytes + tmpl_bytes) / (1024 * 1024), 2)

    manifest = {
        'status': 'SUCCESS',
        'timestamp': start_time.strftime('%Y-%m-%d %H:%M:%S'),
        'duration_seconds': round(duration, 2),
        'target_directory': target_root,
        'archive_folder': os.path.basename(archive_dir),
        'database_records': db_stats,
        'documents_synced': docs_count,
        'templates_synced': tmpl_count,
        'total_backup_size_mb': total_size_mb
    }

    manifest_path = os.path.join(archive_dir, 'manifest.json')
    with open(manifest_path, 'w', encoding='utf-8') as f:
        json.dump(manifest, f, indent=2)
    shutil.copy2(manifest_path, os.path.join(latest_dir, 'manifest.json'))

    # 6. Prune old archives
    pruned = prune_old_archives(os.path.join(target_root, 'archives'), max_keep=max_archives)
    manifest['pruned_old_archives'] = pruned

    # 7. Record to history
    record_backup_history(manifest)

    return manifest

def record_backup_history(entry):
    history = []
    if os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE, 'r', encoding='utf-8') as f:
                history = json.load(f)
        except Exception:
            history = []
    history.insert(0, entry)
    history = history[:20]  # Keep last 20 entries
    try:
        with open(HISTORY_FILE, 'w', encoding='utf-8') as f:
            json.dump(history, f, indent=2)
    except Exception:
        pass

def get_backup_history():
    if os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            pass
    return []

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="Server & External Drive Backup Utility")
    parser.add_argument('--target', type=str, help="Target backup folder on external hard drive")
    parser.add_argument('--list-drives', action='store_true', help="List detected removable and external drives")
    args = parser.parse_args()

    if args.list_drives:
        drives = detect_available_drives()
        print(f"Detected Drives: {len(drives)}")
        for d in drives:
            print(f"  {d['letter']}: [{d['type']}] Writable: {d['writable']} (Suggested: {d['suggested_path']})")
        sys.exit(0)

    print("Beginning Server Synchronization & Backup...")
    result = perform_backup(target_root=args.target)
    print("\nBackup Completed Successfully!")
    print(f"  Target: {result['target_directory']}")
    print(f"  Archive: {result['archive_folder']}")
    print(f"  Size: {result['total_backup_size_mb']} MB")
    print(f"  Duration: {result['duration_seconds']}s")
    print(f"  DB Records: {result['database_records']}")

