import os
import shutil
import re
import math
from datetime import datetime, date, timedelta
import psycopg2
from psycopg2.extras import RealDictCursor
from flask import Flask, render_template, request, redirect, url_for, send_from_directory, flash, jsonify, abort, session
from werkzeug.security import generate_password_hash, check_password_hash
import copy
import openpyxl
from openpyxl.styles import Border, Side
import backup_manager

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "dev_secret_key_change_in_production")
app.permanent_session_lifetime = timedelta(days=14)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_DIR = os.path.join(BASE_DIR, 'excel_templates')
OUTPUT_DIR = os.path.join(BASE_DIR, 'generated_docs')
PASSWORD_FILE = os.path.join(BASE_DIR, '.app_password')
os.makedirs(OUTPUT_DIR, exist_ok=True)

# -------------------------------------------------------------------------
# AUTHENTICATION HELPERS
# -------------------------------------------------------------------------
def get_system_password():
    if os.environ.get("APP_PASSWORD"):
        return os.environ.get("APP_PASSWORD").strip()
    if os.environ.get("TACT_APP_PASSWORD"):
        return os.environ.get("TACT_APP_PASSWORD").strip()
    if os.path.exists(PASSWORD_FILE):
        try:
            with open(PASSWORD_FILE, 'r', encoding='utf-8') as f:
                c = f.read().strip()
                if c:
                    return c
        except Exception:
            pass
    return "admin"

def verify_system_password(entered):
    if not entered:
        return False
    stored = get_system_password()
    if entered == stored:
        return True
    try:
        return check_password_hash(stored, entered)
    except Exception:
        return False

@app.before_request
def require_authentication():
    # Allow login, static files, and favicon without authentication
    if request.endpoint in ('login', 'static') or (request.path and request.path.startswith('/static')):
        return
    if not session.get('authenticated'):
        if request.path.startswith('/api/'):
            return jsonify({'error': 'Authentication required', 'login_url': url_for('login')}), 401
        return redirect(url_for('login', next=request.full_path if request.method == 'GET' else request.path))

@app.route('/login', methods=['GET', 'POST'])
def login():
    if session.get('authenticated'):
        return redirect(url_for('index'))

    if request.method == 'POST':
        entered_pass = request.form.get('password', '').strip()
        if verify_system_password(entered_pass):
            session['authenticated'] = True
            session.permanent = True
            flash("Welcome! Signed in successfully.", "success")
            next_url = request.args.get('next')
            if next_url and next_url.startswith('/') and not next_url.startswith('//'):
                return redirect(next_url)
            return redirect(url_for('index'))
        else:
            flash("Incorrect password. Please try again.", "danger")

    return render_template('login.html')

@app.route('/logout')
def logout():
    session.clear()
    flash("You have been signed out.", "info")
    return redirect(url_for('login'))

# -------------------------------------------------------------------------
# DATABASE CONNECTION
# -------------------------------------------------------------------------
DB_HOST = os.environ.get("DB_HOST", "localhost")
DB_PORT = os.environ.get("DB_PORT", "5432")
DB_NAME = os.environ.get("DB_NAME", "inventory_db")
DB_USER = os.environ.get("DB_USER", "postgres")
DB_PASS = os.environ.get("DB_PASS", "postgres")

def get_db():
    return psycopg2.connect(
        dbname=DB_NAME,
        user=DB_USER,
        password=DB_PASS,
        host=DB_HOST,
        port=DB_PORT
    )

def ensure_schema_migrations():
    try:
        conn = get_db()
        cur = conn.cursor()
        cur.execute("""
            ALTER TABLE quotes ADD COLUMN IF NOT EXISTS quote_number VARCHAR(50) UNIQUE;
            ALTER TABLE quotes ADD COLUMN IF NOT EXISTS quote_base VARCHAR(50);
            ALTER TABLE quotes ADD COLUMN IF NOT EXISTS quote_version INTEGER NOT NULL DEFAULT 1;
            ALTER TABLE orders ADD COLUMN IF NOT EXISTS order_number VARCHAR(50) UNIQUE;
            
            CREATE TABLE IF NOT EXISTS packing_slips (
                slip_id SERIAL PRIMARY KEY,
                order_id INTEGER REFERENCES orders(order_id) ON DELETE CASCADE,
                slip_number VARCHAR(50) NOT NULL UNIQUE,
                delivery_date DATE NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            );
            
            CREATE TABLE IF NOT EXISTS system_settings (
                key VARCHAR(50) PRIMARY KEY,
                value TEXT
            );
        """)
        conn.commit()
        cur.close()
        conn.close()
    except Exception as e:
        print(f"Schema migration error: {e}")

# Run schema migrations check on startup
ensure_schema_migrations()

# -------------------------------------------------------------------------
# NUMBERING HELPER FUNCTIONS (AYYMMDDXX & INYYMMDDXXX)
# -------------------------------------------------------------------------
def generate_next_order_number(cur, order_date=None):
    if not order_date:
        order_date = date.today()
    elif isinstance(order_date, (datetime,)):
        order_date = order_date.date()
    elif isinstance(order_date, str):
        try:
            order_date = datetime.strptime(order_date[:10], '%Y-%m-%d').date()
        except Exception:
            order_date = date.today()
    yymmdd = order_date.strftime('%y%m%d')
    prefix = f"O{yymmdd}"
    cur.execute("SELECT order_number FROM orders WHERE order_number LIKE %s", (f"{prefix}%",))
    rows = cur.fetchall()
    max_seq = 0
    for r in rows:
        num_str = (r.get('order_number') if isinstance(r, dict) else r[0]) or ''
        if len(num_str) >= len(prefix) + 2 and num_str.startswith(prefix):
            seq_part = num_str[len(prefix):len(prefix)+2]
            if seq_part.isdigit():
                max_seq = max(max_seq, int(seq_part))
    next_seq = max_seq + 1
    return f"{prefix}{next_seq:02d}"

def generate_next_quote_number(cur, quote_date=None):
    if not quote_date:
        quote_date = date.today()
    elif isinstance(quote_date, (datetime,)):
        quote_date = quote_date.date()
    elif isinstance(quote_date, str):
        try:
            quote_date = datetime.strptime(quote_date[:10], '%Y-%m-%d').date()
        except Exception:
            quote_date = date.today()
    yymmdd = quote_date.strftime('%y%m%d')
    prefix = f"Q{yymmdd}"
    cur.execute("SELECT quote_base, quote_number FROM quotes WHERE quote_base LIKE %s OR quote_number LIKE %s", (f"{prefix}%", f"{prefix}%"))
    rows = cur.fetchall()
    max_seq = 0
    for r in rows:
        if isinstance(r, dict):
            base_str = r.get('quote_base') or r.get('quote_number') or ''
        else:
            base_str = r[0] or r[1] or ''
        if len(base_str) >= len(prefix) + 2 and base_str.startswith(prefix):
            seq_part = base_str[len(prefix):len(prefix)+2]
            if seq_part.isdigit():
                max_seq = max(max_seq, int(seq_part))
    next_seq = max_seq + 1
    base = f"{prefix}{next_seq:02d}"
    version = 1
    full_number = f"{base}V{version:02d}"
    return base, version, full_number

def generate_next_packing_slip_number(cur, gen_date=None):
    if not gen_date:
        gen_date = date.today()
    elif isinstance(gen_date, (datetime,)):
        gen_date = gen_date.date()
    elif isinstance(gen_date, str):
        try:
            gen_date = datetime.strptime(gen_date[:10], '%Y-%m-%d').date()
        except Exception:
            gen_date = date.today()
    yymmdd = gen_date.strftime('%y%m%d')
    prefix = f"P{yymmdd}"
    cur.execute("SELECT slip_number FROM packing_slips WHERE slip_number LIKE %s", (f"{prefix}%",))
    rows = cur.fetchall()
    max_seq = 0
    for r in rows:
        num_str = (r.get('slip_number') if isinstance(r, dict) else r[0]) or ''
        if len(num_str) >= len(prefix) + 2 and num_str.startswith(prefix):
            seq_part = num_str[len(prefix):len(prefix)+2]
            if seq_part.isdigit():
                max_seq = max(max_seq, int(seq_part))
    next_seq = max_seq + 1
    return f"{prefix}{next_seq:02d}"

def get_last_fiscal_seq(cur):
    cur.execute("SELECT value FROM system_settings WHERE key = 'last_fiscal_order_seq'")
    r = cur.fetchone()
    if r:
        val = r.get('value') if isinstance(r, dict) else r[0]
        if val is not None and str(val).strip().isdigit():
            return int(str(val).strip())
    return None

def set_last_fiscal_seq(cur, seq):
    cur.execute("""
        INSERT INTO system_settings (key, value)
        VALUES ('last_fiscal_order_seq', %s)
        ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value
    """, (str(seq),))

# -------------------------------------------------------------------------
# EXCEL CELL MAPPING CONFIGURATION
# -------------------------------------------------------------------------
CELL_MAPPINGS = {
    # -----------------------------------------------------------------------
    # QUOTE  – billing (customer) address, quote header
    # Template labels: G3=Issue Date, G4=Quote No, G5=Quote Valid
    # "QUOTE TO" block: G8=Name, G9=Lab, G10=Dep, G11=Org,
    #                   G12=Address, G13=City/Prov/Postal, G14=Country
    # Notes block:      A32 (merged A32:I33)
    # -----------------------------------------------------------------------
    'quote': {
        'template': 'quote_template.xlsx',
        'single_fields': {
            'H3': 'created_at_fmt',         # DD-MMM-YY (Issue Date)
            'H4': 'quote_ref_display',      # Quote No (+ Order # if converted)
            'H5': 'expires_at_fmt',         # DD-MMM-YY (Quote Valid)
            'A32': 'customer_notes',
        },
        'customer_blocks': [
            {
                'col': 'G',
                'start_row': 8,
                'end_row': 14,
                'fields': [
                    'customer_name',
                    'lab',
                    'dep',
                    'org',
                    'billing_street_full',
                    'billing_cityline',
                    'billing_country'
                ]
            }
        ],
        'items_start_row': 17,
        'item_columns': {
            'B': 'full_sku',
            'C': 'name',
            'F': 'quantity',
            'G': 'unit_price',
            'H': 'discount',
        }
    },

    # -----------------------------------------------------------------------
    # INVOICE – billing block (A12-A17) + shipping block (G12-G18)
    # Header labels in G3-G7; values in H3-H6:
    #   H3=invoice_number, H4=order_number_display, H5=quote_number_display, H6=order_date_fmt
    # Bill To: A12=billing_name, A13=org, A14=dep, A15=billing_street_full,
    #          A16=billing_cityline, A17=email
    # Ship To: G12=customer_name, G13=org, G14=shipping_unit, G15=dep,
    #          G16=lab, G17=shipping_street, G18=shipping_cityline
    # Details: F21=order_date_fmt, A22=phone, A23=email
    # Totals:  I33=shipping_cost, I35=tax_amount
    # -----------------------------------------------------------------------
    'invoice': {
        'template': 'invoice_template.xlsx',
        'single_fields': {
            # Header
            'H3': 'invoice_number',
            'H4': 'order_number_display',
            'H5': 'quote_number_display',   # Associated Quote No
            'H6': 'order_date_fmt',         # DD-MMM-YY
            # Buyer / Order row
            'F21': 'order_date_fmt',
            'A22': 'phone',
            'A23': 'email',
            # Totals
            'I33': 'shipping_cost',
            'I34': 'ice_charge',
            'I35': 'tax_amount',
        },
        'customer_blocks': [
            {
                'col': 'A',
                'start_row': 12,
                'end_row': 17,
                'fields': [
                    'billing_name_display',
                    'org',
                    'dep',
                    'billing_street_full',
                    'billing_cityline',
                    'email'
                ]
            },
            {
                'col': 'G',
                'start_row': 12,
                'end_row': 18,
                'fields': [
                    'customer_name',
                    'org',
                    'shipping_unit',
                    'dep',
                    'lab',
                    'shipping_street',
                    'shipping_cityline'
                ]
            }
        ],
        'items_start_row': 26,
        'item_columns': {
            'B': 'full_sku',
            'C': 'name',
            'F': 'quantity',
            'G': 'unit_price',
            'H': 'discount',
        }
    },

    # -----------------------------------------------------------------------
    # PACKING SLIP – shipping address, order/quote references
    # Header labels are in E3-E8; values go in the F column beside them:
    #   F4=delivery_date_fmt (Issue Date), F6=quote_number_display (Quote No), F7=order_number_display (Order No)
    # Ship To block starts at E10 label row; values fill E11-E17
    # Details block: B19=customer_name, B20=phone, B21=email, F20=order_date_fmt
    # -----------------------------------------------------------------------
    'packing_slip': {
        'template': 'packing_template.xlsx',
        'single_fields': {
            # Header values (labels are in E column, values go in F)
            'F3': 'slip_number',            # Slip No (PYYMMDDXX)
            'F4': 'delivery_date_fmt',      # DD-MMM-YY (Issue / Delivery Date)
            'F6': 'quote_number_display',   # Associated Quote No
            'F7': 'order_number_display',   # Order No
            # Order & Delivery Details
            'B19': 'customer_name',
            'B20': 'phone',
            'B21': 'email',
            'F20': 'order_date_fmt',        # DD-MMM-YY (Order Date)
            'A33': 'internal_notes',
        },
        'customer_blocks': [
            {
                'col': 'E',
                'start_row': 11,
                'end_row': 17,
                'fields': [
                    'customer_name',
                    'lab',
                    'dep',
                    'org',
                    'shipping_street_full',
                    'shipping_cityline',
                    'shipping_country'
                ]
            }
        ],
        'items_start_row': 24,
        'item_columns': {
            'B': 'full_sku',
            'C': 'name',
            'F': 'quantity',
            'G': 'lot_number',
        }
    },

    # -----------------------------------------------------------------------
    # LABEL – shipping address only
    # A3:B6 is merged (Company return address) — ship-to starts at A7
    # A7=customer_name, A8=shipping_street_full, A9=shipping_cityline, A10=shipping_country
    # -----------------------------------------------------------------------
    'label': {
        'template': 'label_template.xlsx',
        'single_fields': {},
        'customer_blocks': [
            {
                'col': 'A',
                'start_row': 7,
                'end_row': 10,
                'fields': [
                    'customer_name',
                    'shipping_street_full',
                    'shipping_cityline',
                    'shipping_country'
                ]
            }
        ],
        'items_start_row': None,
        'item_columns': {}
    }
}

# -------------------------------------------------------------------------
# EXCEL GENERATOR HELPER
# -------------------------------------------------------------------------
def set_cell_value(ws, cell_address, value, reset_format=False):
    cell = ws[cell_address]
    target_cell = cell
    if type(cell).__name__ == 'MergedCell':
        for rng in ws.merged_cells.ranges:
            if cell.coordinate in rng:
                target_cell = ws.cell(rng.min_row, rng.min_col)
                break
    target_cell.value = value
    if reset_format:
        target_cell.number_format = 'General'

def format_full_sku(sku, sub_sku):
    sku_s = str(sku or '').strip()
    sub_s = str(sub_sku if sub_sku is not None else '').strip()
    if sub_s:
        return f"{sku_s}-{sub_s}"
    return sku_s

def generate_excel_document(doc_type, header_data, item_list, subfolder=None):
    config = CELL_MAPPINGS[doc_type]
    template_path = os.path.join(TEMPLATE_DIR, config['template'])

    wb = openpyxl.load_workbook(template_path)
    # Refer to the second sheet if template has multiple sheets (removes the first junk sheet)
    if len(wb.worksheets) > 1:
        ws = wb.worksheets[1]
        wb.remove(wb.worksheets[0])
    else:
        ws = wb.active

    # --- Build computed fields ---
    data = dict(header_data)  # shallow copy so we don't mutate the caller's dict

    def _cityline(city_key, prov_key, postal_key):
        city   = (data.get(city_key) or '').strip()
        prov   = (data.get(prov_key) or '').strip()
        postal = (data.get(postal_key) or '').strip()
        city_prov = ', '.join(p for p in [city, prov] if p)
        return '  '.join(p for p in [city_prov, postal] if p)

    data['billing_cityline']  = _cityline('billing_city',  'billing_province',  'billing_postal_code')
    data['shipping_cityline'] = _cityline('shipping_city', 'shipping_province', 'shipping_postal_code')

    b_street = (data.get('billing_street') or '').strip()
    b_unit   = (data.get('billing_unit') or '').strip()
    data['billing_street_full'] = f"{b_street}, {b_unit}" if (b_street and b_unit) else (b_street or b_unit)

    s_street = (data.get('shipping_street') or '').strip()
    s_unit   = (data.get('shipping_unit') or '').strip()
    data['shipping_street_full'] = f"{s_street}, {s_unit}" if (s_street and s_unit) else (s_street or s_unit)

    data['billing_name_display'] = (data.get('billing_name') or data.get('customer_name') or data.get('name') or '').strip()
    data['customer_name'] = (data.get('customer_name') or data.get('name') or '').strip()

    data['order_number_display'] = data.get('order_number') or (f"#{data.get('order_id')}" if data.get('order_id') else '')
    data['quote_number_display'] = data.get('quote_number') or (f"#{data.get('quote_id')}" if data.get('quote_id') else '')

    # Quote reference display (includes associated Order # if converted)
    q_disp = data.get('quote_number_display') or ''
    o_disp = data.get('order_number_display') or ''
    if doc_type == 'quote' and q_disp and o_disp:
        data['quote_ref_display'] = f"{q_disp} (Order #{o_disp})"
    else:
        data['quote_ref_display'] = q_disp

    # DD-MMM-YY formatted dates
    def _fmt_date(val):
        if val is None:
            return ''
        if isinstance(val, (datetime, date)):
            return val.strftime('%d-%b-%y')
        try:
            return datetime.strptime(str(val)[:10], '%Y-%m-%d').strftime('%d-%b-%y')
        except ValueError:
            return str(val)

    data['created_at_fmt'] = _fmt_date(data.get('created_at'))
    data['expires_at_fmt'] = _fmt_date(data.get('expires_at'))
    data['order_date_fmt'] = _fmt_date(data.get('order_date'))
    data['delivery_date_fmt'] = _fmt_date(data.get('delivery_date') or data.get('last_packing_slip_date') or date.today())
    data['shipping_date_fmt'] = _fmt_date(data.get('delivery_date') or data.get('last_packing_slip_date') or data.get('shipping_date'))
    # --- End computed fields ---

    # Document-specific adjustments
    if doc_type == 'invoice':
        pay_method = (data.get('payment_method') or 'Credit Card').strip()
        po_num = (data.get('po_number') or '').strip()
        ship_date_val = data.get('shipping_date_fmt') or ''

        # Payment method in H21
        set_cell_value(ws, 'H21', pay_method)

        # Shipping date in G21 if packing slip was generated
        if ship_date_val:
            set_cell_value(ws, 'G21', ship_date_val)
        else:
            set_cell_value(ws, 'G21', '')

        if pay_method.upper() == 'PO':
            # List PO# in H9 and C21
            set_cell_value(ws, 'H9', po_num)
            set_cell_value(ws, 'C21', po_num)
        else:
            # Credit Card or other non-PO:
            # Clear H9 and C21 values, remove headers in C20 and E20, and clear E21 ('Standard')
            set_cell_value(ws, 'H9', '')
            set_cell_value(ws, 'C20', '')
            set_cell_value(ws, 'E20', '')
            set_cell_value(ws, 'C21', '')
            set_cell_value(ws, 'E21', '')

    elif doc_type == 'quote':
        # Ensure top border of G3 and bottom of G1 (merged G1:I2 title block) is preserved
        f2_border = ws['F2'].border
        green_side = copy.copy(f2_border.bottom) if (f2_border and f2_border.bottom and f2_border.bottom.style) else Side(style='medium', color='FF00B050')
        b_g1 = ws['G1'].border
        ws['G1'].border = Border(top=b_g1.top, left=b_g1.left, right=b_g1.right, bottom=green_side)
        b_g3 = ws['G3'].border
        ws['G3'].border = Border(top=green_side, left=b_g3.left, right=b_g3.right, bottom=b_g3.bottom)

    elif doc_type == 'packing_slip':
        pay_method = (data.get('payment_method') or '').strip().upper()
        po_num = (data.get('po_number') or '').strip()

        # Slip number in F3
        slip_num = data.get('slip_number') or f"P{date.today().strftime('%y%m%d')}01"
        set_cell_value(ws, 'F3', slip_num)

        # Delivery / issue date in F4
        deliv_fmt = data.get('delivery_date_fmt') or _fmt_date(date.today())
        set_cell_value(ws, 'F4', deliv_fmt)

        # Order date in F20
        ord_date_fmt = data.get('order_date_fmt') or ''
        set_cell_value(ws, 'F20', ord_date_fmt)

        # Customer PO handling:
        # If PO given beforehand, put in F8 and F21. If not PO, delete E8 and E21.
        if pay_method == 'PO' and po_num:
            set_cell_value(ws, 'E8', 'Cust. PO#')
            set_cell_value(ws, 'F8', po_num)
            set_cell_value(ws, 'E21', 'Cust. PO#:')
            set_cell_value(ws, 'F21', po_num)
        else:
            set_cell_value(ws, 'E8', '')
            set_cell_value(ws, 'F8', '')
            set_cell_value(ws, 'E21', '')
            set_cell_value(ws, 'F21', '')

    elif doc_type == 'label':
        est_weight = data.get('estimated_weight')
        if est_weight is not None and str(est_weight).strip() != '':
            # Format nicely, e.g. "5 lbs" or raw number
            w_str = str(est_weight).strip()
            set_cell_value(ws, 'D24', w_str)

    NUMERIC_SINGLE_FIELDS = {
        'shipping_cost', 'tax_amount', 'ice_charge', 
        'estimated_shipping', 'estimated_tax', 'total_amount'
    }

    # Write Single Field Values
    for cell_address, field_key in config['single_fields'].items():
        if field_key in data:
            raw_val = data[field_key]
            if raw_val is None or str(raw_val).strip() == '':
                if field_key == 'ice_charge':
                    set_cell_value(ws, cell_address, '')
                    continue
                val = ""
            elif field_key in NUMERIC_SINGLE_FIELDS:
                try:
                    val = float(raw_val)
                except (ValueError, TypeError):
                    val = raw_val
            else:
                val = str(raw_val)
            reset_fmt = field_key in ('order_id', 'order_number', 'order_number_display', 'quote_id', 'quote_number', 'quote_number_display', 'invoice_number', 'quote_ref_display', 'slip_number')
            set_cell_value(ws, cell_address, val, reset_format=reset_fmt)

    # Write Customer Data Blocks (cascading / shifting up when rows are empty)
    for block in config.get('customer_blocks', []):
        col = block['col']
        start_row = block['start_row']
        end_row = block['end_row']
        # Collect non-empty values in order
        values = []
        for field_key in block['fields']:
            val = (data.get(field_key) or '').strip()
            if val:
                values.append(val)
        # Write packed values starting at start_row
        for i, row in enumerate(range(start_row, end_row + 1)):
            if i < len(values):
                set_cell_value(ws, f"{col}{row}", values[i])
            else:
                set_cell_value(ws, f"{col}{row}", "")

    # Write Line Items
    if config['items_start_row'] and config['item_columns']:
        start_row = config['items_start_row']
        for idx, raw_item in enumerate(item_list):
            item = dict(raw_item)
            item['full_sku'] = format_full_sku(item.get('sku'), item.get('sub_sku'))
            # Format product name as "name, size" if size is available
            p_name = (item.get('name') or '').strip()
            p_size = (item.get('size') or '').strip()
            if p_size and p_size.lower() not in ('n/a', 'none', '-'):
                item['name'] = f"{p_name}, {p_size}"
            else:
                item['name'] = p_name
            row = start_row + idx
            for col_letter, item_key in config['item_columns'].items():
                if item_key not in item:
                    continue
                raw_val = item[item_key]
                if raw_val is None or str(raw_val).strip() == '':
                    continue

                # Discount formatted as decimal for Excel percentage cell
                if item_key == 'discount':
                    try:
                        disc_val = float(raw_val or 0.0)
                        if disc_val == 0.0:
                            continue
                        excel_disc = disc_val / 100.0 if disc_val > 1.0 else disc_val
                        set_cell_value(ws, f"{col_letter}{row}", excel_disc)
                        continue
                    except (TypeError, ValueError):
                        continue
                elif item_key == 'quantity':
                    try:
                        q_num = float(raw_val)
                        q_val = int(q_num) if q_num.is_integer() else q_num
                        set_cell_value(ws, f"{col_letter}{row}", q_val)
                        continue
                    except (TypeError, ValueError):
                        pass
                elif item_key in ('unit_price', 'price'):
                    try:
                        p_val = float(raw_val)
                        set_cell_value(ws, f"{col_letter}{row}", p_val)
                        continue
                    except (TypeError, ValueError):
                        pass

                set_cell_value(ws, f"{col_letter}{row}", raw_val)

    if doc_type == 'packing_slip':
        doc_identifier = data.get('slip_number') or f"P{date.today().strftime('%y%m%d')}01"
    elif doc_type == 'quote':
        doc_identifier = data.get('quote_number') or f"Q{date.today().strftime('%y%m%d')}01V01"
    elif doc_type == 'invoice':
        doc_identifier = data.get('invoice_number') or f"IN{date.today().strftime('%y%m%d')}001"
    else: # label
        doc_identifier = data.get('order_number') or (f"O{data.get('order_id')}" if data.get('order_id') else 'label')
    filename = f"{doc_type}_{doc_identifier}.xlsx"

    if subfolder:
        target_dir = os.path.join(OUTPUT_DIR, subfolder)
    else:
        target_dir = OUTPUT_DIR
    os.makedirs(target_dir, exist_ok=True)

    # Clean up old differently-named documents of this type for this order/quote if regenerating
    for old_f in os.listdir(target_dir):
        if old_f.startswith(f"{doc_type}_") and old_f.endswith('.xlsx') and old_f != filename:
            try:
                os.remove(os.path.join(target_dir, old_f))
            except Exception:
                pass

    output_path = os.path.join(target_dir, filename)
    wb.save(output_path)
    return filename


# Auto-migrate any flat files and consolidate converted quotes into order subfolders
def migrate_existing_documents():
    if not os.path.exists(OUTPUT_DIR):
        return
    # 1. Flat files migration into subfolders
    for item in os.listdir(OUTPUT_DIR):
        item_path = os.path.join(OUTPUT_DIR, item)
        if os.path.isfile(item_path) and item.endswith('.xlsx'):
            target_subfolder = None
            if item.startswith('quote_'):
                parts = item.replace('.xlsx', '').split('_')
                if len(parts) >= 2 and parts[1].isdigit():
                    target_subfolder = f"quote_{parts[1]}"
            elif 'INV-' in item:
                m = re.search(r'INV-(\d+)', item)
                if m:
                    target_subfolder = f"order_{int(m.group(1))}"
            if target_subfolder:
                dest_dir = os.path.join(OUTPUT_DIR, target_subfolder)
                os.makedirs(dest_dir, exist_ok=True)
                dest_path = os.path.join(dest_dir, item)
                if not os.path.exists(dest_path):
                    shutil.move(item_path, dest_path)

    # 2. Consolidate converted quote folders into their linked order folder
    try:
        conn = get_db()
        cur = conn.cursor()
        cur.execute("SELECT order_id, quote_id FROM orders WHERE quote_id IS NOT NULL")
        rows = cur.fetchall()
        for oid, qid in rows:
            q_dir = os.path.join(OUTPUT_DIR, f"quote_{qid}")
            o_dir = os.path.join(OUTPUT_DIR, f"order_{oid}")
            if os.path.exists(q_dir):
                os.makedirs(o_dir, exist_ok=True)
                for f in os.listdir(q_dir):
                    src = os.path.join(q_dir, f)
                    dst = os.path.join(o_dir, f)
                    if not os.path.exists(dst):
                        shutil.move(src, dst)
                    else:
                        shutil.copy2(src, dst)
                        os.remove(src)
                try:
                    os.rmdir(q_dir)
                except Exception:
                    pass
        cur.close()
        conn.close()
    except Exception:
        pass

migrate_existing_documents()

def regenerate_order_documents(cur, order_id):
    """
    If any documents (invoice, packing slip, shipping label) exist in order_<order_id>,
    re-render and overwrite them with current database values.
    """
    order_folder = os.path.join(OUTPUT_DIR, f"order_{order_id}")
    if not os.path.exists(order_folder):
        return

    cur.execute("""
        SELECT c.*, o.*, COALESCE(c.name, 'Deleted Customer') AS customer_name
        FROM orders o
        LEFT JOIN customers c ON o.customer_id = c.customer_id
        WHERE o.order_id = %s
    """, (order_id,))
    order_data = cur.fetchone()
    if not order_data:
        return

    cur.execute("""
        SELECT oi.*, p.sku, p.sub_sku, p.name, p.size
        FROM order_items oi
        JOIN products p ON oi.product_id = p.product_id
        WHERE oi.order_id = %s
        ORDER BY oi.order_item_id ASC
    """, (order_id,))
    items = cur.fetchall()

    for fname in os.listdir(order_folder):
        if not fname.endswith('.xlsx'):
            continue
        try:
            if fname.startswith('invoice_'):
                generate_excel_document('invoice', order_data, items, subfolder=f"order_{order_id}")
            elif fname.startswith('packing_slip_'):
                # Packing slip uses fulfilled items or current items
                generate_excel_document('packing_slip', order_data, items, subfolder=f"order_{order_id}")
            elif fname.startswith('label_'):
                generate_excel_document('label', order_data, items, subfolder=f"order_{order_id}")
            elif fname.startswith('quote_'):
                # In case quote was moved into this order folder
                if order_data.get('quote_id'):
                    cur.execute("""
                        SELECT c.*, q.*, COALESCE(c.name, 'Deleted Customer') AS customer_name, %s AS order_id
                        FROM quotes q
                        LEFT JOIN customers c ON q.customer_id = c.customer_id
                        WHERE q.quote_id = %s
                    """, (order_id, order_data['quote_id']))
                    q_data = cur.fetchone()
                    if q_data:
                        cur.execute("""
                            SELECT qi.*, p.sku, p.sub_sku, p.name, p.size
                            FROM quote_items qi
                            JOIN products p ON qi.product_id = p.product_id
                            WHERE qi.quote_id = %s
                            ORDER BY qi.quote_item_id ASC
                        """, (order_data['quote_id'],))
                        q_items = cur.fetchall()
                        generate_excel_document('quote', q_data, q_items, subfolder=f"order_{order_id}")
        except Exception:
            pass

def regenerate_quote_document(cur, quote_id):
    """
    If quote document exists (in quote_<quote_id> or order_<order_id>), re-render and overwrite it.
    """
    cur.execute("""
        SELECT c.*, q.*, COALESCE(c.name, 'Deleted Customer') AS customer_name,
               (SELECT o.order_id FROM orders o WHERE o.quote_id = q.quote_id ORDER BY o.order_id DESC LIMIT 1) AS order_id,
               (SELECT o.order_number FROM orders o WHERE o.quote_id = q.quote_id ORDER BY o.order_id DESC LIMIT 1) AS order_number
        FROM quotes q
        LEFT JOIN customers c ON q.customer_id = c.customer_id
        WHERE q.quote_id = %s
    """, (quote_id,))
    quote_data = cur.fetchone()
    if not quote_data:
        return

    subfolder = f"order_{quote_data['order_id']}" if quote_data.get('order_id') else f"quote_{quote_id}"
    quote_folder = os.path.join(OUTPUT_DIR, subfolder)
    if not os.path.exists(quote_folder):
        return

    cur.execute("""
        SELECT qi.*, p.sku, p.sub_sku, p.name, p.size
        FROM quote_items qi
        JOIN products p ON qi.product_id = p.product_id
        WHERE qi.quote_id = %s
        ORDER BY qi.quote_item_id ASC
    """, (quote_id,))
    items = cur.fetchall()

    try:
        generate_excel_document('quote', quote_data, items, subfolder=subfolder)
    except Exception:
        pass

# -------------------------------------------------------------------------
# INVENTORY RESERVATION & LOT ALLOCATION HELPERS (FEFO)
# -------------------------------------------------------------------------
def helper_reserve_lot_fefo(cur, product_id, quantity, preferred_lot_number=None):
    """
    Allocates and reserves stock using FEFO (First Expired, First Out).
    Returns (assigned_lot_number, assigned_lot_id).
    """
    # 1. If preferred lot number is provided, try to reserve against that specific lot
    if preferred_lot_number:
        cur.execute("""
            SELECT lot_id, lot_number, stock, reserved 
            FROM product_lots 
            WHERE product_id = %s AND lot_number = %s
            LIMIT 1
        """, (product_id, preferred_lot_number))
        lot = cur.fetchone()
        if lot:
            cur.execute("""
                UPDATE product_lots 
                SET reserved = reserved + %s 
                WHERE lot_id = %s
            """, (quantity, lot['lot_id']))
            return lot['lot_number'], lot['lot_id']

    # 2. FEFO: allocate to lot with available stock (stock - reserved > 0) closest to expiring
    cur.execute("""
        SELECT lot_id, lot_number, stock, reserved, expiry_date,
               (stock - reserved) AS available
        FROM product_lots 
        WHERE product_id = %s AND (stock - reserved) > 0
        ORDER BY expiry_date ASC, lot_id ASC
        LIMIT 1
    """, (product_id,))
    lot = cur.fetchone()
    if lot:
        cur.execute("""
            UPDATE product_lots 
            SET reserved = reserved + %s 
            WHERE lot_id = %s
        """, (quantity, lot['lot_id']))
        return lot['lot_number'], lot['lot_id']

    # 3. Fallback: if no available stock, check if any lot exists at all for this product
    cur.execute("""
        SELECT lot_id, lot_number, expiry_date 
        FROM product_lots 
        WHERE product_id = %s 
        ORDER BY expiry_date ASC, lot_id ASC 
        LIMIT 1
    """, (product_id,))
    fallback_lot = cur.fetchone()
    if fallback_lot:
        cur.execute("""
            UPDATE product_lots 
            SET reserved = reserved + %s 
            WHERE lot_id = %s
        """, (quantity, fallback_lot['lot_id']))
        return fallback_lot['lot_number'], fallback_lot['lot_id']

    return None, None

def helper_fulfill_order_items(cur, order_id):
    """
    Deducts physical stock and removes reservation when order is marked Fulfilled.
    """
    cur.execute("""
        UPDATE order_items 
        SET fulfilled_quantity = quantity 
        WHERE order_id = %s
    """, (order_id,))
    
    cur.execute("""
        SELECT order_item_id, product_id, quantity, lot_id, lot_number 
        FROM order_items 
        WHERE order_id = %s
    """, (order_id,))
    items = cur.fetchall()
    
    for item in items:
        qty = item['quantity']
        if item['lot_id']:
            cur.execute("""
                UPDATE product_lots 
                SET stock = GREATEST(0, stock - %s),
                    reserved = GREATEST(0, reserved - %s)
                WHERE lot_id = %s
            """, (qty, qty, item['lot_id']))
        elif item['lot_number']:
            cur.execute("""
                UPDATE product_lots 
                SET stock = GREATEST(0, stock - %s),
                    reserved = GREATEST(0, reserved - %s)
                WHERE product_id = %s AND lot_number = %s
            """, (qty, qty, item['product_id'], item['lot_number']))
        else:
            # Pick earliest expiring lot for product
            cur.execute("""
                UPDATE product_lots 
                SET stock = GREATEST(0, stock - %s),
                    reserved = GREATEST(0, reserved - %s)
                WHERE lot_id = (
                    SELECT lot_id FROM product_lots 
                    WHERE product_id = %s 
                    ORDER BY expiry_date ASC LIMIT 1
                )
            """, (qty, qty, item['product_id']))

def helper_unfulfill_order_items(cur, order_id):
    """
    Reverts physical stock deduction and restores reservation if order changes from Fulfilled back to Pending/Unfulfilled.
    """
    cur.execute("""
        UPDATE order_items 
        SET fulfilled_quantity = 0 
        WHERE order_id = %s
    """, (order_id,))

    cur.execute("""
        SELECT order_item_id, product_id, quantity, lot_id, lot_number 
        FROM order_items 
        WHERE order_id = %s
    """, (order_id,))
    items = cur.fetchall()
    
    for item in items:
        qty = item['quantity']
        if item['lot_id']:
            cur.execute("""
                UPDATE product_lots 
                SET stock = stock + %s,
                    reserved = reserved + %s
                WHERE lot_id = %s
            """, (qty, qty, item['lot_id']))
        elif item['lot_number']:
            cur.execute("""
                UPDATE product_lots 
                SET stock = stock + %s,
                    reserved = reserved + %s
                WHERE product_id = %s AND lot_number = %s
            """, (qty, qty, item['product_id'], item['lot_number']))

def helper_release_order_reservations(cur, order_id):
    """
    Releases reserved stock if an unfulfilled/pending order is deleted.
    """
    cur.execute("""
        SELECT order_item_id, product_id, quantity, lot_id, lot_number 
        FROM order_items 
        WHERE order_id = %s
    """, (order_id,))
    items = cur.fetchall()
    
    for item in items:
        qty = item['quantity']
        if item['lot_id']:
            cur.execute("""
                UPDATE product_lots 
                SET reserved = GREATEST(0, reserved - %s)
                WHERE lot_id = %s
            """, (qty, item['lot_id']))
        elif item['lot_number']:
            cur.execute("""
                UPDATE product_lots 
                SET reserved = GREATEST(0, reserved - %s)
                WHERE product_id = %s AND lot_number = %s
            """, (qty, item['product_id'], item['lot_number']))

def sync_lot_reservations(cur):
    """
    Synchronizes product_lots.reserved with actual active (non-Fulfilled) orders
    so orphaned or deleted orders never leave behind phantom reserved stock,
    and unallocated order items get linked to available lots via FEFO.
    """
    # 0. Ensure any order in 'Pending' or 'Unfulfilled' status has fulfilled_quantity reset to 0
    cur.execute("""
        UPDATE order_items oi
        SET fulfilled_quantity = 0
        FROM orders o
        WHERE oi.order_id = o.order_id
          AND o.fulfillment_status IN ('Pending', 'Unfulfilled')
          AND COALESCE(oi.fulfilled_quantity, 0) > 0
    """)

    # 1. Attempt to link unallocated active order items to lots using FEFO if lots exist
    cur.execute("""
        SELECT oi.order_item_id, oi.product_id, 
               (CASE 
                    WHEN o.fulfillment_status = 'Partially Fulfilled' THEN GREATEST(0, oi.quantity - COALESCE(oi.fulfilled_quantity, 0))
                    ELSE oi.quantity 
                END) AS rem_qty
        FROM order_items oi
        JOIN orders o ON oi.order_id = o.order_id
        WHERE o.fulfillment_status NOT IN ('Fulfilled', 'Cancelled')
          AND oi.lot_id IS NULL
          AND (
              CASE 
                  WHEN o.fulfillment_status = 'Partially Fulfilled' THEN (oi.quantity - COALESCE(oi.fulfilled_quantity, 0))
                  ELSE oi.quantity 
              END
          ) > 0
    """)
    unassigned = cur.fetchall()
    for u in unassigned:
        pid = u['product_id']
        cur.execute("""
            SELECT lot_id, lot_number
            FROM product_lots
            WHERE product_id = %s AND (stock - reserved) > 0
            ORDER BY expiry_date ASC, lot_id ASC
            LIMIT 1
        """, (pid,))
        lot = cur.fetchone()
        if not lot:
            cur.execute("""
                SELECT lot_id, lot_number
                FROM product_lots
                WHERE product_id = %s
                ORDER BY expiry_date ASC, lot_id ASC
                LIMIT 1
            """, (pid,))
            lot = cur.fetchone()
        if lot:
            cur.execute("""
                UPDATE order_items
                SET lot_id = %s, lot_number = %s
                WHERE order_item_id = %s
            """, (lot['lot_id'], lot['lot_number'], u['order_item_id']))

    # 2. Sync product_lots.reserved counts
    cur.execute("""
        UPDATE product_lots pl
        SET reserved = COALESCE((
            SELECT SUM(
                CASE 
                    WHEN o.fulfillment_status = 'Partially Fulfilled' THEN GREATEST(0, oi.quantity - COALESCE(oi.fulfilled_quantity, 0))
                    WHEN o.fulfillment_status IN ('Fulfilled', 'Cancelled') THEN 0
                    ELSE oi.quantity
                END
            )
            FROM order_items oi
            JOIN orders o ON oi.order_id = o.order_id
            WHERE o.fulfillment_status NOT IN ('Fulfilled', 'Cancelled')
              AND (
                  oi.lot_id = pl.lot_id
                  OR (oi.lot_id IS NULL AND oi.product_id = pl.product_id AND oi.lot_number = pl.lot_number)
              )
        ), 0)
    """)

# -------------------------------------------------------------------------
# ROUTES: DASHBOARD
# -------------------------------------------------------------------------

@app.route('/')
def index():
    conn = get_db()
    cur = conn.cursor(cursor_factory=RealDictCursor)
    
    # Fetch Quotes with Customer info and associated Order ID if converted
    # Prioritize quotes requiring action (Pending / Active) at top; converted quotes filtered down to bottom
    cur.execute("""
        SELECT q.*, COALESCE(c.name, 'Deleted Customer') AS customer_name, c.email, c.phone,
               (SELECT o.order_id FROM orders o WHERE o.quote_id = q.quote_id ORDER BY o.order_id DESC LIMIT 1) AS linked_order_id,
               (SELECT o.order_number FROM orders o WHERE o.quote_id = q.quote_id ORDER BY o.order_id DESC LIMIT 1) AS linked_order_number
        FROM quotes q
        LEFT JOIN customers c ON q.customer_id = c.customer_id
        ORDER BY 
            CASE WHEN LOWER(COALESCE(q.status, 'pending')) = 'converted' THEN 1 ELSE 0 END ASC,
            q.created_at DESC, 
            q.quote_id DESC
    """)
    quotes = cur.fetchall()
    
    # Fetch Orders with Customer info
    # Prioritize orders with actions remaining (Unfulfilled / Pending / Unpaid) at top; completed orders (Fulfilled & Paid) filtered down
    cur.execute("""
        SELECT o.*, COALESCE(c.name, 'Deleted Customer') AS customer_name, c.email, c.phone,
               q.quote_number
        FROM orders o
        LEFT JOIN customers c ON o.customer_id = c.customer_id
        LEFT JOIN quotes q ON o.quote_id = q.quote_id
        ORDER BY 
            CASE 
                WHEN LOWER(COALESCE(o.fulfillment_status, '')) = 'fulfilled' 
                     AND LOWER(COALESCE(o.payment_status, '')) = 'paid' THEN 2
                WHEN LOWER(COALESCE(o.fulfillment_status, '')) = 'cancelled' THEN 3
                ELSE 0 
            END ASC,
            o.order_date DESC, 
            o.order_id DESC
    """)
    orders = cur.fetchall()

    # Fetch Customers with order/quote counts and total spend for View by Customer
    cur.execute("""
        SELECT c.*,
               (SELECT COUNT(*) FROM orders o WHERE o.customer_id = c.customer_id) AS order_count,
               (SELECT COALESCE(SUM(o.total_amount), 0) FROM orders o WHERE o.customer_id = c.customer_id) AS total_spent,
               (SELECT COUNT(*) FROM quotes q WHERE q.customer_id = c.customer_id) AS quote_count
        FROM customers c 
        ORDER BY c.name ASC
    """)
    customers = cur.fetchall()

    # Fetch Products for line item selection
    cur.execute("""
        SELECT product_id, sku, sub_sku, name, size, price, COALESCE(supplier, 'Default Supplier') AS supplier
        FROM products 
        ORDER BY sku ASC, sub_sku ASC
    """)
    products = cur.fetchall()
    
    cur.close()
    conn.close()
    return render_template('index.html', quotes=quotes, orders=orders, customers=customers, products=products, active_page='dashboard')

# -------------------------------------------------------------------------
# ROUTES: CREATE & EDIT ORDER / QUOTE
# -------------------------------------------------------------------------

def helper_resolve_customer(cur, form):
    customer_type = form.get('customer_type', 'existing')
    if customer_type == 'new':
        name = form.get('new_name', '').strip()
        if not name:
            raise ValueError("Customer name is required.")
        
        email = form.get('new_email', '').strip() or None
        phone = form.get('new_phone', '').strip() or None
        lab = form.get('new_lab', '').strip() or None
        org = form.get('new_org', '').strip() or None
        dep = form.get('new_dep', '').strip() or None
        billing_name = form.get('new_billing_name', '').strip() or name
        
        billing_street = form.get('new_billing_street', '').strip()
        billing_unit = form.get('new_billing_unit', '').strip() or None
        billing_city = form.get('new_billing_city', '').strip()
        billing_province = form.get('new_billing_province', '').strip()
        billing_postal_code = form.get('new_billing_postal_code', '').strip()
        billing_country = form.get('new_billing_country', '').strip() or 'Canada'
        
        same_shipping = form.get('same_shipping') == 'on'
        if same_shipping:
            shipping_street = billing_street
            shipping_unit = billing_unit
            shipping_city = billing_city
            shipping_province = billing_province
            shipping_postal_code = billing_postal_code
            shipping_country = billing_country
        else:
            shipping_street = form.get('new_shipping_street', '').strip() or billing_street
            shipping_unit = form.get('new_shipping_unit', '').strip() or None
            shipping_city = form.get('new_shipping_city', '').strip() or billing_city
            shipping_province = form.get('new_shipping_province', '').strip() or billing_province
            shipping_postal_code = form.get('new_shipping_postal_code', '').strip() or billing_postal_code
            shipping_country = form.get('new_shipping_country', '').strip() or billing_country

        # Ensure check_contact_method constraint (email or phone not null) is satisfied
        if not email and not phone:
            phone = 'N/A'

        cur.execute("""
            INSERT INTO customers (
                name, lab, org, dep, email, phone, billing_name,
                billing_street, billing_unit, billing_city, billing_province, billing_postal_code, billing_country,
                shipping_street, shipping_unit, shipping_city, shipping_province, shipping_postal_code, shipping_country
            )
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            RETURNING customer_id
        """, (
            name, lab, org, dep, email, phone, billing_name,
            billing_street, billing_unit, billing_city, billing_province, billing_postal_code, billing_country,
            shipping_street, shipping_unit, shipping_city, shipping_province, shipping_postal_code, shipping_country
        ))
        row = cur.fetchone()
        return row['customer_id'] if isinstance(row, dict) else row[0]
    else:
        cust_id = form.get('existing_customer_id')
        if not cust_id:
            raise ValueError("Please select an existing customer or create a new one.")
        return int(cust_id)

@app.route('/create_order', methods=['POST'])
def create_order():
    conn = get_db()
    cur = conn.cursor(cursor_factory=RealDictCursor)
    
    try:
        customer_id = helper_resolve_customer(cur, request.form)

        # Parse Order Details
        order_date = request.form.get('order_date') or None
        shipping_cost = float(request.form.get('shipping_cost') or 0.0)
        
        # Tax entered in % (e.g. 13.00 for 13%)
        tax_rate = float(request.form.get('tax_rate') or 0.0)
        
        internal_notes = request.form.get('internal_notes', '').strip() or None

        # Parse Line Items
        product_ids = request.form.getlist('product_id[]')
        quantities = request.form.getlist('quantity[]')
        unit_prices = request.form.getlist('unit_price[]')
        discounts = request.form.getlist('discount[]')
        lot_numbers = request.form.getlist('lot_number[]')

        items_to_insert = []
        calculated_subtotal = 0.0

        for i, p_id in enumerate(product_ids):
            if not p_id or not p_id.strip():
                continue
            try:
                pid_int = int(p_id)
                qty_int = int(quantities[i]) if i < len(quantities) and int(quantities[i]) > 0 else 1
                price_flt = float(unit_prices[i]) if i < len(unit_prices) and unit_prices[i] else 0.0
                disc_pct = float(discounts[i]) if i < len(discounts) and discounts[i] else 0.0
                pref_lot = lot_numbers[i].strip() if i < len(lot_numbers) and lot_numbers[i] else None
                
                # Automatic FEFO reservation and lot allocation
                assigned_lot_number, assigned_lot_id = helper_reserve_lot_fefo(cur, pid_int, qty_int, pref_lot)
                
                line_total = max(0.0, qty_int * price_flt * (1.0 - disc_pct / 100.0))
                calculated_subtotal += line_total
                
                items_to_insert.append({
                    'product_id': pid_int,
                    'quantity': qty_int,
                    'unit_price': price_flt,
                    'discount': disc_pct,  # store as percentage (e.g. 10 for 10%)
                    'lot_number': assigned_lot_number,
                    'lot_id': assigned_lot_id
                })
            except (ValueError, TypeError):
                continue

        if not items_to_insert:
            conn.rollback()
            flash("Please add at least one valid product to the order.", "warning")
            return redirect(url_for('index'))

        # Tax calculated from percentage
        calculated_tax = round(calculated_subtotal * (tax_rate / 100.0), 2)
        total_amount = round(calculated_subtotal + shipping_cost + calculated_tax, 2)
        order_num = generate_next_order_number(cur, order_date)
        cur.execute("""
            INSERT INTO orders (
                customer_id, order_number, order_date, total_amount, fulfillment_status,
                payment_status, invoice_number, shipping_cost, tax_amount, internal_notes
            )
            VALUES (%s, %s, COALESCE(%s::date, CURRENT_DATE), %s, 'Pending', 'Unpaid', NULL, %s, %s, %s)
            RETURNING order_id
        """, (customer_id, order_num, order_date, total_amount, shipping_cost, calculated_tax, internal_notes))
        new_order_id = cur.fetchone()['order_id']

        for item in items_to_insert:
            cur.execute("""
                INSERT INTO order_items (order_id, product_id, quantity, unit_price, discount, lot_number, lot_id)
                VALUES (%s, %s, %s, %s, %s, %s, %s)
            """, (new_order_id, item['product_id'], item['quantity'], item['unit_price'], item['discount'], item['lot_number'], item['lot_id']))

        sync_lot_reservations(cur)
        conn.commit()
        flash(f"Order {order_num} created! Items reserved via FEFO.", "success")
    except Exception as e:
        conn.rollback()
        flash(f"Error creating order: {str(e)}", "danger")
    finally:
        cur.close()
        conn.close()

    return redirect(url_for('index'))

@app.route('/edit_order/<int:order_id>', methods=['POST'])
def edit_order(order_id):
    """
    Updates an existing order and its line items WITHOUT modifying any related quote.
    """
    conn = get_db()
    cur = conn.cursor(cursor_factory=RealDictCursor)
    try:
        cur.execute("SELECT * FROM orders WHERE order_id = %s", (order_id,))
        existing_order = cur.fetchone()
        if not existing_order:
            flash("Order not found.", "danger")
            return redirect(url_for('index'))

        customer_id = int(request.form.get('customer_id') or request.form.get('existing_customer_id') or existing_order['customer_id'])
        order_date = request.form.get('order_date') or None
        fulfillment_status = request.form.get('fulfillment_status') or existing_order['fulfillment_status']
        payment_status = request.form.get('payment_status') or existing_order['payment_status']
        shipping_cost = float(request.form.get('shipping_cost') or 0.0)
        tax_rate = float(request.form.get('tax_rate') or 0.0)
        internal_notes = request.form.get('internal_notes', '').strip() or None

        product_ids = request.form.getlist('product_id[]')
        quantities = request.form.getlist('quantity[]')
        unit_prices = request.form.getlist('unit_price[]')
        discounts = request.form.getlist('discount[]')
        lot_numbers = request.form.getlist('lot_number[]')

        # Revert old stock / reservations before replacing order_items
        was_fulfilled = (existing_order['fulfillment_status'] == 'Fulfilled')
        if was_fulfilled:
            helper_unfulfill_order_items(cur, order_id)
        helper_release_order_reservations(cur, order_id)
        cur.execute("DELETE FROM order_items WHERE order_id = %s", (order_id,))

        items_to_insert = []
        calculated_subtotal = 0.0

        for i, p_id in enumerate(product_ids):
            if not p_id or not p_id.strip():
                continue
            try:
                pid_int = int(p_id)
                qty_int = int(quantities[i]) if i < len(quantities) and int(quantities[i]) > 0 else 1
                price_flt = float(unit_prices[i]) if i < len(unit_prices) and unit_prices[i] else 0.0
                disc_pct = float(discounts[i]) if i < len(discounts) and discounts[i] else 0.0
                pref_lot = lot_numbers[i].strip() if i < len(lot_numbers) and lot_numbers[i] else None

                assigned_lot_number, assigned_lot_id = helper_reserve_lot_fefo(cur, pid_int, qty_int, pref_lot)

                line_total = max(0.0, qty_int * price_flt * (1.0 - disc_pct / 100.0))
                calculated_subtotal += line_total

                items_to_insert.append({
                    'product_id': pid_int,
                    'quantity': qty_int,
                    'unit_price': price_flt,
                    'discount': disc_pct,
                    'lot_number': assigned_lot_number,
                    'lot_id': assigned_lot_id
                })
            except (ValueError, TypeError):
                continue

        if not items_to_insert:
            conn.rollback()
            flash("Order must have at least one valid product.", "warning")
            return redirect(url_for('index'))

        payment_method = request.form.get('payment_method', '').strip() or None
        po_number = request.form.get('po_number', '').strip() or None

        calculated_tax = round(calculated_subtotal * (tax_rate / 100.0), 2)
        total_amount = round(calculated_subtotal + shipping_cost + calculated_tax, 2)

        cur.execute("""
            UPDATE orders
            SET customer_id = %s,
                order_date = COALESCE(%s::date, order_date),
                fulfillment_status = %s,
                payment_status = %s,
                payment_method = %s,
                po_number = %s,
                total_amount = %s,
                shipping_cost = %s,
                tax_amount = %s,
                internal_notes = %s
            WHERE order_id = %s
        """, (customer_id, order_date, fulfillment_status, payment_status, payment_method, po_number, total_amount, shipping_cost, calculated_tax, internal_notes, order_id))

        for item in items_to_insert:
            cur.execute("""
                INSERT INTO order_items (order_id, product_id, quantity, unit_price, discount, lot_number, lot_id)
                VALUES (%s, %s, %s, %s, %s, %s, %s)
            """, (order_id, item['product_id'], item['quantity'], item['unit_price'], item['discount'], item['lot_number'], item['lot_id']))

        if fulfillment_status == 'Fulfilled':
            helper_fulfill_order_items(cur, order_id)

        sync_lot_reservations(cur)

        # Overwrite any previously generated documents with the updated order details
        regenerate_order_documents(cur, order_id)

        conn.commit()
        flash(f"Order #{order_id} updated successfully! Associated documents updated.", "success")
    except Exception as e:
        conn.rollback()
        flash(f"Error updating order: {str(e)}", "danger")
    finally:
        cur.close()
        conn.close()

    return redirect(url_for('index'))

@app.route('/create_quote', methods=['POST'])
def create_quote():
    conn = get_db()
    cur = conn.cursor(cursor_factory=RealDictCursor)
    
    try:
        customer_id = helper_resolve_customer(cur, request.form)

        # Parse Quote Details
        expires_at = request.form.get('expires_at') or None
        estimated_shipping = float(request.form.get('estimated_shipping') or 0.0)
        
        # Tax entered in % (e.g. 13.00 for 13%)
        tax_rate = float(request.form.get('tax_rate') or 0.0)
        customer_notes = request.form.get('customer_notes', '').strip() or None

        # Parse Line Items
        product_ids = request.form.getlist('product_id[]')
        quantities = request.form.getlist('quantity[]')
        unit_prices = request.form.getlist('unit_price[]')
        discounts = request.form.getlist('discount[]')

        items_to_insert = []
        calculated_subtotal = 0.0

        for i, p_id in enumerate(product_ids):
            if not p_id or not p_id.strip():
                continue
            try:
                pid_int = int(p_id)
                qty_int = int(quantities[i]) if i < len(quantities) and int(quantities[i]) > 0 else 1
                price_flt = float(unit_prices[i]) if i < len(unit_prices) and unit_prices[i] else 0.0
                disc_pct = float(discounts[i]) if i < len(discounts) and discounts[i] else 0.0
                
                line_total = max(0.0, qty_int * price_flt * (1.0 - disc_pct / 100.0))
                calculated_subtotal += line_total
                
                items_to_insert.append({
                    'product_id': pid_int,
                    'quantity': qty_int,
                    'unit_price': price_flt,
                    'discount': disc_pct  # store as percentage (e.g. 10 for 10%)
                })
            except (ValueError, TypeError):
                continue

        if not items_to_insert:
            conn.rollback()
            flash("Please add at least one valid product to the quote.", "warning")
            return redirect(url_for('index'))

        # Tax calculated from percentage
        estimated_tax = round(calculated_subtotal * (tax_rate / 100.0), 2)
        total_amount = round(calculated_subtotal + estimated_shipping + estimated_tax, 2)

        q_base, q_ver, q_num = generate_next_quote_number(cur, date.today())
        cur.execute("""
            INSERT INTO quotes (
                customer_id, quote_number, quote_base, quote_version,
                created_at, expires_at, total_amount, status,
                estimated_shipping, estimated_tax, customer_notes
            )
            VALUES (%s, %s, %s, %s, CURRENT_TIMESTAMP, COALESCE(%s::date, CURRENT_DATE + INTERVAL '30 days'), %s, 'Pending', %s, %s, %s)
            RETURNING quote_id
        """, (customer_id, q_num, q_base, q_ver, expires_at, total_amount, estimated_shipping, estimated_tax, customer_notes))
        new_quote_id = cur.fetchone()['quote_id']

        for item in items_to_insert:
            cur.execute("""
                INSERT INTO quote_items (quote_id, product_id, quantity, unit_price, discount)
                VALUES (%s, %s, %s, %s, %s)
            """, (new_quote_id, item['product_id'], item['quantity'], item['unit_price'], item['discount']))

        conn.commit()
        flash(f"Quote {q_num} created successfully!", "success")
    except Exception as e:
        conn.rollback()
        flash(f"Error creating quote: {str(e)}", "danger")
    finally:
        cur.close()
        conn.close()

    return redirect(url_for('index'))

@app.route('/edit_quote/<int:quote_id>', methods=['POST'])
def edit_quote(quote_id):
    """
    Updates an existing quote and its line items WITHOUT modifying any related order.
    """
    conn = get_db()
    cur = conn.cursor(cursor_factory=RealDictCursor)
    try:
        cur.execute("SELECT * FROM quotes WHERE quote_id = %s", (quote_id,))
        existing_quote = cur.fetchone()
        if not existing_quote:
            flash("Quote not found.", "danger")
            return redirect(url_for('index'))

        customer_id = int(request.form.get('customer_id') or request.form.get('existing_customer_id') or existing_quote['customer_id'])
        expires_at = request.form.get('expires_at') or None
        estimated_shipping = float(request.form.get('estimated_shipping') or 0.0)
        tax_rate = float(request.form.get('tax_rate') or 0.0)
        customer_notes = request.form.get('customer_notes', '').strip() or None

        product_ids = request.form.getlist('product_id[]')
        quantities = request.form.getlist('quantity[]')
        unit_prices = request.form.getlist('unit_price[]')
        discounts = request.form.getlist('discount[]')

        items_to_insert = []
        calculated_subtotal = 0.0

        for i, p_id in enumerate(product_ids):
            if not p_id or not p_id.strip():
                continue
            try:
                pid_int = int(p_id)
                qty_int = int(quantities[i]) if i < len(quantities) and int(quantities[i]) > 0 else 1
                price_flt = float(unit_prices[i]) if i < len(unit_prices) and unit_prices[i] else 0.0
                disc_pct = float(discounts[i]) if i < len(discounts) and discounts[i] else 0.0

                line_total = max(0.0, qty_int * price_flt * (1.0 - disc_pct / 100.0))
                calculated_subtotal += line_total

                items_to_insert.append({
                    'product_id': pid_int,
                    'quantity': qty_int,
                    'unit_price': price_flt,
                    'discount': disc_pct
                })
            except (ValueError, TypeError):
                continue

        if not items_to_insert:
            conn.rollback()
            flash("Quote must have at least one valid product.", "warning")
            return redirect(url_for('index'))

        estimated_tax = round(calculated_subtotal * (tax_rate / 100.0), 2)
        total_amount = round(calculated_subtotal + estimated_shipping + estimated_tax, 2)

        cur.execute("SELECT quote_base, quote_version, quote_number FROM quotes WHERE quote_id = %s", (quote_id,))
        q_row = cur.fetchone()
        q_base = (q_row['quote_base'] if (q_row and q_row['quote_base']) else f"Q{date.today().strftime('%y%m%d')}01")
        new_version = ((q_row['quote_version'] or 1) + 1) if q_row else 2
        new_quote_num = f"{q_base}V{new_version:02d}"

        cur.execute("""
            UPDATE quotes
            SET customer_id = %s,
                expires_at = COALESCE(%s::date, expires_at),
                total_amount = %s,
                estimated_shipping = %s,
                estimated_tax = %s,
                customer_notes = %s,
                quote_base = %s,
                quote_version = %s,
                quote_number = %s
            WHERE quote_id = %s
        """, (customer_id, expires_at, total_amount, estimated_shipping, estimated_tax, customer_notes, q_base, new_version, new_quote_num, quote_id))

        cur.execute("DELETE FROM quote_items WHERE quote_id = %s", (quote_id,))
        for item in items_to_insert:
            cur.execute("""
                INSERT INTO quote_items (quote_id, product_id, quantity, unit_price, discount)
                VALUES (%s, %s, %s, %s, %s)
            """, (quote_id, item['product_id'], item['quantity'], item['unit_price'], item['discount']))

        # Overwrite previously generated quote spreadsheet with updated quote details
        regenerate_quote_document(cur, quote_id)

        conn.commit()
        flash(f"Quote {new_quote_num} updated successfully! Associated documents updated.", "success")
    except Exception as e:
        conn.rollback()
        flash(f"Error updating quote: {str(e)}", "danger")
    finally:
        cur.close()
        conn.close()

    return redirect(url_for('index'))

# -------------------------------------------------------------------------
# ROUTES: STATUS UPDATES & INVENTORY SYNCHRONIZATION
# -------------------------------------------------------------------------

@app.route('/order/<int:order_id>/update_status', methods=['POST'])
def update_order_status(order_id):
    conn = get_db()
    cur = conn.cursor(cursor_factory=RealDictCursor)
    try:
        # Check current status before update
        cur.execute("SELECT fulfillment_status, payment_status FROM orders WHERE order_id = %s", (order_id,))
        current_order = cur.fetchone()
        old_fulfillment = current_order['fulfillment_status'] if current_order else None

        fulfillment_status = request.form.get('fulfillment_status')
        payment_status = request.form.get('payment_status')
        
        updates = []
        params = []
        if fulfillment_status:
            updates.append("fulfillment_status = %s")
            params.append(fulfillment_status)
        if payment_status:
            updates.append("payment_status = %s")
            params.append(payment_status)

        if updates:
            params.append(order_id)
            query = f"UPDATE orders SET {', '.join(updates)} WHERE order_id = %s RETURNING fulfillment_status, payment_status"
            cur.execute(query, tuple(params))
            updated = cur.fetchone()

            # Handle Inventory Transitions
            if fulfillment_status and fulfillment_status != old_fulfillment:
                if fulfillment_status == 'Fulfilled' and old_fulfillment != 'Fulfilled':
                    # Deduct physical stock and release reservation
                    helper_fulfill_order_items(cur, order_id)
                elif fulfillment_status in ('Pending', 'Unfulfilled') and old_fulfillment in ('Fulfilled', 'Partially Fulfilled'):
                    # Revert deduction and restore reservation
                    helper_unfulfill_order_items(cur, order_id)

            sync_lot_reservations(cur)
            conn.commit()

        if request.headers.get('X-Requested-With') == 'XMLHttpRequest' or request.is_json:
            return jsonify({
                "success": True,
                "order_id": order_id,
                "fulfillment_status": updated['fulfillment_status'] if updates else None,
                "payment_status": updated['payment_status'] if updates else None
            })
        flash(f"Order #{order_id} status updated successfully.", "success")
    except Exception as e:
        conn.rollback()
        if request.headers.get('X-Requested-With') == 'XMLHttpRequest' or request.is_json:
            return jsonify({"success": False, "error": str(e)}), 400
        flash(f"Error updating order status: {str(e)}", "danger")
    finally:
        cur.close()
        conn.close()

    return redirect(url_for('index'))

@app.route('/order/<int:order_id>/update_item_lot', methods=['POST'])
def update_order_item_lot(order_id):
    order_item_id = request.form.get('order_item_id')
    lot_number = request.form.get('lot_number', '').strip() or None
    
    conn = get_db()
    cur = conn.cursor(cursor_factory=RealDictCursor)
    try:
        # Find matching lot_id
        lot_id = None
        if lot_number:
            cur.execute("""
                SELECT pl.lot_id 
                FROM product_lots pl
                JOIN order_items oi ON pl.product_id = oi.product_id
                WHERE oi.order_item_id = %s AND pl.lot_number = %s
            """, (order_item_id, lot_number))
            row = cur.fetchone()
            if row:
                lot_id = row['lot_id']

        cur.execute("""
            UPDATE order_items 
            SET lot_number = %s, lot_id = %s 
            WHERE order_item_id = %s AND order_id = %s
        """, (lot_number, lot_id, order_item_id, order_id))
        sync_lot_reservations(cur)
        conn.commit()
        return jsonify({"success": True, "lot_number": lot_number, "lot_id": lot_id})
    except Exception as e:
        conn.rollback()
        return jsonify({"success": False, "error": str(e)}), 400
    finally:
        cur.close()
        conn.close()

# -------------------------------------------------------------------------
# ROUTES: DETAILED JSON APIS
# -------------------------------------------------------------------------

@app.route('/api/order/<int:order_id>')
def api_get_order(order_id):
    conn = get_db()
    cur = conn.cursor(cursor_factory=RealDictCursor)
    
    cur.execute("""
        SELECT c.*, o.*, q.quote_number
        FROM orders o
        LEFT JOIN customers c ON o.customer_id = c.customer_id
        LEFT JOIN quotes q ON o.quote_id = q.quote_id
        WHERE o.order_id = %s
    """, (order_id,))
    order = cur.fetchone()
    
    if not order:
        cur.close()
        conn.close()
        return jsonify({"error": "Order not found"}), 404
        
    cur.execute("""
        SELECT oi.*, p.sku, p.sub_sku, p.name, p.size, p.price AS catalog_price
        FROM order_items oi
        JOIN products p ON oi.product_id = p.product_id
        WHERE oi.order_id = %s
        ORDER BY oi.order_item_id ASC
    """, (order_id,))
    items = cur.fetchall()

    for item in items:
        cur.execute("""
            SELECT lot_id, lot_number, expiry_date, stock, reserved,
                   (stock - reserved) AS available
            FROM product_lots 
            WHERE product_id = %s 
            ORDER BY expiry_date ASC
        """, (item['product_id'],))
        item['available_lots'] = cur.fetchall()

    # Check for existing generated documents in order subfolder
    order_subfolder = os.path.join(OUTPUT_DIR, f"order_{order_id}")
    generated_docs = []
    if os.path.exists(order_subfolder):
        for f in os.listdir(order_subfolder):
            if f.endswith('.xlsx'):
                fpath = os.path.join(order_subfolder, f)
                stat = os.stat(fpath)
                doc_type = 'invoice' if 'invoice' in f else ('packing_slip' if 'packing' in f else ('label' if 'label' in f else 'excel'))
                generated_docs.append({
                    "filename": f,
                    "filepath": f"order_{order_id}/{f}",
                    "type": doc_type,
                    "size_kb": f"{round(stat.st_size / 1024, 1)} KB",
                    "modified": datetime.fromtimestamp(stat.st_mtime).strftime('%Y-%m-%d %H:%M')
                })

    last_fiscal_seq = get_last_fiscal_seq(cur)
    is_first_invoice = (last_fiscal_seq is None)
    suggested_seq = (last_fiscal_seq + 1) if last_fiscal_seq is not None else 1
    today_yymmdd = date.today().strftime('%y%m%d')
    suggested_inv_num = f"IN{today_yymmdd}{suggested_seq:03d}"
    next_packing_slip = generate_next_packing_slip_number(cur)

    cur.close()
    conn.close()
    
    return jsonify({
        "order": order,
        "items": items,
        "documents": generated_docs,
        "last_fiscal_seq": last_fiscal_seq,
        "is_first_invoice": is_first_invoice,
        "suggested_seq": suggested_seq,
        "today_yymmdd": today_yymmdd,
        "suggested_invoice_number": suggested_inv_num,
        "next_packing_slip_number": next_packing_slip
    })

@app.route('/api/quote/<int:quote_id>')
def api_get_quote(quote_id):
    conn = get_db()
    cur = conn.cursor(cursor_factory=RealDictCursor)
    cur.execute("""
        SELECT c.*, q.*, COALESCE(c.name, 'Deleted Customer') AS customer_name,
               (SELECT o.order_id FROM orders o WHERE o.quote_id = q.quote_id ORDER BY o.order_id DESC LIMIT 1) AS linked_order_id,
               (SELECT o.order_number FROM orders o WHERE o.quote_id = q.quote_id ORDER BY o.order_id DESC LIMIT 1) AS linked_order_number
        FROM quotes q
        LEFT JOIN customers c ON q.customer_id = c.customer_id
        WHERE q.quote_id = %s
    """, (quote_id,))
    quote = cur.fetchone()
    
    if not quote:
        cur.close()
        conn.close()
        return jsonify({"error": "Quote not found"}), 404

    cur.execute("""
        SELECT qi.*, p.sku, p.sub_sku, p.name, p.size, p.price AS catalog_price
        FROM quote_items qi
        JOIN products p ON qi.product_id = p.product_id
        WHERE qi.quote_id = %s
        ORDER BY qi.quote_item_id ASC
    """, (quote_id,))
    items = cur.fetchall()

    # Check for existing generated documents in quote subfolder
    quote_subfolder = os.path.join(OUTPUT_DIR, f"quote_{quote_id}")
    generated_docs = []
    if os.path.exists(quote_subfolder):
        for f in os.listdir(quote_subfolder):
            if f.endswith('.xlsx'):
                fpath = os.path.join(quote_subfolder, f)
                stat = os.stat(fpath)
                generated_docs.append({
                    "filename": f,
                    "filepath": f"quote_{quote_id}/{f}",
                    "type": "quote",
                    "size_kb": f"{round(stat.st_size / 1024, 1)} KB",
                    "modified": datetime.fromtimestamp(stat.st_mtime).strftime('%Y-%m-%d %H:%M')
                })

    cur.close()
    conn.close()
    return jsonify({
        "quote": quote,
        "customer_name": quote['customer_name'],
        "total_amount": quote['total_amount'],
        "customer_notes": quote['customer_notes'],
        "items": items,
        "documents": generated_docs
    })

@app.route('/api/customer/<int:customer_id>')
def api_get_customer(customer_id):
    conn = get_db()
    cur = conn.cursor(cursor_factory=RealDictCursor)

    cur.execute("SELECT * FROM customers WHERE customer_id = %s", (customer_id,))
    customer = cur.fetchone()
    if not customer:
        cur.close()
        conn.close()
        return jsonify({"error": "Customer not found"}), 404

    cur.execute("""
        SELECT o.*, q.quote_number
        FROM orders o
        LEFT JOIN quotes q ON o.quote_id = q.quote_id
        WHERE o.customer_id = %s
        ORDER BY o.order_date DESC, o.order_id DESC
    """, (customer_id,))
    orders = cur.fetchall()
    for o in orders:
        if o.get('order_date'):
            o['order_date_str'] = o['order_date'].strftime('%Y-%m-%d')
        else:
            o['order_date_str'] = 'N/A'

    cur.execute("""
        SELECT q.*,
               (SELECT o.order_id FROM orders o WHERE o.quote_id = q.quote_id ORDER BY o.order_id DESC LIMIT 1) AS linked_order_id,
               (SELECT o.order_number FROM orders o WHERE o.quote_id = q.quote_id ORDER BY o.order_id DESC LIMIT 1) AS linked_order_number
        FROM quotes q
        WHERE q.customer_id = %s
        ORDER BY q.created_at DESC, q.quote_id DESC
    """, (customer_id,))
    quotes = cur.fetchall()
    for q in quotes:
        if q.get('created_at'):
            q['created_at_str'] = q['created_at'].strftime('%Y-%m-%d')
        else:
            q['created_at_str'] = 'N/A'
        if q.get('expires_at'):
            q['expires_at_str'] = q['expires_at'].strftime('%Y-%m-%d')
        else:
            q['expires_at_str'] = 'N/A'

    total_spent = sum(float(o.get('total_amount') or 0.0) for o in orders)

    cur.close()
    conn.close()
    return jsonify({
        "customer": customer,
        "orders": orders,
        "quotes": quotes,
        "total_spent": round(total_spent, 2)
    })

# -------------------------------------------------------------------------
# ROUTES: MANUAL DELETION WITH CASCADE
# -------------------------------------------------------------------------

@app.route('/delete_order/<int:order_id>', methods=['POST'])
def delete_order(order_id):
    delete_files = request.form.get('delete_files') == '1'
    conn = get_db()
    cur = conn.cursor(cursor_factory=RealDictCursor)
    try:
        cur.execute("""
            SELECT o.fulfillment_status, o.invoice_number, o.quote_id, o.order_number,
                   q.quote_number
            FROM orders o
            LEFT JOIN quotes q ON o.quote_id = q.quote_id
            WHERE o.order_id = %s
        """, (order_id,))
        order = cur.fetchone()
        inv_str = f" ({order['invoice_number']})" if order and order['invoice_number'] else ""
        linked_quote_id = order['quote_id'] if order else None
        o_name = (order['order_number'] if order and order['order_number'] else f"#{order_id}")
        q_name = (order['quote_number'] if order and order['quote_number'] else (f"#{linked_quote_id}" if linked_quote_id else None))

        # Release stock reservation if order was not fulfilled
        if order and order['fulfillment_status'] != 'Fulfilled':
            helper_release_order_reservations(cur, order_id)

        # Delete order items first (FK)
        cur.execute("DELETE FROM order_items WHERE order_id = %s", (order_id,))
        # Delete order
        cur.execute("DELETE FROM orders WHERE order_id = %s", (order_id,))

        # Delete the linked quote (and its items) if one exists
        if linked_quote_id:
            cur.execute("DELETE FROM quote_items WHERE quote_id = %s", (linked_quote_id,))
            cur.execute("DELETE FROM quotes WHERE quote_id = %s", (linked_quote_id,))

        sync_lot_reservations(cur)
        conn.commit()

        # Delete associated files if requested
        if delete_files:
            subfolder = os.path.join(OUTPUT_DIR, f"order_{order_id}")
            if os.path.exists(subfolder):
                shutil.rmtree(subfolder)
            if linked_quote_id:
                q_subfolder = os.path.join(OUTPUT_DIR, f"quote_{linked_quote_id}")
                if os.path.exists(q_subfolder):
                    shutil.rmtree(q_subfolder)

        quote_str = f" and Quote {q_name}" if q_name else ""
        flash(f"Order {o_name}{inv_str}{quote_str} deleted successfully.", "success")
    except Exception as e:
        conn.rollback()
        flash(f"Error deleting order: {str(e)}", "danger")
    finally:
        cur.close()
        conn.close()

    return redirect(url_for('index'))

@app.route('/delete_quote/<int:quote_id>', methods=['POST'])
def delete_quote(quote_id):
    delete_files = request.form.get('delete_files') == '1'
    conn = get_db()
    cur = conn.cursor(cursor_factory=RealDictCursor)
    try:
        cur.execute("SELECT quote_number FROM quotes WHERE quote_id = %s", (quote_id,))
        q_row = cur.fetchone()
        q_name = q_row['quote_number'] if (q_row and q_row['quote_number']) else f"#{quote_id}"

        # Find any order linked to this quote
        cur.execute("SELECT order_id, order_number, fulfillment_status, invoice_number FROM orders WHERE quote_id = %s", (quote_id,))
        linked_order = cur.fetchone()

        if linked_order:
            order_id = linked_order['order_id']
            inv_str = f" ({linked_order['invoice_number']})" if linked_order['invoice_number'] else ""
            # Release stock reservation if order was not fulfilled
            if linked_order['fulfillment_status'] != 'Fulfilled':
                helper_release_order_reservations(cur, order_id)
            cur.execute("DELETE FROM order_items WHERE order_id = %s", (order_id,))
            cur.execute("DELETE FROM orders WHERE order_id = %s", (order_id,))

        # Delete quote items first (FK)
        cur.execute("DELETE FROM quote_items WHERE quote_id = %s", (quote_id,))
        # Delete quote
        cur.execute("DELETE FROM quotes WHERE quote_id = %s", (quote_id,))
        sync_lot_reservations(cur)
        conn.commit()

        # Delete associated files if requested
        if delete_files:
            subfolder = os.path.join(OUTPUT_DIR, f"quote_{quote_id}")
            if os.path.exists(subfolder):
                shutil.rmtree(subfolder)
            if linked_order:
                o_subfolder = os.path.join(OUTPUT_DIR, f"order_{linked_order['order_id']}")
                if os.path.exists(o_subfolder):
                    shutil.rmtree(o_subfolder)

        o_name = (linked_order['order_number'] or f"#{linked_order['order_id']}") if linked_order else ""
        order_str = f" and Order {o_name}{inv_str}" if linked_order else ""
        flash(f"Quote {q_name}{order_str} deleted successfully.", "success")
    except Exception as e:
        conn.rollback()
        flash(f"Error deleting quote: {str(e)}", "danger")
    finally:
        cur.close()
        conn.close()

    return redirect(url_for('index'))

@app.route('/edit_customer/<int:customer_id>', methods=['POST'])
def edit_customer(customer_id):
    name = request.form.get('name', '').strip()
    if not name:
        flash("Customer name is required.", "danger")
        return redirect(url_for('index'))

    email = request.form.get('email', '').strip() or None
    phone = request.form.get('phone', '').strip() or None
    lab = request.form.get('lab', '').strip() or None
    org = request.form.get('org', '').strip() or None
    dep = request.form.get('dep', '').strip() or None
    billing_name = request.form.get('billing_name', '').strip() or name

    billing_street = request.form.get('billing_street', '').strip()
    billing_unit = request.form.get('billing_unit', '').strip() or None
    billing_city = request.form.get('billing_city', '').strip()
    billing_province = request.form.get('billing_province', '').strip()
    billing_postal_code = request.form.get('billing_postal_code', '').strip()
    billing_country = request.form.get('billing_country', '').strip() or 'Canada'

    shipping_street = request.form.get('shipping_street', '').strip() or billing_street
    shipping_unit = request.form.get('shipping_unit', '').strip() or billing_unit
    shipping_city = request.form.get('shipping_city', '').strip() or billing_city
    shipping_province = request.form.get('shipping_province', '').strip() or billing_province
    shipping_postal_code = request.form.get('shipping_postal_code', '').strip() or billing_postal_code
    shipping_country = request.form.get('shipping_country', '').strip() or billing_country

    if not email and not phone:
        phone = 'N/A'

    conn = get_db()
    cur = conn.cursor(cursor_factory=RealDictCursor)
    try:
        cur.execute("""
            UPDATE customers
            SET name = %s, email = %s, phone = %s, lab = %s, org = %s, dep = %s, billing_name = %s,
                billing_street = %s, billing_unit = %s, billing_city = %s, billing_province = %s, billing_postal_code = %s, billing_country = %s,
                shipping_street = %s, shipping_unit = %s, shipping_city = %s, shipping_province = %s, shipping_postal_code = %s, shipping_country = %s
            WHERE customer_id = %s
        """, (
            name, email, phone, lab, org, dep, billing_name,
            billing_street, billing_unit, billing_city, billing_province, billing_postal_code, billing_country,
            shipping_street, shipping_unit, shipping_city, shipping_province, shipping_postal_code, shipping_country,
            customer_id
        ))

        # Overwrite any documents for orders and quotes associated with this customer
        cur.execute("SELECT order_id FROM orders WHERE customer_id = %s", (customer_id,))
        for o_row in cur.fetchall():
            regenerate_order_documents(cur, o_row['order_id'])

        cur.execute("SELECT quote_id FROM quotes WHERE customer_id = %s", (customer_id,))
        for q_row in cur.fetchall():
            regenerate_quote_document(cur, q_row['quote_id'])

        conn.commit()
        flash(f"Customer '{name}' updated successfully! Associated documents updated.", "success")
    except Exception as e:
        conn.rollback()
        flash(f"Error updating customer: {str(e)}", "danger")
    finally:
        cur.close()
        conn.close()

    return redirect(url_for('index'))

@app.route('/create_customer', methods=['POST'])
def create_customer():
    conn = get_db()
    cur = conn.cursor(cursor_factory=RealDictCursor)
    try:
        # Force customer_type to new so helper_resolve_customer reads the new_* fields
        form_dict = dict(request.form)
        form_dict['customer_type'] = 'new'
        new_cust_id = helper_resolve_customer(cur, form_dict)
        conn.commit()
        name = request.form.get('new_name', '').strip()
        flash(f"Customer '{name}' created successfully!", "success")
    except Exception as e:
        conn.rollback()
        flash(f"Error creating customer: {str(e)}", "danger")
    finally:
        cur.close()
        conn.close()

    return redirect(url_for('index'))

@app.route('/delete_customer/<int:customer_id>', methods=['POST'])
def delete_customer(customer_id):
    conn = get_db()
    cur = conn.cursor(cursor_factory=RealDictCursor)
    try:
        cur.execute("SELECT name FROM customers WHERE customer_id = %s", (customer_id,))
        cust = cur.fetchone()
        cust_name = cust['name'] if cust else f"#{customer_id}"

        # Quotes have ON DELETE CASCADE so quote_items and quotes are automatically removed
        # Orders have ON DELETE SET NULL; if customer is deleted, orders remain with customer_id=NULL or user can delete orders first
        cur.execute("DELETE FROM customers WHERE customer_id = %s", (customer_id,))
        conn.commit()
        flash(f"Customer '{cust_name}' deleted successfully.", "success")
    except Exception as e:
        conn.rollback()
        flash(f"Error deleting customer: {str(e)}", "danger")
    finally:
        cur.close()
        conn.close()

    return redirect(url_for('index'))

# -------------------------------------------------------------------------
# ROUTES: CONVERT QUOTE TO ORDER
# -------------------------------------------------------------------------

@app.route('/convert_quote_to_order/<int:quote_id>', methods=['POST'])
def convert_quote_to_order(quote_id):
    conn = get_db()
    cur = conn.cursor(cursor_factory=RealDictCursor)
    try:
        cur.execute("SELECT * FROM quotes WHERE quote_id = %s", (quote_id,))
        quote = cur.fetchone()
        
        cur.execute("SELECT * FROM quote_items WHERE quote_id = %s", (quote_id,))
        quote_items = cur.fetchall()
        
        order_num = generate_next_order_number(cur, date.today())
        cur.execute("""
            INSERT INTO orders (
                customer_id, order_number, order_date, total_amount, fulfillment_status, 
                payment_status, invoice_number, shipping_cost, tax_amount, internal_notes, quote_id
            )
            VALUES (%s, %s, CURRENT_DATE, %s, 'Pending', 'Unpaid', NULL, %s, %s, %s, %s)
            RETURNING order_id
        """, (
            quote['customer_id'], order_num, quote['total_amount'], 
            quote['estimated_shipping'], quote['estimated_tax'],
            None, quote_id
        ))
        new_order_id = cur.fetchone()['order_id']
        
        for item in quote_items:
            # Allocate FEFO lot reservation
            lot_num, lot_id = helper_reserve_lot_fefo(cur, item['product_id'], item['quantity'])
            cur.execute("""
                INSERT INTO order_items (order_id, product_id, quantity, unit_price, discount, lot_number, lot_id)
                VALUES (%s, %s, %s, %s, %s, %s, %s)
            """, (new_order_id, item['product_id'], item['quantity'], item['unit_price'], item['discount'], lot_num, lot_id))
            
        cur.execute("UPDATE quotes SET status = 'Converted' WHERE quote_id = %s", (quote_id,))
        sync_lot_reservations(cur)
        conn.commit()

        # Move any documents in quote_<quote_id> into order_<new_order_id> so docs are in the same folder
        q_dir = os.path.join(OUTPUT_DIR, f"quote_{quote_id}")
        o_dir = os.path.join(OUTPUT_DIR, f"order_{new_order_id}")
        if os.path.exists(q_dir):
            os.makedirs(o_dir, exist_ok=True)
            for f in os.listdir(q_dir):
                src = os.path.join(q_dir, f)
                dst = os.path.join(o_dir, f)
                shutil.move(src, dst)
            try:
                os.rmdir(q_dir)
            except Exception:
                pass

        # Re-render quote document inside order folder to include Order # reference
        regenerate_quote_document(cur, quote_id)

        q_disp = quote.get('quote_number') or f"#{quote_id}"
        flash(f"Quote {q_disp} converted to Order {order_num} successfully! Stock reserved.", "success")
    except Exception as e:
        conn.rollback()
        flash(f"Error converting quote: {str(e)}", "danger")
    finally:
        cur.close()
        conn.close()
        
    return redirect(url_for('index'))

# -------------------------------------------------------------------------
# ROUTES: DOCUMENT GENERATION (DIRECT DOWNLOAD)
# -------------------------------------------------------------------------

@app.route('/generate_quote_excel/<int:quote_id>', methods=['GET', 'POST'])
def generate_quote_excel(quote_id):
    conn = get_db()
    cur = conn.cursor(cursor_factory=RealDictCursor)
    
    if request.method == 'POST':
        updated_amount = request.form.get('total_amount')
        updated_notes = request.form.get('customer_notes')
        
        cur.execute("SELECT total_amount, customer_notes, quote_base, quote_version, quote_number FROM quotes WHERE quote_id = %s", (quote_id,))
        cur_q = cur.fetchone()
        if cur_q:
            amount_changed = (float(updated_amount) != float(cur_q['total_amount'] or 0)) if updated_amount else False
            notes_changed = (str(updated_notes or '').strip() != str(cur_q['customer_notes'] or '').strip())
            if amount_changed or notes_changed:
                q_base = cur_q.get('quote_base') or f"Q{date.today().strftime('%y%m%d')}01"
                new_version = (cur_q.get('quote_version') or 1) + 1
                new_quote_num = f"{q_base}V{new_version:02d}"
                cur.execute("""
                    UPDATE quotes 
                    SET total_amount = %s, customer_notes = %s, quote_base = %s, quote_version = %s, quote_number = %s 
                    WHERE quote_id = %s
                """, (updated_amount, updated_notes, q_base, new_version, new_quote_num, quote_id))
            else:
                cur.execute("""
                    UPDATE quotes 
                    SET total_amount = %s, customer_notes = %s 
                    WHERE quote_id = %s
                """, (updated_amount, updated_notes, quote_id))
            conn.commit()

    cur.execute("""
        SELECT c.*, q.*, COALESCE(c.name, 'Deleted Customer') AS customer_name,
               (SELECT o.order_id FROM orders o WHERE o.quote_id = q.quote_id ORDER BY o.order_id DESC LIMIT 1) AS order_id,
               (SELECT o.order_number FROM orders o WHERE o.quote_id = q.quote_id ORDER BY o.order_id DESC LIMIT 1) AS order_number
        FROM quotes q
        LEFT JOIN customers c ON q.customer_id = c.customer_id
        WHERE q.quote_id = %s
    """, (quote_id,))
    quote = cur.fetchone()

    if quote and not quote.get('quote_number'):
        q_base, q_ver, q_num = generate_next_quote_number(cur, quote.get('created_at'))
        cur.execute("UPDATE quotes SET quote_base=%s, quote_version=%s, quote_number=%s WHERE quote_id=%s", (q_base, q_ver, q_num, quote_id))
        conn.commit()
        quote['quote_number'] = q_num
        quote['quote_base'] = q_base
        quote['quote_version'] = q_ver
    
    cur.execute("""
        SELECT qi.*, p.sku, p.sub_sku, p.name, p.size 
        FROM quote_items qi
        JOIN products p ON qi.product_id = p.product_id
        WHERE qi.quote_id = %s
    """, (quote_id,))
    items = cur.fetchall()
    
    cur.close()
    conn.close()
    
    # If quote has a converted order, store in the order folder so all docs are together
    subfolder = f"order_{quote['order_id']}" if (quote and quote.get('order_id')) else f"quote_{quote_id}"
    filename = generate_excel_document('quote', quote, items, subfolder=subfolder)
    target_dir = os.path.join(OUTPUT_DIR, subfolder)
    return send_from_directory(target_dir, filename, as_attachment=True)

@app.route('/generate_order_document/<int:order_id>/<doc_type>', methods=['GET', 'POST'])
def generate_order_document(order_id, doc_type):
    if doc_type not in ['invoice', 'packing_slip', 'label']:
        return "Invalid document type", 400

    conn = get_db()
    cur = conn.cursor(cursor_factory=RealDictCursor)
    
    # Check for parameters (from modal via GET or POST)
    pay_method = request.values.get('payment_method')
    po_num = request.values.get('po_number')
    est_weight = request.values.get('estimated_weight')
    ice_charge_in = request.values.get('ice_charge')
    
    if pay_method:
        cur.execute("""
            UPDATE orders 
            SET payment_method = %s, po_number = %s 
            WHERE order_id = %s
        """, (pay_method.strip(), (po_num.strip() if po_num else None), order_id))
        conn.commit()

    # If ice charge was submitted for invoice, remember it in database
    if doc_type == 'invoice' and ice_charge_in is not None and str(ice_charge_in).strip() != '':
        try:
            ice_val = float(ice_charge_in)
            cur.execute("""
                UPDATE orders 
                SET ice_charge = %s 
                WHERE order_id = %s
            """, (ice_val, order_id))
            conn.commit()
        except (ValueError, TypeError):
            pass

    cur.execute("""
        SELECT c.*, o.*, COALESCE(c.name, 'Deleted Customer') AS customer_name,
               q.quote_number
        FROM orders o
        LEFT JOIN customers c ON o.customer_id = c.customer_id
        LEFT JOIN quotes q ON o.quote_id = q.quote_id
        WHERE o.order_id = %s
    """, (order_id,))
    order_data = cur.fetchone()
    if not order_data:
        cur.close()
        conn.close()
        flash("Order not found.", "danger")
        return redirect(url_for('index'))

    # If parameters were supplied in this request, merge into order_data
    if pay_method:
        order_data['payment_method'] = pay_method.strip()
        order_data['po_number'] = (po_num.strip() if po_num else None)
    if est_weight:
        order_data['estimated_weight'] = est_weight.strip()

    # Ensure order_number is present
    if not order_data.get('order_number'):
        new_ord_num = generate_next_order_number(cur, order_data.get('order_date'))
        cur.execute("UPDATE orders SET order_number = %s WHERE order_id = %s", (new_ord_num, order_id))
        conn.commit()
        order_data['order_number'] = new_ord_num

    # Handle Invoice number assignment & remembering
    if doc_type == 'invoice':
        existing_inv = order_data.get('invoice_number')
        if existing_inv:
            inv_num = existing_inv
        else:
            today_yymmdd = date.today().strftime('%y%m%d')
            user_seq = request.values.get('fiscal_order_seq')
            user_inv_num = request.values.get('invoice_number', '').strip()

            seq_val = None
            if user_seq and str(user_seq).strip().isdigit():
                seq_val = int(str(user_seq).strip())
            elif user_inv_num:
                m = re.match(r'^IN\d{6}(\d+)$', user_inv_num, re.IGNORECASE)
                if m:
                    seq_val = int(m.group(1))
                elif user_inv_num.isdigit():
                    seq_val = int(user_inv_num)

            if seq_val is None:
                last_seq = get_last_fiscal_seq(cur)
                seq_val = (last_seq + 1) if last_seq is not None else 1

            inv_num = f"IN{today_yymmdd}{seq_val:03d}"
            cur.execute("UPDATE orders SET invoice_number = %s WHERE order_id = %s", (inv_num, order_id))
            set_last_fiscal_seq(cur, seq_val)
            conn.commit()

        order_data['invoice_number'] = inv_num

    cur.execute("""
        SELECT oi.*, p.sku, p.sub_sku, p.name, p.size 
        FROM order_items oi
        JOIN products p ON oi.product_id = p.product_id
        WHERE oi.order_id = %s
        ORDER BY oi.order_item_id ASC
    """, (order_id,))
    all_items = cur.fetchall()

    # Packing slip partial shipment handling
    items_to_render = all_items
    if doc_type == 'packing_slip':
        # Delivery date from modal (defaults to today)
        deliv_date_raw = request.values.get('delivery_date')
        if deliv_date_raw:
            try:
                deliv_date = datetime.strptime(deliv_date_raw.strip(), '%Y-%m-%d').date()
            except Exception:
                deliv_date = date.today()
        else:
            deliv_date = date.today()

        cur.execute("UPDATE orders SET last_packing_slip_date = %s WHERE order_id = %s", (deliv_date, order_id))
        order_data['last_packing_slip_date'] = deliv_date
        order_data['delivery_date'] = deliv_date

        # Slip number from modal or auto-generated
        custom_slip_num = request.values.get('slip_number', '').strip()
        if not custom_slip_num:
            custom_slip_num = generate_next_packing_slip_number(cur, deliv_date)
        order_data['slip_number'] = custom_slip_num

        try:
            cur.execute("""
                INSERT INTO packing_slips (order_id, slip_number, delivery_date)
                VALUES (%s, %s, %s)
                ON CONFLICT (slip_number) DO UPDATE SET delivery_date = EXCLUDED.delivery_date
            """, (order_id, custom_slip_num, deliv_date))
            conn.commit()
        except Exception:
            pass
        
        # Check if selected item IDs were passed
        selected_item_ids = request.values.getlist('selected_items[]') or request.values.getlist('selected_items')
        if not selected_item_ids and request.values.get('selected_item_ids'):
            selected_item_ids = [s.strip() for s in request.values.get('selected_item_ids').split(',') if s.strip()]

        if selected_item_ids:
            # Parse selected integer IDs
            try:
                sel_ids_set = set(int(x) for x in selected_item_ids)
                items_to_render = [dict(it) for it in all_items if it['order_item_id'] in sel_ids_set]
            except (ValueError, TypeError):
                items_to_render = [dict(it) for it in all_items]
        else:
            # By default, exclude already fulfilled items if this order has partial fulfillments
            remaining_items = [dict(it) for it in all_items if (it.get('fulfilled_quantity') or 0) < it['quantity']]
            if remaining_items:
                items_to_render = remaining_items

        # Apply specific quantity shipped per selected item
        is_partial = (len(items_to_render) < len(all_items))
        for it in items_to_render:
            oid_key = f"item_qty_{it['order_item_id']}"
            raw_qty = request.values.get(oid_key)
            current_ful = it.get('fulfilled_quantity') or 0
            max_remaining = max(1, it['quantity'] - current_ful)

            if raw_qty is not None and str(raw_qty).isdigit():
                qty_to_ship = min(max_remaining, max(1, int(raw_qty)))
            else:
                qty_to_ship = max_remaining

            # Display this slip's shipped quantity on the packing slip Excel
            it['quantity'] = qty_to_ship

            orig_qty = next((orig['quantity'] for orig in all_items if orig['order_item_id'] == it['order_item_id']), it['quantity'])
            if qty_to_ship < orig_qty:
                is_partial = True

        mark_as_fulfilled = (request.values.get('mark_as_fulfilled') == '1')
        if mark_as_fulfilled:
            for it in items_to_render:
                cur.execute("""
                    UPDATE order_items 
                    SET fulfilled_quantity = LEAST(quantity, fulfilled_quantity + %s)
                    WHERE order_item_id = %s
                """, (it['quantity'], it['order_item_id']))
            
            # Check if all items in the order are now fulfilled or only some
            cur.execute("""
                SELECT COUNT(*) AS total_items,
                       SUM(CASE WHEN fulfilled_quantity >= quantity THEN 1 ELSE 0 END) AS fulfilled_items,
                       SUM(CASE WHEN fulfilled_quantity > 0 THEN 1 ELSE 0 END) AS any_fulfilled_items
                FROM order_items 
                WHERE order_id = %s
            """, (order_id,))
            f_counts = cur.fetchone()
            tot = f_counts['total_items'] or 0
            all_ful = f_counts['fulfilled_items'] or 0
            any_ful = f_counts['any_fulfilled_items'] or 0

            new_status = 'Pending'
            if all_ful >= tot and tot > 0:
                new_status = 'Fulfilled'
                helper_fulfill_order_items(cur, order_id)
            elif any_ful > 0:
                new_status = 'Partially Fulfilled'

            cur.execute("UPDATE orders SET fulfillment_status = %s WHERE order_id = %s", (new_status, order_id))
            order_data['fulfillment_status'] = new_status
            sync_lot_reservations(cur)
            conn.commit()

        # Handle Packing Slip Delivery Notes:
        # User requirement:
        # 1. Do NOT put "converted from quote#" in delivery notes
        # 2. Allow optional note entered in the modal
        # 3. If it is a partial delivery, note that in delivery notes, otherwise leave blank
        custom_note = request.values.get('delivery_notes', '').strip()
        
        note_parts = []
        if is_partial:
            note_parts.append("Partial Delivery: Remaining items to ship separately.")
        if custom_note:
            note_parts.append(custom_note)
        
        order_data['internal_notes'] = "\n".join(note_parts) if note_parts else ""

    cur.close()
    conn.close()
    
    filename = generate_excel_document(doc_type, order_data, items_to_render, subfolder=f"order_{order_id}")
    target_dir = os.path.join(OUTPUT_DIR, f"order_{order_id}")
    return send_from_directory(target_dir, filename, as_attachment=True)

# -------------------------------------------------------------------------
# ROUTES: WAREHOUSE (STOCK, LOT TRACKING & DEFICIT/EXPIRY ALERTS)
# -------------------------------------------------------------------------

@app.route('/warehouse')
def warehouse_page():
    conn = get_db()
    cur = conn.cursor(cursor_factory=RealDictCursor)

    # Ensure reserved quantities are always in sync with active unfulfilled orders
    sync_lot_reservations(cur)
    conn.commit()
    
    # Fetch products with physical stock and reserved quantities
    # Items with active orders (not Fulfilled/Cancelled) are reserved!
    cur.execute("""
        SELECT p.*, 
               COALESCE(SUM(pl.stock), 0) AS total_stock,
               COALESCE((
                   SELECT SUM(
                       CASE 
                           WHEN o.fulfillment_status = 'Partially Fulfilled' THEN GREATEST(0, oi.quantity - COALESCE(oi.fulfilled_quantity, 0))
                           WHEN o.fulfillment_status IN ('Fulfilled', 'Cancelled') THEN 0
                           ELSE oi.quantity
                       END
                   )
                   FROM order_items oi
                   JOIN orders o ON oi.order_id = o.order_id
                   WHERE o.fulfillment_status NOT IN ('Fulfilled', 'Cancelled')
                     AND oi.product_id = p.product_id
               ), 0) AS total_reserved,
               COUNT(pl.lot_id) AS total_lots
        FROM products p
        LEFT JOIN product_lots pl ON p.product_id = pl.product_id
        GROUP BY p.product_id
        ORDER BY p.sku ASC, p.sub_sku ASC
    """)
    products = cur.fetchall()

    # Fetch all active lots with days until expiry
    cur.execute("""
        SELECT pl.*, p.sku, p.sub_sku, p.name AS product_name, p.size,
               (pl.stock - pl.reserved) AS available,
               (pl.expiry_date - CURRENT_DATE) AS days_left
        FROM product_lots pl
        JOIN products p ON pl.product_id = p.product_id
        ORDER BY pl.expiry_date ASC, pl.lot_number ASC
    """)
    all_lots = cur.fetchall()

    # Attach lots to products and compute critical alert flags
    lots_by_product = {}
    critical_expiring = []
    
    for lot in all_lots:
        pid = lot['product_id']
        if pid not in lots_by_product:
            lots_by_product[pid] = []
        lots_by_product[pid].append(lot)
        
        # Expiring within 60 days
        if lot['stock'] > 0 and lot['days_left'] is not None and lot['days_left'] <= 60:
            lot['is_expiring_soon'] = True
            critical_expiring.append(lot)
        else:
            lot['is_expiring_soon'] = False

    critical_shortages = []
    for p in products:
        p['lots'] = lots_by_product.get(p['product_id'], [])
        p['available_stock'] = int(p['total_stock']) - int(p['total_reserved'])
        p['has_shortage'] = (p['available_stock'] < 0)
        p['has_expiring'] = any(l.get('is_expiring_soon') for l in p['lots'])
        if p['has_shortage']:
            critical_shortages.append(p)

    # Group products by base SKU
    sku_groups = {}
    suppliers_set = set()
    for p in products:
        p['supplier'] = (p.get('supplier') or 'Default Supplier').strip() or 'Default Supplier'
        suppliers_set.add(p['supplier'])
        base_sku = p['sku']
        if base_sku not in sku_groups:
            sku_groups[base_sku] = {
                'sku': base_sku,
                'name': p['name'],
                'supplier': p['supplier'],
                'products': [],
                'group_total_stock': 0,
                'group_total_reserved': 0,
                'group_available_stock': 0,
                'has_stock': False,
                'has_alert': False
            }
        sku_groups[base_sku]['products'].append(p)
        sku_groups[base_sku]['group_total_stock'] += int(p['total_stock'] or 0)
        sku_groups[base_sku]['group_total_reserved'] += int(p['total_reserved'] or 0)
        if p['has_shortage'] or p['has_expiring']:
            sku_groups[base_sku]['has_alert'] = True

    suppliers = sorted(list(suppliers_set))

    group_list = list(sku_groups.values())
    for g in group_list:
        g['group_available_stock'] = g['group_total_stock'] - g['group_total_reserved']
        g['has_stock'] = (g['group_total_stock'] > 0)
        # Inside group: prioritize items with reserved stock first, then in-stock, then sub_sku
        g['products'].sort(key=lambda x: (
            0 if int(x['total_reserved'] or 0) > 0 else 1,
            0 if int(x['total_stock'] or 0) > 0 else 1,
            str(x['sub_sku'] or '')
        ))

    # Sort groups:
    # 1. Groups with reserved stock prioritized FIRST to the very top!
    # 2. Alert groups (shortage/expiring)
    # 3. In-stock products prioritized
    # 4. Alphabetically by base SKU
    group_list.sort(key=lambda x: (
        0 if int(x['group_total_reserved'] or 0) > 0 else 1,
        0 if x['has_alert'] else 1,
        0 if x['has_stock'] else 1,
        x['sku'].lower()
    ))

    # High-level metrics
    total_products_count = len(products)
    total_stock_count = sum(int(p['total_stock'] or 0) for p in products)
    total_reserved_count = sum(int(p['total_reserved'] or 0) for p in products)
    active_lots_count = len(all_lots)
    net_available_count = total_stock_count - total_reserved_count
    attention_count = len(critical_shortages) + len(critical_expiring)

    cur.close()
    conn.close()
    return render_template(
        'warehouse.html',
        products=products,
        sku_groups=group_list,
        suppliers=suppliers,
        critical_shortages=critical_shortages,
        critical_expiring=critical_expiring,
        attention_count=attention_count,
        total_products=total_products_count,
        total_stock=total_stock_count,
        total_reserved=total_reserved_count,
        net_available=net_available_count,
        active_lots=active_lots_count,
        active_page='warehouse'
    )

@app.route('/warehouse/add_stock', methods=['POST'])
def warehouse_add_stock():
    product_id = request.form.get('product_id')
    lot_number = request.form.get('lot_number', '').strip()
    expiry_date = request.form.get('expiry_date')
    stock = request.form.get('stock', '0')
    received_date = request.form.get('received_date') or None

    if not product_id or not lot_number or not expiry_date:
        flash("Product, Lot Number, and Expiry Date are required.", "danger")
        return redirect(url_for('warehouse_page'))

    conn = get_db()
    cur = conn.cursor(cursor_factory=RealDictCursor)
    try:
        stock_int = int(stock)
        cur.execute("""
            INSERT INTO product_lots (product_id, lot_number, expiry_date, stock, reserved, received_date)
            VALUES (%s, %s, %s, %s, 0, COALESCE(%s::date, CURRENT_DATE))
            ON CONFLICT (product_id, lot_number) DO UPDATE
            SET stock = product_lots.stock + EXCLUDED.stock,
                expiry_date = EXCLUDED.expiry_date
        """, (product_id, lot_number, expiry_date, stock_int, received_date))
        sync_lot_reservations(cur)
        conn.commit()
        flash(f"Added {stock_int} units for Lot #{lot_number} successfully.", "success")
    except Exception as e:
        conn.rollback()
        flash(f"Error adding stock: {str(e)}", "danger")
    finally:
        cur.close()
        conn.close()

    return redirect(url_for('warehouse_page'))

@app.route('/warehouse/update_lot', methods=['POST'])
def warehouse_update_lot():
    lot_id = request.form.get('lot_id')
    stock = request.form.get('stock')
    expiry_date = request.form.get('expiry_date')
    
    conn = get_db()
    cur = conn.cursor(cursor_factory=RealDictCursor)
    try:
        cur.execute("""
            UPDATE product_lots 
            SET stock = %s, expiry_date = %s 
            WHERE lot_id = %s
        """, (int(stock), expiry_date, lot_id))
        conn.commit()
        flash("Lot stock updated successfully.", "success")
    except Exception as e:
        conn.rollback()
        flash(f"Error updating lot: {str(e)}", "danger")
    finally:
        cur.close()
        conn.close()

    return redirect(url_for('warehouse_page'))

@app.route('/warehouse/delete_lot/<int:lot_id>', methods=['POST'])
def warehouse_delete_lot(lot_id):
    conn = get_db()
    cur = conn.cursor(cursor_factory=RealDictCursor)
    try:
        cur.execute("SELECT lot_number, reserved FROM product_lots WHERE lot_id = %s", (lot_id,))
        lot = cur.fetchone()
        if not lot:
            flash("Stock lot not found.", "warning")
            return redirect(url_for('warehouse_page'))
        
        lot_num = lot['lot_number']
        # If order items reference this lot_id, set them to NULL so they don't block deletion
        cur.execute("UPDATE order_items SET lot_id = NULL WHERE lot_id = %s", (lot_id,))
        cur.execute("DELETE FROM product_lots WHERE lot_id = %s", (lot_id,))
        sync_lot_reservations(cur)
        conn.commit()
        flash(f"Lot #{lot_num} deleted successfully.", "success")
    except Exception as e:
        conn.rollback()
        flash(f"Error deleting lot: {str(e)}", "danger")
    finally:
        cur.close()
        conn.close()

    return redirect(url_for('warehouse_page'))

@app.route('/warehouse/edit_product/<int:product_id>', methods=['POST'])
def warehouse_edit_product(product_id):
    sku = request.form.get('sku', '').strip()
    sub_sku = request.form.get('sub_sku', '').strip()
    name = request.form.get('name', '').strip()
    size = request.form.get('size', '').strip()
    price = request.form.get('price', '0.00').strip()
    supplier_choice = request.form.get('supplier_choice', '').strip()
    new_supplier = request.form.get('new_supplier', '').strip()
    supplier = new_supplier if (supplier_choice == '__new__' and new_supplier) else supplier_choice

    if not sku or not name or not size or not supplier:
        flash("SKU, Name, Size, and Supplier are required.", "danger")
        return redirect(url_for('warehouse_page'))

    if '-' in sku and not sub_sku:
        sku, sub_sku = sku.split('-', 1)

    conn = get_db()
    cur = conn.cursor(cursor_factory=RealDictCursor)
    try:
        cur.execute("""
            UPDATE products 
            SET sku = %s, sub_sku = %s, name = %s, size = %s, price = %s, supplier = %s 
            WHERE product_id = %s
        """, (sku.strip(), sub_sku.strip(), name, size, float(price), supplier.strip(), product_id))
        conn.commit()
        flash(f"Product {format_full_sku(sku, sub_sku)} ({name}) updated successfully.", "success")
    except Exception as e:
        conn.rollback()
        flash(f"Error updating product: {str(e)}", "danger")
    finally:
        cur.close()
        conn.close()

    return redirect(url_for('warehouse_page'))

@app.route('/warehouse/delete_product/<int:product_id>', methods=['POST'])
def warehouse_delete_product(product_id):
    conn = get_db()
    cur = conn.cursor(cursor_factory=RealDictCursor)
    try:
        cur.execute("SELECT sku, sub_sku, name FROM products WHERE product_id = %s", (product_id,))
        p = cur.fetchone()
        if not p:
            flash("Product not found.", "warning")
            return redirect(url_for('warehouse_page'))

        cat_display = format_full_sku(p['sku'], p['sub_sku'])
        # Check if product is in order_items or quote_items
        cur.execute("SELECT COUNT(*) AS count FROM order_items WHERE product_id = %s", (product_id,))
        o_count = cur.fetchone()['count']
        cur.execute("SELECT COUNT(*) AS count FROM quote_items WHERE product_id = %s", (product_id,))
        q_count = cur.fetchone()['count']

        if o_count > 0 or q_count > 0:
            flash(f"Cannot delete product {cat_display} because it is referenced in {o_count} order line(s) and {q_count} quote line(s).", "danger")
            return redirect(url_for('warehouse_page'))

        # Delete any associated lots first
        cur.execute("DELETE FROM product_lots WHERE product_id = %s", (product_id,))
        cur.execute("DELETE FROM products WHERE product_id = %s", (product_id,))
        conn.commit()
        flash(f"Product {cat_display} ({p['name']}) and its stock lots were deleted successfully.", "success")
    except Exception as e:
        conn.rollback()
        flash(f"Error deleting product: {str(e)}", "danger")
    finally:
        cur.close()
        conn.close()

    return redirect(url_for('warehouse_page'))

@app.route('/warehouse/add_product', methods=['POST'])
def warehouse_add_product():
    sku = request.form.get('sku', '').strip()
    sub_sku = request.form.get('sub_sku', '').strip()
    name = request.form.get('name', '').strip()
    size = request.form.get('size', '').strip()
    price = request.form.get('price', '0.00').strip()
    
    supplier_choice = request.form.get('supplier_choice', '').strip()
    new_supplier = request.form.get('new_supplier', '').strip()
    supplier = new_supplier if (supplier_choice == '__new__' and new_supplier) else supplier_choice

    if not sku or not name or not size or not supplier:
        flash("SKU, Name, Size, and Supplier are required.", "danger")
        return redirect(url_for('warehouse_page'))

    if '-' in sku and not sub_sku:
        sku, sub_sku = sku.split('-', 1)

    conn = get_db()
    cur = conn.cursor(cursor_factory=RealDictCursor)
    try:
        cur.execute("""
            INSERT INTO products (sku, sub_sku, name, size, price, supplier)
            VALUES (%s, %s, %s, %s, %s, %s)
            ON CONFLICT (sku, sub_sku) DO UPDATE
            SET name = EXCLUDED.name, size = EXCLUDED.size, price = EXCLUDED.price, supplier = EXCLUDED.supplier
            RETURNING product_id
        """, (sku.strip(), sub_sku.strip(), name, size, float(price), supplier.strip()))
        conn.commit()
        full_cat = format_full_sku(sku, sub_sku)
        flash(f"Product {full_cat} ({name} - {size}) [{supplier}] saved successfully.", "success")
    except Exception as e:
        conn.rollback()
        flash(f"Error adding product: {str(e)}", "danger")
    finally:
        cur.close()
        conn.close()

    return redirect(url_for('warehouse_page'))

@app.route('/warehouse/import_excel', methods=['POST'])
def warehouse_import_excel():
    uploaded_file = request.files.get('excel_file')
    supplier_choice = request.form.get('supplier_choice', '').strip()
    new_supplier = request.form.get('new_supplier', '').strip()
    supplier = new_supplier if (supplier_choice == '__new__' and new_supplier) else supplier_choice
    if not supplier:
        supplier = request.form.get('supplier', 'Default Supplier').strip() or 'Default Supplier'

    if not uploaded_file or not uploaded_file.filename:
        flash("Please select an Excel (.xlsx) file to upload.", "warning")
        return redirect(url_for('warehouse_page'))

    if not uploaded_file.filename.lower().endswith('.xlsx'):
        flash("Only .xlsx Excel files are supported.", "danger")
        return redirect(url_for('warehouse_page'))

    temp_path = os.path.join(OUTPUT_DIR, f"_temp_import_{int(datetime.now().timestamp())}.xlsx")
    try:
        uploaded_file.save(temp_path)
        from import_excel import import_excel_file
        summary = import_excel_file(temp_path, supplier=supplier)
        sheets_str = ', '.join(summary['sheets_matched']) if summary.get('sheets_matched') else 'None'
        flash(
            f"Catalog sync complete ({supplier} — sheets: {sheets_str})! "
            f"Scanned {summary['total_processed']} SKUs: "
            f"{summary['products_added']} new added, {summary['products_updated']} updated, "
            f"{summary['products_unchanged']} unchanged.",
            "success"
        )
    except Exception as e:
        flash(f"Error importing Excel file: {str(e)}", "danger")
    finally:
        if os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except OSError:
                pass

    return redirect(url_for('warehouse_page'))

# -------------------------------------------------------------------------
# ROUTES: DOCUMENTS REPOSITORY
# -------------------------------------------------------------------------

@app.route('/documents')
def documents_page():
    conn = get_db()
    cur = conn.cursor(cursor_factory=RealDictCursor)

    # Fetch lookup dicts for orders and quotes
    cur.execute("""
        SELECT o.order_id, o.order_number, o.quote_id, o.invoice_number, COALESCE(c.name, 'Deleted Customer') AS customer_name,
               q.quote_number
        FROM orders o
        LEFT JOIN customers c ON o.customer_id = c.customer_id
        LEFT JOIN quotes q ON o.quote_id = q.quote_id
    """)
    orders_map = {row['order_id']: row for row in cur.fetchall()}

    cur.execute("""
        SELECT q.quote_id, q.quote_number, q.status, COALESCE(c.name, 'Deleted Customer') AS customer_name 
        FROM quotes q
        LEFT JOIN customers c ON q.customer_id = c.customer_id
    """)
    quotes_map = {row['quote_id']: row for row in cur.fetchall()}
    cur.close()
    conn.close()

    folders_dict = {}
    all_docs = []
    total_bytes = 0

    if os.path.exists(OUTPUT_DIR):
        for item in sorted(os.listdir(OUTPUT_DIR)):
            item_path = os.path.join(OUTPUT_DIR, item)
            if not os.path.isdir(item_path):
                continue

            folder_key = item
            folder_title = item
            customer_name = 'N/A'
            invoice_number = None
            linked_quote_id = None
            order_id = None
            is_converted = False

            if folder_key.startswith('order_'):
                oid_str = folder_key.replace('order_', '')
                if oid_str.isdigit():
                    order_id = int(oid_str)
                    if order_id in orders_map:
                        o_info = orders_map[order_id]
                        customer_name = o_info['customer_name']
                        invoice_number = o_info['invoice_number']
                        linked_quote_id = o_info.get('quote_id')
                        o_name = o_info.get('order_number') or f"Order #{order_id}"
                        q_name = o_info.get('quote_number') or (f"Quote #{linked_quote_id}" if linked_quote_id else None)
                        if q_name:
                            folder_title = f"{o_name} ({q_name})"
                            is_converted = True
                        else:
                            folder_title = o_name
            elif folder_key.startswith('quote_'):
                qid_str = folder_key.replace('quote_', '')
                if qid_str.isdigit():
                    qid = int(qid_str)
                    if qid in quotes_map:
                        q_info = quotes_map[qid]
                        customer_name = q_info['customer_name']
                        q_name = q_info.get('quote_number') or f"Quote #{qid}"
                        folder_title = f"{q_name} ({q_info.get('status', 'Active')})"

            folder_files = []
            folder_bytes = 0
            folder_mtime = 0

            for f in sorted(os.listdir(item_path)):
                if not f.endswith('.xlsx'):
                    continue
                fpath = os.path.join(item_path, f)
                stat = os.stat(fpath)
                folder_bytes += stat.st_size
                total_bytes += stat.st_size
                if stat.st_mtime > folder_mtime:
                    folder_mtime = stat.st_mtime

                rel_filepath = os.path.relpath(fpath, OUTPUT_DIR).replace('\\', '/')

                doc_type = 'Excel Document'
                badge_class = 'bg-secondary'
                if f.startswith('invoice_'):
                    doc_type = 'Invoice'
                    badge_class = 'bg-primary'
                elif f.startswith('packing_slip_'):
                    doc_type = 'Packing Slip'
                    badge_class = 'bg-warning text-dark'
                elif f.startswith('label_'):
                    doc_type = 'Shipping Label'
                    badge_class = 'bg-info text-dark'
                elif f.startswith('quote_'):
                    doc_type = 'Quote'
                    badge_class = 'bg-success'

                doc_obj = {
                    'filename': f,
                    'filepath': rel_filepath,
                    'folder': folder_key,
                    'entity_label': folder_title,
                    'customer_name': customer_name,
                    'invoice_number': invoice_number,
                    'doc_type': doc_type,
                    'badge_class': badge_class,
                    'size_kb': f"{round(stat.st_size / 1024, 1)} KB",
                    'size_bytes': stat.st_size,
                    'modified_timestamp': stat.st_mtime,
                    'modified_str': datetime.fromtimestamp(stat.st_mtime).strftime('%Y-%m-%d %H:%M')
                }
                folder_files.append(doc_obj)
                all_docs.append(doc_obj)

            if folder_files:
                folders_dict[folder_key] = {
                    'folder_key': folder_key,
                    'folder_title': folder_title,
                    'customer_name': customer_name,
                    'invoice_number': invoice_number,
                    'order_id': order_id,
                    'linked_quote_id': linked_quote_id,
                    'is_converted': is_converted,
                    'file_count': len(folder_files),
                    'total_size_kb': f"{round(folder_bytes / 1024, 1)} KB",
                    'total_bytes': folder_bytes,
                    'latest_mtime': folder_mtime,
                    'modified_str': datetime.fromtimestamp(folder_mtime).strftime('%Y-%m-%d %H:%M') if folder_mtime else '—',
                    'files': folder_files
                }

    folders_list = list(folders_dict.values())
    folders_list.sort(key=lambda x: x['latest_mtime'], reverse=True)
    all_docs.sort(key=lambda x: x['modified_timestamp'], reverse=True)

    highlight = request.args.get('highlight', '')
    open_now = request.args.get('open') == '1'
    total_size_mb = f"{round(total_bytes / (1024 * 1024), 2)} MB"

    return render_template(
        'documents.html',
        folders=folders_list,
        documents=all_docs,
        total_folders=len(folders_list),
        total_docs=len(all_docs),
        total_size_mb=total_size_mb,
        highlight=highlight,
        open_now=open_now,
        active_page='documents'
    )

@app.route('/api/backup/status')
def api_backup_status():
    drives = backup_manager.detect_available_drives()
    history = backup_manager.get_backup_history()
    last_backup = history[0] if history else None
    return jsonify({
        'drives': drives,
        'last_backup': last_backup,
        'history_count': len(history)
    })

@app.route('/api/backup/run', methods=['POST'])
def api_backup_run():
    target_path = request.form.get('target_path') or (request.json.get('target_path') if request.is_json else None)
    if target_path and target_path.strip():
        target_path = target_path.strip()
    else:
        target_path = None
    try:
        res = backup_manager.perform_backup(target_root=target_path)
        return jsonify(res)
    except Exception as e:
        return jsonify({'status': 'ERROR', 'error': str(e)}), 500

@app.route('/documents/download/<path:filepath>')
def documents_download(filepath):
    norm_path = os.path.normpath(filepath)
    if norm_path.startswith('..') or os.path.isabs(norm_path):
        abort(403)
    
    full_path = os.path.join(OUTPUT_DIR, norm_path)
    if not os.path.exists(full_path):
        abort(404)

    folder_dir = os.path.dirname(full_path)
    filename = os.path.basename(full_path)
    return send_from_directory(folder_dir, filename, as_attachment=True)

@app.route('/documents/delete', methods=['POST'])
def documents_delete():
    filepath = request.form.get('filepath', '').strip()
    norm_path = os.path.normpath(filepath)
    if norm_path.startswith('..') or os.path.isabs(norm_path):
        abort(403)

    full_path = os.path.join(OUTPUT_DIR, norm_path)
    if os.path.exists(full_path):
        try:
            os.remove(full_path)
            flash(f"Deleted {os.path.basename(full_path)} successfully.", "success")
        except Exception as e:
            flash(f"Error deleting file: {str(e)}", "danger")
    else:
        flash("File not found.", "warning")

    return redirect(url_for('documents_page'))

@app.route('/documents/delete-folder', methods=['POST'])
def documents_delete_folder():
    folder_key = request.form.get('folder_key', '').strip()
    norm_key = os.path.normpath(folder_key)
    if not norm_key or norm_key in ('.', '..') or norm_key.startswith('..') or os.path.isabs(norm_key):
        abort(403)

    full_folder_path = os.path.join(OUTPUT_DIR, norm_key)
    if os.path.exists(full_folder_path) and os.path.isdir(full_folder_path):
        try:
            shutil.rmtree(full_folder_path)
            flash(f"Folder '{folder_key}' and all its documents were deleted successfully.", "success")
        except Exception as e:
            flash(f"Error deleting folder: {str(e)}", "danger")
    else:
        flash(f"Folder '{folder_key}' not found.", "warning")

    return redirect(url_for('documents_page'))

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)