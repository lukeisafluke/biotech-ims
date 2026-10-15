-- =============================================================================
-- Inventory & Order Management Database Schema (PostgreSQL)
-- =============================================================================

CREATE TABLE IF NOT EXISTS customers (
    customer_id SERIAL PRIMARY KEY,
    name VARCHAR(100) NOT NULL,
    lab VARCHAR(100),
    org VARCHAR(100),
    dep VARCHAR(100),
    email VARCHAR(100),
    phone VARCHAR(20),
    billing_name VARCHAR(100),
    billing_street VARCHAR(255) NOT NULL,
    billing_unit VARCHAR(50),
    billing_city VARCHAR(100) NOT NULL,
    billing_province VARCHAR(50) NOT NULL,
    billing_postal_code VARCHAR(20) NOT NULL,
    billing_country VARCHAR(100) NOT NULL DEFAULT 'Canada',
    shipping_street VARCHAR(255) NOT NULL,
    shipping_unit VARCHAR(50),
    shipping_city VARCHAR(100) NOT NULL,
    shipping_province VARCHAR(50) NOT NULL,
    shipping_postal_code VARCHAR(20) NOT NULL,
    shipping_country VARCHAR(100) NOT NULL DEFAULT 'Canada',
    CONSTRAINT check_contact_method CHECK (((email IS NOT NULL) OR (phone IS NOT NULL)))
);

CREATE TABLE IF NOT EXISTS products (
    product_id SERIAL PRIMARY KEY,
    sku VARCHAR(50) NOT NULL,
    sub_sku VARCHAR(50) NOT NULL DEFAULT '',
    name VARCHAR(255) NOT NULL,
    size VARCHAR(255) NOT NULL,
    price NUMERIC NOT NULL,
    supplier VARCHAR(100) NOT NULL,
    CONSTRAINT unique_product_variant UNIQUE (sku, sub_sku)
);

CREATE TABLE IF NOT EXISTS product_lots (
    lot_id SERIAL PRIMARY KEY,
    product_id INTEGER NOT NULL REFERENCES products(product_id) ON DELETE RESTRICT,
    lot_number VARCHAR(100) NOT NULL,
    expiry_date DATE NOT NULL,
    stock INTEGER NOT NULL DEFAULT 0,
    received_date DATE DEFAULT CURRENT_DATE,
    reserved INTEGER NOT NULL DEFAULT 0,
    CONSTRAINT unique_product_lot UNIQUE (product_id, lot_number)
);

CREATE TABLE IF NOT EXISTS quotes (
    quote_id SERIAL PRIMARY KEY,
    customer_id INTEGER REFERENCES customers(customer_id) ON DELETE CASCADE,
    quote_number VARCHAR(50) UNIQUE,
    quote_base VARCHAR(50),
    quote_version INTEGER NOT NULL DEFAULT 1,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    expires_at DATE DEFAULT (CURRENT_DATE + INTERVAL '30 days'),
    total_amount NUMERIC DEFAULT 0.00,
    status VARCHAR(50) NOT NULL DEFAULT 'Active',
    estimated_shipping NUMERIC DEFAULT 0.00,
    estimated_tax NUMERIC DEFAULT 0.00,
    customer_notes TEXT
);

CREATE TABLE IF NOT EXISTS quote_items (
    quote_item_id SERIAL PRIMARY KEY,
    quote_id INTEGER REFERENCES quotes(quote_id) ON DELETE CASCADE,
    product_id INTEGER REFERENCES products(product_id) ON DELETE RESTRICT,
    quantity INTEGER NOT NULL DEFAULT 1,
    unit_price NUMERIC NOT NULL,
    discount NUMERIC DEFAULT 0.00
);

CREATE TABLE IF NOT EXISTS orders (
    order_id SERIAL PRIMARY KEY,
    customer_id INTEGER REFERENCES customers(customer_id) ON DELETE SET NULL,
    order_number VARCHAR(50) UNIQUE,
    order_date DATE DEFAULT CURRENT_DATE,
    total_amount NUMERIC DEFAULT 0.00,
    fulfillment_status VARCHAR(50) NOT NULL DEFAULT 'Pending',
    payment_status VARCHAR(50) NOT NULL DEFAULT 'Unpaid',
    invoice_number VARCHAR(50) UNIQUE,
    shipping_cost NUMERIC DEFAULT 0.00,
    tax_amount NUMERIC DEFAULT 0.00,
    internal_notes TEXT,
    quote_id INTEGER REFERENCES quotes(quote_id) ON DELETE SET NULL,
    payment_method VARCHAR(50),
    po_number VARCHAR(100),
    last_packing_slip_date DATE,
    ice_charge NUMERIC DEFAULT NULL
);

CREATE TABLE IF NOT EXISTS order_items (
    order_item_id SERIAL PRIMARY KEY,
    order_id INTEGER REFERENCES orders(order_id) ON DELETE CASCADE,
    product_id INTEGER REFERENCES products(product_id) ON DELETE RESTRICT,
    quantity INTEGER NOT NULL DEFAULT 1,
    unit_price NUMERIC NOT NULL,
    discount NUMERIC DEFAULT 0.00,
    lot_number VARCHAR(100),
    lot_id INTEGER REFERENCES product_lots(lot_id) ON DELETE SET NULL,
    fulfilled_quantity INTEGER NOT NULL DEFAULT 0
);

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

