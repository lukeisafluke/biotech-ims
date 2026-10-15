-- =============================================================================
-- Sample Database Data Seed: Inventory & Order Management System
-- =============================================================================
BEGIN;

-- Table: customers (2 records)
INSERT INTO customers (customer_id, name, lab, org, dep, email, phone, billing_name, billing_street, billing_unit, billing_city, billing_province, billing_postal_code, billing_country, shipping_street, shipping_unit, shipping_city, shipping_province, shipping_postal_code, shipping_country) VALUES (1, 'Acme Corporation', 'Main Lab', 'Acme Org', 'Purchasing', 'billing@acmecorp.example', '555-0199', 'Acme Accounts Payable', '123 Anvil Street', 'Suite 100', 'Toronto', 'Ontario', 'M5V 2L7', 'Canada', '456 Coyote Way', 'Dock 4', 'Toronto', 'Ontario', 'M5V 2L8', 'Canada') ON CONFLICT DO NOTHING;
INSERT INTO customers (customer_id, name, lab, org, dep, email, phone, billing_name, billing_street, billing_unit, billing_city, billing_province, billing_postal_code, billing_country, shipping_street, shipping_unit, shipping_city, shipping_province, shipping_postal_code, shipping_country) VALUES (9, 'BioHealth Technologies', 'Genomics Division', 'BioHealth Global', 'Procurement', 'orders@biohealth.example', '555-0142', 'BioHealth Accounts Payable', '789 Discovery Blvd', 'Suite 200', 'Vancouver', 'British Columbia', 'V6B 1A1', 'Canada', '789 Discovery Blvd', 'Dock 2', 'Vancouver', 'British Columbia', 'V6B 1A1', 'Canada') ON CONFLICT DO NOTHING;
SELECT setval('customers_customer_id_seq', COALESCE((SELECT MAX(customer_id) FROM customers), 1), true);

-- Table: products (2 records)
INSERT INTO products (product_id, sku, sub_sku, name, size, price) VALUES (1, 'WIDGET-001', 1, 'Standard Widget', 'Medium', 25.00) ON CONFLICT DO NOTHING;
INSERT INTO products (product_id, sku, sub_sku, name, size, price) VALUES (2, 'GIZMO-100', 2, 'Advanced Gizmo', 'Large', 85.50) ON CONFLICT DO NOTHING;
SELECT setval('products_product_id_seq', COALESCE((SELECT MAX(product_id) FROM products), 1), true);

-- Table: product_lots (4 records)
INSERT INTO product_lots (lot_id, product_id, lot_number, expiry_date, stock, received_date, reserved) VALUES (3, 2, 'GZ2601L', '2027-12-31', 15, '2026-09-13', 0) ON CONFLICT DO NOTHING;
INSERT INTO product_lots (lot_id, product_id, lot_number, expiry_date, stock, received_date, reserved) VALUES (2, 1, 'WG2602M', '2027-09-30', 0, '2026-09-13', 0) ON CONFLICT DO NOTHING;
INSERT INTO product_lots (lot_id, product_id, lot_number, expiry_date, stock, received_date, reserved) VALUES (16, 1, '73092300', '2027-10-21', 10, '2026-09-13', 2) ON CONFLICT DO NOTHING;
INSERT INTO product_lots (lot_id, product_id, lot_number, expiry_date, stock, received_date, reserved) VALUES (1, 1, 'WG2601M', '2027-06-30', 0, '2026-09-13', 2) ON CONFLICT DO NOTHING;
SELECT setval('product_lots_lot_id_seq', COALESCE((SELECT MAX(lot_id) FROM product_lots), 1), true);

-- Table: quotes (5 records)
INSERT INTO quotes (quote_id, customer_id, created_at, expires_at, total_amount, status, estimated_shipping, estimated_tax, customer_notes) VALUES (3, 1, '2026-09-13T04:33:08.224817', '2026-10-13', 161.38, 'Converted', 0.85, 0.03, '') ON CONFLICT DO NOTHING;
INSERT INTO quotes (quote_id, customer_id, created_at, expires_at, total_amount, status, estimated_shipping, estimated_tax, customer_notes) VALUES (1, 1, '2026-09-13T04:42:03.357346', '2026-10-13', 135.50, 'Converted', 10.00, 15.00, 'Initial test quote to verify Excel mapping.') ON CONFLICT DO NOTHING;
INSERT INTO quotes (quote_id, customer_id, created_at, expires_at, total_amount, status, estimated_shipping, estimated_tax, customer_notes) VALUES (9, 1, '2026-09-13T04:46:40.419979', '2026-10-15', 60.00, 'Pending', 10.00, 0.00, 'Test quote notes') ON CONFLICT DO NOTHING;
INSERT INTO quotes (quote_id, customer_id, created_at, expires_at, total_amount, status, estimated_shipping, estimated_tax, customer_notes) VALUES (10, 1, '2026-09-13T04:46:47.581083', '2026-10-15', 66.50, 'Converted', 10.00, 6.50, 'Test quote notes') ON CONFLICT DO NOTHING;
INSERT INTO quotes (quote_id, customer_id, created_at, expires_at, total_amount, status, estimated_shipping, estimated_tax, customer_notes) VALUES (18, 9, '2026-09-13T13:29:26.395893', '2026-09-12', 271.48, 'Pending', 50.00, 25.48, NULL) ON CONFLICT DO NOTHING;
SELECT setval('quotes_quote_id_seq', COALESCE((SELECT MAX(quote_id) FROM quotes), 1), true);

-- Table: quote_items (8 records)
INSERT INTO quote_items (quote_item_id, quote_id, product_id, quantity, unit_price, discount) VALUES (4, 3, 1, 3, 25.00, 0.00) ON CONFLICT DO NOTHING;
INSERT INTO quote_items (quote_item_id, quote_id, product_id, quantity, unit_price, discount) VALUES (5, 3, 2, 1, 85.50, 0.00) ON CONFLICT DO NOTHING;
INSERT INTO quote_items (quote_item_id, quote_id, product_id, quantity, unit_price, discount) VALUES (6, 1, 1, 2, 25.00, 0.00) ON CONFLICT DO NOTHING;
INSERT INTO quote_items (quote_item_id, quote_id, product_id, quantity, unit_price, discount) VALUES (7, 1, 2, 1, 85.50, 0.00) ON CONFLICT DO NOTHING;
INSERT INTO quote_items (quote_item_id, quote_id, product_id, quantity, unit_price, discount) VALUES (13, 9, 1, 2, 25.00, 0.00) ON CONFLICT DO NOTHING;
INSERT INTO quote_items (quote_item_id, quote_id, product_id, quantity, unit_price, discount) VALUES (14, 10, 1, 2, 25.00, 0.00) ON CONFLICT DO NOTHING;
INSERT INTO quote_items (quote_item_id, quote_id, product_id, quantity, unit_price, discount) VALUES (22, 18, 1, 1, 25.00, 0.00) ON CONFLICT DO NOTHING;
INSERT INTO quote_items (quote_item_id, quote_id, product_id, quantity, unit_price, discount) VALUES (23, 18, 2, 2, 85.50, 0.00) ON CONFLICT DO NOTHING;
SELECT setval('quote_items_quote_item_id_seq', COALESCE((SELECT MAX(quote_item_id) FROM quote_items), 1), true);

-- Table: orders (3 records)
INSERT INTO orders (order_id, customer_id, order_date, total_amount, fulfillment_status, payment_status, invoice_number, shipping_cost, tax_amount, internal_notes, quote_id) VALUES (15, NULL, '2026-09-13', 74.50, 'Pending', 'Unpaid', 'INV-0015', 18.00, 6.50, NULL, NULL) ON CONFLICT DO NOTHING;
INSERT INTO orders (order_id, customer_id, order_date, total_amount, fulfillment_status, payment_status, invoice_number, shipping_cost, tax_amount, internal_notes, quote_id) VALUES (13, 1, '2026-09-13', 66.50, 'Pending', 'Unpaid', 'INV-0011', 10.00, 6.50, 'Converted from Quote #11', NULL) ON CONFLICT DO NOTHING;
INSERT INTO orders (order_id, customer_id, order_date, total_amount, fulfillment_status, payment_status, invoice_number, shipping_cost, tax_amount, internal_notes, quote_id) VALUES (14, 1, '2026-09-13', 66.50, 'Fulfilled', 'Paid', 'INV-0010', 10.00, 6.50, 'Converted from Quote #10', 10) ON CONFLICT DO NOTHING;
SELECT setval('orders_order_id_seq', COALESCE((SELECT MAX(order_id) FROM orders), 1), true);

-- Table: order_items (3 records)
INSERT INTO order_items (order_item_id, order_id, product_id, quantity, unit_price, discount, lot_number, lot_id) VALUES (17, 13, 1, 2, 25.00, 0.00, 'WG2601M', 1) ON CONFLICT DO NOTHING;
INSERT INTO order_items (order_item_id, order_id, product_id, quantity, unit_price, discount, lot_number, lot_id) VALUES (18, 14, 1, 2, 25.00, 0.00, 'WG2601M', 1) ON CONFLICT DO NOTHING;
INSERT INTO order_items (order_item_id, order_id, product_id, quantity, unit_price, discount, lot_number, lot_id) VALUES (19, 15, 1, 2, 25.00, 0.00, '73092300', 16) ON CONFLICT DO NOTHING;
SELECT setval('order_items_order_item_id_seq', COALESCE((SELECT MAX(order_item_id) FROM order_items), 1), true);

COMMIT;
