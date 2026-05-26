PRAGMA foreign_keys = ON;

CREATE TABLE IF NOT EXISTS categories (
    category_id INTEGER PRIMARY KEY,
    owner TEXT NOT NULL,
    category TEXT NOT NULL,
    run_name TEXT NOT NULL,
    slug TEXT NOT NULL UNIQUE,
    active INTEGER NOT NULL DEFAULT 1,
    aliases_json TEXT NOT NULL DEFAULT '[]',
    notes TEXT,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS shopify_category_products (
    product_id INTEGER PRIMARY KEY,
    category_id INTEGER NOT NULL REFERENCES categories(category_id) ON DELETE CASCADE,
    sku TEXT,
    product_title TEXT NOT NULL,
    product_type TEXT,
    tags_json TEXT NOT NULL DEFAULT '[]',
    shopify_url TEXT,
    image_url TEXT,
    active_status TEXT,
    category_mapping TEXT,
    source TEXT NOT NULL,
    source_reference TEXT,
    created_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS shopify_spec_attributes (
    attribute_id INTEGER PRIMARY KEY,
    category_id INTEGER NOT NULL REFERENCES categories(category_id) ON DELETE CASCADE,
    sku TEXT,
    attribute_name TEXT NOT NULL,
    attribute_value TEXT NOT NULL,
    source TEXT NOT NULL,
    source_reference TEXT,
    created_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS category_attribute_distribution (
    distribution_id INTEGER PRIMARY KEY,
    category_id INTEGER NOT NULL REFERENCES categories(category_id) ON DELETE CASCADE,
    attribute_name TEXT NOT NULL,
    attribute_value TEXT,
    occurrence_count INTEGER NOT NULL DEFAULT 0,
    product_count INTEGER NOT NULL DEFAULT 0,
    coverage_pct REAL,
    signal_class TEXT NOT NULL,
    source TEXT NOT NULL,
    created_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS stackline_segments (
    segment_id INTEGER PRIMARY KEY,
    category_id INTEGER NOT NULL REFERENCES categories(category_id) ON DELETE CASCADE,
    channel TEXT NOT NULL,
    segment_name TEXT NOT NULL,
    retail_sales REAL,
    units_sold REAL,
    sales_share_pct REAL,
    growth_pct REAL,
    data_start_date TEXT,
    data_end_date TEXT,
    source TEXT NOT NULL,
    source_reference TEXT,
    created_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS stackline_top_products (
    top_product_id INTEGER PRIMARY KEY,
    category_id INTEGER NOT NULL REFERENCES categories(category_id) ON DELETE CASCADE,
    segment_id INTEGER REFERENCES stackline_segments(segment_id) ON DELETE SET NULL,
    channel TEXT NOT NULL,
    brand TEXT,
    title TEXT NOT NULL,
    asin_sku TEXT,
    url TEXT,
    price REAL,
    reviews INTEGER,
    rating REAL,
    retail_sales REAL,
    units_sold REAL,
    sales_share_pct REAL,
    source TEXT NOT NULL,
    source_reference TEXT,
    created_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS category_feature_signal_profile (
    profile_id INTEGER PRIMARY KEY,
    category_id INTEGER REFERENCES categories(category_id) ON DELETE CASCADE,
    profile_key TEXT NOT NULL,
    label TEXT NOT NULL,
    applies_to_json TEXT NOT NULL DEFAULT '[]',
    feature_signals_json TEXT NOT NULL DEFAULT '[]',
    source TEXT NOT NULL,
    source_reference TEXT,
    created_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS gap_evidence (
    evidence_id INTEGER PRIMARY KEY,
    category_id INTEGER NOT NULL REFERENCES categories(category_id) ON DELETE CASCADE,
    source_channel TEXT NOT NULL,
    recommendation TEXT NOT NULL,
    classification TEXT,
    priority TEXT,
    confidence TEXT,
    competitor_example TEXT,
    review_url TEXT,
    sunco_coverage_check TEXT,
    gap_rationale TEXT,
    pm_action TEXT,
    source_systems_json TEXT NOT NULL DEFAULT '[]',
    source_reference TEXT,
    local_image TEXT,
    created_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS refresh_audit (
    audit_id INTEGER PRIMARY KEY,
    source TEXT NOT NULL,
    run_date TEXT NOT NULL,
    data_age_days INTEGER,
    sql_used TEXT,
    row_counts_json TEXT NOT NULL DEFAULT '{}',
    warnings_json TEXT NOT NULL DEFAULT '[]',
    created_at TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_categories_slug ON categories(slug);
CREATE INDEX IF NOT EXISTS idx_products_category_sku ON shopify_category_products(category_id, sku);
CREATE INDEX IF NOT EXISTS idx_specs_category_attr ON shopify_spec_attributes(category_id, attribute_name);
CREATE INDEX IF NOT EXISTS idx_distribution_category_attr ON category_attribute_distribution(category_id, attribute_name);
CREATE INDEX IF NOT EXISTS idx_stackline_segments_category ON stackline_segments(category_id, channel);
CREATE INDEX IF NOT EXISTS idx_stackline_top_products_category ON stackline_top_products(category_id, channel);
CREATE INDEX IF NOT EXISTS idx_gap_evidence_category_channel ON gap_evidence(category_id, source_channel);

CREATE VIEW IF NOT EXISTS category_intelligence_summary AS
SELECT
    c.slug,
    c.owner,
    c.run_name AS category,
    COUNT(DISTINCT p.product_id) AS shopify_product_count,
    COUNT(DISTINCT s.segment_id) AS stackline_segment_count,
    COUNT(DISTINCT t.top_product_id) AS stackline_top_product_count,
    COUNT(DISTINCT g.evidence_id) AS gap_evidence_count
FROM categories c
LEFT JOIN shopify_category_products p ON p.category_id = c.category_id
LEFT JOIN stackline_segments s ON s.category_id = c.category_id
LEFT JOIN stackline_top_products t ON t.category_id = c.category_id
LEFT JOIN gap_evidence g ON g.category_id = c.category_id
GROUP BY c.category_id;
