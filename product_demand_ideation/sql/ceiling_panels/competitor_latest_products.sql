-- Ceiling Panels competitor PDP/spec metadata for Product Demand Ideation.
-- Avoid select * because the current Redshift path can fail on boolean parquet fields.

select
  name,
  price,
  price_high,
  currency,
  sku,
  mpn,
  gtin,
  model,
  brand,
  category,
  product_type,
  description,
  specifications,
  wattage,
  lumens,
  cct,
  cri,
  ip_rating,
  voltage,
  life_hours,
  dimmable,
  rating,
  review_count,
  reviews_sample,
  availability,
  stock_qty,
  stock_status,
  stock_cached_at,
  warehouse,
  restock_date,
  image,
  additional_images,
  url,
  scraped_at,
  scraped_timestamp,
  first_seen_date,
  tracking_days,
  appearances
from public.vw_competitors_scraping_latest
where
  lower(coalesce(name, '')) like '%panel%'
  or lower(coalesce(name, '')) like '%troffer%'
  or lower(coalesce(category, '')) like '%panel%'
  or lower(coalesce(category, '')) like '%troffer%'
  or lower(coalesce(product_type, '')) like '%panel%'
  or lower(coalesce(product_type, '')) like '%troffer%';
