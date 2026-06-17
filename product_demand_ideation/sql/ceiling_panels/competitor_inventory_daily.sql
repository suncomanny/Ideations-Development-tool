-- Ceiling Panels inventory movement for Product Demand Ideation.

select
  url,
  brand,
  sku,
  name,
  scrape_date,
  stock_qty,
  stock_status,
  availability,
  price,
  currency,
  scraped_at,
  prev_scrape_date,
  days_since_prev,
  prev_stock_qty,
  stock_qty_delta,
  prev_price,
  price_delta
from public.v_competitors_inventory_daily
where
  lower(coalesce(name, '')) like '%panel%'
  or lower(coalesce(name, '')) like '%troffer%';


