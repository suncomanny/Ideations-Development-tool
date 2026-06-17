-- Ceiling Panels Sunco catalog coverage for Product Demand Ideation.

select
  coalesce(a.master_sku, s.master_sku) as master_sku,
  a.asin,
  coalesce(a.title, a.name, s.product_core, a.family_name) as title,
  a.family_name,
  s.category,
  s.product_core,
  coalesce(a.pack_size, s.pack_size) as pack_size,
  s.status as shopify_status,
  s.available as shopify_available,
  s.now_price as shopify_price,
  a.now_price as amazon_price,
  a.onhand as amazon_onhand
from public.sunco_pilot_amazon a
full outer join public.sunco_pilot_shopify s
  on a.master_sku = s.master_sku
where
  lower(coalesce(a.title, '')) like '%panel%'
  or lower(coalesce(a.title, '')) like '%troffer%'
  or lower(coalesce(a.name, '')) like '%panel%'
  or lower(coalesce(a.name, '')) like '%troffer%'
  or lower(coalesce(a.family_name, '')) like '%panel%'
  or lower(coalesce(a.family_name, '')) like '%troffer%'
  or lower(coalesce(s.category, '')) like '%panel%'
  or lower(coalesce(s.category, '')) like '%troffer%'
  or lower(coalesce(s.product_core, '')) like '%panel%'
  or lower(coalesce(s.product_core, '')) like '%troffer%';


