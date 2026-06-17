-- Ceiling Panels Stackline/Amazon marketplace validation for Product Demand Ideation.
-- This feeds both Amazon Recommendations and validation for competitor inventory-led recommendations.

select
  'tb_scrapping_bsr_gold' as source_table,
  asin,
  title,
  brand,
  category,
  subcategory,
  bsr,
  rating,
  ratings_total,
  price,
  wattage,
  lumens,
  cct,
  date
from public.tb_scrapping_bsr_gold
where
  lower(coalesce(title, '')) like '%panel%'
  or lower(coalesce(title, '')) like '%troffer%'

union all

select
  'competitorpricinganalysis' as source_table,
  asin,
  title,
  brand,
  null as category,
  null as subcategory,
  null as bsr,
  rating,
  reviews as ratings_total,
  price,
  null as wattage,
  null as lumens,
  null as cct,
  date
from public.competitorpricinganalysis
where
  lower(coalesce(title, '')) like '%panel%'
  or lower(coalesce(title, '')) like '%troffer%';


