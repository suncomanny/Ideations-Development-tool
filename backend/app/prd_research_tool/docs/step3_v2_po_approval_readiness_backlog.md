# Step 3 V2 PO Approval Readiness Backlog

Purpose: capture the Step 3 improvements that would better prepare Product Managers for Sam/Simon PO approval conversations without turning the current Step 1-2-3 workflow into a final PO Ready process.

Current release stance:

- Step 3 remains a market and PRD-research report, not a PO Ready form.
- Step 3 should not invent vendor-final facts such as final MOQ, lead time, monthly capacity, patent coverage, incoterm, or final landed cost.
- Vendor-final fields should remain later-gate inputs after RFQ, sample/EVT review, vendor PO Ready form, and internal PO approval prep.

V2 would-like-to-have section:

Add a "PO Approval Readiness" section near the end of each Step 3 ideation sheet with three buckets:

1. Covered by Step 3
   - Closest current Sunco/NSL product comparison.
   - Clear differentiation from current Sunco/NSL products.
   - Why this product is worth adding now.
   - Verified competitor examples and whether they are apples-to-apples.
   - Competitor demand evidence from Stackline, ecommerce snapshots, verified listings, ratings/reviews, or other approved sources.
   - Related Sunco family sales trend where reference-family data exists.
   - Stackline segment quality check and unrelated-product warning.
   - Same, similar, new product, or new category classification.
   - Initial recommended channel posture: Amazon, Shopify, both, or staged test.
   - Initial pack-size rationale using competitor norms and current Sunco pack performance where available.

2. Needs PM/Data Team follow-up
   - Explanation for weak or declining related Sunco sales before using those SKUs as support.
   - Stronger product-traffic signal when ecommerce/paid-search/ad-library data becomes available.
   - Pack-size recommendation from Felipe/data team.
   - First-order quantity logic tied to test strategy, inventory risk, MOQ expectations, or demand.
   - Confidence/risk score inputs that can be prefilled before RiskCalculatorV3 is finalized.
   - Product and packaging dimensions when supplier or sample data is not yet available.

3. Needs Vendor/RFQ/EVT/PO form
   - Final MOQ.
   - Final vendor cost, landed cost, total investment, and gross margin.
   - Incoterm recommendation and matching landed-cost assumptions.
   - Product dimensions and packaging dimensions if not already confirmed.
   - Compatibility notes, accessories, sensors, clips, installation constraints, and included parts that require supplier confirmation.
   - Required certifications and who pays for them.
   - Patent/legal risk status and whether vendor covers legal fees.
   - Vendor lead time, reorder lead time, and monthly capacity.
   - Final RiskCalculatorV3 confidence score.

Desired output shape:

- Keep the existing Step 3 sections intact.
- Add a concise readiness summary, not a new workbook.
- Use direct labels: `Supported`, `Partial`, `Missing later-gate input`, and `Blocker for PO Ready`.
- End each ideation sheet with the likely approval question:
  "Should we move forward, and what quantity/pack/channel should be investigated next?"

Implementation note:

This is a Version 2 backlog item. Do not implement before the current Step 2-3 smoke tests and team deployment package are stable.
