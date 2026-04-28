# Template Fill Task TODO

## Objective
Fill template with source data: 
- "Current Plan Description" ← plan name from source
- "Monthly Total Premium" ← total premium from "current premium" column

## Steps
- [ ] Step 1: Inspect data sheets to locate 'plan name' and 'current premium' columns (e.g. 'SalesConsultantBroker List').
- [ ] Step 2: Update template.py config: source_sheet to correct sheet (e.g. 'SalesConsultantBroker List'), ensure fields map 'plan name' → 'N', 'current premium' → 'O'.
- [ ] Step 3: Identify template file (likely invoice_pdf_filtered.xlsx as template, source.xlsx as source).
- [x] Step 1: Inspect data sheets (inspect_plans.py done - Census is form-like, data sheets unnamed/headers lower).
- [x] Step 3: Created config.json for 'Employee Details' form_sheet (filtered.xlsx has it, columns 'plan name', 'current premium').
- [x] Step 4: Ran filler - created filled_invoice.xlsx (0 rows filled - no matches, likely because source Census unnamed, no 'full name' col).
- [ ] Step 2: Fix source loading (header=None, map index for plan name/current premium from inspect).
- [ ] Step 5: Re-run filler.
- [x] Complete? Output created, but need data mapping.

Progress:
- [x] Inspections & initial run
- [ ] ...

