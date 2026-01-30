MRR Automation for Warehouse (Jan 21/26)

Had a teams call with Keith January 21, 2026 11:31 AM. He asked if we could use the Flat BOM data from an IFP to populate the MRR (Material Receiving Report) excel template. It would need to read a PO (pdf) and extract PNs and compare them to the Flat BOM then extract the data from the flat bom and process a bit and dump it into a new sheet at the end of the MRR template and then begin migrating data to MRR tabs.

Inputs Files

- PO from NetSuite (pdf)
- Flat BOM from IFP (excel)
- MRR Template (excel)

Action Steps (command 1)

- Load each input file
- Read PNs from PO
- Compare PNs from PO to Flat BOM
- Extract Flat BOM rows corresponding to PNs from PO
- Extract material from Description ito new column + leave description without material (removing `,` )
- deposit Qty, Description, Material, PN into single table in Template in new sheet at the end
- (stop automation)
- user: would delete rows (or mark in far right column an “x” to signify it should not be included in data migration to MRR template sheets)

(nice to have) (command - continue)

- new function/command to continue next step in processing — (click button)
- evaluate total rows without “x”, if > 20, duplicate sheet 1 and rename to MRR 2 and so forth, until enoguh sheets of 20 rows each are greater than total rows in raw data
- copy first 20 rows into MRR sheet 1 — ensure fields from raw table are matched to MRR columns
- copy next 20 rows …
- until all rows are copied
- produce MRR as excel to save

(other nice to have)

- after initial upload and command 1 is complete
- user: column to show how many times to split that 1 PN row from raw data into X rows

(edge cases)

- ?
- if PO contains PNs not in BOM - flag items at bottom of raw data table and notify user with prompt