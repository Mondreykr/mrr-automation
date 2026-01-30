# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

MRR (Material Receiving Report) Automation tool for Enercorp warehouse operations. The system extracts data from Purchase Orders (PDFs) and Flat BOMs (Excel), cross-references part numbers, and populates MRR Excel templates.

**End Users**: Keith and warehouse staff (non-technical, need simple GUI)
**Usage Frequency**: Daily

## Technology Stack

**Decision: HTML/JavaScript single-file application** (like the existing BOM tool)

- **PDF parsing**: pdf.js (CDN: `https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js`)
- **Excel reading (BOM)**: SheetJS (xlsx library) - for data extraction from Flat BOM
- **Excel template & export**: ExcelJS - **preserves all formatting** (colors, borders, fonts, merged cells, images)
- **GUI**: HTML/CSS in browser
- **Distribution**: Single HTML file (~500KB-1MB), shareable via email/OneDrive

**Why ExcelJS for templates:** SheetJS community edition strips formatting when reading/writing. ExcelJS preserves styles automatically by manipulating the XML structure directly rather than deserializing to an object model.

Rationale: Small file size, easy distribution, cross-platform, consistent with existing tooling.

## Input Files

### PO PDF (from NetSuite)
- Standardized format, 2+ pages
- Page 1: Exhibit 9.2 boilerplate (skip)
- Page 2+: Line items table with Part Numbers
- **PN format**: `1029200 | Shell, Rolled Plate...` - extract 7-digit number before pipe
- **Extraction regex**: `/(\d{7})\s*\|/g` or `/\b(10\d{5})\b/g`

### Flat BOM Excel (from IFP tool)
- Filename pattern: `{JobNumber}-IFP REV{X}-Flat BOM-{YYYYMMDD}.xlsx`
- **Columns**: Qty, Part Number, Component Type, Description, Length (Decimal), Length (Fractional), UofM, Purchase Description, State, Revision, NS Item Type
- Part Numbers are 7-digit integers (e.g., `1029834`)
- ~200 rows typical

### MRR Template Excel
- **Sheets**: MRR, MRR 2, MRR 3, Data sheet
- **MRR sheet layout**: 16 data rows per sheet (rows 6-21), header in row 5
- **MRR sheet columns**: Qty, Circ., Diam., Thickness, Heat #, Product Description, Material Spec., Part #, Color Code
- **Column I (Color Code)**: Contains VLOOKUP formulas that auto-calculate from Material Spec.
- **Data sheet**: Contains reference list of known materials with color codes (used for material validation and color lookup)

## Data Mapping

| Flat BOM | MRR Field |
|----------|-----------|
| Qty | Qty |
| Description (material removed) | Product Description |
| Extracted material | Material Spec. |
| Part Number | Part # |
| Lookup from Data sheet | Color Code |
| (blank - filled on receipt) | Circ., Diam., Thickness, Heat # |

## Material Parsing Logic

1. Split description at **last comma**
2. Text after comma = candidate material
3. Validate against known materials list (from MRR Template "Data sheet")
4. If match found: extract as material, assign color code
5. If no match: leave description intact, flag for user review

**Known materials** (from Data sheet): SA516-70N, SA350-LF2 CL.1, CSA G40.21 44W, A333-6, AR400, A193-B8M, A240-316L, etc.

## Workflow

1. User loads 3 files via drag-drop or file picker
2. App extracts PNs from PO PDF
3. App cross-references PNs against Flat BOM
4. App parses material from descriptions
5. App displays review table:
   - Checkbox to exclude rows
   - **Split column**: Enter number to create multiple rows for one item (V2)
   - Editable Qty for split rows (V2)
   - Editable material field (for parsing corrections)
   - Flag for PNs not found in BOM
6. User reviews, makes corrections, adjusts splits
7. App exports populated MRR Excel (cloned from template to preserve formulas)

## V2 Features

### Row Splitting (Heat-Based)
- **Use case**: Material arrives in multiple heats, user needs separate rows for each heat
- **UI**: Inline number input in "Split" column (default 1)
- **Behavior**: Enter 2 = creates 2 rows from this item, user edits Qty per row
- **Visual**: Split rows highlighted with same background color (grouped)
- **Data propagation**: Excluding/material changes propagate to all rows in group

### Multi-Sheet Distribution
- Template has **16 data rows per sheet** (rows 6-21)
- Auto-distributes rows across MRR, MRR 2, MRR 3
- **Dynamic sheet creation**: If > 48 rows, creates MRR 4, MRR 5, etc.

### Template Cloning (ExcelJS)
- V2 uses **ExcelJS** to clone the template workbook (preserves ALL formatting)
- Deep copy via write/read buffer cycle maintains: colors, borders, fonts, merged cells, images
- **Preserves Column I formulas**: VLOOKUP auto-calculates Color Code
- Writes to columns A, F, G, H only (Qty, Product Desc, Material Spec, Part #)

### Processed Data Sheet
- Complete audit trail added as last sheet in export
- Includes ALL rows (excluded and included)
- Columns: Qty, Part #, Product Description, Material Spec., Color Code, Excluded (Yes/No), Split Group

## Edge Cases

- **PN in PO but not in BOM**: Flag at bottom of table, notify user (e.g., Impact Test Coupons)
- **Material parsing uncertain**: Show for user review, allow manual edit
- **Multiple matches for fuzzy material names**: Use Data sheet as canonical reference
- **All rows excluded**: Export only Processed Data sheet, warn user
- **Split count = 0**: Treated as 1 (minimum)
- **Large split (e.g., 50)**: Allowed, but overflow warning shows sheet count

## Development Notes

- Python 3.13 is installed on dev machine
- pandas and openpyxl available for testing/validation scripts
- Test files available in project root (PO PDFs, Flat BOM, MRR Template)
