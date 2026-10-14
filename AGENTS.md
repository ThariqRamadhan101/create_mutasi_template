# Agent Notes

## What this repo does
Python + openpyxl scripts that read `Data.xlsx` and generate four styled Excel outputs from `Template 1.xlsx` ... `Template 4.xlsx`. Each generator does `shutil.copy2(template, output)` then mutates the copy, so templates (which hold all styles) are never modified and the outputs are safe to delete/regenerate.

## Commands
```bash
# Regenerate all outputs (must succeed before validating)
python generate_from_data.py

# Validate generated outputs against Data.xlsx; prints per-template error counts
python validate_outputs.py
```
There is no linter, test framework, or package manifest — these two scripts are the whole toolchain.

## Critical gotchas
- **Close all `Output Template *.xlsx` files in Excel before running `generate_from_data.py`**. openpyxl cannot overwrite an open file and will raise `PermissionError: [WinError 32]`.
- `PENEMPATAN_ORDER` (33 entries) is hard-coded in `generate_from_data.py`. If a new Unit Kerja Penempatan appears in `Data.xlsx`, add it there or it sorts to the end.
- `validate_outputs.py` imports helpers from `generate_from_data.py`; keep them in sync when changing grouping/column logic.
- Style is captured from a fixed sample row and re-applied to inserted rows via `capture_row_style` / `apply_row_style`. When row layout changes, update the sampled row index in the relevant builder.
- Template 2 must `unmerge_cells` the leftover merged ranges from the template before `delete_rows`, or stale merges corrupt the rebuilt sheet.
- Print settings (`print_title_rows` + page-number header) are NOT copied by `copy_worksheet` — `apply_print_settings(ws, template_key)` must run at the end of every per-sheet builder. Print title rows per template: T1/T2 = 5:6, T3 = 4:4, T4 = 4:5.
- Page number lives in the **header** (top), not the footer: format `"- &P -"` with `firstPageNumber = 2` (first printed page shows "- 2 -"). Font matches each template's body font via `HEADER_FONT`: T1 = Bookman Old Style 10, T2/T3/T4 = Arial 10.
- Template 2: the sample TOTAL row lives at template row 8 (merged A:E, values in F & G). `build_template2_sheet` must capture its style before `delete_rows` and rebuild a TOTAL row under the data (F = G = grand total of all people). `validate_outputs.py` expects this row in its `expected` list.

## Input / output files
- `Data.xlsx` — source data. Columns A–J: Nama, NIP, Pangkat, Golongan, TMT, Jabatan, Unit Organisasi, Unit Kerja, Propinsi, Balai Penugasan. Row 1 is a header; data starts at row 2.
- `Template 1.xlsx` ... `Template 4.xlsx` — read-only style/structure sources (headers, title cells, column widths).
- `Output Template 1.xlsx` ... `Output Template 4.xlsx` — generated results.

## Sheet layout (all templates)
Templates 1, 2, and 4 emit an `all` sheet plus one sheet per Unit Organisasi Asal. Sheet names are produced by `sanitize_sheet_name` (31-char Excel limit; `all` is reserved and collisions get an `_N` suffix). Template 3 is a single sheet.

## Template 2 layout
Columns A–G: No, Penempatan, Unit Organisasi Asal, UPT Asal, Jabatan, Jumlah (per jabatan), Jumlah (group total). Merges are hierarchical:
- A & B merge across all rows with the same Penempatan (`No` runs 1–33, one per penempatan).
- C merges across same Penempatan + Unit Organisasi Asal.
- D & G merge within a group (Penempatan + Unit Org + UPT).
- E (Jabatan) and F (Jumlah) are never merged.

## Dependencies
Only `openpyxl` is used:
```bash
pip install openpyxl
```

## Validation notes
`validate_outputs.py` prints `Template N errors: 0` for each and ends with `Semua output valid!` when clean. Merged cells store their value only in the top-left cell, so the validator uses `get_cell_value_with_merge` to read the displayed value of every row — use it for any new per-row checks, otherwise merged continuation rows read as `None`.
