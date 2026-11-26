# Form-Builder

Windows app. Overlay customer data on fixed PDF forms. Original PDF never changed. Output flattened PDF.

## Use
1. Open PDF.
2. Load Data (Excel/CSV). One row = one customer.
3. Click on form where text goes. Type column name.
   - `=Literal text` fixed text
   - `check:Column` draws X when value is yes/true/1/x
4. Save Tpl (JSON, tied to PDF hash).
5. Pick row, Generate. Or "ALL ROWS".
6. Print at 100% / Actual size.

## Build
Push to main or tag `v*`. GitHub Actions builds `FormBuilder.exe` (Actions artifact; tag makes Release).
