Hi y'all,

here there is the documentation of the Excel_to_Json_Csv_Sql.

This is a small desktop app (Tkinter) that extracts tables from Excel workbooks and exports them as JSON, CSV, and/or Oracle DDL (SQL).
Tables are detected by a header pattern you define in the UI.

WHAT IT DOES
 - Open an .xlsx workbook and scan selected sheets.
 - Detect one or more tables in each sheet based on a header row you specify.
 - Parse each table into structured data (table_name, header, rows, …).
 - Export:
    - JSON: one file with all parsed tables.
    - CSV: one file per table.
    - SQL (DDL): one file per table (+ an _ALL_TABLES.sql bundle) with CREATE TABLE, PRIMARY KEY, and COMMENT statements.
 - Assign an Oracle schema per sheet (or use a default) for the SQL.

REQUIREMENTS AND PYTHON VERSION
-> read ".python-version" and "requirements.txt"

HOW DETECTION WORKS
 - The app searches each sheet for a consecutive row whose cells match your Header Pattern (case/spacing-insensitive).
 - The table name is read from the row above the header (merged cells supported). If absent, the sheet name is used.
 - Data rows are read downwards until:
    - an empty row is found (if “Stop on empty row” is enabled), or
    - another header pattern is encountered (start of a new table).

GUI — fIELDS AND ACTIONS
 - Source Excel
    - Path — choose the .xlsx to process.
    - Load Sheets — loads sheet names into the list.

 - Header Pattern (defines table start)
    - Set Frame (format: col_a, col_b, …) — type the exact header labels in order, separated by commas.
      Example on ./data/EXAMPLE_EXCEL.xslx: COLUMN NAME, DATA TYPE, PK, NULL, DEFAULT, DESCRIPTION, COMMENTS
 
 - Output 
    - JSON — writes tables.json containing all parsed tables.
    - CSV — writes one CSV per table into out/<timestamp>/csv/.
    - SQL (DDL) — writes one .sql per table into out/<timestamp>/sql/ and _ALL_TABLES.sql.
    - Stop on empty row — if enabled, the parser stops a table at the first empty row in the header’s column span.

 - Output Directory
    - Choose where to write json/, csv/, and sql/.

 - Sheets & Schemas
    - List of sheets — select the sheets to process (if none selected, all sheets are processed).
    - Schema per sheet — optional text field for each sheet.
      The app builds a map {sheet → schema}; blank entries use the first non-blank schema you typed, or fallback MYSCHEMA.
 
 - Run
    - Launches parsing and exports according to your selections.

RUNNING THE APP

<python app_gui.py> from the terminal on the app folder

1. Click Browse… and choose your Excel.
2. Click Load Sheets.
3. Enter the Header Pattern (comma-separated).
4. Select sheets (optional).
5. Fill schemas (optional; leave blank to use default).
6. Select export formats and output folder.
7. Click Run.

TROUBLESHOOTING
- “No tables found with the given header pattern.”
   - Check the header labels: they must match exactly after trimming each token (case/spacing is normalized).
   - Verify you included all header columns and correct order.
   -If tables have empty spacer rows, try unchecking Stop on empty row.
- Wrong table name
   - The app reads the row above the header (merged cells supported). If blank, it falls back to the sheet name.
- SQL errors in Oracle
   - Make sure DATA TYPE values in Excel are valid Oracle types (e.g., VARCHAR2(50), NUMBER(10,2), DATE, TIMESTAMP(6)).
   - If you need strict date/timestamp defaults, add TO_DATE/TO_TIMESTAMP formatting in ddl_module.py.
- Long names truncated
   - Oracle identifiers are limited to 30 chars. The app truncates after sanitization.

KNOWN LIMITATIONS
- Detection assumes headers are contiguous and in a single row.
- Stops at the first empty row within the header’s span (if enabled).
- Does not infer data types from sample values—relies on Excel DATA TYPE column.