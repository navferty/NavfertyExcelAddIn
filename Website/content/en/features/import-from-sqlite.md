{
  "title": "Import from SQLite",
  "description": "Import SQLite tables, views or a SELECT query into a new Excel workbook."
}

[Home](/en/) › [Commands](/en/#commands) › Import from SQLite

# Import from SQLite

Import tables and views from a SQLite database into a new workbook. Each selected table, view or query result gets its own worksheet, with column names in the first row.

## How to use

1. On the **Navferty's Tools** ribbon, choose **Import from SQLite**.
2. Select a `.db`, `.sqlite` or `.sqlite3` database file. The file picker also offers an all-files filter.
3. Check one or more tables or views. Record counts load in the background and may be unavailable for an unreadable object. Use **Select all** to check every item at once, then uncheck any you do not need. If a count exceeds Excel's row limit, the dialog warns you; leave that item unchecked or use a query to limit the results.
4. Optionally enter a single `SELECT` query without a semicolon. You can import the query alone or together with selected tables and views.
5. Click **Import**. The results appear in a new workbook.

The database is opened read-only. SQL `NULL` becomes an empty cell, BLOB values are encoded as base64 text, and integers exceeding 15 digits are imported as text to preserve their precision. Text is formatted as text to prevent formulas from being executed.

Results exceeding Excel's limits (1,048,575 data rows plus a header, 16,384 columns, or 32,767 characters per cell) cannot be imported. Invalid or locked files and other import errors are reported without leaving a partial workbook.

To create a database from an existing workbook, see [Export to SQLite](/en/features/export-to-sqlite/).

The add-in is free and open source. [Installation instructions](/en/install/).
