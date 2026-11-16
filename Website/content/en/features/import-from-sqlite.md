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

The database is opened read-only.

## How values are converted

SQLite allows values of different storage classes in the same column. For each value in a table, view or query result, the SQLite provider first attempts its normal conversion, which can use the declared column type. If it throws a format, invalid-cast or overflow exception, the importer reads the value using its SQLite storage class instead. Other errors still stop the import.

The resulting values are written as follows:

| Value returned by the provider or fallback | Excel result |
| --- | --- |
| `NULL` | Empty cell |
| Integer (`Int64`) | Number if it has at most 15 digits; otherwise text, to avoid loss of precision |
| Real (`Double`) | Number |
| Text | Text |
| BLOB | Base64-encoded text (for example, `X'00FF'` becomes `AP8=`) |
| .NET `DateTime` | Round-trip date/time text, preserving fractional seconds and any kind/offset indicator; it is not an Excel date cell |

For example, in a `DATETIME` column, `2026-01-02 03:04:05.1234567` can be converted by the provider to a `DateTime` and imported as `2026-01-02T03:04:05.1234567`. If another row contains `not a date`, that conversion fails and the text is imported unchanged. The importer continues with subsequent rows.

The fallback only runs when the provider throws one of the conversion exceptions above. A successful provider conversion may already have changed the value or its representation, so exact preservation of every original SQLite value is **not** guaranteed. SQLite date/time functions such as `datetime()` return text; they do not themselves create a .NET `DateTime`.

The destination ranges are formatted as text before writing to prevent imported text from being executed as formulas. Numeric values are still written as numbers.

Results exceeding Excel's limits (1,048,575 data rows plus a header, 16,384 columns, or 32,767 characters per cell) cannot be imported. Invalid or locked files and other import errors are reported without leaving a partial workbook.

To create a database from an existing workbook, see [Export to SQLite](/en/features/export-to-sqlite/).

The add-in is free and open source. [Installation instructions](/en/install/).
