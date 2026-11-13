{
  "title": "Export to SQLite",
  "description": "Export worksheet data to SQLite tables with automatic column type detection.",
  "image": "https://navferty.net/images/ExportToSqlite1.png"
}

[Home](/en/) › [Commands](/en/#commands) › Export to SQLite

# <img alt="" src="/images/icons/validation.png" class="icons" width="64" height="64" />Export to SQLite

Export workbook worksheets to a SQLite database file, with automatic column type detection. Sheets without exportable data are skipped.

- Use first row as column headers
- Skip initial rows if needed
- Automatic data type detection

<div class="shots">

<img src="/images/ExportToSqlite1.png" loading="lazy" width="546" height="680" alt="Export dialog" /> <img src="/images/ExportToSqlite2.png" loading="lazy" width="1412" height="900" alt="SQLite database result" />

</div>

<div class="info-box">

**Related guide:** [Exporting Excel data to SQLite](/en/articles/export-excel-data-to-sqlite/)

</div>

## How to use

1. Open the workbook and choose **Export to SQLite**.
2. Choose whether to use the first row after the skipped rows as column names, and specify how many initial rows to skip.
3. Choose a database file to save.

For example, if a sheet's used range starts with a report title followed by headings, skip one row and enable column names. Each exported worksheet's used range becomes a table, with column types inferred from its data. Names are sanitized for SQLite and empty header cells receive generated names. Empty data cells become SQL `NULL`.

**Important:** An existing database at the chosen path is deleted and replaced. Choose a new filename or back it up first. Formatting and Excel formulas are not exported as workbook features; inspect the resulting data and types before relying on them.

The add-in is free and open source. [Installation instructions](/en/install/).
