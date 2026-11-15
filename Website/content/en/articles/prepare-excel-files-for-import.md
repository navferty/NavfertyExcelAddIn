{
  "title": "Preparing Excel files for import into 1C, Power BI and SQL",
  "description": "What import tools expect from a spreadsheet and how to get the file into that shape.",
  "image": "https://navferty.net/images/NavfertyToolsRibbonEn.png"
}

[← All articles](/en/articles/)

# Preparing Excel files for import into 1C, Power BI and SQL

What import tools expect from a spreadsheet and how to get the file into that shape.

Import tools are strict where people are forgiving. A human reads a title, a merged header and a total row and understands the table. A loader sees a mess. Most import failures come down to the same few things.

## Shape of the table

- One header row, directly above the data. No title block above it, or know how many rows to skip.
- No merged cells. Fill the repeated values instead.
- No subtotal or grand total rows inside the data.
- No blank rows or columns in the middle.
- One table per sheet.

## Column names

- Unique, non-empty, short. Duplicated names get suffixes or cause an error.
- For SQL, avoid spaces, punctuation and a leading digit. `order_date` is easier to work with than `Order date (UTC)`.
- Do not rename columns between exports. Many loaders map by name.

## Types

- Every column has one type. A single `-` or `n/a` in a numeric column turns the entire column to text in Power Query or in an automatic type detection.
- Numbers are real numbers, not text with spaces. Dates are real dates. See [text numbers](/en/articles/text-numbers-to-numeric-values/) and [extra spaces](/en/articles/remove-extra-spaces-in-excel/).
- Identifiers with leading zeros (codes, postal codes, phone numbers) are text, and stay text.
- Keep the decimal separator consistent with the locale of the target system, or use an unambiguous format.
- Use ISO dates (`2024-03-12`) in anything that goes to a database. Day/month order is the most common source of silent errors.

## Values

- Formulas are fine in Excel but some importers read the cached value, and a few read the formula text. If in doubt, paste the values into a clean sheet.
- Remove leading and trailing spaces and non-breaking spaces from keys.
- Check for error values (`#N/A`, `#REF!`) before the import: a loader may stop or load them as text. See [finding errors](/en/articles/find-all-errors-in-excel/).
- For XML-based imports, make sure text has no characters that are illegal in XML.

## Tool-specific notes

**Power BI / Power Query.** Prefer an Excel table (<kbd>Ctrl</kbd>+<kbd>T</kbd>) or a named range as the source: it keeps working when rows are added. Rename columns in the source rather than in the query when you can.

**SQL.** CSV with UTF-8 is the least troublesome format (see [CSV and encoding](/en/articles/csv-files-without-encoding-issues/)). For SQLite, an [export straight from the workbook](/en/articles/export-excel-data-to-sqlite/) is an option.

**1C.** Typical import processing expects a fixed column order and a single header row. Check the required template for the document type, and verify tax IDs, codes and dates before the load. A wrong TIN or a text date is rejected row by row, which is slow to fix after the fact.

## A dry run

Import ten rows first. It is much faster to diagnose a problem on ten rows than on fifty thousand.
