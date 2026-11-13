{
  "title": "Exporting Excel data to SQLite",
  "description": "When a spreadsheet is no longer enough, SQLite gives you real queries in a single file.",
  "image": "https://navferty.net/images/NavfertyToolsRibbonEn.png"
}

[← All articles](/en/articles/)

# Exporting Excel data to SQLite

When a spreadsheet is no longer enough, SQLite gives you real queries in a single file.

Excel handles a few hundred thousand rows, but joins across sheets, grouping with conditions and repeated lookups get slow and fragile. SQLite is a small database stored in one file, with no server. Many tools can read it: DB Browser for SQLite, DBeaver, Python, R, Power BI and most programming languages.

## Why move data there

- You need a `JOIN` across several tables instead of nested `VLOOKUP`s.
- You want to run the same query repeatedly on a new data set.
- The workbook has become slow and large.
- A developer or analyst asks for the data in a format a program can read reliably.

## Prepare the sheets

Whatever tool you use, the result is only as good as the input:

- One header row, one table per sheet, no merged cells and no blank rows or columns in the middle.
- Unique and non-empty header names. Avoid spaces and special characters if you plan to write SQL by hand.
- One type per column. A column with mostly numbers and a few `n/a` strings becomes text.
- Dates as real dates, not as text.

## Other ways to do it

- Save each sheet as CSV and import it with the `.import` command of the `sqlite3` command-line tool, or with the import dialog in DB Browser for SQLite.
- Use Python with pandas: `pd.read_excel(...).to_sql(...)`.
- Connect through ODBC and write the data with VBA. This is flexible and needs the most setup.

## With the add-in

**Export to SQLite** writes the whole workbook to a `.db` file. Each sheet with exportable data becomes a table, the column types are detected automatically, and you can choose whether the first row holds the column names and how many leading rows to skip, which helps when a report has a title block above the table. **Choose a new filename or back up the destination first: an existing database at the selected path is deleted before export begins, including all its existing tables, even if the export subsequently fails.**

## A first query

Open the file in any SQLite client. If the sheet was called `Orders` and has columns `Customer` and `Amount`:

    SELECT Customer, SUM(Amount) AS Total
    FROM Orders
    GROUP BY Customer
    ORDER BY Total DESC;

The add-in replaces spaces and other characters outside letters, digits and underscores with underscores in table and column names. For example, `Order date` becomes `Order_date`. Names starting with a digit also gain a leading underscore. Use the exported names in your queries.

The add-in command for this: [Export to SQLite](/en/features/export-to-sqlite/).
