{
  "title": "How to speed up large Excel workbooks",
  "description": "Why a workbook becomes slow and what to change first.",
  "image": "https://navferty.net/images/NavfertyToolsRibbonEn.png"
}

[← All articles](/en/articles/)

# How to speed up large Excel workbooks

Why a workbook becomes slow and what to change first.

A workbook that takes a minute to open or freezes on every edit has usually one or two clear causes. Find those before trying everything.

## Measure first

Switch to manual calculation (**Formulas → Calculation Options → Manual**). If the file becomes responsive, formulas are the issue. If it is still slow, look at file size, formatting, objects and add-ins. Press <kbd>F9</kbd> to calculate manually. Remember to switch back, or tell whoever uses the file.

## Formulas

- **Whole-column references in heavy functions.** `SUMIFS` on `A:A` is OK in modern Excel, but array formulas and `SUMPRODUCT` over entire columns are slow. Limit ranges or use tables.
- **Thousands of `VLOOKUP`s with exact match over a large range.** Sort the data and use binary search where it makes sense, or add one helper column with `MATCH` and reuse its result with `INDEX`. `XLOOKUP` and a lookup against a proper table are also faster than repeating the same search in many columns.
- **Volatile functions** (`NOW`, `TODAY`, `RAND`, `OFFSET`, `INDIRECT`) recalculate on every change and everything that depends on them too. Replace `OFFSET` and `INDIRECT` with `INDEX` or table references where you can.
- **Formulas that never change.** Old months in a report do not need to be calculated each time. Paste them as values and keep a copy of the formulas somewhere.
- **Chains of dependency across many sheets and external links.** Check **Data → Edit Links** and remove links to files that no longer exist.

## File size

- **Used range that is much bigger than the data.** Press <kbd>Ctrl</kbd>+<kbd>End</kbd>. If it jumps far beyond your data, select the empty rows and columns below and to the right, delete them (not just clear), and save.
- **Formatting applied to entire columns or rows** that were then filled with fancy conditional formatting. Check **Conditional Formatting → Manage Rules** and delete duplicates.
- **Images and shapes.** Compress pictures (**Picture Format → Compress Pictures**) and delete invisible shapes. **Home → Find & Select → Selection Pane** lists all objects, including hidden ones.
- **Pivot caches.** Each pivot table may carry a copy of the data. Use the same cache for pivots built on the same source, and untick **Save source data with file** where acceptable.
- **Format.** Saving as `.xlsb` (binary workbook) is usually smaller and opens faster than `.xlsx`. The trade-off is that some non-Microsoft tools cannot read it.

## Environment

- Use 64-bit Excel if you work with very large data sets.
- Disable add-ins you do not use (**File → Options → Add-ins**) and check whether the file is slow without them.
- Files on network drives open slowly; copy locally to test.

## When Excel is the wrong place

If a workbook holds millions of rows or many joins, move the data to a database and use Excel for reports on top of it. Power Query with the data model, or a SQLite or other database file, handles volumes where formulas do not. See [exporting to SQLite](/en/articles/export-excel-data-to-sqlite/).
