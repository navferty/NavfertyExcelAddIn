{
  "title": "Automating routine Excel tasks without VBA",
  "description": "Tables, Power Query, Flash Fill, dynamic arrays and other built-in tools that replace most simple macros.",
  "image": "https://navferty.net/images/NavfertyToolsRibbonEn.png"
}

[← All articles](/en/articles/)

# Automating routine Excel tasks without VBA

Tables, Power Query, Flash Fill, dynamic arrays and other built-in tools that replace most simple macros.

Macros are powerful, but they need maintenance, they get blocked by security settings, and they do not run in Excel for the web. A lot of routine work can be done with built-in features that anyone can inspect and change.

## Tables

A table (<kbd>Ctrl</kbd>+<kbd>T</kbd>) extends formulas, formatting, filters and chart ranges automatically when you add rows. This alone removes many “update the range” macros.

## Power Query

**Data → Get Data** is the right tool for repeated imports and clean-up: combine all files from a folder, remove columns, change types, split, unpivot, merge two tables on a key. Every step is recorded. Next month you drop in the new file and click **Refresh All**.

Typical cases: a monthly export from an accounting system with the same layout; a folder of branch reports to stack into one table; a report that needs unpivoting from a wide to a long format.

## Dynamic array functions

In Microsoft 365 and Excel 2021 and later, one formula can return a whole range:

- `=UNIQUE(A2:A1000)` – a list of distinct values.
- `=SORT(A2:C1000, 2, -1)` – sorted copy of the data.
- `=FILTER(A2:C1000, C2:C1000>1000)` – only the rows that match.
- `=XLOOKUP(E2, A:A, C:C, "not found")` – a lookup with a default value.
- `=TEXTSPLIT`, `=TEXTJOIN`, `=LET` to split, join and name intermediate results.

## Flash Fill

Type the result for the first row (for example the initials extracted from a full name), then press <kbd>Ctrl</kbd>+<kbd>E</kbd>. Excel infers the pattern. It is not a formula, so it does not update when the source changes.

## Pivot tables and slicers

A pivot table replaces most hand-written summary formulas, and refreshes with a click. Add slicers for interactive filters.

## Data validation and conditional formatting

Dropdown lists and rules that highlight problems (an empty required field, a date in the past, a duplicate) prevent errors at entry time, so there is less to clean afterwards.

## Office Scripts and Power Automate

For Excel on the web, Office Scripts (TypeScript) can automate edits, and Power Automate can run them on a schedule or when a file arrives. This is a closer replacement for VBA in a cloud workflow.

## Add-ins for repeated chores

Small operations that appear in many files, such as trimming text, converting text to numbers, highlighting duplicates or exporting to another format, fit well in an add-in with a ribbon button. See [free add-ins](/en/articles/best-free-excel-add-ins/).

## When VBA still makes sense

When you need custom dialogs, interaction with other desktop applications, or logic that does not fit a formula or a query, VBA is still the straightforward way. The VBA articles linked from the [article index](/en/articles/) may help there.
