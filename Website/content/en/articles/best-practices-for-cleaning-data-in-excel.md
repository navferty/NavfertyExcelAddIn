{
  "title": "Best practices for cleaning data in Excel",
  "description": "A repeatable order of steps for turning a messy export into a table you can rely on.",
  "image": "https://navferty.net/images/NavfertyToolsRibbonEn.png"
}

[← All articles](/en/articles/)

# Best practices for cleaning data in Excel

A repeatable order of steps for turning a messy export into a table you can rely on.

Data cleaning goes faster and causes fewer mistakes when you do it in the same order every time and keep the original untouched.

## 1. Keep the original

Copy the raw data to a sheet called `raw` and never edit it. Work on a copy, or better, use Power Query so the steps are recorded and can be replayed on next month's file.

## 2. Look at the data before changing it

- Turn on filters and look at the unique values of each column. Misspellings and variants of the same category show up in seconds.
- Sort a column ascending and descending: blanks, text among numbers and unexpected extremes end up at the edges.
- Check the row count and the totals, so that you can compare after cleaning.

## 3. Fix the structure

- Remove title rows, notes and total rows inside the data.
- Unmerge cells and fill in the blanks. **Go To Special → Blanks**, type `=` and the cell above, then <kbd>Ctrl</kbd>+<kbd>Enter</kbd> fills down.
- One header row, unique column names.

### Format the result as a table

Once the structure is clean, select any cell in the data and press <kbd>Ctrl</kbd>+<kbd>T</kbd> (Home → Format as Table). Tick "My table has headers". It costs nothing and helps later:

- The range grows by itself when you add rows, so formulas, charts and pivot tables based on it do not need updating.
- Filters and banded rows are added, and the header stays visible while scrolling.
- Formulas can use column names: `=SUM(Table1[Amount])` instead of `=SUM(C2:C500)`. A formula typed in one row fills the whole column.
- Give the table a meaningful name on the Table Design tab (for example `Orders`).

Tables do not allow merged cells, and a table cannot contain subtotal rows inside the data, which fits well with the rules above.

## 4. Fix the text

- Remove extra spaces and non-printable characters ([details](/en/articles/remove-extra-spaces-in-excel/)).
- Normalise case with `UPPER`, `LOWER`, `PROPER`. `PROPER` mishandles names like `McDonald` and abbreviations, so review the result.
- Look for lookalike letters. Latin `A` and Cyrillic `А` are different characters, and a code that mixes them will not match anything. This is common in Russian-language data with codes and licence plates.
- Split combined fields with **Text to Columns** or **Flash Fill** (<kbd>Ctrl</kbd>+<kbd>E</kbd>). Check the Flash Fill results, because it guesses a pattern.

## 5. Fix the types

- Convert text numbers and text dates to real ones ([details](/en/articles/text-numbers-to-numeric-values/)).
- Keep identifiers as text if they have leading zeros.
- Use data validation lists on columns where people type values by hand.

## 6. Handle duplicates and missing values

- Look at duplicates before removing them ([details](/en/articles/highlight-duplicate-values-in-excel/)). Two rows with the same key may be different records.
- Decide per column what a blank means: zero, unknown or not applicable. Do not fill in with zeros blindly, it changes averages.

## 7. Validate the result

- Compare row count and totals with the original.
- Check for formula errors ([details](/en/articles/find-all-errors-in-excel/)).
- Test a few rows by hand against the source system.

## Write down what you did

A small “Notes” sheet with the steps will save you the next time the same file arrives, and lets a colleague trust the result.
