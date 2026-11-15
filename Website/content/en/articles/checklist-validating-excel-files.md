{
  "title": "Checklist for validating Excel files before sending",
  "description": "A short list of checks to run on a workbook before it leaves your computer.",
  "image": "https://navferty.net/images/NavfertyToolsRibbonEn.png"
}

[← All articles](/en/articles/)

# Checklist for validating Excel files before sending

A short list of checks to run on a workbook before it leaves your computer.

Most embarrassing workbook mistakes are discovered by the recipient. These checks take about ten minutes and cover the common ones.

## Content

- No error values anywhere ([how to find them](/en/articles/find-all-errors-in-excel/)).
- Totals agree with an independent number: the source system, a pivot table or a quick estimate.
- Calculation mode is Automatic and you pressed <kbd>F9</kbd> once.
- No numbers stored as text, no stray spaces in keys ([text numbers](/en/articles/text-numbers-to-numeric-values/), [spaces](/en/articles/remove-extra-spaces-in-excel/)).
- Duplicates in key columns have been reviewed.
- Dates are right: the year, the day/month order, and the period named in the title.
- Units and currency are written near the numbers.
- Spelling checked (<kbd>F7</kbd>), including sheet names and chart titles.

## Hidden things

- Hidden sheets, rows and columns: do you want the recipient to see them? Right-click a sheet tab to see if **Unhide** is active.
- **File → Info → Check for Issues → Inspect Document** finds comments, hidden content, document properties and personal information.
- Comments and notes that were meant for you only.
- Tracked changes and old versions in the file history.
- Author name and path in the properties.

## Links and data sources

- **Data → Edit Links**: a workbook linked to files on your drive shows a security warning on the recipient's side and displays old values. Break the links or paste values.
- Pivot tables are refreshed.
- If you send a report, think about whether formulas that reference other files should stay as formulas.

## Confidential data

- Columns the recipient does not need (salaries, personal data, margins) are deleted, not just hidden.
- If you sent the file with sheet protection, remember it is not security. Use encryption with a password for sensitive files, and send the password by another channel.
- Pivot caches can contain the whole source data even when the pivot shows only a summary.

## Presentation

- Print preview: page orientation, scale, print area, repeated header rows.
- The active sheet, selected cell and scroll position are on the first cell of the first sheet (<kbd>Ctrl</kbd>+<kbd>Home</kbd> on each sheet before saving).
- Column widths are enough to show numbers without `####`.
- Freeze panes on large tables.

## File

- File name tells what it is and when: `Sales_2024-03_v2.xlsx`.
- File size is reasonable (see [speeding up workbooks](/en/articles/speed-up-large-excel-workbooks/)).
- If the recipient may use a different system, think about whether a PDF or CSV is a better format.
- If the file will be imported by a program, run it through the checks in [preparing files for import](/en/articles/prepare-excel-files-for-import/).
