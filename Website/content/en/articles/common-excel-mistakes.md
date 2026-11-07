{
  "title": "Common Excel mistakes and how to fix them",
  "description": "Problems that show up again and again in real workbooks, with a way to avoid each one.",
  "image": "https://navferty.net/images/NavfertyToolsRibbonEn.png"
}

[← All articles](/en/articles/)

# Common Excel mistakes and how to fix them

Problems that show up again and again in real workbooks, with a way to avoid each one.

Most spreadsheet errors are not complicated. They come from a handful of habits that are easy to change.

## Numbers stored as text

Totals that do not add up and lookups that fail often come from text-formatted numbers. See [converting text numbers](/en/articles/text-numbers-to-numeric-values/) for the fixes.

## Hard-coded numbers inside formulas

`=B2*1.2` works until the tax rate changes and you have to find every place where 1.2 was typed. Put the rate in a labelled cell and reference it: `=B2*$F$1`. Name the cell (**Formulas → Define Name**) and the formula reads `=B2*VAT`.

## Relative and absolute references mixed up

If a formula copied down gives odd results, check the dollar signs. `$A$1` never moves, `A$1` keeps the row, `$A1` keeps the column. <kbd>F4</kbd> cycles through them while editing.

## Ranges that do not cover new rows

`=SUM(B2:B100)` silently misses row 101. Convert the data to a table (<kbd>Ctrl</kbd>+<kbd>T</kbd>) and use structured references such as `=SUM(Sales[Amount])`, which grow with the data.

## Merged cells in the data area

Merged cells break sorting, filtering, pivot tables and copy-paste. Use **Center Across Selection** (**Format Cells → Alignment → Horizontal**) for the visual effect. If you inherited a merged table, unmerge it and fill each cell with the original value.

## Dates that are really text

If `01.02.2024` will not sort or calculate, it is text, or it was read with the wrong day/month order. Check with `=ISNUMBER(A1)`. Convert with `DATE`, `DATEVALUE` or Text to Columns with the date format set explicitly.

## Approximate match in lookups

`VLOOKUP(x, range, 2)` without the last argument defaults to approximate match and needs sorted data. It can return a wrong value instead of an error. Always write `FALSE` (or use `XLOOKUP` where available, which defaults to an exact match).

## Hidden rows, columns and sheets

`SUM` includes hidden rows, `SUBTOTAL(109, ...)` does not. Filtered rows are the same story: pick the function that matches what you intend. Before you send a file, check **Review → Inspect Document** or right-click a sheet tab and look at **Unhide**.

## Copying formulas over values and back

Paste Special → Values removes the link to source cells. Sometimes that is intended, but a report where half of the figures are frozen is a trap. Colour inputs differently from formulas, and keep one sheet with the raw data you never touch.

## No second pair of eyes

A few cheap checks catch a surprising amount: totals compared with an independent calculation, a count of rows before and after, and **Formulas → Show Formulas** (<kbd>Ctrl</kbd>+<kbd>\`</kbd>) to look for a cell where a formula was replaced by a number. See also the [pre-send checklist](/en/articles/checklist-validating-excel-files/).
