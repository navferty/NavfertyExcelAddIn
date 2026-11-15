{
  "title": "Beginner's guide to Excel formulas",
  "description": "The basics of writing formulas, with the functions most people use every week.",
  "image": "https://navferty.net/images/NavfertyToolsRibbonEn.png"
}

[← All articles](/en/articles/)

# Beginner's guide to Excel formulas

The basics of writing formulas, with the functions most people use every week.

A formula starts with `=` and can contain numbers, cell references, operators and functions. Excel recalculates it every time the inputs change.

## The basics

    =A1+B1        add
    =A1*B1        multiply
    =A1/B1        divide
    =A1^2         power
    =(A1+B1)*C1   parentheses change the order

Excel follows the usual order: parentheses, powers, multiplication and division, then addition and subtraction.

## References

- `A1` is relative: when you copy the formula one row down, it becomes `A2`.
- `$A$1` is absolute and does not change when copied. Press <kbd>F4</kbd> to add the dollar signs.
- `$A1` and `A$1` lock only the column or only the row.
- `Sheet2!A1` refers to another sheet. If the name has spaces, quote it: `'Sales 2024'!A1`.
- A range is written `A1:A10`.

## Ten functions worth knowing

- `SUM(A1:A10)`, `AVERAGE`, `MIN`, `MAX`, `COUNT` (numbers), `COUNTA` (non-empty).
- `IF(A1>100, "high", "low")` – a choice. Combine conditions with `AND`, `OR`.
- `SUMIF(range, criteria, sum_range)` and `SUMIFS` – sum with one or several conditions. `COUNTIF` and `COUNTIFS` do the same for counting.
- `XLOOKUP(key, lookup_range, result_range)` or `VLOOKUP(key, table, column, FALSE)` – find a value by key ([more on VLOOKUP](/en/articles/vlookup-in-excel/)).
- `INDEX` and `MATCH` – the flexible lookup combination that works in all versions and in any direction.
- `IFERROR(formula, "")` – a fallback for when a formula returns an error. Use it deliberately (see [errors](/en/articles/find-all-errors-in-excel/)).
- `LEFT`, `RIGHT`, `MID`, `LEN`, `FIND` – work with parts of the text. `TRIM` and `SUBSTITUTE` clean it.
- `TEXT(A1, "yyyy-mm")` – format a number or date as text. The format codes depend on the locale of Excel.
- `&` or `TEXTJOIN` – join text: `=A1&" "&B1`.
- `TODAY()`, `DATE(y,m,d)`, `EOMONTH`, `DATEDIF` – dates.

For totals and counts by category you may not need formulas at all: see [pivot tables basics](/en/articles/pivot-tables-basics/). [Keyboard shortcuts](/en/articles/excel-keyboard-shortcuts/) speed up everything else.

## Useful habits

- Put inputs (rates, thresholds) in their own cells, not inside the formulas.
- Convert a range to a table and use column names in formulas.
- Click a formula and press <kbd>F2</kbd>: Excel outlines the referenced cells in colour. **Formulas → Evaluate Formula** walks through the calculation step by step.
- Check results with an example you can calculate by hand.
- Type the first letters of a function and press <kbd>Tab</kbd> to complete it, then <kbd>Ctrl</kbd>+<kbd>Shift</kbd>+<kbd>A</kbd> to see the argument names.

## Locale notes

In some language versions the list separator is a semicolon rather than a comma (`=IF(A1>100;"high";"low")`), and function names are localised (`SUM` is `СУММ` in Russian Excel). If a formula copied from a web page does not work, check this first.

## Common first errors

Parentheses that do not match, text without quotes, a number stored as text, a circular reference (a formula that refers to its own cell). [Common mistakes](/en/articles/common-excel-mistakes/) lists more.
