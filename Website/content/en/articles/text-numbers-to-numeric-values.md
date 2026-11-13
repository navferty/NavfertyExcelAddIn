{
  "title": "Converting text numbers to numeric values",
  "description": "Why numbers become text in Excel and several ways to convert them back.",
  "image": "https://navferty.net/images/articles/text-numbers-to-numeric-values/failed-sum.jpg"
}

[← All articles](/en/articles/)

# Converting text numbers to numeric values

Why numbers become text in Excel and several ways to convert them back.

Numbers copied from a web page, a PDF, an accounting system or a CSV often arrive as text. They look normal, but `SUM` ignores them, sorting puts `100` before `20`, and `VLOOKUP` refuses to match them with real numbers.

<img src="/images/articles/text-numbers-to-numeric-values/failed-sum.jpg" loading="lazy" width="547" height="282" alt="SUM returns 0 for a column of numbers stored as text" />

## How to recognise text numbers

- They are left-aligned by default while real numbers are right-aligned.
- A small green triangle appears in the corner of the cell.
- The status bar shows only **Count** for the selection, not **Sum**.
- `=ISNUMBER(A1)` returns `FALSE`.

## Quick fixes in Excel

**The green triangle.** Select the cells, click the warning icon that appears and choose **Convert to Number**. It works only when Excel flags the cells.

<img src="/images/articles/text-numbers-to-numeric-values/number-as-text-warning.jpg" loading="lazy" width="513" height="382" alt="The Number Stored as Text warning menu with the Convert to Number command" />

**Paste Special.** Type `1` in an empty cell and copy it. Select the text numbers, choose **Paste Special → Multiply**. Excel has to do arithmetic on each value, so it converts them.

<img src="/images/articles/text-numbers-to-numeric-values/paste-special-multiply.jpg" loading="lazy" width="792" height="710" alt="The Paste Special dialog with the Multiply operation selected" />

**Text to Columns.** Select one column, open **Data → Text to Columns** and click **Finish** right away. This re-parses the values and converts numbers and dates. It is a fast trick for a single column.

<img src="/images/articles/text-numbers-to-numeric-values/text-to-columns.jpg" loading="lazy" width="1028" height="632" alt="The last step of the Text to Columns wizard with the General column format" />

**A formula.** `=VALUE(A1)` or the shorter `=--A1` returns a number. If the source has odd spaces, combine it with cleaning: `=VALUE(TRIM(SUBSTITUTE(A1,CHAR(160),"")))`. Non-breaking spaces (`CHAR(160)`) are common in data from websites and thousand separators in Russian or French formats, and `TRIM` does not remove them.

## Things that go wrong

- **Decimal separator.** `1,5` is text on a system that expects `1.5`. Replace the separator first, or use `NUMBERVALUE(A1, ",", " ")` to state the decimal and group separators explicitly.
- **Leading apostrophe.** It is invisible in the cell but marks it as text. Paste Special → Multiply handles it.
- **Minus at the end.** Some systems export negative numbers as `100-`. Convert with `=IF(RIGHT(A1)="-",-LEFT(A1,LEN(A1)-1),A1)`.
- **Leading zeros.** Phone numbers, postal codes and product codes should stay text. Converting them removes the zeros for good, so check before you convert a whole column.
- **Cell format.** Changing a cell's format to Number does not convert what is already inside. You need to re-enter the value (<kbd>F2</kbd>, <kbd>Enter</kbd>) or use one of the methods above.

## With the add-in

The **Parse Numerics** command converts the selected text cells to numbers in one step, including the cases with non-breaking spaces. If the result is wrong, **Undo Last Action** restores the previous values.

The add-in command for this: [Parse Numerics](/en/features/parse-numerics/).
