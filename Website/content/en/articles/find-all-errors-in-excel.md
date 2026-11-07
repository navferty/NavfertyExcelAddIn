{
  "title": "How to find all errors in Excel (#VALUE!, #REF!, #DIV/0!)",
  "description": "Ways to locate every formula error in a sheet, and what each error usually means.",
  "image": "https://navferty.net/images/articles/find-all-errors-in-excel/sample-range-with-errors.jpg"
}

[← All articles](/en/articles/)

# How to find all errors in Excel (#VALUE!, \#REF!, \#DIV/0!)

Ways to locate every formula error in a sheet, and what each error usually means.

A report that goes out with a `#REF!` in the totals looks bad, and in a large workbook a single error cell is easy to miss. Excel has a few built-in ways to find them all.

Example range used below, with the common error types:

<img src="/images/articles/find-all-errors-in-excel/sample-range-with-errors.jpg" loading="lazy" width="308" height="398" alt="A sheet with the errors #DIV/0!, #REF!, #VALUE!, #N/A, #NAME? and #NUM! mixed with ordinary numbers" />

## Go To Special

Select the range (or press <span class="kbd">Ctrl</span>+<span class="kbd">A</span> for the whole sheet), then press <span class="kbd">F5</span> (or <span class="kbd">Ctrl</span>+<span class="kbd">G</span>) → **Special…** → **Formulas** and leave only **Errors** ticked. Excel selects every cell that currently shows an error. You can then fill them with a colour or step through them with <span class="kbd">Enter</span>.

<img src="/images/articles/find-all-errors-in-excel/go-to-main.jpg" loading="lazy" width="488" height="412" alt="The Go To dialog with the Special button" />

<img src="/images/articles/find-all-errors-in-excel/go-to-special.jpg" loading="lazy" width="388" height="451" alt="The Go To Special dialog with Formulas and Errors selected" />

<img src="/images/articles/find-all-errors-in-excel/go-to-special-errors-result-selection.jpg" loading="lazy" width="382" height="381" alt="Every cell with an error is selected" />

Limitation: it works per sheet, and it gives you a selection, not a list. You still have to jump around to see what is where.

## Find with a wildcard

Press <span class="kbd">Ctrl</span>+<span class="kbd">F</span>, search for `#`, set **Look in** to **Values** and click **Find All**. This gives a list of cells with their addresses. It also finds legitimate text that contains a hash sign, so it is a bit noisy, and it does not work well if the sheet uses `#` in headers.

## Conditional formatting

For a sheet that is edited constantly, a rule makes errors visible without any searching. Select the range, choose **Home → Conditional Formatting → New Rule → Use a formula** and enter (adjust A1 to the top-left cell of your selection):

    =ISERROR(A1)

<img src="/images/articles/find-all-errors-in-excel/conditional-formatting-errors-rule.jpg" loading="lazy" width="893" height="780" alt="The New Formatting Rule dialog with the formula =ISERROR(A1)" />

Pick a red fill. Any error that appears later is highlighted right away.

**Mind the active cell.** In a conditional formatting formula, a relative reference like `A1` is relative to the *active* cell, which is not always the top-left cell of the selection. If you select the range from the bottom up, or click a cell in the middle before opening the dialog, the rule is shifted and some cells (typically the first rows) are not checked. The fix is to make the top-left cell active before creating the rule: select the range starting from its first cell, and check the Name Box (left of the formula bar) shows that address. You can also edit the rule later via **Manage Rules** and compare the formula with the top-left cell of **Applies to**.

<img src="/images/articles/find-all-errors-in-excel/conditional-formatting-errors-result.jpg" loading="lazy" width="403" height="415" alt="The result: every error cell is highlighted, including #DIV/0! and #REF!" />

## Count them

To get a quick number for a checklist or a dashboard cell:

    =SUMPRODUCT(--ISERROR(A1:Z1000))

`ISERROR` also treats `#N/A` as an error. If you want to leave `#N/A` out, use `ISERR` instead.

## What the errors mean

- `#DIV/0!` – division by zero or by an empty cell. Often a lookup that returned nothing or a month with no data yet.
- `#VALUE!` – the formula got the wrong type of value, typically text where a number was expected. A number with a trailing space or a text-formatted number is a very common cause (see [converting text numbers](/en/articles/text-numbers-to-numeric-values/)).
- `#REF!` – the formula points to a cell that no longer exists, usually because a row, column or sheet was deleted.
- `#N/A` – a lookup found no match. Check for extra spaces and for numbers stored as text on either side.
- `#NAME?` – a misspelled function name, a missing quote around text, or a function that your Excel version does not have.
- `#NUM!` – an impossible numeric operation, such as the square root of a negative number, or a result too large for Excel.
- `#NULL!` – two ranges in a formula are separated by a space instead of a comma or colon, and they do not intersect.

## Handling them

Wrapping everything in `IFERROR` makes the error go away, but it also hides real problems. Use it where an empty result is a normal outcome (a lookup that may find nothing) and let errors show everywhere else, so you notice when something breaks.

## With the add-in

The **Find All Cells Containing Errors** command scans the selected range and shows a list of all error cells with their type. Clicking an entry jumps to that cell. It is handy when the errors are spread over a big range and you want to work through them one by one.

<img src="/images/articles/find-all-errors-in-excel/find-errors-with-add-in.jpg" loading="lazy" width="1193" height="618" alt="The list of error cells found by the add-in command, with the error type and cell address" />

The add-in command for this: [Find All Cells Containing Errors](/en/features/find-errors/).
