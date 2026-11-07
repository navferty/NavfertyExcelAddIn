{
  "title": "Preparing Excel reports for management",
  "description": "How to make a workbook that the reader understands in a minute.",
  "image": "https://navferty.net/images/NavfertyToolsRibbonEn.png"
}

[← All articles](/en/articles/)

# Preparing Excel reports for management

How to make a workbook that the reader understands in a minute.

A manager reading your report wants to know what happened, whether it is good or bad, and what to do about it. Any time spent on a cell colour should serve that.

## Start with the question

Before opening Excel, write down what the reader needs to decide. A report without a question tends to become a collection of every number available.

## Structure of the workbook

- **Summary sheet first**: three to seven key figures, with the previous period or plan next to them, and a sentence of conclusion.
- **Details after**, in the order the reader would ask the next question.
- **Raw data and calculations** on separate sheets, so the reader is not required to look at them.
- Sheets have meaningful names, not `Sheet1` and `Copy (2)`.

## Tables

- Align numbers to the right, with the same number of decimals in a column and a thousands separator.
- Round to the precision that matters. Thousands or millions in a header (`USD, thousand`) read better than long numbers.
- Use a few borders. Light horizontal lines and a bold header are usually enough.
- Freeze the header row on long tables.
- Show totals clearly and make sure they add up. A table whose total differs by one from the sum of the rows because of rounding invites questions. Use `ROUND` in the calculation, or add a note.

## Charts

- Pick the type for the question: lines for change over time, bars for comparison, and avoid pies with more than four slices.
- The title says the point: “Sales fell 8% in Q3”, not “Sales by quarter”.
- Start bar chart axes at zero. Remove 3D, gradients and extra gridlines.
- Label directly instead of using a legend when you can.

## Conditional formatting, sparingly

Highlight what needs attention (plan missed, overdue), not everything. When ten cells are red, nothing is.

## Make it refreshable

You will do this again next month. Use tables, Power Query and pivot tables so that updating means pasting new data and clicking **Refresh All**. A manually assembled report is likely to contain a copy-paste mistake sooner or later.

## Before sending

Go through the [checklist](/en/articles/checklist-validating-excel-files/), set the print area, and consider sending a PDF if the reader only needs to read. If they need to interact, send the workbook with the summary sheet active.
