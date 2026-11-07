{
  "title": "Pivot tables in Excel: a practical introduction",
  "description": "How to build a first pivot table, summarise by category and by date, switch between sum and count, and keep it up to date.",
  "image": "https://navferty.net/images/articles/pivot-tables/inserting-pivot-table.jpg"
}

[← All articles](/en/articles/)

# Pivot tables in Excel: a practical introduction

How to build a first pivot table, summarise by category and by date, switch between sum and count, and keep it up to date.

A pivot table summarises a long list without a single formula: totals by category, counts, averages, by month or year. It is worth learning because the same task with `SUMIFS` means a formula for every cell. This article covers the basics only; the feature is deep, and the <a href="https://support.microsoft.com/en-us/office/create-a-pivottable-to-analyze-worksheet-data-a9a84538-bfe9-40a9-a8e9-f99134456576" rel="noopener" target="_blank">Microsoft documentation</a> goes further.

## Prepare the source data

A pivot table needs a flat list: one header row, one record per row, no empty rows or columns, no merged cells, no subtotals inside. The example below has four columns: *Order Id*, *Order Date*, *Department Id* and *Amount*. If the data is messy, see [cleaning data](/en/articles/best-practices-for-cleaning-data-in-excel/) first. Dates and amounts must be real dates and numbers, not text ([converting text numbers](/en/articles/text-numbers-to-numeric-values/)).

It helps to turn the list into a table first (<kbd>Ctrl</kbd>+<kbd>T</kbd>). A pivot table built on a table picks up new rows after a refresh, and does not include blank rows.

## Create the pivot table

Select any cell in the data and choose **Insert → PivotTable**. Check the range Excel suggests and choose where to put the result; a new worksheet is the safe default.

<img src="/images/articles/pivot-tables/inserting-pivot-table.jpg" loading="lazy" width="1097" height="825" alt="The Insert PivotTable dialog with the source range SourceData!$A:$D and a New Worksheet option" />

In this example the range is written as whole columns (`SourceData!$A:$D`). It works, but the pivot table will then contain a `(blank)` item for the empty cells below the data. A table or an exact range such as `$A$1:$D$25` avoids it.

## Build the report

The empty pivot table shows the **PivotTable Fields** pane. The four areas below the field list decide what goes where:

- **Rows** – the values that become row labels.
- **Columns** – the values that become column headings.
- **Values** – what to calculate (sum, count, average).
- **Filters** – a drop-down above the table to limit the whole report.

<img src="/images/articles/pivot-tables/fields-selection.jpg" loading="lazy" width="1102" height="787" alt="An empty PivotTable with the field list and the Filters, Columns, Rows and Values areas" />

Drag *Department Id* to **Rows** and *Amount* to **Values**. This is the first example: total amount per department. Excel chooses **Sum** for numbers.

## Count instead of sum

To count orders per department, click the field in the **Values** area, choose **Value Field Settings** and pick **Count**. Counting works on any column, and for text columns Count is the default. The same dialog has Average, Max, Min and the **Show Values As** tab with percentages of the total.

## Group dates by year or month

Drag *Order Date* to **Columns**. In recent versions of Excel the dates are grouped automatically (Years, Quarters, Months fields appear). If they are not, right-click a date and choose **Group**. Here the report has departments in the rows, years in the columns and the sum of amounts in between:

<img src="/images/articles/pivot-tables/pivot-configuration.jpg" loading="lazy" width="1311" height="833" alt="A pivot table with departments in rows, years in columns and Sum of Amount in values, with the fields pane showing the Columns, Rows and Values areas" />

The grouping only works if every value in the column is a real date. One text value or an empty cell and Excel will not offer it.

## Refresh

A pivot table is a snapshot of the data and does not recalculate when the source changes. Right-click it and choose **Refresh**, or press <kbd>Alt</kbd>+<kbd>F5</kbd>. **Data → Refresh All** (<kbd>Ctrl</kbd>+<kbd>Alt</kbd>+<kbd>F5</kbd>) refreshes every pivot table in the workbook.

## Limitations to know about

- You cannot type over the values in the pivot area; the numbers come from the source.
- New rows are included only if they fall inside the source range, which is the point of using a table.
- A column format (like a thousands separator) can reset on refresh. Set it through **Value Field Settings → Number Format**.
- Read cells from a pivot with `GETPIVOTDATA`. Excel inserts it for you when you click a pivot cell from a formula; its address-style alternative breaks when the layout changes.

Related: [VLOOKUP](/en/articles/vlookup-in-excel/) for bringing in names for the ids.
