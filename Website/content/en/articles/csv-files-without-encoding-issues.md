{
  "title": "Working with CSV files without encoding issues",
  "description": "Why Cyrillic and accented characters turn into garbage and how to open and save CSV correctly.",
  "image": "https://navferty.net/images/articles/csv-files-without-encoding-issues/wrong-encoding-csv-opened-in-excel.jpg"
}

[← All articles](/en/articles/)

# Working with CSV files without encoding issues

Why Cyrillic and accented characters turn into garbage and how to open and save CSV correctly.

You open a CSV in Excel and see `РџСЂРёРІРµС‚` instead of the expected text, or `Ã©` in place of `é`. The file is fine. Excel has simply guessed the wrong encoding.

<img src="/images/articles/csv-files-without-encoding-issues/wrong-encoding-csv-opened-in-excel.jpg" loading="lazy" width="822" height="346" alt="A CSV file opened by double-click, with accented letters such as René and Müller garbled" />

## What is happening

A CSV file is plain text, and nothing in it says which encoding is used. Double-clicking a file makes Excel use the system's legacy code page (for example Windows-1251 on Russian Windows, Windows-1252 on Western European systems), unless the file starts with a UTF-8 byte-order mark (BOM). A UTF-8 file without a BOM is therefore misread.

## Opening a file correctly

1.  Do not double-click. Open a blank workbook.
2.  **Data → From Text/CSV**.
3.  Select the file. In the preview, change **File Origin** until the text looks right (usually **65001: Unicode (UTF-8)**, or **1251: Cyrillic (Windows)** for older Russian files).
4.  Check the delimiter, then click **Load** or **Transform Data**.

<img src="/images/articles/csv-files-without-encoding-issues/csv-import-with-encoding.jpg" loading="lazy" width="1151" height="663" alt="The From Text/CSV preview with File Origin set to 65001: Unicode (UTF-8), and the accented letters shown correctly" />

<img src="/images/articles/csv-files-without-encoding-issues/csv-import-result.jpg" loading="lazy" width="391" height="207" alt="The loaded data on the sheet, with René, Müller, São Paulo and Café intact" />

## Saving a file correctly

- **CSV UTF-8 (Comma delimited)** in the Save As type list writes UTF-8 with a BOM. It is the one to choose when the file goes to another program or when it opens in Excel on another computer.
- The plain **CSV (Comma delimited)** option saves in the system code page and loses characters that are not part of it.
- Some strict importers reject the BOM. Then remove it with a text editor such as Notepad++ (**Encoding → UTF-8**) or save from there.

## Delimiter problems

In many European locales, including Russian and German, Excel uses a semicolon as the list separator, since the comma is the decimal separator. A comma-separated file opened by double-click will land in one column, and a semicolon-separated file opened on an English system will do the same. In the From Text/CSV dialog, set the delimiter explicitly. In a text file, an `sep=;` line at the top forces Excel to use a semicolon, but other tools will see it as data, so use that only for files meant for Excel.

## Other traps

- **Leading zeros** disappear when Excel guesses the column as a number. In the import dialog, set the column type to Text.
- **Long numbers** (more than 15 digits) are rounded and shown in scientific notation. Import them as text.
- **Dates** are converted with the local day/month order. `03/04/2024` may mean the third of April or the fourth of March.
- **Fields with line breaks or delimiters** must be in quotes, and quotes inside doubled. Programs that write CSV by concatenating strings often miss this.
- **Saving a CSV over and over** in Excel changes the content (numbers, dates, zeros). Keep the original and the working .xlsx separate.

## For recurring files

Use Power Query: it remembers the encoding, delimiter and column types, so the next file is a refresh away.
