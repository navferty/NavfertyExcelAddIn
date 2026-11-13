{
  "title": "Excel keyboard shortcuts that save the most time",
  "description": "A practical set of shortcuts for editing, navigating, selecting, filling and pasting in Excel, grouped by task.",
  "image": "https://navferty.net/images/NavfertyToolsRibbonEn.png"
}

[← All articles](/en/articles/)

# Excel keyboard shortcuts that save the most time

A practical set of shortcuts for editing, navigating, selecting, filling and pasting in Excel, grouped by task.

Most of the time saved with the keyboard comes from a small set of shortcuts used constantly. These are grouped by what you are doing. Learn a few from one group, use them for a week, then add more.

## Ribbon shortcuts with Alt KeyTips

In desktop Excel for Windows, press and release <kbd>Alt</kbd> to show small KeyTip labels for the ribbon tabs and other controls:

<img src="/images/articles/excel-keyboard-shortcuts/alt-keytips-en.png" alt="Excel showing KeyTips for ribbon tabs after pressing Alt, including H for Home" width="610" height="229" loading="lazy">

Press the letter shown for a tab to open it. Excel then displays new KeyTips for that tab's commands; follow the labels one step at a time. For example, in English Excel, select a cell in a column and press <kbd>Alt</kbd>, <kbd>H</kbd>, <kbd>O</kbd>, <kbd>I</kbd> in sequence to AutoFit the column width. Release each key before pressing the next; do not hold them all together. Press <kbd>Esc</kbd> to back out.

The add-in also has a ribbon KeyTip: press <kbd>Alt</kbd>, <kbd>Z</kbd>, <kbd>X</kbd> to open **Navferty's Tools**, then follow the displayed labels to choose a command.

KeyTip letters can differ with the Excel interface language and version. Use the labels shown in your installation rather than assuming an English sequence works everywhere.

## Editing a cell

- <kbd>F2</kbd> – edit the cell in place, with the cursor at the end of the text. Press it again to switch between "enter" and "edit" modes, where the arrow keys move the cursor or select another cell.
- <kbd>Alt+Enter</kbd> – new line inside a cell.
- <kbd>Ctrl+Enter</kbd> – confirm and fill the whole selection with the same value or formula.
- <kbd>F4</kbd> – while typing a reference, cycle between `A1`, `$A$1`, `A$1` and `$A1`. Outside of editing it repeats the last action.
- <kbd>Ctrl+;</kbd> – insert today's date as a fixed value. <kbd>Ctrl+Shift+;</kbd> inserts the current time.
- <kbd>Ctrl+Z / Ctrl+Y</kbd> – undo and redo.

## Moving around

- <kbd>Ctrl+Arrow</kbd> – jump to the edge of the current block of data. Pressed in an empty area it jumps to the next filled cell.
- <kbd>Ctrl+Home / Ctrl+End</kbd> – go to `A1` or to the last used cell.
- <kbd>Ctrl+PgDn / Ctrl+PgUp</kbd> – next and previous sheet.
- <kbd>F5 or Ctrl+G</kbd> – Go To: type a cell address or a name. The **Special…** button selects cells by type (blanks, errors, formulas, visible cells only).
- <kbd>Ctrl+F / Ctrl+H</kbd> – find and replace.

## Selecting

- <kbd>Ctrl+Space</kbd> – select the whole column. If it switches the keyboard layout instead, change that shortcut in the Windows language settings.
- <kbd>Shift+Space</kbd> – select the whole row.
- <kbd>Ctrl+A</kbd> – select the current data region; pressed again, the whole sheet.
- <kbd>Shift+Arrow</kbd> – extend the selection by one cell. Combined with <kbd>Ctrl</kbd> it extends to the edge of the data: <kbd>Ctrl+Shift+Arrow</kbd>.
- <kbd>Ctrl+Shift+End</kbd> – extend the selection to the last used cell.
- <kbd>Alt+;</kbd> – select only visible cells. Useful before copying a filtered list, so hidden rows are not copied.

## Filling and formulas

- <kbd>Ctrl+D</kbd> – fill down: copy the top cell of the selection into the cells below. <kbd>Ctrl+R</kbd> fills right.
- <kbd>Alt+=</kbd> – insert `SUM` for the neighbouring range. With several cells selected it adds totals to all of them at once.
- <kbd>Ctrl+E</kbd> – Flash Fill: type the result for the first row, then press it to continue the pattern.
- <kbd>Ctrl+\`</kbd> – show formulas instead of values (the key to the left of <kbd>1</kbd>). Press again to switch back.
- <kbd>Ctrl+Shift+Enter</kbd> – enter an array formula in older versions of Excel.

## Copying and pasting

- <kbd>Ctrl+C / Ctrl+X / Ctrl+V</kbd> – copy, cut, paste.
- <kbd>Ctrl+Alt+V</kbd> – Paste Special. In the dialog, press the underlined letter of the option, for example to paste values only or to multiply by a number (see [converting text numbers to numeric values](/en/articles/text-numbers-to-numeric-values/)).
- <kbd>Ctrl+Shift++ / Ctrl+-</kbd> – insert or delete cells, rows or columns. Select a whole row or column first.
- <kbd>Ctrl+9 / Ctrl+0</kbd> – hide the selected rows or columns.

## Formatting and files

- <kbd>Ctrl+1</kbd> – Format Cells dialog.
- <kbd>Ctrl+T</kbd> – turn the range into a table (see [cleaning data](/en/articles/best-practices-for-cleaning-data-in-excel/)).
- <kbd>Ctrl+Shift+L</kbd> – toggle filters.
- <kbd>Ctrl+B / Ctrl+I / Ctrl+U</kbd> – bold, italic, underline.
- <kbd>F12</kbd> – Save As.
- <kbd>Ctrl+S</kbd> – save. Worth pressing every few minutes.

Many shortcuts are the same in English and Russian Excel, but ribbon KeyTips can differ. On a Mac most <kbd>Ctrl</kbd> shortcuts use <kbd>Cmd</kbd>, and some differ, so check the Excel help for your version.

Related: [Automating routine Excel tasks without VBA](/en/articles/automating-excel-tasks-without-vba/).
