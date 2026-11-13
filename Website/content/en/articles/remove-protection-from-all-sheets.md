{
  "title": "Removing protection from all sheets at once",
  "description": "What sheet and workbook protection actually does, and how to unprotect many sheets quickly.",
  "image": "https://navferty.net/images/NavfertyToolsRibbonEn.png"
}

[← All articles](/en/articles/)

# Removing protection from all sheets at once

What sheet and workbook protection actually does, and how to unprotect many sheets quickly.

You receive a workbook with twenty protected sheets and need to change something. Clicking **Review → Unprotect Sheet** on each one is tedious, and when nobody remembers the password it stops there.

## What kinds of protection exist

- **Sheet protection** blocks editing of locked cells, formatting, inserting rows and so on. It is a convenience feature, not security: the password is stored in the file in a form that can be bypassed.
- **Workbook structure protection** prevents adding, deleting, renaming and moving sheets.
- **VBA project protection** locks the macro code view.
- **File encryption** (**File → Info → Protect Workbook → Encrypt with Password**) is the only one that is real security. Without the password the file cannot be opened.

## When you know the password

Unprotecting each sheet by hand works but is slow. A short macro loops through sheets:

    Sub UnprotectAll()
        Dim ws As Worksheet
        For Each ws In ActiveWorkbook.Worksheets
            ws.Unprotect "password"
        Next
    End Sub

If the passwords differ between sheets, this will raise an error on the first mismatch, so wrap it in error handling.

## When you do not

First, ask the author. If the file belongs to you or your company and the author is gone, the protection can be removed because sheet protection is not encryption. An .xlsx file is a ZIP archive; renaming it to .zip and deleting the `<sheetProtection>` element from `xl/worksheets/sheetN.xml` removes the protection from that sheet. It works but takes several steps per sheet and is easy to get wrong.

Only do this with files you have the right to edit.

## With the add-in

**Unprotect Workbook** removes worksheet and workbook-structure protection from saved `.xlsx` and `.xlsm` files without a password and attempts to unlock an `.xlsm` file's VBA project. Save your changes and make a separate backup before confirming: the command closes the workbook without saving, modifies the original file on disk, and reopens it. Unsaved changes are lost, and successful VBA recovery is not guaranteed. It does not open password-encrypted workbooks.

## A note on protecting your own files

If you protect a sheet to prevent accidental edits, sheet protection is enough. If the data is confidential, use file encryption or do not send the file at all.

The add-in command for this: [Unprotect Workbook](/en/features/unprotect-workbook/), [Worksheets Protection](/en/features/worksheets-protection/).
