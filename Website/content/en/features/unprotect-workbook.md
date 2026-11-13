{
  "title": "Unprotect Workbook",
  "description": "Remove worksheet and workbook structure protection from a saved workbook and attempt to unlock its VBA project.",
  "image": "https://navferty.net/images/UnprotectWorkbook1.png"
}

[Home](/en/) › [Commands](/en/#commands) › Unprotect Workbook

# <img alt="" src="/images/icons/unprotectWorkbook.png" class="icons" width="64" height="64" />Unprotect Workbook

Remove worksheet and workbook structure protection and attempt to unlock a VBA project. *This does not open encrypted workbooks.*

<div class="shots">

<img src="/images/UnprotectWorkbook1.png" loading="lazy" width="663" height="601" alt="Excel warning about editing a protected sheet, and the Unprotect Workbook button" />

</div>

<div class="info-box">

**Related guide:** [Removing protection from all sheets at once](/en/articles/remove-protection-from-all-sheets/)

</div>

## How to use

Use this only on a workbook you are authorized to modify. Save the workbook, make a backup, then choose **Remove File Protection** and confirm the warning. The command closes the workbook without saving further changes, edits the saved file, and reopens it.

For example, it can remove worksheet and workbook-structure protection from a saved `.xlsx` file. It also attempts to modify VBA project protection in an `.xlsm` file.

**Important:** Unsaved changes are lost on confirmation. The command modifies the original file and supports `.xlsx` and `.xlsm`, not password-to-open encryption. Work on a backup, especially for files with macros; do not assume VBA projects will always be recovered successfully.

The add-in is free and open source. [Installation instructions](/en/install/).
