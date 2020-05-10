# Navferty's Excel Add-In

## Common tools for MS Excel ##

![Navferty's Tools Ribbon Tab in MS Excel](images/NavfertyToolsRibbonEn.png)


* Common Functions:
    * [Highlight Duplicates (different colors for groups of same values)](#highlight-duplications)
    * [Parse Numerics (convert numbers stored as text into numeric values)](#parse-numerics)
    * [Toggle Case (toggle lowercase-UPPERCASE-Camel Case)](#toggle-case)
    * [Trim Spaces (delete trailing spaces and extra line breaks in selection)](#trim-spaces)
    * [Unmerge Cells (unmerge cells and fill each cell with original value)](#unmerge-cells)
    * [Unprotect Workbook (remove protection for workbook and each worksheet)](#unprotect-workbook)
    * [Export table to markdown (table in markdown format will be placed to clipboard)](#export-table-to-markdown)
    * [Validate cell values (numerics, dates, XML text etc.)](#validate-cell-values)
    * [Find all cells containing errors on sheet](#find-all-cells-containing-errors)
    * Cut Names (make commonly used names shorter) *under development*

* [XML Functions:]()
    * [Create Sample XML based on XSD (you need to select XSD file)](#create-sample-xml-based-on-xsd)
    * [Validate XML with XSD (check selected xml based on XSD-schema)](#validate-xml-with-xsd)

* [How to install](#how-to-install)

## Highlight duplications ##
Fill different droups of duplicated values with different colors.

![Navferty's Tools Ribbon Tab in MS Excel](images/Duplicates.png)


## Parse Numerics ##
Convert numbers stored as text to numeric format.

![Navferty's Tools Ribbon Tab in MS Excel](images/ParseNumerics1.png)
![Navferty's Tools Ribbon Tab in MS Excel](images/ParseNumerics2.png)


## Toggle Case ##
Toggle text case in selected cells (UPPERCASE->lowercase->Camel Case).

![Navferty's Tools Ribbon Tab in MS Excel](images/ToggleCase1.png)
![Navferty's Tools Ribbon Tab in MS Excel](images/ToggleCase2.png)
![Navferty's Tools Ribbon Tab in MS Excel](images/ToggleCase3.png)
![Navferty's Tools Ribbon Tab in MS Excel](images/ToggleCase4.png)


## Unmerge Cells ##
![Navferty's Tools Ribbon Tab in MS Excel](images/Unmerge1.png)
![Navferty's Tools Ribbon Tab in MS Excel](images/Unmerge2.png)

Unmerge cells and fill each cell of merge area with initial value.


## Trim Spaces ##
Trim spaces in text values, remove extra space symbols and new line symbols. Delete values in empty cells.

![Navferty's Tools Ribbon Tab in MS Excel](images/TrimSpaces1.png)
![Navferty's Tools Ribbon Tab in MS Excel](images/TrimSpaces2.png)


## Validate cell values ##
Check values in selected cells as numerics, valid dates, valid text for XML contents, russian TIN (known as 'ИНН') etc.

![Navferty's Tools Ribbon Tab in MS Excel](images/Validate.png)


## Find all cells containing errors ##
Find all cells with errors like '#VALUE!', '#REF!' etc. on current worksheet.

![Navferty's Tools Ribbon Tab in MS Excel](images/FindErrorValues.png)


## Unprotect Workbook ##
Remove protection without password from workbook and all worksheets, unlock VBA project if it exists.

![Navferty's Tools Ribbon Tab in MS Excel](images/Unprotect1.png)
![Navferty's Tools Ribbon Tab in MS Excel](images/Unprotect2.png)


## Export table to markdown ##
Contents of celected cells will be copied to clipboard in markdown format.

![Navferty's Tools Ribbon Tab in MS Excel](images/ExportToMarkdown.png)


## Create Sample XML based on XSD ##
Select file with an XSD schema and create a sampe XML based on that shema.

![Navferty's Tools Ribbon Tab in MS Excel](images/SampleXml1.png)
![Navferty's Tools Ribbon Tab in MS Excel](images/SampleXml2.png)



## Validate XML with XSD ##
Check XML file with XSD schema. Select xml and xsd files, and report with all validation errors and warnings will be created in new workbook.

![Navferty's Tools Ribbon Tab in MS Excel](images/ValidateXml1.png)
![Navferty's Tools Ribbon Tab in MS Excel](images/ValidateXml2.png)
![Navferty's Tools Ribbon Tab in MS Excel](images/ValidateXml3.png)


# How to install #

The solution is build in Azure, you can download installation files from there.
Visit https://navferty.visualstudio.com/NavfertyExcelAddIn/_build?definitionId=3
then select latest build and download installation files as build artifacts:

![Navferty's Tools Ribbon Tab in MS Excel](images/Install1.png)
![Navferty's Tools Ribbon Tab in MS Excel](images/Install2.png)

Extract files to a folder and run '.vsto' file. *Using desktop folder is recommended - for installing updates you will need to do it from the same folder that was used to install add-in for the first time!*

![Navferty's Tools Ribbon Tab in MS Excel](images/Install3.png)

After installation process is completed, run (or restart) Excel application, and you will see new tab:

![Navferty's Tools Ribbon Tab in MS Excel](images/Install4.png)

*Used icons are designed by iconarchive, Flaticon*
*"Find errors" icon made by turkkub from www.flaticon.com*
