![](images/ru.png) [Russian](README_RU.md) | ![](images/en.png) **English**

# Navferty's Excel Add-In

A free Excel add-in that enhances productivity with tools for data manipulation, formatting, and validation.

![Navferty's Tools Ribbon Tab in MS Excel](images/NavfertyToolsRibbonEn.png)

## Features
 - [Parse Numerics](#parse-numerics)
 - [Replace (Transliteration or Analogues)](#replace)
 - [Stringify Numerics](#stringify-numerics)
 - [Toggle Case](#toggle-case)
 - [Trim Spaces](#trim-spaces)
 - [Unprotect Workbook](#unprotect-workbook)
 - [Worksheets Protection](#worksheets-protection)
 - [Highlight Duplications](#highlight-duplications)
 - [Unmerge Cells](#unmerge-cells)
 - [Find All Cells Containing Errors](#find-all-cells-containing-errors)
 - [Copy as Markdown](#copy-as-markdown)
 - [Validate Values](#validate-values)
 - [Export to SQLite](#export-to-sqlite)
 - [Create Sample XML based on XSD](#create-sample-xml-based-on-xsd)
 - [Validate XML with XSD](#validate-xml-with-xsd)
 - [Undo Last Action](#undo-last-action)

## Installation

Download, unpack and run [the installation file](https://deploy.navferty.ru/setup.zip) to install the add-in for Excel.

After installation, restart Excel if it was open. You should see the new **Navferty's Tools** tab in the ribbon.

**Note:** If you encounter a security warning about the certificate, see [detailed installation instructions](#how-to-install-the-add-in) below.

---

## Parse Numerics

Convert numbers stored as text to proper numeric format, making them usable in calculations.

<details>
  <summary>View screenshots</summary>

![Parse Numerics - Before](images/ParseNumerics1.png)
![Parse Numerics - After](images/ParseNumerics2.png)
</details>

## Replace

Replace Russian characters using different transformation methods.

### With Transliteration

Complete transliteration of Russian alphabet to English. For example: "Ж" → "Zh", "Щ" → "Shch". Based on ICAO Doc [9303](https://www.icao.int/publications/Documents/9303_p3_cons_en.pdf).

<details>
  <summary>View screenshots</summary>

![Transliteration - Before](images/Transliterate1.png)
![Transliteration - After](images/Transliterate2.png)
</details>

### Replace Chars

Replace only visually similar letters: А↔A, В↔B, Е↔E, К↔K, М↔M, Н↔H, О↔O, Р↔P, С↔C, Т↔T, У↔Y, Х↔X.

<details>
  <summary>View screenshots</summary>

![Replace Chars - Before](images/ReplaceChars1.png)
![Replace Chars - After](images/ReplaceChars2.png)
</details>

## Stringify Numerics

Convert numbers to their word representation in multiple languages:
- Russian
- English
- French

<details>
  <summary>View screenshots</summary>

![Stringify Numerics - Input](images/Stringify1.png)
![Stringify Numerics - Output](images/Stringify2.png)
</details>

[Up](#navfertys-excel-add-in)

---

## Toggle Case

Cycle through text case transformations: `Abcde` → `abcde` → `ABCDE`

<details>
  <summary>View screenshots</summary>

![Toggle Case - Example 1](images/ToggleCase1.png)
![Toggle Case - Example 2](images/ToggleCase2.png)
![Toggle Case - Example 3](images/ToggleCase3.png)
![Toggle Case - Example 4](images/ToggleCase4.png)
</details>

## Trim Spaces

Remove extra spaces, line breaks, and leading/trailing whitespace from text cells.

<details>
  <summary>View screenshots</summary>

![Trim Spaces - Before](images/TrimSpaces1.png)
![Trim Spaces - After](images/TrimSpaces2.png)
</details>

[Up](#navfertys-excel-add-in)

---

## Unprotect Workbook

Remove protection from all worksheets, workbook structure, and VBA projects without needing passwords. 

*Note: Does not work with encrypted workbooks.*

## Worksheets Protection

Set or remove protection on multiple worksheets at once.

## Highlight Duplications

Automatically color-code duplicate values in the selected range. Different colors indicate different duplicate groups.

<details>
  <summary>View screenshot</summary>

![Highlighted duplicates with different colors](images/Duplicates.png)
</details>

## Unmerge Cells

Unmerge cells while filling each cell with the original merged value.

<details>
  <summary>View screenshots</summary>

![Unmerge Cells - Before](images/Unmerge1.png)
![Unmerge Cells - After](images/Unmerge2.png)
</details>

## Find All Cells Containing Errors

Locate all formula errors in the selected range: `#N/A`, `#NAME?`, `#DIV/0!`, `#REF!`, `#VALUE!`, `#NUM!`, `#NULL!`

<details>
  <summary>View screenshot</summary>

![Error cells highlighted](images/FindErrorValues.png)
</details>

## Copy as Markdown

Copy selected cells to clipboard as a markdown-formatted table.

<details>
  <summary>View screenshot</summary>

![Markdown export example](images/ExportToMarkdown.png)
</details>

## Validate Values

Check cell values against specific formats:
- Number
- Date
- TIN of an individual (12 digits with verification)
- TIN of a legal entity (10 digits with verification)
- Text for XML (no invalid characters)

*Note: Valid TIN format doesn't guarantee the organization or individual exists.*

<details>
  <summary>View screenshot</summary>

![Validation results](images/Validate.png)
</details>

[Up](#navfertys-excel-add-in)

---

## Export to SQLite

Export the entire workbook to a SQLite database file. Each worksheet becomes a separate table with automatic column type detection.

<details>
  <summary>View screenshots</summary>

![Export dialog](images/ExportToSqlite1.png)
![SQLite database result](images/ExportToSqlite2.png)
</details>

[Up](#navfertys-excel-add-in)

---

## Create Sample XML based on XSD

Generate a sample XML file from an XSD schema to visualize the structure.

<details>
  <summary>View example</summary>

**Input XSD Schema:**

```XML
<xsd:schema xmlns:xsd="http://www.w3.org/2001/XMLSchema"
           xmlns:tns="http://tempuri.org/PurchaseOrderSchema.xsd"
           targetNamespace="http://tempuri.org/PurchaseOrderSchema.xsd"
           elementFormDefault="qualified">
 <xsd:element name="PurchaseOrder" type="tns:PurchaseOrderType"/>
 <xsd:complexType name="PurchaseOrderType">
  <xsd:sequence>
   <xsd:element name="ShipTo" type="tns:USAddress" maxOccurs="2"/>
   <xsd:element name="BillTo" type="tns:USAddress"/>
  </xsd:sequence>
  <xsd:attribute name="OrderDate" type="xsd:date"/>
 </xsd:complexType>

 <xsd:complexType name="USAddress">
  <xsd:sequence>
   <xsd:element name="name"   type="xsd:string"/>
   <xsd:element name="street" type="xsd:string"/>
   <xsd:element name="city"   type="xsd:string"/>
   <xsd:element name="state"  type="xsd:string"/>
   <xsd:element name="zip"    type="xsd:integer"/>
  </xsd:sequence>
  <xsd:attribute name="country" type="xsd:NMTOKEN" fixed="US"/>
 </xsd:complexType>
</xsd:schema>
```

**Generated XML Output:**

```XML
<PurchaseOrder xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" OrderDate="1900-01-01" xmlns="http://tempuri.org/PurchaseOrderSchema.xsd">
  <ShipTo country="US">
    <name>name1</name>
    <street>street1</street>
    <city>city1</city>
    <state>state1</state>
    <zip>1</zip>
  </ShipTo>
  <ShipTo country="US">
    <name>name2</name>
    <street>street2</street>
    <city>city2</city>
    <state>state2</state>
    <zip>-79228162514264337593543950335</zip>
  </ShipTo>
  <BillTo country="US">
    <name>name1</name>
    <street>street1</street>
    <city>city1</city>
    <state>state1</state>
    <zip>1</zip>
  </BillTo>
</PurchaseOrder>
```
</details>

## Validate XML with XSD

Validate an XML file against an XSD schema. A detailed error report will be created in a new workbook.

<details>
  <summary>View sample error report</summary>

Sample error report:

|Severity|Element|Message|
|---|---|---|
|Error|city|The element 'shipto' has invalid child element 'city'. List of possible elements expected: 'address'.|
|Error|quantity|The 'quantity' element is invalid - The value '-5' is invalid according to its datatype 'http://www.w3.org/2001/XMLSchema:positiveInteger' - Value '-5' was either too large or too small for PositiveInteger.|
|Error|price|The 'price' element is invalid - The value 'asdasd' is invalid according to its datatype 'http://www.w3.org/2001/XMLSchema:decimal' - The string 'не число' is not a valid Decimal value.|

</details>

[Up](#navfertys-excel-add-in)

---

## Undo Last Action

Undo the last action performed with this add-in. Available for some functions in the 'Converting values' and 'Formatting values' sections, only if cells haven't been edited since.

[Up](#navfertys-excel-add-in)

---

## How to Install the Add-In

### Quick Install (Recommended)

1. Download and unpack [setup.exe](https://deploy.navferty.ru/setup.zip)
2. Run the installer and follow the wizard
3. Restart Excel if it was open during installation
4. You'll see the new **Navferty's Tools** tab in the ribbon

### Certificate Installation (if needed)

If you encounter a security warning about the certificate:


1. Right-click on `setup.exe` and select **Properties**
2. Go to the **Digital Signatures** tab
3. Select the signature and click **Details**
4. Click **View Certificate** → **Install Certificate**
5. Choose **Local Machine** and place it in **Trusted Root Certification Authorities**

*Note: You can remove the certificate after installation, but it will be required for future updates.*

### Advanced: Offline Install

The solution is built in Azure. You can download the full archive with installation files:

1. Visit https://navferty.visualstudio.com/NavfertyExcelAddIn/_build?definitionId=3
2. Select the latest build of 'NavfertyExcelAddIn - Publish' pipeline
3. Download published installation files
4. Extract files to a folder and run the '.vsto' file

> **Important:** Using the desktop folder is recommended - updates are only permitted from the same folder where the add-in was first installed.

After installation, restart Excel and you'll see the new **Navferty's Tools** tab.

[Up](#navfertys-excel-add-in)
