![](images/ru.png) [Russian](README_RU.md) | ![](images/en.png) **English**

# Navferty's Excel Add-In

![Navferty's Tools Ribbon Tab in MS Excel](images/NavfertyToolsRibbonEn.png)

## Available features
 - [Undo Last Action](#undo-last-action)
 - [Parse Numerics](#parse-numerics)
 - [Replace Chars (Transliteration or analogues)](#replace)
 - [Stringify Numerics to Words](#stringify-numerics-into-words)
 - [Toggle Case](#toggle-case)
 - [Trim Spaces](#trim-spaces)
 - [Unprotect Workbook](#unprotect-workbook)
 - [Highlight Duplicates](#highlight-duplications)
 - [Unmerge Cells](#unmerge-cells)
 - [Find Formula Errors in the selected range](#find-all-cells-containing-errors)
 - [Copy as Markdown](#copy-as-markdown)
 - [Validate Values](#validate-values)
 - [Create Sample XML based on XSD](#create-sample-xml-based-on-xsd)
 - [Validate XML with XSD](#validate-xml-with-xsd)

## [How to install](#how-to-install-the-add-in)

---

## Undo Last Action

|||
|:-:|---|
|![](images/icons/undo.png)|Undo the last action performed with this add-in. Canceling is possible for some functions in the 'Converting values' and 'Formatting values' sections, and only if the range of cells was not edited after the action was performed.|

[Up](#navfertys-excel-add-in)

---

## Parse Numerics

|||
|:-:|---|
|![](images/icons/parseNumerics.png)|Convert numbers stored as text to numeric format.|

<details>
  <summary>View screenshots</summary>

![Navferty's Tools Ribbon Tab in MS Excel](images/ParseNumerics1.png)
![Navferty's Tools Ribbon Tab in MS Excel](images/ParseNumerics2.png)
</details>

## Replace

|||
|:-:|---|
|![](images/icons/replace.png)|Replace Russian characters in the match table.|

*   ### Transliteration

    |||
    |:-:|---|
    |![](images/icons/transliterate.png)|The entire Russian alphabet is completely changed to English. For example, the letter "Ж" will be replaced with "Zh", and the letter "Щ" with "Shch". Based on ICAO Doc [9303](https://www.icao.int/publications/Documents/9303_p3_cons_en.pdf).|

    <details>
      <summary>View screenshots</summary>

    ![Navferty's Tools Ribbon Tab in MS Excel](images/Transliterate1.png)
    ![Navferty's Tools Ribbon Tab in MS Excel](images/Transliterate2.png)
    </details>

* ### Replace Chars

    |||
    |:-:|---|
    |![](images/icons/replaceChars.png)|Only uppercase letters of the alphabets will be replaced, such as: Аа, Вв, Ее, Кк, Мм, Нн, Оо, Рр, Сс, Тт, Уу, Хх.|

    <details>
      <summary>View screenshots</summary>

    ![Navferty's Tools Ribbon Tab in MS Excel](images/ReplaceChars1.png)
    ![Navferty's Tools Ribbon Tab in MS Excel](images/ReplaceChars2.png)
    </details>

## Stringify Numerics into Words

|||
|:-:|---|
|![](images/icons/stringifyNumerics.png)|Rewrites numeric values in the text with the decryption<br>- In Russian<br>- In English<br>- In French|

<details>
  <summary>View screenshots</summary>

![Navferty's Tools Ribbon Tab in MS Excel](images/Stringify1.png)
![Navferty's Tools Ribbon Tab in MS Excel](images/Stringify2.png)
</details>

[Up](#navfertys-excel-add-in)

---

## Toggle Case

|||
|:-:|---|
|![](images/icons/toggleCase.png)|Case switching for text values in selected cells according to the scheme:<br>`Abcde` -> `abcde` -> `ABCDE`|

<details>
  <summary>View screenshots</summary>

![Navferty's Tools Ribbon Tab in MS Excel](images/ToggleCase1.png)
![Navferty's Tools Ribbon Tab in MS Excel](images/ToggleCase2.png)
![Navferty's Tools Ribbon Tab in MS Excel](images/ToggleCase3.png)
![Navferty's Tools Ribbon Tab in MS Excel](images/ToggleCase4.png)
</details>

## Trim Spaces

|||
|:-:|---|
|![](images/icons/trimSpaces.png)|Clear the text content of the selected cells from unnecessary spaces. Removes repeated spaces and line breaks, as well as beginning and ending spaces in cells that have a text format.|

<details>
  <summary>View screenshots</summary>

![Navferty's Tools Ribbon Tab in MS Excel](images/TrimSpaces1.png)
![Navferty's Tools Ribbon Tab in MS Excel](images/TrimSpaces2.png)
</details>

[Up](#navfertys-excel-add-in)

---

## Unprotect Workbook

|||
|:-:|---|
|![](images/icons/unprotectWorkbook.png)|Allows you to unprotect all the pages of an open book as the entire book, no password, and also unlock VBA project (if any) to which the password is set. This feature does not apply to encrypted books.|

## Highlight duplications

|||
|:-:|---|
|![](images/icons/highlightDuplicates.png)|Sets the color of cells that contain duplicate values in the selected range. Different colors correspond to different groups of duplicates.|

<details>
  <summary>View a screenshot</summary>

![Navferty's Tools Ribbon Tab in MS Excel](images/Duplicates.png)
</details>

## Unmerge Cells

|||
|:-:|---|
|![](images/icons/unmergeCells.png)|Unmerge cells and fill each cell of merge area with initial value.|

<details>
  <summary>View screenshots</summary>

![Navferty's Tools Ribbon Tab in MS Excel](images/Unmerge1.png)
![Navferty's Tools Ribbon Tab in MS Excel](images/Unmerge2.png)
</details>

## Find all cells containing errors

|||
|:-:|---|
|![](images/icons/findErrors.png)|Search for all cells in the selected cells that contain calculation errors:<br><br>Excel formula errors types:<br>`#N/A`<br>`#NAME?`<br>`#DIV/0!`<br>`#REF!`<br>`#VALUE!`<br>`#NUM!`<br>`#NULL!`|

<details>
  <summary>View a screenshot</summary>

![Navferty's Tools Ribbon Tab in MS Excel](images/FindErrorValues.png)
</details>

## Copy as Markdown

|||
|:-:|---|
|![](images/icons/parseNumerics.png)|Contents of selected cells will be copied to clipboard in markdown format.|

<details>
  <summary>View a screenshot</summary>

![Navferty's Tools Ribbon Tab in MS Excel](images/ExportToMarkdown.png)
</details>

## Validate values

|||
|:-:|---|
|![](images/icons/validation.png)|<p>Check the cell values in the selected range for a specific format.<br><br>Supported format: <br>- Number<br>- Date<br>- TIN of an individual\* (12 digits, with two verification digits)<br>- TIN of the legal entity\* (10 digits, with one verification digit)<br>- Text for XML (no `<` and `>` characters or other invalid characters for XML content)<br><br>\* _- The correct TIN does not guarantee the existence of an organization or individual who would own this INN_</p>

<details>
  <summary>View a screenshot</summary>

![Navferty's Tools Ribbon Tab in MS Excel](images/Validate.png)
</details>

[Up](#navfertys-excel-add-in)

---

## Create Sample XML based on XSD

|||
|:-:|---|
|![](images/icons/createSampleXml.png)|Select file with an XSD schema and create a sampe XML based on that schema.|

For example, for the diagram below
<details>
  <summary>XML schema sample - `sample.xsd`</summary>

```XML
<?xml version="1.0" encoding="UTF-8" ?>
<xs:schema xmlns:xs="http://www.w3.org/2001/XMLSchema">

<xs:element name="shiporder">
  <xs:complexType>
    <xs:sequence>
      <xs:element name="orderperson" type="xs:string"/>
      <xs:element name="shipto">
        <xs:complexType>
          <xs:sequence>
            <xs:element name="name" type="xs:string"/>
            <xs:element name="address" type="xs:string"/>
            <xs:element name="city" type="xs:string"/>
            <xs:element name="country" type="xs:string"/>
          </xs:sequence>
        </xs:complexType>
      </xs:element>
      <xs:element name="item" maxOccurs="unbounded">
        <xs:complexType>
          <xs:sequence>
            <xs:element name="title" type="xs:string"/>
            <xs:element name="note" type="xs:string" minOccurs="0"/>
            <xs:element name="quantity" type="xs:positiveInteger"/>
            <xs:element name="price" type="xs:decimal"/>
          </xs:sequence>
        </xs:complexType>
      </xs:element>
    </xs:sequence>
    <xs:attribute name="orderid" type="xs:string" use="required"/>
  </xs:complexType>
</xs:element>

</xs:schema>
```
</details>

The following xml-file will be generated:

<details>
  <summary>XML output - `sample.xml`</summary>

```XML
<shiporder xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" orderid="orderid1">
  <orderperson>orderperson1</orderperson>
  <shipto>
    <name>name1</name>
    <address>address1</address>
    <city>city1</city>
    <country>country1</country>
  </shipto>
  <item>
    <title>title1</title>
    <note>note1</note>
    <quantity>1</quantity>
    <price>1</price>
  </item>
  <item>
    <title>title2</title>
    <note>note2</note>
    <quantity>79228162514264337593543950335</quantity>
    <price>-79228162514264337593543950335</price>
  </item>
  <item>
    <title>title3</title>
    <note>note3</note>
    <quantity>2</quantity>
    <price>79228162514264337593543950335</price>
  </item>
  <item>
    <title>title4</title>
    <note>note4</note>
    <quantity>79228162514264337593543950334</quantity>
    <price>0.9</price>
  </item>
  <item>
    <title>title5</title>
    <note>note5</note>
    <quantity>3</quantity>
    <price>1.1</price>
  </item>
</shiporder>
```
</details>

## Validate XML with XSD

|||
|:-:|---|
|![](images/icons/validateXml.png)|Check XML file with XSD schema. Select xml and xsd files, and report with all validation errors and warnings will be created in new workbook.|

Sample error report

|Severity|Element|Message|
|---|---|---|
|Error|city|The element 'shipto' has invalid child element 'city'. List of possible elements expected: 'address'.|
|Error|quantity|The 'quantity' element is invalid - The value '-5' is invalid according to its datatype 'http://www.w3.org/2001/XMLSchema:positiveInteger' - Value '-5' was either too large or too small for PositiveInteger.|
|Error|price|The 'price' element is invalid - The value 'asdasd' is invalid according to its datatype 'http://www.w3.org/2001/XMLSchema:decimal' - The string 'не число' is not a valid Decimal value.|

[Up](#navfertys-excel-add-in)

---

## How to Install the Add-In

### Online Install

You can install the add-in from official website of the project:
[navferty.ru](https://www.navferty.ru). Just download and run the setup.exe file.

You may need to manually import the self-signed certificate before the installation process can be finished.

The installation process also requires internet connection to load latest version.

### Offline Install

The solution is build in Azure, you can download full archive with installation files from there:

* Visit https://navferty.visualstudio.com/NavfertyExcelAddIn/_build?definitionId=3

* Select latest build of 'NavfertyExcelAddIn - Publish' pipeline:

![Navferty's Tools Ribbon Tab in MS Excel](images/Install1.png)

* Download published installation files:

![Navferty's Tools Ribbon Tab in MS Excel](images/Install2.png)

* Extract files to a folder and run '.vsto' file:

> Using desktop folder is highly recommended - installing updates
> is permitted only from the same folder where it was installed
> for the first time!

![Navferty's Tools Ribbon Tab in MS Excel](images/Install3.png)

After installation process is completed, run (or restart) Excel application, and you will see new tab:

![Navferty's Tools Ribbon Tab in MS Excel](images/Install4.png)

[Up](#navfertys-excel-add-in)
