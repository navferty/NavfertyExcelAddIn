{
  "title": "Validate XML with XSD",
  "description": "Validate an XML file against XSD schemas and review reported issues in Excel.",
  "image": "https://navferty.net/images/ValidateXml1.png"
}

[Home](/en/) › [Commands](/en/#commands) › Validate XML with XSD

# <img alt="" src="/images/icons/validateXml.png" class="icons" width="64" height="64" />Validate XML with XSD

Validate an XML file against one or more XSD schemas. If issues are reported, a detailed report is created in a new workbook.

<div class="shots">

<img src="/images/ValidateXml1.png" loading="lazy" width="883" height="257" alt="Validation report on a sheet: severity, element, value and message" />

</div>

## View sample error report

<div class="tbl">

| Severity | Element  | Message                                                                                                                                    |
|----------|----------|--------------------------------------------------------------------------------------------------------------------------------------------|
| Error    | city     | The element 'shipto' has invalid child element 'city'. List of possible elements expected: 'address'.                                      |
| Error    | quantity | The 'quantity' element is invalid - The value '-5' is invalid according to its datatype 'http://www.w3.org/2001/XMLSchema:positiveInteger' |
| Error    | price    | The 'price' element is invalid - The value 'asdasd' is invalid according to its datatype 'http://www.w3.org/2001/XMLSchema:decimal'        |

</div>

<div class="info-box">

**Related guide:** [Validating XML against XSD inside Excel](/en/articles/validate-xml-against-xsd-in-excel/)

</div>

## How to use

Click **Validate XML with XSD**, select one XML file, then select the XSD file or files that make up the schema. If validation reports issues, the new workbook lists severity, element, value and message. If there are no reported issues, a success message appears instead.

For example, a text value where an integer is required produces a report identifying the element and explaining the type mismatch. Fix the XML and run validation again.

Validation checks the selected file against the supplied schemas; it does not repair it or verify business meaning. Missing, malformed or incompatible files can prevent validation.

The add-in is free and open source. [Installation instructions](/en/install/).
