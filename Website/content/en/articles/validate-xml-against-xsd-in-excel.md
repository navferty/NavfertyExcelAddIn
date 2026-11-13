{
  "title": "Validating XML against XSD inside Excel",
  "description": "How to check an XML file against its schema and read the errors as a table.",
  "image": "https://navferty.net/images/NavfertyToolsRibbonEn.png"
}

[← All articles](/en/articles/)

# Validating XML against XSD inside Excel

How to check an XML file against its schema and read the errors as a table.

Many systems exchange data as XML: tax and customs declarations, bank statements, e-invoicing, product feeds. Often the XML is produced from an Excel table by a script or a conversion tool, and the receiving system rejects it with a message like “file does not conform to schema” without saying where the mistake is.

## What an XSD is

An XSD (XML Schema) file describes which elements are allowed, in what order, which are required and what type each value has. Validating means checking an XML file against those rules. A file can be well-formed (correct tags) and still invalid (a required element is missing, or a text value sits in a date field).

## Typical problems

- A required element is missing or in the wrong order. Order matters in `xs:sequence`.
- A value does not match its type: `-5` for a `positiveInteger`, `12.03.2024` for an `xs:date` that expects `2024-03-12`.
- Decimal comma instead of a point.
- Strings longer than `maxLength` or not matching a `pattern` (for example a tax ID of the wrong length).
- Characters like `<` and `&` that were not escaped when the XML was written.

## Other ways to validate

- `xmllint --noout --schema schema.xsd file.xml` on Linux or macOS, and available for Windows too.
- Editors with schema support, such as Visual Studio Code with an XML extension, Oxygen or Notepad++ with the XML Tools plugin.
- Excel's own XML Maps (**Developer → Source**) can import an XSD and map it to cells. It is limited with nested or repeating structures.

## With the add-in

**Validate XML with XSD** takes an XML file and an XSD file and writes a new workbook with one row per problem: severity, element and the message. Because it is a normal sheet, you can filter, sort and send it to the person who produced the data.

When you only have the schema and need to understand what a valid file looks like, **Create Sample XML based on XSD** generates an example file from it. For checks before the XML is even created, **Validate Values** checks that the cells hold numbers, dates, valid tax IDs, or text without characters that break XML.

## Reading the messages

Messages mention the element name and the expected alternatives. Fix the first error first: later messages are often consequences of the first one, for example a missing element makes everything after it look out of order.

The add-in command for this: [Create Sample XML based on XSD](/en/features/create-sample-xml/), [Validate XML with XSD](/en/features/validate-xml/).
