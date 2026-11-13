{
  "title": "Conditional Formatting",
  "description": "Fill formatting across selected rows using conditional formatting in a source row.",
  "image": "https://navferty.net/images/ConditionalFormatting1.png"
}

[Home](/en/) › [Commands](/en/#commands) › Conditional Formatting

# Conditional Formatting

Select the target range with the source conditional formatting rules in its first row and at least one additional row. The command copies formatting from the source row, replacing ordinary cell formatting, including number formats, as well as conditional formatting, so work on a copy. Selecting the entire worksheet uses its used range; selecting a single cell does nothing.

<div class="shots">

<img src="/images/ConditionalFormatting1.png" loading="lazy" width="718" height="646" alt="Conditional formatting set on the first row only" /> <img src="/images/ConditionalFormatting2.png" loading="lazy" width="716" height="646" alt="The same format applied to the whole table" />

</div>

## Example and important notes

If row 2 contains the correct rules and rows 3–10 have inconsistent formatting, select the cells in rows 2–10 and run **Conditional Formatting**. The first row is the source; if it has no rules, the second row is used instead.

The command deletes existing conditional-format rules in the destination and pastes **all formatting**, not just rules. Number formats, fonts, fills and borders can therefore change. It requires at least two selected rows and does nothing for a single cell. Save a copy and verify the resulting rules and cell formats.

The add-in is free and open source. [Installation instructions](/en/install/).
