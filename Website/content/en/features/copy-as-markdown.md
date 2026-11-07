{
  "title": "Copy as Markdown",
  "description": "Copy selected cells to clipboard as a markdown-formatted table.",
  "image": "https://navferty.net/images/ExportToMarkdown.png"
}

[Home](/en/) › [Commands](/en/#commands) › Copy as Markdown

# <img alt="" src="/images/icons/markdown.png" class="icons" width="64" height="64" />Copy as Markdown

Copy selected cells to clipboard as a markdown-formatted table.

<div class="shots">

<img src="/images/ExportToMarkdown.png" loading="lazy" width="858" height="306" alt="Markdown export example" />

</div>

## How to use

Select a rectangular range, including its header row, and click **Markdown**. Paste the clipboard into a Markdown editor or GitHub comment. For a two-column selection the output looks like:

```text
|Name|Amount|
|---|---|
|Apples|12|
```

The first selected row is always treated as the header. The command copies cell values, not formulas or formatting, and collapses whitespace in cell text. Literal `|` characters are not escaped automatically, so review the pasted table.

The add-in is free and open source. [Installation instructions](/en/install/).
