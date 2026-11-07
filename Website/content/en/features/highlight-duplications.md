{
  "title": "Highlight Duplications",
  "description": "Automatically color-code duplicate values in the selected range. Different colors indicate different duplicate groups.",
  "image": "https://navferty.net/images/Duplicates.png"
}

[Home](/en/) › [Commands](/en/#commands) › Highlight Duplications

# <img alt="" src="/images/icons/highlightDuplicates.png" class="icons" width="64" height="64" />Highlight Duplications

Automatically color-code duplicate values in the selected range. Different colors indicate different duplicate groups.

<div class="shots">

<img src="/images/Duplicates.png" loading="lazy" width="858" height="306" alt="Highlighted duplicates with different colors" />

</div>

<div class="info-box">

**Related guide:** [Highlighting duplicate values without formulas](/en/articles/highlight-duplicate-values-in-excel/)

</div>

## How to use

Select the cells to compare, then click **Colorize Duplicates**. For `apple`, `pear`, `apple`, both `apple` cells receive the same fill; a different repeated value receives another color.

Comparison ignores letter case and skips empty or whitespace-only values. It compares individual cell values across the entire selection, not complete rows. Colors are drawn from a finite palette and can repeat when there are many groups.

**Important:** Existing cell fills are changed for duplicates. This is a one-time formatting operation, not a live rule; rerun it after the data changes and review existing colors.

The add-in is free and open source. [Installation instructions](/en/install/).
