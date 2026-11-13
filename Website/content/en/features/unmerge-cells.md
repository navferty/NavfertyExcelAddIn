{
  "title": "Unmerge Cells",
  "description": "Unmerge cells while filling each cell with the original merged value.",
  "image": "https://navferty.net/images/Unmerge1.png"
}

[Home](/en/) › [Commands](/en/#commands) › Unmerge Cells

# <img alt="" src="/images/icons/unmergeCells.png" class="icons" width="64" height="64" />Unmerge Cells

Unmerge cells while filling each cell with the original merged value.

<div class="shots">

<img src="/images/Unmerge1.png" loading="lazy" width="550" height="199" alt="Unmerge Cells - Before" /> <img src="/images/Unmerge2.png" loading="lazy" width="550" height="199" alt="Unmerge Cells - After" />

</div>

## How to use

Select cells that include the merged areas and click **Unmerge**. For example, if `A1:A3` is merged and contains `North`, the command unmerges it and fills `A1`, `A2` and `A3` with `North`.

The entire merged area is processed even when the selection includes only part of it. The original top-left cell's formula is assigned to every cell in the area, so review formula results as well as values. Save a copy before making changes.

The add-in is free and open source. [Installation instructions](/en/install/).
