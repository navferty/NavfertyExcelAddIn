{
  "title": "Undo Last Action",
  "description": "Undo the most recent supported value transformation made by the add-in.",
  "image": "https://navferty.net/images/NavfertyToolsRibbonEn.png"
}

[Home](/en/) › [Commands](/en/#commands) › Undo Last Action

# <img alt="" src="/images/icons/undo.png" class="icons" width="64" height="64" />Undo Last Action

Undo the last action performed with this add-in. Available for some functions in the 'Converting values' and 'Formatting values' sections, only if cells haven't been edited since.

## How to use

After a supported value transformation, remain on the same worksheet and click **Undo Last Action** on Navferty's Tools. For example, use it after trimming text to restore the values recorded before that action.

**Important:** This is a separate, one-action undo for supported value transformations, not a general Excel undo history. It does not undo every command, such as SQLite export or protection changes. It only restores the recorded action on the same workbook and worksheet, and only if the recorded cells still match the action's result. For a multi-cell area, that can include cells that were not changed by the transformation. It records values rather than formulas, so use a backup when formulas may be affected.

The add-in is free and open source. [Installation instructions](/en/install/).
