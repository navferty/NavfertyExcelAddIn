{
  "title": "Removing extra spaces in Excel",
  "description": "TRIM, non-breaking spaces and line breaks: how to clean text cells properly.",
  "image": "https://navferty.net/images/articles/remove-extra-spaces-in-excel/extra-spaces.jpg"
}

[← All articles](/en/articles/)

# Removing extra spaces in Excel

TRIM, non-breaking spaces and line breaks: how to clean text cells properly.

Extra spaces are a classic reason for failed lookups, duplicates that are not detected and imports that reject a file. `"Smith"` and `"Smith "` look the same, but Excel treats them as different values.

## TRIM

`=TRIM(A2)` removes spaces at the start and the end of the text and reduces repeated spaces between words to one. Copy the formulas, then paste them over the original column with **Paste Special → Values**.

## When TRIM does nothing

If the value still looks wrong, it probably contains a character that is not an ordinary space:

- **Non-breaking space**, character code 160. It appears in text copied from web pages and some systems. `TRIM` ignores it. Use `=TRIM(SUBSTITUTE(A2,CHAR(160)," "))`.
- **Line breaks**, codes 10 and 13. `=CLEAN(A2)` removes them (and other non-printable characters up to code 31). If you want a space instead of the break, use `SUBSTITUTE(A2,CHAR(10)," ")` first.
- **Tabs**, code 9. `CLEAN` removes them as well.

To see which character is the culprit, use `=UNICODE(RIGHT(A2,1))` or `=UNICODE(LEFT(A2,1))`. A result of 160 means a non-breaking space.

## A combined formula

    =TRIM(CLEAN(SUBSTITUTE(A2,CHAR(160)," ")))

This covers most cases in real data.

## Without formulas

**Find & Replace** (<kbd>Ctrl</kbd>+<kbd>H</kbd>) can remove all spaces from a selection: put a space into **Find what**, leave **Replace with** empty. It is fine for codes and numbers like `1 234 567`, but it will also glue words together, so do not use it on names or addresses.

To enter a non-breaking space in the Find box, hold <kbd>Alt</kbd> and type `0160` on the numeric keypad.

## Check before and after

Compare `=LEN(A2)` with `=LEN(TRIM(A2))` to count how many cells actually have the problem:

<img src="/images/articles/remove-extra-spaces-in-excel/extra-spaces.jpg" loading="lazy" width="1002" height="251" alt="A table comparing text length, the code of the first character and the result of TRIM, including a non-breaking space that TRIM leaves in place" />

    =SUMPRODUCT(--(LEN(A2:A1000)<>LEN(TRIM(A2:A1000))))

## With the add-in

**Trim Spaces** works directly on the selected cells: it removes leading and trailing whitespace, line breaks and extra spaces without helper columns or pasting values back.

The add-in command for this: [Trim Spaces](/en/features/trim-spaces/).
