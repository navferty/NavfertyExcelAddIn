{
  "title": "VLOOKUP in Excel: syntax, examples and common errors",
  "description": "How VLOOKUP works, a worked example with orders and users, why it returns #N/A, and what to use instead.",
  "image": "https://navferty.net/images/articles/vlookup/orders-table.jpg"
}

[← All articles](/en/articles/)

# VLOOKUP in Excel: syntax, examples and common errors

How VLOOKUP works, a worked example with orders and users, why it returns \#N/A, and what to use instead.

A typical situation: one sheet has orders with a user id, another has a list of users, and you want the user name next to each order. Copying names by hand does not scale. VLOOKUP (in Russian Excel, `ВПР`) does this: it takes a key, finds it in the first column of a table and returns a value from the same row.

## Syntax

    =VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])

- `lookup_value` – what to find, for example the user id in the current row.
- `table_array` – where to look. The key must be in the **first** column of this range.
- `col_index_num` – which column of the range to return, counting from 1.
- `range_lookup` – `FALSE` (or `0`) for an exact match, `TRUE` for an approximate one. If omitted, Excel uses `TRUE`, which is rarely what you want for ids and names.

## An example

The Orders sheet has the order id, date, amount and user id:

<img src="/images/articles/vlookup/orders-table.jpg" loading="lazy" width="536" height="365" alt="The Orders sheet with columns Order Id, Order Date, Amount and User Id" />

The Users sheet has the reference list:

<img src="/images/articles/vlookup/users-reference.jpg" loading="lazy" width="507" height="371" alt="The Users sheet with columns User Id and User Name" />

In cell `E2` of the Orders sheet start the formula. After typing the first argument, switch to the Users sheet and select the columns `A:B`; Excel fills in the second argument:

<img src="/images/articles/vlookup/vlookup-selecting-table_array.jpg" loading="lazy" width="832" height="355" alt="Selecting the table_array range A:B on the Users sheet while typing the formula" />

    =VLOOKUP(D2, Users!A:B, 2, FALSE)

<img src="/images/articles/vlookup/vlookup-formula.jpg" loading="lazy" width="841" height="366" alt="The formula being typed in E2, with the argument tooltip below the cell" />

Here the key is in `D2`, the table is `Users!A:B`, and we return the second column (the name). Copy the formula down:

<img src="/images/articles/vlookup/vlookup-result.jpg" loading="lazy" width="707" height="372" alt="VLOOKUP results in column E: user names, and #N/A for the user id that is not in the Users sheet" />

The order with user id 8 shows `#N/A`: there is no such id in the Users sheet. This is the correct behaviour. It tells you the data does not match, which is often exactly what you are looking for.

## Why you get \#N/A (or a wrong value)

- **The key is really missing.** Check the reference list.
- **Text versus number.** The id `1` as a number does not match `"1"` as text. See [converting text numbers to numeric values](/en/articles/text-numbers-to-numeric-values/).
- **Extra spaces.** `"John "` is not `"John"`. See [removing extra spaces](/en/articles/remove-extra-spaces-in-excel/).
- **The key column is not first.** VLOOKUP only looks to the right of the first column of `table_array`.
- **Approximate match.** Without `FALSE` Excel assumes the first column is sorted and can quietly return a wrong row. Always write the last argument.
- **Hard-coded column number.** If someone inserts a column in the lookup table, `2` now points at the wrong one. Check formulas after changing the table structure.
- **A moving range.** When you copy the formula, `A2:B6` shifts too. Use whole columns (`A:B`) or lock the range (`$A$2:$B$6`).

## XLOOKUP: a modern alternative to VLOOKUP

XLOOKUP is available in Excel 2021 and Microsoft 365 and removes several VLOOKUP limitations. It can search in any column and return a value from a column either to its left or right. You do not need to count a column index, exact match is the default, and an optional argument lets you choose what to display when no key is found.

### Syntax

    =XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])

The first three arguments are required: what to look for, where to look, and which range to return a result from. The lookup and return arrays must be the same size. The remaining arguments are optional and set the message for a missing key, the matching rule, and the search direction.

### Example: look up a currency name by code

On the Orders sheet, column `F` contains a currency code. On the Currencies sheet, codes are in column `C` and names are in `B`. VLOOKUP would not work with this column order: the lookup column is to the right of the return column. XLOOKUP takes these ranges separately, so it can look to the left.

<img src="/images/articles/vlookup/xlookup-orders.png" loading="lazy" width="1022" height="603" alt="Orders sheet with currency codes in column F and an empty Currency Name column for the result" />

On the Currencies sheet, column `C` contains the keys and column `B` contains the currency names:

<img src="/images/articles/vlookup/xlookup-currencies.png" loading="lazy" width="893" height="588" alt="Currency reference table with names in column B and codes in column C" />

In cell `G2`, start the formula and enter the value to look up: the code in `F2`.

<img src="/images/articles/vlookup/xlookup-lookup-value.png" loading="lazy" width="1438" height="607" alt="Beginning an XLOOKUP formula with a reference to the currency code in F2" />

For the lookup array, select the code column, `Currencies!C:C`:

<img src="/images/articles/vlookup/xlookup-lookup-array.png" loading="lazy" width="1430" height="598" alt="The XLOOKUP formula uses Currencies!C:C as the currency-code lookup array" />

For the return array, select the names in `Currencies!B:B`. This column can be to the left of the lookup array:

<img src="/images/articles/vlookup/xlookup-return-array.png" loading="lazy" width="1447" height="601" alt="The XLOOKUP formula uses Currencies!B:B as the return array, to the left of the lookup column" />

The fourth argument is optional. Enter a clear message, such as `"Currency not found"`, to display instead of `#N/A` when a code is missing:

<img src="/images/articles/vlookup/xlookup-if-not-found.png" loading="lazy" width="1466" height="597" alt="The XLOOKUP formula specifies Currency not found for a missing code" />

The complete formula for cell `G2` is:

    =XLOOKUP(F2, Currencies!C:C, Currencies!B:B, "Currency not found")

Press Enter and copy the formula down. The missing code `AMD` displays the message you specified:

<img src="/images/articles/vlookup/xlookup-result.png" loading="lazy" width="1207" height="592" alt="XLOOKUP results with currency names filled in and Currency not found shown for the missing AMD code" />

### Optional modes

If you need something other than an exact match, the fifth argument, `match_mode`, changes the lookup rule. You can usually omit it: XLOOKUP uses an exact match by default. Use `0` to explicitly request an exact match; `-1` for an exact match or the next smaller item; `1` for an exact match or the next larger item; and `2` to enable wildcard matching. In supported Microsoft 365 builds, `3` enables regular-expression matching.

<img src="/images/articles/vlookup/xlookup-match-mode.png" loading="lazy" width="1495" height="598" alt="Excel tooltip listing XLOOKUP match_mode values" />

The sixth argument, `search_mode`, sets the direction: `1` (the default) searches from first to last, while `-1` searches from last to first. Values `2` and `-2` enable binary search and require the data to be sorted in ascending or descending order respectively; leave the default for ordinary lookup lists.

<img src="/images/articles/vlookup/xlookup-search-model.png" loading="lazy" width="1557" height="597" alt="Excel tooltip listing XLOOKUP search_mode values" />

For example, the result screenshot explicitly sets exact matching and a first-to-last search: `=XLOOKUP(F2, Currencies!C:C, Currencies!B:B, "Currency not found", 0, 1)`. You can omit the last two arguments when the default exact search is all you need.

## Alternatives

**INDEX and MATCH** work in every version and also look in any direction:

    =INDEX(Users!B:B, MATCH(D2, Users!A:A, 0))

If the lookup list is a table (see [cleaning data](/en/articles/best-practices-for-cleaning-data-in-excel/)), you can refer to its columns by name, and the range grows when you add rows.

A formula like `IFERROR(VLOOKUP(...), "")` hides the errors, including real data problems. Use it only after you have checked why the values are missing.

More on formulas in general: [Beginner's guide to Excel formulas](/en/articles/beginners-guide-to-excel-formulas/).
