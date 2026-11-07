{
  "title": "Currency rates from the National Bank of Ukraine website in VBA",
  "description": "Two Excel user-defined functions that load National Bank of Ukraine exchange rates for a given date or for today, with notes on when not to use them.",
  "image": "https://navferty.net/images/articles/currency-rates-from-nbu-vba/1593904881155676818.png"
}

[← All articles](/en/articles/)

# Currency rates from the National Bank of Ukraine website in VBA

Two Excel user-defined functions that load National Bank of Ukraine exchange rates for a given date or for today, with notes on when not to use them.

A reader of the [previous article](/en/articles/currency-rates-from-cbr-vba/) about loading Bank of Russia exchange rates asked about the same task for the National Bank of Ukraine (NBU). Here is a ready-made macro for it. The process is the same as the one described in detail in that article, so this one goes straight to the result.

The macro declares two user-defined functions (UDF) that can be used in worksheet formulas.

## Rate on a given date

    =GetNbuCurrency("RUB","5/31/2020")

Loads the rate of the chosen currency for the given date. The data comes from the official NBU API:

    https://bank.gov.ua/NBU_Exchange/exchange?date=31.12.2019

## Rate for today

    =NbuTodayCurrency(C10,NOW())

Loads the rate for the current date. The value of the second argument can be anything. `NOW` is used only to force Excel to recalculate the value (and query the API) every time that argument changes: when the workbook is opened, and even when any cell in any open workbook changes.

Be careful with this function. If a workbook has many cells with such a formula, you risk flooding the NBU website with requests, which can get you banned. It is wiser to put the formula in a single place (for example on a separate sheet) and refer to that cell from the rest of your formulas.

<img src="/images/articles/currency-rates-from-nbu-vba/1593904881155676818.png" loading="lazy" width="502" height="678" alt="A worksheet with GetNbuCurrency results for several currencies and dates, and a table of today&#39;s rates" />

## Source code

The repository with the workbook, examples and source code: <a href="https://github.com/navferty/NBU-VBA-Currencies" rel="noopener" target="_blank">https://github.com/navferty/NBU-VBA-Currencies</a>. If you find bugs or have ideas for improvement, issues are welcome.

Originally published in Russian on <a href="https://pikabu.ru/story/kursyi_valyut_s_sayta_nbu_na_vba_7565687" rel="noopener" target="_blank">Pikabu</a>.
