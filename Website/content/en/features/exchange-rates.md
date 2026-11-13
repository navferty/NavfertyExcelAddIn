{
  "title": "Exchange rates",
  "description": "Select a currency exchange rate from an online source and paste it into the current cell. Select a cell first: the rate is written into the selection.",
  "image": "https://navferty.net/images/ExchangeRates1.png"
}

[Home](/en/) › [Commands](/en/#commands) › Exchange rates

# Exchange rates

Select a currency exchange rate from an online source and paste it into the selected cell or range.

<div class="shots">

<img src="/images/ExchangeRates1.png" loading="lazy" width="796" height="693" alt="Exchange Rates dialog with the data source, currency filter and rates table" /> <img src="/images/ExchangeRates2.png" loading="lazy" width="508" height="173" alt="The selected rate pasted into cell A1" />

</div>

<div class="info-box">

**Related guides:** [Currency rates from the Bank of Russia website in VBA](/en/articles/currency-rates-from-cbr-vba/), [Currency rates from the National Bank of Ukraine website in VBA](/en/articles/currency-rates-from-nbu-vba/)

</div>

## How to use

Select a destination cell, click **Exchange rates**, choose a source and date, and filter the list by currency. Select a row and click **Select**, or double-click it, to insert its rate.

For example, choose the European Central Bank, filter for `USD`, and select the displayed US dollar rate. Available sources include the European Central Bank, the Bank of Russia and the National Bank of Ukraine.

**Important:** This writes a numeric value into the entire selected range, replacing its existing contents or formulas. The inserted value is a snapshot, not an automatically updating formula. Check the source, date and quotation units before using it in a report; network access is required.

The add-in is free and open source. [Installation instructions](/en/install/).
