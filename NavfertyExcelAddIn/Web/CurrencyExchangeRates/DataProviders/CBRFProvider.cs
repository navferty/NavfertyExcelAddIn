using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using NavfertyExcelAddIn.Localization;

namespace NavfertyExcelAddIn.Web.CurrencyExchangeRates.Providers
{
	internal class CBRFProvider : IExchangeRatesDataProvider
	{
		private static readonly CultureInfo ciRU = CultureInfo.GetCultureInfo("ru-RU");

		public string GetTitle() => UIStrings.CurrencyExchangeRates_Sources_CBRF;

		public async Task<WebResultRow[]> GetExchabgeRatesForDate(DateTime dt)
		{
			var exchangeRatesRows = await CurrencyExchangeRates.GetCurrencyExchabgeRates_CBRF(dt);
			return exchangeRatesRows;
		}

		public CultureInfo GetCulture() => ciRU;

		public override string ToString() => GetTitle();
	}
}
