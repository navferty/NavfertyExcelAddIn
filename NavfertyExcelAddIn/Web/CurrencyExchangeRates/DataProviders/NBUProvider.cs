using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using NavfertyExcelAddIn.Localization;

namespace NavfertyExcelAddIn.Web.CurrencyExchangeRates.Providers
{
	internal class NBUProvider : IExchangeRatesDataProvider
	{
		private static readonly CultureInfo ciUA = CultureInfo.GetCultureInfo("uk-UA");

		public string GetTitle() => UIStrings.CurrencyExchangeRates_Sources_NBU;

		public async Task<WebResultRow[]> GetExchabgeRatesForDate(DateTime dt)
		{
			var exchangeRatesRows = await CurrencyExchangeRates.GetCurrencyExchabgeRates_NBU(dt);
			return exchangeRatesRows;
		}

		public CultureInfo GetCulture() => ciUA;

		public override string ToString() => GetTitle();
	}
}
