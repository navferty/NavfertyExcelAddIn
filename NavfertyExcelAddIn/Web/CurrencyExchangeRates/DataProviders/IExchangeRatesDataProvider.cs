using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NavfertyExcelAddIn.Web.CurrencyExchangeRates.Providers
{
	internal interface IExchangeRatesDataProvider
	{
		public string GetTitle();

		public Task<WebResultRow[]> GetExchabgeRatesForDate(DateTime dt);

		public CultureInfo GetCulture();
	}
}
