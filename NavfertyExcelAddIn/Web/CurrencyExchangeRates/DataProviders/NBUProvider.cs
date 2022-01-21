using System;
using System.Globalization;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;

using NavfertyExcelAddIn.Localization;

using Newtonsoft.Json;

namespace NavfertyExcelAddIn.Web.CurrencyExchangeRates.Providers
{
	internal class NBUProvider : ExchangeRatesDataProviderBaase
	{
		private static readonly CultureInfo ci = CultureInfo.GetCultureInfo("uk-UA");

		public override string Title => UIStrings.CurrencyExchangeRates_Sources_NBU;

		protected override async Task<WebResultRow[]> GetExchabgeRatesForDate_Core(DateTime dt)
		{
			//https://bank.gov.ua/ua/open-data/api-dev
			//https://bank.gov.ua/NBUStatService/v1/statdirectory/exchange
			//https://bank.gov.ua/NBUStatService/v1/statdirectory/exchange?date=20200302&json

			string sDateForNBU = dt.ToString("yyyyMMdd");
			var urlNBUExchangeForDate = @$"https://bank.gov.ua/NBUStatService/v1/statdirectory/exchange?date={sDateForNBU}&json";
			//Debug.WriteLine(urlNBUExchangeForDate);

			using (var htc = new HttpClient())
			{
				var sJson = await (await htc.GetAsync(urlNBUExchangeForDate)).
					EnsureSuccessStatusCode().
					Content.ReadAsStringAsync();

				var nbuResultRows = JsonConvert.DeserializeObject<NBU.ExchangeRatesForDateRecord[]>(sJson);
				var rows = nbuResultRows.Select(row => new WebResultRow(row)).ToArray();
				return rows;
			}
		}

		public override CultureInfo Culture => ci;
	}
}
