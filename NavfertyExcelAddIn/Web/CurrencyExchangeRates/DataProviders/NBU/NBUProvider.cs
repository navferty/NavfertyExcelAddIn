using System;
using System.Globalization;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;

using NavfertyExcelAddIn.Localization;

using Newtonsoft.Json;

using NLog;

namespace NavfertyExcelAddIn.Web.CurrencyExchangeRates.Providers
{
	internal class NBUProvider : ExchangeRatesDataProviderBaase
	{
		private static readonly CultureInfo ci = CultureInfo.GetCultureInfo("uk-UA");

		public override string Title => UIStrings.CurrencyExchangeRates_Sources_NBU;


		private readonly ILogger logger = LogManager.GetCurrentClassLogger();
		public override ILogger Logger => logger;

		private string rawJson = String.Empty;
		private NBU.JsonExchangeRatesForDateRecord[] rawJsonRows = Array.Empty<NBU.JsonExchangeRatesForDateRecord>();
		private HttpClient web = new();

		protected override async Task<WebResultRow[]> DownloadWebResultRowsForDate(DateTime dt)
		{
			rawJson = String.Empty;
			rawJsonRows = Array.Empty<NBU.JsonExchangeRatesForDateRecord>();

			//https://bank.gov.ua/ua/open-data/api-dev
			//https://bank.gov.ua/NBUStatService/v1/statdirectory/exchange
			//https://bank.gov.ua/NBUStatService/v1/statdirectory/exchange?date=20200302&json

			string sDateForNBU = dt.ToString("yyyyMMdd");
			var urlNBUExchangeForDate = @$"https://bank.gov.ua/NBUStatService/v1/statdirectory/exchange?date={sDateForNBU}&json";
			logger.Debug($"NBU: ExchangeRates query url: {urlNBUExchangeForDate}");

			rawJson = await (await web.GetAsync(urlNBUExchangeForDate)).
				EnsureSuccessStatusCode().
				Content.ReadAsStringAsync();

			try
			{
				rawJsonRows = JsonConvert.DeserializeObject<NBU.JsonExchangeRatesForDateRecord[]>(rawJson);
			}
			catch (Exception ex)
			{
				logger.Error(ex, $"NBU: Failed to deserialize Json via JsonConvert.DeserializeObject<NBU.JsonExchangeRatesForDateRecord[]>(rawJson):\n\n{rawJson}");
				throw new Exception(UIStrings.CurrencyExchangeRates_Error_ParseError);
			}

			try
			{
				var rows = rawJsonRows.Select(row => new WebResultRow(row)).ToArray();
				return rows;
			}
			catch (Exception ex)
			{
				logger.Error(ex, $"NBU: Failed to convert 'JsonExchangeRatesForDateRecord' to 'WebResultRow'!");
				throw new Exception(UIStrings.CurrencyExchangeRates_Error_ParseError);
			}
		}

		public override CultureInfo Culture => ci;
	}
}
