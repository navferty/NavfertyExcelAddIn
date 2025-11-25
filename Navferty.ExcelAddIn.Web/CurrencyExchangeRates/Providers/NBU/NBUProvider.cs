using System;
using System.Globalization;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;

using Navferty.ExcelAddIn.Web.Localization;

using Newtonsoft.Json;

using NLog;

#nullable enable

namespace Navferty.ExcelAddIn.Web.CurrencyExchangeRates.Providers.NBU;

internal class NBUProvider : ExchangeRatesProviderBase
{

	//https://bank.gov.ua/ua/open-data/api-dev
	//https://bank.gov.ua/NBUStatService/v1/statdirectory/exchange
	//https://bank.gov.ua/NBUStatService/v1/statdirectory/exchange?date=20200302&json
	private const string WEB_URL = @"https://bank.gov.ua/NBUStatService/v1/statdirectory/exchange?date={0}&json";
	private const string WEB_DATE_FORMAT = @"yyyyMMdd";

	private static readonly CultureInfo ci = CultureInfo.GetCultureInfo("uk-UA");

	public override string Title => UIStrings.CurrencyExchangeRates_Sources_NBU;


	private readonly ILogger logger = LogManager.GetCurrentClassLogger();
	public override ILogger Logger => logger;

	private string rawJson = String.Empty;
	private NBU.JsonExchangeRatesForDateRecord[] rawJsonRows = Array.Empty<NBU.JsonExchangeRatesForDateRecord>();
	private HttpClient web = new();

	protected override async Task<ExchangeRateRecord[]> DownloadExchangeRatesForDayAsync(DateTime dt)
	{
		rawJson = String.Empty;
		rawJsonRows = Array.Empty<NBU.JsonExchangeRatesForDateRecord>();

		string url = string.Format(WEB_URL, dt.ToString(WEB_DATE_FORMAT));
		logger.Debug($"Query url: {url}");

		var webResponce = (await web.GetAsync(url));
		logger.Debug($"webResponce.StatusCode: {webResponce.StatusCode}, IsSuccessStatusCode = {webResponce.IsSuccessStatusCode}");
		rawJson = await webResponce.EnsureSuccessStatusCode().Content.ReadAsStringAsync();

		try
		{
			rawJsonRows = JsonConvert.DeserializeObject<NBU.JsonExchangeRatesForDateRecord[]>(rawJson)!;
		}
		catch (Exception ex)
		{
			logger.Error(ex, $"Failed to deserialize NBU Json via 'JsonConvert.DeserializeObject<NBU.JsonExchangeRatesForDateRecord[]>(rawJson)':\nrawJson:\n{rawJson}");
			throw new Exception(UIStrings.CurrencyExchangeRates_Error_ParseError);
		}

		try
		{
			var rows = rawJsonRows.Select(row => new ExchangeRateRecord(row)).ToArray();
			return rows;
		}
		catch (Exception ex)
		{
			logger.Error(ex, $"Failed to convert '{rawJsonRows.GetType()}' to '{typeof(ExchangeRateRecord)}'!");
			throw new Exception(UIStrings.CurrencyExchangeRates_Error_ParseError);
		}
	}

	public override CultureInfo Culture => ci;
}
