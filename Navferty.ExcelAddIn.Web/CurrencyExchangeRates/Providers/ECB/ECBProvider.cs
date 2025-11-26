using System;
using System.Globalization;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using System.Xml;

using Navferty.ExcelAddIn.Web.Localization;

using NLog;

#nullable enable

namespace Navferty.ExcelAddIn.Web.CurrencyExchangeRates.Providers.ECB;

internal class ECBProvider : ExchangeRatesProviderBase
{
	private const string WEB_URL = @"/service/data/EXR?startPeriod={0}&endPeriod={0}";
	private const string WEB_DATE_FORMAT = @"yyyy-MM-dd";

	private const string C_EURO_ISO = "EUR";
	private const char C_EURO = '€';

	private static readonly Lazy<CultureInfo> ci = new(() =>
	{
		CultureInfo ciNew = (CultureInfo)CultureInfo.GetCultureInfo("en-GB").Clone();
		ciNew.NumberFormat.CurrencySymbol = C_EURO.ToString();
		return ciNew;
	});

	public override string Title => UIStrings.CurrencyExchangeRates_Sources_ECB;

	private readonly ILogger logger = LogManager.GetCurrentClassLogger();
	public override ILogger Logger => logger;

	private readonly HttpClient web = new()
	{
		BaseAddress = new Uri("https://data-api.ecb.europa.eu"),
	};

	protected override async Task<ExchangeRateRecord[]> DownloadExchangeRatesForDayAsync(DateTime dt)
	{
		string url = string.Format(WEB_URL, dt.ToString(WEB_DATE_FORMAT));
		logger.Debug($"Query url: {url}");

		var rawXML = string.Empty;

		try
		{
			var webResponce = await web.GetAsync(url);
			logger.Debug($"webResponce.StatusCode: {webResponce.StatusCode}, IsSuccessStatusCode = {webResponce.IsSuccessStatusCode}");
			webResponce.EnsureSuccessStatusCode();
			rawXML = await webResponce.Content.ReadAsStringAsync();
		}
		catch (Exception ex)
		{
			logger.Error(ex, "Failed to download xml from web service!");
			throw;
		}

		if (string.IsNullOrWhiteSpace(rawXML))
		{
			logger.Error("Web service answer xml = NULL!");
			throw new Exception(UIStrings.CurrencyExchangeRates_Error_Network);
		}

		ECB.ECBExchangeRatesRecord[] rawRows;
		try
		{
			rawRows = ParseECBXml(rawXML);
		}
		catch (Exception ex)
		{
			logger.Error(ex, $"Failed to parse xml:\nrawXML:\n{rawXML}");
			throw new Exception(UIStrings.CurrencyExchangeRates_Error_ParseError);
		}

		try
		{
			return rawRows.Select(row => new ExchangeRateRecord(row)).ToArray();
		}
		catch (Exception ex)
		{
			logger.Error(ex, $"Failed to convert '{rawRows.GetType()}' to '{typeof(ExchangeRateRecord)}'!");

			throw new Exception(UIStrings.CurrencyExchangeRates_Error_ParseError);

		}
	}

	private ECB.ECBExchangeRatesRecord[] ParseECBXml(string xmlText)
	{
		var doc = new XmlDocument();
		doc.LoadXml(xmlText);

		//Get all xml nodes with <generic:Series> tag.
		//This tag is about some asset exchange rate. Not only money!
		var seriesTags = doc.GetElementsByTagName("generic:Series").Cast<XmlElement>().ToArray();
		var rawRows = seriesTags
			.Select(tagSeries => new ECB.ECBExchangeRatesRecord(tagSeries))
			.Where(x => x.CurrencyDenom == C_EURO_ISO)//Select only rates related to Euro ('EUR'). Full list contains many other asset exchange rates.
			.ToArray();

		return rawRows;
	}

	public override CultureInfo Culture => ci.Value;
}
