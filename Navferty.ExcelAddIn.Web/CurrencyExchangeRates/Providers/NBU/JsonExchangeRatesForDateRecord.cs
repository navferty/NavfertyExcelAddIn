
using Newtonsoft.Json;

#nullable enable

namespace Navferty.ExcelAddIn.Web.CurrencyExchangeRates.Providers.NBU
{
	public class JsonExchangeRatesForDateRecord
	{
		[JsonProperty("r030")]
		public int R030 { get; set; } = 0;

		[JsonProperty("txt")]
		public string Name { get; set; } = string.Empty;

		[JsonProperty("rate")]
		public string RateString { get; set; } = string.Empty;

		[JsonProperty("cc")]
		public string ISOCode { get; set; } = string.Empty;

		[JsonProperty("exchangedate")]
		public string ValidFrom { get; set; } = string.Empty;
	}
}
