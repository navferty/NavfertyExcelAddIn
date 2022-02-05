using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Newtonsoft.Json;

namespace Navferty.ExcelAddIn.Web.CurrencyExchangeRates.Providers.NBU
{
	public class JsonExchangeRatesForDateRecord
	{
		[JsonProperty("r030")]
		public int r030 { get; set; }

		[JsonProperty("txt")]
		public string Name { get; set; }

		[JsonProperty("rate")]
		public string RateString { get; set; }

		[JsonProperty("cc")]
		public string ISOCode { get; set; }

		[JsonProperty("exchangedate")]
		public string ValidFrom { get; set; }
	}
}
