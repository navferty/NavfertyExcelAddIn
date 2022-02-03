using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using System.Xml;

using NavfertyExcelAddIn.Localization;

using Newtonsoft.Json;

namespace NavfertyExcelAddIn.Web.CurrencyExchangeRates.Providers
{
	internal class ECBProvider : ExchangeRatesDataProviderBaase
	{
		private const string C_EURO_ISO = "EUR";
		private const char C_EURO = '€';

		private static readonly Lazy<CultureInfo> ci = new Lazy<CultureInfo>(() =>
		{
			CultureInfo ciNew = (CultureInfo)CultureInfo.GetCultureInfo("en-GB").Clone();
			ciNew.NumberFormat.CurrencySymbol = C_EURO.ToString();
			return ciNew;
		});

		public override string Title => UIStrings.CurrencyExchangeRates_Sources_ECB;


		private string rawXMLString = String.Empty;
		private NBU.JsonExchangeRatesForDateRecord[] rawJsonRows;

		protected override async Task<WebResultRow[]> DownloadWebResultRowsForDate(DateTime dt)
		{
			//https://sdw-wsrest.ecb.europa.eu/service/data/EXR?startPeriod=2022-02-01&endPeriod=2022-02-01
			string sDate = dt.ToString("yyyy-MM-dd");
			var urlECBExchangeForDate = @$"https://sdw-wsrest.ecb.europa.eu/service/data/EXR?startPeriod={sDate}&endPeriod={sDate}";
			Debug.WriteLine(urlECBExchangeForDate);

			using (var htc = new HttpClient())
			{
				var rawECBString = await (await htc.GetAsync(urlECBExchangeForDate)).
					EnsureSuccessStatusCode().
					Content.ReadAsStringAsync();

				if (string.IsNullOrWhiteSpace(rawECBString)) throw new Exception("Нет данных на указанную дату");


				var rows = ParseECBXml(rawECBString);
				return rows.Select(row => new WebResultRow(row)).ToArray();
			}
		}

		private ECB.ECBExchangeRatesRecord[] ParseECBXml(string xmlText)
		{
			var doc = new XmlDocument();
			doc.LoadXml(xmlText);

			var ChildNodesAsAttributes = new Func<XmlNode, Dictionary<string, string>>(node =>
			{
				var attrsList = node.ChildNodes.Cast<XmlElement>().ToArray();
				var dic = attrsList
				.Where(attr => (attr.Attributes.Count == 1 || attr.Attributes.Count == 2))
				.Select(attr =>
				{
					var nodeAttributes = attr.Attributes.Cast<XmlAttribute>().ToArray();
					string Key = attr.LocalName;
					string Vaue = nodeAttributes[0].Value;

					if (nodeAttributes.Length == 2)
					{
						Key = nodeAttributes[0].Value;
						Vaue = nodeAttributes[1].Value;
					}
					return new { Key, Vaue };
				}).ToArray().ToDictionary(x => x.Key, x => x.Vaue);

				return dic;
			});

			var xSeriesKeys = doc.GetElementsByTagName("generic:Series").Cast<XmlElement>().ToArray();
			var rawRows = xSeriesKeys
				.Select(nodeSeries => new ECB.ECBExchangeRatesRecord(nodeSeries))
				.Where(x => x.CurrencyDenom == C_EURO_ISO)
				.ToArray();

			return rawRows;
		}

		public override CultureInfo Culture => ci.Value;
	}
}
