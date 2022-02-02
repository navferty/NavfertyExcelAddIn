using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace NavfertyExcelAddIn.Web.CurrencyExchangeRates.Providers.ECB
{
	public class ECBExchangeRatesRecord
	{
		public readonly string ISO;
		public readonly string CurrencyDenom;

		public readonly string Title;
		public readonly string Description;
		public readonly string Unit;
		public readonly string UnitMult;
		public readonly string Decimals;

		public readonly string ObsDimension;
		public readonly string ObsValue;

		internal ECBExchangeRatesRecord(XmlElement nodeSeries)
		{
			var seriesKey = nodeSeries.GetElementsByTagName("generic:SeriesKey").Cast<XmlElement>().FirstOrDefault();
			var attrKey = nodeSeries.GetElementsByTagName("generic:Attributes").Cast<XmlElement>().FirstOrDefault();
			var obsKey = nodeSeries.GetElementsByTagName("generic:Obs").Cast<XmlElement>().FirstOrDefault();

			var dicSeriesKey = ChildNodesAsAttributes(seriesKey);
			var dicAttributes = ChildNodesAsAttributes(attrKey);
			var dicObs = ChildNodesAsAttributes(obsKey);

			dicSeriesKey.TryGetValue("CURRENCY", out ISO);
			dicSeriesKey.TryGetValue("CURRENCY_DENOM", out CurrencyDenom);

			dicAttributes.TryGetValue("TITLE", out Title);
			dicAttributes.TryGetValue("TITLE_COMPL", out Description);
			dicAttributes.TryGetValue("UNIT", out Unit);
			dicAttributes.TryGetValue("UNIT_MULT", out UnitMult);
			dicAttributes.TryGetValue("DECIMALS", out Decimals);

			dicObs.TryGetValue("ObsDimension", out ObsDimension);
			dicObs.TryGetValue("ObsValue", out ObsValue);
		}


		private static Dictionary<string, string> ChildNodesAsAttributes(XmlNode node)
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
		}
	}
}
