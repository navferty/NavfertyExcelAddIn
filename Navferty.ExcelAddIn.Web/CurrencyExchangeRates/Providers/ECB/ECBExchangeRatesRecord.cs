using System.Collections.Generic;
using System.Linq;
using System.Xml;

#nullable enable

namespace Navferty.ExcelAddIn.Web.CurrencyExchangeRates.Providers.ECB;

internal class ECBExchangeRatesRecord
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

	internal ECBExchangeRatesRecord(XmlElement tagSeries)
	{
		/* The 'tagSeries' xml block looks like:
			<generic:Series>
				<generic:SeriesKey>
					<generic:Value id="FREQ" value="D"/>
					<generic:Value id="CURRENCY" value="CAD"/>
					<generic:Value id="CURRENCY_DENOM" value="EUR"/>
					<generic:Value id="EXR_TYPE" value="SP00"/>
					<generic:Value id="EXR_SUFFIX" value="A"/>
				</generic:SeriesKey>
				<generic:Attributes>
					<generic:Value id="UNIT" value="CAD"/>
					<generic:Value id="TIME_FORMAT" value="P1D"/>
					<generic:Value id="COLLECTION" value="A"/>
					<generic:Value id="TITLE" value="Canadian dollar/Euro"/>
					<generic:Value id="SOURCE_AGENCY" value="DE2"/>
					<generic:Value id="UNIT_MULT" value="0"/>
					<generic:Value id="DECIMALS" value="4"/>
					<generic:Value id="TITLE_COMPL" value="ECB reference exchange rate, Canadian dollar/Euro, 2:15 pm (C.E.T.)"/>
				</generic:Attributes>
				<generic:Obs>
					<generic:ObsDimension value="2022-02-01"/>
					<generic:ObsValue value="1.4299"/>
					<generic:Attributes>
					<generic:Value id="OBS_STATUS" value="A"/>
					<generic:Value id="OBS_CONF" value="F"/>
					</generic:Attributes>
				</generic:Obs>
				<generic:Obs>
					<generic:ObsDimension value="2022-02-02"/>
					<generic:ObsValue value="1.433"/>
					<generic:Attributes>
						<generic:Value id="OBS_STATUS" value="A"/>
						<generic:Value id="OBS_CONF" value="F"/>
					</generic:Attributes>
				</generic:Obs>
			</generic:Series>			 
		 */

		var tagSeriesKey = tagSeries.GetElementsByTagName("generic:SeriesKey").Cast<XmlElement>().FirstOrDefault();
		var tagAttributes = tagSeries.GetElementsByTagName("generic:Attributes").Cast<XmlElement>().FirstOrDefault();
		var tagObs = tagSeries.GetElementsByTagName("generic:Obs").Cast<XmlElement>().FirstOrDefault();

		var dicSeriesKey = ChildNodesAsAttributes(tagSeriesKey);
		var dicAttributes = ChildNodesAsAttributes(tagAttributes);
		var dicObs = ChildNodesAsAttributes(tagObs);

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

	/// <summary>
	/// Return xml node children nodes like 'generic:Value id="FREQ" value="D"' as a dictionary, where key=1-st node attribute, and value=2-nd node attribute.
	/// In case of attributes count < 2 - dictionary key=node.LocalName, and value=1-st attribute.
	/// </summary>
	/// <returns>Dictionary of ID-Value pairs built from attributes of child nodes
	/// If the xml structure will be changed, we will fall :)
	/// </returns>
	private static Dictionary<string, string> ChildNodesAsAttributes(XmlNode node)
	{
		/* SAMPLE of XML nodes to parse:
		 
		1)
			<generic:SeriesKey>
				<generic:Value id="FREQ" value="D"/>
				<generic:Value id="CURRENCY" value="CAD"/>
				<generic:Value id="CURRENCY_DENOM" value="EUR"/>
				<generic:Value id="EXR_TYPE" value="SP00"/>
				<generic:Value id="EXR_SUFFIX" value="A"/>
			</generic:SeriesKey>

			in this case will return:
			FREQ/D
			CURRENCY/CAD
			CURRENCY_DENOM/EUR
			EXR_TYPE/SP00
			EXR_SUFFIX/A

		*********************************************
		2) 
			<generic:Obs>
				<generic:ObsDimension value="2022-02-02"/>
				<generic:ObsValue value="1.433"/>
				<generic:Attributes>
					<generic:Value id="OBS_STATUS" value="A"/>
					<generic:Value id="OBS_CONF" value="F"/>
				</generic:Attributes>
			</generic:Obs>

			in this case will return:
				ObsDimension/"2022-02-02"
				ObsValue/"1.433"
		 */

		var childNodes = node.ChildNodes.Cast<XmlElement>().ToArray();
		var dic = childNodes
		.Where(nodeChild => (nodeChild.Attributes.Count == 1 || nodeChild.Attributes.Count == 2))
		.Select(nodeChild =>
		{
			var nodeAttributes = nodeChild.Attributes.Cast<XmlAttribute>().ToArray();
			string Key = nodeChild.LocalName;
			string Vaue = nodeAttributes[0].Value;

			if (nodeAttributes.Length == 2)
			{
				Key = nodeAttributes[0].Value;
				Vaue = nodeAttributes[1].Value;
			}
			return new { Key, Vaue };
		}).ToDictionary(x => x.Key, x => x.Vaue);

		return dic;
	}
}
