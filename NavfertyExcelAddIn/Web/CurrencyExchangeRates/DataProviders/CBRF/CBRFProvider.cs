using System;
using System.Data;
using System.Diagnostics;
using System.Dynamic;
using System.Globalization;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using System.Xml.Linq;

using NavfertyExcelAddIn.Commons;
using NavfertyExcelAddIn.Localization;

namespace NavfertyExcelAddIn.Web.CurrencyExchangeRates.Providers
{
	internal class CBRFProvider : ExchangeRatesDataProviderBaase
	{
		private static readonly CultureInfo ci = CultureInfo.GetCultureInfo("ru-RU");

		public override string Title => UIStrings.CurrencyExchangeRates_Sources_CBRF;

		public override CultureInfo Culture => ci;

		private DataTable rawDataTable = null;
		protected override async Task<WebResultRow[]> DownloadWebResultRowsForDate(DateTime dt)
		{
			await TestECB(dt);

			using (var cbr = new Web.CBR.DailyInfoSoapClient())
			{
				var dtsResult = await cbr.GetCursOnDateAsync(dt);
				if (dtsResult == null) throw new Exception("Failed to get remote data with no errors!");

				rawDataTable = dtsResult.Tables.Cast<DataTable>().FirstOrDefault();
				if (rawDataTable == default) throw new Exception("Remote dstaset does not containt Tables!");
				var rows = rawDataTable.RowsAsEnumerable().Select(row => new WebResultRow(row, dt)).ToArray();
				return rows;
			};
		}

		private async Task TestECB(DateTime dt)
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

				if (string.IsNullOrWhiteSpace(rawECBString))
					return;

				// Load an XML document.

				var xr = new System.IO.StringReader(rawECBString);
				var xDoc = XDocument.Load(xr);

				dynamic root = new ExpandoObject();
				NavfertyExcelAddIn.Web.CurrencyExchangeRates.DataProviders.CBRF.XmlDeserializerToObject.Parse(root, xDoc.Root);


				int UUU = 9;
				//return rawECBString;
			}
		}
	}
}
