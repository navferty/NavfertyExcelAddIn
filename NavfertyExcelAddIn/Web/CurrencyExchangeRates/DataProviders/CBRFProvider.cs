using System;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;

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
		protected override async Task<WebResultRow[]> GetExchabgeRatesForDate_Core(DateTime dt)
		{
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
	}
}
