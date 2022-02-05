using System;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;

using Navferty.Common;

using Navferty.ExcelAddIn.Web.Localization;

using NLog;

#nullable enable

namespace Navferty.ExcelAddIn.Web.CurrencyExchangeRates.Providers
{
	internal class CBRFProvider : ExchangeRatesDataProviderBaase
	{
		private static readonly CultureInfo ci = CultureInfo.GetCultureInfo("ru-RU");

		public override string Title => UIStrings.CurrencyExchangeRates_Sources_CBRF;

		public override CultureInfo Culture => ci;

		private readonly ILogger logger = LogManager.GetCurrentClassLogger();
		public override ILogger Logger => logger;

		private DataTable? rawDataTable = null;
		protected override async Task<WebResultRow[]> DownloadWebResultRowsForDate(DateTime dt)
		{

			using (var cbr = new cbrwebservice.DailyInfoSoapClient())
			{
				var dtsResult = await cbr.GetCursOnDateAsync(dt);
				if (dtsResult == null)
				{
					logger.Error("CBR: await cbr.GetCursOnDateAsync(dt) return NULL with no errors!");
					throw new Exception(UIStrings.CurrencyExchangeRates_Error_Network);
				}

				rawDataTable = dtsResult.Tables.Cast<DataTable>().FirstOrDefault();
				if (rawDataTable == default)
				{
					logger.Error("CBR: dstaset does not containt any Tables!");
					throw new Exception(UIStrings.CurrencyExchangeRates_Error_Network);
				}

				try
				{
					var rows = rawDataTable.RowsAsEnumerable().Select(row => new WebResultRow(row, dt)).ToArray();
					return rows;
				}
				catch (Exception ex)
				{
					logger.Error(ex, "CBR: Failed to convert 'CBR.Datatable[0].DataRow' to 'WebResultRow'!");
					throw new Exception(UIStrings.CurrencyExchangeRates_Error_ParseError);
				}
			};
		}
	}
}
