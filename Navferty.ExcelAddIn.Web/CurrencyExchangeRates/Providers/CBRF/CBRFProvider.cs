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
	internal class CBRFProvider : ExchangeRatesProviderBase
	{
		private static readonly CultureInfo ci = CultureInfo.GetCultureInfo("ru-RU");

		public override string Title => UIStrings.CurrencyExchangeRates_Sources_CBRF;

		public override CultureInfo Culture => ci;

		private readonly ILogger logger = LogManager.GetCurrentClassLogger();
		public override ILogger Logger => logger;

		private DataTable? rawDataTable = null;
		protected override async Task<ExchangeRateRecord[]> DownloadExchangeRatesForDayAsync(DateTime dt)
		{

			using (var cbr = new cbrwebservice.DailyInfoSoapClient())
			{
				var dtsResult = await cbr.GetCursOnDateAsync(dt);
				if (dtsResult == null)
				{
					logger.Error("async await cbr.GetCursOnDateAsync(dt) return NULL with no errors!");
					throw new Exception(UIStrings.CurrencyExchangeRates_Error_Network);
				}

				rawDataTable = dtsResult.Tables.Cast<DataTable>().FirstOrDefault();
				if (rawDataTable == default)
				{
					logger.Error($"Failed to get frist DataTable from web dstaset ('{dtsResult.GetType()}'). Dataset does not containt any Tables!");
					throw new Exception(UIStrings.CurrencyExchangeRates_Error_Network);
				}

				var dataTableRows = rawDataTable.RowsAsEnumerable().ToArray();
				logger.Debug($"dataTableRows.Count = {dataTableRows.Count()}");
				try
				{
					var rows = dataTableRows.Select(row => new ExchangeRateRecord(row, dt)).ToArray();
					return rows;
				}
				catch (Exception ex)
				{

					logger.Error(ex, $"Failed to convert '{dataTableRows.GetType()}' to '{typeof(ExchangeRateRecord)}'!");
					throw new Exception(UIStrings.CurrencyExchangeRates_Error_ParseError);
				}
			};
		}
	}
}
