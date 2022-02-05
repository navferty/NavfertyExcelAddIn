using System;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;

using Navferty.ExcelAddIn.Web.Localization;

using NLog;

#nullable enable


namespace Navferty.ExcelAddIn.Web.CurrencyExchangeRates.Providers
{
	internal abstract class ExchangeRatesProviderBase
	{
		public abstract string Title { get; }
		public abstract ILogger Logger { get; }

		public abstract CultureInfo Culture { get; }

		public uint Priority
		{
			get
			{
				var bIsCurrent = (Culture.LCID == CultureInfo.CurrentUICulture.LCID);

				return bIsCurrent ? 1u : uint.MaxValue;
			}
		}

		public async Task<CurrencyExchangeRatesDataset.ExchangeRatesDataTable> GetExchangeRatesForDate(
			DateTime dt,
			Func<ExchangeRateRecord, uint?> cbGetCurrencyPriority)
		{
			dt = dt.Date;//Cut off time			

			var result = await TryGetExchangeRatesForDate(dt, 1);
			if (!result.HasData)
			{
				var dtPrevDay = dt.AddDays(-3); //We're going back a few days.
												//For some sources, the data may not be available at the end of the banking day in their local time,
												//so a request for -1 day ago, according to the user's calendar,
												//may also not work and lead to errors.
												//So we go straight to 3 days ago, to avoid "false" errors!


				var result2 = await TryGetExchangeRatesForDate(dtPrevDay, 2);
				if (!result2.HasData) throw result.Err!;  // both web requests failed! Looks like we have some network problems...

				// 2-nd web request finished good! The source does not yet contain data for the specified date!
				string sErr = string.Format(UIStrings.CurrencyExchangeRates_Error_NotAvailYet,
					Title,
					dt.ToLongDateString());

				Logger.Debug($"1st web request for '{dt}' failed, but 2-nd web request for '{dtPrevDay}' finished good! Looks like web source does not yet provide data for '{dt}'!");
				throw new Exception(sErr);
			}

			var webRows = result.Rows!;
			if (null != cbGetCurrencyPriority)
			{
				//Get priority for each row
				webRows.ToList().ForEach(wrr =>
			   {
				   var priority = cbGetCurrencyPriority.Invoke(wrr);
				   if (priority.HasValue) wrr.PriorityInGrid = priority.Value;
			   });
			}

			return ExchangeRateRecord.ToDataTable(webRows);
		}

		private struct WebResult
		{
			public readonly ExchangeRateRecord[]? Rows = null;
			public readonly Exception? Err = null;

			public WebResult(ExchangeRateRecord[] rows)
			{
				this.Rows = rows;
				if (!HasData) Err = new Exception("Web query finished without error (Unknown error)!");
			}

			public WebResult(Exception e) { Err = e; }

			public bool HasData => (null != Rows && Rows.Any());
		}

		private async Task<WebResult> TryGetExchangeRatesForDate(DateTime dt, int nTry)
		{
			string sDate = dt.ToString("yyyy-MM-dd");
			Logger.Debug($"Starting Web query attempt #{nTry}, Date: '{sDate}'...");
			var sw = new Stopwatch();
			try
			{
				sw.Start();
				var webRows = await DownloadExchangeRatesForDayAsync(dt);
				sw.Stop();
				Logger.Debug($"Web query finished OK. Elapsed: {sw.Elapsed.TotalMilliseconds}ms.");
				var wr = new WebResult(webRows);
				if (!wr.HasData) Logger.Error("Web query result: webRows is NULL or Empty rows!");
				return wr;
			}
			catch (Exception ex)
			{
				sw.Stop();
				Logger.Error(ex, $"Web query Failed! Elapsed: {sw.Elapsed.TotalMilliseconds}ms.");
				return new WebResult(ex);
			}
		}

		protected abstract Task<ExchangeRateRecord[]> DownloadExchangeRatesForDayAsync(DateTime dt);

		public override string ToString() => Title;
	}
}
