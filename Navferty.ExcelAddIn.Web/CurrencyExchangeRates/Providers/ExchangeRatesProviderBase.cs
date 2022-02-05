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
			ExchangeRateRecord[]? webRowsForDay = null;
			ExchangeRateRecord[]? webRowsForPrevDay = null;
			Exception? ex1 = null;
			Exception? ex2 = null;

			dt = dt.Date;//Cut off time
			var dtPrevDay = dt.AddDays(-1);

			var sw = new Stopwatch();
			try
			{
				Logger.Debug($"1st web query attempt for '{dt}' starting...");
				sw.Start();
				webRowsForDay = await DownloadExchangeRatesForDayAsync(dt);
				sw.Stop();
				Logger.Debug($"1st web query finished OK. Elapsed: {sw.Elapsed.TotalMilliseconds}ms.");

				if (null == webRowsForDay || !webRowsForDay.Any())
				{
					Logger.Error("1st webRowsForDay = NULL!");
					throw new Exception(UIStrings.CurrencyExchangeRates_Error_Network);
				}
			}
			catch (Exception e1)
			{
				sw.Stop();
				ex1 = e1;
				Logger.Error(ex1, $"1st web query '{dt}' Failed! Elapsed: {sw.Elapsed.TotalMilliseconds}ms.");
			}

			if (null != ex1)
			{
				Logger.Debug($"2nd web query attempt for '{dtPrevDay}' starting...");
				sw = new Stopwatch();
				try
				{
					sw.Start();
					webRowsForPrevDay = await DownloadExchangeRatesForDayAsync(dtPrevDay);
					sw.Stop();
					Logger.Debug($"2nd web query finished OK. Elapsed: {sw.Elapsed.TotalMilliseconds}ms.");

					if (null == webRowsForPrevDay || !webRowsForPrevDay.Any())
					{
						Logger.Error("2nd webRowsForPrevDay = NULL!");
						throw new Exception(UIStrings.CurrencyExchangeRates_Error_Network);
					}
				}
				catch (Exception e2)
				{
					sw.Stop();
					ex2 = e2;
					Logger.Error(ex2, $"2nd web query '{dtPrevDay}' Failed! Elapsed: {sw.Elapsed.TotalMilliseconds}ms.");
				}
			}

			if (ex1 != null && ex2 != null)
			{
				// both web requests failed!
				// Looks like we have some network problems...
				throw ex1;
			}

			if (ex1 != null && ex2 == null)
			{
				// 2-nd web request finished good!
				// The source does not yet contain data for the specified date!

				string sErr = string.Format(UIStrings.CurrencyExchangeRates_Error_NotAvailYet,
					Title,
					dt.ToLongDateString());

				Logger.Debug($"1st web request for '{dt}' failed, but 2-nd web request for '{dtPrevDay}' finished good!\nLooks like web source does not yet provide data for '{dt}'!");
				//Logger.Debug(sErr);
				throw new Exception(sErr);
			}

			CurrencyExchangeRatesDataset.ExchangeRatesDataTable dtResult = new();
			if (null != cbGetCurrencyPriority)
			{
				//Get priority for each row
				webRowsForDay.ToList().ForEach(wrr =>
			   {
				   var priority = cbGetCurrencyPriority.Invoke(wrr);
				   if (priority.HasValue) wrr.PriorityInGrid = priority.Value;
			   });
			}

			(from r in webRowsForDay
			 orderby r.PriorityInGrid ascending, r.Name ascending   //Sort by grid priority and title
			 select r)
			 .ToList().ForEach(wrr  //Populate our result datatable with rows...
			 =>
			 {
				 var newRow = dtResult.NewExchangeRatesRow();
				 {
					 newRow.Raw = wrr;
					 newRow.Name = wrr.DisplayName;
					 newRow.ISO = wrr.ISOCode;
					 newRow.Rate = wrr.Curs;
				 }
				 dtResult.Rows.Add(newRow);
			 });

			return dtResult;
		}

		protected abstract Task<ExchangeRateRecord[]> DownloadExchangeRatesForDayAsync(DateTime dt);

		public override string ToString() => Title;
	}
}
