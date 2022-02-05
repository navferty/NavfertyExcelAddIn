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
	internal abstract class ExchangeRatesDataProviderBaase
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
			ExchangeRateRecord[] webRowsForDay = { };
			ExchangeRateRecord[] webRowsForPrevDay = { };
			Exception? ex1 = null;
			Exception? ex2 = null;

			var sw = new Stopwatch();
			try
			{
				Logger.Debug($"{this.GetType()} ({this.Title}): 1ST web query attempt for '{dt}' starting...");
				sw.Start();
				webRowsForDay = await DownloadWebResultRowsForDate(dt);
				sw.Stop();
				Logger.Debug($"\t1ST web query finished OK. Elapsed: {sw.Elapsed.TotalMilliseconds}ms.");
			}
			catch (Exception e1)
			{
				sw.Stop();
				ex1 = e1;
				Logger.Error(ex1, $"1ST web query (for date '{dt}') Failed! Elapsed: {sw.Elapsed.TotalMilliseconds}ms.");
			}
			if (null != ex1 || !webRowsForDay.Any())
			{
				var dtPrevDay = dt.AddDays(-1);
				Logger.Debug($"{this.GetType()} ({this.Title}): 2ND web query attempt for '{dtPrevDay}' starting...");
				sw = new Stopwatch();
				try
				{
					sw.Start();
					webRowsForPrevDay = await DownloadWebResultRowsForDate(dtPrevDay);
					sw.Stop();
					Logger.Debug($"\t2ND web query finished OK. Elapsed: {sw.Elapsed.TotalMilliseconds}ms.");
				}
				catch (Exception e2)
				{
					sw.Stop();
					ex2 = e2;
					Logger.Debug(ex2, $"2ND web query (for date '{dtPrevDay}') Failed! Elapsed: {sw.Elapsed.TotalMilliseconds}ms.");
				}
			}


			if (ex1 != null && ex2 != null)
			{
				// No one web request finished good.
				// Looks like we have some network problems...
				throw ex1 ?? new Exception(UIStrings.CurrencyExchangeRates_Error_Network);
			}
			if (ex1 != null && ex2 == null)
			{
				// 2-nd web request finished good!
				// The source does not yet contain data for the specified date!

				var sErr = string.Format(UIStrings.CurrencyExchangeRates_Error_NotAvailYet,
					Title,
					dt.ToLongDateString());

				Logger.Debug(sErr);
				throw new Exception(sErr);
			}

			var dtResult = new CurrencyExchangeRatesDataset.ExchangeRatesDataTable();
			if (!webRowsForDay.Any()) return dtResult;

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
				 select r).ToList().ForEach(wrr
				 =>
				 {
					 var newRow = dtResult.NewExchangeRatesRow();
					 newRow.Raw = wrr;
					 newRow.Name = wrr.DisplayName;
					 newRow.ISO = wrr.ISOCode;
					 newRow.Rate = wrr.Curs;
					 dtResult.Rows.Add(newRow);//Populate our result datatable with rows...
				 });
			return dtResult;
		}

		protected abstract Task<ExchangeRateRecord[]> DownloadWebResultRowsForDate(DateTime dt);

		public override string ToString() => Title;
	}
}
