using System;
using System.Globalization;
using System.Threading.Tasks;
using System.Linq;

using NavfertyExcelAddIn.Localization;
using System.Collections.Generic;

namespace NavfertyExcelAddIn.Web.CurrencyExchangeRates.Providers
{
	internal abstract class ExchangeRatesDataProviderBaase
	{
		public abstract string Title { get; }

		public abstract CultureInfo Culture { get; }

		public uint Priority
		{
			get
			{
				var bIsCurrent = (Culture.LCID == CultureInfo.CurrentUICulture.LCID);

				return bIsCurrent ? 1u : uint.MaxValue;
			}
		}

		public async Task<CurrencyExchangeRatesDataset.ExchangeRatesDataTable> GetExchabgeRatesForDate(DateTime dt, Func<WebResultRow, uint?> cbGetCurrencyPriority)
		{
			try
			{
				var dtView = new CurrencyExchangeRatesDataset.ExchangeRatesDataTable();
				WebResultRow[] webRows = await DownloadWebResultRowsForDate(dt);
				if (!webRows.Any()) return dtView;

				if (null != cbGetCurrencyPriority)
				{
					//Get priority for each row
					webRows.ToList().ForEach(wrr =>
				   {
					   var priority = cbGetCurrencyPriority.Invoke(wrr);
					   if (priority.HasValue) wrr.PriorityInGrid = priority.Value;
				   });
				}

				//Sort by grid priority and title
				webRows = (from r in webRows
						   orderby r.PriorityInGrid ascending, r.Name ascending
						   select r).ToArray();

				foreach (var wrr in webRows)
				{
					var newRow = dtView.NewExchangeRatesRow();
					newRow.Raw = wrr;
					newRow.Name = wrr.DisplayName;
					newRow.ISO = wrr.ISOCode;
					newRow.Rate = wrr.Curs;
					dtView.Rows.Add(newRow);
				}
				return dtView;
			}
			catch (Exception ex)
			{
				throw new Exception(UIStrings.CurrencyExchangeRates_Error_RemoteSource, ex);
			}
		}

		protected abstract Task<WebResultRow[]> DownloadWebResultRowsForDate(DateTime dt);

		public override string ToString() => Title;
	}
}
