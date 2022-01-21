using System;
using System.Globalization;
using System.Threading.Tasks;

using NavfertyExcelAddIn.Localization;

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

		public async Task<WebResultRow[]> GetExchabgeRatesForDate(DateTime dt)
		{
			try
			{
				return await GetExchabgeRatesForDate_Core(dt);
			}
			catch (Exception ex)
			{
				throw new Exception(UIStrings.CurrencyExchangeRates_Error_RemoteSource, ex);
			}
		}

		protected abstract Task<WebResultRow[]> GetExchabgeRatesForDate_Core(DateTime dt);

		public override string ToString() => Title;
	}
}
