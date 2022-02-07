
using System.Windows.Forms;

using Navferty.Common;

#nullable enable

namespace Navferty.ExcelAddIn.Web.CurrencyExchangeRates
{
	public static class Manager
	{
		public static ExchangeRateRecord? SelectExchageRate(IDialogService dialogService)
		{
			using (var f = new frmExchangeRates(dialogService))
			{
				if (f.ShowDialog() != DialogResult.OK) return null;
				return f.SelectedExchangeRate;
			};
		}
	}
}
