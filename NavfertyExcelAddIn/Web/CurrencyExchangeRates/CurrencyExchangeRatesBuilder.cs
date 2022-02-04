
using Microsoft.Office.Interop.Excel;

using NavfertyCommon;

using NavfertyExcelAddIn.Localization;

#nullable enable

namespace NavfertyExcelAddIn.Web.CurrencyExchangeRates
{
	public class CurrencyExchangeRatesBuilder : ICurrencyExchangeRates
	{
		internal readonly IDialogService dialogService;
		private Microsoft.Office.Interop.Excel.Application App => Globals.ThisAddIn.Application;

		public CurrencyExchangeRatesBuilder(IDialogService dialogService)
			=> this.dialogService = dialogService;


		public void ShowCurrencyExchangeRates(Workbook wb)
		{
			Range? sel = App.Selection;
			if (null == sel || sel.Cells == null || sel.Cells.Count < 1)
			{
				dialogService.ShowError(UIStrings.CurrencyExchangeRates_Error_NedAnyCellSelection);
				return;
			}

			var rslt = frmExchangeRates.SelectExchageRates(dialogService);
			if (rslt == null) return;
			var exchangeRate = rslt.CursFor1Unit;
			sel.Value = exchangeRate;
		}
	}
}
