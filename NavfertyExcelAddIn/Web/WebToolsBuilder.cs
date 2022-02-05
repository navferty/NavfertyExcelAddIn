
using Microsoft.Office.Interop.Excel;

using Navferty.Common;

using NavfertyExcelAddIn.Localization;

#nullable enable

namespace NavfertyExcelAddIn.Web
{
	public class WebToolsBuilder : IWebTools
	{
		internal readonly IDialogService dialogService;
		private Microsoft.Office.Interop.Excel.Application App => Globals.ThisAddIn.Application;

		public WebToolsBuilder(IDialogService dialogService)
			=> this.dialogService = dialogService;


		public void CurrencyExchangeRates_Show()
		{
			Range? sel = App.Selection;
			if (null == sel || sel.Cells == null || sel.Cells.Count < 1)
			{
				dialogService.ShowError(UIStrings.CurrencyExchangeRates_Error_NedAnyCellSelection);
				return;
			}

			var rslt = Navferty.ExcelAddIn.Web.CurrencyExchangeRates.Manager.SelectExchageRate(dialogService);
			if (rslt == null) return;//User cancel

			var exchangeRate = rslt.CursFor1Unit;
			sel.Value = exchangeRate;
		}
	}
}
