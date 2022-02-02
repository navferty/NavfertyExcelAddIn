using System.Windows.Forms;

using Microsoft.Office.Interop.Excel;

using NavfertyExcelAddIn.Commons;
using NavfertyExcelAddIn.Localization;


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

			using (var f = new frmExchangeRates(this, wb))
			{
				if (f.ShowDialog() != DialogResult.OK) return;
			};
		}
	}
}
