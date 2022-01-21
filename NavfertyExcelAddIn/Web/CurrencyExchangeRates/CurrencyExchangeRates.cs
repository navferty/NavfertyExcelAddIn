using System.Windows.Forms;

using Microsoft.Office.Interop.Excel;

using NavfertyExcelAddIn.Commons;
using NavfertyExcelAddIn.Localization;


namespace NavfertyExcelAddIn.Web.CurrencyExchangeRates
{
	public class CurrencyExchangeRates : ICurrencyExchangeRates
	{
		internal readonly IDialogService dialogService;
		private Microsoft.Office.Interop.Excel.Application App => Globals.ThisAddIn.Application;

		public CurrencyExchangeRates(IDialogService dialogService)
			=> this.dialogService = dialogService;


		public void ShowCurrencyExchangeRates(Workbook wb)
		{
			if (App.Selection == null
				|| ((Range)App.Selection).Cells == null
				|| ((Range)App.Selection).Cells.Count != 1)
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
