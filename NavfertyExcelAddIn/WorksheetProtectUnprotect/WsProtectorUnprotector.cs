using System;
using System.Windows.Forms;

using Microsoft.Office.Interop.Excel;

using Navferty.Common;

using NavfertyExcelAddIn.Localization;

namespace NavfertyExcelAddIn.WorksheetProtectUnprotect
{
	public class WsProtectorUnprotector : IWsProtectorUnprotector
	{

		internal readonly IDialogService dialogService;
		private Microsoft.Office.Interop.Excel.Application App => Globals.ThisAddIn.Application;

		public WsProtectorUnprotector(IDialogService dialogService)
			=> this.dialogService = dialogService;

		public void ProtectUnprotectSelectedWorksheets(Workbook wb)
		{
			try
			{
				Sheets wbSheets = wb.Worksheets;
				if (wbSheets.Count < 1)
					throw new Exception(string.Format(UIStrings.WorkSheetsNotFound, wb.FullName));

				using var f = new frmWorksheetsProtection(this, wb);
				if (f.ShowDialog() != DialogResult.OK) return;
			}
			catch (Exception ex) { dialogService.ShowError(ex); }
		}
	}
}
