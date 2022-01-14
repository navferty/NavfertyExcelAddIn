using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;

using Microsoft.Office.Interop.Excel;

using NavfertyExcelAddIn.Commons;
using NavfertyExcelAddIn.Localization;

namespace NavfertyExcelAddIn.WorksheetProtectorUnprotector
{
	public class WsProtectorUnprotector : IWsProtectorUnprotector
	{

		internal readonly IDialogService dialogService;
		private Microsoft.Office.Interop.Excel.Application App => Globals.ThisAddIn.Application;

		public WsProtectorUnprotector(IDialogService dialogService)
			=> this.dialogService = dialogService;

		public void ProtectUnprotectSelectedWorksheets(Workbook wb)
		{
			Sheets wbSheets = wb.Worksheets;
			if (wbSheets.Count < 1)
			{
				dialogService.ShowError(string.Format(UIStrings.WorkSheetsNotFound, wb.FullName));
			}

			using (var f = new frmAskWorksheetProtectionPassword(this, wb))
			{
				if (f.ShowDialog() != DialogResult.OK) return;
			}
			//MessageBox.Show($"ProtectUnprotectWorksheets {wb.FullName}", wb.FullName);
		}
	}
}
