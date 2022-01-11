using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Microsoft.Office.Interop.Excel;

using NavfertyExcelAddIn.Localization;

namespace NavfertyExcelAddIn.WorksheetProtectorUnprotector
{
	public partial class frmAskWorksheetProtectionPassword : Form
	{

		private Microsoft.Office.Interop.Excel.Application App => Globals.ThisAddIn.Application;

		private readonly WsProtectorUnprotector creator = null;
		private readonly Workbook wb = null;

		private Sheets GetSheets() => wb.Worksheets;

		public frmAskWorksheetProtectionPassword()
		{
			InitializeComponent();
		}

		public frmAskWorksheetProtectionPassword(WsProtectorUnprotector Creator, Workbook wb) : this()
		{
			this.creator = Creator;
			this.wb = wb;


			Text = RibbonLabels.ProtectUnprotectWorksheets;
			radioModeProtect.Text = UIStrings.Protection_Set;
			radioModeUnProtect.Text = UIStrings.Protection_Clear;

			lblModeDesription.Text = UIStrings.ProtectionForSheets;
			lblPWD.Text = UIStrings.Password;

			btnExecProtectionAction.Text = UIStrings.Execute;
			btnExecProtectionAction.Enabled = false;

			this.Load += (s, e) => OnFormLoad();

		}

		private void OnFormLoad()
		{
			OnSelectProtectAction();
			radioModeProtect.CheckedChanged += (s, e) => OnSelectProtectAction();
			btnExecProtectionAction.Click += (s, e) => OnExecProtectAction();
		}

		private void OnSelectProtectAction()
		{
			Cursor = Cursors.WaitCursor;
			try
			{

				bool bProtect = radioModeProtect.Checked;

				var rows = GetSheets().Cast<Worksheet>()
					.Where(ws => !(ws.ProtectionMode == bProtect))
					.Select(ws => new WorksheetRow(ws)).ToArray();

				lstWorksheets.Items.Clear();
				lstWorksheets.Items.AddRange(rows);

				int i = 0;
				lstWorksheets.Items.Cast<WorksheetRow>().ToList().ForEach(item =>
				{
					//lstWorksheets.SetItemChecked(i, !item.Sheet.ProtectionMode);
					lstWorksheets.SetItemChecked(i, true);
					i++;
				});

				txtPWD.Enabled = rows.Any();
				btnExecProtectionAction.Enabled = rows.Any();

			}
			finally { Cursor = Cursors.Default; }
		}

		private void OnExecProtectAction()
		{
			Cursor = Cursors.WaitCursor;
			try
			{
				var rowsToProcees = lstWorksheets.CheckedItems.Cast<WorksheetRow>().ToArray();
				if (!rowsToProcees.Any()) return;

				var pwd = txtPWD.Text;
				bool bProtect = radioModeProtect.Checked;
				bool wasErrorsInProcessingSheets = false;

				rowsToProcees.ToList().ForEach(item =>
				{
					try
					{
						item.Sheet.Protect(Password: pwd);
					}
					catch (Exception ex)
					{
						wasErrorsInProcessingSheets = true;
						creator.dialogService.ShowError(ex.Message);
					}
				});
				if (!wasErrorsInProcessingSheets) DialogResult = DialogResult.OK;

			}
			catch (Exception ex)
			{
				creator.dialogService.ShowError(ex.Message);
			}
			finally { Cursor = Cursors.Default; }

			//refill sheets list with updated protectiob status
			OnSelectProtectAction();
		}



	}
}
