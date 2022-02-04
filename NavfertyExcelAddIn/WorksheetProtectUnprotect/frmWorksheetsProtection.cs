using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Linq;

using Microsoft.Office.Interop.Excel;

using NavfertyExcelAddIn.Commons;
using NavfertyExcelAddIn.Localization;


namespace NavfertyExcelAddIn.WorksheetProtectUnprotect
{
	internal partial class frmWorksheetsProtection : Commons.Controls.FormEx
	{

		private Microsoft.Office.Interop.Excel.Application App => Globals.ThisAddIn.Application;

		private readonly WsProtectorUnprotector creator = null;
		private readonly Workbook wb = null;

		private Sheets GetSheets() => wb.Worksheets;

		public frmWorksheetsProtection()
		{
			InitializeComponent();
		}

		public frmWorksheetsProtection(WsProtectorUnprotector Creator, Workbook wb) : this()
		{
			this.creator = Creator;
			this.wb = wb;


			Text = $"{RibbonLabels.ProtectUnprotectWorksheets} '{wb.Name}'";
			radioModeProtect.Text = UIStrings.SheetProtection_Set;
			radioModeUnProtect.Text = UIStrings.SheetProtection_Clear;

			lblModeDesription.Text = UIStrings.SheetProtection_ProtectionForSheets;
			lblPWD.Text = UIStrings.SheetProtection_Password;
			txtPWD.Text = string.Empty;

			btnExecProtectionAction.Text = UIStrings.SheetProtection_Execute;
			btnExecProtectionAction.Enabled = false;

			lstWorksheets.EmptyText = UIStrings.NoMatchingWorkSheets;
			lstWorksheets.EmptyTextAlign = ContentAlignment.MiddleCenter;

			this.Load += (s, e) => OnFormLoad();
		}

		private void OnFormLoad()
		{
			txtPWD.SetVistaCueBanner(UIStrings.SheetProtection_PasswordBanner);

			OnSelectProtectAction();

			radioModeProtect.CheckedChanged += (s, e) => OnSelectProtectAction();
			lstWorksheets.ItemCheck += OnSheetInListChecked;
			btnExecProtectionAction.Click += (s, e) => OnExecProtectAction();
		}

		private bool hasSheetsToProcess = false;

		private void OnSelectProtectAction()
		{
			Cursor = Cursors.WaitCursor;
			try
			{
				hasSheetsToProcess = false;

				bool bProtect = radioModeProtect.Checked;
				var rows = GetSheets().Cast<Worksheet>()
					.Select(ws => new WorksheetRow(ws))
					.Where(wsr => (wsr.HasAnyProtectedObjects() != bProtect))
					.ToArray();

				hasSheetsToProcess = rows.Any();

				lstWorksheets.Items.Clear();
				if (hasSheetsToProcess)
				{
					lstWorksheets.Items.AddRange(rows);
					int i = 0;
					lstWorksheets.Items.Cast<WorksheetRow>().ToList().ForEach(item =>
					{
						lstWorksheets.SetItemChecked(i, true);
						i++;
					});
				}
				else
				{
					//lstWorksheets.Items.Add(UIStrings.NoMatchingWorkSheets);
				}

			}
			catch (Exception ex)
			{
				creator.dialogService.ShowError(ex.Message);
			}
			finally
			{
				lstWorksheets.Enabled = hasSheetsToProcess;//Disable listbox if it have not valid items
				AfterSheetChecked();
				Cursor = Cursors.Default;
			}
		}

		private void OnSheetInListChecked(object sender, ItemCheckEventArgs e)
		{
			//inside OnSheetInListChecked() handler, lst.CheckedItems still not return actual row checked staus before methos finished
			//we shcedule our action to execute after exit control events handler pipeline
			((System.Action)AfterSheetChecked).RunDelayed();
			//And now we exit from event handler, but our delayed AfterSheetChecked() will be run soon...
		}

		private void AfterSheetChecked()
		{
			var rowsToProcees = lstWorksheets.CheckedItems.Cast<WorksheetRow>().ToArray();
			hasSheetsToProcess = rowsToProcees.Any();

			txtPWD.Enabled = hasSheetsToProcess;
			btnExecProtectionAction.Enabled = hasSheetsToProcess;
		}

		private void OnExecProtectAction()
		{
			UseWaitCursor = true;
			try
			{
				var rowsToProcees = lstWorksheets.CheckedItems.Cast<WorksheetRow>().ToArray();
				if (!rowsToProcees.Any())
				{
					creator.dialogService.ShowError(UIStrings.NoMatchingWorkSheets);
					return;
				}

				var pwd = txtPWD.Text;
				bool hasPWD = (null != pwd) && (!string.IsNullOrEmpty(pwd));
				bool bProtect = radioModeProtect.Checked;
				bool wasErrorsInProcessingSheets = false;

				rowsToProcees.ToList().ForEach(item =>
				{
					try
					{
						Worksheet ws = item.Sheet;
						if (bProtect)
						{
							if (hasPWD)
								ws.Protect(Password: pwd);
							else
								ws.Protect();
						}
						else
						{
							if (hasPWD)
								ws.Unprotect(Password: pwd);
							else
								ws.Unprotect();
						}
					}
					catch (Exception ex)
					{
						wasErrorsInProcessingSheets = true;
						creator.dialogService.ShowError(ex.Message);
					}
				});
				if (!wasErrorsInProcessingSheets) DialogResult = DialogResult.OK;

				//was any errors, do not close dialog
			}
			catch (Exception ex)
			{
				creator.dialogService.ShowError(ex.Message);
			}
			finally { UseWaitCursor = false; }

			//refill list of sheets with updated protectiob status
			OnSelectProtectAction();
		}
	}
}
