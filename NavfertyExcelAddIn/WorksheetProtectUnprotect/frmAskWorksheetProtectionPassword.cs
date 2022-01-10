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
<<<<<<< HEAD
<<<<<<< HEAD
			radioModeProtect.Text = UIStrings.SheetProtection_Set;
			radioModeUnProtect.Text = UIStrings.SheetProtection_Clear;

			lblModeDesription.Text = UIStrings.SheetProtection_ProtectionForSheets;
			lblPWD.Text = UIStrings.SheetProtection_Password;
			txtPWD.Text = string.Empty;

			btnExecProtectionAction.Text = UIStrings.SheetProtection_Execute;
			btnExecProtectionAction.Enabled = false;

			this.Load += (s, e) => OnFormLoad();
=======
			radioModeProtect.Text = UIStrings.Protection_Set;
			radioModeUnProtect.Text = UIStrings.Protection_Clear;
=======
			radioModeProtect.Text = UIStrings.SheetProtection_Set;
			radioModeUnProtect.Text = UIStrings.SheetProtection_Clear;
>>>>>>> 2bbf414 (v1 finished)

			lblModeDesription.Text = UIStrings.SheetProtection_ProtectionForSheets;
			lblPWD.Text = UIStrings.SheetProtection_Password;
			txtPWD.Text = string.Empty;

			btnExecProtectionAction.Text = UIStrings.SheetProtection_Execute;
			btnExecProtectionAction.Enabled = false;

			this.Load += (s, e) => OnFormLoad();
<<<<<<< HEAD

>>>>>>> add758c (creating ProtectUnprotectSelectedWorksheets)
=======
>>>>>>> 2bbf414 (v1 finished)
		}

		private void OnFormLoad()
		{
			OnSelectProtectAction();
<<<<<<< HEAD
<<<<<<< HEAD

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
						//lstWorksheets.SetItemChecked(i, !item.Sheet.ProtectionMode);
						lstWorksheets.SetItemChecked(i, true);
						i++;
					});
				}
				else
				{
					lstWorksheets.Items.Add(UIStrings.NoMatchingWorkSheets);
				}

			}
			catch (Exception ex)
			{
				creator.dialogService.ShowError(ex.Message);
			}
			finally
			{
				lstWorksheets.Enabled = hasSheetsToProcess;

				AfterSheetChecked();
				Cursor = Cursors.Default;
			}
		}

		private void OnSheetInListChecked(object sender, ItemCheckEventArgs e)
		{
			//inside OnSheetInListChecked() handler, lst.CheckedItems still not return actual row checked staus before methos finished
			//we create small timer, which delays row checking after we exited OnSheetInListChecked.

			var tmrUIProcessPause = new System.Windows.Forms.Timer()
			{
				Interval = 100, //set as short time as we can, to only allow to finish OnSheetInListChecked(), but not allow any overheaded UI events...
				Enabled = false //do not start timer untill we finish it's setup
			};
			tmrUIProcessPause.Tick += (s, te) =>
			{
				//This Work!!!
				tmrUIProcessPause.Stop();//first stop and dispose our timer, to avoid double execution
				tmrUIProcessPause.Dispose();

				AfterSheetChecked();//Now start real row checking processing...
			};


			tmrUIProcessPause.Start();//Start pause timer

			//now we exit from event handler, but our delayed (timered) AfterSheetChecked must be run...
		}

		private void AfterSheetChecked()
		{
			var rowsToProcees = lstWorksheets.CheckedItems.Cast<WorksheetRow>().ToArray();
			hasSheetsToProcess = rowsToProcees.Any();

			txtPWD.Enabled = hasSheetsToProcess;
			btnExecProtectionAction.Enabled = hasSheetsToProcess;
=======
=======

>>>>>>> 2bbf414 (v1 finished)
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
						//lstWorksheets.SetItemChecked(i, !item.Sheet.ProtectionMode);
						lstWorksheets.SetItemChecked(i, true);
						i++;
					});
				}
				else
				{
					lstWorksheets.Items.Add(UIStrings.NoMatchingWorkSheets);
				}

			}
<<<<<<< HEAD
			finally { Cursor = Cursors.Default; }
>>>>>>> add758c (creating ProtectUnprotectSelectedWorksheets)
=======
			catch (Exception ex)
			{
				creator.dialogService.ShowError(ex.Message);
			}
			finally
			{
				lstWorksheets.Enabled = hasSheetsToProcess;

				AfterSheetChecked();
				Cursor = Cursors.Default;
			}
		}

		private void OnSheetInListChecked(object sender, ItemCheckEventArgs e)
		{
			//inside OnSheetInListChecked() handler, lst.CheckedItems still not return actual row checked staus before methos finished
			//we create small timer, which delays row checking after we exited OnSheetInListChecked.

			var tmrUIProcessPause = new System.Windows.Forms.Timer()
			{
				Interval = 100, //set as short time as we can, to only allow to finish OnSheetInListChecked(), but not allow any overheaded UI events...
				Enabled = false //do not start timer untill we finish it's setup
			};
			tmrUIProcessPause.Tick += (s, te) =>
			{
				//This Work!!!
				tmrUIProcessPause.Stop();//first stop and dispose our timer, to avoid double execution
				tmrUIProcessPause.Dispose();

				AfterSheetChecked();//Now start real row checking processing...
			};


			tmrUIProcessPause.Start();//Start pause timer

			//now we exit from event handler, but our delayed (timered) AfterSheetChecked must be run...
		}

		private void AfterSheetChecked()
		{
			var rowsToProcees = lstWorksheets.CheckedItems.Cast<WorksheetRow>().ToArray();
			hasSheetsToProcess = rowsToProcees.Any();

			txtPWD.Enabled = hasSheetsToProcess;
			btnExecProtectionAction.Enabled = hasSheetsToProcess;
>>>>>>> 2bbf414 (v1 finished)
		}

		private void OnExecProtectAction()
		{
			Cursor = Cursors.WaitCursor;
			try
			{
				var rowsToProcees = lstWorksheets.CheckedItems.Cast<WorksheetRow>().ToArray();
<<<<<<< HEAD
<<<<<<< HEAD
=======
>>>>>>> 2bbf414 (v1 finished)
				if (!rowsToProcees.Any())
				{
					creator.dialogService.ShowError(UIStrings.NoMatchingWorkSheets);
					return;
				}
<<<<<<< HEAD

				var pwd = txtPWD.Text;
				bool hasPWD = (null != pwd) && (!string.IsNullOrEmpty(pwd));
=======
				if (!rowsToProcees.Any()) return;

				var pwd = txtPWD.Text;
>>>>>>> add758c (creating ProtectUnprotectSelectedWorksheets)
=======

				var pwd = txtPWD.Text;
				bool hasPWD = (null != pwd) && (!string.IsNullOrEmpty(pwd));
>>>>>>> 2bbf414 (v1 finished)
				bool bProtect = radioModeProtect.Checked;
				bool wasErrorsInProcessingSheets = false;

				rowsToProcees.ToList().ForEach(item =>
				{
					try
					{
<<<<<<< HEAD
<<<<<<< HEAD
=======
>>>>>>> 2bbf414 (v1 finished)
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
<<<<<<< HEAD
=======
						item.Sheet.Protect(Password: pwd);
>>>>>>> add758c (creating ProtectUnprotectSelectedWorksheets)
=======
>>>>>>> 2bbf414 (v1 finished)
					}
					catch (Exception ex)
					{
						wasErrorsInProcessingSheets = true;
						creator.dialogService.ShowError(ex.Message);
					}
				});
				if (!wasErrorsInProcessingSheets) DialogResult = DialogResult.OK;

<<<<<<< HEAD
<<<<<<< HEAD
				//was any errors, do not close dialog
=======
>>>>>>> add758c (creating ProtectUnprotectSelectedWorksheets)
=======
				//was any errors, do not close dialog
>>>>>>> 2bbf414 (v1 finished)
			}
			catch (Exception ex)
			{
				creator.dialogService.ShowError(ex.Message);
			}
			finally { Cursor = Cursors.Default; }

<<<<<<< HEAD
<<<<<<< HEAD
			//refill list of sheets with updated protectiob status
			OnSelectProtectAction();
		}

=======
			//refill sheets list with updated protectiob status
			OnSelectProtectAction();
		}



>>>>>>> add758c (creating ProtectUnprotectSelectedWorksheets)
=======
			//refill list of sheets with updated protectiob status
			OnSelectProtectAction();
		}

>>>>>>> 2bbf414 (v1 finished)
	}
}
