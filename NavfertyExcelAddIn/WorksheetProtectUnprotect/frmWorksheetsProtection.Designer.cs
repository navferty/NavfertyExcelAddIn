

namespace NavfertyExcelAddIn.WorksheetProtectUnprotect
{
	partial class frmWorksheetsProtection
	{
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing && (components != null))
			{
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region Windows Form Designer generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.tlpMain = new System.Windows.Forms.TableLayoutPanel();
			this.btnExecProtectionAction = new System.Windows.Forms.Button();
			this.radioModeProtect = new System.Windows.Forms.RadioButton();
			this.lblModeDesription = new System.Windows.Forms.Label();
			this.radioModeUnProtect = new System.Windows.Forms.RadioButton();
			this.lstWorksheets = new NavfertyExcelAddIn.Commons.Controls.CheckedListBoxEx();
			this.lblPWD = new System.Windows.Forms.Label();
			this.txtPWD = new System.Windows.Forms.TextBox();
			this.panel1 = new System.Windows.Forms.Panel();
			this.tlpMain.SuspendLayout();
			this.SuspendLayout();
			// 
			// tlpMain
			// 
			this.tlpMain.ColumnCount = 2;
			this.tlpMain.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
			this.tlpMain.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
			this.tlpMain.Controls.Add(this.btnExecProtectionAction, 1, 7);
			this.tlpMain.Controls.Add(this.radioModeProtect, 0, 0);
			this.tlpMain.Controls.Add(this.lblModeDesription, 0, 2);
			this.tlpMain.Controls.Add(this.radioModeUnProtect, 0, 1);
			this.tlpMain.Controls.Add(this.lstWorksheets, 0, 3);
			this.tlpMain.Controls.Add(this.lblPWD, 0, 5);
			this.tlpMain.Controls.Add(this.txtPWD, 1, 5);
			this.tlpMain.Controls.Add(this.panel1, 0, 4);
			this.tlpMain.Dock = System.Windows.Forms.DockStyle.Fill;
			this.tlpMain.Location = new System.Drawing.Point(8, 8);
			this.tlpMain.Name = "tlpMain";
			this.tlpMain.RowCount = 8;
			this.tlpMain.RowStyles.Add(new System.Windows.Forms.RowStyle());
			this.tlpMain.RowStyles.Add(new System.Windows.Forms.RowStyle());
			this.tlpMain.RowStyles.Add(new System.Windows.Forms.RowStyle());
			this.tlpMain.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
			this.tlpMain.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 4F));
			this.tlpMain.RowStyles.Add(new System.Windows.Forms.RowStyle());
			this.tlpMain.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 8F));
			this.tlpMain.RowStyles.Add(new System.Windows.Forms.RowStyle());
			this.tlpMain.Size = new System.Drawing.Size(537, 264);
			this.tlpMain.TabIndex = 0;
			// 
			// btnExecProtectionAction
			// 
			this.btnExecProtectionAction.Dock = System.Windows.Forms.DockStyle.Right;
			this.btnExecProtectionAction.Location = new System.Drawing.Point(406, 229);
			this.btnExecProtectionAction.Name = "btnExecProtectionAction";
			this.btnExecProtectionAction.Size = new System.Drawing.Size(128, 32);
			this.btnExecProtectionAction.TabIndex = 0;
			this.btnExecProtectionAction.Text = "Set / clear";
			this.btnExecProtectionAction.UseVisualStyleBackColor = true;
			// 
			// radioModeProtect
			// 
			this.radioModeProtect.AutoSize = true;
			this.radioModeProtect.Checked = true;
			this.tlpMain.SetColumnSpan(this.radioModeProtect, 2);
			this.radioModeProtect.Dock = System.Windows.Forms.DockStyle.Top;
			this.radioModeProtect.Location = new System.Drawing.Point(3, 3);
			this.radioModeProtect.Name = "radioModeProtect";
			this.radioModeProtect.Size = new System.Drawing.Size(531, 17);
			this.radioModeProtect.TabIndex = 1;
			this.radioModeProtect.TabStop = true;
			this.radioModeProtect.Text = "set";
			this.radioModeProtect.UseVisualStyleBackColor = true;
			// 
			// lblModeDesription
			// 
			this.lblModeDesription.AutoSize = true;
			this.tlpMain.SetColumnSpan(this.lblModeDesription, 2);
			this.lblModeDesription.Location = new System.Drawing.Point(3, 46);
			this.lblModeDesription.Name = "lblModeDesription";
			this.lblModeDesription.Size = new System.Drawing.Size(106, 13);
			this.lblModeDesription.TabIndex = 3;
			this.lblModeDesription.Text = "protection for sheets:";
			// 
			// radioModeUnProtect
			// 
			this.radioModeUnProtect.AutoSize = true;
			this.tlpMain.SetColumnSpan(this.radioModeUnProtect, 2);
			this.radioModeUnProtect.Dock = System.Windows.Forms.DockStyle.Top;
			this.radioModeUnProtect.Location = new System.Drawing.Point(3, 26);
			this.radioModeUnProtect.Name = "radioModeUnProtect";
			this.radioModeUnProtect.Size = new System.Drawing.Size(531, 17);
			this.radioModeUnProtect.TabIndex = 2;
			this.radioModeUnProtect.TabStop = true;
			this.radioModeUnProtect.Text = "reset";
			this.radioModeUnProtect.UseVisualStyleBackColor = true;
			// 
			// lstWorksheets
			// 
			this.lstWorksheets.CheckOnClick = true;
			this.tlpMain.SetColumnSpan(this.lstWorksheets, 2);
			this.lstWorksheets.Dock = System.Windows.Forms.DockStyle.Fill;
			this.lstWorksheets.FormattingEnabled = true;
			this.lstWorksheets.IntegralHeight = false;
			this.lstWorksheets.Location = new System.Drawing.Point(3, 62);
			this.lstWorksheets.Name = "lstWorksheets";
			this.lstWorksheets.Size = new System.Drawing.Size(531, 123);
			this.lstWorksheets.TabIndex = 4;
			// 
			// lblPWD
			// 
			this.lblPWD.AutoSize = true;
			this.lblPWD.Dock = System.Windows.Forms.DockStyle.Right;
			this.lblPWD.Location = new System.Drawing.Point(3, 192);
			this.lblPWD.Name = "lblPWD";
			this.lblPWD.Size = new System.Drawing.Size(27, 26);
			this.lblPWD.TabIndex = 5;
			this.lblPWD.Text = "pwd";
			this.lblPWD.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// txtPWD
			// 
			this.txtPWD.Dock = System.Windows.Forms.DockStyle.Top;
			this.txtPWD.HideSelection = false;
			this.txtPWD.Location = new System.Drawing.Point(36, 195);
			this.txtPWD.Name = "txtPWD";
			this.txtPWD.Size = new System.Drawing.Size(498, 20);
			this.txtPWD.TabIndex = 6;
			this.txtPWD.Text = "tttttttttttttttt";
			this.txtPWD.UseSystemPasswordChar = true;
			this.txtPWD.WordWrap = false;
			// 
			// panel1
			// 
			this.panel1.BackColor = System.Drawing.SystemColors.ControlDarkDark;
			this.tlpMain.SetColumnSpan(this.panel1, 2);
			this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.panel1.Location = new System.Drawing.Point(3, 191);
			this.panel1.Name = "panel1";
			this.panel1.Size = new System.Drawing.Size(531, 1);
			this.panel1.TabIndex = 7;
			// 
			// frmAskWorksheetProtectionPassword
			// 
			this.AcceptButton = this.btnExecProtectionAction;
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(553, 280);
			this.Controls.Add(this.tlpMain);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmAskWorksheetProtectionPassword";
			this.Padding = new System.Windows.Forms.Padding(8);
			this.ShowIcon = false;
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
			this.Text = "frmAskWorksheetProtectionPassword";
			this.tlpMain.ResumeLayout(false);
			this.tlpMain.PerformLayout();
			this.ResumeLayout(false);

		}

		#endregion

		private System.Windows.Forms.TableLayoutPanel tlpMain;
		internal System.Windows.Forms.Button btnExecProtectionAction;
		internal System.Windows.Forms.RadioButton radioModeProtect;
		internal System.Windows.Forms.RadioButton radioModeUnProtect;
		internal System.Windows.Forms.Label lblModeDesription;
		internal NavfertyExcelAddIn.Commons.Controls.CheckedListBoxEx lstWorksheets;
		internal System.Windows.Forms.Label lblPWD;
		internal System.Windows.Forms.TextBox txtPWD;
		private System.Windows.Forms.Panel panel1;
	}
}
