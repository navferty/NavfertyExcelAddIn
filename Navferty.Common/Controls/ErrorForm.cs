using System;
using System.Drawing;

using Navferty.Common.Feedback;

#nullable enable

namespace Navferty.Common.Controls
{
	internal class ErrorForm : FormEx
	{
		private System.Windows.Forms.Button btnOk;
		private System.Windows.Forms.Label lblMessage;
		private System.Windows.Forms.PictureBox picIcon;
		private System.Windows.Forms.TableLayoutPanel tlpMessage;
		private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
		private System.Windows.Forms.LinkLabel llSendErrorReport;
		private System.Windows.Forms.Label lblHeader;
		private System.Windows.Forms.TableLayoutPanel tlpMain;

		internal ErrorForm() : base()
		{
			InitializeComponent();
			picIcon!.Image = SystemIcons.Exclamation.ToBitmap();
			lblMessage!.Text = string.Empty;
			lblMessage!.BackColor = SystemColors.Window;
			llSendErrorReport!.Text = Localization.UIStrings.Feedback_SendFeedback;
			btnOk!.Text = Localization.UIStrings.ErrorWindow_OkButton;
		}
		private void InitializeComponent()
		{
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ErrorForm));
            this.tlpMain = new System.Windows.Forms.TableLayoutPanel();
            this.lblHeader = new System.Windows.Forms.Label();
            this.tlpMessage = new System.Windows.Forms.TableLayoutPanel();
            this.lblMessage = new System.Windows.Forms.Label();
            this.picIcon = new System.Windows.Forms.PictureBox();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.btnOk = new System.Windows.Forms.Button();
            this.llSendErrorReport = new System.Windows.Forms.LinkLabel();
            this.tlpMain.SuspendLayout();
            this.tlpMessage.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.picIcon)).BeginInit();
            this.tableLayoutPanel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tlpMain
            // 
            this.tlpMain.AutoSize = true;
            this.tlpMain.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.tlpMain.ColumnCount = 1;
            this.tlpMain.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tlpMain.Controls.Add(this.lblHeader, 0, 0);
            this.tlpMain.Controls.Add(this.tlpMessage, 0, 1);
            this.tlpMain.Controls.Add(this.tableLayoutPanel1, 0, 2);
            this.tlpMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tlpMain.Location = new System.Drawing.Point(0, 0);
            this.tlpMain.Name = "tlpMain";
            this.tlpMain.RowCount = 3;
            this.tlpMain.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tlpMain.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tlpMain.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tlpMain.Size = new System.Drawing.Size(384, 161);
            this.tlpMain.TabIndex = 0;
            // 
            // lblHeader
            // 
            this.lblHeader.AutoSize = true;
            this.lblHeader.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblHeader.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblHeader.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.lblHeader.Location = new System.Drawing.Point(3, 0);
            this.lblHeader.Name = "lblHeader";
            this.lblHeader.Padding = new System.Windows.Forms.Padding(8);
            this.lblHeader.Size = new System.Drawing.Size(378, 35);
            this.lblHeader.TabIndex = 5;
            this.lblHeader.Text = "label1";
            this.lblHeader.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // tlpMessage
            // 
            this.tlpMessage.AutoSize = true;
            this.tlpMessage.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.tlpMessage.BackColor = System.Drawing.SystemColors.Window;
            this.tlpMessage.ColumnCount = 3;
            this.tlpMessage.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 8F));
            this.tlpMessage.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tlpMessage.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tlpMessage.Controls.Add(this.lblMessage, 2, 0);
            this.tlpMessage.Controls.Add(this.picIcon, 1, 0);
            this.tlpMessage.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tlpMessage.Location = new System.Drawing.Point(3, 38);
            this.tlpMessage.Name = "tlpMessage";
            this.tlpMessage.RowCount = 1;
            this.tlpMessage.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tlpMessage.Size = new System.Drawing.Size(378, 66);
            this.tlpMessage.TabIndex = 3;
            this.tlpMessage.Paint += new System.Windows.Forms.PaintEventHandler(this.tlpMessage_Paint);
            // 
            // lblMessage
            // 
            this.lblMessage.AutoSize = true;
            this.lblMessage.BackColor = System.Drawing.SystemColors.Info;
            this.lblMessage.Dock = System.Windows.Forms.DockStyle.Left;
            this.lblMessage.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblMessage.Location = new System.Drawing.Point(47, 0);
            this.lblMessage.Name = "lblMessage";
            this.lblMessage.Padding = new System.Windows.Forms.Padding(8);
            this.lblMessage.Size = new System.Drawing.Size(68, 66);
            this.lblMessage.TabIndex = 1;
            this.lblMessage.Text = "label1";
            this.lblMessage.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // picIcon
            // 
            this.picIcon.Dock = System.Windows.Forms.DockStyle.Left;
            this.picIcon.Image = ((System.Drawing.Image)(resources.GetObject("picIcon.Image")));
            this.picIcon.Location = new System.Drawing.Point(11, 3);
            this.picIcon.Name = "picIcon";
            this.picIcon.Size = new System.Drawing.Size(30, 60);
            this.picIcon.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.picIcon.TabIndex = 2;
            this.picIcon.TabStop = false;
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 2;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tableLayoutPanel1.Controls.Add(this.btnOk, 1, 0);
            this.tableLayoutPanel1.Controls.Add(this.llSendErrorReport, 0, 0);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(3, 110);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.Padding = new System.Windows.Forms.Padding(8);
            this.tableLayoutPanel1.RowCount = 1;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(378, 48);
            this.tableLayoutPanel1.TabIndex = 4;
            // 
            // btnOk
            // 
            this.btnOk.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btnOk.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnOk.Location = new System.Drawing.Point(285, 11);
            this.btnOk.Name = "btnOk";
            this.btnOk.Size = new System.Drawing.Size(82, 26);
            this.btnOk.TabIndex = 0;
            this.btnOk.Text = "Close";
            this.btnOk.UseVisualStyleBackColor = true;
            // 
            // llSendErrorReport
            // 
            this.llSendErrorReport.AutoSize = true;
            this.llSendErrorReport.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.llSendErrorReport.Location = new System.Drawing.Point(11, 27);
            this.llSendErrorReport.Name = "llSendErrorReport";
            this.llSendErrorReport.Size = new System.Drawing.Size(268, 13);
            this.llSendErrorReport.TabIndex = 1;
            this.llSendErrorReport.TabStop = true;
            this.llSendErrorReport.Text = "Send error report";
            this.llSendErrorReport.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.llSendErrorReport_LinkClicked);
            // 
            // ErrorForm
            // 
            this.AcceptButton = this.btnOk;
            this.AutoSize = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.ClientSize = new System.Drawing.Size(384, 161);
            this.Controls.Add(this.tlpMain);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(400, 200);
            this.Name = "ErrorForm";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.tlpMain.ResumeLayout(false);
            this.tlpMain.PerformLayout();
            this.tlpMessage.ResumeLayout(false);
            this.tlpMessage.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.picIcon)).EndInit();
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

		}

		private readonly Exception? Err = null;
		private readonly bool AllowErrorReporting = false;

		internal ErrorForm(Exception ex, string? title, bool allowErrorReporting = true) : this()
		{
			Err = ex;
			AllowErrorReporting = allowErrorReporting && (null != Err);
			Text = Localization.UIStrings.ErrorWindow_Title;
			bool HasTitle = !string.IsNullOrWhiteSpace(title);
			{
				lblHeader.Visible = HasTitle;
				if (HasTitle) lblHeader.Text = title;
			}
			lblMessage.Text = (null == Err) ? "Unknown Error!" : ex.Message;
			llSendErrorReport.Visible = allowErrorReporting;
		}

		private void llSendErrorReport_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
		{
			if (!AllowErrorReporting) return;
			FeedbackManager.ShowFeedbackUI(Err?.Message);
		}

		private void tlpMessage_Paint(object sender, System.Windows.Forms.PaintEventArgs e)
		{

		}
	}
}
