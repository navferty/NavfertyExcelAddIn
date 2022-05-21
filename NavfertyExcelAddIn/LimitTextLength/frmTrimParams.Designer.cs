namespace NavfertyExcelAddIn.LimitTextLength
{
	partial class frmTrimParams
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
			this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
			this.btnOk = new System.Windows.Forms.Button();
			this.lblTextLength = new System.Windows.Forms.Label();
			this.numMaxLength = new System.Windows.Forms.NumericUpDown();
			this.chkTrimStartEnd = new System.Windows.Forms.CheckBox();
			this.chkTrimFullSpaces = new System.Windows.Forms.CheckBox();
			this.tableLayoutPanel1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.numMaxLength)).BeginInit();
			this.SuspendLayout();
			// 
			// tableLayoutPanel1
			// 
			this.tableLayoutPanel1.ColumnCount = 2;
			this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
			this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
			this.tableLayoutPanel1.Controls.Add(this.btnOk, 1, 4);
			this.tableLayoutPanel1.Controls.Add(this.lblTextLength, 0, 0);
			this.tableLayoutPanel1.Controls.Add(this.numMaxLength, 1, 0);
			this.tableLayoutPanel1.Controls.Add(this.chkTrimStartEnd, 0, 1);
			this.tableLayoutPanel1.Controls.Add(this.chkTrimFullSpaces, 0, 2);
			this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.tableLayoutPanel1.Location = new System.Drawing.Point(8, 8);
			this.tableLayoutPanel1.Name = "tableLayoutPanel1";
			this.tableLayoutPanel1.RowCount = 5;
			this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
			this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
			this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
			this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
			this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 32F));
			this.tableLayoutPanel1.Size = new System.Drawing.Size(288, 136);
			this.tableLayoutPanel1.TabIndex = 0;
			// 
			// btnOk
			// 
			this.btnOk.DialogResult = System.Windows.Forms.DialogResult.OK;
			this.btnOk.Dock = System.Windows.Forms.DockStyle.Right;
			this.btnOk.Location = new System.Drawing.Point(179, 107);
			this.btnOk.Name = "btnOk";
			this.btnOk.Size = new System.Drawing.Size(106, 26);
			this.btnOk.TabIndex = 0;
			this.btnOk.Text = "Ok";
			this.btnOk.UseVisualStyleBackColor = true;
			// 
			// lblTextLength
			// 
			this.lblTextLength.AutoSize = true;
			this.lblTextLength.Dock = System.Windows.Forms.DockStyle.Fill;
			this.lblTextLength.Location = new System.Drawing.Point(3, 0);
			this.lblTextLength.Name = "lblTextLength";
			this.lblTextLength.Size = new System.Drawing.Size(170, 26);
			this.lblTextLength.TabIndex = 1;
			this.lblTextLength.Text = "Maximum Text Length:";
			this.lblTextLength.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// numMaxLength
			// 
			this.numMaxLength.Dock = System.Windows.Forms.DockStyle.Top;
			this.numMaxLength.Location = new System.Drawing.Point(179, 3);
			this.numMaxLength.Name = "numMaxLength";
			this.numMaxLength.Size = new System.Drawing.Size(106, 20);
			this.numMaxLength.TabIndex = 2;
			this.numMaxLength.ThousandsSeparator = true;
			// 
			// chkTrimStartEnd
			// 
			this.chkTrimStartEnd.AutoSize = true;
			this.tableLayoutPanel1.SetColumnSpan(this.chkTrimStartEnd, 2);
			this.chkTrimStartEnd.Dock = System.Windows.Forms.DockStyle.Top;
			this.chkTrimStartEnd.Location = new System.Drawing.Point(3, 29);
			this.chkTrimStartEnd.Name = "chkTrimStartEnd";
			this.chkTrimStartEnd.Size = new System.Drawing.Size(282, 14);
			this.chkTrimStartEnd.TabIndex = 3;
			this.chkTrimStartEnd.Text = "trimStartEndSpaces";
			this.chkTrimStartEnd.UseVisualStyleBackColor = true;
			// 
			// chkTrimFullSpaces
			// 
			this.chkTrimFullSpaces.AutoSize = true;
			this.tableLayoutPanel1.SetColumnSpan(this.chkTrimFullSpaces, 2);
			this.chkTrimFullSpaces.Dock = System.Windows.Forms.DockStyle.Top;
			this.chkTrimFullSpaces.Location = new System.Drawing.Point(3, 49);
			this.chkTrimFullSpaces.Name = "chkTrimFullSpaces";
			this.chkTrimFullSpaces.Size = new System.Drawing.Size(282, 14);
			this.chkTrimFullSpaces.TabIndex = 4;
			this.chkTrimFullSpaces.Text = "trimFullSpacedStrings";
			this.chkTrimFullSpaces.UseVisualStyleBackColor = true;
			// 
			// frmTrimParams
			// 
			this.AcceptButton = this.btnOk;
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(304, 152);
			this.Controls.Add(this.tableLayoutPanel1);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmTrimParams";
			this.Padding = new System.Windows.Forms.Padding(8);
			this.ShowIcon = false;
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
			this.Text = "Trim Params";
			this.tableLayoutPanel1.ResumeLayout(false);
			this.tableLayoutPanel1.PerformLayout();
			((System.ComponentModel.ISupportInitialize)(this.numMaxLength)).EndInit();
			this.ResumeLayout(false);

		}

		#endregion

		private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
		private System.Windows.Forms.Button btnOk;
		private System.Windows.Forms.Label lblTextLength;
		internal System.Windows.Forms.NumericUpDown numMaxLength;
		internal System.Windows.Forms.CheckBox chkTrimStartEnd;
		internal System.Windows.Forms.CheckBox chkTrimFullSpaces;
	}
}
