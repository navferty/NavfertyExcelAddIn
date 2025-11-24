using System;
using System.Windows.Forms;

using Navferty.Common.Controls;

using NavfertyExcelAddIn.Localization;

namespace NavfertyExcelAddIn.SqliteExport;

internal partial class SqliteExportOptionsForm : FormEx
{
	public SqliteExportOptions Options { get; private set; }

	public SqliteExportOptionsForm()
	{
		InitializeComponent();
		Options = new SqliteExportOptions();
	}

	private void InitializeComponent()
	{
		this.checkBoxUseFirstRow = new System.Windows.Forms.CheckBox();
		this.numericUpDownSkipRows = new System.Windows.Forms.NumericUpDown();
		this.labelSkipRows = new System.Windows.Forms.Label();
		this.buttonOK = new System.Windows.Forms.Button();
		this.buttonCancel = new System.Windows.Forms.Button();
		((System.ComponentModel.ISupportInitialize)(this.numericUpDownSkipRows)).BeginInit();
		this.SuspendLayout();
		// 
		// checkBoxUseFirstRow
		// 
		this.checkBoxUseFirstRow.AutoSize = true;
		this.checkBoxUseFirstRow.Checked = true;
		this.checkBoxUseFirstRow.CheckState = System.Windows.Forms.CheckState.Checked;
		this.checkBoxUseFirstRow.Location = new System.Drawing.Point(12, 12);
		this.checkBoxUseFirstRow.Name = "checkBoxUseFirstRow";
		this.checkBoxUseFirstRow.Size = new System.Drawing.Size(200, 17);
		this.checkBoxUseFirstRow.TabIndex = 0;
		this.checkBoxUseFirstRow.Text = UIStrings.SqliteExport_UseFirstRowAsHeaders;
		this.checkBoxUseFirstRow.UseVisualStyleBackColor = true;
		// 
		// labelSkipRows
		// 
		this.labelSkipRows.AutoSize = true;
		this.labelSkipRows.Location = new System.Drawing.Point(12, 45);
		this.labelSkipRows.Name = "labelSkipRows";
		this.labelSkipRows.Size = new System.Drawing.Size(150, 13);
		this.labelSkipRows.TabIndex = 1;
		this.labelSkipRows.Text = UIStrings.SqliteExport_RowsToSkip;
		// 
		// numericUpDownSkipRows
		// 
		this.numericUpDownSkipRows.Location = new System.Drawing.Point(170, 43);
		this.numericUpDownSkipRows.Maximum = new decimal(new int[] {
            1000,
            0,
            0,
            0});
		this.numericUpDownSkipRows.Name = "numericUpDownSkipRows";
		this.numericUpDownSkipRows.Size = new System.Drawing.Size(100, 20);
		this.numericUpDownSkipRows.TabIndex = 2;
		// 
		// buttonOK
		// 
		this.buttonOK.DialogResult = System.Windows.Forms.DialogResult.OK;
		this.buttonOK.Location = new System.Drawing.Point(114, 85);
		this.buttonOK.Name = "buttonOK";
		this.buttonOK.Size = new System.Drawing.Size(75, 23);
		this.buttonOK.TabIndex = 3;
		this.buttonOK.Text = UIStrings.OK;
		this.buttonOK.UseVisualStyleBackColor = true;
		this.buttonOK.Click += new System.EventHandler(this.ButtonOK_Click);
		// 
		// buttonCancel
		// 
		this.buttonCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
		this.buttonCancel.Location = new System.Drawing.Point(195, 85);
		this.buttonCancel.Name = "buttonCancel";
		this.buttonCancel.Size = new System.Drawing.Size(75, 23);
		this.buttonCancel.TabIndex = 4;
		this.buttonCancel.Text = UIStrings.Cancel;
		this.buttonCancel.UseVisualStyleBackColor = true;
		// 
		// SqliteExportOptionsForm
		// 
		this.AcceptButton = this.buttonOK;
		this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
		this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
		this.CancelButton = this.buttonCancel;
		this.ClientSize = new System.Drawing.Size(284, 121);
		this.Controls.Add(this.buttonCancel);
		this.Controls.Add(this.buttonOK);
		this.Controls.Add(this.numericUpDownSkipRows);
		this.Controls.Add(this.labelSkipRows);
		this.Controls.Add(this.checkBoxUseFirstRow);
		this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
		this.MaximizeBox = false;
		this.MinimizeBox = false;
		this.Name = "SqliteExportOptionsForm";
		this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
		this.Text = UIStrings.SqliteExport_OptionsTitle;
		((System.ComponentModel.ISupportInitialize)(this.numericUpDownSkipRows)).EndInit();
		this.ResumeLayout(false);
		this.PerformLayout();

	}

	private System.Windows.Forms.CheckBox checkBoxUseFirstRow = null!;
	private System.Windows.Forms.NumericUpDown numericUpDownSkipRows = null!;
	private System.Windows.Forms.Label labelSkipRows = null!;
	private System.Windows.Forms.Button buttonOK = null!;
	private System.Windows.Forms.Button buttonCancel = null!;

	private void ButtonOK_Click(object sender, EventArgs e)
	{
		Options.UseFirstRowAsHeaders = checkBoxUseFirstRow.Checked;
		Options.RowsToSkip = (int)numericUpDownSkipRows.Value;
	}
}
