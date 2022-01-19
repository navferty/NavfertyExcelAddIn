namespace NavfertyExcelAddIn.Web.CurrencyExchangeRates
{
	partial class frmExchangeRates
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
			this.btnApplyResult = new System.Windows.Forms.Button();
			this.dtpDate = new System.Windows.Forms.DateTimePicker();
			this.gridResult = new System.Windows.Forms.DataGridView();
			this.tlpMain = new System.Windows.Forms.TableLayoutPanel();
			this.txtFilter = new System.Windows.Forms.TextBox();
			((System.ComponentModel.ISupportInitialize)(this.gridResult)).BeginInit();
			this.tlpMain.SuspendLayout();
			this.SuspendLayout();
			// 
			// btnGet
			// 
			this.btnApplyResult.Dock = System.Windows.Forms.DockStyle.Right;
			this.btnApplyResult.Location = new System.Drawing.Point(388, 618);
			this.btnApplyResult.Name = "btnGet";
			this.btnApplyResult.Size = new System.Drawing.Size(140, 42);
			this.btnApplyResult.TabIndex = 3;
			this.btnApplyResult.Text = "Выбрать";
			this.btnApplyResult.UseVisualStyleBackColor = true;
			this.btnApplyResult.Click += new System.EventHandler(this.btnApplyResult_Click);
			// 
			// dtpDate
			// 
			this.dtpDate.Dock = System.Windows.Forms.DockStyle.Left;
			this.dtpDate.Location = new System.Drawing.Point(3, 3);
			this.dtpDate.Name = "dtpDate";
			this.dtpDate.Size = new System.Drawing.Size(219, 20);
			this.dtpDate.TabIndex = 0;
			// 
			// gridResult
			// 
			this.gridResult.AllowUserToAddRows = false;
			this.gridResult.AllowUserToDeleteRows = false;
			this.gridResult.AllowUserToOrderColumns = true;
			this.gridResult.AllowUserToResizeColumns = false;
			this.gridResult.AllowUserToResizeRows = false;
			this.gridResult.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
			this.gridResult.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
			this.gridResult.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			this.gridResult.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
			this.gridResult.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
			this.tlpMain.SetColumnSpan(this.gridResult, 2);
			this.gridResult.Dock = System.Windows.Forms.DockStyle.Fill;
			this.gridResult.Location = new System.Drawing.Point(3, 29);
			this.gridResult.MultiSelect = false;
			this.gridResult.Name = "gridResult";
			this.gridResult.ReadOnly = true;
			this.gridResult.RowHeadersVisible = false;
			this.gridResult.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;
			this.gridResult.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
			this.gridResult.Size = new System.Drawing.Size(525, 583);
			this.gridResult.StandardTab = true;
			this.gridResult.TabIndex = 2;
			// 
			// tlpMain
			// 
			this.tlpMain.ColumnCount = 2;
			this.tlpMain.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
			this.tlpMain.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
			this.tlpMain.Controls.Add(this.dtpDate, 0, 0);
			this.tlpMain.Controls.Add(this.gridResult, 0, 1);
			this.tlpMain.Controls.Add(this.btnApplyResult, 1, 2);
			this.tlpMain.Controls.Add(this.txtFilter, 1, 0);
			this.tlpMain.Dock = System.Windows.Forms.DockStyle.Fill;
			this.tlpMain.Location = new System.Drawing.Point(8, 8);
			this.tlpMain.Name = "tlpMain";
			this.tlpMain.RowCount = 3;
			this.tlpMain.RowStyles.Add(new System.Windows.Forms.RowStyle());
			this.tlpMain.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
			this.tlpMain.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 48F));
			this.tlpMain.Size = new System.Drawing.Size(531, 663);
			this.tlpMain.TabIndex = 3;
			// 
			// txtFilter
			// 
			this.txtFilter.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtFilter.Dock = System.Windows.Forms.DockStyle.Top;
			this.txtFilter.HideSelection = false;
			this.txtFilter.Location = new System.Drawing.Point(228, 3);
			this.txtFilter.MaxLength = 100;
			this.txtFilter.Name = "txtFilter";
			this.txtFilter.ShortcutsEnabled = false;
			this.txtFilter.Size = new System.Drawing.Size(300, 20);
			this.txtFilter.TabIndex = 1;
			// 
			// frmExchangeRates
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(547, 679);
			this.Controls.Add(this.tlpMain);
			this.MinimizeBox = false;
			this.Name = "frmExchangeRates";
			this.Padding = new System.Windows.Forms.Padding(8);
			this.ShowIcon = false;
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
			this.Text = "ExchangeRates";
			this.Load += new System.EventHandler(this.Form_Load);
			this.Shown += new System.EventHandler(this.Form_Displayed);
			((System.ComponentModel.ISupportInitialize)(this.gridResult)).EndInit();
			this.tlpMain.ResumeLayout(false);
			this.tlpMain.PerformLayout();
			this.ResumeLayout(false);

		}

		#endregion

		private System.Windows.Forms.Button btnApplyResult;
		private System.Windows.Forms.DateTimePicker dtpDate;
		private System.Windows.Forms.DataGridView gridResult;
		private System.Windows.Forms.TableLayoutPanel tlpMain;
		private System.Windows.Forms.TextBox txtFilter;
	}
}
