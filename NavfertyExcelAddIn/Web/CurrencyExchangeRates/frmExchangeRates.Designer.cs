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
            this.btnPasteResult = new System.Windows.Forms.Button();
            this.dtpDate = new System.Windows.Forms.DateTimePicker();
            this.tlpMain = new System.Windows.Forms.TableLayoutPanel();
            this.gridResult = new NavfertyExcelAddIn.Commons.Controls.DataGridViewEx();
            this.txtFilter = new System.Windows.Forms.TextBox();
            this.cbProvider = new System.Windows.Forms.ComboBox();
            this.lblSource = new System.Windows.Forms.Label();
            this.lblFilterTitle = new System.Windows.Forms.Label();
            this.tlpMain.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gridResult)).BeginInit();
            this.SuspendLayout();
            // 
            // btnPasteResult
            // 
            this.btnPasteResult.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btnPasteResult.Location = new System.Drawing.Point(496, 498);
            this.btnPasteResult.Name = "btnPasteResult";
            this.btnPasteResult.Size = new System.Drawing.Size(174, 34);
            this.btnPasteResult.TabIndex = 4;
            this.btnPasteResult.Text = "Select";
            this.btnPasteResult.UseVisualStyleBackColor = true;
            this.btnPasteResult.Click += new System.EventHandler(this.btnPasteResult_Click);
            // 
            // dtpDate
            // 
            this.dtpDate.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dtpDate.Location = new System.Drawing.Point(496, 3);
            this.dtpDate.Name = "dtpDate";
            this.dtpDate.Size = new System.Drawing.Size(174, 20);
            this.dtpDate.TabIndex = 2;
            // 
            // tlpMain
            // 
            this.tlpMain.ColumnCount = 3;
            this.tlpMain.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tlpMain.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tlpMain.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 180F));
            this.tlpMain.Controls.Add(this.dtpDate, 2, 0);
            this.tlpMain.Controls.Add(this.gridResult, 0, 2);
            this.tlpMain.Controls.Add(this.btnPasteResult, 2, 3);
            this.tlpMain.Controls.Add(this.txtFilter, 1, 1);
            this.tlpMain.Controls.Add(this.cbProvider, 1, 0);
            this.tlpMain.Controls.Add(this.lblSource, 0, 0);
            this.tlpMain.Controls.Add(this.lblFilterTitle, 0, 1);
            this.tlpMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tlpMain.Location = new System.Drawing.Point(8, 8);
            this.tlpMain.Name = "tlpMain";
            this.tlpMain.RowCount = 4;
            this.tlpMain.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tlpMain.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tlpMain.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tlpMain.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 40F));
            this.tlpMain.Size = new System.Drawing.Size(673, 535);
            this.tlpMain.TabIndex = 0;
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
            this.tlpMain.SetColumnSpan(this.gridResult, 3);
            this.gridResult.Dock = System.Windows.Forms.DockStyle.Fill;
            this.gridResult.Location = new System.Drawing.Point(3, 56);
            this.gridResult.MultiSelect = false;
            this.gridResult.Name = "gridResult";
            this.gridResult.ReadOnly = true;
            this.gridResult.RowHeadersVisible = false;
            this.gridResult.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;
            this.gridResult.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.gridResult.Size = new System.Drawing.Size(667, 436);
            this.gridResult.StandardEnter = true;
            this.gridResult.StandardTab = true;
            this.gridResult.TabIndex = 0;
            this.gridResult.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.gridResult_CellDoubleClick);
            // 
            // txtFilter
            // 
            this.txtFilter.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.tlpMain.SetColumnSpan(this.txtFilter, 2);
            this.txtFilter.Dock = System.Windows.Forms.DockStyle.Top;
            this.txtFilter.HideSelection = false;
            this.txtFilter.Location = new System.Drawing.Point(48, 30);
            this.txtFilter.MaxLength = 100;
            this.txtFilter.Name = "txtFilter";
            this.txtFilter.ShortcutsEnabled = false;
            this.txtFilter.Size = new System.Drawing.Size(622, 20);
            this.txtFilter.TabIndex = 3;
            // 
            // cbProvider
            // 
            this.cbProvider.Dock = System.Windows.Forms.DockStyle.Top;
            this.cbProvider.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbProvider.FormattingEnabled = true;
            this.cbProvider.Location = new System.Drawing.Point(48, 3);
            this.cbProvider.Name = "cbProvider";
            this.cbProvider.Size = new System.Drawing.Size(442, 21);
            this.cbProvider.TabIndex = 1;
            // 
            // lblSource
            // 
            this.lblSource.AutoSize = true;
            this.lblSource.Dock = System.Windows.Forms.DockStyle.Left;
            this.lblSource.Location = new System.Drawing.Point(3, 0);
            this.lblSource.Name = "lblSource";
            this.lblSource.Size = new System.Drawing.Size(39, 27);
            this.lblSource.TabIndex = 5;
            this.lblSource.Text = "source";
            this.lblSource.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblFilterTitle
            // 
            this.lblFilterTitle.AutoSize = true;
            this.lblFilterTitle.Dock = System.Windows.Forms.DockStyle.Right;
            this.lblFilterTitle.Location = new System.Drawing.Point(24, 27);
            this.lblFilterTitle.Name = "lblFilterTitle";
            this.lblFilterTitle.Size = new System.Drawing.Size(18, 26);
            this.lblFilterTitle.TabIndex = 6;
            this.lblFilterTitle.Text = "fltr";
            this.lblFilterTitle.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // frmExchangeRates
            // 
            this.AcceptButton = this.btnPasteResult;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(689, 551);
            this.Controls.Add(this.tlpMain);
            this.KeyPreview = true;
            this.MinimizeBox = false;
            this.Name = "frmExchangeRates";
            this.Padding = new System.Windows.Forms.Padding(8);
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "ExchangeRates";
            this.Load += new System.EventHandler(this.Form_Load);
            this.Shown += new System.EventHandler(this.Form_Displayed);
            this.tlpMain.ResumeLayout(false);
            this.tlpMain.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gridResult)).EndInit();
            this.ResumeLayout(false);

		}

		#endregion

		private System.Windows.Forms.Button btnPasteResult;
		private System.Windows.Forms.DateTimePicker dtpDate;
		private System.Windows.Forms.TableLayoutPanel tlpMain;
		private System.Windows.Forms.TextBox txtFilter;
		private System.Windows.Forms.ComboBox cbProvider;
		private Commons.Controls.DataGridViewEx gridResult;
		private System.Windows.Forms.Label lblSource;
		private System.Windows.Forms.Label lblFilterTitle;
	}
}
