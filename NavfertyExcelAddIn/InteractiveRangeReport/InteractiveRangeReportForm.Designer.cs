using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

namespace NavfertyExcelAddIn.InteractiveRangeReport
{
    partial class InteractiveRangeReportForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private IContainer components = null;

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
            ErrorMessage = new DataGridViewTextBoxColumn();
            Value = new DataGridViewTextBoxColumn();
            Address = new DataGridViewTextBoxColumn();
            WorksheetName = new DataGridViewTextBoxColumn();
            RangesGridView = new DataGridView();
            ((ISupportInitialize)(RangesGridView)).BeginInit();
            SuspendLayout();

            // 
            // ErrorMessage
            // 
            ErrorMessage.DataPropertyName = "ErrorMessage";
            ErrorMessage.Name = "ErrorMessage";
            ErrorMessage.ReadOnly = true;

            // 
            // Value
            // 
            Value.DataPropertyName = "Value";
            Value.Name = "Value";
            Value.ReadOnly = true;
            Value.Width = 250;

            // 
            // Address
            // 
            Address.DataPropertyName = "Address";
            Address.Name = "Address";
            Address.ReadOnly = true;

            // 
            // WorksheetName
            // 
            WorksheetName.DataPropertyName = "WorksheetName";
            WorksheetName.Name = "WorksheetName";
            WorksheetName.ReadOnly = true;

            // 
            // RangesGridView
            // 
            RangesGridView.AllowUserToAddRows = false;
            RangesGridView.AllowUserToDeleteRows = false;
            RangesGridView.AllowUserToOrderColumns = true;
            RangesGridView.Anchor =
                ((AnchorStyles)((((AnchorStyles.Top | AnchorStyles.Bottom) | AnchorStyles.Left) | AnchorStyles.Right)));
            RangesGridView.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            RangesGridView.Columns.AddRange(new DataGridViewColumn[] { ErrorMessage, Value, Address, WorksheetName });
            RangesGridView.Location = new System.Drawing.Point(12, 12);
            RangesGridView.Name = "RangesGridView";
            RangesGridView.ReadOnly = true;
            RangesGridView.Size = new System.Drawing.Size(776, 426);
            RangesGridView.TabIndex = 0;

            // 
            // SearchRangeResultForm
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new System.Drawing.Size(800, 450);
            Controls.Add(RangesGridView);

            var resources = new ComponentResourceManager(typeof(InteractiveRangeReportForm));
            Icon = (Icon)resources.GetObject("ExcelIcon");

            Name = "SearchRangeResultForm";
            TopMost = false;
            ((ISupportInitialize)(RangesGridView)).EndInit();
            ResumeLayout(false);

        }
        #endregion

        private DataGridView RangesGridView;
        private DataGridViewTextBoxColumn ErrorMessage;
        private DataGridViewTextBoxColumn Value;
        private DataGridViewTextBoxColumn Address;
        private DataGridViewTextBoxColumn WorksheetName;
    }
}