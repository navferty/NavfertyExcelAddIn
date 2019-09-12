using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;

using Microsoft.Office.Interop.Excel;

using Application = Microsoft.Office.Interop.Excel.Application;

namespace NavfertyExcelAddIn.InteractiveRangeReport
{
    // TODO rename InteractiveRangeReportForm
    public partial class InteractiveRangeReportForm : Form
    {
        private readonly InteractiveErrorItem[] errorItems;
        private readonly Worksheet worksheet;

        private Application App => Globals.ThisAddIn.Application;

        public InteractiveRangeReportForm(IReadOnlyCollection<InteractiveErrorItem> errorItems, Worksheet worksheet)
        {
            this.errorItems = errorItems.ToArray();
            this.worksheet = worksheet;

            InitializeComponent();

            ErrorMessage.HeaderText = Localization.UIStrings.Error;
            Value.HeaderText = Localization.UIStrings.Formula;
            Address.HeaderText = Localization.UIStrings.Address;
            WorksheetName.HeaderText = Localization.UIStrings.WsName;
            Text = Localization.UIStrings.SearchResults;

            RangesGridView.DataSource = this.errorItems;
            RangesGridView.SelectionChanged += OnSelectionChanged;
            RangesGridView.AutoGenerateColumns = false;

            worksheet.BeforeDelete += () => Close();
            ((Workbook)worksheet.Parent).BeforeClose += (ref bool cancel) => Close();
        }

        private void OnSelectionChanged(object sender, EventArgs e)
        {
            var name = worksheet.Name;
            var selectedRows = RangesGridView.SelectedCells.Cast<DataGridViewCell>().ToArray();

            if (selectedRows.Length == 0)
                return;

            if (selectedRows.Length == 1)
                errorItems[selectedRows[0].RowIndex].Range.Select();

            var range = selectedRows
                .Select(x => errorItems[x.RowIndex].Range)
                .Aggregate((r1, r2) => App.Union(r1, r2));

            range.Select();
        }
    }
}
