using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;

using Microsoft.Office.Interop.Excel;

using Application = Microsoft.Office.Interop.Excel.Application;

namespace NavfertyExcelAddIn.Commons
{
    public partial class CellsResult : Form
    {
        private readonly IList<ListRangeItem> items;
        private Application App => Globals.ThisAddIn.Application;

        public CellsResult(IReadOnlyCollection<ErroredRange> ranges)
        {
            InitializeComponent();
            
            ErrorType.HeaderText = Localization.UIStrings.ErrorType;
            Formula.HeaderText = Localization.UIStrings.Formula;
            Address.HeaderText = Localization.UIStrings.Address;
            WsName.HeaderText = Localization.UIStrings.WsName;
            Text = Localization.UIStrings.SearchResults;

            items = ranges.Select(r => new ListRangeItem(r)).ToArray();
            RangesGridView.DataSource = items;
            RangesGridView.SelectionChanged += OnSelectionChanged;
        }

        private void OnSelectionChanged(object sender, EventArgs e)
        {
            var selectedRows = RangesGridView.SelectedCells.Cast<DataGridViewCell>().ToArray();

            if (selectedRows.Length == 0)
                return;

            if (selectedRows.Length == 1)
                items[selectedRows[0].RowIndex].Range.Select();

            var range = selectedRows
                .Select(x => items[x.RowIndex].Range)
                .Aggregate((r1, r2) => App.Union(r1, r2));

            range.Select();
        }

        private class ListRangeItem
        {
            private readonly ErroredRange erroredRange;

            public ListRangeItem(ErroredRange erroredRange)
            {
                this.erroredRange = erroredRange;
            }
            public string ErrorType => erroredRange.ErrorType.GetEnumDescription();
            public string Formula => (string)Range.Formula;
            public string Address => Range.Address[false, false];
            public string WsName => Range.Worksheet.Name;
            public Range Range => erroredRange.Range;
        }
    }
}
