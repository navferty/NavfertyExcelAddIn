using System.Collections.Generic;
using System.Linq;

using Microsoft.Office.Interop.Excel;

using NavfertyExcelAddIn.Commons;

namespace NavfertyExcelAddIn.WorksheetCellsEditing
{
    public class CellsUnmerger : ICellsUnmerger
    {
        public void Unmerge(Range range)
        {
            range.OfType<Range>()
                .Select(x => x.MergeArea)
                .Where(x => x.Count != 1)
                .Distinct(new MergeAreaEqualityComparer())
                .ForEach(UnmergeArea);
        }

        private void UnmergeArea(Range currentRange)
        {
            var formula = currentRange.Cells.OfType<Range>().First().Formula;

            currentRange.UnMerge();
            currentRange.Formula = formula;
        }

        private class MergeAreaEqualityComparer : IEqualityComparer<Range>
        {
            public bool Equals(Range x, Range y)
            {
                return y != null
                    && x != null
                    && x.Row == y.Row
                    && x.Column == y.Column;
            }

            public int GetHashCode(Range obj)
            {
                return obj.Row * 9973 + obj.Column;
            }
        }
    }
}
