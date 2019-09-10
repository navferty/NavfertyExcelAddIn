using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Excel;

namespace NavfertyExcelAddIn.FindFormulaErrors
{
    public class ErrorFinder : IErrorFinder
    {
        public IReadOnlyCollection<ErroredRange> GetAllErrorCells(Range range)
        {
            return GetErroredCells(range).ToArray();
        }

        private IEnumerable<ErroredRange> GetErroredCells(Range range)
        {
            var rangeValue = range.Value;

            if (rangeValue == null)
                yield break;

            // only single cell exists on sheet
            if (!(rangeValue is object[,] values))
            {
                if (IsXlCvErr(rangeValue))
                {
                    yield return new ErroredRange(range, (CVErrEnum)rangeValue);
                }
                yield break;
            }

            int upperI = values.GetUpperBound(0); // Columns
            int upperJ = values.GetUpperBound(1); // Rows

            for (int i = values.GetLowerBound(0); i <= upperI; i++)
            {
                for (int j = values.GetLowerBound(1); j <= upperJ; j++)
                {
                    var value = values[i, j];
                    if (IsXlCvErr(value))
                    {
                        yield return new ErroredRange((Range)range.Cells[i, j], (CVErrEnum)value);
                    }
                }
            }
        }

        private bool IsXlCvErr(object obj)
        {
            // only CVErr values in Excel are interpreted as Int32 values in .NET
            // https://stackoverflow.com/questions/16217350/vba-looking-for-error-values-in-a-specific-column
            // https://xldennis.wordpress.com/2006/11/22/dealing-with-cverr-values-in-net-%E2%80%93-part-i-the-problem/
            return (obj) is Int32;
        }
    }
}
