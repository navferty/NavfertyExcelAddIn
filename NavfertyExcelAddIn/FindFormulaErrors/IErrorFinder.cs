using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

namespace NavfertyExcelAddIn.FindFormulaErrors
{
    public interface IErrorFinder
    {
        IReadOnlyCollection<ErroredRange> GetAllErrorCells(Range range);
    }
}
