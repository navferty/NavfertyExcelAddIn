using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using NavfertyExcelAddIn.InteractiveRangeReport;

namespace NavfertyExcelAddIn.FindFormulaErrors
{
    public interface IErrorFinder
    {
        IReadOnlyCollection<InteractiveErrorItem> GetAllErrorCells(Range range);
    }
}
