using Microsoft.Office.Interop.Excel;

namespace NavfertyExcelAddIn.FindFormulaErrors
{
    public class ErroredRange
    {
        public ErroredRange(Range range, CVErrEnum errorType)
        {
            Range = range;
            ErrorType = errorType;
        }

        public Range Range { get; }
        public CVErrEnum ErrorType { get; } 
    }
}
