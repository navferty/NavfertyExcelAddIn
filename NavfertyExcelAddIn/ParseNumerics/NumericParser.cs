using Microsoft.Office.Interop.Excel;
using NavfertyExcelAddIn.Commons;

namespace NavfertyExcelAddIn.ParseNumerics
{
    public class NumericParser : INumericParser
    {
        public void Parse(Range selection)
        {
            selection.ApplyForEachCellOfType<string, decimal?>(s => s.ParseDecimal());
        }
    }
}
