using Microsoft.Office.Interop.Excel;

namespace NavfertyExcelAddIn.Commons
{
    public interface ICellsUnmerger
    {
        void Unmerge(Range range);
    }
}