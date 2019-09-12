using Microsoft.Office.Interop.Excel;

namespace NavfertyExcelAddIn.WorksheetCellsEditing
{
    public interface IEmptySpaceTrimmer
    {
        void TrimSpaces(Range range);
    }
}