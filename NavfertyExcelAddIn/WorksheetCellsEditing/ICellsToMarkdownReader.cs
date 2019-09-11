using Microsoft.Office.Interop.Excel;

namespace NavfertyExcelAddIn.WorksheetCellsEditing
{
    public interface ICellsToMarkdownReader
    {
        string ReadToMarkdown(Range range);
    }
}