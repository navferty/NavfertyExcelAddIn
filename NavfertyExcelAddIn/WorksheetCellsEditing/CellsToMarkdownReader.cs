using System.Text;
using Microsoft.Office.Interop.Excel;
using NavfertyExcelAddIn.Commons;

namespace NavfertyExcelAddIn.WorksheetCellsEditing
{
    public class CellsToMarkdownReader : ICellsToMarkdownReader
    {
        public string ReadToMarkdown(Range range)
        {
            var markdown = new StringBuilder();

            foreach (Range row in range.Rows)
            {
                foreach (Range cell in row.Cells)
                {
                    markdown.Append('|');
                    var value = (string)cell.Value?.ToString();
                    markdown.Append(value.TrimSpaces() ?? " ");
                }
                markdown.Append('|');
                markdown.Append("\r\n");
            }

            return markdown.ToString();
        }
    }
}
