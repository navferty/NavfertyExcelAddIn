using System.Linq;
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

            var isHeaderRow = true;

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

                if (isHeaderRow)
                {
                    Enumerable
                        .Repeat("|---", range.Columns.Count)
                        .ForEach(x => markdown.Append(x));

                    markdown.Append("|\r\n");

                    isHeaderRow = false;
                }
            }

            return markdown.ToString();
        }
    }
}
