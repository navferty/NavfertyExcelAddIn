using Microsoft.Office.Interop.Excel;

namespace NavfertyExcelAddIn.WorksheetCellsEditing
{
	public interface IDuplicatesHighlighter
	{
		void HighlightDuplicates(Range range);
	}
}