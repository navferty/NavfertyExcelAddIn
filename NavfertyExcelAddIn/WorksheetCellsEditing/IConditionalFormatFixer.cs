using Microsoft.Office.Interop.Excel;

namespace NavfertyExcelAddIn.WorksheetCellsEditing
{
	public interface IConditionalFormatFixer
	{
		void FillRange(Range range);
	}
}
