using Microsoft.Office.Interop.Excel;

namespace NavfertyExcelAddIn.WorksheetCellsEditing
{
	public interface IEmptySpaceTrimmer
	{
		void TrimExtraSpaces(Range range);
		void RemoveAllSpaces(Range range);
	}
}
