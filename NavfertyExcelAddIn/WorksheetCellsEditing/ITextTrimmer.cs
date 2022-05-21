using Microsoft.Office.Interop.Excel;

namespace NavfertyExcelAddIn.WorksheetCellsEditing
{
	public interface ITextTrimmer
	{
		void TrimExtraSpaces(Range range);
		void RemoveAllSpaces(Range range);
		void TrimTextByLengthUIDisplay(Range range);
	}
}
