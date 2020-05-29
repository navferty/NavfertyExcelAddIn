using Microsoft.Office.Interop.Excel;

namespace NavfertyExcelAddIn.WorksheetCellsEditing
{
	public interface ICaseToggler
	{
		void ToggleCase(Range range);
	}
}