using Microsoft.Office.Interop.Excel;

namespace NavfertyExcelAddIn.WorksheetCellsEditing
{
	public interface ICellsUnmerger
	{
		void Unmerge(Range range);
	}
}