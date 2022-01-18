using Microsoft.Office.Interop.Excel;

namespace NavfertyExcelAddIn.WorksheetProtectUnprotect
{
	public interface IWsProtectorUnprotector
	{
		void ProtectUnprotectSelectedWorksheets(Workbook wb);
	}
}
