using Microsoft.Office.Interop.Excel;

namespace NavfertyExcelAddIn.WorksheetProtectorUnprotector
{
	public interface IWsProtectorUnprotector
	{
		void ProtectUnprotectSelectedWorksheets(Workbook wb);
	}
}
