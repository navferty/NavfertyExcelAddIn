using Microsoft.Office.Interop.Excel;

namespace NavfertyExcelAddIn.SqliteExport
{
	public interface ISqliteExporter
	{
		void ExportToSqlite(Workbook workbook);
		void ExportToSqlite(Workbook workbook, SqliteExportOptions options);
	}
}
