namespace NavfertyExcelAddIn.SqliteExport
{
	public class SqliteExportOptions
	{
		public bool UseFirstRowAsHeaders { get; set; }
		public int RowsToSkip { get; set; }

		public SqliteExportOptions()
		{
			UseFirstRowAsHeaders = true;
			RowsToSkip = 0;
		}
	}
}
