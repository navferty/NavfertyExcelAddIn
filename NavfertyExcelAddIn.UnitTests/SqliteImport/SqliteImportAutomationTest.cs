using Microsoft.Office.Interop.Excel;

using Moq;

using Navferty.Common;
using NavfertyExcelAddIn.SqliteExport;
using NavfertyExcelAddIn.SqliteImport;

using Range = Microsoft.Office.Interop.Excel.Range;

namespace NavfertyExcelAddIn.UnitTests.SqliteImport;

[Category("Automation")]
[NotInParallel("Automation")]
public class SqliteImportAutomationTest : AutomationTestsBase
{
	[Test]
	public async Task ExportImportRoundTrip()
	{
		var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.db");
		Workbook? original = null;
		Workbook? imported = null;
		try
		{
			original = app.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
			var source = (Worksheet)original.Worksheets[1];
			source.Name = "Products";
			var expected = new object[,]
			{
				{ "Name", "Quantity", "Price" },
				{ "Apple", 2d, 1.5d },
				{ "Pear", 4d, 3.25d },
				{ "=1+1", 1d, 2.5d }
			};
			var sourceRange = source.Range[source.Cells[1, 1], source.Cells[4, 3]];
			sourceRange.NumberFormat = "@";
			sourceRange.Value2 = expected;
			var dialog = new Mock<IDialogService>();
			dialog.Setup(x => x.AskForSaveFile(FileType.Db)).Returns(path);
			new SqliteExporter(dialog.Object, Mock.Of<ISqliteExportOptionsProvider>())
				.ExportToSqlite(original, new SqliteExportOptions());
			using (var database = new SqliteDatabase(path))
				new SqliteImporter().ImportWorkbook(app, database, new[] { "Products" }, "");
			imported = app.ActiveWorkbook;
			var target = (Worksheet)imported.Worksheets[1];
			await Assert.That(target.Name).IsEqualTo("Products");
			for (int row = 1; row <= 4; row++)
				for (int col = 1; col <= 3; col++)
					await Assert.That((object)((Range)target.Cells[row, col]).Value2)
						.IsEqualTo((object)((Range)source.Cells[row, col]).Value2);
			await Assert.That((bool)((Range)target.Cells[4, 1]).HasFormula).IsFalse();
		}
		finally
		{
			imported?.Close(false);
			original?.Close(false);
			System.Data.SQLite.SQLiteConnection.ClearAllPools();
			if (File.Exists(path)) File.Delete(path);
		}
	}
}
