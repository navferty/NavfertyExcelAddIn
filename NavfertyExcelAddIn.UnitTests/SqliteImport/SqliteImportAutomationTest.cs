using System.Data.SQLite;
using System.Reflection;

using Microsoft.Office.Interop.Excel;

using Moq;

using Navferty.Common;
using NavfertyExcelAddIn.SqliteExport;
using NavfertyExcelAddIn.SqliteImport;

using Range = Microsoft.Office.Interop.Excel.Range;
using Action = System.Action;
using Window = Microsoft.Office.Interop.Excel.Window;

namespace NavfertyExcelAddIn.UnitTests.SqliteImport;

[Category("Automation")]
[NotInParallel("Automation")]
public class SqliteImportAutomationTest : AutomationTestsBase
{
	[Test]
	public async Task ImportKeepsWorkbookHiddenAndCleansUpOnCancellationOrFailure()
	{
		var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.db");
		Workbook? original = null;
		Workbook? imported = null;
		bool originalScreenUpdating = app.ScreenUpdating;
		try
		{
			using (var connection = new SQLiteConnection($"Data Source={path};Version=3;"))
			{
				connection.Open();
				using var command = new SQLiteCommand("CREATE TABLE items (value INTEGER); INSERT INTO items VALUES (1), (2)", connection);
				command.ExecuteNonQuery();
			}
			original = app.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
			int originalCount = app.Workbooks.Count;
			using var database = new SqliteDatabase(path);
			var core = typeof(SqliteImporter).GetMethod("ImportWorkbookCore", BindingFlags.NonPublic | BindingFlags.Static)!;

			foreach (string scenario in new[] { "success", "cancel", "error" })
			{
				app.ScreenUpdating = scenario != "cancel";
				bool expectedScreenUpdating = app.ScreenUpdating;
				bool alwaysHidden = true, committed = false;
				int reports = 0;
				using var cancellation = new CancellationTokenSource();
				Action<Action> dispatch = action => action();
				Action<string, int> report = (name, count) =>
				{
					var pending = app.Workbooks[originalCount + 1];
					reports++;
					alwaysHidden &= !((Window)pending.Windows[1]).Visible;
					if (scenario == "cancel" && count > 0) cancellation.Cancel();
				};
				Action complete = () => committed = true;
				Exception? error = null;
				try
				{
					core.Invoke(null, new object[] { app, database, scenario == "error" ? Array.Empty<string>() : new[] { "items" },
						scenario == "error" ? "SELECT FROM items" : "", dispatch, report, complete, cancellation.Token });
				}
				catch (TargetInvocationException ex) { error = ex.InnerException; }

				await Assert.That(reports > 0 && alwaysHidden).IsTrue();
				await Assert.That(app.ScreenUpdating).IsEqualTo(expectedScreenUpdating);
				if (scenario == "success")
				{
					imported = app.Workbooks[originalCount + 1];
					await Assert.That(error == null).IsTrue();
					await Assert.That(committed).IsTrue();
					await Assert.That(((Window)imported.Windows[1]).Visible).IsTrue();
					imported.Close(false);
					imported = null;
				}
				else
				{
					await Assert.That(error != null).IsTrue();
					await Assert.That(scenario != "cancel" || error is OperationCanceledException).IsTrue();
					await Assert.That(committed).IsFalse();
					await Assert.That(app.Workbooks.Count).IsEqualTo(originalCount);
				}
				await Assert.That(((Window)original.Windows[1]).Visible).IsTrue();
			}
		}
		finally
		{
			imported?.Close(false);
			original?.Close(false);
			app.ScreenUpdating = originalScreenUpdating;
			SQLiteConnection.ClearAllPools();
			if (File.Exists(path)) File.Delete(path);
		}
	}

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
