using System.Data.SQLite;
using System.Reflection;

using Microsoft.Office.Interop.Excel;

using Moq;

using Navferty.Common;
using NavfertyExcelAddIn.SqliteExport;
using NavfertyExcelAddIn.SqliteImport;

using Range = Microsoft.Office.Interop.Excel.Range;

namespace NavfertyExcelAddIn.UnitTests.SqliteImport;

public class SqliteDatabaseTests
{
	private string path = null!;

	[Before(HookType.Test)]
	public void Setup() => path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.sqlite");

	[After(HookType.Test)]
	public void Cleanup()
	{
		SQLiteConnection.ClearAllPools();
		if (File.Exists(path)) File.Delete(path);
	}

	[Test]
	public async Task ListsAndReadsTablesViewsAndTypes()
	{
		using (var connection = new SQLiteConnection($"Data Source={path};Version=3;"))
		{
			connection.Open();
			using var command = new SQLiteCommand("CREATE TABLE \"a\"\"b\" (id INTEGER, real REAL, text TEXT, blob BLOB, missing); INSERT INTO \"a\"\"b\" VALUES (123, 2.5, '=1+1', X'00FF', NULL); CREATE VIEW v AS SELECT * FROM \"a\"\"b\"", connection);
			command.ExecuteNonQuery();
		}
		using var database = new SqliteDatabase(path);
		var objects = database.GetObjects();
		await Assert.That(objects.Count).IsEqualTo(2);
		await Assert.That(objects.Single(x => x.Name == "a\"b").Count).IsEqualTo(1L);
		await Assert.That(objects.Single(x => x.Name == "v").Count).IsEqualTo(1L);
		string[] headers = null!;
		object[,] values = null!;
		database.ReadTable("a\"b", h => headers = h, (v, count) => values = v);
		await Assert.That(string.Join(",", headers)).IsEqualTo("id,real,text,blob,missing");
		await Assert.That(values[0, 0]).IsEqualTo((object)123d);
		await Assert.That(values[0, 1]).IsEqualTo((object)2.5d);
		await Assert.That(values[0, 2]).IsEqualTo((object)"=1+1");
		await Assert.That(values[0, 3]).IsEqualTo((object)"AP8=");
		await Assert.That(values[0, 4]).IsNull();
		database.ReadQuery("SELECT text FROM v", h => headers = h, (v, count) => values = v);
		await Assert.That(headers[0]).IsEqualTo("text");
		await Assert.That(values[0, 0]).IsEqualTo((object)"=1+1");
		Assert.Throws<ArgumentException>(() => database.ReadQuery("SELECT 1; DELETE FROM v", _ => { }, (_, _) => { }));
		Assert.Throws<ArgumentException>(() => database.ReadQuery("DELETE FROM v", _ => { }, (_, _) => { }));
	}

	[Test]
	public async Task MissingDatabaseIsNotCreated()
	{
		Assert.Throws<SQLiteException>(() => new SqliteDatabase(path));
		await Assert.That(File.Exists(path)).IsFalse();
	}

	[Test]
	public async Task ReadsInBatchesAndPreservesLargeIntegers()
	{
		using (var connection = new SQLiteConnection($"Data Source={path};Version=3;"))
		{
			connection.Open();
			using var command = new SQLiteCommand("CREATE TABLE t (n INTEGER); WITH RECURSIVE seq(n) AS (SELECT 1 UNION ALL SELECT n+1 FROM seq WHERE n<501) INSERT INTO t SELECT n FROM seq; INSERT INTO t VALUES (9223372036854775807)", connection);
			command.ExecuteNonQuery();
		}
		using var database = new SqliteDatabase(path);
		var counts = new List<int>();
		object last = null!;
		database.ReadTable("t", _ => { }, (values, count) => { counts.Add(count); last = values[count - 1, 0]; });
		await Assert.That(string.Join(",", counts)).IsEqualTo("500,2");
		await Assert.That(last).IsEqualTo((object)"9223372036854775807");
	}

	[Test]
	public async Task ExportThenImportMatchesWorkbookValues()
	{
		var workbook = new Mock<Workbook>();
		var sheets = new Mock<Sheets>();
		var sheet = new Mock<Worksheet>();
		var range = new Mock<Range>();
		var data = Array.CreateInstance(typeof(object), new[] { 3, 3 }, new[] { 1, 1 });
		data.SetValue("Name", 1, 1); data.SetValue("Count", 1, 2); data.SetValue("Price", 1, 3);
		data.SetValue("Apple", 2, 1); data.SetValue(2d, 2, 2); data.SetValue(1.5d, 2, 3);
		data.SetValue("Pear", 3, 1); data.SetValue(4d, 3, 2); data.SetValue(3.25d, 3, 3);
		sheet.Setup(x => x.Name).Returns("Products");
		sheet.Setup(x => x.UsedRange).Returns(range.Object);
		range.Setup(x => x.Rows.Count).Returns(3);
		range.Setup(x => x.Columns.Count).Returns(3);
		range.Setup(x => x.get_Value(Missing.Value)).Returns(data);
		sheets.Setup(x => x.Count).Returns(1);
		sheets.Setup(x => x.GetEnumerator()).Returns(new[] { sheet.Object }.GetEnumerator());
		workbook.Setup(x => x.Worksheets).Returns(sheets.Object);
		var dialog = new Mock<IDialogService>();
		dialog.Setup(x => x.AskForSaveFile(FileType.Db)).Returns(path);
		new SqliteExporter(dialog.Object, Mock.Of<ISqliteExportOptionsProvider>())
			.ExportToSqlite(workbook.Object, new SqliteExportOptions());

		using var database = new SqliteDatabase(path);
		string[] headers = null!;
		object[,] values = null!;
		database.ReadTable("Products", h => headers = h, (v, count) => values = v);
		for (int col = 0; col < 3; col++)
		{
			await Assert.That(headers[col]).IsEqualTo((string)data.GetValue(1, col + 1)!);
			for (int row = 0; row < 2; row++)
				await Assert.That(values[row, col]).IsEqualTo(data.GetValue(row + 2, col + 1));
		}
	}
}
