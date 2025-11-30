using System.Data.SQLite;
using System.Reflection;

using Microsoft.Office.Interop.Excel;

using Moq;

using Navferty.Common;

using NavfertyExcelAddIn.SqliteExport;

using Range = Microsoft.Office.Interop.Excel.Range;

namespace NavfertyExcelAddIn.UnitTests.SqliteExport;

public class SqliteExporterTests : TestsBase
{
	private Mock<IDialogService> dialogService = null!;
	private Mock<ISqliteExportOptionsProvider> optionsProvider = null!;
	private string testDbPath = string.Empty;

	[Before(HookType.Test)]
	public void BeforeEachTest()
	{
		dialogService = new Mock<IDialogService>(MockBehavior.Strict);
		optionsProvider = new Mock<ISqliteExportOptionsProvider>(MockBehavior.Strict);
		testDbPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid()}.db");
	}

	[After(HookType.Test)]
	public void AfterEachTest()
	{
		if (File.Exists(testDbPath))
		{
			File.Delete(testDbPath);
		}
	}

	[Test]
	public async Task ExportToSqlite_NullWorkbook_ThrowsArgumentNullException()
	{
		var exporter = new SqliteExporter(dialogService.Object, optionsProvider.Object);

		Assert.Throws<ArgumentNullException>(() => exporter.ExportToSqlite(null!));
	}

	[Test]
	public async Task ExportToSqlite_UserCancelsOptionsDialog_DoesNothing()
	{
		var workbook = new Mock<Workbook>(MockBehavior.Strict);
		var exporter = new SqliteExporter(dialogService.Object, optionsProvider.Object);

		optionsProvider
			.Setup(x => x.TryGetOptions(out It.Ref<SqliteExportOptions?>.IsAny))
			.Returns(false);

		exporter.ExportToSqlite(workbook.Object);

		// Verify that no dialog service methods were called since user cancelled
		dialogService.VerifyAll();
	}

	[Test]
	public async Task ExportToSqlite_NoDbPathSelected_DoesNothing()
	{
		var workbook = new Mock<Workbook>(MockBehavior.Strict);
		var exporter = new SqliteExporter(dialogService.Object, optionsProvider.Object);

		var testOptions = new SqliteExportOptions
		{
			UseFirstRowAsHeaders = true,
			RowsToSkip = 0
		};

		optionsProvider
			.Setup(x => x.TryGetOptions(out testOptions))
			.Returns(true);

		dialogService
			.Setup(x => x.AskForSaveFile(FileType.Db))
			.Returns(string.Empty);

		exporter.ExportToSqlite(workbook.Object);

		dialogService.VerifyAll();
	}

	[Test]
	public async Task ExportToSqlite_ValidWorkbook_CreatesDatabase()
	{
		var workbook = new Mock<Workbook>(MockBehavior.Loose);
		var worksheets = new Mock<Sheets>(MockBehavior.Loose);
		var worksheet = new Mock<Worksheet>(MockBehavior.Loose);
		var usedRange = new Mock<Range>(MockBehavior.Loose);

		var testData = new object[3, 3];
		testData[1, 1] = "A1";
		testData[1, 2] = "B1";
		testData[2, 1] = "A2";
		testData[2, 2] = "B2";

		worksheet.Setup(x => x.Name).Returns("TestSheet");
		worksheet.Setup(x => x.UsedRange).Returns(usedRange.Object);
		usedRange.Setup(x => x.Rows.Count).Returns(2);
		usedRange.Setup(x => x.Columns.Count).Returns(2);
		usedRange.Setup(x => x.get_Value(Missing.Value)).Returns(testData);

		worksheets.Setup(x => x.Count).Returns(1);
		worksheets.Setup(x => x.GetEnumerator()).Returns(new[] { worksheet.Object }.GetEnumerator());

		workbook.Setup(x => x.Worksheets).Returns(worksheets.Object);

		var testOptions = new SqliteExportOptions
		{
			UseFirstRowAsHeaders = false,
			RowsToSkip = 0
		};

		optionsProvider
			.Setup(x => x.TryGetOptions(out testOptions))
			.Returns(true);

		dialogService
			.Setup(x => x.AskForSaveFile(FileType.Db))
			.Returns(testDbPath);

		dialogService
			.Setup(x => x.ShowInfo(It.IsAny<string>()));
		var exporter = new SqliteExporter(dialogService.Object, optionsProvider.Object);

		exporter.ExportToSqlite(workbook.Object);

		await Assert.That(File.Exists(testDbPath)).IsTrue();

		using (var connection = new SQLiteConnection($"Data Source={testDbPath};Version=3;"))
		{
			connection.Open();
			using (var command = new SQLiteCommand("SELECT name FROM sqlite_master WHERE type='table'", connection))
			using (var reader = command.ExecuteReader())
			{
				await Assert.That(reader.Read()).IsTrue();
				await Assert.That(reader.GetString(0)).IsEqualTo("TestSheet");
			}

			using (var command = new SQLiteCommand("SELECT COUNT(*) FROM TestSheet", connection))
			{
				var count = (long)command.ExecuteScalar();
				await Assert.That(count).IsEqualTo(2);
			}
		}

		dialogService.VerifyAll();
	}

	[Test]
	public async Task ExportToSqlite_SheetNameWithSpecialChars_SanitizesTableName()
	{
		var workbook = new Mock<Workbook>(MockBehavior.Loose);
		var worksheets = new Mock<Sheets>(MockBehavior.Loose);
		var worksheet = new Mock<Worksheet>(MockBehavior.Loose);
		var usedRange = new Mock<Range>(MockBehavior.Loose);

		var testData = new object[2, 2];
		testData[1, 1] = "Value";

		worksheet.Setup(x => x.Name).Returns("Test-Sheet!@#");
		worksheet.Setup(x => x.UsedRange).Returns(usedRange.Object);
		usedRange.Setup(x => x.Rows.Count).Returns(1);
		usedRange.Setup(x => x.Columns.Count).Returns(1);
		usedRange.Setup(x => x.get_Value(Missing.Value)).Returns(testData);

		worksheets.Setup(x => x.Count).Returns(1);
		worksheets.Setup(x => x.GetEnumerator()).Returns(new[] { worksheet.Object }.GetEnumerator());

		workbook.Setup(x => x.Worksheets).Returns(worksheets.Object);

		var testOptions = new SqliteExportOptions
		{
			UseFirstRowAsHeaders = false,
			RowsToSkip = 0
		};

		optionsProvider
			.Setup(x => x.TryGetOptions(out testOptions))
			.Returns(true);

		dialogService
			.Setup(x => x.AskForSaveFile(FileType.Db))
			.Returns(testDbPath);

		dialogService
			.Setup(x => x.ShowInfo(It.IsAny<string>()));
		var exporter = new SqliteExporter(dialogService.Object, optionsProvider.Object);

		exporter.ExportToSqlite(workbook.Object);

		await Assert.That(File.Exists(testDbPath)).IsTrue();

		using (var connection = new SQLiteConnection($"Data Source={testDbPath};Version=3;"))
		{
			connection.Open();
			using var command = new SQLiteCommand("SELECT name FROM sqlite_master WHERE type='table'", connection);
			using var reader = command.ExecuteReader();
			await Assert.That(reader.Read()).IsTrue();
			var tableName = reader.GetString(0);
			await Assert.That(tableName).IsEqualTo("Test_Sheet___");
		}

		dialogService.VerifyAll();
	}

	[Test]
	public async Task ExportToSqlite_EmptyWorksheet_SkipsEmptySheet()
	{
		var workbook = new Mock<Workbook>(MockBehavior.Loose);
		var worksheets = new Mock<Sheets>(MockBehavior.Loose);
		var worksheet = new Mock<Worksheet>(MockBehavior.Loose);

		worksheet.Setup(x => x.Name).Returns("EmptySheet");
		worksheet.Setup(x => x.UsedRange).Returns((Range)null!);

		worksheets.Setup(x => x.Count).Returns(1);
		worksheets.Setup(x => x.GetEnumerator()).Returns(new[] { worksheet.Object }.GetEnumerator());

		workbook.Setup(x => x.Worksheets).Returns(worksheets.Object);

		var testOptions = new SqliteExportOptions
		{
			UseFirstRowAsHeaders = false,
			RowsToSkip = 0
		};

		optionsProvider
			.Setup(x => x.TryGetOptions(out testOptions))
			.Returns(true);

		dialogService
			.Setup(x => x.AskForSaveFile(FileType.Db))
			.Returns(testDbPath);

		dialogService
			.Setup(x => x.ShowInfo(It.IsAny<string>()));
		var exporter = new SqliteExporter(dialogService.Object, optionsProvider.Object);

		exporter.ExportToSqlite(workbook.Object);

		await Assert.That(File.Exists(testDbPath)).IsTrue();

		using (var connection = new SQLiteConnection($"Data Source={testDbPath};Version=3;"))
		{
			connection.Open();
			using var command = new SQLiteCommand("SELECT COUNT(*) FROM sqlite_master WHERE type='table'", connection);
			var tableCount = (long)command.ExecuteScalar();
			await Assert.That(tableCount).IsEqualTo(0);
		}

		dialogService.VerifyAll();
	}

	[Test]
	public async Task ExportToSqlite_WithOptions_UseFirstRowAsHeaders_UsesHeaderNames()
	{
		var workbook = new Mock<Workbook>(MockBehavior.Loose);
		var worksheets = new Mock<Sheets>(MockBehavior.Loose);
		var worksheet = new Mock<Worksheet>(MockBehavior.Loose);
		var usedRange = new Mock<Range>(MockBehavior.Loose);

		var testData = new object[3, 3];
		testData[1, 1] = "Name";
		testData[1, 2] = "Age";
		testData[2, 1] = "John";
		testData[2, 2] = 30;

		worksheet.Setup(x => x.Name).Returns("TestSheet");
		worksheet.Setup(x => x.UsedRange).Returns(usedRange.Object);
		usedRange.Setup(x => x.Rows.Count).Returns(2);
		usedRange.Setup(x => x.Columns.Count).Returns(2);
		usedRange.Setup(x => x.get_Value(Missing.Value)).Returns(testData);

		worksheets.Setup(x => x.Count).Returns(1);
		worksheets.Setup(x => x.GetEnumerator()).Returns(new[] { worksheet.Object }.GetEnumerator());

		workbook.Setup(x => x.Worksheets).Returns(worksheets.Object);

		var options = new SqliteExportOptions
		{
			UseFirstRowAsHeaders = true,
			RowsToSkip = 0
		};

		dialogService
			.Setup(x => x.AskForSaveFile(FileType.Db))
			.Returns(testDbPath);

		dialogService
			.Setup(x => x.ShowInfo(It.IsAny<string>()));
		var exporter = new SqliteExporter(dialogService.Object, optionsProvider.Object);

		exporter.ExportToSqlite(workbook.Object, options);

		await Assert.That(File.Exists(testDbPath)).IsTrue();

		using (var connection = new SQLiteConnection($"Data Source={testDbPath};Version=3;"))
		{
			connection.Open();
			using var command = new SQLiteCommand("PRAGMA table_info(TestSheet)", connection);
			using var reader = command.ExecuteReader();
			await Assert.That(reader.Read()).IsTrue();
			await Assert.That(reader.GetString(1)).IsEqualTo("Name");
			await Assert.That(reader.Read()).IsTrue();
			await Assert.That(reader.GetString(1)).IsEqualTo("Age");
		}

		dialogService.VerifyAll();
	}

	[Test]
	public async Task ExportToSqlite_WithOptions_SkipRows_SkipsSpecifiedRows()
	{
		var workbook = new Mock<Workbook>(MockBehavior.Loose);
		var worksheets = new Mock<Sheets>(MockBehavior.Loose);
		var worksheet = new Mock<Worksheet>(MockBehavior.Loose);
		var usedRange = new Mock<Range>(MockBehavior.Loose);

		var testData = new object[5, 2];
		testData[1, 1] = "Skip1";
		testData[2, 1] = "Skip2";
		testData[3, 1] = "Header";
		testData[4, 1] = "Data1";

		worksheet.Setup(x => x.Name).Returns("TestSheet");
		worksheet.Setup(x => x.UsedRange).Returns(usedRange.Object);
		usedRange.Setup(x => x.Rows.Count).Returns(4);
		usedRange.Setup(x => x.Columns.Count).Returns(1);
		usedRange.Setup(x => x.get_Value(Missing.Value)).Returns(testData);

		worksheets.Setup(x => x.Count).Returns(1);
		worksheets.Setup(x => x.GetEnumerator()).Returns(new[] { worksheet.Object }.GetEnumerator());

		workbook.Setup(x => x.Worksheets).Returns(worksheets.Object);

		var options = new SqliteExportOptions
		{
			UseFirstRowAsHeaders = true,
			RowsToSkip = 2
		};

		dialogService
			.Setup(x => x.AskForSaveFile(FileType.Db))
			.Returns(testDbPath);

		dialogService
			.Setup(x => x.ShowInfo(It.IsAny<string>()));
		var exporter = new SqliteExporter(dialogService.Object, optionsProvider.Object);

		exporter.ExportToSqlite(workbook.Object, options);

		await Assert.That(File.Exists(testDbPath)).IsTrue();

		using (var connection = new SQLiteConnection($"Data Source={testDbPath};Version=3;"))
		{
			connection.Open();
			using var command = new SQLiteCommand("SELECT COUNT(*) FROM TestSheet", connection);
			var count = (long)command.ExecuteScalar();
			await Assert.That(count).IsEqualTo(1); // Only one data row after skipping 2 and using 1 as header
		}

		dialogService.VerifyAll();
	}

	[Test]
	public async Task ExportToSqlite_WithOptions_IntegerColumn_CreatesIntegerType()
	{
		var workbook = new Mock<Workbook>(MockBehavior.Loose);
		var worksheets = new Mock<Sheets>(MockBehavior.Loose);
		var worksheet = new Mock<Worksheet>(MockBehavior.Loose);
		var usedRange = new Mock<Range>(MockBehavior.Loose);

		var testData = new object[4, 2];
		testData[1, 1] = "ID";
		testData[2, 1] = 1;
		testData[3, 1] = 2;

		worksheet.Setup(x => x.Name).Returns("TestSheet");
		worksheet.Setup(x => x.UsedRange).Returns(usedRange.Object);
		usedRange.Setup(x => x.Rows.Count).Returns(3);
		usedRange.Setup(x => x.Columns.Count).Returns(1);
		usedRange.Setup(x => x.get_Value(Missing.Value)).Returns(testData);

		worksheets.Setup(x => x.Count).Returns(1);
		worksheets.Setup(x => x.GetEnumerator()).Returns(new[] { worksheet.Object }.GetEnumerator());

		workbook.Setup(x => x.Worksheets).Returns(worksheets.Object);

		var options = new SqliteExportOptions
		{
			UseFirstRowAsHeaders = true,
			RowsToSkip = 0
		};

		dialogService
			.Setup(x => x.AskForSaveFile(FileType.Db))
			.Returns(testDbPath);

		dialogService
			.Setup(x => x.ShowInfo(It.IsAny<string>()));
		var exporter = new SqliteExporter(dialogService.Object, optionsProvider.Object);

		exporter.ExportToSqlite(workbook.Object, options);

		await Assert.That(File.Exists(testDbPath)).IsTrue();

		using (var connection = new SQLiteConnection($"Data Source={testDbPath};Version=3;"))
		{
			connection.Open();
			using var command = new SQLiteCommand("PRAGMA table_info(TestSheet)", connection);
			using var reader = command.ExecuteReader();
			await Assert.That(reader.Read()).IsTrue();
			await Assert.That(reader.GetString(1)).IsEqualTo("ID");
			await Assert.That(reader.GetString(2)).IsEqualTo("INTEGER");
		}

		dialogService.VerifyAll();
	}

	[Test]
	public async Task ExportToSqlite_WithOptions_RealColumn_CreatesRealType()
	{
		var workbook = new Mock<Workbook>(MockBehavior.Loose);
		var worksheets = new Mock<Sheets>(MockBehavior.Loose);
		var worksheet = new Mock<Worksheet>(MockBehavior.Loose);
		var usedRange = new Mock<Range>(MockBehavior.Loose);

		var testData = new object[4, 2];
		testData[1, 1] = "Price";
		testData[2, 1] = 19.99;
		testData[3, 1] = 29.99;

		worksheet.Setup(x => x.Name).Returns("TestSheet");
		worksheet.Setup(x => x.UsedRange).Returns(usedRange.Object);
		usedRange.Setup(x => x.Rows.Count).Returns(3);
		usedRange.Setup(x => x.Columns.Count).Returns(1);
		usedRange.Setup(x => x.get_Value(Missing.Value)).Returns(testData);

		worksheets.Setup(x => x.Count).Returns(1);
		worksheets.Setup(x => x.GetEnumerator()).Returns(new[] { worksheet.Object }.GetEnumerator());

		workbook.Setup(x => x.Worksheets).Returns(worksheets.Object);

		var options = new SqliteExportOptions
		{
			UseFirstRowAsHeaders = true,
			RowsToSkip = 0
		};

		dialogService
			.Setup(x => x.AskForSaveFile(FileType.Db))
			.Returns(testDbPath);

		dialogService
			.Setup(x => x.ShowInfo(It.IsAny<string>()));
		var exporter = new SqliteExporter(dialogService.Object, optionsProvider.Object);

		exporter.ExportToSqlite(workbook.Object, options);

		await Assert.That(File.Exists(testDbPath)).IsTrue();

		using (var connection = new SQLiteConnection($"Data Source={testDbPath};Version=3;"))
		{
			connection.Open();
			using var command = new SQLiteCommand("PRAGMA table_info(TestSheet)", connection);
			using var reader = command.ExecuteReader();
			await Assert.That(reader.Read()).IsTrue();
			await Assert.That(reader.GetString(1)).IsEqualTo("Price");
			await Assert.That(reader.GetString(2)).IsEqualTo("REAL");
		}

		dialogService.VerifyAll();
	}

	[Test]
	public async Task ExportToSqlite_WithOptions_NoHeaders_UsesDefaultColumnNames()
	{
		var workbook = new Mock<Workbook>(MockBehavior.Loose);
		var worksheets = new Mock<Sheets>(MockBehavior.Loose);
		var worksheet = new Mock<Worksheet>(MockBehavior.Loose);
		var usedRange = new Mock<Range>(MockBehavior.Loose);

		var testData = new object[3, 3];
		testData[1, 1] = "Value1";
		testData[1, 2] = "Value2";
		testData[2, 1] = "Value3";
		testData[2, 2] = "Value4";

		worksheet.Setup(x => x.Name).Returns("TestSheet");
		worksheet.Setup(x => x.UsedRange).Returns(usedRange.Object);
		usedRange.Setup(x => x.Rows.Count).Returns(2);
		usedRange.Setup(x => x.Columns.Count).Returns(2);
		usedRange.Setup(x => x.get_Value(Missing.Value)).Returns(testData);

		worksheets.Setup(x => x.Count).Returns(1);
		worksheets.Setup(x => x.GetEnumerator()).Returns(new[] { worksheet.Object }.GetEnumerator());

		workbook.Setup(x => x.Worksheets).Returns(worksheets.Object);

		var options = new SqliteExportOptions
		{
			UseFirstRowAsHeaders = false,
			RowsToSkip = 0
		};

		dialogService
			.Setup(x => x.AskForSaveFile(FileType.Db))
			.Returns(testDbPath);

		dialogService
			.Setup(x => x.ShowInfo(It.IsAny<string>()));
		var exporter = new SqliteExporter(dialogService.Object, optionsProvider.Object);

		exporter.ExportToSqlite(workbook.Object, options);

		await Assert.That(File.Exists(testDbPath)).IsTrue();

		using (var connection = new SQLiteConnection($"Data Source={testDbPath};Version=3;"))
		{
			connection.Open();
			using (var command = new SQLiteCommand("PRAGMA table_info(TestSheet)", connection))
			using (var reader = command.ExecuteReader())
			{
				await Assert.That(reader.Read()).IsTrue();
				await Assert.That(reader.GetString(1)).IsEqualTo("Column1");
				await Assert.That(reader.Read()).IsTrue();
				await Assert.That(reader.GetString(1)).IsEqualTo("Column2");
			}

			using (var command = new SQLiteCommand("SELECT COUNT(*) FROM TestSheet", connection))
			{
				var count = (long)command.ExecuteScalar();
				await Assert.That(count).IsEqualTo(2); // Both rows should be data
			}
		}

		dialogService.VerifyAll();
	}
}
