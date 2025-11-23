using System;
using System.Data.SQLite;
using System.IO;

using Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;

using Moq;

using Navferty.Common;

using NavfertyExcelAddIn.SqliteExport;

namespace NavfertyExcelAddIn.UnitTests.SqliteExport
{
	[TestClass]
	public class SqliteExporterTests : TestsBase
	{
		private Mock<IDialogService> dialogService;
		private SqliteExporter exporter;
		private string testDbPath;

		[TestInitialize]
		public void BeforeEachTest()
		{
			dialogService = new Mock<IDialogService>(MockBehavior.Strict);
			exporter = new SqliteExporter(dialogService.Object);
			testDbPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid()}.db");
		}

		[TestCleanup]
		public void AfterEachTest()
		{
			if (File.Exists(testDbPath))
			{
				File.Delete(testDbPath);
			}
		}

		[TestMethod]
		public void ExportToSqlite_NullWorkbook_ThrowsArgumentNullException()
		{
			Assert.ThrowsException<ArgumentNullException>(() => exporter.ExportToSqlite(null));
		}

		[TestMethod]
		public void ExportToSqlite_NoDbPathSelected_DoesNothing()
		{
			var workbook = new Mock<Workbook>(MockBehavior.Strict);

			dialogService
				.Setup(x => x.AskForSaveFile(FileType.Db))
				.Returns(string.Empty);

			exporter.ExportToSqlite(workbook.Object);

			dialogService.VerifyAll();
		}

		[TestMethod]
		public void ExportToSqlite_ValidWorkbook_CreatesDatabase()
		{
			var workbook = new Mock<Workbook>(MockBehavior.Loose);
			var worksheets = new Mock<Sheets>(MockBehavior.Loose);
			var worksheet = new Mock<Worksheet>(MockBehavior.Loose);
			var usedRange = new Mock<Range>(MockBehavior.Loose);

			var testData = new object[3, 2];
			testData[1, 1] = "A1";
			testData[1, 2] = "B1";
			testData[2, 1] = "A2";
			testData[2, 2] = "B2";

			worksheet.Setup(x => x.Name).Returns("TestSheet");
			worksheet.Setup(x => x.UsedRange).Returns(usedRange.Object);
			usedRange.Setup(x => x.Rows.Count).Returns(2);
			usedRange.Setup(x => x.Columns.Count).Returns(2);
			usedRange.Setup(x => x.Value).Returns(testData);

			worksheets.Setup(x => x.Count).Returns(1);
			worksheets.Setup(x => x.GetEnumerator()).Returns(new[] { worksheet.Object }.GetEnumerator());

			workbook.Setup(x => x.Worksheets).Returns(worksheets.Object);

			dialogService
				.Setup(x => x.AskForSaveFile(FileType.Db))
				.Returns(testDbPath);

			dialogService
				.Setup(x => x.ShowInfo(It.IsAny<string>()));

			exporter.ExportToSqlite(workbook.Object);

			Assert.IsTrue(File.Exists(testDbPath));

			using (var connection = new SQLiteConnection($"Data Source={testDbPath};Version=3;"))
			{
				connection.Open();
				using (var command = new SQLiteCommand("SELECT name FROM sqlite_master WHERE type='table'", connection))
				using (var reader = command.ExecuteReader())
				{
					Assert.IsTrue(reader.Read());
					Assert.AreEqual("TestSheet", reader.GetString(0));
				}

				using (var command = new SQLiteCommand("SELECT COUNT(*) FROM TestSheet", connection))
				{
					var count = (long)command.ExecuteScalar();
					Assert.AreEqual(2, count);
				}
			}

			dialogService.VerifyAll();
		}

		[TestMethod]
		public void ExportToSqlite_SheetNameWithSpecialChars_SanitizesTableName()
		{
			var workbook = new Mock<Workbook>(MockBehavior.Loose);
			var worksheets = new Mock<Sheets>(MockBehavior.Loose);
			var worksheet = new Mock<Worksheet>(MockBehavior.Loose);
			var usedRange = new Mock<Range>(MockBehavior.Loose);

			var testData = new object[2, 1];
			testData[1, 1] = "Value";

			worksheet.Setup(x => x.Name).Returns("Test-Sheet!@#");
			worksheet.Setup(x => x.UsedRange).Returns(usedRange.Object);
			usedRange.Setup(x => x.Rows.Count).Returns(1);
			usedRange.Setup(x => x.Columns.Count).Returns(1);
			usedRange.Setup(x => x.Value).Returns(testData);

			worksheets.Setup(x => x.Count).Returns(1);
			worksheets.Setup(x => x.GetEnumerator()).Returns(new[] { worksheet.Object }.GetEnumerator());

			workbook.Setup(x => x.Worksheets).Returns(worksheets.Object);

			dialogService
				.Setup(x => x.AskForSaveFile(FileType.Db))
				.Returns(testDbPath);

			dialogService
				.Setup(x => x.ShowInfo(It.IsAny<string>()));

			exporter.ExportToSqlite(workbook.Object);

			Assert.IsTrue(File.Exists(testDbPath));

			using (var connection = new SQLiteConnection($"Data Source={testDbPath};Version=3;"))
			{
				connection.Open();
				using (var command = new SQLiteCommand("SELECT name FROM sqlite_master WHERE type='table'", connection))
				using (var reader = command.ExecuteReader())
				{
					Assert.IsTrue(reader.Read());
					var tableName = reader.GetString(0);
					Assert.AreEqual("Test_Sheet___", tableName);
				}
			}

			dialogService.VerifyAll();
		}

		[TestMethod]
		public void ExportToSqlite_EmptyWorksheet_SkipsEmptySheet()
		{
			var workbook = new Mock<Workbook>(MockBehavior.Loose);
			var worksheets = new Mock<Sheets>(MockBehavior.Loose);
			var worksheet = new Mock<Worksheet>(MockBehavior.Loose);

			worksheet.Setup(x => x.Name).Returns("EmptySheet");
			worksheet.Setup(x => x.UsedRange).Returns((Range)null);

			worksheets.Setup(x => x.Count).Returns(1);
			worksheets.Setup(x => x.GetEnumerator()).Returns(new[] { worksheet.Object }.GetEnumerator());

			workbook.Setup(x => x.Worksheets).Returns(worksheets.Object);

			dialogService
				.Setup(x => x.AskForSaveFile(FileType.Db))
				.Returns(testDbPath);

			dialogService
				.Setup(x => x.ShowInfo(It.IsAny<string>()));

			exporter.ExportToSqlite(workbook.Object);

			Assert.IsTrue(File.Exists(testDbPath));

			using (var connection = new SQLiteConnection($"Data Source={testDbPath};Version=3;"))
			{
				connection.Open();
				using (var command = new SQLiteCommand("SELECT COUNT(*) FROM sqlite_master WHERE type='table'", connection))
				{
					var tableCount = (long)command.ExecuteScalar();
					Assert.AreEqual(0, tableCount);
				}
			}

			dialogService.VerifyAll();
		}
	}
}
