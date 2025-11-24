using System;
using System.Data.SQLite;
using System.IO;
using System.Threading;
using System.Windows.Forms;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace NavfertyExcelAddIn.UnitTests.SqliteExport;

[TestClass]
public class SqliteExportAutomationTest : AutomationTestsBase
{
	private string testDbPath;

	[TestInitialize]
	public override void Initialize()
	{
		testDbPath = Path.Combine(Path.GetTempPath(), $"test_automation_{Guid.NewGuid()}.db");
		base.Initialize();
	}

	[TestMethod]
	[TestCategory("Automation")]
	public void ExportToSqlite_ExpectedDataAndDataTypes()
	{
		var workbookPath = GetFilePath("SqliteExport/SqliteExportTestData.xlsx");

		Assert.IsTrue(File.Exists(workbookPath), $"Test data file not found: {workbookPath}");

		var workbook = app.Workbooks.Open(workbookPath);
		Thread.Sleep(defaultSleep.Add(defaultSleep));

		// send keys Alt separately, then Z, X: open the Navferty tab
		Thread.Sleep(defaultSleep);
		SendKeys.SendWait("%"); // Alt
		SendKeys.SendWait("ZXES");

		Thread.Sleep(defaultSleep);

		// Tab, arrow up (to input 1 in skip rows), enter
		SendKeys.SendWait("{TAB}{UP}");
		Thread.Sleep(defaultSleep);
		SendKeys.SendWait("{ENTER}");
		Thread.Sleep(defaultSleep);

		// enter the database path and enter to confirm
		SendKeys.SendWait(testDbPath + "{ENTER}");
		Thread.Sleep(defaultSleep);

		// check text in message box to be success
		var messageText = GetMessageBoxText();
		TestContext.WriteLine($"Message box text: {messageText}");
		Assert.IsTrue(messageText.Contains("Workbook successfully exported to SQLite database"), "Export did not complete successfully");

		// send esc to confirm success message box
		SendKeys.SendWait("{ESC}");

		workbook.Close(false);

		Assert.IsTrue(File.Exists(testDbPath), "Database file was not created");

		using var connection = new SQLiteConnection($"Data Source={testDbPath};Version=3;");
		connection.Open();

		// Verify tables were created (one for each worksheet)
		using var command = new SQLiteCommand("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name", connection);
		using var reader = command.ExecuteReader();
		Assert.IsTrue(reader.HasRows, "No tables were created");

		while (reader.Read())
		{
			var tableName = reader.GetString(0);
			TestContext.WriteLine($"Found table: {tableName}");
			VerifyTableStructureAndData(connection, tableName);
		}
	}

	private void VerifyTableStructureAndData(SQLiteConnection connection, string tableName)
	{
		TestContext.WriteLine($"Verifying table: {tableName}");

		// Verify table has columns
		using (var command = new SQLiteCommand($"PRAGMA table_info([{tableName}])", connection))
		using (var reader = command.ExecuteReader())
		{
			Assert.IsTrue(reader.HasRows, $"Table {tableName} has no columns");
			
			int columnCount = 0;
			while (reader.Read())
			{
				columnCount++;
				var columnName = reader.GetString(1);
				var columnType = reader.GetString(2);
				TestContext.WriteLine($"  Column {columnCount}: {columnName} ({columnType})");
				
				// Verify column type is one of the expected SQLite types
				Assert.IsTrue(
					columnType == "TEXT" || 
					columnType == "INTEGER" || 
					columnType == "REAL" || 
					columnType == "NUMERIC" || 
					columnType == "BLOB",
					$"Invalid column type: {columnType}");
			}
			
			Assert.IsTrue(columnCount > 0, $"Table {tableName} has no columns");
		}

		// Verify table has data
		using (var command = new SQLiteCommand($"SELECT COUNT(*) FROM [{tableName}]", connection))
		{
			var rowCount = (long)command.ExecuteScalar();
			TestContext.WriteLine($"  Row count: {rowCount}");
			Assert.IsTrue(rowCount > 0, $"Table {tableName} has no data rows");
		}

		// Verify we can read data from the table
		using (var command = new SQLiteCommand($"SELECT * FROM [{tableName}] LIMIT 5", connection))
		using (var reader = command.ExecuteReader())
		{
			if (reader.HasRows)
			{
				TestContext.WriteLine($"  Sample data from {tableName}:");
				int rowNum = 0;
				while (reader.Read() && rowNum < 3)
				{
					rowNum++;
					var values = new object[reader.FieldCount];
					reader.GetValues(values);
					TestContext.WriteLine($"    Row {rowNum}: {string.Join(", ", values)}");
				}
			}
		}
	}

	[TestCleanup]
	public override void Cleanup()
	{
		if (File.Exists(testDbPath))
		{
			try
			{
				File.Delete(testDbPath);
			}
			catch (Exception ex)
			{
				TestContext.WriteLine($"Warning: Could not delete test database: {ex.Message}");
			}
		}

		base.Cleanup();
	}
}
