using System;
using System.Data.SQLite;
using System.IO;
using System.Text;

using Microsoft.Office.Interop.Excel;

using Navferty.Common;

using NavfertyExcelAddIn.Localization;

using NLog;

namespace NavfertyExcelAddIn.SqliteExport
{
	public class SqliteExporter : ISqliteExporter
	{
		private readonly IDialogService dialogService;
		private readonly Logger logger = LogManager.GetCurrentClassLogger();

		public SqliteExporter(IDialogService dialogService)
		{
			this.dialogService = dialogService;
		}

		public void ExportToSqlite(Workbook workbook)
		{
			if (workbook == null)
			{
				throw new ArgumentNullException(nameof(workbook));
			}

			var dbPath = dialogService.AskForSaveFile(FileType.Db);
			if (string.IsNullOrEmpty(dbPath))
			{
				return;
			}

			logger.Debug($"Exporting workbook to SQLite: {dbPath}");

			try
			{
				ExportWorkbookToDatabase(workbook, dbPath);
				dialogService.ShowInfo(UIStrings.SqliteExportSuccess);
			}
			catch (Exception ex)
			{
				logger.Error(ex, "Failed to export to SQLite");
				dialogService.ShowError($"{UIStrings.SqliteExportError}: {ex.Message}");
			}
		}

		private void ExportWorkbookToDatabase(Workbook workbook, string dbPath)
		{
			if (File.Exists(dbPath))
			{
				File.Delete(dbPath);
			}

			var connectionString = $"Data Source={dbPath};Version=3;";

			using (var connection = new SQLiteConnection(connectionString))
			{
				connection.Open();

				foreach (Worksheet worksheet in workbook.Worksheets)
				{
					try
					{
						logger.Debug($"Exporting worksheet: {worksheet.Name}");
						ExportWorksheetToTable(worksheet, connection);
					}
					catch (Exception ex)
					{
						logger.Error(ex, $"Failed to export worksheet: {worksheet.Name}");
						throw;
					}
				}
			}

			logger.Debug($"Successfully exported {workbook.Worksheets.Count} worksheets to {dbPath}");
		}

		private void ExportWorksheetToTable(Worksheet worksheet, SQLiteConnection connection)
		{
			var usedRange = worksheet.UsedRange;
			if (usedRange == null || usedRange.Rows.Count == 0)
			{
				logger.Debug($"Worksheet {worksheet.Name} is empty, skipping");
				return;
			}

			var tableName = SanitizeTableName(worksheet.Name);
			var rowCount = usedRange.Rows.Count;
			var colCount = usedRange.Columns.Count;

			var values = (object[,])usedRange.Value;
			if (values == null)
			{
				logger.Debug($"Worksheet {worksheet.Name} has no values, skipping");
				return;
			}

			CreateTable(connection, tableName, colCount);

			InsertData(connection, tableName, values, rowCount, colCount);

			logger.Debug($"Exported {rowCount} rows to table {tableName}");
		}

		private void CreateTable(SQLiteConnection connection, string tableName, int columnCount)
		{
			var createTableSql = new StringBuilder();
			createTableSql.Append($"CREATE TABLE [{tableName}] (");

			for (var i = 1; i <= columnCount; i++)
			{
				createTableSql.Append($"[Column{i}] TEXT");
				if (i < columnCount)
				{
					createTableSql.Append(", ");
				}
			}

			createTableSql.Append(")");

			using (var command = new SQLiteCommand(createTableSql.ToString(), connection))
			{
				command.ExecuteNonQuery();
			}
		}

		private void InsertData(SQLiteConnection connection, string tableName, object[,] values, int rowCount, int colCount)
		{
			using (var transaction = connection.BeginTransaction())
			{
				var insertSql = new StringBuilder();
				insertSql.Append($"INSERT INTO [{tableName}] VALUES (");

				for (var i = 1; i <= colCount; i++)
				{
					insertSql.Append($"@col{i}");
					if (i < colCount)
					{
						insertSql.Append(", ");
					}
				}

				insertSql.Append(")");

				using (var command = new SQLiteCommand(insertSql.ToString(), connection))
				{
					for (var row = 1; row <= rowCount; row++)
					{
						command.Parameters.Clear();

						for (var col = 1; col <= colCount; col++)
						{
							var value = values[row, col];
							var stringValue = value != null ? value.ToString() : string.Empty;
							command.Parameters.AddWithValue($"@col{col}", stringValue);
						}

						command.ExecuteNonQuery();
					}
				}

				transaction.Commit();
			}
		}

		private string SanitizeTableName(string name)
		{
			if (string.IsNullOrWhiteSpace(name))
			{
				return "Sheet";
			}

			var sanitized = new StringBuilder();
			foreach (var c in name)
			{
				if (char.IsLetterOrDigit(c) || c == '_')
				{
					sanitized.Append(c);
				}
				else
				{
					sanitized.Append('_');
				}
			}

			var result = sanitized.ToString();
			if (string.IsNullOrWhiteSpace(result))
			{
				return "Sheet";
			}

			if (char.IsDigit(result[0]))
			{
				result = "_" + result;
			}

			return result;
		}
	}
}
