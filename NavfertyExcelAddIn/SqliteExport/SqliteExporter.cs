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
		private readonly ISqliteExportOptionsProvider optionsProvider;
		private readonly Logger logger = LogManager.GetCurrentClassLogger();

		public SqliteExporter(IDialogService dialogService, ISqliteExportOptionsProvider optionsProvider)
		{
			this.dialogService = dialogService;
			this.optionsProvider = optionsProvider;
		}

		public void ExportToSqlite(Workbook workbook)
		{
			if (workbook == null)
			{
				throw new ArgumentNullException(nameof(workbook));
			}

			// Get options from provider
			if (!optionsProvider.TryGetOptions(out var options))
			{
				return;
			}

			ExportToSqlite(workbook, options);
		}

		public void ExportToSqlite(Workbook workbook, SqliteExportOptions options)
		{
			if (workbook == null)
			{
				throw new ArgumentNullException(nameof(workbook));
			}

			if (options == null)
			{
				throw new ArgumentNullException(nameof(options));
			}

			var dbPath = dialogService.AskForSaveFile(FileType.Db);
			if (string.IsNullOrEmpty(dbPath))
			{
				return;
			}

			logger.Debug($"Exporting workbook to SQLite: {dbPath}");

			try
			{
				ExportWorkbookToDatabase(workbook, dbPath, options);
				dialogService.ShowInfo(UIStrings.SqliteExportSuccess);
			}
			catch (Exception ex)
			{
				logger.Error(ex, "Failed to export to SQLite");
				dialogService.ShowError($"{UIStrings.SqliteExportError}: {ex.Message}");
			}
		}

		private void ExportWorkbookToDatabase(Workbook workbook, string dbPath, SqliteExportOptions options)
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
						ExportWorksheetToTable(worksheet, connection, options);
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

		private void ExportWorksheetToTable(Worksheet worksheet, SQLiteConnection connection, SqliteExportOptions options)
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

			// Calculate header row index (1-based indexing)
			int headerRowIndex = 1 + options.RowsToSkip;
			
			// Validate header row is within bounds
			if (headerRowIndex > rowCount)
			{
				logger.Debug($"Worksheet {worksheet.Name} has no rows after skipping {options.RowsToSkip} rows, skipping");
				return;
			}
			
			// Calculate actual data start row (after skipped rows and optional header)
			int dataStartRow = headerRowIndex;
			string[] columnNames = null;

			if (options.UseFirstRowAsHeaders)
			{
				// Extract column names from the first non-skipped row
				columnNames = new string[colCount];
				for (int col = 1; col <= colCount; col++)
				{
					var headerValue = values[headerRowIndex, col];
					var headerText = headerValue != null ? headerValue.ToString().Trim() : string.Empty;
					columnNames[col - 1] = string.IsNullOrWhiteSpace(headerText) ? $"Column{col}" : SanitizeColumnName(headerText);
				}
				// Data starts from the row after the header
				dataStartRow = headerRowIndex + 1;
			}

			// Check if there's any data to export
			if (dataStartRow > rowCount)
			{
				logger.Debug($"Worksheet {worksheet.Name} has no data rows after skipping, skipping");
				return;
			}

			// Detect column types
			var columnTypes = new ColumnTypeDetector.SqliteColumnType[colCount];
			for (int col = 1; col <= colCount; col++)
			{
				columnTypes[col - 1] = ColumnTypeDetector.DetectColumnType(values, col, dataStartRow, rowCount);
			}

			CreateTable(connection, tableName, colCount, columnNames, columnTypes);

			InsertData(connection, tableName, values, dataStartRow, rowCount, colCount, columnTypes);

			logger.Debug($"Exported {rowCount - dataStartRow + 1} rows to table {tableName}");
		}

		private void CreateTable(SQLiteConnection connection, string tableName, int columnCount, string[] columnNames, ColumnTypeDetector.SqliteColumnType[] columnTypes)
		{
			var createTableSql = new StringBuilder();
			createTableSql.Append($"CREATE TABLE [{tableName}] (");

			for (var i = 1; i <= columnCount; i++)
			{
				var columnName = columnNames != null ? columnNames[i - 1] : $"Column{i}";
				var columnType = columnTypes[i - 1].ToString();
				createTableSql.Append($"[{columnName}] {columnType}");
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

		private void InsertData(SQLiteConnection connection, string tableName, object[,] values, int startRow, int endRow, int colCount, ColumnTypeDetector.SqliteColumnType[] columnTypes)
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
					for (var row = startRow; row <= endRow; row++)
					{
						command.Parameters.Clear();

						for (var col = 1; col <= colCount; col++)
						{
							var value = values[row, col];
							var columnType = columnTypes[col - 1];

							if (value == null || (value is string str && string.IsNullOrWhiteSpace(str)))
							{
								command.Parameters.AddWithValue($"@col{col}", DBNull.Value);
							}
							else if (columnType == ColumnTypeDetector.SqliteColumnType.INTEGER)
							{
								if (value is bool boolVal)
								{
									command.Parameters.AddWithValue($"@col{col}", boolVal ? 1 : 0);
								}
								else if (value is double dbl)
								{
									// Validate that it's actually an integer value before casting
									if (Math.Abs(dbl - Math.Round(dbl)) < ColumnTypeDetector.FloatingPointTolerance)
									{
										command.Parameters.AddWithValue($"@col{col}", (long)Math.Round(dbl));
									}
									else
									{
										// This shouldn't happen if type detection is correct, but handle it gracefully
										logger.Warn($"Non-integer value {dbl} in INTEGER column, converting to REAL");
										command.Parameters.AddWithValue($"@col{col}", dbl);
									}
								}
								else if (value is float flt)
								{
									if (Math.Abs(flt - Math.Round(flt)) < ColumnTypeDetector.FloatingPointTolerance)
									{
										command.Parameters.AddWithValue($"@col{col}", (long)Math.Round(flt));
									}
									else
									{
										logger.Warn($"Non-integer value {flt} in INTEGER column, converting to REAL");
										command.Parameters.AddWithValue($"@col{col}", flt);
									}
								}
								else if (value is decimal dec)
								{
									if (Math.Abs((double)(dec - Math.Round(dec))) < ColumnTypeDetector.FloatingPointTolerance)
									{
										command.Parameters.AddWithValue($"@col{col}", (long)dec);
									}
									else
									{
										logger.Warn($"Non-integer value {dec} in INTEGER column, converting to REAL");
										command.Parameters.AddWithValue($"@col{col}", (double)dec);
									}
								}
								else if (value is int || value is long || value is short || value is byte)
								{
									command.Parameters.AddWithValue($"@col{col}", Convert.ToInt64(value));
								}
								else if (long.TryParse(value.ToString(), out long parsed))
								{
									command.Parameters.AddWithValue($"@col{col}", parsed);
								}
								else
								{
									command.Parameters.AddWithValue($"@col{col}", value.ToString());
								}
							}
							else if (columnType == ColumnTypeDetector.SqliteColumnType.REAL)
							{
								if (value is double || value is float || value is decimal)
								{
									command.Parameters.AddWithValue($"@col{col}", Convert.ToDouble(value));
								}
								else if (double.TryParse(value.ToString(), out double parsed))
								{
									command.Parameters.AddWithValue($"@col{col}", parsed);
								}
								else
								{
									command.Parameters.AddWithValue($"@col{col}", value.ToString());
								}
							}
							else if (columnType == ColumnTypeDetector.SqliteColumnType.NUMERIC)
							{
								if (value is DateTime dateTime)
								{
									// Store DateTime as ISO 8601 string format which SQLite can parse
									command.Parameters.AddWithValue($"@col{col}", dateTime.ToString("yyyy-MM-dd HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture));
								}
								else
								{
									command.Parameters.AddWithValue($"@col{col}", value.ToString());
								}
							}
							else if (columnType == ColumnTypeDetector.SqliteColumnType.BLOB)
							{
								if (value is byte[] bytes)
								{
									command.Parameters.AddWithValue($"@col{col}", bytes);
								}
								else
								{
									command.Parameters.AddWithValue($"@col{col}", DBNull.Value);
								}
							}
							else
							{
								command.Parameters.AddWithValue($"@col{col}", value.ToString());
							}
						}

						command.ExecuteNonQuery();
					}
				}

				transaction.Commit();
			}
		}

		private string SanitizeTableName(string name)
		{
			return SanitizeName(name, "Sheet");
		}

		private string SanitizeColumnName(string name)
		{
			return SanitizeName(name, "Column");
		}

		private string SanitizeName(string name, string defaultName)
		{
			if (string.IsNullOrWhiteSpace(name))
			{
				return defaultName;
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
				return defaultName;
			}

			// Ensure the name doesn't start with a digit
			if (char.IsDigit(result[0]))
			{
				result = "_" + result;
			}

			return result;
		}
	}
}
