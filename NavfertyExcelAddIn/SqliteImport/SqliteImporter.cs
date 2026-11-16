using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

using Microsoft.Office.Interop.Excel;

using NavfertyExcelAddIn.Localization;

using ExcelApplication = Microsoft.Office.Interop.Excel.Application;

namespace NavfertyExcelAddIn.SqliteImport;

public sealed class SqliteImporter
{
	public void Import(ExcelApplication app)
	{
		using (var picker = new OpenFileDialog
		{
			Filter = "SQLite database (*.db;*.sqlite;*.sqlite3)|*.db;*.sqlite;*.sqlite3|All files (*.*)|*.*",
			CheckFileExists = true
		})
		{
			if (picker.ShowDialog() != DialogResult.OK) return;
			try
			{
				using (var database = new SqliteDatabase(picker.FileName))
				using (var form = new SqliteImportForm(database.GetObjects(), picker.FileName))
				{
					if (form.ShowDialog() != DialogResult.OK) return;
					ImportWorkbook(app, database, form.SelectedObjects.Select(x => x.Name).ToArray(), form.Query);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show($"{UIStrings.ResourceManager.GetString("SqliteImportError")}: {ex.Message}",
					UIStrings.Error, MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}
	}

	public void ImportWorkbook(ExcelApplication app, SqliteDatabase database, IReadOnlyList<string> tables, string query)
	{
		Workbook? workbook = null;
		try
		{
			workbook = app.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
			var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
			int index = 0;
			foreach (var table in tables)
			{
				var sheet = NewSheet(workbook, index++, table, names);
				WriteResult(sheet, (headers, rows) => database.ReadTable(table, headers, rows));
			}
			if (!string.IsNullOrWhiteSpace(query))
			{
				var sheet = NewSheet(workbook, index++, "Query", names);
				WriteResult(sheet, (headers, rows) => database.ReadQuery(query, headers, rows));
			}
			if (index == 0) throw new ArgumentException("Select a table or enter a query.");
		}
		catch
		{
			workbook?.Close(false);
			throw;
		}
	}

	private static Worksheet NewSheet(Workbook workbook, int index, string name, HashSet<string> names)
	{
		var sheet = index == 0
			? (Worksheet)workbook.Worksheets[1]
			: (Worksheet)workbook.Worksheets.Add(After: workbook.Worksheets[workbook.Worksheets.Count]);
		sheet.Name = GetSheetName(name, names);
		return sheet;
	}

	private static string GetSheetName(string name, HashSet<string> names)
	{
		var safe = new string(name.Select(c => "[]:*?/\\".IndexOf(c) >= 0 || char.IsControl(c) ? '_' : c).ToArray()).Trim('\'');
		if (safe.Length > 31) safe = safe.Substring(0, 31);
		safe = safe.Trim('\'');
		if (string.IsNullOrWhiteSpace(safe)) safe = "Sheet";
		if (string.Equals(safe, "History", StringComparison.OrdinalIgnoreCase)) safe = safe + "_";
		string candidate = safe;
		for (int suffix = 2; !names.Add(candidate); suffix++)
		{
			string ending = $" ({suffix})";
			candidate = safe.Substring(0, Math.Min(safe.Length, 31 - ending.Length)).TrimEnd('\'') + ending;
		}
		return candidate;
	}

	private static void WriteResult(Worksheet sheet, Action<Action<string[]>, Action<object[,], int>> read)
	{
		int columnCount = 0;
		int nextRow = 2;
		read(headers =>
		{
			columnCount = headers.Length;
			var values = new object[1, columnCount];
			for (int col = 0; col < columnCount; col++) values[0, col] = headers[col];
			var range = sheet.Range[sheet.Cells[1, 1], sheet.Cells[1, columnCount]];
			range.NumberFormat = "@";
			range.Value2 = values;
		}, (values, count) =>
		{
			var range = sheet.Range[sheet.Cells[nextRow, 1], sheet.Cells[nextRow + count - 1, columnCount]];
			range.NumberFormat = "@";
			range.Value2 = values;
			nextRow += count;
		});
	}
}
