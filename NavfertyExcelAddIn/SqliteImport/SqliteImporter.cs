using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Windows.Forms;

using Microsoft.Office.Interop.Excel;

using NavfertyExcelAddIn.Localization;

using Action = System.Action;
using ExcelApplication = Microsoft.Office.Interop.Excel.Application;

namespace NavfertyExcelAddIn.SqliteImport;

public sealed class SqliteImporter
{
	private static int importing;

	public void Import(ExcelApplication app)
	{
		if (Interlocked.CompareExchange(ref importing, 1, 0) != 0) return;
		try
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
						var tables = form.SelectedObjects.Select(x => x.Name).ToArray();
						var query = form.Query;
						using (var progress = new SqliteImportProgressForm((dispatch, report, complete, token) =>
							ImportWorkbookCore(app, database, tables, query, dispatch, report, complete, token)))
						{
							progress.ShowDialog(new ExcelWindow(app.Hwnd));
							if (progress.Error != null) throw progress.Error;
						}
					}
				}
				catch (OperationCanceledException) { }
				catch (Exception ex)
				{
					MessageBox.Show($"{UIStrings.ResourceManager.GetString("SqliteImportError")}: {ex.Message}",
						UIStrings.Error, MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
			}
		}
		finally { Interlocked.Exchange(ref importing, 0); }
	}

	public void ImportWorkbook(ExcelApplication app, SqliteDatabase database, IReadOnlyList<string> tables, string query)
	{
		if (Interlocked.CompareExchange(ref importing, 1, 0) != 0)
			throw new InvalidOperationException("A SQLite import is already in progress.");
		try
		{
			ImportWorkbookCore(app, database, tables, query, action => action(), (name, count) => { }, () => { }, CancellationToken.None);
		}
		finally { Interlocked.Exchange(ref importing, 0); }
	}

	private static void ImportWorkbookCore(ExcelApplication app, SqliteDatabase database, IReadOnlyList<string> tables,
		string query, Action<Action> dispatch, Action<string, int> report, Action complete, CancellationToken token)
	{
		Workbook? workbook = null;
		try
		{
			dispatch(() =>
			{
				token.ThrowIfCancellationRequested();
				workbook = app.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
			});
			var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
			int index = 0;
			foreach (var table in tables)
			{
				token.ThrowIfCancellationRequested();
				Worksheet? sheet = null;
				dispatch(() =>
				{
					token.ThrowIfCancellationRequested();
					sheet = NewSheet(workbook!, index++, table, names);
					report(table, 0);
				});
				WriteResult(sheet!, (headers, rows) => database.ReadTable(table, headers, rows, token), dispatch, report, table, token);
			}
			if (!string.IsNullOrWhiteSpace(query))
			{
				token.ThrowIfCancellationRequested();
				Worksheet? sheet = null;
				dispatch(() =>
				{
					token.ThrowIfCancellationRequested();
					sheet = NewSheet(workbook!, index++, "Query", names);
					report("Query", 0);
				});
				WriteResult(sheet!, (headers, rows) => database.ReadQuery(query, headers, rows, token), dispatch, report, "Query", token);
			}
			if (index == 0) throw new ArgumentException("Select a table or enter a query.");
			dispatch(() =>
			{
				token.ThrowIfCancellationRequested();
				complete();
			});
		}
		catch
		{
			dispatch(() => workbook?.Close(false));
			throw;
		}
	}

	private sealed class ExcelWindow : IWin32Window
	{
		public ExcelWindow(int handle) => Handle = new IntPtr(handle);
		public IntPtr Handle { get; }
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

	private static void WriteResult(Worksheet sheet, Action<Action<string[]>, Action<object[,], int>> read,
		Action<Action> dispatch, Action<string, int> report, string name, CancellationToken token)
	{
		int columnCount = 0;
		int nextRow = 2;
		read(headers =>
		{
			dispatch(() =>
			{
				token.ThrowIfCancellationRequested();
				columnCount = headers.Length;
				var values = new object[1, columnCount];
				for (int col = 0; col < columnCount; col++) values[0, col] = headers[col];
				var range = sheet.Range[sheet.Cells[1, 1], sheet.Cells[1, columnCount]];
				range.NumberFormat = "@";
				range.Value2 = values;
			});
		}, (values, count) =>
		{
			dispatch(() =>
			{
				token.ThrowIfCancellationRequested();
				var range = sheet.Range[sheet.Cells[nextRow, 1], sheet.Cells[nextRow + count - 1, columnCount]];
				range.NumberFormat = "@";
				range.Value2 = values;
				nextRow += count;
				report(name, nextRow - 2);
			});
		});
	}
}
