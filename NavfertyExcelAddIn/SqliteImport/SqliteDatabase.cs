using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Threading;

using NavfertyExcelAddIn.Localization;

namespace NavfertyExcelAddIn.SqliteImport;

public sealed class SqliteDatabase : IDisposable
{
	public const int MaxRows = 1048575;
	public const int MaxColumns = 16384;
	private const int MaxCellLength = 32767;
	private readonly SQLiteConnection connection;

	public SqliteDatabase(string path)
	{
		var builder = new SQLiteConnectionStringBuilder
		{
			DataSource = path,
			ReadOnly = true,
			FailIfMissing = true,
			DefaultTimeout = 3
		};
		connection = new SQLiteConnection(builder.ConnectionString);
		try { connection.Open(); }
		catch { connection.Dispose(); throw; }
	}

	public IReadOnlyList<SqliteObject> GetObjects()
	{
		var objects = new List<SqliteObject>();
		using (var command = new SQLiteCommand("SELECT name, type FROM sqlite_master WHERE type IN ('table', 'view') AND name NOT LIKE 'sqlite\\_%' ESCAPE '\\' ORDER BY name", connection))
		using (var reader = command.ExecuteReader())
		{
			while (reader.Read())
				objects.Add(new SqliteObject(reader.GetString(0), reader.GetString(1)));
		}
		return objects;
	}

	public void LoadCounts(IReadOnlyList<SqliteObject> objects, Action<int, long?> report, CancellationToken cancellationToken)
	{
		SQLiteProgressEventHandler progress = (sender, args) =>
		{
			if (cancellationToken.IsCancellationRequested)
				args.ReturnCode = SQLiteProgressReturnCode.Interrupt;
		};
		connection.ProgressOps = 1000;
		connection.Progress += progress;
		try
		{
			using (cancellationToken.Register(() => connection.Cancel()))
			{
				for (int i = 0; i < objects.Count; i++)
				{
					cancellationToken.ThrowIfCancellationRequested();
					long? count;
					try
					{
						using (var command = new SQLiteCommand("SELECT COUNT(*) FROM " + QuoteIdentifier(objects[i].Name), connection))
							count = Convert.ToInt64(command.ExecuteScalar(), CultureInfo.InvariantCulture);
					}
					catch (SQLiteException)
					{
						cancellationToken.ThrowIfCancellationRequested();
						count = null;
					}
					cancellationToken.ThrowIfCancellationRequested();
					report(i, count);
				}
			}
		}
		finally { connection.Progress -= progress; }
	}

	public static string QuoteIdentifier(string name) => "\"" + name.Replace("\"", "\"\"") + "\"";

	public void ReadTable(string name, Action<string[]> headers, Action<object[,], int> rows)
		=> Read("SELECT * FROM " + QuoteIdentifier(name), headers, rows);

	public void ReadQuery(string query, Action<string[]> headers, Action<object[,], int> rows)
	{
		// Restrict the input to one SELECT statement. The connection is also read-only.
		if (string.IsNullOrWhiteSpace(query) || !Regex.IsMatch(query, @"^\s*SELECT\b", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant) || query.IndexOf(';') >= 0)
			throw new ArgumentException("Enter a single SELECT statement without a semicolon.");
		Read(query, headers, rows);
	}

	private void Read(string sql, Action<string[]> headers, Action<object[,], int> rows)
	{
		using (var command = new SQLiteCommand(sql, connection))
		using (var reader = command.ExecuteReader())
		{
			int columns = reader.FieldCount;
			if (columns == 0 || columns > MaxColumns)
				throw new InvalidOperationException("The result has too many columns for Excel.");
			var names = new string[columns];
			for (int col = 0; col < columns; col++)
				names[col] = CheckLength(reader.GetName(col));
			headers(names);
			int batchSize = Math.Max(1, Math.Min(500, 250000 / columns));
			var batch = new List<object[]>();
			int total = 0;
			while (reader.Read())
			{
				if (++total > MaxRows)
					throw new InvalidOperationException("The result has too many rows for an Excel worksheet.");
				var row = new object[columns];
				for (int col = 0; col < columns; col++)
					row[col] = ToExcelValue(ReadValue(reader, col));
				batch.Add(row);
				if (batch.Count == batchSize)
					Flush(batch, columns, rows);
			}
			if (batch.Count > 0)
				Flush(batch, columns, rows);
		}
	}

	private static object ReadValue(SQLiteDataReader reader, int column)
	{
		var affinity = reader.GetFieldAffinity(column);
		try { return reader.GetValue(column); }
		catch (Exception ex) when (ex is FormatException || ex is InvalidCastException || ex is OverflowException)
		{
			switch (affinity)
			{
				case TypeAffinity.Int64: return reader.GetInt64(column);
				case TypeAffinity.Double: return reader.GetDouble(column);
				case TypeAffinity.Text: return reader.GetString(column);
				case TypeAffinity.Null: return DBNull.Value;
				case TypeAffinity.Blob:
					long length = reader.GetBytes(column, 0, null, 0, 0);
					if (length > (MaxCellLength / 4) * 3)
						throw new InvalidOperationException("A BLOB exceeds Excel's 32,767 character cell limit when encoded as base64.");
					var bytes = new byte[(int)length];
					reader.GetBytes(column, 0, bytes, 0, bytes.Length);
					return bytes;
				default: throw;
			}
		}
	}

	private static void Flush(List<object[]> batch, int columns, Action<object[,], int> rows)
	{
		var values = new object[batch.Count, columns];
		for (int row = 0; row < batch.Count; row++)
			for (int col = 0; col < columns; col++)
				values[row, col] = batch[row][col];
		rows(values, batch.Count);
		batch.Clear();
	}

	private static object ToExcelValue(object value)
	{
		if (value == DBNull.Value) return null!;
		if (value is byte[] bytes)
		{
			if (bytes.Length > (MaxCellLength / 4) * 3)
				throw new InvalidOperationException("A BLOB exceeds Excel's 32,767 character cell limit when encoded as base64.");
			return Convert.ToBase64String(bytes);
		}
		if (value is string text) return CheckLength(text);
		if (value is long number)
			return number > 999999999999999L || number < -999999999999999L
				? (object)number.ToString(CultureInfo.InvariantCulture) : (double)number;
		if (value is DateTime date) return date.ToString("O", CultureInfo.InvariantCulture);
		return value;
	}

	private static string CheckLength(string value)
	{
		if (value.Length > MaxCellLength)
			throw new InvalidOperationException("A value exceeds Excel's 32,767 character cell limit.");
		return value;
	}

	public void Dispose() => connection.Dispose();
}

public sealed class SqliteObject
{
	public string Name { get; }
	public string Type { get; }
	public long? Count { get; internal set; }
	public bool CountUnavailable { get; internal set; }
	public SqliteObject(string name, string type) { Name = name; Type = type; }
	public override string ToString()
	{
		var count = Count.HasValue ? Count.Value.ToString("N0") :
			UIStrings.ResourceManager.GetString(CountUnavailable ? "SqliteImportCountUnavailable" : "SqliteImportCountLoading");
		return $"{Name} ({Type}, {count})";
	}
}
