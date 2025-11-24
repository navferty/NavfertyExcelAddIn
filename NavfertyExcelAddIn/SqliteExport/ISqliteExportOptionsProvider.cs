namespace NavfertyExcelAddIn.SqliteExport;

#nullable enable
public interface ISqliteExportOptionsProvider
{
	/// <summary>
	/// Gets the export options from the user.
	/// </summary>
	/// <param name="options">The export options if user confirmed, null if cancelled.</param>
	/// <returns>True if user confirmed the options, false if cancelled.</returns>
	bool TryGetOptions(out SqliteExportOptions? options);
}
