using System.Windows.Forms;

namespace NavfertyExcelAddIn.SqliteExport;

#nullable enable
public class SqliteExportOptionsFormProvider : ISqliteExportOptionsProvider
{
	public bool TryGetOptions(out SqliteExportOptions? options)
	{
		using (var optionsForm = new SqliteExportOptionsForm())
		{
			if (optionsForm.ShowDialog() == DialogResult.OK)
			{
				options = optionsForm.Options;
				return true;
			}
		}

		options = null;
		return false;
	}
}
