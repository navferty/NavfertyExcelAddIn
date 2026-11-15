using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

using NavfertyExcelAddIn.Localization;

namespace NavfertyExcelAddIn.SqliteImport;

internal sealed class SqliteImportForm : Form
{
	private readonly CheckedListBox objects = new CheckedListBox();
	private readonly TextBox query = new TextBox();
	public IEnumerable<SqliteObject> SelectedObjects => objects.CheckedItems.Cast<SqliteObject>();
	public string Query => query.Text.Trim();

	private static string TextFor(string key) => UIStrings.ResourceManager.GetString(key) ?? key;

	public SqliteImportForm(IReadOnlyList<SqliteObject> items)
	{
		Text = TextFor("SqliteImportTitle");
		StartPosition = FormStartPosition.CenterScreen;
		Size = new Size(600, 440);
		MinimumSize = new Size(450, 350);
		var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 6, Padding = new Padding(10) };
		layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
		layout.RowStyles.Add(new RowStyle(SizeType.Percent, 60));
		layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
		layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
		layout.RowStyles.Add(new RowStyle(SizeType.Percent, 40));
		layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
		var selectionControls = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
		selectionControls.Controls.Add(new Label { Text = TextFor("SqliteImportObjects"), AutoSize = true, Anchor = AnchorStyles.Left });
		var selectAll = new Button { Name = "selectAll", Text = TextFor("SqliteImportSelectAll"), AutoSize = true };
		selectAll.Click += (sender, args) =>
		{
			for (int i = 0; i < objects.Items.Count; i++)
				objects.SetItemChecked(i, true);
		};
		selectionControls.Controls.Add(selectAll);
		layout.Controls.Add(selectionControls, 0, 0);
		objects.Dock = DockStyle.Fill;
		objects.CheckOnClick = true;
		foreach (var item in items) objects.Items.Add(item);
		layout.Controls.Add(objects, 0, 1);
		if (items.Any(item => item.Count > SqliteDatabase.MaxRows))
		{
			layout.Controls.Add(new Label
			{
				Name = "rowLimitWarning",
				Text = string.Format(TextFor("SqliteImportRowLimitWarning"), SqliteDatabase.MaxRows.ToString("N0")),
				AutoSize = true,
				ForeColor = Color.DarkRed,
				MaximumSize = new Size(400, 0)
			}, 0, 2);
		}
		layout.Controls.Add(new Label { Text = TextFor("SqliteImportQuery"), AutoSize = true }, 0, 3);
		query.Multiline = true;
		query.ScrollBars = ScrollBars.Vertical;
		query.Dock = DockStyle.Fill;
		layout.Controls.Add(query, 0, 4);
		var buttons = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, FlowDirection = FlowDirection.RightToLeft };
		var cancel = new Button { Text = TextFor("SqliteImportCancel"), DialogResult = DialogResult.Cancel, AutoSize = true };
		var ok = new Button { Text = TextFor("SqliteImportButton"), AutoSize = true };
		ok.Click += (sender, args) =>
		{
			if (Query.Length == 0 && objects.CheckedItems.Count == 0)
			{
				MessageBox.Show(TextFor("SqliteImportSelect"), Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			DialogResult = DialogResult.OK;
		};
		buttons.Controls.Add(cancel);
		buttons.Controls.Add(ok);
		layout.Controls.Add(buttons, 0, 5);
		Controls.Add(layout);
		AcceptButton = ok;
		CancelButton = cancel;
	}
}
