using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

using NavfertyExcelAddIn.Localization;

namespace NavfertyExcelAddIn.SqliteImport;

internal sealed class SqliteImportForm : Form
{
	private readonly ListView objects = new ListView { Name = "objects", View = View.Details, CheckBoxes = true, FullRowSelect = true, Dock = DockStyle.Fill };
	private readonly TextBox search = new TextBox { Name = "searchBox", Width = 180 };
	private readonly TextBox query = new TextBox { Name = "query", Multiline = true, ScrollBars = ScrollBars.Both, Dock = DockStyle.Fill, AcceptsReturn = true, WordWrap = false };
	private readonly Label selectionSummary = new Label { Name = "selectionSummary", AutoSize = true };
	private readonly Label rowLimitWarning = new Label { Name = "rowLimitWarning", AutoSize = true, ForeColor = Color.DarkRed, MaximumSize = new Size(700, 0), Visible = false };
	private readonly Label queryCountStatus = new Label { Name = "queryCountStatus", AutoSize = true, MaximumSize = new Size(700, 0) };
	private readonly Button countQuery = new Button { Name = "countQuery", AutoSize = true };
	private readonly CancellationTokenSource countCancellation = new CancellationTokenSource();
	private readonly IReadOnlyList<SqliteObject> items;
	private readonly HashSet<SqliteObject> selected = new HashSet<SqliteObject>();
	private readonly string? databasePath;
	private CancellationTokenSource? queryCountCancellation;
	private bool updatingList;
	public IEnumerable<SqliteObject> SelectedObjects => items.Where(selected.Contains);
	public string Query => query.Text.Trim();

	private static string TextFor(string key) => UIStrings.ResourceManager.GetString(key) ?? key;

	public SqliteImportForm(IReadOnlyList<SqliteObject> items) : this(items, null) { }

	public SqliteImportForm(IReadOnlyList<SqliteObject> items, string? databasePath)
	{
		this.items = items;
		this.databasePath = databasePath;
		Text = TextFor("SqliteImportTitle");
		StartPosition = FormStartPosition.CenterScreen;
		Size = new Size(760, 650);
		MinimumSize = new Size(560, 480);
		AutoScaleMode = AutoScaleMode.Font;
		var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 10, Padding = new Padding(10) };
		layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
		layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
		layout.RowStyles.Add(new RowStyle(SizeType.Percent, 55));
		layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
		layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
		layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
		layout.RowStyles.Add(new RowStyle(SizeType.Percent, 45));
		layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
		layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
		layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

		var fileLabel = new Label { AutoSize = true, Text = string.Format(TextFor("SqliteImportDatabase"), databasePath == null ? "" : Path.GetFileName(databasePath)), AccessibleName = TextFor("SqliteImportDatabaseName") };
		if (databasePath != null) new ToolTip().SetToolTip(fileLabel, databasePath);
		layout.Controls.Add(fileLabel, 0, 0);
		var selectionControls = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, WrapContents = true };
		selectionControls.Controls.Add(new Label { Text = TextFor("SqliteImportSearch"), AutoSize = true, Anchor = AnchorStyles.Left });
		search.AccessibleName = TextFor("SqliteImportSearch");
		search.TextChanged += (sender, args) => RefreshObjects();
		selectionControls.Controls.Add(search);
		var selectAll = new Button { Name = "selectAll", Text = TextFor("SqliteImportSelectAll"), AutoSize = true };
		selectAll.Click += (sender, args) =>
		{
			foreach (var item in items)
				if (!(item.Count > SqliteDatabase.MaxRows)) selected.Add(item);
			RefreshObjects();
		};
		selectionControls.Controls.Add(selectAll);
		var clearAll = new Button { Name = "clearAll", Text = TextFor("SqliteImportClearAll"), AutoSize = true };
		clearAll.Click += (sender, args) => { selected.Clear(); RefreshObjects(); };
		selectionControls.Controls.Add(clearAll);
		layout.Controls.Add(selectionControls, 0, 1);
		objects.AccessibleName = TextFor("SqliteImportObjects");
		objects.Columns.Add(TextFor("SqliteImportColumnName"), 300);
		objects.Columns.Add(TextFor("SqliteImportColumnType"), 100);
		objects.Columns.Add(TextFor("SqliteImportColumnRows"), 220);
		objects.ItemCheck += (sender, args) =>
		{
			if (updatingList) return;
			var item = (SqliteObject)objects.Items[args.Index].Tag;
			if (args.NewValue == CheckState.Checked) selected.Add(item); else selected.Remove(item);
			UpdateSelection();
		};
		layout.Controls.Add(objects, 0, 2);
		layout.Controls.Add(selectionSummary, 0, 3);
		layout.Controls.Add(rowLimitWarning, 0, 4);
		layout.Controls.Add(new Label { Text = TextFor("SqliteImportQueryHelp"), AutoSize = true, MaximumSize = new Size(700, 0) }, 0, 5);
		query.AccessibleName = TextFor("SqliteImportQuery");
		query.TextChanged += (sender, args) => InvalidateQueryCount();
		layout.Controls.Add(query, 0, 6);
		var queryControls = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
		countQuery.Text = TextFor("SqliteImportCountRows");
		countQuery.Enabled = databasePath != null;
		countQuery.Click += async (sender, args) =>
		{
			if (queryCountCancellation != null) { InvalidateQueryCount(); return; }
			await CountQueryAsync();
		};
		queryControls.Controls.Add(countQuery);
		queryControls.Controls.Add(queryCountStatus);
		layout.Controls.Add(queryControls, 0, 7);
		layout.Controls.Add(new Label { Text = TextFor("SqliteImportCountDisclaimer"), AutoSize = true, MaximumSize = new Size(700, 0) }, 0, 8);
		var buttons = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, FlowDirection = FlowDirection.RightToLeft };
		var cancel = new Button { Text = TextFor("SqliteImportCancel"), DialogResult = DialogResult.Cancel, AutoSize = true };
		var ok = new Button { Text = TextFor("SqliteImportButton"), AutoSize = true };
		ok.Click += (sender, args) =>
		{
			if (Query.Length == 0 && selected.Count == 0)
			{
				MessageBox.Show(TextFor("SqliteImportSelect"), Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			if (SelectedObjects.Any(item => !item.Count.HasValue) &&
				MessageBox.Show(TextFor("SqliteImportUnknownCountConfirm"), Text,
					MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes)
				return;
			DialogResult = DialogResult.OK;
		};
		buttons.Controls.Add(cancel);
		buttons.Controls.Add(ok);
		layout.Controls.Add(buttons, 0, 9);
		Controls.Add(layout);
		AcceptButton = ok;
		CancelButton = cancel;
		RefreshObjects();
		if (databasePath != null) Shown += async (sender, args) => await LoadCountsAsync(databasePath, items);
		FormClosed += (sender, args) => { countCancellation.Cancel(); InvalidateQueryCount(); };
	}

	private void RefreshObjects()
	{
		updatingList = true;
		objects.BeginUpdate();
		try
		{
			objects.Items.Clear();
			foreach (var item in items.Where(x => x.Name.IndexOf(search.Text, StringComparison.CurrentCultureIgnoreCase) >= 0))
			{
				var row = new ListViewItem(new[] { item.Name, item.Type, CountText(item) }) { Tag = item, Checked = selected.Contains(item) };
				objects.Items.Add(row);
			}
		}
		finally { objects.EndUpdate(); updatingList = false; }
		UpdateSelection();
	}

	private static string CountText(SqliteObject item)
	{
		var count = item.Count.HasValue ? item.Count.Value.ToString("N0") : TextFor(item.CountUnavailable ? "SqliteImportCountUnavailable" : "SqliteImportCountLoading");
		return item.Count > SqliteDatabase.MaxRows ? count + " " + TextFor("SqliteImportOverLimit") : count;
	}

	private void UpdateSelection()
	{
		selectionSummary.Text = string.Format(TextFor("SqliteImportSelectionSummary"), selected.Count);
		var oversized = items.Where(item => selected.Contains(item) && item.Count > SqliteDatabase.MaxRows).ToArray();
		rowLimitWarning.Visible = oversized.Length > 0;
		if (oversized.Length == 1)
			rowLimitWarning.Text = string.Format(TextFor("SqliteImportSingleLimitWarning"),
				oversized[0].Name, oversized[0].Count!.Value.ToString("N0"), SqliteDatabase.MaxRows.ToString("N0"));
		else if (oversized.Length > 1)
			rowLimitWarning.Text = string.Format(TextFor("SqliteImportMultipleLimitWarning"),
				oversized.Length.ToString("N0"), oversized.Max(item => item.Count!.Value).ToString("N0"), SqliteDatabase.MaxRows.ToString("N0"));
		else
			rowLimitWarning.Text = "";
	}

	private void InvalidateQueryCount()
	{
		var cancellation = queryCountCancellation;
		queryCountCancellation = null;
		cancellation?.Cancel();
		queryCountStatus.Text = "";
		countQuery.Text = TextFor("SqliteImportCountRows");
	}

	private async Task CountQueryAsync()
	{
		if (databasePath == null) return;
		string sql = Query;
		var cancellation = new CancellationTokenSource();
		queryCountCancellation = cancellation;
		countQuery.Text = TextFor("SqliteImportCancelCount");
		queryCountStatus.Text = TextFor("SqliteImportCounting");
		queryCountStatus.ForeColor = SystemColors.ControlText;
		try
		{
			long count = await Task.Run(() =>
			{
				using (var database = new SqliteDatabase(databasePath))
					return database.CountQuery(sql, cancellation.Token);
			});
			if (queryCountCancellation != cancellation || IsDisposed) return;
			queryCountStatus.Text = string.Format(TextFor("SqliteImportQueryCount"), count.ToString("N0"));
			if (count > SqliteDatabase.MaxRows)
			{
				queryCountStatus.Text += " " + string.Format(TextFor("SqliteImportQueryLimit"), SqliteDatabase.MaxRows.ToString("N0"));
				queryCountStatus.ForeColor = Color.DarkRed;
			}
		}
		catch (OperationCanceledException) { }
		catch (Exception ex)
		{
			if (queryCountCancellation != cancellation || IsDisposed) return;
			queryCountStatus.Text = string.Format(TextFor("SqliteImportQueryCountError"), ex.Message);
			queryCountStatus.ForeColor = Color.DarkRed;
		}
		finally
		{
			if (queryCountCancellation == cancellation)
			{
				queryCountCancellation = null;
				if (!IsDisposed) countQuery.Text = TextFor("SqliteImportCountRows");
			}
			cancellation.Dispose();
		}
	}

	private async Task LoadCountsAsync(string path, IReadOnlyList<SqliteObject> items)
	{
		var token = countCancellation.Token;
		var progress = new Progress<(int Index, long? Count)>(result =>
		{
			if (!token.IsCancellationRequested && !IsDisposed) ApplyCount(items, result.Index, result.Count);
		});
		try
		{
			await Task.Run(() =>
			{
				using (var database = new SqliteDatabase(path))
					database.LoadCounts(items, (index, count) => ((IProgress<(int, long?)>)progress).Report((index, count)), token);
			});
		}
		catch (OperationCanceledException) { }
		catch (Exception)
		{
			if (!token.IsCancellationRequested && !IsDisposed)
				for (int i = 0; i < items.Count; i++)
					if (!items[i].Count.HasValue) ((IProgress<(int, long?)>)progress).Report((i, null));
		}
	}

	private void ApplyCount(IReadOnlyList<SqliteObject> items, int index, long? count)
	{
		var item = items[index];
		item.Count = count;
		item.CountUnavailable = !count.HasValue;
		foreach (ListViewItem row in objects.Items)
			if (ReferenceEquals(row.Tag, item)) { row.SubItems[2].Text = CountText(item); break; }
		UpdateSelection();
	}
}
