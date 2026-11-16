using System;
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

using NavfertyExcelAddIn.Localization;

namespace NavfertyExcelAddIn.SqliteImport;

internal sealed class SqliteImportProgressForm : Form
{
	private readonly CancellationTokenSource cancellation = new CancellationTokenSource();
	private readonly Label status = new Label();
	private readonly Button cancel = new Button();
	private readonly Action<Action<Action>, Action<string, int>, Action, CancellationToken> import;
	private bool finished;
	private bool committed;

	public Exception? Error { get; private set; }

	public SqliteImportProgressForm(Action<Action<Action>, Action<string, int>, Action, CancellationToken> import)
	{
		this.import = import;
		Text = TextFor("SqliteImportProgressTitle", "Importing SQLite data");
		StartPosition = FormStartPosition.CenterParent;
		FormBorderStyle = FormBorderStyle.FixedDialog;
		MinimizeBox = false;
		MaximizeBox = false;
		ShowInTaskbar = false;
		ClientSize = new Size(420, 115);

		var layout = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(12), ColumnCount = 1, RowCount = 3 };
		layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
		layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
		layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
		status.Dock = DockStyle.Fill;
		status.AutoEllipsis = true;
		status.Text = TextFor("SqliteImportProgressTitle", "Importing SQLite data");
		layout.Controls.Add(status, 0, 0);
		layout.Controls.Add(new ProgressBar { Dock = DockStyle.Fill, Style = ProgressBarStyle.Marquee }, 0, 1);
		cancel.Text = TextFor("SqliteImportProgressCancel", "Cancel");
		cancel.AutoSize = true;
		cancel.Anchor = AnchorStyles.Right;
		cancel.Click += (sender, args) => RequestCancellation();
		layout.Controls.Add(cancel, 0, 2);
		Controls.Add(layout);
		CancelButton = cancel;
		Shown += async (sender, args) => await RunImportAsync();
	}

	private static string TextFor(string key, string fallback) => UIStrings.ResourceManager.GetString(key) ?? fallback;

	private void RequestCancellation()
	{
		if (finished || committed || cancellation.IsCancellationRequested) return;
		cancel.Enabled = false;
		status.Text = TextFor("SqliteImportProgressCancelling", "Cancelling...");
		cancellation.Cancel();
	}

	protected override void OnFormClosing(FormClosingEventArgs e)
	{
		if (!finished)
		{
			e.Cancel = true;
			RequestCancellation();
		}
		base.OnFormClosing(e);
	}

	private async Task RunImportAsync()
	{
		try
		{
			await Task.Run(() => import(
				action => Invoke(action),
				(name, count) => status.Text = string.Format(TextFor("SqliteImportProgressText", "{0}: {1:N0} rows"), name, count),
				() => { committed = true; cancel.Enabled = false; },
				cancellation.Token));
		}
		catch (OperationCanceledException) when (cancellation.IsCancellationRequested) { }
		catch (Exception ex) { Error = ex; }
		finally
		{
			finished = true;
			Close();
		}
	}

	protected override void Dispose(bool disposing)
	{
		if (disposing) cancellation.Dispose();
		base.Dispose(disposing);
	}
}
