using System.Reflection;
using System.Windows.Forms;

using NavfertyExcelAddIn.SqliteImport;

namespace NavfertyExcelAddIn.UnitTests.SqliteImport;

public class SqliteImportFormTests
{
	private static readonly Type FormType = typeof(SqliteDatabase).Assembly.GetType("NavfertyExcelAddIn.SqliteImport.SqliteImportForm", true)!;

	private static void OnStaThread(Action action)
	{
		Exception? error = null;
		var thread = new Thread(() =>
		{
			try { action(); }
			catch (Exception ex) { error = ex; }
		});
		thread.SetApartmentState(ApartmentState.STA);
		thread.Start();
		thread.Join();
		if (error != null) System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(error).Throw();
	}

	private static Form CreateForm(SqliteObject[] items)
	{
		var form = (Form)Activator.CreateInstance(FormType, new object[] { items })!;
		form.Show();
		return form;
	}

	private static T Control<T>(Form form, string name) where T : Control =>
		(T)form.Controls.Find(name, true).Single();

	private static void Click(Button button) =>
		typeof(Button).GetMethod("OnClick", BindingFlags.Instance | BindingFlags.NonPublic)!
			.Invoke(button, new object[] { EventArgs.Empty });

	private static string Selected(Form form) =>
		string.Join(",", ((IEnumerable<SqliteObject>)FormType.GetProperty("SelectedObjects")!.GetValue(form)!)
			.Select(item => item.Name));

	private static void ApplyCount(Form form, SqliteObject[] items, int index, long? count) =>
		FormType.GetMethod("ApplyCount", BindingFlags.Instance | BindingFlags.NonPublic)!
			.Invoke(form, new object?[] { items, index, count });

	[Test]
	public async Task CountUpdatesPreserveSelectionAndUpdateRowsAndWarning()
	{
		string selected = "", countText = "", unavailableText = "", loadingText = "";
		bool warningInitially = true, warningAfterCount = false, warningAfterUncheck = true;
		bool unavailable = false, tagsMatch = false, checkPreserved = false;
		OnStaThread(() =>
		{
			var items = new[] { new SqliteObject("large", "table"), new SqliteObject("broken", "view") };
			using var form = CreateForm(items);
			var list = Control<ListView>(form, "objects");
			var warning = Control<Label>(form, "rowLimitWarning");
			tagsMatch = list.Columns.Count >= 3 && ReferenceEquals(list.Items[0].Tag, items[0]) &&
				ReferenceEquals(list.Items[1].Tag, items[1]);
			loadingText = list.Items[1].SubItems[2].Text;
			list.Items[0].Checked = true;
			warningInitially = warning.Visible;
			ApplyCount(form, items, 0, (long)SqliteDatabase.MaxRows + 1);
			ApplyCount(form, items, 1, null);
			selected = Selected(form);
			checkPreserved = list.Items[0].Checked && !list.Items[1].Checked;
			warningAfterCount = warning.Visible;
			countText = list.Items[0].SubItems[2].Text;
			unavailableText = list.Items[1].SubItems[2].Text;
			unavailable = items[1].CountUnavailable;
			list.Items[0].Checked = false;
			warningAfterUncheck = warning.Visible;
		});
		await Assert.That(tagsMatch).IsTrue();
		await Assert.That(selected).IsEqualTo("large");
		await Assert.That(checkPreserved).IsTrue();
		await Assert.That(warningInitially).IsFalse();
		await Assert.That(warningAfterCount).IsTrue();
		await Assert.That(warningAfterUncheck).IsFalse();
		await Assert.That(countText).Contains(((long)SqliteDatabase.MaxRows + 1).ToString("N0"));
		await Assert.That(unavailable).IsTrue();
		await Assert.That(unavailableText == loadingText).IsFalse();
	}

	[Test]
	public async Task SelectAllSkipsKnownOversizedObjectsAndClearAllClearsSelection()
	{
		string selected = "", afterIndividualChange = "", afterClear = "";
		string emptySummary = "", selectedSummary = "", clearedSummary = "", emptySelection = "not empty";
		bool warningAtBoundary = true, warningForLargeView = false, warningAfterClear = true;
		OnStaThread(() =>
		{
			var items = new[]
			{
				new SqliteObject("boundary", "table"), new SqliteObject("large", "view"),
				new SqliteObject("unknown", "view"), new SqliteObject("unavailable", "table")
			};
			using var form = CreateForm(items);
			var list = Control<ListView>(form, "objects");
			var warning = Control<Label>(form, "rowLimitWarning");
			var summary = Control<Label>(form, "selectionSummary");
			emptySummary = summary.Text;
			ApplyCount(form, items, 0, SqliteDatabase.MaxRows);
			ApplyCount(form, items, 1, (long)SqliteDatabase.MaxRows + 1);
			ApplyCount(form, items, 3, null);
			Click(Control<Button>(form, "selectAll"));
			selected = Selected(form);
			selectedSummary = summary.Text;
			warningAtBoundary = warning.Visible;
			list.Items[1].Checked = true;
			warningForLargeView = warning.Visible;
			list.Items[0].Checked = false;
			afterIndividualChange = Selected(form);
			Click(Control<Button>(form, "clearAll"));
			afterClear = Selected(form);
			clearedSummary = summary.Text;
			warningAfterClear = warning.Visible;
			using var emptyForm = CreateForm(Array.Empty<SqliteObject>());
			Click(Control<Button>(emptyForm, "selectAll"));
			emptySelection = Selected(emptyForm);
		});
		await Assert.That(selected).IsEqualTo("boundary,unknown,unavailable");
		await Assert.That(warningAtBoundary).IsFalse();
		await Assert.That(warningForLargeView).IsTrue();
		await Assert.That(afterIndividualChange).IsEqualTo("large,unknown,unavailable");
		await Assert.That(afterClear).IsEqualTo("");
		await Assert.That(selectedSummary == emptySummary).IsFalse();
		await Assert.That(selectedSummary).Contains("3");
		await Assert.That(clearedSummary).IsEqualTo(emptySummary);
		await Assert.That(warningAfterClear).IsFalse();
		await Assert.That(emptySelection).IsEqualTo("");
	}

	[Test]
	public async Task OversizedWarningTracksAllSelectedObjectsIncludingHiddenOnes()
	{
		string first = "", second = "", both = "", afterDeselect = "", hidden = "", afterCountChange = "", afterSecondDeselect = "";
		string firstCount = "", secondCount = "", updatedCount = "", limit = "";
		bool afterLastDeselect = true, afterCountBecomesSafe = true;
		OnStaThread(() =>
		{
			var items = new[] { new SqliteObject("first", "table"), new SqliteObject("second", "view") };
			using var form = CreateForm(items);
			var list = Control<ListView>(form, "objects");
			var warning = Control<Label>(form, "rowLimitWarning");
			long count1 = (long)SqliteDatabase.MaxRows + 123;
			long count2 = (long)SqliteDatabase.MaxRows + 456;
			long updated = count2 + 789;
			firstCount = count1.ToString("N0");
			secondCount = count2.ToString("N0");
			updatedCount = updated.ToString("N0");
			limit = SqliteDatabase.MaxRows.ToString("N0");
			ApplyCount(form, items, 0, count1);
			ApplyCount(form, items, 1, count2);
			list.Items[1].Checked = true;
			second = warning.Text;
			list.Items[1].Checked = false;
			list.Items[0].Checked = true;
			first = warning.Text;
			list.Items[1].Checked = true;
			both = warning.Text;
			list.Items[0].Checked = false;
			afterDeselect = warning.Visible ? warning.Text : "";
			Control<TextBox>(form, "searchBox").Text = "first";
			hidden = warning.Visible ? warning.Text : "";
			ApplyCount(form, items, 1, updated);
			afterCountChange = warning.Visible ? warning.Text : "";
			Control<TextBox>(form, "searchBox").Text = "";
			list.Items[0].Checked = true;
			list.Items[1].Checked = false;
			afterSecondDeselect = warning.Visible ? warning.Text : "";
			list.Items[0].Checked = false;
			afterLastDeselect = warning.Visible;
			list.Items[1].Checked = true;
			ApplyCount(form, items, 1, SqliteDatabase.MaxRows);
			afterCountBecomesSafe = warning.Visible;
		});
		await Assert.That(first).Contains("first");
		await Assert.That(first).Contains(firstCount);
		await Assert.That(first).Contains(limit);
		await Assert.That(second).Contains("second");
		await Assert.That(second).Contains(secondCount);
		await Assert.That(second).Contains(limit);
		await Assert.That(second.Contains(firstCount)).IsFalse();
		await Assert.That(both).Contains(2.ToString("N0"));
		await Assert.That(both).Contains(secondCount);
		await Assert.That(both).Contains(limit);
		await Assert.That(both.Contains(firstCount)).IsFalse();
		await Assert.That(afterDeselect).IsEqualTo(second);
		await Assert.That(hidden).IsEqualTo(second);
		await Assert.That(afterCountChange).Contains("second");
		await Assert.That(afterCountChange).Contains(updatedCount);
		await Assert.That(afterSecondDeselect).IsEqualTo(first);
		await Assert.That(afterLastDeselect).IsFalse();
		await Assert.That(afterCountBecomesSafe).IsFalse();
	}

	[Test]
	public async Task FilteringPreservesHiddenSelectionAndCountUpdates()
	{
		string whileFiltered = "", afterFiltering = "", afterClear = "";
		int visibleCount = -1;
		bool hiddenWarning = false, restoredCheck = false;
		OnStaThread(() =>
		{
			var items = new[] { new SqliteObject("alpha", "table"), new SqliteObject("beta", "view") };
			using var form = CreateForm(items);
			var list = Control<ListView>(form, "objects");
			var search = Control<TextBox>(form, "searchBox");
			list.Items[0].Checked = true;
			search.Text = "BETA";
			visibleCount = list.Items.Count;
			ApplyCount(form, items, 0, (long)SqliteDatabase.MaxRows + 1);
			whileFiltered = Selected(form);
			hiddenWarning = Control<Label>(form, "rowLimitWarning").Visible;
			list.Items[0].Checked = true;
			search.Text = "";
			afterFiltering = Selected(form);
			restoredCheck = list.Items[0].Checked && list.Items[1].Checked;
			search.Text = "beta";
			Click(Control<Button>(form, "clearAll"));
			afterClear = Selected(form);
		});
		await Assert.That(visibleCount).IsEqualTo(1);
		await Assert.That(whileFiltered).IsEqualTo("alpha");
		await Assert.That(hiddenWarning).IsTrue();
		await Assert.That(afterFiltering).IsEqualTo("alpha,beta");
		await Assert.That(restoredCheck).IsTrue();
		await Assert.That(afterClear).IsEqualTo("");
	}

	[Test]
	public async Task EditingQueryClearsCountStatusAndCancelsActiveCount()
	{
		bool noDatabaseDisabled = false, databaseEnabled = false, cancelled = false, cancelledByButton = false;
		string statusAfterEdit = "", buttonAfterEdit = "", buttonBefore = "", statusAfterClick = "";
		OnStaThread(() =>
		{
			using var noDatabase = (Form)Activator.CreateInstance(FormType, new object[] { Array.Empty<SqliteObject>() })!;
			noDatabaseDisabled = !Control<Button>(noDatabase, "countQuery").Enabled;
			using var form = (Form)Activator.CreateInstance(FormType, new object[] { Array.Empty<SqliteObject>(), "test.db" })!;
			var button = Control<Button>(form, "countQuery");
			var status = Control<Label>(form, "queryCountStatus");
			var query = Control<TextBox>(form, "query");
			var field = FormType.GetField("queryCountCancellation", BindingFlags.Instance | BindingFlags.NonPublic)!;
			databaseEnabled = button.Enabled;
			buttonBefore = button.Text;
			using var cancellation = new CancellationTokenSource();
			field.SetValue(form, cancellation);
			button.Text = "Counting";
			status.Text = "Stale count";
			query.Text = "SELECT 1";
			cancelled = cancellation.IsCancellationRequested;
			statusAfterEdit = status.Text;
			buttonAfterEdit = button.Text;
			using var secondCancellation = new CancellationTokenSource();
			field.SetValue(form, secondCancellation);
			status.Text = "Counting";
			Click(button);
			cancelledByButton = secondCancellation.IsCancellationRequested;
			statusAfterClick = status.Text;
		});
		await Assert.That(noDatabaseDisabled).IsTrue();
		await Assert.That(databaseEnabled).IsTrue();
		await Assert.That(cancelled).IsTrue();
		await Assert.That(statusAfterEdit).IsEqualTo("");
		await Assert.That(buttonAfterEdit).IsEqualTo(buttonBefore);
		await Assert.That(cancelledByButton).IsTrue();
		await Assert.That(statusAfterClick).IsEqualTo("");
	}
}
