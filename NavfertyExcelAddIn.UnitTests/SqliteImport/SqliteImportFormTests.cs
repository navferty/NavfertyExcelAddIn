using System.Reflection;
using System.Windows.Forms;

using NavfertyExcelAddIn.SqliteImport;

namespace NavfertyExcelAddIn.UnitTests.SqliteImport;

public class SqliteImportFormTests
{
	[Test]
	public async Task RowLimitWarningAppearsOnlyWhenAnyTableOrViewExceedsLimit()
	{
		Exception? error = null;
		int[] warningCounts = [];
		string warningText = "";
		int checkedCount = -1;
		var thread = new Thread(() =>
		{
			try
			{
				var formType = typeof(SqliteDatabase).Assembly.GetType("NavfertyExcelAddIn.SqliteImport.SqliteImportForm", true)!;
				SqliteObject Item(string name, string type, long count)
				{
					var item = new SqliteObject(name, type);
					typeof(SqliteObject).GetProperty("Count")!.SetValue(item, count);
					return item;
				}
				var cases = new[]
				{
					Array.Empty<SqliteObject>(),
					new[] { Item("boundary", "table", SqliteDatabase.MaxRows) },
					new[] { Item("small", "table", 1), Item("large", "table", (long)SqliteDatabase.MaxRows + 1) },
					new[] { Item("large_view", "view", (long)SqliteDatabase.MaxRows + 1) }
				};
				var counts = new List<int>();
				foreach (var items in cases)
				{
					using var form = (Form)Activator.CreateInstance(formType, new object[] { items })!;
					var warnings = form.Controls.Find("rowLimitWarning", true);
					counts.Add(warnings.Length);
					if (warnings.Length != 0)
					{
						warningText = warnings[0].Text;
						var layout = (TableLayoutPanel)form.Controls[0];
						checkedCount = ((CheckedListBox)layout.GetControlFromPosition(0, 1)!).CheckedItems.Count;
					}
				}
				warningCounts = counts.ToArray();
			}
			catch (Exception ex) { error = ex; }
		});
		thread.SetApartmentState(ApartmentState.STA);
		thread.Start();
		thread.Join();
		if (error != null) throw error;
		await Assert.That(string.Join(",", warningCounts)).IsEqualTo("0,0,1,1");
		await Assert.That(warningText).Contains(SqliteDatabase.MaxRows.ToString("N0"));
		await Assert.That(checkedCount).IsEqualTo(0);
	}

	[Test]
	public async Task SelectAllChecksTablesAndViewsAndAllowsIndividualChanges()
	{
		Exception? error = null;
		string[] selected = [];
		string[] afterUnchecking = [];
		int emptyCount = -1;
		var thread = new Thread(() =>
		{
			try
			{
				var formType = typeof(SqliteDatabase).Assembly.GetType("NavfertyExcelAddIn.SqliteImport.SqliteImportForm", true)!;
				using var form = (Form)Activator.CreateInstance(formType, new object[] {
					new[] { new SqliteObject("first", "table"), new SqliteObject("second", "view") }
				})!;
				var layout = (TableLayoutPanel)form.Controls[0];
				var list = (CheckedListBox)layout.GetControlFromPosition(0, 1)!;
				var button = (Button)form.Controls.Find("selectAll", true).Single();
				list.SetItemChecked(0, true);
				typeof(Button).GetMethod("OnClick", BindingFlags.Instance | BindingFlags.NonPublic)!
					.Invoke(button, new object[] { EventArgs.Empty });
				selected = ((IEnumerable<SqliteObject>)formType.GetProperty("SelectedObjects")!.GetValue(form)!)
					.Select(x => x.Name).ToArray();
				list.SetItemChecked(0, false);
				afterUnchecking = ((IEnumerable<SqliteObject>)formType.GetProperty("SelectedObjects")!.GetValue(form)!)
					.Select(x => x.Name).ToArray();
				using var emptyForm = (Form)Activator.CreateInstance(formType, new object[] { Array.Empty<SqliteObject>() })!;
				var emptyButton = (Button)emptyForm.Controls.Find("selectAll", true).Single();
				typeof(Button).GetMethod("OnClick", BindingFlags.Instance | BindingFlags.NonPublic)!
					.Invoke(emptyButton, new object[] { EventArgs.Empty });
				emptyCount = ((IEnumerable<SqliteObject>)formType.GetProperty("SelectedObjects")!.GetValue(emptyForm)!).Count();
			}
			catch (Exception ex) { error = ex; }
		});
		thread.SetApartmentState(ApartmentState.STA);
		thread.Start();
		thread.Join();
		if (error != null) throw error;
		await Assert.That(string.Join(",", selected)).IsEqualTo("first,second");
		await Assert.That(string.Join(",", afterUnchecking)).IsEqualTo("second");
		await Assert.That(emptyCount).IsEqualTo(0);
	}
}
