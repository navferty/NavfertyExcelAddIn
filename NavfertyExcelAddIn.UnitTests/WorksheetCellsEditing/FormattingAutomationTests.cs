using System.Threading;
using System.Windows.Forms;

using Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;

using NavfertyExcelAddIn.UnitTests.SqliteExport;

namespace NavfertyExcelAddIn.UnitTests.WorksheetCellsEditing;

#nullable enable

[TestClass]
public class FormattingAutomationTests : AutomationTestsBase
{
	[TestMethod]
	[TestCategory("Automation")]
	[Description("Test ToggleCase feature - cycle through lowercase, UPPERCASE, and Titlecase")]
	public void ToggleCase_TextValues_CaseToggled()
	{
		TestContext.WriteLine("Testing ToggleCase feature");

		var workbook = app.Workbooks.Add();
		Thread.Sleep(defaultSleep);

		var sheet = (Worksheet)workbook.Sheets[1];
		Assert.IsNotNull(sheet);

		// Fill with lowercase text
		var values = new object[,]
		{
			{ "hello", "world", "test" },
			{ "one", "two", "three" },
		};

		sheet.Range[sheet.Cells[1, 1], sheet.Cells[2, 3]].Value = values;
		sheet.UsedRange.Select();
		Thread.Sleep(defaultSleep);

		// Alt + ZX + TC (ToggleCase keytip)
		SendKeys.SendWait("%"); // Alt
		SendKeys.SendWait("ZXTC");
		Thread.Sleep(defaultSleep);

		var currentValues = (object?[,])sheet.UsedRange.Cells.Value;
		Assert.IsNotNull(currentValues);

		foreach (var cell in currentValues)
		{
			TestContext.WriteLine(cell?.ToString() ?? "<null>");
		}

		// After first toggle: should be UPPERCASE
		var expectedValues = new object[,]
		{
			{ "HELLO", "WORLD", "TEST" },
			{ "ONE", "TWO", "THREE" },
		};

		CollectionAssert.AreEqual(expectedValues, currentValues);
	}

	[TestMethod]
	[TestCategory("Automation")]
	[Description("Test TrimExtraSpaces feature - remove extra spaces from text")]
	public void TrimExtraSpaces_TextWithSpaces_ExtraSpacesRemoved()
	{
		TestContext.WriteLine("Testing TrimExtraSpaces feature");

		var workbook = app.Workbooks.Add();
		Thread.Sleep(defaultSleep);

		var sheet = (Worksheet)workbook.Sheets[1];
		Assert.IsNotNull(sheet);

		// Fill with text containing extra spaces
		var values = new object[,]
		{
			{ "  hello  ", "world  test", "  trim  me  " },
			{ "one   two", "  three", "four  " },
		};

		sheet.Range[sheet.Cells[1, 1], sheet.Cells[2, 3]].Value = values;
		sheet.UsedRange.Select();
		Thread.Sleep(defaultSleep);

		// Alt + ZX + TS + X (TrimExtraSpaces keytip)
		SendKeys.SendWait("%"); // Alt
		SendKeys.SendWait("ZXTSX");
		Thread.Sleep(defaultSleep);

		var currentValues = (object?[,])sheet.UsedRange.Cells.Value;
		Assert.IsNotNull(currentValues);

		foreach (var cell in currentValues)
		{
			TestContext.WriteLine($"'{cell?.ToString() ?? "<null>"}'");
		}

		// Expected: extra spaces removed, single spaces preserved
		var expectedValues = new object[,]
		{
			{ "hello", "world test", "trim me" },
			{ "one two", "three", "four" },
		};

		CollectionAssert.AreEqual(expectedValues, currentValues);
	}

	[TestMethod]
	[TestCategory("Automation")]
	[Description("Test RemoveAllSpaces feature - remove all spaces from text")]
	public void RemoveAllSpaces_TextWithSpaces_AllSpacesRemoved()
	{
		TestContext.WriteLine("Testing RemoveAllSpaces feature");

		var workbook = app.Workbooks.Add();
		Thread.Sleep(defaultSleep);

		var sheet = (Worksheet)workbook.Sheets[1];
		Assert.IsNotNull(sheet);

		// Fill with text containing spaces
		var values = new object[,]
		{
			{ "hello world", "test data", "remove spaces" },
			{ "one two three", "a b c", "x y z" },
		};

		sheet.Range[sheet.Cells[1, 1], sheet.Cells[2, 3]].Value = values;
		sheet.UsedRange.Select();
		Thread.Sleep(defaultSleep);

		// Alt + ZX + TS + A (RemoveAllSpaces keytip)
		SendKeys.SendWait("%"); // Alt
		SendKeys.SendWait("ZXTSA");
		Thread.Sleep(defaultSleep);

		var currentValues = (object?[,])sheet.UsedRange.Cells.Value;
		Assert.IsNotNull(currentValues);

		foreach (var cell in currentValues)
		{
			TestContext.WriteLine(cell?.ToString() ?? "<null>");
		}

		// Expected: all spaces removed
		var expectedValues = new object[,]
		{
			{ "helloworld", "testdata", "removespaces" },
			{ "onetwothree", "abc", "xyz" },
		};

		CollectionAssert.AreEqual(expectedValues, currentValues);
	}

	[TestMethod]
	[TestCategory("Automation")]
	[Description("Test TrimSpaces (default) feature - remove extra spaces from text")]
	public void TrimSpaces_TextWithSpaces_ExtraSpacesRemoved()
	{
		TestContext.WriteLine("Testing TrimSpaces (default) feature");

		var workbook = app.Workbooks.Add();
		Thread.Sleep(defaultSleep);

		var sheet = (Worksheet)workbook.Sheets[1];
		Assert.IsNotNull(sheet);

		// Fill with text containing extra spaces
		var values = new object[,]
		{
			{ "  leading", "trailing  ", "  both  " },
		};

		sheet.Range[sheet.Cells[1, 1], sheet.Cells[1, 3]].Value = values;
		sheet.UsedRange.Select();
		Thread.Sleep(defaultSleep);

		// Alt + ZX + TS (TrimSpaces keytip - default action)
		SendKeys.SendWait("%"); // Alt
		SendKeys.SendWait("ZXTS");
		Thread.Sleep(defaultSleep);

		var currentValues = (object?[,])sheet.UsedRange.Cells.Value;
		Assert.IsNotNull(currentValues);

		foreach (var cell in currentValues)
		{
			TestContext.WriteLine($"'{cell?.ToString() ?? "<null>"}'");
		}

		// TrimSpaces calls TrimExtraSpaces
		var expectedValues = new object[,]
		{
			{ "leading", "trailing", "both" },
		};

		CollectionAssert.AreEqual(expectedValues, currentValues);
	}
}
