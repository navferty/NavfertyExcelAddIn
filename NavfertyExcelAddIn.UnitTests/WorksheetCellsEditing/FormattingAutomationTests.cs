using Microsoft.Office.Interop.Excel;

namespace NavfertyExcelAddIn.UnitTests.WorksheetCellsEditing;

[Category("Automation")]
[NotInParallel("Automation")]
public class FormattingAutomationTests : AutomationTestsBase
{
	[Test]
	//[Description("Test ToggleCase feature - cycle through lowercase, UPPERCASE, and Titlecase")]
	public async Task ToggleCase_TextValues_CaseToggled()
	{
		TestContext.Output.WriteLine("Testing ToggleCase feature");

		var workbook = app.Workbooks.Add();
		Thread.Sleep(defaultSleep);

		var sheet = (Worksheet)workbook.Sheets[1];
		await Assert.That(sheet).IsNotNull();

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
		await Assert.That(currentValues).IsNotNull();

		foreach (var cell in currentValues)
		{
			TestContext.Output.WriteLine(cell?.ToString() ?? "<null>");
		}

		// After first toggle: should be UPPERCASE
		var expectedValues = new object[,]
		{
			{ "HELLO", "WORLD", "TEST" },
			{ "ONE", "TWO", "THREE" },
		};

		await Assert.That(currentValues).IsEquivalentTo(expectedValues);
	}

	[Test]
	//[Description("Test TrimExtraSpaces feature - remove extra spaces from text")]
	public async Task TrimExtraSpaces_TextWithSpaces_ExtraSpacesRemoved()
	{
		TestContext.Output.WriteLine("Testing TrimExtraSpaces feature");

		var workbook = app.Workbooks.Add();
		Thread.Sleep(defaultSleep);

		var sheet = (Worksheet)workbook.Sheets[1];
		await Assert.That(sheet).IsNotNull();

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
		SendKeys.SendWait("ZXTS{DOWN}{ENTER}");
		Thread.Sleep(defaultSleep);

		var currentValues = (object?[,])sheet.UsedRange.Cells.Value;
		await Assert.That(currentValues).IsNotNull();

		foreach (var cell in currentValues)
		{
			TestContext.Output.WriteLine($"'{cell?.ToString() ?? "<null>"}'");
		}

		// Expected: extra spaces removed, single spaces preserved
		var expectedValues = new object[,]
		{
			{ "hello", "world test", "trim me" },
			{ "one two", "three", "four" },
		};

		await Assert.That(currentValues).IsEquivalentTo(expectedValues);
	}

	[Test]
	//[Description("Test RemoveAllSpaces feature - remove all spaces from text")]
	public async Task RemoveAllSpaces_TextWithSpaces_AllSpacesRemoved()
	{
		TestContext.Output.WriteLine("Testing RemoveAllSpaces feature");

		var workbook = app.Workbooks.Add();
		Thread.Sleep(defaultSleep);

		var sheet = (Worksheet)workbook.Sheets[1];
		await Assert.That(sheet).IsNotNull();

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
		SendKeys.SendWait("ZXTS{ENTER}");
		Thread.Sleep(defaultSleep);

		var currentValues = (object?[,])sheet.UsedRange.Cells.Value;
		await Assert.That(currentValues).IsNotNull();

		foreach (var cell in currentValues)
		{
			TestContext.Output.WriteLine(cell?.ToString() ?? "<null>");
		}

		// Expected: all spaces removed
		var expectedValues = new object[,]
		{
			{ "helloworld", "testdata", "removespaces" },
			{ "onetwothree", "abc", "xyz" },
		};

		await Assert.That(currentValues).IsEquivalentTo(expectedValues);
	}

	[Test]
	//[Description("Test TrimSpaces (default) feature - remove extra spaces from text")]
	public async Task TrimSpaces_TextWithSpaces_ExtraSpacesRemoved()
	{
		TestContext.Output.WriteLine("Testing TrimSpaces (default) feature");

		var workbook = app.Workbooks.Add();
		Thread.Sleep(defaultSleep);

		var sheet = (Worksheet)workbook.Sheets[1];
		await Assert.That(sheet).IsNotNull();

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
		SendKeys.SendWait("ZXTS{DOWN}{ENTER}");
		Thread.Sleep(defaultSleep);

		var currentValues = (object?[,])sheet.UsedRange.Cells.Value;
		await Assert.That(currentValues).IsNotNull();

		foreach (var cell in currentValues)
		{
			TestContext.Output.WriteLine($"'{cell?.ToString() ?? "<null>"}'");
		}

		// TrimSpaces calls TrimExtraSpaces
		var expectedValues = new object[,]
		{
			{ "leading", "trailing", "both" },
		};

		await Assert.That(currentValues).IsEquivalentTo(expectedValues);
	}
}
