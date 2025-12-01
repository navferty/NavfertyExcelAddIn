using Microsoft.Office.Interop.Excel;

namespace NavfertyExcelAddIn.UnitTests.ParseNumerics;

[Category("Automation")]
[NotInParallel("Automation")]
public class ParseNumericsAutomationTests : AutomationTestsBase
{
	[Test]
	[Property("Description", "Create blank workbook, fill it with numeric strings and parse them to numbers")]
	public async Task ParseNumerics_NumericStrings_ConvertedToNumbers()
	{
		TestContext.Output.WriteLine("Testing ParseNumerics feature");

		var workbook = app.Workbooks.Add();
		Thread.Sleep(defaultSleep);

		var sheet = (Worksheet)workbook.Sheets[1];
		await Assert.That(sheet).IsNotNull();

		// Fill with numeric strings
		var values = new object[,]
		{
			{ "123", "456.78", "1000" },
			{ "1.5", "2.25", "100.99" },
			{ "-50", "0", "999.999" },
		};

		sheet.Range[sheet.Cells[1, 1], sheet.Cells[3, 3]].Value = values;
		sheet.UsedRange.Select();
		Thread.Sleep(defaultSleep);

		// Alt + ZX + N (ParseNumerics keytip)
		SendKeys.SendWait("%"); // Alt
		SendKeys.SendWait("ZXN");
		Thread.Sleep(defaultSleep);

		var currentValues = (object?[,])sheet.UsedRange.Cells.Value;
		await Assert.That(currentValues).IsNotNull();

		// Verify all values are now doubles
		for (int i = 0; i < 3; i++)
		{
			for (int j = 0; j < 3; j++)
			{
				var cellValue = currentValues[i + 1, j + 1];
				TestContext.Output.WriteLine($"Cell [{i + 1},{j + 1}]: {cellValue} (Type: {cellValue?.GetType().Name})");
				await Assert.That(cellValue is double).IsTrue();
			}
		}

		// Verify specific values
		await Assert.That(currentValues[1, 1]).IsEqualTo(123.0);
		await Assert.That(currentValues[1, 2]).IsEqualTo(456.78);
		await Assert.That(currentValues[3, 1]).IsEqualTo(-50.0);
	}

	[Test]
	[Category("Automation")]
	[Property("Description", "Test ParseNumerics with mixed content - strings, numbers, and text")]
	public async Task ParseNumerics_MixedContent_OnlyNumericStringsParsed()
	{
		TestContext.Output.WriteLine("Testing ParseNumerics with mixed content");

		var workbook = app.Workbooks.Add();
		Thread.Sleep(defaultSleep);

		var sheet = (Worksheet)workbook.Sheets[1];
		await Assert.That(sheet).IsNotNull();

		// Fill with mixed content
		var values = new object[,]
		{
			{ "100", "text", "200.5" },
			{ 50, "abc123", "300" },
		};

		sheet.Range[sheet.Cells[1, 1], sheet.Cells[2, 3]].Value = values;
		sheet.UsedRange.Select();
		Thread.Sleep(defaultSleep);

		// Alt + ZX + N (ParseNumerics keytip)
		SendKeys.SendWait("%"); // Alt
		SendKeys.SendWait("ZXN");
		Thread.Sleep(defaultSleep);

		var currentValues = (object?[,])sheet.UsedRange.Cells.Value;
		await Assert.That(currentValues).IsNotNull();

		// "100" should be converted to double
		await Assert.That(currentValues[1, 1]).IsEqualTo(100.0);
		await Assert.That(currentValues[1, 1] is double).IsTrue();

		// "text" should remain as string
		await Assert.That(currentValues[1, 2]).IsEqualTo("text");
		await Assert.That(currentValues[1, 2] is string).IsTrue();

		// "200.5" should be converted to double
		await Assert.That(currentValues[1, 3]).IsEqualTo(200.5);
		await Assert.That(currentValues[1, 3] is double).IsTrue();

		// 50 was already a number, should remain double
		await Assert.That(currentValues[2, 1]).IsEqualTo(50.0);
		await Assert.That(currentValues[2, 1] is double).IsTrue();

		// "abc123" should remain as string
		await Assert.That(currentValues[2, 2]).IsEqualTo("abc123");
		await Assert.That(currentValues[2, 2] is string).IsTrue();
	}
}
