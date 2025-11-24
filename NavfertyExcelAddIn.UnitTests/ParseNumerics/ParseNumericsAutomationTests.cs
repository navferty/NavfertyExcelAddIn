using System.Threading;
using System.Windows.Forms;

using Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;

using NavfertyExcelAddIn.UnitTests.SqliteExport;

namespace NavfertyExcelAddIn.UnitTests.ParseNumerics;

#nullable enable

[TestClass]
public class ParseNumericsAutomationTests : AutomationTestsBase
{
	[TestMethod]
	[TestCategory("Automation")]
	[Description("Create blank workbook, fill it with numeric strings and parse them to numbers")]
	public void ParseNumerics_NumericStrings_ConvertedToNumbers()
	{
		TestContext.WriteLine("Testing ParseNumerics feature");

		var workbook = app.Workbooks.Add();
		Thread.Sleep(defaultSleep);

		var sheet = (Worksheet)workbook.Sheets[1];
		Assert.IsNotNull(sheet);

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
		Assert.IsNotNull(currentValues);

		// Verify all values are now doubles
		for (int i = 0; i < 3; i++)
		{
			for (int j = 0; j < 3; j++)
			{
				var cellValue = currentValues[i + 1, j + 1];
				TestContext.WriteLine($"Cell [{i + 1},{j + 1}]: {cellValue} (Type: {cellValue?.GetType().Name})");
				Assert.IsInstanceOfType(cellValue, typeof(double), $"Cell [{i + 1},{j + 1}] should be double");
			}
		}

		// Verify specific values
		Assert.AreEqual(123.0, currentValues[1, 1]);
		Assert.AreEqual(456.78, currentValues[1, 2]);
		Assert.AreEqual(-50.0, currentValues[3, 1]);
	}

	[TestMethod]
	[TestCategory("Automation")]
	[Description("Test ParseNumerics with mixed content - strings, numbers, and text")]
	public void ParseNumerics_MixedContent_OnlyNumericStringsParsed()
	{
		TestContext.WriteLine("Testing ParseNumerics with mixed content");

		var workbook = app.Workbooks.Add();
		Thread.Sleep(defaultSleep);

		var sheet = (Worksheet)workbook.Sheets[1];
		Assert.IsNotNull(sheet);

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
		Assert.IsNotNull(currentValues);

		// "100" should be converted to double
		Assert.AreEqual(100.0, currentValues[1, 1]);
		Assert.IsInstanceOfType(currentValues[1, 1], typeof(double));

		// "text" should remain as string
		Assert.AreEqual("text", currentValues[1, 2]);
		Assert.IsInstanceOfType(currentValues[1, 2], typeof(string));

		// "200.5" should be converted to double
		Assert.AreEqual(200.5, currentValues[1, 3]);
		Assert.IsInstanceOfType(currentValues[1, 3], typeof(double));

		// 50 was already a number, should remain double
		Assert.AreEqual(50.0, currentValues[2, 1]);
		Assert.IsInstanceOfType(currentValues[2, 1], typeof(double));

		// "abc123" should remain as string
		Assert.AreEqual("abc123", currentValues[2, 2]);
		Assert.IsInstanceOfType(currentValues[2, 2], typeof(string));
	}
}
