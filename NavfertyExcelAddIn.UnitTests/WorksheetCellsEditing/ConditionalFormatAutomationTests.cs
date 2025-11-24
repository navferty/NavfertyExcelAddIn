using System.Drawing;
using System.Threading;
using System.Windows.Forms;

using Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;

using NavfertyExcelAddIn.UnitTests.SqliteExport;

namespace NavfertyExcelAddIn.UnitTests.WorksheetCellsEditing;

#nullable enable

[TestClass]
public class ConditionalFormatAutomationTests : AutomationTestsBase
{
	[TestMethod]
	[TestCategory("Automation")]
	[Description("Test RepairConditionalFormat feature - copy conditional format from first row to entire range")]
	public void RepairConditionalFormat_WithConditionalFormatInFirstRow_CopiedToAllRows()
	{
		TestContext.WriteLine("Testing RepairConditionalFormat feature");

		var workbook = app.Workbooks.Add();
		Thread.Sleep(defaultSleep);

		var sheet = (Worksheet)workbook.Sheets[1];
		Assert.IsNotNull(sheet);

		// Fill with numeric data
		var values = new object[,]
		{
			{ 10, 20, 30 },
			{ 15, 25, 35 },
			{ 5, 15, 25 },
			{ 20, 30, 40 },
		};

		sheet.Range[sheet.Cells[1, 1], sheet.Cells[4, 3]].Value = values;

		// Add conditional formatting to the first row only
		var firstRow = sheet.Range[sheet.Cells[1, 1], sheet.Cells[1, 3]];
		var formatCondition = firstRow.FormatConditions.Add(
			Type: XlFormatConditionType.xlCellValue,
			Operator: XlFormatConditionOperator.xlGreater,
			Formula1: "15");
		
		formatCondition.Interior.Color = ColorTranslator.ToOle(Color.Yellow);

		Thread.Sleep(defaultSleep);

		// Verify initial state - only first row has conditional formatting
		Assert.IsTrue(firstRow.FormatConditions.Count > 0, "First row should have conditional formatting");
		
		var secondRow = sheet.Range[sheet.Cells[2, 1], sheet.Cells[2, 3]];
		Assert.AreEqual(0, secondRow.FormatConditions.Count, "Second row should not have conditional formatting initially");

		// Select the entire range
		sheet.Range[sheet.Cells[1, 1], sheet.Cells[4, 3]].Select();
		Thread.Sleep(defaultSleep);

		// Alt + ZX + CF (RepairConditionalFormat keytip)
		SendKeys.SendWait("%"); // Alt
		SendKeys.SendWait("ZXCF");
		Thread.Sleep(defaultSleep);

		// Verify conditional formatting was copied to all rows
		for (int row = 1; row <= 4; row++)
		{
			var currentRow = sheet.Range[sheet.Cells[row, 1], sheet.Cells[row, 3]];
			Assert.IsTrue(currentRow.FormatConditions.Count > 0, $"Row {row} should have conditional formatting");
			TestContext.WriteLine($"Row {row} has {currentRow.FormatConditions.Count} conditional format(s)");
		}
	}

	[TestMethod]
	[Ignore("Only first row repair for conditional formattin is implemented")]
	[TestCategory("Automation")]
	[Description("Test RepairConditionalFormat with conditional format in second row (first row empty)")]
	public void RepairConditionalFormat_WithConditionalFormatInSecondRow_CopiedToAllRows()
	{
		TestContext.WriteLine("Testing RepairConditionalFormat with format in second row");

		var workbook = app.Workbooks.Add();
		Thread.Sleep(defaultSleep);

		var sheet = (Worksheet)workbook.Sheets[1];
		Assert.IsNotNull(sheet);

		// Fill with numeric data
		var values = new object[,]
		{
			{ 10, 20, 30 },
			{ 15, 25, 35 },
			{ 5, 15, 25 },
		};

		sheet.Range[sheet.Cells[1, 1], sheet.Cells[3, 3]].Value = values;

		// Add conditional formatting to the second row only (first row has none)
		var secondRow = sheet.Range[sheet.Cells[2, 1], sheet.Cells[2, 3]];
		var formatCondition = secondRow.FormatConditions.Add(
			Type: XlFormatConditionType.xlCellValue,
			Operator: XlFormatConditionOperator.xlLess,
			Formula1: "20");
		
		formatCondition.Interior.Color = ColorTranslator.ToOle(Color.LightBlue);

		Thread.Sleep(defaultSleep);

		// Verify initial state
		var firstRow = sheet.Range[sheet.Cells[1, 1], sheet.Cells[1, 3]];
		Assert.AreEqual(0, firstRow.FormatConditions.Count, "First row should not have conditional formatting initially");
		Assert.IsTrue(secondRow.FormatConditions.Count > 0, "Second row should have conditional formatting");

		// Select the entire range
		sheet.Range[sheet.Cells[1, 1], sheet.Cells[3, 3]].Select();
		Thread.Sleep(defaultSleep);

		// Alt + ZX + CF (RepairConditionalFormat keytip)
		SendKeys.SendWait("%"); // Alt
		SendKeys.SendWait("ZXCF");
		Thread.Sleep(defaultSleep);

		// Verify conditional formatting was copied to all rows
		for (int row = 1; row <= 3; row++)
		{
			var currentRow = sheet.Range[sheet.Cells[row, 1], sheet.Cells[row, 3]];
			Assert.IsTrue(currentRow.FormatConditions.Count > 0, $"Row {row} should have conditional formatting");
			TestContext.WriteLine($"Row {row} has {currentRow.FormatConditions.Count} conditional format(s)");
		}
	}
}
