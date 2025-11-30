using Microsoft.Office.Interop.Excel;

namespace NavfertyExcelAddIn.UnitTests.WorksheetCellsEditing;

[Category("Automation")]
[NotInParallel("Automation")]
public class ConditionalFormatAutomationTests : AutomationTestsBase
{
	[Test]
	//[Description("Test RepairConditionalFormat feature - copy conditional format from first row to entire range")]
	public async Task RepairConditionalFormat_WithConditionalFormatInFirstRow_CopiedToAllRows()
	{
		TestContext.Output.WriteLine("Testing RepairConditionalFormat feature");

		var workbook = app.Workbooks.Add();
		Thread.Sleep(defaultSleep);

		var sheet = (Worksheet)workbook.Sheets[1];
		Assert.NotNull(sheet);

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
		Range firstRow = sheet.Range[sheet.Cells[1, 1], sheet.Cells[1, 3]];
		var formatCondition = firstRow.FormatConditions.Add(
			Type: XlFormatConditionType.xlCellValue,
			Operator: XlFormatConditionOperator.xlGreater,
			Formula1: "15");
		
		formatCondition.Interior.Color = ColorTranslator.ToOle(Color.Yellow);

		Thread.Sleep(defaultSleep);

		// Verify initial state - only first row has conditional formatting
		await Assert.That(GetFormatConditionsCount(firstRow)).IsGreaterThan(0); //, "First row should have conditional formatting");
		
		var secondRow = sheet.Range[sheet.Cells[2, 1], sheet.Cells[2, 3]];
		var formatConditionsCount = (int)secondRow.FormatConditions.Count;
		await Assert.That(formatConditionsCount).IsEqualTo(0); //, "Second row should not have conditional formatting initially");

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
			Range currentRow = sheet.Range[sheet.Cells[row, 1], sheet.Cells[row, 3]];
			await Assert.That(GetFormatConditionsCount(currentRow)).IsGreaterThan(0); //, $"Row {row} should have conditional formatting");
			TestContext.Output.WriteLine($"Row {row} has {currentRow.FormatConditions.Count} conditional format(s)");
		}
	}

	[Test]
	[Skip("Only first row repair for conditional formattin is implemented")]
	//[Description("Test RepairConditionalFormat with conditional format in second row (first row empty)")]
	public async Task RepairConditionalFormat_WithConditionalFormatInSecondRow_CopiedToAllRows()
	{
		TestContext.Output.WriteLine("Testing RepairConditionalFormat with format in second row");

		var workbook = app.Workbooks.Add();
		Thread.Sleep(defaultSleep);

		var sheet = (Worksheet)workbook.Sheets[1];
		Assert.NotNull(sheet);

		// Fill with numeric data
		var values = new object[,]
		{
			{ 10, 20, 30 },
			{ 15, 25, 35 },
			{ 5, 15, 25 },
		};

		sheet.Range[sheet.Cells[1, 1], sheet.Cells[3, 3]].Value = values;

		// Add conditional formatting to the second row only (first row has none)
		Range secondRow = sheet.Range[sheet.Cells[2, 1], sheet.Cells[2, 3]];
		var formatCondition = secondRow.FormatConditions.Add(
			Type: XlFormatConditionType.xlCellValue,
			Operator: XlFormatConditionOperator.xlLess,
			Formula1: "20");
		
		formatCondition.Interior.Color = ColorTranslator.ToOle(Color.LightBlue);

		Thread.Sleep(defaultSleep);

		// Verify initial state
		Range firstRow = sheet.Range[sheet.Cells[1, 1], sheet.Cells[1, 3]];
		await Assert.That(GetFormatConditionsCount(firstRow)).EqualTo(0); //, "First row should not have conditional formatting initially");
		await Assert.That(GetFormatConditionsCount(secondRow)).IsGreaterThan(0); //, "Second row should have conditional formatting");

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
			Range currentRow = sheet.Range[sheet.Cells[row, 1], sheet.Cells[row, 3]];
			await Assert.That(GetFormatConditionsCount(currentRow)).IsGreaterThan(0); // > 0, $"Row {row} should have conditional formatting");
			TestContext.Output.WriteLine($"Row {row} has {currentRow.FormatConditions.Count} conditional format(s)");
		}
	}

	private static int GetFormatConditionsCount(Range range)
	{
		return (int)range.FormatConditions.Count;
	}
}
