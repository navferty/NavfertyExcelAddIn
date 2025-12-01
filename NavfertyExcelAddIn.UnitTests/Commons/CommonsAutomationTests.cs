using Microsoft.Office.Interop.Excel;

namespace NavfertyExcelAddIn.UnitTests.Commons;

[Category("Automation")]
[NotInParallel("Automation")]
public class CommonsAutomationTests : AutomationTestsBase
{
	[Test]
	[DisplayName("Test HighlightDuplicates feature - highlight duplicate values in selection")]
	public async Task HighlightDuplicates_DuplicateValues_Highlighted()
	{
		TestContext.Output.WriteLine("Testing HighlightDuplicates feature");

		var workbook = app.Workbooks.Add();
		Thread.Sleep(defaultSleep);

		var sheet = (Worksheet)workbook.Sheets[1];
		await Assert.That(sheet).IsNotNull();

		// Fill with data containing duplicates
		var values = new object[,]
		{
			{ "apple", "banana", "apple" },
			{ "orange", "banana", "grape" },
			{ "apple", "kiwi", "orange" },
		};

		sheet.Range[sheet.Cells[1, 1], sheet.Cells[3, 3]].Value = values;
		sheet.UsedRange.Select();
		Thread.Sleep(defaultSleep);

		// Alt + ZX + D (HighlightDuplicates keytip)
		SendKeys.SendWait("%"); // Alt
		SendKeys.SendWait("ZXD");
		Thread.Sleep(defaultSleep);

		// Verify that cells with same values have same color
		// apple appears in: A1, C1, A3
		var appleColor1 = (XlColorIndex)((Range)sheet.Cells[1, 1]).Interior.ColorIndex;
		var appleColor2 = (XlColorIndex)((Range)sheet.Cells[1, 3]).Interior.ColorIndex;
		var appleColor3 = (XlColorIndex)((Range)sheet.Cells[3, 1]).Interior.ColorIndex;
		
		TestContext.Output.WriteLine($"Apple cells - A1 color: {appleColor1}, C1 color: {appleColor2}, A3 color: {appleColor3}");
		await Assert.That(appleColor2).IsEqualTo(appleColor1).And.IsEqualTo(appleColor3);
		await Assert.That(appleColor1).IsNotEqualTo(XlColorIndex.xlColorIndexNone);

		// banana appears in: B1, B2
		var bananaColor1 = (XlColorIndex)((Range)sheet.Cells[1, 2]).Interior.ColorIndex;
		var bananaColor2 = (XlColorIndex)((Range)sheet.Cells[2, 2]).Interior.ColorIndex;
		
		TestContext.Output.WriteLine($"Banana cells - B1 color: {bananaColor1}, B2 color: {bananaColor2}");
		await Assert.That(bananaColor2).IsEqualTo(bananaColor1);
		await Assert.That(bananaColor1).IsNotEqualTo(XlColorIndex.xlColorIndexNone);

		// orange appears in: A2, C3
		var orangeColor1 = (XlColorIndex)((Range)sheet.Cells[2, 1]).Interior.ColorIndex;
		var orangeColor2 = (XlColorIndex)((Range)sheet.Cells[3, 3]).Interior.ColorIndex;
		
		TestContext.Output.WriteLine($"Orange cells - A2 color: {orangeColor1}, C3 color: {orangeColor2}");
		await Assert.That(orangeColor2).IsEqualTo(orangeColor1);
		await Assert.That(orangeColor1).IsNotEqualTo(XlColorIndex.xlColorIndexNone);

		// Different duplicate groups should have different colors
		await Assert.That(appleColor1).IsNotEqualTo(bananaColor1);
		await Assert.That(appleColor1).IsNotEqualTo(orangeColor1);
		await Assert.That(bananaColor1).IsNotEqualTo(orangeColor1);

		// grape and kiwi appear only once, should not be colored
		var grapeColor = (XlColorIndex)((Range)sheet.Cells[2, 3]).Interior.ColorIndex;
		var kiwiColor = (XlColorIndex)((Range)sheet.Cells[3, 2]).Interior.ColorIndex;
		
		TestContext.Output.WriteLine($"Unique cells - grape color: {grapeColor}, kiwi color: {kiwiColor}");
		await Assert.That(grapeColor).IsEqualTo(XlColorIndex.xlColorIndexNone); //, , "'grape' should not be colored (not a duplicate)");
		await Assert.That(kiwiColor).IsEqualTo(XlColorIndex.xlColorIndexNone); //, , "'kiwi' should not be colored (not a duplicate)");
	}

	[Test]
	[Property("Description", "Test UnmergeCells feature - unmerge merged cells and fill with values")]
	public async Task UnmergeCells_MergedCells_Unmerged()
	{
		TestContext.Output.WriteLine("Testing UnmergeCells feature");

		var workbook = app.Workbooks.Add();
		Thread.Sleep(defaultSleep);

		var sheet = (Worksheet)workbook.Sheets[1];
		await Assert.That(sheet).IsNotNull();

		// Create merged cells
		var range1 = sheet.Range[sheet.Cells[1, 1], sheet.Cells[1, 3]];
		range1.Value = "Merged";
		range1.Merge();

		var range2 = sheet.Range[sheet.Cells[2, 1], sheet.Cells[3, 1]];
		range2.Value = "Vertical";
		range2.Merge();

		Thread.Sleep(defaultSleep);

		// Select the area with merged cells
		sheet.Range[sheet.Cells[1, 1], sheet.Cells[3, 3]].Select();
		Thread.Sleep(defaultSleep);

		// Alt + ZX + UM (UnmergeCells keytip)
		SendKeys.SendWait("%"); // Alt
		SendKeys.SendWait("ZXUM");
		Thread.Sleep(defaultSleep);

		// Verify cells are unmerged
		var cell1 = (Range)sheet.Cells[1, 1];
		var cell2 = (Range)sheet.Cells[1, 2];
		var cell3 = (Range)sheet.Cells[2, 1];
		var cell4 = (Range)sheet.Cells[3, 1];

		await Assert.That((bool)cell1.MergeCells).IsFalse();
		await Assert.That((bool)cell2.MergeCells).IsFalse();
		await Assert.That((bool)cell3.MergeCells).IsFalse();
		await Assert.That((bool)cell4.MergeCells).IsFalse();

		// Verify values are filled
		await Assert.That(cell1.Value as object).IsEqualTo("Merged");
		await Assert.That(cell2.Value as object).IsEqualTo("Merged");
		await Assert.That(cell3.Value as object).IsEqualTo("Vertical");
		await Assert.That(cell4.Value as object).IsEqualTo("Vertical");
	}

	[Test]
	[Property("Description", "Test FindErrors feature - find cells with formula errors")]
	public async Task FindErrors_FormulasWithErrors_Found()
	{
		TestContext.Output.WriteLine("Testing FindErrors feature");

		var workbook = app.Workbooks.Add();
		Thread.Sleep(defaultSleep);

		var sheet = (Worksheet)workbook.Sheets[1];
		await Assert.That(sheet).IsNotNull();

		// Create some formulas with errors
		sheet.Cells[1, 1] = 10;
		sheet.Cells[2, 1] = 0;
		((Range)sheet.Cells[3, 1]).Formula = "=A1/A2"; // Division by zero: #DIV/0!
		((Range)sheet.Cells[4, 1]).Formula = "=A1/B1"; // Reference to empty cell (0)
		((Range)sheet.Cells[5, 1]).Formula = "=VLOOKUP(1,A1:A2,2,FALSE)"; // #N/A error

		Thread.Sleep(defaultSleep);

		// Select the worksheet
		sheet.UsedRange.Select();
		Thread.Sleep(defaultSleep);

		// Alt + ZX + ER (FindErrors keytip)
		SendKeys.SendWait("%"); // Alt
		SendKeys.SendWait("ZXER");
		Thread.Sleep(defaultSleep);

		SendKeys.SendWait("{DOWN}{DOWN}"); // Move to third error in A5 cell

		var selectedCell = app.ActiveCell;
		await Assert.That(selectedCell).IsNotNull();
		TestContext.Output.WriteLine($"Active cell after FindErrors: {selectedCell.Address}");
		await Assert.That(selectedCell.Address).IsEqualTo("$A$5");

		// The FindErrors feature should identify and highlight error cells
		// We can verify it completed without errors
		var cell3Value = ((Range)sheet.Cells[3, 1]).Value;
		TestContext.Output.WriteLine($"Cell A3 value: {cell3Value}");
		
		// Verify the error value is present
		await Assert.That(cell3Value as object).IsNotNull();
	}

	[Test]
	[Property("Description", "Test CopyAsMarkdown feature - copy range as markdown table")]
	public async Task CopyAsMarkdown_Range_CopiedToClipboard()
	{
		TestContext.Output.WriteLine("Testing CopyAsMarkdown feature");

		var workbook = app.Workbooks.Add();
		Thread.Sleep(defaultSleep);

		var sheet = (Worksheet)workbook.Sheets[1];
		await Assert.That(sheet).IsNotNull();

		// Fill with data for markdown table
		var values = new object[,]
		{
			{ "Name", "Age", "City" },
			{ "Alice", 30, "New York" },
			{ "Bob", 25, "Los Angeles" },
		};

		sheet.Range[sheet.Cells[1, 1], sheet.Cells[3, 3]].Value = values;
		sheet.Range[sheet.Cells[1, 1], sheet.Cells[3, 3]].Select();
		Thread.Sleep(defaultSleep);

		// Alt + ZX + MD (CopyAsMarkdown keytip)
		SendKeys.SendWait("%"); // Alt
		SendKeys.SendWait("ZXMD");
		Thread.Sleep(defaultSleep);

		// Verify clipboard contains markdown
		if (Clipboard.ContainsText())
		{
			var clipboardText = Clipboard.GetText();
			TestContext.Output.WriteLine("Clipboard content:");
			TestContext.Output.WriteLine(clipboardText);

			// Verify markdown format (should contain pipe characters for table)
			await Assert.That(clipboardText).Contains("|");
			await Assert.That(clipboardText).Contains("Name");
			await Assert.That(clipboardText).Contains("Alice");
		}
	}

	[Test]
	[Property("Description", "Test HighlightDuplicates with numeric values")]
	public async Task HighlightDuplicates_NumericValues_Highlighted()
	{
		TestContext.Output.WriteLine("Testing HighlightDuplicates with numbers");

		var workbook = app.Workbooks.Add();
		Thread.Sleep(defaultSleep);

		var sheet = (Worksheet)workbook.Sheets[1];
		await Assert.That(sheet).IsNotNull();

		// Fill with numeric data containing duplicates
		var values = new object[,]
		{
			{ 100, 200, 100 },
			{ 300, 200, 400 },
			{ 100, 500, 300 },
		};

		sheet.Range[sheet.Cells[1, 1], sheet.Cells[3, 3]].Value = values;
		sheet.UsedRange.Select();
		Thread.Sleep(defaultSleep);

		// Alt + ZX + D (HighlightDuplicates keytip)
		SendKeys.SendWait("%"); // Alt
		SendKeys.SendWait("ZXD");
		Thread.Sleep(defaultSleep);

		// Verify that cells with same values have same color
		// 100 appears in: A1, C1, A3
		var value100Color1 = (XlColorIndex)((Range)sheet.Cells[1, 1]).Interior.ColorIndex;
		var value100Color2 = (XlColorIndex)((Range)sheet.Cells[1, 3]).Interior.ColorIndex;
		var value100Color3 = (XlColorIndex)((Range)sheet.Cells[3, 1]).Interior.ColorIndex;
		
		TestContext.Output.WriteLine($"100 cells - A1 color: {value100Color1}, C1 color: {value100Color2}, A3 color: {value100Color3}");
		await Assert.That(value100Color2).IsEqualTo(value100Color1).And.IsEqualTo(value100Color3);
		await Assert.That(value100Color1).IsNotEqualTo(XlColorIndex.xlColorIndexNone);

		// 200 appears in: B1, B2
		var value200Color1 = (XlColorIndex)((Range)sheet.Cells[1, 2]).Interior.ColorIndex;
		var value200Color2 = (XlColorIndex)((Range)sheet.Cells[2, 2]).Interior.ColorIndex;
		
		TestContext.Output.WriteLine($"200 cells - B1 color: {value200Color1}, B2 color: {value200Color2}");
		await Assert.That(value200Color2).IsEqualTo(value200Color1);
		await Assert.That(value200Color1).IsNotEqualTo(XlColorIndex.xlColorIndexNone);

		// 300 appears in: A2, C3
		var value300Color1 = (XlColorIndex)((Range)sheet.Cells[2, 1]).Interior.ColorIndex;
		var value300Color2 = (XlColorIndex)((Range)sheet.Cells[3, 3]).Interior.ColorIndex;
		
		TestContext.Output.WriteLine($"300 cells - A2 color: {value300Color1}, C3 color: {value300Color2}");
		await Assert.That(value300Color2).IsEqualTo(value300Color1);
		await Assert.That(value300Color1).IsNotEqualTo(XlColorIndex.xlColorIndexNone);

		// Different duplicate groups should have different colors
		await Assert.That(value100Color1).IsNotEqualTo(value200Color1);
		await Assert.That(value100Color1).IsNotEqualTo(value300Color1);
		await Assert.That(value200Color1).IsNotEqualTo(value300Color1);

		// 400 and 500 appear only once, should not be colored
		var value400Color = (XlColorIndex)((Range)sheet.Cells[2, 3]).Interior.ColorIndex;
		var value500Color = (XlColorIndex)((Range)sheet.Cells[3, 2]).Interior.ColorIndex;
		
		TestContext.Output.WriteLine($"Unique cells - 400 color: {value400Color}, 500 color: {value500Color}");
		await Assert.That(value400Color).IsEqualTo(XlColorIndex.xlColorIndexNone);
		await Assert.That(value500Color).IsEqualTo(XlColorIndex.xlColorIndexNone);
	}
}
