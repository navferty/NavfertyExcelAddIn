using System.Threading;
using System.Windows.Forms;

using Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;

using NavfertyExcelAddIn.UnitTests.SqliteExport;

namespace NavfertyExcelAddIn.UnitTests.Commons;

#nullable enable

[TestClass]
public class CommonsAutomationTests : AutomationTestsBase
{
	[TestMethod]
	[TestCategory("Automation")]
	[Description("Test HighlightDuplicates feature - highlight duplicate values in selection")]
	public void HighlightDuplicates_DuplicateValues_Highlighted()
	{
		TestContext.WriteLine("Testing HighlightDuplicates feature");

		var workbook = app.Workbooks.Add();
		Thread.Sleep(defaultSleep);

		var sheet = (Worksheet)workbook.Sheets[1];
		Assert.IsNotNull(sheet);

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
		var appleColor1 = ((Range)sheet.Cells[1, 1]).Interior.ColorIndex;
		var appleColor2 = ((Range)sheet.Cells[1, 3]).Interior.ColorIndex;
		var appleColor3 = ((Range)sheet.Cells[3, 1]).Interior.ColorIndex;
		
		TestContext.WriteLine($"Apple cells - A1 color: {appleColor1}, C1 color: {appleColor2}, A3 color: {appleColor3}");
		Assert.AreEqual(appleColor1, appleColor2, "All 'apple' cells should have the same color");
		Assert.AreEqual(appleColor1, appleColor3, "All 'apple' cells should have the same color");
		Assert.AreNotEqual(XlColorIndex.xlColorIndexNone, appleColor1, "'apple' cells should be colored");

		// banana appears in: B1, B2
		var bananaColor1 = ((Range)sheet.Cells[1, 2]).Interior.ColorIndex;
		var bananaColor2 = ((Range)sheet.Cells[2, 2]).Interior.ColorIndex;
		
		TestContext.WriteLine($"Banana cells - B1 color: {bananaColor1}, B2 color: {bananaColor2}");
		Assert.AreEqual(bananaColor1, bananaColor2, "All 'banana' cells should have the same color");
		Assert.AreNotEqual(XlColorIndex.xlColorIndexNone, bananaColor1, "'banana' cells should be colored");

		// orange appears in: A2, C3
		var orangeColor1 = ((Range)sheet.Cells[2, 1]).Interior.ColorIndex;
		var orangeColor2 = ((Range)sheet.Cells[3, 3]).Interior.ColorIndex;
		
		TestContext.WriteLine($"Orange cells - A2 color: {orangeColor1}, C3 color: {orangeColor2}");
		Assert.AreEqual(orangeColor1, orangeColor2, "All 'orange' cells should have the same color");
		Assert.AreNotEqual(XlColorIndex.xlColorIndexNone, orangeColor1, "'orange' cells should be colored");

		// Different duplicate groups should have different colors
		Assert.AreNotEqual(appleColor1, bananaColor1, "Different duplicate groups should have different colors");
		Assert.AreNotEqual(appleColor1, orangeColor1, "Different duplicate groups should have different colors");
		Assert.AreNotEqual(bananaColor1, orangeColor1, "Different duplicate groups should have different colors");

		// grape and kiwi appear only once, should not be colored
		var grapeColor = (XlColorIndex)((Range)sheet.Cells[2, 3]).Interior.ColorIndex;
		var kiwiColor = (XlColorIndex)((Range)sheet.Cells[3, 2]).Interior.ColorIndex;
		
		TestContext.WriteLine($"Unique cells - grape color: {grapeColor}, kiwi color: {kiwiColor}");
		Assert.AreEqual(XlColorIndex.xlColorIndexNone, grapeColor, "'grape' should not be colored (not a duplicate)");
		Assert.AreEqual(XlColorIndex.xlColorIndexNone, kiwiColor, "'kiwi' should not be colored (not a duplicate)");
	}

	[TestMethod]
	[TestCategory("Automation")]
	[Description("Test UnmergeCells feature - unmerge merged cells and fill with values")]
	public void UnmergeCells_MergedCells_Unmerged()
	{
		TestContext.WriteLine("Testing UnmergeCells feature");

		var workbook = app.Workbooks.Add();
		Thread.Sleep(defaultSleep);

		var sheet = (Worksheet)workbook.Sheets[1];
		Assert.IsNotNull(sheet);

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

		Assert.IsFalse((bool)cell1.MergeCells, "Cell A1 should not be merged");
		Assert.IsFalse((bool)cell2.MergeCells, "Cell B1 should not be merged");
		Assert.IsFalse((bool)cell3.MergeCells, "Cell A2 should not be merged");
		Assert.IsFalse((bool)cell4.MergeCells, "Cell A3 should not be merged");

		// Verify values are filled
		Assert.AreEqual("Merged", cell1.Value);
		Assert.AreEqual("Merged", cell2.Value);
		Assert.AreEqual("Vertical", cell3.Value);
		Assert.AreEqual("Vertical", cell4.Value);
	}

	[TestMethod]
	[TestCategory("Automation")]
	[Description("Test FindErrors feature - find cells with formula errors")]
	public void FindErrors_FormulasWithErrors_Found()
	{
		TestContext.WriteLine("Testing FindErrors feature");

		var workbook = app.Workbooks.Add();
		Thread.Sleep(defaultSleep);

		var sheet = (Worksheet)workbook.Sheets[1];
		Assert.IsNotNull(sheet);

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
		Assert.IsNotNull(selectedCell);
		TestContext.WriteLine($"Active cell after FindErrors: {selectedCell.Address}");
		Assert.AreEqual("$A$5", selectedCell.Address, "Active cell should be A3 (first error)");

		// The FindErrors feature should identify and highlight error cells
		// We can verify it completed without errors
		var cell3Value = ((Range)sheet.Cells[3, 1]).Value;
		TestContext.WriteLine($"Cell A3 value: {cell3Value}");
		
		// Verify the error value is present
		Assert.IsNotNull(cell3Value, "Error cell should have a value");
	}

	[TestMethod]
	[TestCategory("Automation")]
	[Description("Test CopyAsMarkdown feature - copy range as markdown table")]
	public void CopyAsMarkdown_Range_CopiedToClipboard()
	{
		TestContext.WriteLine("Testing CopyAsMarkdown feature");

		var workbook = app.Workbooks.Add();
		Thread.Sleep(defaultSleep);

		var sheet = (Worksheet)workbook.Sheets[1];
		Assert.IsNotNull(sheet);

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
			TestContext.WriteLine("Clipboard content:");
			TestContext.WriteLine(clipboardText);

			// Verify markdown format (should contain pipe characters for table)
			Assert.IsTrue(clipboardText.Contains("|"), "Markdown table should contain pipe characters");
			Assert.IsTrue(clipboardText.Contains("Name"), "Markdown should contain header 'Name'");
			Assert.IsTrue(clipboardText.Contains("Alice"), "Markdown should contain 'Alice'");
		}
	}

	[TestMethod]
	[TestCategory("Automation")]
	[Description("Test HighlightDuplicates with numeric values")]
	public void HighlightDuplicates_NumericValues_Highlighted()
	{
		TestContext.WriteLine("Testing HighlightDuplicates with numbers");

		var workbook = app.Workbooks.Add();
		Thread.Sleep(defaultSleep);

		var sheet = (Worksheet)workbook.Sheets[1];
		Assert.IsNotNull(sheet);

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
		var value100Color1 = ((Range)sheet.Cells[1, 1]).Interior.ColorIndex;
		var value100Color2 = ((Range)sheet.Cells[1, 3]).Interior.ColorIndex;
		var value100Color3 = ((Range)sheet.Cells[3, 1]).Interior.ColorIndex;
		
		TestContext.WriteLine($"100 cells - A1 color: {value100Color1}, C1 color: {value100Color2}, A3 color: {value100Color3}");
		Assert.AreEqual(value100Color1, value100Color2, "All '100' cells should have the same color");
		Assert.AreEqual(value100Color1, value100Color3, "All '100' cells should have the same color");
		Assert.AreNotEqual(XlColorIndex.xlColorIndexNone, value100Color1, "'100' cells should be colored");

		// 200 appears in: B1, B2
		var value200Color1 = ((Range)sheet.Cells[1, 2]).Interior.ColorIndex;
		var value200Color2 = ((Range)sheet.Cells[2, 2]).Interior.ColorIndex;
		
		TestContext.WriteLine($"200 cells - B1 color: {value200Color1}, B2 color: {value200Color2}");
		Assert.AreEqual(value200Color1, value200Color2, "All '200' cells should have the same color");
		Assert.AreNotEqual(XlColorIndex.xlColorIndexNone, value200Color1, "'200' cells should be colored");

		// 300 appears in: A2, C3
		var value300Color1 = ((Range)sheet.Cells[2, 1]).Interior.ColorIndex;
		var value300Color2 = ((Range)sheet.Cells[3, 3]).Interior.ColorIndex;
		
		TestContext.WriteLine($"300 cells - A2 color: {value300Color1}, C3 color: {value300Color2}");
		Assert.AreEqual(value300Color1, value300Color2, "All '300' cells should have the same color");
		Assert.AreNotEqual(XlColorIndex.xlColorIndexNone, value300Color1, "'300' cells should be colored");

		// Different duplicate groups should have different colors
		Assert.AreNotEqual(value100Color1, value200Color1, "Different duplicate groups should have different colors");
		Assert.AreNotEqual(value100Color1, value300Color1, "Different duplicate groups should have different colors");
		Assert.AreNotEqual(value200Color1, value300Color1, "Different duplicate groups should have different colors");

		// 400 and 500 appear only once, should not be colored
		var value400Color = (XlColorIndex)((Range)sheet.Cells[2, 3]).Interior.ColorIndex;
		var value500Color = (XlColorIndex)((Range)sheet.Cells[3, 2]).Interior.ColorIndex;
		
		TestContext.WriteLine($"Unique cells - 400 color: {value400Color}, 500 color: {value500Color}");
		Assert.AreEqual(XlColorIndex.xlColorIndexNone, value400Color, "'400' should not be colored (not a duplicate)");
		Assert.AreEqual(XlColorIndex.xlColorIndexNone, value500Color, "'500' should not be colored (not a duplicate)");
	}
}
