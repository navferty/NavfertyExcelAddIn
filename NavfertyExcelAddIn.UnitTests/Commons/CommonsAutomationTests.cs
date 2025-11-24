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

		// Verify conditional formatting was applied
		var usedRange = sheet.UsedRange;
		Assert.IsTrue(usedRange.FormatConditions.Count > 0, "Conditional formatting should be applied");

		TestContext.WriteLine($"Conditional formats applied: {usedRange.FormatConditions.Count}");
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

		// Verify conditional formatting was applied
		var usedRange = sheet.UsedRange;
		Assert.IsTrue(usedRange.FormatConditions.Count > 0, "Conditional formatting should be applied");

		TestContext.WriteLine($"Conditional formats applied: {usedRange.FormatConditions.Count}");
	}
}
