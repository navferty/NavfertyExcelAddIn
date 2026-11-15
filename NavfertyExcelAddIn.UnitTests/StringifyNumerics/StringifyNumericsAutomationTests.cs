using Autofac;

using Microsoft.Office.Interop.Excel;

using NavfertyExcelAddIn.Commons;
using NavfertyExcelAddIn.Undo;

namespace NavfertyExcelAddIn.UnitTests.StringifyNumerics;

[Category("Automation")]
[NotInParallel("Automation")]
public class StringifyNumericsAutomationTests : AutomationTestsBase
{
	[Test]
	public async Task Undo_RestoresFormulasAndValues()
	{
		var workbook = app.Workbooks.Add();
		var sheet = (Worksheet)workbook.Sheets[1];
		((Range)sheet.Cells[1, 1]).Value = 1;
		((Range)sheet.Cells[2, 1]).Value = 2;
		((Range)sheet.Cells[3, 1]).Formula = "=A1+A2";
		((Range)sheet.Cells[1, 2]).Value = "text";
		var selection = sheet.Range["A1", "B3"];
		var undoManager = NavfertyRibbon.Container.Resolve<UndoManager>();

		selection.ApplyForEachCellOfType<double, string>(x => $"number {x}");
		await Assert.That((string)((Range)sheet.Cells[3, 1]).Value).IsEqualTo("number 3");
		undoManager.UndoLastAction(sheet);

		await Assert.That((string)((Range)sheet.Cells[3, 1]).Formula).IsEqualTo("=A1+A2");
		await Assert.That((double)((Range)sheet.Cells[1, 1]).Value).IsEqualTo(1d);
		await Assert.That((double)((Range)sheet.Cells[2, 1]).Value).IsEqualTo(2d);
		await Assert.That((string)((Range)sheet.Cells[1, 2]).Value).IsEqualTo("text");
		await Assert.That((object?)((Range)sheet.Cells[2, 2]).Value).IsNull();

		var formulaCell = (Range)sheet.Cells[3, 1];
		formulaCell.ApplyForEachCellOfType<double, string>(x => $"number {x}");
		undoManager.UndoLastAction(sheet);

		await Assert.That((string)formulaCell.Formula).IsEqualTo("=A1+A2");

		selection.ApplyForEachCellOfType<double, string>(x => $"number {x}");
		((Range)sheet.Cells[1, 1]).Value = "edited";
		undoManager.UndoLastAction(sheet);

		await Assert.That((string)((Range)sheet.Cells[1, 1]).Value).IsEqualTo("edited");
		await Assert.That((string)formulaCell.Value).IsEqualTo("number 3");
	}

	[Test]
	[MethodDataSource(nameof(GetStringifyTestCases))]
	[Property("Description", "Create blank workbook, fill it with numbers and stringify numerics to each of 3 language options")]
	public async Task StringifyNumerics(string language, string keySequence, object[,] expectedValues)
	{
		TestContext.Output.WriteLine($"Testing {language} stringification");

		var workbook = app.Workbooks.Add();
		Thread.Sleep(defaultSleep);

		var sheet = (Worksheet)workbook.Sheets[1];
		await Assert.That(sheet).IsNotNull();

		var values = new[,]
		{
			{ 1, 2, 3 },
			{ 10, 11, 12 },
			{ 100, 200, 300 },
		};

		sheet.Range[sheet.Cells[1, 1], sheet.Cells[3, 3]].Value = values;
		sheet.UsedRange.Select();
		Thread.Sleep(defaultSleep);

		SendKeys.SendWait("%"); // Alt
		SendKeys.SendWait("ZXSN");

		// arrow down to select language
		SendKeys.SendWait(keySequence);

		// enter
		SendKeys.SendWait("{ENTER}");
		Thread.Sleep(defaultSleep);

		var currentValues = (object?[,])sheet.UsedRange.Cells.Value;
		foreach (var cell in currentValues)
		{
			TestContext.Output.WriteLine(cell?.ToString() ?? "<null>");
		}

		await Assert.That(currentValues).IsNotNull();

		// TODO
		await Assert.That(currentValues).IsEquivalentTo(expectedValues);
	}

	public static IEnumerable<object[]> GetStringifyTestCases()
	{
		// Russian
		yield return new object[]
		{
			"Russian",
			"", // No arrow key needed (first option)
			new object[,]
			{
				{ "один", "два", "три" },
				{ "десять", "одиннадцать", "двенадцать" },
				{ "сто", "двести", "триста" },
			}
		};

		// English
		yield return new object[]
		{
			"English",
			"{DOWN}", // Arrow down to select English
			new object[,]
			{
				{ "one", "two", "three" },
				{ "ten", "eleven", "twelve" },
				{ "one hundred", "two hundred", "three hundred" },
			}
		};

		// French
		yield return new object[]
		{
			"French",
			"{DOWN}{DOWN}", // Arrow down twice to select French
			new object[,]
			{
				{ "un", "deux", "trois" },
				{ "dix", "onze", "douze" },
				{ "cent", "deux cent", "trois cent" },
			}
		};
	}
}
