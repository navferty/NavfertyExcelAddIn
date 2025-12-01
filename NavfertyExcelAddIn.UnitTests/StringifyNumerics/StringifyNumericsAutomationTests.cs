using Microsoft.Office.Interop.Excel;

namespace NavfertyExcelAddIn.UnitTests.StringifyNumerics;

[Category("Automation")]
[NotInParallel("Automation")]
public class StringifyNumericsAutomationTests : AutomationTestsBase
{
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
