using System.Collections.Generic;
using System.Threading;
using System.Windows.Forms;

using Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;

using NavfertyExcelAddIn.UnitTests.SqliteExport;

namespace NavfertyExcelAddIn.UnitTests.StringifyNumerics;

#nullable enable

[TestClass]
public class StringifyNumericsAutomationTests : AutomationTestsBase
{
	[TestMethod]
	[DynamicData(nameof(GetStringifyTestCases))]
	[TestCategory("Automation")]
	[Description("Create blank workbook, fill it with numbers and stringify numerics to each of 3 language options")]
	public void StringifyNumerics(string language, string keySequence, object[,] expectedValues)
	{
		TestContext.WriteLine($"Testing {language} stringification");

		var workbook = app.Workbooks.Add();
		Thread.Sleep(defaultSleep);

		var sheet = (Worksheet)workbook.Sheets[1];
		Assert.IsNotNull(sheet);

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
			TestContext.WriteLine(cell?.ToString() ?? "<null>");
		}

		Assert.IsNotNull(currentValues);
		CollectionAssert.AreEqual(expectedValues, currentValues);
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
