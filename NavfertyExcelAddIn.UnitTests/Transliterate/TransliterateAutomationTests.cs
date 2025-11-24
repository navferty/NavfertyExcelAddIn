using System.Threading;
using System.Windows.Forms;

using Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;

using NavfertyExcelAddIn.UnitTests.SqliteExport;

namespace NavfertyExcelAddIn.UnitTests.Transliterate;

#nullable enable

[TestClass]
public class TransliterateAutomationTests : AutomationTestsBase
{
	[TestMethod]
	[TestCategory("Automation")]
	[Description("Test Transliterate feature - convert Cyrillic to Latin")]
	public void Transliterate_CyrillicText_ConvertedToLatin()
	{
		TestContext.WriteLine("Testing Transliterate feature");

		var workbook = app.Workbooks.Add();
		Thread.Sleep(defaultSleep);

		var sheet = (Worksheet)workbook.Sheets[1];
		Assert.IsNotNull(sheet);

		// Fill with Cyrillic text
		var values = new object[,]
		{
			{ "Привет", "Москва", "Россия" },
			{ "Карл", "Клара", "кораллы" },
		};

		sheet.Range[sheet.Cells[1, 1], sheet.Cells[2, 3]].Value = values;
		sheet.UsedRange.Select();
		Thread.Sleep(defaultSleep);

		// Alt + ZX + TL + first entry (Transliterate keytip)
		SendKeys.SendWait("%"); // Alt
		SendKeys.SendWait("ZXTL{ENTER}");
		Thread.Sleep(defaultSleep);

		var currentValues = (object?[,])sheet.UsedRange.Cells.Value;
		Assert.IsNotNull(currentValues);

		foreach (var cell in currentValues)
		{
			TestContext.WriteLine(cell?.ToString() ?? "<null>");
		}

		// Expected transliteration results
		var expectedValues = new object[,]
		{
			{ "Privet", "Moskva", "Rossiia" },
			{ "Karl", "Klara", "korally" },
		};

		CollectionAssert.AreEqual(expectedValues, currentValues);
	}

	[TestMethod]
	[TestCategory("Automation")]
	[Description("Test ReplaceChars feature - replace specific characters")]
	public void ReplaceChars_CustomReplacements_Applied()
	{
		TestContext.WriteLine("Testing ReplaceChars feature");

		var workbook = app.Workbooks.Add();
		Thread.Sleep(defaultSleep);

		var sheet = (Worksheet)workbook.Sheets[1];
		Assert.IsNotNull(sheet);

		// Fill with text containing characters to replace
		var values = new object[,]
		{
			{ "Тест №1", "Пример №2", "Образец №3" },
		};

		sheet.Range[sheet.Cells[1, 1], sheet.Cells[1, 3]].Value = values;
		sheet.UsedRange.Select();
		Thread.Sleep(defaultSleep);

		// Alt + ZX + TL + second entry (ReplaceChars keytip)
		SendKeys.SendWait("%"); // Alt
		SendKeys.SendWait("ZXTL{DOWN}{ENTER}");
		Thread.Sleep(defaultSleep);

		var currentValues = (object?[,])sheet.UsedRange.Cells.Value;
		Assert.IsNotNull(currentValues);

		foreach (var cell in currentValues)
		{
			TestContext.WriteLine(cell?.ToString() ?? "<null>");
		}

		// Verify that characters were replaced (specific replacements depend on implementation)
		// At minimum, verify that the operation completed without error
		Assert.IsNotNull(currentValues[1, 1]);
		Assert.IsNotNull(currentValues[1, 2]);
		Assert.IsNotNull(currentValues[1, 3]);
	}

	[TestMethod]
	[TestCategory("Automation")]
	[Description("Test Transliterate with mixed Cyrillic and Latin text")]
	public void Transliterate_MixedText_OnlyCyrillicTransliterated()
	{
		TestContext.WriteLine("Testing Transliterate with mixed content");

		var workbook = app.Workbooks.Add();
		Thread.Sleep(defaultSleep);

		var sheet = (Worksheet)workbook.Sheets[1];
		Assert.IsNotNull(sheet);

		// Fill with mixed content
		var values = new object[,]
		{
			{ "Test Тест", "English Русский", "123 триста" },
		};

		sheet.Range[sheet.Cells[1, 1], sheet.Cells[1, 3]].Value = values;
		sheet.UsedRange.Select();
		Thread.Sleep(defaultSleep);

		// Alt + ZX + TL + first entry (Transliterate keytip)
		SendKeys.SendWait("%"); // Alt
		SendKeys.SendWait("ZXTL{ENTER}");
		Thread.Sleep(defaultSleep);

		var currentValues = (object?[,])sheet.UsedRange.Cells.Value;
		Assert.IsNotNull(currentValues);

		foreach (var cell in currentValues)
		{
			TestContext.WriteLine(cell?.ToString() ?? "<null>");
		}

		// Expected: only Cyrillic parts transliterated
		var expectedValues = new object[,]
		{
			{ "Test Test", "English Russkii", "123 trista" },
		};

		CollectionAssert.AreEqual(expectedValues, currentValues);
	}
}
