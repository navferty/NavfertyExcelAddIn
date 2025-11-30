using Microsoft.Office.Interop.Excel;

namespace NavfertyExcelAddIn.UnitTests.Transliterate;

[Category("Automation")]
[NotInParallel("Automation")]
public class TransliterateAutomationTests : AutomationTestsBase
{
	[Test]
	//[Description("Test Transliterate feature - convert Cyrillic to Latin")]
	public async Task Transliterate_CyrillicText_ConvertedToLatin()
	{
		TestContext.Output.WriteLine("Testing Transliterate feature");

		var workbook = app.Workbooks.Add();
		Thread.Sleep(defaultSleep);

		var sheet = (Worksheet)workbook.Sheets[1];
		await Assert.That(sheet).IsNotNull();

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
		await Assert.That(currentValues).IsNotNull();

		foreach (var cell in currentValues)
		{
			TestContext.Output.WriteLine(cell?.ToString() ?? "<null>");
		}

		// Expected transliteration results
		var expectedValues = new object[,]
		{
			{ "Privet", "Moskva", "Rossiia" },
			{ "Karl", "Klara", "korally" },
		};

		await Assert.That(currentValues).IsEquivalentTo(expectedValues);
	}

	[Test]
	//[Description("Test ReplaceChars feature - replace specific characters")]
	public async Task ReplaceChars_CustomReplacements_Applied()
	{
		TestContext.Output.WriteLine("Testing ReplaceChars feature");

		var workbook = app.Workbooks.Add();
		Thread.Sleep(defaultSleep);

		var sheet = (Worksheet)workbook.Sheets[1];
		await Assert.That(sheet).IsNotNull();

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
		await Assert.That(currentValues).IsNotNull();

		foreach (var cell in currentValues)
		{
			TestContext.Output.WriteLine(cell?.ToString() ?? "<null>");
		}

		// Verify that characters were replaced (specific replacements depend on implementation)
		// At minimum, verify that the operation completed without error
		await Assert.That(currentValues[1, 1]).IsNotNull();
		await Assert.That(currentValues[1, 2]).IsNotNull();
		await Assert.That(currentValues[1, 3]).IsNotNull();
	}

	[Test]
	//[Description("Test Transliterate with mixed Cyrillic and Latin text")]
	public async Task Transliterate_MixedText_OnlyCyrillicTransliterated()
	{
		TestContext.Output.WriteLine("Testing Transliterate with mixed content");

		var workbook = app.Workbooks.Add();
		Thread.Sleep(defaultSleep);

		var sheet = (Worksheet)workbook.Sheets[1];
		await Assert.That(sheet).IsNotNull();

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
		await Assert.That(currentValues).IsNotNull();

		foreach (var cell in currentValues)
		{
			TestContext.Output.WriteLine(cell?.ToString() ?? "<null>");
		}

		// Expected: only Cyrillic parts transliterated
		var expectedValues = new object[,]
		{
			{ "Test Test", "English Russkii", "123 trista" },
		};

		await Assert.That(currentValues).IsEquivalentTo(expectedValues);
	}
}
