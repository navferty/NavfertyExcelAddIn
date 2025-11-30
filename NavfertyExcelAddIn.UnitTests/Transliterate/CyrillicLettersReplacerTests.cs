using NavfertyExcelAddIn.Transliterate;

namespace NavfertyExcelAddIn.UnitTests.Transliterate;

public class CyrillicLettersReplacerTests
{
	[Test]
	[Arguments("\u0410123\u0412\u0415", "A123BE")]
	[Arguments("ABCDEFG\u0410\u0411\u0412\u0413\u0414\u0415", "ABCDEFGAБBГДE")]
	[Arguments("абвгдеёжзийклмнопрстуфхцчшщъыьэюя", "aбbгдeёжзийkлmhoпpctyфxцчшщъыьэюя")]
	[Arguments("АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯ", "AБBГДEЁЖЗИЙKЛMHOПPCTYФXЦЧШЩЪЫЬЭЮЯ")]
	public async Task Transliterate(string input, string expected)
	{
		var replacer = new CyrillicLettersReplacer();

		var output = replacer.ReplaceCyrillicCharsWithLatin(input);

		await Assert.That(output).IsEqualTo(expected);
	}
}
