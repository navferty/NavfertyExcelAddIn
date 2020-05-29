using Microsoft.VisualStudio.TestTools.UnitTesting;

using NavfertyExcelAddIn.Transliterate;

namespace NavfertyExcelAddIn.UnitTests.Transliterate
{
	[TestClass]
	public class CyrillicLettersReplacerTests
	{
		private CyrillicLettersReplacer replacer;

		[TestInitialize]
		public void Initialize()
		{
			replacer = new CyrillicLettersReplacer();
		}

		[TestMethod]
		[DataRow("\u0410123\u0412\u0415", "A123BE")]
		[DataRow("ABCDEFG\u0410\u0411\u0412\u0413\u0414\u0415", "ABCDEFGAБBГДE")]
		[DataRow("абвгдеёжзийклмнопрстуфхцчшщъыьэюя", "aбbгдeёжзийkлmhoпpctyфxцчшщъыьэюя")]
		[DataRow("АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯ", "AБBГДEЁЖЗИЙKЛMHOПPCTYФXЦЧШЩЪЫЬЭЮЯ")]
		public void Transliterate(string input, string expected)
		{
			var output = replacer.ReplaceCyrillicCharsWithLatin(input);

			Assert.AreEqual(expected, output);
		}

	}
}
