using Microsoft.VisualStudio.TestTools.UnitTesting;
using NavfertyExcelAddIn.Transliterate;

namespace NavfertyExcelAddIn.UnitTests.Transliterate
{
	[TestClass]
	public class TransliteratorTests
	{
		private Transliterator transliterator;

		[TestInitialize]
		public void Initialize()
		{
			transliterator = new Transliterator();
		}

		[TestMethod]
		[DataRow("Карл у Клары украл кораллы", "Karl u Klary ukral korally")]
		[DataRow("Слишком много ножек у сороконожек", "Slishkom mnogo nozhek u sorokonozhek")]
		[DataRow("абвгдеёжзийклмнопрстуфхцчшщъыьэюя", "abvgdezhziiklmnoprstufkhtschshshchieyeiuia")]
		[DataRow("АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯ", "ABVGDEZhZIIKLMNOPRSTUFKhTsChShShchIeYEIuIa")]
		public void Transliterate(string input, string expected)
		{
			var output = transliterator.Transliterate(input);

			Assert.AreEqual(expected, output);
		}

		[TestMethod]
		public void Transliterate_EmptyString()
		{
			var output = transliterator.Transliterate(string.Empty);

			Assert.AreEqual(string.Empty, output);
		}

		[TestMethod]
		public void Transliterate_VeryLong()
		{
			string input = new string('ы', 1_000_000);

			var output = transliterator.Transliterate(input);

			Assert.AreEqual(1_000_000, output.Length);
			Assert.AreEqual("yyy", output.Substring(0, 3));
		}

		[TestMethod]
		public void Transliterate_SkipNonCyrillic()
		{
			// Greek Capital Letter Delta, Cyrillic Capital Letter Komi Lje, Arabic Letter Ghain
			string input = "Буквы \u0394\u0508\u063A"; // ΔԈغ

			var output = transliterator.Transliterate(input);

			Assert.AreEqual("Bukvy ", output);
		}
	}
}
