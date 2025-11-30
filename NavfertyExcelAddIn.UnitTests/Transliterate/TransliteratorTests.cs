using NavfertyExcelAddIn.Transliterate;

namespace NavfertyExcelAddIn.UnitTests.Transliterate
{
	public class TransliteratorTests
	{
		[Test]
		[Arguments("Карл у Клары украл кораллы", "Karl u Klary ukral korally")]
		[Arguments("Слишком много ножек у сороконожек", "Slishkom mnogo nozhek u sorokonozhek")]
		[Arguments("абвгдеёжзийклмнопрстуфхцчшщъыьэюя", "abvgdeezhziiklmnoprstufkhtschshshchieyeiuia")]
		[Arguments("АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯ", "ABVGDEEZhZIIKLMNOPRSTUFKhTsChShShchIeYEIuIa")]
		public async Task Transliterate(string input, string expected)
		{
			var transliterator = new Transliterator();

			var output = transliterator.Transliterate(input);

			await Assert.That(output).IsEqualTo(expected);
		}

		[Test]
		public async Task Transliterate_EmptyString()
		{
			var transliterator = new Transliterator();

			var output = transliterator.Transliterate(string.Empty);

			await Assert.That(output).IsEqualTo(string.Empty);
		}

		[Test]
		public async Task Transliterate_VeryLong()
		{
			string input = new string('ы', 1_000_000);
			var transliterator = new Transliterator();

			var output = transliterator.Transliterate(input);

			await Assert.That(output.Length).IsEqualTo(1_000_000);
			await Assert.That(output.Substring(0, 3)).IsEqualTo("yyy");
		}

		[Test]
		public async Task Transliterate_SkipNonCyrillic()
		{
			// Greek Capital Letter Delta, Cyrillic Capital Letter Komi Lje, Arabic Letter Ghain
			string input = "Буквы \u0394\u0508\u063A"; // ???
			var transliterator = new Transliterator();

			var output = transliterator.Transliterate(input);

			await Assert.That(output).IsEqualTo("Bukvy ");
		}
	}
}
