using NavfertyExcelAddIn.StringifyNumerics;

namespace NavfertyExcelAddIn.UnitTests.StringifyNumerics;

public class FrenchNumericStringifierTests
{
	[Test]
	[Arguments(0, "zéro")]
	[Arguments(7, "sept")]
	[Arguments(42, "quarante-deux")]
	[Arguments(-5, "moins cinq")]
	[Arguments(12.3, "douze virgule trois")]
	[Arguments(121.3, "cent vingt et un virgule trois")]
	[Arguments(1.23, "un virgule vingt-trois")]
	[Arguments(11.023, "onze virgule zéro vingt-trois")]
	[Arguments(1.03, "un virgule zéro trois")]
	[Arguments(0.023, "zéro virgule zéro vingt-trois")]
	[Arguments(0.003, "zéro virgule zéro zéro trois")]
	[Arguments(0.0003, "zéro")]
	[Arguments(1000, "mille")]
	[Arguments(9_000_000_000, "neuf milliards")]
	[Arguments(1_234_567, "un million deux cent trente-quatre mille cinq cent soixante-sept")]
	[Arguments(-1_234_567, "moins un million deux cent trente-quatre mille cinq cent soixante-sept")]
	[Arguments(2_123_654_321.059, "deux milliards cent vingt-trois millions six cent cinquante-quatre mille trois cent vingt et un virgule zéro cinquante-neuf")]
	[Arguments(2_987_654_321.059, "deux milliards neuf cent quatre-vingt-sept millions six cent cinquante-quatre mille trois cent vingt et un virgule zéro cinquante-neuf")]
	[Arguments(-2_987_654_321.059, "moins deux milliards neuf cent quatre-vingt-sept millions six cent cinquante-quatre mille trois cent vingt et un virgule zéro cinquante-neuf")]
	public async Task StringifyNumber(double input, string expected)
	{
		var numericStringifier = new FrenchNumericStringifier();

		var output = numericStringifier.StringifyNumber(input);

		await Assert.That(output).IsEqualTo(expected);
	}
}
