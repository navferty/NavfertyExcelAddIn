using NavfertyExcelAddIn.StringifyNumerics;

namespace NavfertyExcelAddIn.UnitTests.StringifyNumerics;

public class EnglishNumericStringifierTests
{
	[Test]
	[Arguments(0, "zero")]
	[Arguments(7, "seven")]
	[Arguments(42, "forty-two")]
	[Arguments(-5, "minus five")]
	[Arguments(12.3, "twelve and three tenths")]
	[Arguments(121.3, "one hundred and twenty-one and three tenths")]
	[Arguments(1.23, "one and twenty-three hundredths")]
	[Arguments(-11.023, "minus eleven and twenty-three thousandths")]
	[Arguments(1.03, "one and three hundredths")]
	[Arguments(0.023, "zero and twenty-three thousandths")]
	[Arguments(0.003, "zero and three thousandths")]
	[Arguments(0.0003, "zero")]
	[Arguments(1000, "one thousand")]
	[Arguments(9_000_000_000, "nine billion")]
	[Arguments(1_234_567, "one million two hundred and thirty-four thousand five hundred and sixty-seven")]
	[Arguments(-1_234_567, "minus one million two hundred and thirty-four thousand five hundred and sixty-seven")]
	[Arguments(2_123_654_321.059, "two billion one hundred and twenty-three million six hundred and fifty-four thousand three hundred and twenty-one and fifty-nine thousandths")]
	[Arguments(2_987_654_321.059, "two billion nine hundred and eighty-seven million six hundred and fifty-four thousand three hundred and twenty-one and fifty-nine thousandths")]
	[Arguments(-2_987_654_321.059, "minus two billion nine hundred and eighty-seven million six hundred and fifty-four thousand three hundred and twenty-one and fifty-nine thousandths")]
	public async Task StringifyNumber(double input, string expected)
	{
		var numericStringifier = new EnglishNumericStringifier();

		var output = numericStringifier.StringifyNumber(input);

		await Assert.That(output).IsEqualTo(expected);
	}
}
