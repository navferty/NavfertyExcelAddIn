using NavfertyExcelAddIn.StringifyNumerics;

namespace NavfertyExcelAddIn.UnitTests.StringifyNumerics;

public class RussianNumericStringifierTests
{
	[Test]
	[Arguments(0, "ноль")]
	[Arguments(7, "семь")]
	[Arguments(42, "сорок два")]
	[Arguments(-5, "минус пять")]
	[Arguments(1001.3, "одна тысяча одна целая и три десятых")] // двадцать три сотых или сотые?
	[Arguments(1.23, "одна целая и двадцать три сотых")] // двадцать три сотых или сотые?
	[Arguments(0.023, "ноль целых и двадцать три тысячных")]
	[Arguments(0.0003, "ноль")]
	[Arguments(1000, "одна тысяча")]
	[Arguments(1_234_567, "один миллион двести тридцать четыре тысячи пятьсот шестьдесят семь")]
	[Arguments(-1_234_567, "минус один миллион двести тридцать четыре тысячи пятьсот шестьдесят семь")]
	[Arguments(351, "триста пятьдесят один")]
	[Arguments(351.1, "триста пятьдесят одна целая и одна десятая")]
	[Arguments(351.2, "триста пятьдесят одна целая и две десятых")]
	[Arguments(352.1, "триста пятьдесят две целых и одна десятая")]
	[Arguments(0.01, "ноль целых и одна сотая")]
	[Arguments(0.02, "ноль целых и две сотых")]
	[Arguments(0.001, "ноль целых и одна тысячная")]
	[Arguments(0.002, "ноль целых и две тысячных")]
	public async Task StringifyNumber(double input, string expected)
	{
		var numericStringifier = new RussianNumericStringifier();

		var output = numericStringifier.StringifyNumber(input);

		await Assert.That(output).IsEqualTo(expected);
	}
}
