using Microsoft.VisualStudio.TestTools.UnitTesting;

using NavfertyExcelAddIn.StringifyNumerics;

namespace NavfertyExcelAddIn.UnitTests.StringifyNumerics
{
	[TestClass]
	public class RussianNumericStringifierTests
	{
		private RussianNumericStringifier numericStringifier;

		[TestInitialize]
		public void Initialize()
		{
			numericStringifier = new RussianNumericStringifier();
		}

		[TestMethod]
		[DataRow(0, "ноль")]
		[DataRow(7, "семь")]
		[DataRow(42, "сорок два")]
		[DataRow(-5, "минус пять")]
		[DataRow(1001.3, "одна тысяча одна целая и три десятых")] // двадцать три сотых или сотые?
		[DataRow(1.23, "одна целая и двадцать три сотых")] // двадцать три сотых или сотые?
		[DataRow(0.023, "ноль целых и двадцать три тысячных")]
		[DataRow(0.0003, "ноль")]
		[DataRow(1000, "одна тысяча")]
		[DataRow(1_234_567, "один миллион двести тридцать четыре тысячи пятьсот шестьдесят семь")]
		[DataRow(-1_234_567, "минус один миллион двести тридцать четыре тысячи пятьсот шестьдесят семь")]
		[DataRow(351, "триста пятьдесят один")]
		[DataRow(351.1, "триста пятьдесят одна целая и одна десятая")]
		[DataRow(351.2, "триста пятьдесят одна целая и две десятых")]
		[DataRow(352.1, "триста пятьдесят две целых и одна десятая")]
		[DataRow(0.01, "ноль целых и одна сотая")]
		[DataRow(0.02, "ноль целых и две сотых")]
		[DataRow(0.001, "ноль целых и одна тысячная")]
		[DataRow(0.002, "ноль целых и две тысячных")]
		public void StringifyNumber(double input, string expected)
		{
			var output = numericStringifier.StringifyNumber(input);

			Assert.AreEqual(expected, output);
		}
	}
}
