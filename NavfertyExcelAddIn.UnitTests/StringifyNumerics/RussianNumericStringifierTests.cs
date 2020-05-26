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
		[DataRow(1.230, "одна целая и двести тридцать тысячных")] // 230 тысячных...
		[DataRow(0.023, "ноль целых и двадцать три тысячные")]
		[DataRow(1000, "одна тысяча")]
		[DataRow(1_234_567, "один миллион двести тридцать четыре тысячи пятьсот шестьдесят семь")]
		[DataRow(-1_234_567, "минус один миллион двести тридцать четыре тысячи пятьсот шестьдесят семь")]
		public void StringifyNumber(double input, string expected)
		{
			var output = numericStringifier.StringifyNumber(input);

			Assert.AreEqual(expected, output);
		}
	}
}
