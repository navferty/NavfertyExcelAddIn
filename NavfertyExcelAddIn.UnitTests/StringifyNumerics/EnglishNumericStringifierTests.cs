using Microsoft.VisualStudio.TestTools.UnitTesting;
using NavfertyExcelAddIn.StringifyNumerics;

namespace NavfertyExcelAddIn.UnitTests.StringifyNumerics
{
	[TestClass]
	public class EnglishNumericStringifierTests
	{
		private EnglishNumericStringifier numericStringifier;

		[TestInitialize]
		public void Initialize()
		{
			numericStringifier = new EnglishNumericStringifier();
		}

		[TestMethod]
		[DataRow(0, "zero")]
		[DataRow(7, "seven")]
		[DataRow(42, "forty-two")]
		[DataRow(-5, "minus five")]
		[DataRow(12.3, "twelve and three tenths")]
		[DataRow(1.23, "one and twenty-three hundredths")]
		[DataRow(1.03, "one and three hundredths")]
		[DataRow(0.023, "zero and twenty-three thousandths")]
		[DataRow(0.003, "zero and three thousandths")]
		[DataRow(1000, "one thousand")]
		[DataRow(1_234_567, "one million two hundred and thirty-four thousand five hundred and sixty-seven")]
		[DataRow(-1_234_567, "minus one million two hundred and thirty-four thousand five hundred and sixty-seven")]
		public void StringifyNumber(double input, string expected)
		{
			var output = numericStringifier.StringifyNumber(input);

			Assert.AreEqual(expected, output);
		}
	}
}
