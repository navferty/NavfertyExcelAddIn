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
		[DataRow(1.23, "one point twenty-three")]
		[DataRow(0.023, "zero point twenty-three")]
		[DataRow(1000, "one thousand")]
		[DataRow(1_234_567, "")]
		[DataRow(-1_234_567, "")]
		public void StringifyNumber(double input, string expected)
		{
			var output = numericStringifier.StringifyNumber(input);

			Assert.AreEqual(expected, output);
		}
	}
}
