using Microsoft.VisualStudio.TestTools.UnitTesting;
using NavfertyExcelAddIn.StringifyNumerics;

namespace NavfertyExcelAddIn.UnitTests.StringifyNumerics
{
	[TestClass]
	public class FrenchNumericStringifierTests
	{
		private FrenchNumericStringifier numericStringifier;

		[TestInitialize]
		public void Initialize()
		{
			numericStringifier = new FrenchNumericStringifier();
		}

		[TestMethod]
		[DataRow(0, "zéro")]
		[DataRow(7, "sept")]
		[DataRow(42, "quarante-deux")]
		[DataRow(-5, "moins cinq")]
		[DataRow(12.3, "douze virgule trois")]
		[DataRow(1.23, "un virgule vingt-trois")]
		[DataRow(11.023, "onze virgule zéro vingt-trois")]
		[DataRow(1.03, "un virgule zéro trois")]
		[DataRow(0.023, "zéro virgule zéro vingt-trois")]
		[DataRow(0.003, "zéro virgule zéro zéro trois")]
		[DataRow(1000, "mille")]
		[DataRow(1_234_567, "un million deux cent trente-quatre mille cinq cent soixante-sept")]
		[DataRow(-1_234_567, "moins un million deux cent trente-quatre mille cinq cent soixante-sept")]
		public void StringifyNumber(double input, string expected)
		{
			var output = numericStringifier.StringifyNumber(input);

			Assert.AreEqual(expected, output);
		}
	}
}
