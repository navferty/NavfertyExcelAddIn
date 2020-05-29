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
		[DataRow(121.3, "cent vingt et un virgule trois")]
		[DataRow(1.23, "un virgule vingt-trois")]
		[DataRow(11.023, "onze virgule zéro vingt-trois")]
		[DataRow(1.03, "un virgule zéro trois")]
		[DataRow(0.023, "zéro virgule zéro vingt-trois")]
		[DataRow(0.003, "zéro virgule zéro zéro trois")]
		[DataRow(0.0003, "zéro")]
		[DataRow(1000, "mille")]
		[DataRow(9_000_000_000, "neuf milliards")]
		[DataRow(1_234_567, "un million deux cent trente-quatre mille cinq cent soixante-sept")]
		[DataRow(-1_234_567, "moins un million deux cent trente-quatre mille cinq cent soixante-sept")]
		[DataRow(2_123_654_321.059, "deux milliards cent vingt-trois millions six cent cinquante-quatre mille trois cent vingt et un virgule zéro cinquante-neuf")]
		[DataRow(2_987_654_321.059, "deux milliards neuf cent quatre-vingt-sept millions six cent cinquante-quatre mille trois cent vingt et un virgule zéro cinquante-neuf")]
		[DataRow(-2_987_654_321.059, "moins deux milliards neuf cent quatre-vingt-sept millions six cent cinquante-quatre mille trois cent vingt et un virgule zéro cinquante-neuf")]
		public void StringifyNumber(double input, string expected)
		{
			var output = numericStringifier.StringifyNumber(input);

			Assert.AreEqual(expected, output);
		}
	}
}
