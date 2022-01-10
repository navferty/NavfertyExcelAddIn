using System;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using NavfertyExcelAddIn.ParseNumerics;

namespace NavfertyExcelAddIn.UnitTests.ParseNumerics
{
	[TestClass]
	public class DecimalParserTests
	{
		[DataRow("0", 0, null)]
		[DataRow("123", 123, null)]
		[DataRow("1.1", 1.1, null)]
		[DataRow("1,1", 1.1, null)]
		[DataRow("1 000", 1000, null)]
		[DataRow("12  00 00 . 12", 120000.12, null)]
		[DataRow("0.1", 0.1, null)]
		[DataRow(".123", 0.123, null)]
		[DataRow("10.000.000", 10000000, null)]
		[DataRow("10.000,12", 10000.12, null)]
		[DataRow("12,345.12", 12345.12, null)]
		[DataRow("12,345,000", 12345000, null)]
		[DataRow("1E6", 1000000, null)]
		[DataRow("1.2E3", 1200, null)]
		[DataRow("-1.3E-5", -0.000013, null)]
		[DataRow("$5666", 5666, "$")]
		[DataRow("5666$", 5666, "$")]
		[DataRow("$56.66", 56.66, "$")]
		[DataRow("56.66$", 56.66, "$")]
		[DataRow("$56,66", 56.66, "$")]
		[DataRow("56,66$", 56.66, "$")]
		[DataRow("₽5666", 5666, "₽")]
		[DataRow("5666₽", 5666, "₽")]
		[DataRow("₽56.66", 56.66, "₽")]
		[DataRow("56.66₽", 56.66, "₽")]
		[DataRow("₽56,66", 56.66, "₽")]
		[DataRow("56,66₽", 56.66, "₽")]
		[DataRow("£5666", 5666, "£")]
		[DataRow("5666£", 5666, "£")]
		[DataRow("£56.66", 56.66, "£")]
		[DataRow("56.66£", 56.66, "£")]
		[DataRow("£56.66", 56.66, "£")]
		[DataRow("56.66£", 56.66, "£")]
		[TestMethod]
		public void ParseDecimal_Success(string sourceValue, double targetDoubleValue, string currencySymbol)
		{
			var targetValue = new NumericParseResult((decimal)targetDoubleValue, currencySymbol);
			// DataRow can't into decimals - they are not primitive type
			//var targetValue = Convert.ToDecimal(targetDoubleValue);
			Assert.AreEqual(targetValue, sourceValue.ParseDecimal());
		}

		[TestMethod]
		public void ParseDecimal_MissingValue()
		{
			var input = "no value";
			Assert.AreEqual(null, input.ParseDecimal());
		}

		[TestMethod]
		public void ParseDecimal_UnknownCurrency()
		{
			var input = "56.66%";
			Assert.AreEqual(null, input.ParseDecimal());
		}

		[TestMethod]
		public void ParseDecimal_MultipleCurrencyInOneString()
		{
			Assert.AreEqual(null, "56.66£₽".ParseDecimal());
			Assert.AreEqual(null, "56.66₽£".ParseDecimal());
			Assert.AreEqual(null, "₽£56.66".ParseDecimal());
		}



	}
}
