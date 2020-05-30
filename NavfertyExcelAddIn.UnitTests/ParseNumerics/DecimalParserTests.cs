using System;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using NavfertyExcelAddIn.ParseNumerics;

namespace NavfertyExcelAddIn.UnitTests.ParseNumerics
{
	[TestClass]
	public class DecimalParserTests
	{
		[DataRow("0", 0)]
		[DataRow("123", 123)]
		[DataRow("1.1", 1.1)]
		[DataRow("1,1", 1.1)]
		[DataRow("1 000", 1000)]
		[DataRow("12  00 00 . 12", 120000.12)]
		[DataRow("0.1", 0.1)]
		[DataRow(".123", 0.123)]
		[DataRow("10.000.000", 10000000)]
		[DataRow("10.000,12", 10000.12)]
		[DataRow("12,345.12", 12345.12)]
		[DataRow("12,345,000", 12345000)]
		[DataRow("1E6", 1000000)]
		[DataRow("1.2E3", 1200)]
		[DataRow("-1.3E-5", -0.000013)]
		[TestMethod]
		public void ParseDecimal_Success(string sourceValue, double targetDoubleValue)
		{
			// DataRow can't into decimals - they are not primitive type
			var targetValue = Convert.ToDecimal(targetDoubleValue);

			Assert.AreEqual(targetValue, sourceValue.ParseDecimal());
		}

		[TestMethod]
		public void ParseDecimal_MissingValue()
		{
			var input = "no value";

			Assert.AreEqual(null, input.ParseDecimal());
		}

	}
}
