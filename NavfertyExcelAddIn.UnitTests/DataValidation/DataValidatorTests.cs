using System;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using NavfertyExcelAddIn.DataValidation;

namespace NavfertyExcelAddIn.UnitTests.DataValidation
{
	[TestClass]
	public class DataValidatorTests
	{
		private IValidatorFactory _dataValidatorFactory;

		[TestInitialize]
		public void BeforeEachTest()
		{
			_dataValidatorFactory = new ValidatorFactory();
		}

		[TestMethod]
		[DataRow(ValidationType.Date, "01.01.2011", true)]
		[DataRow(ValidationType.Date, "21.11.2011", true)]
		[DataRow(ValidationType.Date, "11/21/2011", true)]
		[DataRow(ValidationType.Date, "2011-11-21", true)]
		[DataRow(ValidationType.Date, "29.02.2012", true)]
		[DataRow(ValidationType.Date, "29.02.2011", false)]
		[DataRow(ValidationType.Date, "29.02.20111", false)]
		[DataRow(ValidationType.Date, "29022011", false)]
		[DataRow(ValidationType.Date, "123", false)]
		[DataRow(ValidationType.Date, "something", false)]

		[DataRow(ValidationType.TinOrganization, "something", false)]
		[DataRow(ValidationType.TinOrganization, "123", false)]
		[DataRow(ValidationType.TinOrganization, "1234567890", false)]
		[DataRow(ValidationType.TinOrganization, "7734052340", true)]
		[DataRow(ValidationType.TinOrganization, "7734110842", true)]

		[DataRow(ValidationType.TinPersonal, "7734110842", false)]
		[DataRow(ValidationType.TinPersonal, "123456789012", false)]
		[DataRow(ValidationType.TinPersonal, "526317984687", false)]
		[DataRow(ValidationType.TinPersonal, "526317984610", false)]
		[DataRow(ValidationType.TinPersonal, "526317984689", true)]

		[DataRow(ValidationType.Numeric, "5", true)]
		[DataRow(ValidationType.Numeric, "526", true)]
		[DataRow(ValidationType.Numeric, "1,2", true)]
		[DataRow(ValidationType.Numeric, "1.2", true)]
		[DataRow(ValidationType.Numeric, "526317984689", true)]
		[DataRow(ValidationType.Numeric, "abc", false)]
		[DataRow(ValidationType.Numeric, "123w", false)]
		[DataRow(ValidationType.Numeric, "0x0011", false)]
		[DataRow(ValidationType.Numeric, "0x00HH", false)]
		[DataRow(ValidationType.Numeric, "01.01.2015", false)]
		[DataRow(ValidationType.Numeric, "01/01/2015", false)]

		[DataRow(ValidationType.Xml, "12", true)]
		[DataRow(ValidationType.Xml, "1.2asd", true)]
		[DataRow(ValidationType.Xml, "asdads", true)]
		[DataRow(ValidationType.Xml, "русские буквы и корректные знаки препинания !№%:?*()@#$%^*()_+[]{}|\\/;", true)]
		[DataRow(ValidationType.Xml, "abc ", true)]
		[DataRow(ValidationType.Xml, "abc <", false)]
		[DataRow(ValidationType.Xml, "abc >", false)]
		[DataRow(ValidationType.Xml, "abc &", false)]
		[DataRow(ValidationType.Xml, "abc \"", false)]
		[DataRow(ValidationType.Xml, "abc '", false)]
		[DataRow(ValidationType.Xml, "abc ~", false)]
		[DataRow(ValidationType.Xml, "abc Џ", false)]
		[DataRow(ValidationType.Xml, "abc ´", false)]
		[DataRow(ValidationType.Xml, "abc `", false)]
		public void Validate(ValidationType validationType, string value, bool expected)
		{
			var validator = _dataValidatorFactory.CreateValidator(validationType);

			var result = validator.CheckValue(value);

			Assert.AreEqual(expected, result.IsSuccess);
		}

		[TestMethod]
		[DataRow("2012-01-01", true)]
		[DataRow("1799-01-01", true)]
		public void ValidateDate(string value, bool expected)
		{
			var validator = _dataValidatorFactory.CreateValidator(ValidationType.Date);
			var date = DateTime.Parse(value);

			var result = validator.CheckValue(date);

			Assert.AreEqual(expected, result.IsSuccess);
		}
	}
}
