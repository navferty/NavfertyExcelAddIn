using NavfertyExcelAddIn.DataValidation;

namespace NavfertyExcelAddIn.UnitTests.DataValidation;

public class DataValidatorTests
{
	[Test]
	[Arguments(ValidationType.Date, "01.01.2011", true)]
	[Arguments(ValidationType.Date, "21.11.2011", true)]
	[Arguments(ValidationType.Date, "11/21/2011", true)]
	[Arguments(ValidationType.Date, "2011-11-21", true)]
	[Arguments(ValidationType.Date, "29.02.2012", true)]
	[Arguments(ValidationType.Date, "29.02.2011", false)]
	[Arguments(ValidationType.Date, "29.02.20111", false)]
	[Arguments(ValidationType.Date, "29022011", false)]
	[Arguments(ValidationType.Date, "123", false)]
	[Arguments(ValidationType.Date, "something", false)]

	[Arguments(ValidationType.TinOrganization, "something", false)]
	[Arguments(ValidationType.TinOrganization, "123", false)]
	[Arguments(ValidationType.TinOrganization, "1234567890", false)]
	[Arguments(ValidationType.TinOrganization, "7734052340", true)]
	[Arguments(ValidationType.TinOrganization, "7734110842", true)]

	[Arguments(ValidationType.TinPersonal, "7734110842", false)]
	[Arguments(ValidationType.TinPersonal, "123456789012", false)]
	[Arguments(ValidationType.TinPersonal, "526317984687", false)]
	[Arguments(ValidationType.TinPersonal, "526317984610", false)]
	[Arguments(ValidationType.TinPersonal, "526317984689", true)]

	[Arguments(ValidationType.Numeric, "5", true)]
	[Arguments(ValidationType.Numeric, "526", true)]
	[Arguments(ValidationType.Numeric, "1,2", true)]
	[Arguments(ValidationType.Numeric, "1.2", true)]
	[Arguments(ValidationType.Numeric, "526317984689", true)]
	[Arguments(ValidationType.Numeric, "abc", false)]
	[Arguments(ValidationType.Numeric, "123w", false)]
	[Arguments(ValidationType.Numeric, "0x0011", false)]
	[Arguments(ValidationType.Numeric, "0x00HH", false)]
	[Arguments(ValidationType.Numeric, "01.01.2015", false)]
	[Arguments(ValidationType.Numeric, "01/01/2015", false)]

	[Arguments(ValidationType.Xml, "12", true)]
	[Arguments(ValidationType.Xml, "1.2asd", true)]
	[Arguments(ValidationType.Xml, "asdads", true)]
	[Arguments(ValidationType.Xml, "русские буквы и корректные знаки препинания !№%:?*()@#$%^*()_+[]{}|\\/;", true)]
	[Arguments(ValidationType.Xml, "abc ", true)]
	[Arguments(ValidationType.Xml, "abc <", false)]
	[Arguments(ValidationType.Xml, "abc >", false)]
	[Arguments(ValidationType.Xml, "abc &", false)]
	[Arguments(ValidationType.Xml, "abc \"", false)]
	[Arguments(ValidationType.Xml, "abc '", false)]
	[Arguments(ValidationType.Xml, "abc ~", false)]
	[Arguments(ValidationType.Xml, "abc Џ", false)]
	[Arguments(ValidationType.Xml, "abc ´", false)]
	[Arguments(ValidationType.Xml, "abc `", false)]
	public async Task Validate(ValidationType validationType, string value, bool expected)
	{
		var dataValidatorFactory = new ValidatorFactory();
		var validator = dataValidatorFactory.CreateValidator(validationType);

		var result = validator.CheckValue(value);

		await Assert.That(result.IsSuccess).IsEqualTo(expected);
	}

	[Test]
	[Arguments("2012-01-01", true)]
	[Arguments("1799-01-01", true)]
	public async Task ValidateDate(string value, bool expected)
	{
		var dataValidatorFactory = new ValidatorFactory();
		var validator = dataValidatorFactory.CreateValidator(ValidationType.Date);
		var date = DateTime.Parse(value);

		var result = validator.CheckValue(date);

		await Assert.That(result.IsSuccess).IsEqualTo(expected);
	}
}
