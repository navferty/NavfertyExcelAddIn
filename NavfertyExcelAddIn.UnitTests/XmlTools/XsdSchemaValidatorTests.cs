using System.Xml;

using NavfertyExcelAddIn.XmlTools;

namespace NavfertyExcelAddIn.UnitTests.XmlTools;

public class XsdSchemaValidatorTests : TestsBase
{
	private const string XsdFile = "XmlTools/Samples/NO_PRIB_1_002_00_05_07_05.xsd";
	private const string CorrectXml = "XmlTools/Samples/Прибыль_Correct.xml";
	private const string IncorrectXml = "XmlTools/Samples/Прибыль_Incorrect.xml";
	private const string InvalidXml = "XmlTools/Samples/Прибыль_Invalid.xml";

	[Before(HookType.Test)]
	public void BeforeEachTest()
	{
		SetCulture();
	}

	[Test]
	public async Task ValidateXml_CorrectFile_NoErrors()
	{
		string xmlFilename = GetFilePath(CorrectXml);
		string xsdFilename = GetFilePath(XsdFile);
		var validator = new XsdSchemaValidator();

		var result = validator.Validate(xmlFilename, [xsdFilename]);

		await Assert.That(result).IsEmpty();
	}

	[Test]
	public async Task ValidateXml_IncorrectFile_AllErrorsFound()
	{
		string xmlFilename = GetFilePath(IncorrectXml);
		string xsdFilename = GetFilePath(XsdFile);
		var validator = new XsdSchemaValidator();

		var result = validator.Validate(xmlFilename, [xsdFilename]);

		await Assert.That(result).Count().EqualTo(3);
		await Assert.That(result.All(x => x.Severity == ValidationErrorSeverity.Error)).IsTrue();

		var errors = result.ToArray();
		await Assert.That(errors[0].ElementName).IsEqualTo("ОКТМО");
		await Assert.That(errors[0].Value).IsEqualTo("xxx");
		await Assert.That(errors[0].Message).Contains("ОКТМОТип");

		await Assert.That(errors[1].ElementName).IsEqualTo("КБК");
		await Assert.That(errors[1].Value).IsEqualTo("123");
		await Assert.That(errors[1].Message).Contains("КБКТип");

		await Assert.That(errors[2].ElementName).IsEqualTo("КБК");
		await Assert.That(errors[2].Value).IsEqualTo("321");
		await Assert.That(errors[2].Message).Contains("КБКТип");
	}

	[Test]
	public async Task ValidateXml_InvalidXmlFile_NoErrors()
	{
		string xmlFilename = GetFilePath(InvalidXml);
		string xsdFilename = GetFilePath(XsdFile);
		var validator = new XsdSchemaValidator();

		var ex = Assert.Throws<XmlException>(() => validator.Validate(xmlFilename, [xsdFilename]));

		await Assert.That(ex.Message).Contains("Документ");
	}
}
