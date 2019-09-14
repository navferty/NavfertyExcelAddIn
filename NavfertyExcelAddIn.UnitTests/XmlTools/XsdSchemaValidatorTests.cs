using System.Linq;
using System.Xml;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using NavfertyExcelAddIn.XmlTools;

namespace NavfertyExcelAddIn.UnitTests.XmlTools
{
    [TestClass]
    public class XsdSchemaValidatorTests : TestsBase
    {

        private XsdSchemaValidator validator;

        private const string XsdFile = "XmlTools/Samples/NO_PRIB_1_002_00_05_07_05.xsd";
        private const string CorrectXml = "XmlTools/Samples/Прибыль_Correct.xml";
        private const string IncorrectXml = "XmlTools/Samples/Прибыль_Incorrect.xml";
        private const string InvalidXml = "XmlTools/Samples/Прибыль_Invalid.xml";

        [TestInitialize]
        public void BeforeEachTest()
        {
            SetCulture();

            validator = new XsdSchemaValidator();
        }

        [TestMethod]
        public void ValidateXml_CorrectFile_NoErrors()
        {
            string xmlFilename = GetFilePath(CorrectXml);
            string xsdFilename = GetFilePath(XsdFile);


            var result = validator.Validate(xmlFilename, new[] { xsdFilename });

            Assert.AreEqual(0, result.Count);
        }

        [TestMethod]
        public void ValidateXml_IncorrectFile_AllErrorsFound()
        {
            string xmlFilename = GetFilePath(IncorrectXml);
            string xsdFilename = GetFilePath(XsdFile);


            var result = validator.Validate(xmlFilename, new[] { xsdFilename });

            Assert.AreEqual(3, result.Count);
            Assert.IsTrue(result.All(x => x.Severity == ValidationErrorSeverity.Error));

            var errors = result.ToArray();
            Assert.AreEqual("ОКТМО", errors[0].ElementName);
            Assert.AreEqual("xxx", errors[0].Value);
            StringAssert.Contains(errors[0].Message, "ОКТМОТип");

            Assert.AreEqual("КБК", errors[1].ElementName);
            Assert.AreEqual("123", errors[1].Value);
            StringAssert.Contains(errors[1].Message, "КБКТип");

            Assert.AreEqual("КБК", errors[2].ElementName);
            Assert.AreEqual("321", errors[2].Value);
            StringAssert.Contains(errors[2].Message, "КБКТип");
        }

        [TestMethod]
        public void ValidateXml_InvalidXmlFile_NoErrors()
        {
            string xmlFilename = GetFilePath(InvalidXml);
            string xsdFilename = GetFilePath(XsdFile);


            var ex = Assert.ThrowsException<XmlException>(() => validator.Validate(xmlFilename, new[] { xsdFilename }));

            StringAssert.Contains(ex.Message, "Документ");
        }
    }
}
