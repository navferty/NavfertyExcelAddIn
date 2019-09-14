using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;

using NavfertyExcelAddIn.XmlTools;
using NavfertyExcelAddIn.Commons;

namespace NavfertyExcelAddIn.UnitTests.XmlTools
{
    [TestClass]
    public class XmlValidatorTests : TestsBase
    {
        private Mock<IDialogService> dialogService;
        private Mock<IXsdSchemaValidator> xsdSchemaValidator;

        private XmlValidator validator;

        [TestInitialize]
        public void BeforeEachTest()
        {
            dialogService = new Mock<IDialogService>(MockBehavior.Strict);
            xsdSchemaValidator = new Mock<IXsdSchemaValidator>(MockBehavior.Strict);

            validator = new XmlValidator(dialogService.Object, xsdSchemaValidator.Object);
        }

        [TestMethod]
        public void Validate_XmlNotSelected()
        {
            dialogService
                .Setup(x => x.AskForFiles(It.IsAny<bool>(), FileType.Xml))
                .Returns(Array.Empty<string>());


            validator.Validate();

            dialogService.VerifyAll();
        }

        [TestMethod]
        public void Validate_XsdNotSelected()
        {
            dialogService
                .Setup(x => x.AskForFiles(It.IsAny<bool>(), FileType.Xml))
                .Returns(new[] { "file.xml" });

            dialogService
                .SetupSequence(x => x.AskForFiles(It.IsAny<bool>(), FileType.Xsd))
                .Returns(Array.Empty<string>());


            validator.Validate();

            dialogService.VerifyAll();
        }

        [TestMethod]
        public void Validate_FilesNotExist_Throws()
        {
            dialogService
                .SetupSequence(x => x.AskForFiles(It.IsAny<bool>(), FileType.Xsd))
                .Returns(new[] { "file.xsd" });
            dialogService
                .SetupSequence(x => x.AskForFiles(It.IsAny<bool>(), FileType.Xml))
                .Returns(new[] { "file.xml" });


            var ex = Assert.ThrowsException<ArgumentException>(() => validator.Validate());

            Assert.AreEqual("One or more files not found", ex.Message);
            dialogService.VerifyAll();
        }


        [TestMethod]
        public void Validate_NoErrors_SuccessMessageShown()
        {
            dialogService
                .SetupSequence(x => x.AskForFiles(It.IsAny<bool>(), FileType.Xsd))
                .Returns(new[] { GetFilePath("XmlTools/Samples/NO_PRIB_1_002_00_05_07_05.xsd") });
            dialogService
                .SetupSequence(x => x.AskForFiles(It.IsAny<bool>(), FileType.Xml))
                .Returns(new[] { GetFilePath("XmlTools/Samples/Прибыль_Correct.xml") });

            xsdSchemaValidator
                .Setup(x => x.Validate(It.IsAny<string>(), It.IsAny<string[]>()))
                .Returns(Array.Empty<XmlValidationError>());

            dialogService
                .SetupSequence(x => x.ShowInfo(It.IsAny<string>()));

            SetCulture();

            validator.Validate();

            dialogService.VerifyAll();
            dialogService.Verify(x => x.ShowInfo(It.Is<string>(s => s.Contains("Successfully"))));
        }
    }
}
