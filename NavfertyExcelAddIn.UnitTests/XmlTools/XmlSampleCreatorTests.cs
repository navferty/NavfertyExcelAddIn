using System;
using System.IO;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using Moq;

using NavfertyExcelAddIn.Commons;
using NavfertyExcelAddIn.XmlTools;

namespace NavfertyExcelAddIn.UnitTests.XmlTools
{
    [TestClass]
    public class XmlSampleCreatorTests : TestsBase
    {
        private Mock<IDialogService> dialogService;

        private XmlSampleCreator xmlSampleCreator;

        private const string XsdFile = "XmlTools/Samples/NO_PRIB_1_002_00_05_07_05.xsd";

        [TestInitialize]
        public void BeforeEachTest()
        {
            dialogService = new Mock<IDialogService>(MockBehavior.Strict);

            xmlSampleCreator = new XmlSampleCreator(dialogService.Object);
        }

        [TestMethod]
        public void CreateSample_XsdNotSelected()
        {
            dialogService
                .Setup(x => x.AskForFiles(It.IsAny<bool>(), FileType.Xsd))
                .Returns(Array.Empty<string>());

            xmlSampleCreator.CreateSampleXml();

            dialogService.VerifyAll();
        }

        [TestMethod]
        public void CreateSample_TargetXmlNotSelected()
        {
            dialogService
                .Setup(x => x.AskForFiles(false, FileType.Xsd))
                .Returns(new[] { GetFilePath(XsdFile) });

            dialogService
                .Setup(x => x.AskFileNameSaveAs(It.IsAny<string>(), FileType.Xml))
                .Returns(string.Empty);

            xmlSampleCreator.CreateSampleXml();

            dialogService.VerifyAll();
        }

        [TestMethod]
        public void CreateSample_CreatedSuccessfully()
        {
            dialogService
                .Setup(x => x.AskForFiles(false, FileType.Xsd))
                .Returns(new[] { GetFilePath(XsdFile) });

            var path = GetFilePath($"sample_{DateTime.Now:yyyy-MM-d_HH-mm-ss}.xml");
            Assert.IsFalse(File.Exists(path));

            dialogService
                .Setup(x => x.AskFileNameSaveAs(It.IsAny<string>(), FileType.Xml))
                .Returns(path);

            xmlSampleCreator.CreateSampleXml();

            dialogService.VerifyAll();
            Assert.IsTrue(File.Exists(path));
        }
    }
}
