using Moq;

using Navferty.Common;

using NavfertyExcelAddIn.XmlTools;

namespace NavfertyExcelAddIn.UnitTests.XmlTools;

public class XmlSampleCreatorTests : TestsBase
{
    private Mock<IDialogService> dialogService = null!;
    private XmlSampleCreator xmlSampleCreator = null!;

    private const string XsdFile = "XmlTools/Samples/NO_PRIB_1_002_00_05_07_05.xsd";

    [Before(HookType.Test)]
    public void BeforeEachTest()
    {
        dialogService = new Mock<IDialogService>(MockBehavior.Strict);

        xmlSampleCreator = new XmlSampleCreator(dialogService.Object);
    }

    [Test]
    public async Task CreateSample_XsdNotSelected()
    {
        dialogService
            .Setup(x => x.AskForFiles(It.IsAny<bool>(), FileType.Xsd))
            .Returns([]);

        xmlSampleCreator.CreateSampleXml();

        dialogService.VerifyAll();
    }

    [Test]
    public async Task CreateSample_TargetXmlNotSelected()
    {
        dialogService
            .Setup(x => x.AskForFiles(false, FileType.Xsd))
            .Returns([GetFilePath(XsdFile)]);

        dialogService
            .Setup(x => x.AskFileNameSaveAs(It.IsAny<string>(), FileType.Xml))
            .Returns(string.Empty);

        xmlSampleCreator.CreateSampleXml();

        dialogService.VerifyAll();
    }

    [Test]
    public async Task CreateSample_CreatedSuccessfully()
    {
        dialogService
            .Setup(x => x.AskForFiles(false, FileType.Xsd))
            .Returns([GetFilePath(XsdFile)]);

        var path = GetFilePath($"sample_{DateTime.Now:yyyy-MM-d_HH-mm-ss}.xml");
        await Assert.That(File.Exists(path)).IsFalse();

        dialogService
            .Setup(x => x.AskFileNameSaveAs(It.IsAny<string>(), FileType.Xml))
            .Returns(path);

        xmlSampleCreator.CreateSampleXml();

        dialogService.VerifyAll();
        await Assert.That(File.Exists(path)).IsTrue();
    }
}
