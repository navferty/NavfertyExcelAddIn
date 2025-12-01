using System.Reflection;

using Microsoft.Office.Interop.Excel;

using Moq;

using Navferty.Common;

using NavfertyExcelAddIn.UnitTests.Builders;
using NavfertyExcelAddIn.XmlTools;

using Range = Microsoft.Office.Interop.Excel.Range;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace NavfertyExcelAddIn.UnitTests.XmlTools;

// TODO
[NotInParallel]
public class XmlValidatorTests : TestsBase
{
    private Mock<IDialogService> dialogService = null!;
    private Mock<IXsdSchemaValidator> xsdSchemaValidator = null!;
    private Mock<Application> app = null!;
    private Mock<Range> range = null!;
    private XmlValidator validator = null!;

    [Before(HookType.Test)]
    public void BeforeEachTest()
    {
        dialogService = new Mock<IDialogService>(MockBehavior.Strict);
        xsdSchemaValidator = new Mock<IXsdSchemaValidator>(MockBehavior.Strict);

        SetupApplicationForAddWorkbook();

        validator = new XmlValidator(dialogService.Object, xsdSchemaValidator.Object);
    }

    [Test]
    public async Task Validate_XmlNotSelected()
    {
        dialogService
            .Setup(x => x.AskForFiles(It.IsAny<bool>(), FileType.Xml))
            .Returns([]);


        validator.Validate(app.Object);

        dialogService.VerifyAll();
    }

    [Test]
    public async Task Validate_XsdNotSelected()
    {
        dialogService
            .Setup(x => x.AskForFiles(It.IsAny<bool>(), FileType.Xml))
            .Returns(["file.xml"]);

        dialogService
            .SetupSequence(x => x.AskForFiles(It.IsAny<bool>(), FileType.Xsd))
            .Returns([]);


        validator.Validate(app.Object);

        dialogService.VerifyAll();
    }

    [Test]
    public async Task Validate_FilesNotExist_Throws()
    {
        dialogService
            .SetupSequence(x => x.AskForFiles(It.IsAny<bool>(), FileType.Xsd))
            .Returns(["file.xsd"]);
        dialogService
            .SetupSequence(x => x.AskForFiles(It.IsAny<bool>(), FileType.Xml))
            .Returns(["file.xml"]);


        var ex = Assert.Throws<ArgumentException>(() => validator.Validate(app.Object));

        await Assert.That(ex.Message).IsEqualTo("One or more files not found");
        dialogService.VerifyAll();
    }


    [Test]
    public async Task Validate_NoErrors_SuccessMessageShown()
    {
        dialogService
            .SetupSequence(x => x.AskForFiles(It.IsAny<bool>(), FileType.Xsd))
            .Returns([GetFilePath("XmlTools/Samples/NO_PRIB_1_002_00_05_07_05.xsd")]);
        dialogService
            .SetupSequence(x => x.AskForFiles(It.IsAny<bool>(), FileType.Xml))
            .Returns([GetFilePath("XmlTools/Samples/Прибыль_Correct.xml")]);

        xsdSchemaValidator
            .Setup(x => x.Validate(It.IsAny<string>(), It.IsAny<IReadOnlyCollection<string>>()))
            .Returns([]);

        dialogService
            .SetupSequence(x => x.ShowInfo(It.IsAny<string>()));

        SetCulture();

        validator.Validate(app.Object);

        dialogService.VerifyAll();
        dialogService.Verify(x => x.ShowInfo(It.Is<string>(s => s.Contains("Successfully"))));
    }

    [Test]
    public async Task Validate_HasErrors_MessageShown()
    {
        dialogService
            .SetupSequence(x => x.AskForFiles(It.IsAny<bool>(), FileType.Xsd))
            .Returns([GetFilePath("XmlTools/Samples/NO_PRIB_1_002_00_05_07_05.xsd")]);
        dialogService
            .SetupSequence(x => x.AskForFiles(It.IsAny<bool>(), FileType.Xml))
            .Returns([GetFilePath("XmlTools/Samples/Прибыль_Correct.xml")]);

        var error = new XmlValidationError(ValidationErrorSeverity.Error, "message", "value", "element name");
        xsdSchemaValidator
            .Setup(x => x.Validate(It.IsAny<string>(), It.IsAny<IReadOnlyCollection<string>>()))
                .Returns([error]);

        SetCulture();

        validator.Validate(app.Object);

        range.Verify(x => x.set_Value(It.IsAny<object>(), It.Is<string>(s => s.Contains("element name"))));
        range.Verify(x => x.set_Value(It.IsAny<object>(), It.IsAny<string>()), Times.Exactly(8));
    }

    private void SetupApplicationForAddWorkbook()
    {
        app = new Mock<Application>(MockBehavior.Strict);
        var wbs = new Mock<Workbooks>(MockBehavior.Strict);
        var wb = new Mock<Workbook>(MockBehavior.Strict);
        var wss = new Mock<Sheets>(MockBehavior.Strict);
        var ws = new Mock<Worksheet>(MockBehavior.Strict);
        //wss.Setup(x => x[1]).re
        var rangeBuilder = new RangeBuilder().WithSetValue();
        range = rangeBuilder.MockObject;
        range.Setup(x => x.get_Value(It.IsAny<object>()));
        range.Setup(x => x.set_Value(It.IsAny<object>(), It.IsAny<object>()));

        ws.Setup(x => x.Cells[It.IsAny<int>(), It.IsAny<int>()]).Returns(rangeBuilder.Build());
        wb.Setup(x => x.Worksheets).Returns(wss.Object);
        wss.Setup(x => x[1]).Returns(ws.Object);
        wbs.Setup(x => x.Add(Missing.Value)).Returns(wb.Object);
        app.Setup(x => x.Workbooks).Returns(wbs.Object);
    }
}
