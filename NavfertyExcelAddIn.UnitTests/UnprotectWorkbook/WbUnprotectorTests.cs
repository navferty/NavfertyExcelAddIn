using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

using Moq;

using Navferty.Common;

using NavfertyExcelAddIn.UnprotectWorkbook;

using Range = Microsoft.Office.Interop.Excel.Range;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace NavfertyExcelAddIn.UnitTests.UnprotectWorkbook;

// ignored in azure-pipelines-CI.yml by TestCategory 'Automation'
[Category("Automation")]
[NotInParallel("Automation")]
public class WbUnprotectorTests : TestsBase
{
    private Application app = null!;


    [Before(HookType.Test)]
    public void Initialize()
    {
        app = OpenNewExcelApp();
    }

    [Test]
    public async Task UnprotectWorkbookAndWorksheet_CanEditAndAddWorksheet()
    {
        var dialogService = new Mock<IDialogService>(MockBehavior.Loose);
        var wbUnprotector = new WbUnprotector(dialogService.Object);

        var path = GetFilePath("UnprotectWorkbook/ProtectedWorkbook.xlsx");

        await Assert.That(File.Exists(path)).IsTrue();

        wbUnprotector.UnprotectWorkbookWithAllWorksheets(path);

        var wb = app.Workbooks.Open(path);
        var ws = (Worksheet)wb.Worksheets[1];

        // test ws is unlocked
        ((Range)ws.Cells[1, 1]).Value = "cba";

        // test wb is unlocked
        wb.Worksheets.Add();
    }

    private static Application OpenNewExcelApp()
    {
        return new Application
        {
            Visible = false,
            EnableEvents = false,
            DisplayAlerts = false,
            AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable
        };
    }

    [After(HookType.Test)]
    public void Cleanup()
    {
        app.Quit();
    }
}
