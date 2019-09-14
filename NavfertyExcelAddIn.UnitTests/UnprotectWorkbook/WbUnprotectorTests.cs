﻿using System.IO;
using Moq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

using NavfertyExcelAddIn.Commons;
using NavfertyExcelAddIn.UnprotectWorkbook;

using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace NavfertyExcelAddIn.UnitTests.UnprotectWorkbook
{
    [TestClass]
    public class WbUnprotectorTests : TestsBase
    {
        private WbUnprotector wbUnprotector;
        private Application app;


        [TestInitialize]
        public void Initialize()
        {
            app = OpenNewExcelApp();
        }

        [TestMethod]
        // ignored in azure-pipelines-CI.yml by name
        public void UnprotectWorkbookAndWorksheet_CanEditAndAddWorksheet()
        {
            var dialogService = new Mock<IDialogService>(MockBehavior.Loose);
            wbUnprotector = new WbUnprotector(dialogService.Object);

            var path = GetFilePath("UnprotectWorkbook/ProtectedWorkbook.xlsx");

            Assert.IsTrue(File.Exists(path));

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

        [TestCleanup]
        public void Cleanup()
        {
            app.Quit();
        }
    }
}
