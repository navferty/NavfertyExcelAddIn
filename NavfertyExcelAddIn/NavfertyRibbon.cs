using System;
using System.Linq;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Diagnostics.CodeAnalysis;

using NavfertyExcelAddIn.ParseNumerics;
using NavfertyExcelAddIn.FindFormulaErrors;
using NavfertyExcelAddIn.UnprotectWorkbook;
using NavfertyExcelAddIn.Localization;
using NavfertyExcelAddIn.Commons;

using NLog;
using Autofac;

using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

using Application = Microsoft.Office.Interop.Excel.Application;

namespace NavfertyExcelAddIn
{
    [ComVisible(true)]
    [SuppressMessage("Style", "IDE0060:Remove unused parameter", Justification = "Ribbon callbacks must have certain signature")]
    public class NavfertyRibbon : IRibbonExtensibility, IDisposable
    {
        private static readonly IContainer container = Registry.CreateContainer();
        private readonly ILogger logger = LogManager.GetCurrentClassLogger();
        private readonly IDialogService dialogService = container.Resolve<IDialogService>();

        private Application App => Globals.ThisAddIn.Application;

        #region Forms
        private SearchRangeResultForm form;
        #endregion

        public NavfertyRibbon()
        {
            // TODO
        }

        #region IRibbonExtensibility
        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText();
        }
        #endregion

        #region Ribbon callbacks
        public void RibbonLoad(IRibbonUI ribbonUI)
        {
            // TODO
        }

        public void ParseNumerics(IRibbonControl ribbonControl)
        {
            var selection = (Range)App.Selection;

            if (selection == null)
                return;

            logger.Debug($"Parse numerics for range {selection.Address}");

            using (var scope = container.BeginLifetimeScope())
            {
                var parser = scope.Resolve<INumericParser>();
                parser.Parse(selection);
            }
        }

        public void UnprotectWorkbook(IRibbonControl ribbonControl)
        {
            var wb = App.ActiveWorkbook;
            var path = wb.FullName;

            var extension = path.Split('.').LastOrDefault();

            if (extension != "xlsx" && extension != "xlsm")
            {
                dialogService.ShowError(UIStrings.CannotUnlockPleaseSaveAsXml);
                return;
            }

            if (!dialogService.Ask(UIStrings.UnsavedChangesWillBeLostPrompt, UIStrings.Warning))
            {
                return;
            }

            wb.Close(false);

            using (var scope = container.BeginLifetimeScope())
            {
                var wbUnprotector = scope.Resolve<IWbUnprotector>();
                wbUnprotector.UnprotectWorkbookWithAllWorksheets(path);
            }

            App.Workbooks.Open(path);
        }

        public void CutNames(IRibbonControl ribbonControl)
        {
            var selection = (Range)App.Selection;

            if (selection == null)
                return;

            logger.Debug($"CutNames. Range selected is {selection.Address}");

            // TODO
        }

        public void HighlightDuplicates(IRibbonControl ribbonControl)
        {
            var selection = (Range)App.Selection;

            if (selection == null)
                return;

            logger.Debug($"HighlightDuplicates. Range selected is {selection.Address}");

            // TODO
        }

        public void ToggleCase(IRibbonControl ribbonControl)
        {
             var selection = (Range)App.Selection;

            if (selection == null)
                return;

            logger.Debug($"ToggleCase. Range selected is {selection.Address}");

            // TODO
        }

        public void TrimSpaces(IRibbonControl ribbonControl)
        {
             var selection = (Range)App.Selection;

            if (selection == null)
                return;

            logger.Debug($"TrimSpaces. Range selected is {selection.Address}");

            // TODO
        }

        public void UnmergeCells(IRibbonControl ribbonControl)
        {
             var selection = (Range)App.Selection;

            if (selection == null)
                return;

            logger.Debug($"UnmergeCells. Range selected is {selection.Address}");
            
            using (var scope = container.BeginLifetimeScope())
            {
                var cellsUnmerger = scope.Resolve<ICellsUnmerger>();
                cellsUnmerger.Unmerge(selection);
            }

        }

        public void ValidateValues(IRibbonControl ribbonControl)
        {
            var selection = (Range)App.Selection;

            if (selection == null)
                return;

            logger.Debug($"UnmergeCells. Range selected is {selection.Address}");

            // TODO
        }

        public void FindErrors(IRibbonControl ribbonControl)
        {
            var activeSheet = (Worksheet)App.ActiveSheet;
            var range = activeSheet.UsedRange;

            IReadOnlyCollection<ErroredRange> allErrors;
            using (var scope = container.BeginLifetimeScope())
            {
                var errorFinder = scope.Resolve<IErrorFinder>();
                allErrors = errorFinder.GetAllErrorCells(range);
            }

            if (allErrors.Count == 0)
            {
                dialogService.ShowInfo(UIStrings.NoErrors);
                return;
            }
            form = new SearchRangeResultForm(allErrors, activeSheet);
            form.Show();
        }

        #region XML Tools
        public void CreateSampleXml(IRibbonControl ribbonControl)
        {
            logger.Debug("CreateSampleXml");
            // TODO
        }
        public void ValidateXml(IRibbonControl ribbonControl)
        {
            logger.Debug("ValidateXml");
            // TODO
        }
        #endregion
        #endregion

        #region Utils
        public string GetLabel(IRibbonControl ribbonControl)
        {
            return RibbonLabels.ResourceManager.GetString(ribbonControl.Id);
        }
        public Bitmap GetImage(string imageName)
        {
            return (Bitmap)RibbonIcons.ResourceManager.GetObject(imageName);
        }

        private static string GetResourceText()
        {
            var asm = typeof(NavfertyRibbon).Assembly;

            using (var stream = asm.GetManifestResourceStream("NavfertyExcelAddIn.NavfertyRibbon.xml"))
            using (var resourceReader = new StreamReader(stream))
            {
                return resourceReader.ReadToEnd();
            }
        }

        public void Dispose()
        {
            form?.Dispose();
            container.Dispose();
        }
        #endregion
    }
}
