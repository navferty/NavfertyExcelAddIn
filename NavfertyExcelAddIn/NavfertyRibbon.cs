using System;
using System.Linq;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Diagnostics.CodeAnalysis;

using NavfertyExcelAddIn.ParseNumerics;
using NavfertyExcelAddIn.Localization;

using NLog;
using Autofac;

using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

using Application = Microsoft.Office.Interop.Excel.Application;
using NavfertyExcelAddIn.Commons;

namespace NavfertyExcelAddIn
{
    [ComVisible(true)]
    [SuppressMessage("Style", "IDE0060:Remove unused parameter", Justification = "Ribbon callbacks must have certain signature")]
    public class NavfertyRibbon : IRibbonExtensibility, IDisposable
    {
        private static readonly IContainer container = Registry.CreateContainer();
        private readonly ILogger logger = LogManager.GetCurrentClassLogger();
        private Application App => Globals.ThisAddIn.Application;

        #region Forms
        private CellsResult form;
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

            // TODO
        }

        public void FindErrors(IRibbonControl ribbonControl)
        {
            var range = ((Worksheet)App.ActiveSheet).UsedRange;

            ErroredRange[] allErrors;
            using (var scope = container.BeginLifetimeScope())
            {
                var errorFinder = scope.Resolve<IErrorFinder>();
                allErrors = errorFinder.GetAllErrorCells(range).ToArray();
            }
            form = new CellsResult(allErrors);
            form.Show();
        }

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

        #region Utils
        public string GetLabel(IRibbonControl ribbonControl)
        {
            return RibbonLabels.ResourceManager.GetString(ribbonControl.Id);
        }
        public Bitmap GetImage(string imageName)
        {
            return (Bitmap)RibbonIcons.ResourceManager.GetObject(imageName);
        }
        #endregion

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
