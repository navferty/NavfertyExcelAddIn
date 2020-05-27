using System;
using System.Linq;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Diagnostics.CodeAnalysis;
using System.Windows.Forms;

using NavfertyExcelAddIn.DataValidation;
using NavfertyExcelAddIn.FindFormulaErrors;
using NavfertyExcelAddIn.ParseNumerics;
using NavfertyExcelAddIn.UnprotectWorkbook;
using NavfertyExcelAddIn.WorksheetCellsEditing;
using NavfertyExcelAddIn.InteractiveRangeReport;
using NavfertyExcelAddIn.XmlTools;
using NavfertyExcelAddIn.Localization;
using NavfertyExcelAddIn.Commons;
using NavfertyExcelAddIn.Transliterate;

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
        #region Private members
        private static readonly IContainer container = Registry.CreateContainer();
        private readonly ILogger logger = LogManager.GetCurrentClassLogger();
        private readonly IDialogService dialogService = container.Resolve<IDialogService>();

        private readonly Dictionary<string, ValidationType> validationTypeByButtonId =
            new Dictionary<string, ValidationType>
            {
                { "ValidateValuesNumerics", ValidationType.Numeric },
                { "ValidateValuesXml", ValidationType.Xml },
                { "ValidateValuesDate", ValidationType.Date },
                { "ValidateValuesTinPersonal", ValidationType.TinPersonal },
                { "ValidateValuesTinOrganization", ValidationType.TinOrganization }
            };

        private Application App => Globals.ThisAddIn.Application;

        #region Forms
        private InteractiveRangeReportForm form;
        #endregion
        #endregion

        #region IRibbonExtensibility
        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText();
        }
        #endregion

        #region Ribbon callbacks
        public void RibbonLoad(IRibbonUI ribbonUI)
        {
            logger.Debug($"Ribbon loaded");
        }

        #region Common tools
        public void ParseNumerics(IRibbonControl ribbonControl)
        {
            var range = GetSelectionOrUsedRange(App.ActiveSheet);

            if (range == null)
                return;

            logger.Debug($"Parse numerics for range {range.Address}");

            using (var scope = container.BeginLifetimeScope())
            {
                var parser = scope.Resolve<INumericParser>();
                parser.Parse(range);
            }
        }

        public void ReplaceChars(IRibbonControl ribbonControl)
        {
            var replacer = GetService<ICyrillicLettersReplacer>();

            Range selection = GetSelectionOrUsedRange(App.ActiveSheet);
            selection.ApplyForEachCellOfType<string, object>(
                value =>
                {
                    var newValue = replacer.ReplaceCyrillicCharsWithLatin(value);
                    return (object)newValue ?? value;
                });
        }

        public void Transliterate(IRibbonControl ribbonControl)
        {
            var transliterator = GetService<ITransliterator>();

            Range selection = GetSelectionOrUsedRange(App.ActiveSheet);
            selection.ApplyForEachCellOfType<string, object>(
                value =>
                {
                    var newValue = transliterator.Transliterate(value);
                    return (object)newValue ?? value;
                });
        }

        public void UnprotectWorkbook(IRibbonControl ribbonControl)
        {
            var wb = App.ActiveWorkbook;
            var path = wb.FullName;

            logger.Debug($"UnprotectWorkbook {path}");

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
            var range = GetSelectionOrUsedRange(App.ActiveSheet);

            if (range == null)
                return;

            logger.Debug($"CutNames. Range selected is {range.Address}");

            // TODO
        }

        public void HighlightDuplicates(IRibbonControl ribbonControl)
        {
            var range = GetSelectionOrUsedRange(App.ActiveSheet);

            if (range == null)
                return;

            logger.Debug($"HighlightDuplicates. Range selected is {range.Address}");

            using (var scope = container.BeginLifetimeScope())
            {
                var duplicatesHighlighter = scope.Resolve<IDuplicatesHighlighter>();
                duplicatesHighlighter.HighlightDuplicates(range);
            }
        }

        public void ToggleCase(IRibbonControl ribbonControl)
        {
            var range = GetSelectionOrUsedRange(App.ActiveSheet);

            if (range == null)
                return;

            logger.Debug($"ToggleCase. Range selected is {range.Address}");

            using (var scope = container.BeginLifetimeScope())
            {
                var caseToggler = scope.Resolve<ICaseToggler>();
                caseToggler.ToggleCase(range);
            }
        }

        public void TrimSpaces(IRibbonControl ribbonControl)
        {
            var range = GetSelectionOrUsedRange(App.ActiveSheet);

            if (range == null)
                return;

            logger.Debug($"TrimSpaces. Range selected is {range.Address}");

            using (var scope = container.BeginLifetimeScope())
            {
                var trimmer = scope.Resolve<IEmptySpaceTrimmer>();
                trimmer.TrimSpaces(range);
            }
        }

        public void UnmergeCells(IRibbonControl ribbonControl)
        {
            var range = GetSelectionOrUsedRange(App.ActiveSheet);

            if (range == null)
                return;

            logger.Debug($"UnmergeCells. Range selected is {range.Address}");

            using (var scope = container.BeginLifetimeScope())
            {
                var cellsUnmerger = scope.Resolve<ICellsUnmerger>();
                cellsUnmerger.Unmerge(range);
            }

        }

        public void ValidateValues(IRibbonControl ribbonControl)
        {
            var activeSheet = (Worksheet)App.ActiveSheet;
            var range = GetSelectionOrUsedRange(activeSheet);

            if (range == null)
                return;

            logger.Debug($"ValidateValues. Range selected is {range.Address}");

            if (!validationTypeByButtonId.TryGetValue(ribbonControl.Id, out var validationType))
            {
                dialogService.ShowError($"Invalid control id '{ribbonControl.Id}'");
                throw new ArgumentOutOfRangeException($"Invalid control id '{ribbonControl.Id}'");
            }

            logger.Debug($"ValidateValues. Range selected is {range.Address}, validation type {validationType}");

            IReadOnlyCollection<InteractiveErrorItem> results;
            using (var scope = container.BeginLifetimeScope())
            {
                var validator = scope.Resolve<ICellsValueValidator>();
                results = validator.Validate(range, validationType);
            }

            form = new InteractiveRangeReportForm(results, activeSheet);
            form.Show();
        }

        public void FindErrors(IRibbonControl ribbonControl)
        {
            var activeSheet = (Worksheet)App.ActiveSheet;
            var range = GetSelectionOrUsedRange(activeSheet);

            if (range == null)
                return;

            logger.Debug($"FindErrors. Range selected is {range.Address}");

            IReadOnlyCollection<InteractiveErrorItem> allErrors;
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
            form = new InteractiveRangeReportForm(allErrors, activeSheet);
            form.Show();
        }

        public void CopyAsMarkdown(IRibbonControl ribbonControl)
        {
            var range = GetSelectionOrUsedRange(App.ActiveSheet);

            if (range == null)
                return;

            logger.Debug($"CopyAsMarkdown. Range selected is {range.Address}");

            string table;
            using (var scope = container.BeginLifetimeScope())
            {
                var markdownReader = scope.Resolve<ICellsToMarkdownReader>();
                table = markdownReader.ReadToMarkdown(range);
            }

            if (!string.IsNullOrWhiteSpace(table))
            {
                Clipboard.SetText(table);
            }
        }
        #endregion

        #region XML Tools
        public void CreateSampleXml(IRibbonControl ribbonControl)
        {
            logger.Debug("CreateSampleXml pressed");

            using (var scope = container.BeginLifetimeScope())
            {
                var xmlSampleCreator = scope.Resolve<IXmlSampleCreator>();
                xmlSampleCreator.CreateSampleXml();
            }
        }
        public void ValidateXml(IRibbonControl ribbonControl)
        {
            logger.Debug("ValidateXml pressed");

            using (var scope = container.BeginLifetimeScope())
            {
                var validator = scope.Resolve<IXmlValidator>();
                validator.Validate(App);
            }
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

        private T GetService<T>()
        {
            using (var scope = container.BeginLifetimeScope())
                return scope.Resolve<T>();
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

        private Range GetSelectionOrUsedRange(Worksheet activeSheet)
        {
            var selection = (Range)App.Selection;

            return selection.Cells.Count > 1
                ? selection
                : activeSheet.UsedRange;
        }

        public void Dispose()
        {
            form?.Dispose();
            container.Dispose();
        }
        #endregion
    }
}
