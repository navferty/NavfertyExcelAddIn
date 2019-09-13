using System;
using System.IO;
using System.Linq;

using NavfertyExcelAddIn.Commons;
using NavfertyExcelAddIn.Localization;

using NLog;

namespace NavfertyExcelAddIn.XmlTools
{
    public class XmlValidator : IXmlValidator
    {
        private readonly IDialogService dialogService;
        private readonly IXsdSchemaValidator xsdValidator;
        private readonly Logger logger = LogManager.GetCurrentClassLogger();

        public XmlValidator(IDialogService dialogService, IXsdSchemaValidator xsdValidator)
        {
            this.dialogService = dialogService;
            this.xsdValidator = xsdValidator;
        }

        public void Validate()
        {
            var xmlFileName = dialogService.AskForFiles(false, FileType.Xml).FirstOrDefault();

            if (string.IsNullOrEmpty(xmlFileName))
            {
                return;
            }

            // allow multiple xsd files as one schema
            var xsdFileNames = dialogService.AskForFiles(true, FileType.Xsd).ToArray();

            if (string.IsNullOrEmpty(xsdFileNames.FirstOrDefault()))
            {
                return;
            }

            if (!File.Exists(xmlFileName) || xsdFileNames.Any(path => !File.Exists(path)))
            {
                throw new ArgumentException("One or more files not found");
            }

            logger.Debug($"Try to validate {xmlFileName} by {string.Join(", ", xsdFileNames)}");

            var validationErrors = xsdValidator.Validate(xmlFileName, xsdFileNames);

            if (!validationErrors.Any())
            {
                dialogService.ShowInfo(UIStrings.SuccessfullyValidatedMessage);
                return;
            }

            // dialogService.ShowInfo(UIStrings.SuccessfullyValidatedMessage);
            logger.Debug($"{validationErrors.Count} errors");

            // TODO write errors to excel worksheet in created workbook
            dialogService.ShowInfo($"Validation messages:\r\n" + string.Join("\r\n", validationErrors));

            //var headers = new[] { new CellItem(1, 1, "Errors"), new CellItem(1, 2, "Warnings") };
            //var i = 2;
            //var errors = xsdValidator.Errors.Select(error => new CellItem(i++, 1, error)).ToList();
            //i = 2;
            //var warnings = xsdValidator.Warnings.Select(warning => new CellItem(i++, 2, warning)).ToList();
            //var cells = headers.Concat(errors).Concat(warnings).ToList();
            //_excelAdapter.WriteCellsToNewSheet(cells, $"Errors {DateTime.Now:yy-MM-dd HH-mm-ss}");
        }
    }
}
